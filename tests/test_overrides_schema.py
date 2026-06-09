# SPDX-License-Identifier: MIT
"""Shape + freeze tests for the learned-overrides block (Cluster B, B3a).

These run WITHOUT a model and WITHOUT any reader consuming overrides yet: they
exercise the additive schema sink (``empty_overrides`` / ``_validate_overrides`` /
``OverrideKind`` / ``LEARNABLE_CHECKS``), its wiring into ``build_envelope`` and
``validate``, and the sha-bound presence test (``store.overrides_are_present``).
The block mirrors the comprehension contract one-for-one: a profile without it (or
with the reserved-empty ``rules.overrides == {}``, or ``status='absent'``) stays
valid and on the deterministic, byte-identical path.
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

import unittest

from brandkit.profile import schema, store


class OverridesShapeTest(unittest.TestCase):
    def test_schema_version_unchanged(self):
        # B3 is additive: the envelope major.minor.patch must NOT bump.
        self.assertEqual(schema.SCHEMA_VERSION, "1.2.0")

    def test_override_kind_vocab_is_closed(self):
        self.assertEqual(
            schema.OVERRIDE_KINDS,
            frozenset({"reroute_role", "number_format", "register_demo_clear"}),
        )
        self.assertEqual(
            {e.value for e in schema.OverrideKind},
            schema.OVERRIDE_KINDS,
        )

    def test_learnable_checks_are_the_unambiguous_ids(self):
        self.assertEqual(
            schema.LEARNABLE_CHECKS,
            frozenset(
                {
                    "resolver_targets_exist",
                    "style_fallback",
                    "no_residual_template_text",
                }
            ),
        )

    def test_empty_overrides_is_shaped_and_absent(self):
        ov = schema.empty_overrides()
        self.assertEqual(ov["schema_version"], schema.OVERRIDES_SCHEMA_VERSION)
        self.assertEqual(ov["status"], "absent")
        self.assertIsNone(ov["source_shell_sha256"])
        self.assertIsNone(ov["generated_by"])
        self.assertEqual(ov["confidence"], 0.0)
        self.assertEqual(ov["reroute_roles"], {})
        self.assertEqual(ov["number_format_swaps"], {})
        self.assertEqual(ov["demo_clears"], [])
        self.assertEqual(ov["provenance"], {})

    def test_build_envelope_stamps_empty_overrides(self):
        for kind in ("docx", "pptx", "xlsx"):
            prof = schema.build_envelope(kind, {"name": "t"})
            self.assertEqual(prof["rules"]["overrides"], schema.empty_overrides())
            self.assertEqual(prof["rules"]["overrides"]["status"], "absent")

    def test_override_targets_exist_declared_in_l0(self):
        self.assertIn("override_targets_exist", schema.DEFAULT_L0_INVARIANTS)
        prof = schema.build_envelope("xlsx", {"name": "t"})
        self.assertIn("override_targets_exist", prof["qa"]["l0_invariants"])


class OverridesValidateTest(unittest.TestCase):
    def test_absent_overrides_validate_clean(self):
        prof = schema.build_envelope("docx", {"name": "t"})
        self.assertEqual(prof["rules"]["overrides"]["status"], "absent")
        self.assertEqual(schema.validate(prof), [])

    def test_validate_overrides_absent_yields_empty(self):
        self.assertEqual(schema._validate_overrides(None), [])
        self.assertEqual(schema._validate_overrides({}), [])
        self.assertEqual(
            schema._validate_overrides({"status": "absent"}),
            [],
        )

    def test_old_profile_with_reserved_empty_slot_still_validates(self):
        # A profile written before B3 carries the reserved ``rules.overrides == {}``.
        prof = schema.build_envelope("pptx", {"name": "t"})
        prof["rules"]["overrides"] = {}
        self.assertEqual(schema.validate(prof), [])

    def test_profile_without_rules_key_still_valid(self):
        prof = schema.build_envelope("docx", {"name": "t"})
        del prof["rules"]
        # _validate_overrides reads (rules or {}).get("overrides") -> None -> [].
        self.assertNotIn(
            "overrides",
            " ".join(schema.validate(prof)),
        )

    def test_present_well_shaped_overrides_validate(self):
        prof = schema.build_envelope("xlsx", {"name": "t"})
        prof["provenance"]["shell"]["sha256"] = "deadbeef"
        prof["rules"]["overrides"] = {
            "schema_version": schema.OVERRIDES_SCHEMA_VERSION,
            "status": "present",
            "source_shell_sha256": "deadbeef",
            "generated_by": "learn",
            "confidence": 0.8,
            "reroute_roles": {"heading.9": "heading.1"},
            "number_format_swaps": {"metric.value": "#,##0"},
            "demo_clears": ["Lorem ipsum"],
            "provenance": {
                "reroute_role:heading.9": {
                    "check": "resolver_targets_exist",
                    "location": "heading.9",
                    "recurred_runs": 3,
                }
            },
        }
        self.assertEqual(schema.validate(prof), [])

    def test_bad_enum_and_shape_are_reported(self):
        prof = schema.build_envelope("xlsx", {"name": "t"})
        prof["rules"]["overrides"] = {
            "status": "weird",
            "confidence": 5.0,
            "source_shell_sha256": 123,
            "reroute_roles": {"heading.9": 7},
            "number_format_swaps": {"metric.value": ""},
            "demo_clears": ["ok", 9],
            "provenance": "nope",
        }
        problems = schema.validate(prof)
        self.assertTrue(any("status" in p for p in problems))
        self.assertTrue(any("confidence" in p for p in problems))
        self.assertTrue(any("source_shell_sha256" in p for p in problems))
        self.assertTrue(any("reroute_roles" in p for p in problems))
        self.assertTrue(any("number_format_swaps" in p for p in problems))
        self.assertTrue(any("demo_clears" in p for p in problems))
        self.assertTrue(any("provenance" in p for p in problems))

    def test_validate_is_shape_only_not_membership(self):
        # A reroute target that is NOT a declared role is still SHAPE-valid: the
        # fail-closed membership check (check_override_targets) owns that ERROR, not
        # the never-required structural validator.
        prof = schema.build_envelope("docx", {"name": "t"})
        prof["provenance"]["shell"]["sha256"] = "abc"
        prof["rules"]["overrides"] = {
            "status": "present",
            "source_shell_sha256": "abc",
            "reroute_roles": {"heading.9": "role.does.not.exist"},
        }
        self.assertEqual(schema.validate(prof), [])

    def test_non_object_overrides_flagged(self):
        self.assertEqual(
            schema._validate_overrides([1, 2, 3]),
            ["rules.overrides: must be an object"],
        )


class OverridesReadHelpersTest(unittest.TestCase):
    def test_overrides_block_defaults_to_empty_on_missing(self):
        self.assertEqual(schema.overrides_block({}), {})
        self.assertEqual(schema.overrides_block({"rules": "nope"}), {})
        self.assertEqual(schema.overrides_block({"rules": {}}), {})

    def test_list_overrides_flattens_all_three_kinds(self):
        prof = schema.build_envelope("xlsx", {"name": "t"})
        prof["rules"]["overrides"]["reroute_roles"] = {"heading.9": "heading.1"}
        prof["rules"]["overrides"]["number_format_swaps"] = {"m.v": "#,##0"}
        prof["rules"]["overrides"]["demo_clears"] = ["Lorem"]
        entries = schema.list_overrides(prof)
        self.assertIn(("reroute_role", "heading.9", "heading.1"), entries)
        self.assertIn(("number_format", "m.v", "#,##0"), entries)
        self.assertIn(("register_demo_clear", "Lorem", "Lorem"), entries)
        self.assertEqual(len(entries), 3)

    def test_list_overrides_empty_by_default(self):
        prof = schema.build_envelope("docx", {"name": "t"})
        self.assertEqual(schema.list_overrides(prof), [])


class OverridesPresenceFreezeTest(unittest.TestCase):
    def _present_profile(self) -> dict:
        prof = schema.build_envelope("xlsx", {"name": "t"})
        prof["provenance"]["shell"]["sha256"] = "shell-sha-1"
        ov = prof["rules"]["overrides"]
        ov["status"] = "present"
        ov["source_shell_sha256"] = "shell-sha-1"
        ov["reroute_roles"] = {"heading.9": "heading.1"}
        return prof

    def test_absent_overrides_not_present(self):
        prof = schema.build_envelope("docx", {"name": "t"})
        prof["provenance"]["shell"]["sha256"] = "s"
        self.assertFalse(store.overrides_are_present(prof))

    def test_present_and_current_is_present(self):
        self.assertTrue(store.overrides_are_present(self._present_profile()))

    def test_present_but_empty_containers_not_present(self):
        prof = self._present_profile()
        prof["rules"]["overrides"]["reroute_roles"] = {}
        self.assertFalse(store.overrides_are_present(prof))

    def test_drifted_shell_sha_freezes_overrides_out(self):
        prof = self._present_profile()
        # A re-extract re-stamps provenance.shell.sha256: the recorded sha no longer
        # matches, so the lesson is frozen out and the resolver reverts.
        prof["provenance"]["shell"]["sha256"] = "shell-sha-2"
        self.assertFalse(store.overrides_are_present(prof))

    def test_missing_or_malformed_rules_not_present(self):
        self.assertFalse(store.overrides_are_present({}))
        self.assertFalse(store.overrides_are_present({"rules": "x"}))
        self.assertFalse(store.overrides_are_present({"rules": {"overrides": "x"}}))


if __name__ == "__main__":
    unittest.main()
