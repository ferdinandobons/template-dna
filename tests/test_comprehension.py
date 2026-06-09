# SPDX-License-Identifier: MIT
"""Foundation tests for the model-comprehension layer (schema 1.2.0).

These run WITHOUT a model: they exercise the schema sink, the fail-closed
membership contract, the gate wiring, the sha-bound cache, and idempotency -
exactly the path CI exercises. The comprehension block is additive and optional,
so a profile without it (every other test) stays valid and on the deterministic
path.
"""

from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from brandkit.profile import comprehension as comp_mod
from brandkit.profile import schema, store
from brandkit.qa.gate import run_qa


def _docx_profile_with_inventory() -> dict:
    """A docx profile carrying a surfaced cover-anchor/field/region inventory."""
    prof = schema.build_envelope("docx", {"name": "t"})
    prof["provenance"]["shell"]["sha256"] = "abc123"
    prof["surface"]["docx"]["cover_anchors"] = [{"id": "title"}, {"id": "subtitle"}]
    prof["surface"]["docx"]["fields"] = [{"id": "tot.1", "seq": "Table"}]
    prof["surface"]["docx"]["regions"] = [{"id": "body.demo"}]
    prof["roles"] = {
        "_index": ["caption", "cover.title"],
        "caption": {"resolver": {"type": "named_style", "style_id": "Caption"}},
        "cover.title": {"resolver": {"type": "named_style", "style_id": "Title"}},
    }
    return prof


def _valid_comp() -> dict:
    return {
        "confidence": 0.9,
        "cover_slots": {
            "title": {"fill_rule": "in_place", "binds_to": "title", "demo_value": "X"}
        },
        "conventions": {
            "indexes": [
                {
                    "index_ref": "tot.1",
                    "reconcile": "regenerate",
                    "seq_id": "Table",
                    "feeds_from_role_id": "caption",
                    "kind": "table_index",
                }
            ],
            "sections": [{"region_ref": "body.demo", "required": False}],
        },
        "role_annotations": {"caption": {"purpose": "captions"}},
        "demo_classification": {
            "regions": [
                {"region_ref": "body.demo", "verdict": "demo", "evidence": "sample"}
            ]
        },
    }


class SchemaAdditiveTest(unittest.TestCase):
    def test_version_bumped(self):
        self.assertEqual(schema.SCHEMA_VERSION, "1.2.0")

    def test_current_major_profile_validates_clean(self):
        # A normal 1.2.0 profile passes the major floor and validates clean.
        prof = schema.build_envelope("docx", {"name": "t"})
        self.assertEqual(prof["schema_version"], "1.2.0")
        self.assertEqual(schema.validate(prof), [])

    def test_newer_major_is_a_single_clear_error(self):
        # A future MAJOR (2.0.0) must short-circuit to ONE actionable message,
        # not a scatter of per-field enum errors from the rest of validate().
        prof = schema.build_envelope("docx", {"name": "t"})
        prof["schema_version"] = "2.0.0"
        problems = schema.validate(prof)
        self.assertEqual(len(problems), 1, problems)
        self.assertIn("major", problems[0])
        self.assertIn("2.0.0", problems[0])
        self.assertIn(str(schema.SUPPORTED_MAJOR), problems[0])

    def test_supported_major_boundary(self):
        # The floor is exclusive: the supported major itself is fine; the next
        # one up is refused.
        prof = schema.build_envelope("docx", {"name": "t"})
        prof["schema_version"] = f"{schema.SUPPORTED_MAJOR}.99.99"
        self.assertEqual(schema.validate(prof), [])
        prof["schema_version"] = f"{schema.SUPPORTED_MAJOR + 1}.0.0"
        self.assertEqual(len(schema.validate(prof)), 1)

    def test_migrate_is_identity_today(self):
        prof = schema.build_envelope("pptx", {"name": "t"})
        self.assertEqual(schema.migrate(prof), prof)

    def test_absent_comprehension_is_valid(self):
        prof = schema.build_envelope("docx", {"name": "t"})
        self.assertEqual(prof["comprehension"]["status"], "absent")
        self.assertEqual(schema.validate(prof), [])

    def test_profile_without_comprehension_key_still_valid(self):
        prof = schema.build_envelope("pptx", {"name": "t"})
        del prof["comprehension"]
        self.assertEqual(schema.validate(prof), [])

    def test_present_well_shaped_is_valid(self):
        prof = _docx_profile_with_inventory()
        prof["comprehension"]["status"] = "present"
        prof["comprehension"]["cover_slots"] = {"title": {"fill_rule": "in_place"}}
        self.assertEqual(schema.validate(prof), [])

    def test_bad_executor_enums_are_reported(self):
        prof = _docx_profile_with_inventory()
        prof["comprehension"]["status"] = "weird"
        prof["comprehension"]["cover_slots"] = {"title": {"fill_rule": "nuke"}}
        prof["comprehension"]["conventions"]["indexes"] = [
            {"index_ref": "tot.1", "reconcile": "zap"}
        ]
        prof["comprehension"]["demo_classification"]["regions"] = [
            {"region_ref": "body.demo", "verdict": "maybe"}
        ]
        problems = schema.validate(prof)
        self.assertTrue(any("status" in p for p in problems))
        self.assertTrue(any("fill_rule" in p for p in problems))
        self.assertTrue(any("reconcile" in p for p in problems))
        self.assertTrue(any("verdict" in p for p in problems))

    def test_open_region_token_accepted_bad_syntax_rejected(self):
        prof = schema.build_envelope("pptx", {"name": "t"})
        prof["structure"] = {
            "ordered": True,
            "skeleton": [{"region": "agenda", "order": 1, "demo": True}],
        }
        self.assertEqual(schema.validate(prof), [])
        prof["structure"]["skeleton"][0]["region"] = "Bad Name!"
        self.assertTrue(any("region" in p for p in schema.validate(prof)))


class MembershipFailClosedTest(unittest.TestCase):
    def test_valid_comprehension_merges(self):
        prof = _docx_profile_with_inventory()
        res = comp_mod.merge(prof, _valid_comp())
        self.assertTrue(res.ok, res.problems)
        self.assertEqual(prof["comprehension"]["status"], "present")

    def test_dangling_anchor_ref_is_rejected(self):
        prof = _docx_profile_with_inventory()
        comp = _valid_comp()
        comp["cover_slots"] = {"ghost": {"fill_rule": "in_place"}}
        res = comp_mod.merge(prof, comp)
        self.assertFalse(res.ok)
        self.assertEqual(prof["comprehension"]["status"], "rejected")
        self.assertTrue(any("ghost" in p for p in res.problems))

    def test_empty_inventory_ref_is_error_not_skipped(self):
        """A ref into an EMPTY inventory is an ERROR (fail-closed), unlike the
        namespace-guarded resolver-consistency check which no-ops on empty."""
        prof = schema.build_envelope("docx", {"name": "t"})  # no surfaced fields
        prof["provenance"]["shell"]["sha256"] = "x"
        comp = {
            "conventions": {"indexes": [{"index_ref": "ghost", "reconcile": "clear"}]}
        }
        res = comp_mod.merge(prof, comp)
        self.assertFalse(res.ok)
        self.assertTrue(any("ghost" in p and "fields" in p for p in res.problems))

    def test_dangling_role_annotation_rejected(self):
        prof = _docx_profile_with_inventory()
        comp = _valid_comp()
        comp["role_annotations"] = {"nonexistent.role": {"purpose": "x"}}
        res = comp_mod.merge(prof, comp)
        self.assertFalse(res.ok)


class GateWiringTest(unittest.TestCase):
    """The contract is enforced by WIRING into run_qa, not tuple membership."""

    def test_dangling_ref_fails_the_gate_through_run_qa(self):
        prof = _docx_profile_with_inventory()
        # Hand-craft a PRESENT comprehension with a dangling anchor (bypassing the
        # merge writer) to prove run_qa independently rejects it.
        prof["comprehension"] = {
            "schema_version": "comprehension-1",
            "status": "present",
            "source_shell_sha256": "abc123",
            "confidence": 0.5,
            "cover_slots": {"ghost": {"fill_rule": "in_place"}},
            "conventions": {"indexes": [], "sections": []},
            "role_annotations": {},
            "demo_classification": {"regions": []},
        }
        report = run_qa(None, prof, qa="fast", shell=None)
        self.assertFalse(report.passed)
        self.assertTrue(
            any(
                f.check == "comprehension_targets_exist" and f.severity == "ERROR"
                for f in report.findings
            ),
            [f.message for f in report.findings],
        )

    def test_absent_comprehension_passes_the_gate(self):
        prof = _docx_profile_with_inventory()
        report = run_qa(None, prof, qa="fast", shell=None)
        self.assertFalse(
            any(f.check == "comprehension_targets_exist" for f in report.findings)
        )

    def test_invariant_id_declared(self):
        self.assertIn("comprehension_targets_exist", schema.DEFAULT_L0_INVARIANTS)
        self.assertIn("no_net_structure_loss", schema.DEFAULT_L0_INVARIANTS)


class CacheBindingTest(unittest.TestCase):
    def test_present_only_when_sha_matches(self):
        prof = _docx_profile_with_inventory()
        comp_mod.merge(prof, _valid_comp())
        self.assertTrue(store.comprehension_is_present(prof))
        # Drift the shell hash -> comprehension no longer counts as present.
        prof["provenance"]["shell"]["sha256"] = "different"
        self.assertFalse(store.comprehension_is_present(prof))

    def test_absent_is_not_present(self):
        prof = schema.build_envelope("docx", {"name": "t"})
        self.assertFalse(store.comprehension_is_present(prof))


class IdempotencyTest(unittest.TestCase):
    def test_comprehend_twice_byte_identical(self):
        prof_a = _docx_profile_with_inventory()
        prof_b = _docx_profile_with_inventory()
        comp_mod.merge(prof_a, _valid_comp())
        comp_mod.merge(prof_b, _valid_comp())
        self.assertEqual(
            json.dumps(prof_a["comprehension"], sort_keys=True),
            json.dumps(prof_b["comprehension"], sort_keys=True),
        )

    def test_remerge_is_stable(self):
        prof = _docx_profile_with_inventory()
        comp_mod.merge(prof, _valid_comp())
        first = json.dumps(prof["comprehension"], sort_keys=True)
        comp_mod.merge(prof, _valid_comp())
        second = json.dumps(prof["comprehension"], sort_keys=True)
        self.assertEqual(first, second)


def _comp_with_fragment(kind="component", ref="note_box", blocks=None):
    """A comprehension proposing one reusable fragment (no inventory refs)."""
    if blocks is None:
        blocks = [{"type": "callout", "intent": "note", "runs": [{"t": "{{body}}"}]}]
    return {
        "confidence": 0.9,
        "fragments": [
            {"ref": ref, "kind": kind, "purpose": "reusable note", "blocks": blocks}
        ],
    }


class FragmentPopulationTest(unittest.TestCase):
    """A model-proposed fragment is validated fail-closed and, on a clean merge,
    DERIVED into the profile's components/sections registry (the milestone:
    auto-population through the comprehend boundary, no hardcoded catalog)."""

    def test_valid_component_fragment_lands_in_registry(self):
        prof = _docx_profile_with_inventory()
        res = comp_mod.merge(prof, _comp_with_fragment())
        self.assertTrue(res.ok, res.problems)
        self.assertEqual(prof["comprehension"]["status"], "present")
        self.assertIn("note_box", prof["components"])
        self.assertEqual(prof["components"]["note_box"]["blocks"][0]["type"], "callout")
        self.assertEqual(prof["components"]["note_box"]["purpose"], "reusable note")
        # The proposal is also recorded in the canonical comprehension block.
        self.assertEqual(len(prof["comprehension"]["fragments"]), 1)
        # Sections registry stays empty (a component proposal only feeds components).
        self.assertEqual(prof["sections"], {})

    def test_section_fragment_lands_in_sections(self):
        prof = _docx_profile_with_inventory()
        res = comp_mod.merge(
            prof,
            _comp_with_fragment(
                kind="section",
                ref="opener",
                blocks=[
                    {"type": "heading", "level": 1, "text": "X"},
                    {"type": "divider"},
                ],
            ),
        )
        self.assertTrue(res.ok, res.problems)
        self.assertIn("opener", prof["sections"])
        self.assertEqual(prof["components"], {})

    def test_bad_block_type_is_rejected_fail_closed(self):
        prof = _docx_profile_with_inventory()
        res = comp_mod.merge(prof, _comp_with_fragment(blocks=[{"type": "bogus"}]))
        self.assertFalse(res.ok)
        self.assertEqual(prof["comprehension"]["status"], "rejected")
        self.assertEqual(prof["components"], {})  # registry left untouched
        self.assertTrue(any("bogus" in p for p in res.problems), res.problems)

    def test_missing_blocks_is_rejected(self):
        prof = _docx_profile_with_inventory()
        res = comp_mod.merge(prof, {"fragments": [{"ref": "x", "kind": "component"}]})
        self.assertFalse(res.ok)
        self.assertEqual(prof["components"], {})
        self.assertTrue(any("blocks" in p for p in res.problems), res.problems)

    def test_bad_kind_is_rejected(self):
        prof = _docx_profile_with_inventory()
        res = comp_mod.merge(
            prof,
            {
                "fragments": [
                    {"ref": "x", "kind": "widget", "blocks": [{"type": "divider"}]}
                ]
            },
        )
        self.assertFalse(res.ok)
        self.assertTrue(any("kind" in p for p in res.problems), res.problems)

    def test_dangling_nested_ref_is_rejected(self):
        prof = _docx_profile_with_inventory()
        res = comp_mod.merge(
            prof,
            _comp_with_fragment(
                kind="section",
                ref="wrap",
                blocks=[{"type": "component", "ref": "ghost"}],
            ),
        )
        self.assertFalse(res.ok)
        self.assertEqual(prof["sections"], {})
        self.assertTrue(any("ghost" in p for p in res.problems), res.problems)

    def test_nested_ref_to_proposed_fragment_is_allowed(self):
        prof = _docx_profile_with_inventory()
        res = comp_mod.merge(
            prof,
            {
                "fragments": [
                    {
                        "ref": "leaf",
                        "kind": "component",
                        "blocks": [{"type": "paragraph", "text": "x"}],
                    },
                    {
                        "ref": "wrap",
                        "kind": "section",
                        "blocks": [{"type": "component", "ref": "leaf"}],
                    },
                ]
            },
        )
        self.assertTrue(res.ok, res.problems)
        self.assertIn("leaf", prof["components"])
        self.assertIn("wrap", prof["sections"])

    def test_duplicate_ref_is_rejected(self):
        prof = _docx_profile_with_inventory()
        res = comp_mod.merge(
            prof,
            {
                "fragments": [
                    {
                        "ref": "dup",
                        "kind": "component",
                        "blocks": [{"type": "divider"}],
                    },
                    {
                        "ref": "dup",
                        "kind": "component",
                        "blocks": [{"type": "divider"}],
                    },
                ]
            },
        )
        self.assertFalse(res.ok)
        self.assertTrue(any("duplicate" in p for p in res.problems), res.problems)

    def test_fragment_merge_is_idempotent(self):
        prof_a = _docx_profile_with_inventory()
        prof_b = _docx_profile_with_inventory()
        comp_mod.merge(prof_a, _comp_with_fragment())
        comp_mod.merge(prof_b, _comp_with_fragment())
        self.assertEqual(
            json.dumps(prof_a["comprehension"], sort_keys=True),
            json.dumps(prof_b["comprehension"], sort_keys=True),
        )
        self.assertEqual(
            json.dumps(prof_a["components"], sort_keys=True),
            json.dumps(prof_b["components"], sort_keys=True),
        )

    def test_fragmentless_comprehend_leaves_registries_empty(self):
        # A comprehension with no fragments (the norm) populates nothing.
        prof = _docx_profile_with_inventory()
        res = comp_mod.merge(prof, _valid_comp())
        self.assertTrue(res.ok, res.problems)
        self.assertEqual(prof["components"], {})
        self.assertEqual(prof["sections"], {})
        self.assertEqual(prof["comprehension"]["fragments"], [])

    def test_input_status_does_not_bypass_fragment_validation(self):
        # merge DISPOSES status: a model-supplied status='rejected'/'absent' must
        # NOT short-circuit fragment validation (merge derives the registry
        # regardless of status), so a bad fragment is still rejected, writing
        # nothing into the registries.
        for status in ("rejected", "absent", "present"):
            prof = _docx_profile_with_inventory()
            res = comp_mod.merge(
                prof,
                {
                    "status": status,
                    "fragments": [
                        {
                            "ref": "bad",
                            "kind": "component",
                            "blocks": [{"type": "BOGUS"}],
                        }
                    ],
                },
            )
            self.assertFalse(res.ok, f"status={status} bypassed validation")
            self.assertEqual(prof["comprehension"]["status"], "rejected")
            self.assertEqual(prof["components"], {})

    def test_cyclic_fragment_refs_are_rejected(self):
        prof = _docx_profile_with_inventory()
        res = comp_mod.merge(
            prof,
            {
                "fragments": [
                    {
                        "ref": "a",
                        "kind": "section",
                        "blocks": [{"type": "section", "ref": "b"}],
                    },
                    {
                        "ref": "b",
                        "kind": "section",
                        "blocks": [{"type": "section", "ref": "a"}],
                    },
                ]
            },
        )
        self.assertFalse(res.ok)
        self.assertEqual(prof["sections"], {})
        self.assertTrue(any("cyclic" in p for p in res.problems), res.problems)

    def test_self_referential_fragment_is_rejected(self):
        prof = _docx_profile_with_inventory()
        res = comp_mod.merge(
            prof,
            {
                "fragments": [
                    {
                        "ref": "loop",
                        "kind": "section",
                        "blocks": [{"type": "section", "ref": "loop"}],
                    }
                ]
            },
        )
        self.assertFalse(res.ok)
        self.assertTrue(any("cyclic" in p for p in res.problems), res.problems)

    def test_diamond_dag_of_fragments_is_not_a_false_cycle(self):
        # A->B, A->C, B->D, C->D is a DAG (no cycle): it must merge clean, proving
        # the cycle detector does not false-positive on shared descendants.
        prof = _docx_profile_with_inventory()

        def sec(ref, *refs):
            return {
                "ref": ref,
                "kind": "section",
                "blocks": [{"type": "section", "ref": r} for r in refs]
                or [{"type": "divider"}],
            }

        res = comp_mod.merge(
            prof,
            {"fragments": [sec("a", "b", "c"), sec("b", "d"), sec("c", "d"), sec("d")]},
        )
        self.assertTrue(res.ok, res.problems)
        self.assertEqual(set(prof["sections"]), {"a", "b", "c", "d"})

    def test_nested_ref_to_unproposed_existing_entry_is_rejected(self):
        # The registry is rebuilt from the proposal ALONE, so a nested ref to an
        # existing-but-not-reproposed entry would be dangling after the rebuild and
        # is rejected (also keeps the merge outcome a pure function of the input).
        prof = _docx_profile_with_inventory()
        prof["components"] = {"pre": {"blocks": [{"type": "divider"}]}}
        res = comp_mod.merge(
            prof,
            _comp_with_fragment(
                kind="section", ref="wrap", blocks=[{"type": "component", "ref": "pre"}]
            ),
        )
        self.assertFalse(res.ok)
        self.assertTrue(any("pre" in p for p in res.problems), res.problems)


class CliComprehendTest(unittest.TestCase):
    def test_comprehend_input_and_comprehend_roundtrip(self):
        import os
        from brandkit.cli import main

        sys.path.insert(0, str(Path(__file__).resolve().parents[0]))
        from test_smoke import _synthetic_template

        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            old = Path.cwd()
            os.chdir(tmp)
            try:
                template = tmp / "t.docx"
                _synthetic_template(template)
                self.assertEqual(
                    main(
                        [
                            "extract",
                            "--name",
                            "acme",
                            "--template",
                            str(template),
                            "--scope",
                            "project",
                        ]
                    ),
                    0,
                )
                # comprehend-input prints a valid bundle.
                import io
                from contextlib import redirect_stdout

                buf = io.StringIO()
                with redirect_stdout(buf):
                    rc = main(
                        ["comprehend-input", "--name", "acme", "--scope", "project"]
                    )
                self.assertEqual(rc, 0)
                bundle = json.loads(buf.getvalue())
                self.assertIn("inventories", bundle["facts"])
                anchors = bundle["facts"]["inventories"]["cover_anchors"]
                anchor_ids = [a["id"] for a in anchors]
                # Cover anchors carry stable, positional ids (one per slot), not a
                # single hardcoded "title". The synthetic ``{{title}}`` placeholder
                # surfaces as a ``para.<i>`` slot.
                title_anchor = next(
                    a["id"]
                    for a in anchors
                    if "{{title}}" in (a.get("placeholder") or "")
                )
                self.assertTrue(title_anchor.startswith("para."), anchor_ids)

                # comprehend merges a valid block over the real surfaced anchor.
                comp = tmp / "comp.json"
                comp.write_text(
                    json.dumps(
                        {
                            "cover_slots": {
                                title_anchor: {
                                    "fill_rule": "in_place",
                                    "binds_to": "title",
                                    "demo_value": "{{title}}",
                                }
                            }
                        }
                    ),
                    encoding="utf-8",
                )
                self.assertEqual(
                    main(
                        [
                            "comprehend",
                            "--name",
                            "acme",
                            "--input",
                            str(comp),
                            "--scope",
                            "project",
                        ]
                    ),
                    0,
                )
                prof = json.loads(
                    (tmp / "brand-kit" / "acme" / "profile.json").read_text()
                )
                self.assertEqual(prof["comprehension"]["status"], "present")
                self.assertTrue(prof["comprehension"]["source_shell_sha256"])

                # A dangling ref is rejected at the CLI boundary (exit 1).
                bad = tmp / "bad.json"
                bad.write_text(
                    json.dumps({"cover_slots": {"ghost": {"fill_rule": "in_place"}}}),
                    encoding="utf-8",
                )
                self.assertEqual(
                    main(
                        [
                            "comprehend",
                            "--name",
                            "acme",
                            "--input",
                            str(bad),
                            "--scope",
                            "project",
                        ]
                    ),
                    1,
                )
                prof2 = json.loads(
                    (tmp / "brand-kit" / "acme" / "profile.json").read_text()
                )
                self.assertEqual(prof2["comprehension"]["status"], "rejected")
            finally:
                os.chdir(old)


if __name__ == "__main__":
    unittest.main()
