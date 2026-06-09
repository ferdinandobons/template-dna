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


class AuditSinkTest(unittest.TestCase):
    """C1: the persisted L2 visual-audit verdict (`comprehension.audit`)."""

    def _checklist_id(self, prof: dict) -> str:
        from brandkit.qa.visual import visual_checklist_ids

        ids = visual_checklist_ids(prof)
        self.assertTrue(ids, "test profile must derive a non-empty checklist")
        return ids[0]

    def test_audit_verdict_against_nonmember_id_rejected(self):
        prof = _docx_profile_with_inventory()
        comp = _valid_comp()
        comp["audit"] = {"definitely-not-a-checklist-id": {"verdict": "PASS"}}
        res = comp_mod.merge(prof, comp)
        self.assertFalse(res.ok)
        self.assertEqual(prof["comprehension"]["status"], "rejected")
        self.assertTrue(
            any("definitely-not-a-checklist-id" in p for p in res.problems),
            res.problems,
        )

    def test_valid_audit_verdict_merges(self):
        prof = _docx_profile_with_inventory()
        cid = self._checklist_id(prof)
        comp = _valid_comp()
        comp["audit"] = {
            cid: {"verdict": "PASS", "evidence": "looks right", "content_sha256": "x"}
        }
        res = comp_mod.merge(prof, comp)
        self.assertTrue(res.ok, res.problems)
        self.assertEqual(prof["comprehension"]["audit"][cid]["verdict"], "PASS")

    def test_bad_audit_verdict_value_rejected(self):
        prof = _docx_profile_with_inventory()
        cid = self._checklist_id(prof)
        comp = _valid_comp()
        comp["audit"] = {cid: {"verdict": "MAYBE"}}
        res = comp_mod.merge(prof, comp)
        self.assertFalse(res.ok)
        self.assertTrue(any("verdict" in p for p in res.problems), res.problems)

    def test_audit_into_empty_checklist_is_error(self):
        # An audit key when the derived checklist is EMPTY has nothing to bind to ->
        # ERROR (reject-never-skip), identical to anchor/index/region refs.
        from unittest.mock import patch

        prof = _docx_profile_with_inventory()
        comp = _valid_comp()
        comp["audit"] = {"any-id": {"verdict": "PASS"}}
        with patch("brandkit.qa.visual.visual_checklist_ids", lambda p: []):
            res = comp_mod.merge(prof, comp)
        self.assertFalse(res.ok)
        self.assertTrue(any("any-id" in p for p in res.problems), res.problems)

    def test_audit_survives_merge_round_trip_byte_identical(self):
        # Guards the _canonicalize audit arm: a key omitted there is silently
        # dropped, and a re-merge must be byte-identical (idempotency).
        prof_a = _docx_profile_with_inventory()
        prof_b = _docx_profile_with_inventory()
        cid_a = self._checklist_id(prof_a)
        comp = _valid_comp()
        comp["audit"] = {cid_a: {"verdict": "PASS", "evidence": "ok"}}
        res_a = comp_mod.merge(prof_a, comp)
        self.assertTrue(res_a.ok, res_a.problems)
        # The verdict survived canonicalization.
        self.assertIn(cid_a, prof_a["comprehension"]["audit"])
        # Two merges of the same proposal are byte-identical, and a re-merge stable.
        comp_mod.merge(
            prof_b, dict(comp, audit={cid_a: {"verdict": "PASS", "evidence": "ok"}})
        )
        self.assertEqual(
            json.dumps(prof_a["comprehension"], sort_keys=True),
            json.dumps(prof_b["comprehension"], sort_keys=True),
        )
        before = json.dumps(prof_a["comprehension"], sort_keys=True)
        comp_mod.merge(
            prof_a, dict(comp, audit={cid_a: {"verdict": "PASS", "evidence": "ok"}})
        )
        self.assertEqual(before, json.dumps(prof_a["comprehension"], sort_keys=True))

    def test_audit_targets_invariant_declared(self):
        self.assertIn("audit_targets_exist", schema.DEFAULT_L0_INVARIANTS)

    def test_audit_nonmember_fails_the_gate_through_run_qa(self):
        # The gate independently rejects a hand-crafted present block with a bad
        # audit key, attributing it to its own ``audit_targets_exist`` id.
        prof = _docx_profile_with_inventory()
        prof["comprehension"] = {
            "schema_version": "comprehension-1",
            "status": "present",
            "source_shell_sha256": "abc123",
            "confidence": 0.5,
            "cover_slots": {},
            "conventions": {"indexes": [], "sections": []},
            "role_annotations": {},
            "demo_classification": {"regions": []},
            "audit": {"ghost-id": {"verdict": "PASS"}},
        }
        report = run_qa(None, prof, qa="fast", shell=None)
        self.assertFalse(report.passed)
        audit_errs = [f for f in report.findings if f.check == "audit_targets_exist"]
        self.assertTrue(audit_errs, [f.message for f in report.findings])
        self.assertEqual(audit_errs[0].severity, "ERROR")
        # EXCLUSIVITY: the bad audit key is attributed ONLY to audit_targets_exist -
        # check_comprehension_targets skips comprehension.audit problems, so the same
        # key never double-reports under comprehension_targets_exist.
        comp_errs = [
            f
            for f in report.findings
            if f.check == "comprehension_targets_exist"
            and "comprehension.audit" in f.message
        ]
        self.assertEqual(comp_errs, [], [f.message for f in report.findings])

    def test_absent_audit_is_byte_identical(self):
        # An old profile with no `audit`/`triage` key round-trips merge unchanged.
        prof = _docx_profile_with_inventory()
        comp = _valid_comp()  # no audit/triage key
        res = comp_mod.merge(prof, comp)
        self.assertTrue(res.ok, res.problems)
        self.assertEqual(prof["comprehension"]["audit"], {})
        self.assertEqual(prof["comprehension"]["triage"], [])


class TriageSinkTest(unittest.TestCase):
    """C2: the model-assisted QA-triage list (`comprehension.triage`)."""

    def test_triage_targets_invariant_declared(self):
        self.assertIn("triage_targets_exist", schema.DEFAULT_L0_INVARIANTS)

    def test_triage_eligible_set_is_closed(self):
        # The eligible set is exactly the three WARNING-only ambiguous checks, and
        # NONE of them is an ERROR-emitting check (the merge-side belt that makes an
        # ERROR-aimed triage entry impossible).
        self.assertEqual(
            schema.AMBIGUOUS_TRIAGE_CHECKS,
            frozenset({"visual.blank_page", "visual.edge_bleed", "component_survival"}),
        )

    def test_triage_dispositions_are_closed(self):
        self.assertEqual(schema.TRIAGE_DISPOSITIONS, frozenset({"expected", "defect"}))

    def test_valid_triage_entry_merges(self):
        prof = _docx_profile_with_inventory()
        comp = _valid_comp()
        comp["triage"] = [
            {
                "check": "visual.edge_bleed",
                "location": "page:1:bottom",
                "disposition": "expected",
                "evidence": "full-bleed cover by design",
            }
        ]
        res = comp_mod.merge(prof, comp)
        self.assertTrue(res.ok, res.problems)
        self.assertEqual(len(prof["comprehension"]["triage"]), 1)
        self.assertEqual(prof["comprehension"]["triage"][0]["disposition"], "expected")

    def test_triage_entry_naming_error_check_rejected_at_merge(self):
        # A triage entry naming an ERROR-emitting check (NOT in the eligible set) is
        # fail-closed rejected at merge: it can never reach the gate to demote an ERROR.
        prof = _docx_profile_with_inventory()
        comp = _valid_comp()
        comp["triage"] = [
            {
                "check": "no_residual_template_text",  # an ERROR check, not eligible
                "location": None,
                "disposition": "expected",
            }
        ]
        res = comp_mod.merge(prof, comp)
        self.assertFalse(res.ok)
        self.assertEqual(prof["comprehension"]["status"], "rejected")
        self.assertTrue(
            any("no_residual_template_text" in p for p in res.problems), res.problems
        )

    def test_unknown_disposition_rejected(self):
        prof = _docx_profile_with_inventory()
        comp = _valid_comp()
        comp["triage"] = [
            {
                "check": "visual.blank_page",
                "location": "page:2",
                "disposition": "ignore",  # not in the closed enum
            }
        ]
        res = comp_mod.merge(prof, comp)
        self.assertFalse(res.ok)
        self.assertTrue(any("disposition" in p for p in res.problems), res.problems)

    def test_duplicate_check_location_rejected(self):
        prof = _docx_profile_with_inventory()
        comp = _valid_comp()
        comp["triage"] = [
            {
                "check": "component_survival",
                "location": "tables",
                "disposition": "expected",
            },
            {
                "check": "component_survival",
                "location": "tables",
                "disposition": "defect",
            },
        ]
        res = comp_mod.merge(prof, comp)
        self.assertFalse(res.ok)
        self.assertTrue(
            any("duplicate" in p.lower() for p in res.problems), res.problems
        )

    def test_triage_survives_merge_round_trip_byte_identical(self):
        # Guards the _canonicalize triage arm: omit it and the triage is silently
        # dropped. A re-merge must be byte-identical (idempotency).
        prof_a = _docx_profile_with_inventory()
        prof_b = _docx_profile_with_inventory()
        entry = {
            "check": "visual.edge_bleed",
            "location": "page:1:bottom",
            "disposition": "expected",
            "evidence": "ok",
        }
        comp = dict(_valid_comp(), triage=[dict(entry)])
        res_a = comp_mod.merge(prof_a, comp)
        self.assertTrue(res_a.ok, res_a.problems)
        self.assertEqual(len(prof_a["comprehension"]["triage"]), 1)
        comp_mod.merge(prof_b, dict(_valid_comp(), triage=[dict(entry)]))
        self.assertEqual(
            json.dumps(prof_a["comprehension"], sort_keys=True),
            json.dumps(prof_b["comprehension"], sort_keys=True),
        )
        before = json.dumps(prof_a["comprehension"], sort_keys=True)
        comp_mod.merge(prof_a, dict(_valid_comp(), triage=[dict(entry)]))
        self.assertEqual(before, json.dumps(prof_a["comprehension"], sort_keys=True))

    def test_triage_entry_naming_error_check_fails_the_gate(self):
        # The gate independently rejects a hand-crafted present block whose triage
        # names a non-eligible check, attributing it to ``triage_targets_exist``.
        prof = _docx_profile_with_inventory()
        prof["comprehension"] = {
            "schema_version": "comprehension-1",
            "status": "present",
            "source_shell_sha256": "abc123",
            "confidence": 0.5,
            "cover_slots": {},
            "conventions": {"indexes": [], "sections": []},
            "role_annotations": {},
            "demo_classification": {"regions": []},
            "triage": [
                {"check": "no_residual_template_text", "disposition": "expected"}
            ],
        }
        report = run_qa(None, prof, qa="fast", shell=None)
        self.assertFalse(report.passed)
        self.assertTrue(
            any(
                f.check == "triage_targets_exist" and f.severity == "ERROR"
                for f in report.findings
            ),
            [f.message for f in report.findings],
        )

    def test_absent_triage_is_byte_identical(self):
        prof = _docx_profile_with_inventory()
        res = comp_mod.merge(prof, _valid_comp())  # no triage key
        self.assertTrue(res.ok, res.problems)
        self.assertEqual(prof["comprehension"]["triage"], [])


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


def _palette_entry(theme_slot):
    """The deterministic palette entry shape (model fields null)."""
    return {
        "ref": {"kind": "theme", "theme": theme_slot},
        "provenance": [],
        "frequency": "rare",
        "name": None,
        "purpose": None,
        "use_when": None,
    }


class PaletteAnnotationsTest(unittest.TestCase):
    """Model-driven color: the comprehension surfaces a 'palette' inventory, binds
    palette_annotations fail-closed against it, and on a clean merge mirrors the
    model's NAMES onto theme.palette without ever touching the captured ref."""

    def _profile_with_palette(self):
        prof = _docx_profile_with_inventory()
        prof["theme"]["palette"] = {
            "accent1": _palette_entry("accent1"),
            "dk1": _palette_entry("dk1"),
        }
        return prof

    def test_palette_inventory_is_surfaced(self):
        prof = self._profile_with_palette()
        inv = comp_mod.surface_inventories(prof)
        self.assertEqual(inv["palette"], ["accent1", "dk1"])

    def test_palette_facts_in_bundle_carry_ref_and_null_name(self):
        prof = self._profile_with_palette()
        bundle = comp_mod.comprehend_input_bundle(prof)
        palette_facts = bundle["facts"]["palette"]
        self.assertEqual([p["key"] for p in palette_facts], ["accent1", "dk1"])
        self.assertEqual(palette_facts[0]["ref"], {"kind": "theme", "theme": "accent1"})
        # The model never receives a name in the deterministic path.
        self.assertIsNone(palette_facts[0]["name"])

    def test_annotation_binds_and_mirrors_onto_palette(self):
        prof = self._profile_with_palette()
        comp = {
            "palette_annotations": {
                "accent1": {
                    "name": "primary brand",
                    "purpose": "headings",
                    "use_when": "section titles",
                    "semantic_role": "accent",
                }
            }
        }
        res = comp_mod.merge(prof, comp)
        self.assertTrue(res.ok, res.problems)
        entry = prof["theme"]["palette"]["accent1"]
        # Names mirrored onto the deterministic entry...
        self.assertEqual(entry["name"], "primary brand")
        self.assertEqual(entry["purpose"], "headings")
        self.assertEqual(entry["use_when"], "section titles")
        self.assertEqual(entry["semantic_role"], "accent")
        # ...without ever touching the captured ref.
        self.assertEqual(entry["ref"], {"kind": "theme", "theme": "accent1"})
        # The canonical comprehension also carries the annotation.
        self.assertIn("accent1", prof["comprehension"]["palette_annotations"])

    def test_annotation_key_absent_from_palette_is_rejected(self):
        prof = self._profile_with_palette()
        res = comp_mod.merge(
            prof, {"palette_annotations": {"accent9": {"name": "ghost"}}}
        )
        self.assertFalse(res.ok)
        self.assertEqual(prof["comprehension"]["status"], "rejected")
        self.assertTrue(
            any("accent9" in p and "palette" in p for p in res.problems), res.problems
        )

    def test_annotation_into_empty_palette_is_error_not_skipped(self):
        # Fail-closed on empty, same rule as anchor/index/region.
        prof = _docx_profile_with_inventory()  # no theme.palette set -> empty
        res = comp_mod.merge(prof, {"palette_annotations": {"accent1": {"name": "x"}}})
        self.assertFalse(res.ok)
        self.assertTrue(
            any("accent1" in p and "palette" in p for p in res.problems), res.problems
        )

    def test_palette_annotation_merge_is_idempotent(self):
        prof_a = self._profile_with_palette()
        prof_b = self._profile_with_palette()
        comp = {"palette_annotations": {"accent1": {"name": "primary"}}}
        comp_mod.merge(prof_a, dict(comp))
        comp_mod.merge(prof_b, dict(comp))
        self.assertEqual(
            json.dumps(prof_a["comprehension"], sort_keys=True),
            json.dumps(prof_b["comprehension"], sort_keys=True),
        )
        self.assertEqual(
            json.dumps(prof_a["theme"]["palette"], sort_keys=True),
            json.dumps(prof_b["theme"]["palette"], sort_keys=True),
        )

    def test_empty_comprehension_has_palette_annotations_default(self):
        comp = schema.empty_comprehension()
        self.assertEqual(comp["palette_annotations"], {})

    def test_bad_annotation_shape_is_reported(self):
        prof = self._profile_with_palette()
        prof["comprehension"]["status"] = "present"
        prof["comprehension"]["palette_annotations"] = {"accent1": {"name": 123}}
        problems = schema.validate(prof)
        self.assertTrue(any("palette_annotations" in p for p in problems), problems)


class OverlayRefinementTest(unittest.TestCase):
    """C3: the pure overlay primitive (delta over the EXISTING sinks, closed-key)."""

    def test_overlay_preserves_existing_sinks(self):
        # The overlay trap: a raw delta would wipe the other sinks; overlay must keep
        # role_annotations/conventions etc. that the delta does not mention.
        prof = _docx_profile_with_inventory()
        comp_mod.merge(prof, _valid_comp())
        existing = prof["comprehension"]
        delta = {"role_annotations": {"cover.title": {"purpose": "cover headline"}}}
        out = comp_mod.overlay_refinement(existing, delta)
        # The delta's new role_annotation is present...
        self.assertIn("cover.title", out["role_annotations"])
        # ...and the pre-existing caption annotation was NOT dropped.
        self.assertIn("caption", out["role_annotations"])
        # Untouched sinks survive verbatim.
        self.assertEqual(out["cover_slots"], existing["cover_slots"])
        self.assertEqual(out["conventions"], existing["conventions"])

    def test_overlay_is_pure_does_not_mutate_inputs(self):
        existing = {"role_annotations": {"caption": {"purpose": "old"}}}
        existing_snapshot = json.dumps(existing, sort_keys=True)
        comp_mod.overlay_refinement(
            existing, {"role_annotations": {"caption": {"purpose": "new"}}}
        )
        self.assertEqual(json.dumps(existing, sort_keys=True), existing_snapshot)

    def test_overlay_map_sink_replaces_matching_key(self):
        existing = {"role_annotations": {"caption": {"purpose": "old"}}}
        out = comp_mod.overlay_refinement(
            existing, {"role_annotations": {"caption": {"purpose": "new"}}}
        )
        self.assertEqual(out["role_annotations"]["caption"]["purpose"], "new")

    def test_overlay_lists_merge_by_ref_no_dup(self):
        # demo_classification.regions and conventions.* merge BY ref, never concat.
        existing = {
            "demo_classification": {
                "regions": [{"region_ref": "body.demo", "verdict": "demo"}]
            },
            "conventions": {
                "indexes": [{"index_ref": "tot.1", "reconcile": "regenerate"}],
                "sections": [{"region_ref": "body.demo", "required": False}],
            },
        }
        delta = {
            "demo_classification": {
                "regions": [{"region_ref": "body.demo", "verdict": "real"}]
            },
            "conventions": {
                "sections": [{"region_ref": "body.demo", "required": True}],
            },
        }
        out = comp_mod.overlay_refinement(existing, delta)
        # Region replaced in place, not duplicated.
        regs = out["demo_classification"]["regions"]
        self.assertEqual(len(regs), 1)
        self.assertEqual(regs[0]["verdict"], "real")
        # Section replaced in place, not duplicated.
        secs = out["conventions"]["sections"]
        self.assertEqual(len(secs), 1)
        self.assertTrue(secs[0]["required"])
        # The index the delta did not mention survives.
        self.assertEqual(len(out["conventions"]["indexes"]), 1)

    def test_overlay_appends_new_ref(self):
        existing = {
            "demo_classification": {
                "regions": [{"region_ref": "body.demo", "verdict": "demo"}]
            }
        }
        out = comp_mod.overlay_refinement(
            existing,
            {
                "demo_classification": {
                    "regions": [{"region_ref": "body.new", "verdict": "real"}]
                }
            },
        )
        refs = {r["region_ref"] for r in out["demo_classification"]["regions"]}
        self.assertEqual(refs, {"body.demo", "body.new"})

    def test_overlay_ignores_unknown_and_structural_keys(self):
        # A delta cannot smuggle a new field or shadow a structural/QA-verdict one.
        existing = {"role_annotations": {"caption": {"purpose": "x"}}}
        out = comp_mod.overlay_refinement(
            existing,
            {
                "source_shell_sha256": "evil",
                "status": "present",
                "audit": {"x": {"verdict": "PASS"}},
                "triage": [{"check": "visual.blank_page", "disposition": "expected"}],
                "fragments": [{"ref": "z", "kind": "component"}],
                "bogus_field": 1,
            },
        )
        # None of those keys leaked through the overlay.
        self.assertNotIn("source_shell_sha256", out)
        self.assertNotIn("bogus_field", out)
        self.assertNotIn("audit", out)
        self.assertNotIn("triage", out)
        self.assertNotIn("fragments", out)
        # The pre-existing sink is untouched.
        self.assertEqual(out["role_annotations"], existing["role_annotations"])


class CliRefineTest(unittest.TestCase):
    """C3: the ``refine`` CLI verb routes the overlaid delta through merge."""

    def _extracted(self, tmp):
        import os
        from brandkit.cli import main

        sys.path.insert(0, str(Path(__file__).resolve().parents[0]))
        from test_smoke import _synthetic_template

        template = tmp / "t.docx"
        _synthetic_template(template)
        rc = main(
            [
                "extract",
                "--name",
                "acme",
                "--template",
                str(template),
                "--scope",
                "project",
            ]
        )
        self.assertEqual(rc, 0)
        return os, main

    def _surfaced_role(self, tmp):
        prof = json.loads((tmp / "brand-kit" / "acme" / "profile.json").read_text())
        roles = schema.list_role_ids(prof)
        self.assertTrue(roles, "test profile must surface at least one role")
        return roles[0]

    def test_refine_without_accept_does_not_persist_live(self):
        import os

        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            old = Path.cwd()
            os.chdir(tmp)
            try:
                _os, main = self._extracted(tmp)
                rid = self._surfaced_role(tmp)
                ref = tmp / "ref.json"
                ref.write_text(
                    json.dumps({"role_annotations": {rid: {"purpose": "feedback"}}}),
                    encoding="utf-8",
                )
                # No --accept: returns 0 but the on-disk block stays absent (prior).
                self.assertEqual(
                    main(
                        [
                            "refine",
                            "--name",
                            "acme",
                            "--input",
                            str(ref),
                            "--scope",
                            "project",
                        ]
                    ),
                    0,
                )
                prof = json.loads(
                    (tmp / "brand-kit" / "acme" / "profile.json").read_text()
                )
                self.assertEqual(prof["comprehension"]["status"], "absent")
                self.assertEqual(prof["comprehension"]["role_annotations"], {})
                # --accept persists the refined present block.
                self.assertEqual(
                    main(
                        [
                            "refine",
                            "--name",
                            "acme",
                            "--input",
                            str(ref),
                            "--scope",
                            "project",
                            "--accept",
                        ]
                    ),
                    0,
                )
                prof2 = json.loads(
                    (tmp / "brand-kit" / "acme" / "profile.json").read_text()
                )
                self.assertEqual(prof2["comprehension"]["status"], "present")
                self.assertEqual(
                    prof2["comprehension"]["role_annotations"][rid]["purpose"],
                    "feedback",
                )
                self.assertTrue(prof2["comprehension"]["source_shell_sha256"])
            finally:
                os.chdir(old)

    def test_refine_delta_binds_only_surfaced_ids(self):
        import os

        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            old = Path.cwd()
            os.chdir(tmp)
            try:
                _os, main = self._extracted(tmp)
                ref = tmp / "ref.json"
                ref.write_text(
                    json.dumps({"role_annotations": {"ghost.role": {"purpose": "x"}}}),
                    encoding="utf-8",
                )
                self.assertEqual(
                    main(
                        [
                            "refine",
                            "--name",
                            "acme",
                            "--input",
                            str(ref),
                            "--scope",
                            "project",
                            "--accept",
                        ]
                    ),
                    1,
                )
                # All-or-nothing: a rejected refinement leaves the prior block untouched
                # (still absent), never a half-written present block.
                prof = json.loads(
                    (tmp / "brand-kit" / "acme" / "profile.json").read_text()
                )
                self.assertEqual(prof["comprehension"]["status"], "absent")
            finally:
                os.chdir(old)

    def test_refine_merges_as_delta_over_existing_sinks(self):
        # A second refinement of a DIFFERENT sink key must not drop the first one:
        # the overlay is a per-sink delta over the EXISTING block, not a replace-all
        # (which is what passing the raw delta to merge would do - the overlay trap).
        import os

        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            old = Path.cwd()
            os.chdir(tmp)
            try:
                _os, main = self._extracted(tmp)
                prof0 = json.loads(
                    (tmp / "brand-kit" / "acme" / "profile.json").read_text()
                )
                roles = schema.list_role_ids(prof0)
                self.assertGreaterEqual(
                    len(roles), 2, "test needs two surfaced roles for this case"
                )
                rid_a, rid_b = roles[0], roles[1]
                ref1 = tmp / "ref1.json"
                ref1.write_text(
                    json.dumps({"role_annotations": {rid_a: {"purpose": "first"}}}),
                    encoding="utf-8",
                )
                self.assertEqual(
                    main(
                        [
                            "refine",
                            "--name",
                            "acme",
                            "--input",
                            str(ref1),
                            "--scope",
                            "project",
                            "--accept",
                        ]
                    ),
                    0,
                )
                # A second refinement annotates a DIFFERENT role; the overlay must add
                # it WITHOUT dropping the first role's annotation.
                ref2 = tmp / "ref2.json"
                ref2.write_text(
                    json.dumps({"role_annotations": {rid_b: {"purpose": "second"}}}),
                    encoding="utf-8",
                )
                self.assertEqual(
                    main(
                        [
                            "refine",
                            "--name",
                            "acme",
                            "--input",
                            str(ref2),
                            "--scope",
                            "project",
                            "--accept",
                        ]
                    ),
                    0,
                )
                prof = json.loads(
                    (tmp / "brand-kit" / "acme" / "profile.json").read_text()
                )
                ann = prof["comprehension"]["role_annotations"]
                # Both refinements coexist (delta-over-existing, not replace-all).
                self.assertEqual(ann[rid_a]["purpose"], "first")
                self.assertEqual(ann[rid_b]["purpose"], "second")
            finally:
                os.chdir(old)


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
