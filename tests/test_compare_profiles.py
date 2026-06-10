# SPDX-License-Identifier: MIT
"""The ``compare-profiles`` verb - read-only cross-template drift detection.

Unit tests pin the drift semantics of the pure comparator (what counts as
DRIFT vs information), and a CLI smoke test proves the verb is genuinely
read-only and deterministic against two really-extracted example profiles.
"""

from __future__ import annotations

import hashlib
import io
import json
import os
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from brandkit.profile import compare  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]


def _profile(
    *,
    kind: str = "docx",
    colors: dict | None = None,
    fonts: dict | None = None,
    palette: dict | None = None,
    palette_roles: dict | None = None,
    roles: list[str] | None = None,
) -> dict:
    return {
        "$schema": "https://brand-docs/schema/profile-1.json",
        "identity": {"name": "acme", "locale": "en-US"},
        "surface": {kind: {}},
        "theme": {
            "colors": {k: {"hex": v} for k, v in (colors or {}).items()},
            "fonts": fonts or {},
            "palette": palette or {},
            "palette_roles": {
                k: {"theme": v} for k, v in (palette_roles or {}).items()
            },
        },
        "roles": {"_index": roles or []},
    }


class CompareSemanticsTest(unittest.TestCase):
    def test_identical_brand_facts_are_aligned(self):
        a = _profile(
            colors={"accent1": "2B7CD3", "dk1": "16213F"},
            fonts={"major": {"latin": "Arial"}, "minor": {"latin": "Calibri"}},
            palette_roles={"primary": "accent1"},
            roles=["paragraph", "heading.1"],
        )
        b = _profile(
            kind="pptx",
            colors={"accent1": "2B7CD3", "dk1": "16213F"},
            fonts={"major": {"latin": "Arial"}, "minor": {"latin": "Calibri"}},
            palette_roles={"primary": "accent1"},
            roles=["paragraph", "title"],  # different coverage: informational
        )
        result = compare.compare_profiles(a, b)
        self.assertEqual(result["verdict"], compare.VERDICT_ALIGNED)
        self.assertEqual(result["theme_colors"]["differ"], {})
        # Coverage difference is reported but never drift.
        self.assertEqual(result["roles"]["only_a"], ["heading.1"])
        self.assertEqual(result["roles"]["only_b"], ["title"])

    def test_same_slot_different_hex_is_drift(self):
        a = _profile(colors={"accent1": "2B7CD3"})
        b = _profile(colors={"accent1": "C24D2B"})
        result = compare.compare_profiles(a, b)
        self.assertEqual(result["verdict"], compare.VERDICT_DRIFT)
        self.assertIn("accent1", result["theme_colors"]["differ"])

    def test_hex_comparison_is_case_insensitive(self):
        a = _profile(colors={"accent1": "2b7cd3"})
        b = _profile(colors={"accent1": "2B7CD3"})
        result = compare.compare_profiles(a, b)
        self.assertEqual(result["verdict"], compare.VERDICT_ALIGNED)

    def test_slot_present_in_only_one_profile_is_informational(self):
        a = _profile(colors={"accent1": "2B7CD3", "accent2": "E0742B"})
        b = _profile(colors={"accent1": "2B7CD3"})
        result = compare.compare_profiles(a, b)
        self.assertEqual(result["verdict"], compare.VERDICT_ALIGNED)
        self.assertEqual(result["theme_colors"]["only_a"], {"accent2": "E0742B"})

    def test_font_family_disagreement_is_drift(self):
        a = _profile(fonts={"major": {"latin": "Arial"}})
        b = _profile(fonts={"major": {"latin": "Roboto"}})
        result = compare.compare_profiles(a, b)
        self.assertEqual(result["verdict"], compare.VERDICT_DRIFT)
        self.assertIn("major", result["fonts"]["differ"])

    def test_captured_body_size_disagreement_is_drift(self):
        a = _profile(fonts={"body": {"latin": "Calibri", "size_hp": 22}})
        b = _profile(fonts={"body": {"latin": "Calibri", "size_hp": 24}})
        result = compare.compare_profiles(a, b)
        self.assertEqual(result["verdict"], compare.VERDICT_DRIFT)
        self.assertIn("body_size_hp", result["fonts"]["differ"])

    def test_palette_role_rebinding_is_drift(self):
        a = _profile(palette_roles={"primary": "accent1"})
        b = _profile(palette_roles={"primary": "accent2"})
        result = compare.compare_profiles(a, b)
        self.assertEqual(result["verdict"], compare.VERDICT_DRIFT)
        self.assertIn("primary", result["palette_roles"]["differ"])

    def test_raw_hex_in_one_that_is_theme_slot_in_other_is_drift(self):
        # The sharpest signal: the corporate blue is wired through the theme in
        # B but hardcoded as a raw run color in A.
        a = _profile(palette={"hex:2B7CD3": {}})
        b = _profile(colors={"accent1": "2B7CD3"})
        result = compare.compare_profiles(a, b)
        self.assertEqual(result["verdict"], compare.VERDICT_DRIFT)
        self.assertEqual(
            result["off_theme"]["a_raw_is_b_slot"],
            [{"hex": "2B7CD3", "other_slots": "accent1"}],
        )

    def test_raw_hex_unknown_to_the_other_theme_is_not_drift(self):
        a = _profile(palette={"hex:C24D2B": {}})
        b = _profile(colors={"accent1": "2B7CD3"})
        result = compare.compare_profiles(a, b)
        self.assertEqual(result["verdict"], compare.VERDICT_ALIGNED)
        self.assertEqual(result["off_theme"]["a_raw_is_b_slot"], [])

    def test_render_report_names_every_drift(self):
        a = _profile(
            colors={"accent1": "2B7CD3"},
            fonts={"major": {"latin": "Arial"}},
        )
        b = _profile(
            colors={"accent1": "C24D2B"},
            fonts={"major": {"latin": "Roboto"}},
        )
        text = compare.render_report(compare.compare_profiles(a, b))
        self.assertIn("DRIFT accent1", text)
        self.assertIn("DRIFT major", text)
        self.assertIn("verdict: drift_detected", text)


class CompareCliTest(unittest.TestCase):
    """The verb end to end on really-extracted profiles: read-only, stable."""

    def test_cli_compare_is_read_only_and_deterministic(self):
        from brandkit import cli

        docx_t = ROOT / "examples" / "templates" / "branddocs_template.docx"
        pptx_t = ROOT / "examples" / "templates" / "branddocs_template.pptx"
        with tempfile.TemporaryDirectory() as td:
            old = Path.cwd()
            os.chdir(td)
            try:
                for name, tpl in (("cmp-a", docx_t), ("cmp-b", pptx_t)):
                    rc = cli.main(
                        [
                            "extract",
                            "--name",
                            name,
                            "--template",
                            str(tpl),
                            "--scope",
                            "project",
                        ]
                    )
                    self.assertEqual(rc, 0)
                before = {
                    n: hashlib.sha256(
                        (Path(td) / "brand-kit" / n / "profile.json").read_bytes()
                    ).hexdigest()
                    for n in ("cmp-a", "cmp-b")
                }

                def run(extra: list[str]) -> tuple[int, str]:
                    buf = io.StringIO()
                    with redirect_stdout(buf):
                        rc = cli.main(
                            [
                                "compare-profiles",
                                "--name-a",
                                "cmp-a",
                                "--name-b",
                                "cmp-b",
                                "--scope-a",
                                "project",
                                "--scope-b",
                                "project",
                            ]
                            + extra
                        )
                    return rc, buf.getvalue()

                rc1, out1 = run([])
                rc2, out2 = run([])
                self.assertIn(rc1, (0, 1))
                self.assertEqual((rc1, out1), (rc2, out2))  # deterministic
                self.assertIn("verdict:", out1)
                rcj, outj = run(["--json"])
                parsed = json.loads(outj)
                self.assertEqual(rcj, rc1)
                self.assertIn(
                    parsed["verdict"],
                    (compare.VERDICT_ALIGNED, compare.VERDICT_DRIFT),
                )
                self.assertEqual(parsed["a"]["kind"], "docx")
                self.assertEqual(parsed["b"]["kind"], "pptx")
                after = {
                    n: hashlib.sha256(
                        (Path(td) / "brand-kit" / n / "profile.json").read_bytes()
                    ).hexdigest()
                    for n in ("cmp-a", "cmp-b")
                }
                self.assertEqual(before, after)  # strictly read-only
            finally:
                os.chdir(old)


if __name__ == "__main__":
    unittest.main()
