# SPDX-License-Identifier: MIT
"""Regression tests for the PPTX correctness + brand-guarantee fixes.

Covers the confirmed findings:
  - C3  the extractor DERIVES roles from the real parsed layouts: every
        resolver points at a layout.name that actually exists in the deck, and
        when no suitable layout exists the role is an honest ``stub`` (no
        fabricated ``"Title Slide"`` / ``"Title and Content"``). The legacy
        literal ``surface.pptx.role_layout_map`` is gone.
  - M6  generation resolves layouts from the profile's real role data, never a
        hardcoded layout name or positional ``slide_layouts[0]/[1]`` fallback.
  - M7  slides are built from the IR block stream — one slide per heading (its
        own runs as the title), following blocks as that slide's body — with no
        flattening and without dropping tables / quotes / captions / lists.

All decks are synthesized in a temp dir with python-pptx and never committed.
"""
from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from pptx import Presentation
from pptx.util import Inches

from brandkit.formats.pptx import extract as px
from brandkit.formats.pptx import generate as pg
from brandkit.ir.model import parse_idoc
from brandkit.profile import schema


# ---------------------------------------------------------------------------
# Synthetic deck builders (temp only — NEVER committed)
# ---------------------------------------------------------------------------
def _branded_template(path: Path) -> None:
    """A deck whose layouts carry NON-default names, to prove the extractor and
    generator never depend on literal ``"Title Slide"`` / ``"Title and Content"``.

    python-pptx's default index-0 layout has a CENTER_TITLE + SUBTITLE (a cover)
    and index-1 has a TITLE + OBJECT body (a content layout); we only rename them.
    """
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    prs.slide_layouts[0].name = "BrandCover"
    prs.slide_layouts[1].name = "BrandContent"
    prs.save(path)


def _placeholderless_template(path: Path) -> None:
    """A deck whose layouts expose NO title/body placeholders at all (forces the
    honest ``stub`` path)."""
    prs = Presentation()
    for layout in prs.slide_layouts:
        sp_tree = layout.shapes._spTree
        for ph in list(layout.placeholders):
            sp_tree.remove(ph._element)
    prs.save(path)


def _extract_profile(template: Path, name: str = "deck") -> dict:
    """Build a profile the way ``extract`` does, in-memory (no disk store)."""
    prs = Presentation(template)
    layouts = px._layouts(prs)
    roles = px._roles(prs, layouts)
    profile = schema.build_envelope(
        "pptx",
        {"name": name, "display_name": name},
        theme=px._theme(),
        roles=roles,
        surface={
            "pptx": {
                "slide_size_emu": {"w": int(prs.slide_width), "h": int(prs.slide_height)},
                "layouts": layouts,
                "safe_area_emu": {"l": 457200, "t": 457200, "r": 457200, "b": 457200},
            }
        },
    )
    profile["capabilities"] = px._capabilities()
    return profile


def _slide_body(slide) -> str:
    for shape in slide.placeholders:
        if shape != slide.shapes.title and shape.has_text_frame:
            return shape.text
    return ""


def _all_text(prs: Presentation) -> str:
    return "\n".join(
        shape.text
        for slide in prs.slides
        for shape in slide.placeholders
        if shape.has_text_frame
    )


# ---------------------------------------------------------------------------
# C3 — roles derived from the REAL deck, not fabricated
# ---------------------------------------------------------------------------
class C3RolesFromRealLayouts(unittest.TestCase):
    def test_resolver_layouts_actually_exist_in_deck(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            template = Path(td) / "branded.pptx"
            _branded_template(template)
            profile = _extract_profile(template)

            layout_names = set(profile["surface"]["pptx"]["layouts"])
            roles = profile["roles"]
            for rid in roles["_index"]:
                layout = roles[rid]["resolver"].get("layout")
                # Every non-stub role names a layout the deck actually has.
                if layout is not None:
                    self.assertIn(
                        layout,
                        layout_names,
                        f"role {rid} points at fabricated layout {layout!r}",
                    )

            # The specific renamed real layouts are the ones chosen.
            self.assertEqual(roles["cover.title"]["resolver"]["layout"], "BrandCover")
            self.assertEqual(roles["heading.1"]["resolver"]["layout"], "BrandContent")
            self.assertEqual(roles["paragraph"]["resolver"]["layout"], "BrandContent")

    def test_no_fabricated_literal_names(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            template = Path(td) / "branded.pptx"
            _branded_template(template)
            profile = _extract_profile(template)

            roles = profile["roles"]
            for rid in roles["_index"]:
                layout = roles[rid]["resolver"].get("layout")
                self.assertNotIn(layout, {"Title Slide", "Title and Content"})

            # The legacy literal map must be gone from the surface.
            self.assertNotIn("role_layout_map", profile["surface"]["pptx"])

    def test_profile_validates_clean(self) -> None:
        # The intra-profile consistency check (Core) would flag a fabricated
        # layout; a derived profile must pass with zero problems.
        with tempfile.TemporaryDirectory() as td:
            template = Path(td) / "branded.pptx"
            _branded_template(template)
            profile = _extract_profile(template)
            self.assertEqual(schema.validate(profile), [])

    def test_placeholder_idx_is_real(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            template = Path(td) / "branded.pptx"
            _branded_template(template)
            profile = _extract_profile(template)
            layouts = profile["surface"]["pptx"]["layouts"]
            roles = profile["roles"]
            for rid in roles["_index"]:
                resolver = roles[rid]["resolver"]
                layout = resolver.get("layout")
                ph_idx = resolver.get("ph_idx")
                if layout is None:
                    continue
                idxs = {ph["idx"] for ph in layouts[layout]["placeholders"]}
                self.assertIn(ph_idx, idxs, f"{rid}: ph_idx {ph_idx} absent from {layout}")

    def test_honest_stub_when_no_suitable_layout(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            template = Path(td) / "bare.pptx"
            _placeholderless_template(template)
            profile = _extract_profile(template)

            roles = profile["roles"]
            for rid in ("cover.title", "heading.1", "paragraph"):
                self.assertEqual(roles[rid]["status"], schema.Status.STUB.value)
                self.assertIsNone(roles[rid]["resolver"]["layout"])
                self.assertEqual(roles[rid]["confidence"], 0.0)
                self.assertFalse(roles[rid]["verified"])

            # A stub profile still validates (no fabricated target to contradict).
            self.assertEqual(schema.validate(profile), [])

    def test_real_extract_reports_cover_anchor_honestly(self) -> None:
        # Drive the real on-disk extract() to confirm anchors track reality:
        # present for a deck with a cover layout, absent for a placeholder-less one.
        import json

        for builder, expect_present in ((_branded_template, True), (_placeholderless_template, False)):
            with tempfile.TemporaryDirectory() as td:
                tp = Path(td)
                template = tp / "tpl.pptx"
                builder(template)
                profile_path = px.extract(template, "deck", scope="project", cwd=tp)
                profile = json.loads(Path(profile_path).read_text())
                cover = profile["anchors"]["cover"]
                self.assertEqual(cover["slots_found"], 1 if expect_present else 0)
                self.assertEqual(
                    cover["kind"],
                    schema.AnchorKind.PLACEHOLDER.value if expect_present else schema.AnchorKind.NONE.value,
                )
                # The on-disk profile must also validate clean either way.
                self.assertEqual(schema.validate(profile), [])


# ---------------------------------------------------------------------------
# M6 / M7 — generation uses real layouts + the IR block stream
# ---------------------------------------------------------------------------
class M6M7GenerateFromRealLayouts(unittest.TestCase):
    def _generate(self, template: Path, idoc_dict: dict, out: Path):
        profile = _extract_profile(template)
        idoc = parse_idoc(idoc_dict)
        pg.generate(profile, template, idoc, out)
        return Presentation(out)

    def test_generation_uses_only_real_layouts(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            tp = Path(td)
            template = tp / "branded.pptx"
            _branded_template(template)
            out = tp / "out.pptx"
            res = self._generate(
                template,
                {
                    "cover": {"title": "Board Update", "subtitle": "Q2"},
                    "blocks": [
                        {"type": "heading", "level": 1, "text": "Section A"},
                        {"type": "paragraph", "text": "Body."},
                    ],
                },
                out,
            )
            used = {s.slide_layout.name for s in res.slides}
            # Only the real renamed layouts — never a fabricated literal.
            self.assertTrue(used <= {"BrandCover", "BrandContent"}, used)
            self.assertNotIn("Title Slide", used)
            self.assertNotIn("Title and Content", used)

    def test_one_slide_per_heading_with_distinct_titles(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            tp = Path(td)
            template = tp / "branded.pptx"
            _branded_template(template)
            out = tp / "out.pptx"
            res = self._generate(
                template,
                {
                    "cover": {"title": "Deck"},
                    "blocks": [
                        {"type": "heading", "level": 1, "text": "Section A"},
                        {"type": "paragraph", "text": "Alpha."},
                        {"type": "heading", "level": 1, "text": "Section B"},
                        {"type": "paragraph", "text": "Beta."},
                    ],
                },
                out,
            )
            titles = [s.shapes.title.text for s in res.slides if s.shapes.title and s.shapes.title.text]
            # Both headings are their OWN slide titles (not "Section A (2)").
            self.assertIn("Section A", titles)
            self.assertIn("Section B", titles)

    def test_heading_text_not_duplicated_into_body(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            tp = Path(td)
            template = tp / "branded.pptx"
            _branded_template(template)
            out = tp / "out.pptx"
            res = self._generate(
                template,
                {
                    "blocks": [
                        {"type": "heading", "level": 1, "text": "Section A"},
                        {"type": "paragraph", "text": "Just the body here."},
                    ]
                },
                out,
            )
            for slide in res.slides:
                body = _slide_body(slide)
                self.assertNotIn("Section A", body)

    def test_tables_quotes_captions_lists_not_dropped(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            tp = Path(td)
            template = tp / "branded.pptx"
            _branded_template(template)
            out = tp / "out.pptx"
            res = self._generate(
                template,
                {
                    "blocks": [
                        {"type": "heading", "level": 1, "text": "Mixed"},
                        {"type": "list", "items": [{"text": "first bullet"}, {"text": "second bullet"}]},
                        {"type": "table", "columns": ["Area", "Status"], "rows": [["Pipeline", "Healthy"]]},
                        {"type": "quote", "text": "A pithy remark.", "attribution": "Anon"},
                        {"type": "caption", "text": "Figure 1. A caption."},
                        {"type": "callout", "intent": "info", "text": "Mind the gap."},
                    ]
                },
                out,
            )
            text = _all_text(res)
            for needle in (
                "first bullet",
                "second bullet",
                "Pipeline",
                "Healthy",
                "pithy remark",
                "Anon",
                "Figure 1",
                "Mind the gap",
            ):
                self.assertIn(needle, text, f"content {needle!r} was dropped")

    def test_long_section_splits_into_continuation_slides(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            tp = Path(td)
            template = tp / "branded.pptx"
            _branded_template(template)
            out = tp / "out.pptx"
            res = self._generate(
                template,
                {
                    "cover": {"title": "Long"},
                    "blocks": [
                        {"type": "heading", "level": 1, "text": "Long Section"},
                        {"type": "paragraph", "text": " ".join(["capacity"] * 420)},
                    ],
                },
                out,
            )
            # cover + at least two content slides for the over-capacity paragraph.
            self.assertGreaterEqual(len(res.slides), 3)
            content_titles = [
                s.shapes.title.text for s in res.slides if s.shapes.title and s.shapes.title.text
            ]
            # The continuation slide carries the SAME section title (suffixed).
            self.assertIn("Long Section", content_titles)
            self.assertTrue(any(t.startswith("Long Section (") for t in content_titles))

    def test_no_cover_slide_when_idoc_has_no_cover(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            tp = Path(td)
            template = tp / "branded.pptx"
            _branded_template(template)
            out = tp / "out.pptx"
            res = self._generate(
                template,
                {"blocks": [{"type": "heading", "level": 1, "text": "Only Content"}]},
                out,
            )
            self.assertTrue(all(s.slide_layout.name != "BrandCover" for s in res.slides))


if __name__ == "__main__":  # pragma: no cover
    unittest.main()
