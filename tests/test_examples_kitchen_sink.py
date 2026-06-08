# SPDX-License-Identifier: MIT
"""End-to-end "kitchen sink" regression against the committed example templates.

Extracts each examples/templates/branddocs_template.{docx,pptx,xlsx}, generates from
a content document that exercises EVERY IR block type plus inline rich-run emphasis
(bold/italic/underline/hyperlink), and runs the QA gate. This is both the loop that
tests the skill against the showcase templates and a permanent regression net:

  * generation never raises and QA never FAILS (degradations are warnings, not errors);
  * inline emphasis + hyperlinks SURVIVE on docx (they used to be silently flattened);
  * `divider` is now a native docx artifact (no longer a block_degraded warning);
  * the still-deferred native writers (kpi/chart/smartart/image) degrade LOUDLY
    (a visible block_degraded finding), never silently.

The templates are 100% synthetic BrandDocs showcases; no proprietary file is used.
"""

from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from brandkit.formats.docx import extract as docx_extract
from brandkit.formats.docx import generate as docx_generate
from brandkit.formats.pptx import extract as pptx_extract
from brandkit.formats.pptx import generate as pptx_generate
from brandkit.formats.xlsx import extract as xlsx_extract
from brandkit.formats.xlsx import generate as xlsx_generate
from brandkit.grid.model import GridDocument
from brandkit.ir.model import parse_idoc
from brandkit.profile import store
from brandkit.qa.gate import run_qa

TEMPLATES = Path(__file__).resolve().parents[1] / "examples" / "templates"

_RICH = [
    {"t": "Plain "},
    {"t": "bold", "b": True},
    {"t": " "},
    {"t": "italic", "i": True},
    {"t": " "},
    {"t": "underline", "u": True},
    {"t": " and a "},
    {"t": "link", "link": "https://example.com"},
    {"t": "."},
]

# Content exercising every block type. component/section are omitted (they require a
# profile-defined fragment registry and otherwise raise fail-closed by design).
KITCHEN = {
    "cover": {
        "title": "Kitchen Sink Report",
        "subtitle": "Every component",
        "fields": {
            "doc_id": "KS-2026-001",
            "date": "2026-06-08",
            "author": "BrandDocs",
        },
    },
    "blocks": [
        {
            "type": "heading",
            "level": 1,
            "runs": [{"t": "Executive "}, {"t": "Summary", "b": True}],
        },
        {"type": "paragraph", "runs": _RICH},
        {"type": "heading", "level": 2, "text": "Details"},
        {
            "type": "list",
            "ordered": False,
            "items": [
                {
                    "runs": [{"t": "First ", "b": True}, {"t": "bullet"}],
                    "level": 0,
                    "items": [{"text": "nested", "level": 1}],
                },
                {"text": "Second"},
            ],
        },
        {
            "type": "list",
            "ordered": True,
            "items": [{"text": "Step one"}, {"text": "Step two"}],
        },
        {
            "type": "callout",
            "intent": "info",
            "title": "Note",
            "runs": [{"t": "Info with "}, {"t": "emphasis", "i": True}],
        },
        {"type": "callout", "intent": "warning", "text": "Careful"},
        {
            "type": "table",
            "columns": ["Area", "Status"],
            "rows": [["Pipeline", "Healthy"], ["Delivery", "Green"]],
            "caption": "Status table",
        },
        {
            "type": "quote",
            "runs": [{"t": "A memorable quote."}],
            "attribution": "Jane Doe",
        },
        {"type": "caption", "target": "figure", "text": "Figure caption"},
        {
            "type": "kpi",
            "items": [
                {"label": "Revenue", "value": "$1.2M", "delta": "+8%"},
                {"label": "Churn", "value": "2.1%", "delta": "-0.3%"},
            ],
        },
        {
            "type": "chart",
            "chart_type": "bar",
            "title": "Revenue",
            "categories": ["Q1", "Q2"],
            "series": [{"name": "Net", "values": [10, 12]}],
        },
        {
            "type": "smartart",
            "diagram": "process",
            "nodes": [{"text": "Plan"}, {"text": "Build"}, {"text": "Ship"}],
        },
        {
            "type": "image",
            "src": "missing.png",
            "alt": "A figure",
            "caption": "An image",
        },
        {"type": "divider"},
        {"type": "toc", "title": "Contents", "max_level": 3},
        {"type": "pagebreak"},
        {"type": "heading", "level": 1, "text": "Appendix"},
        {"type": "paragraph", "text": "End."},
    ],
}

# Block types with no native writer yet: they must degrade LOUDLY (a block_degraded
# finding), never silently. Updated as native writers land.
_DOCX_DEGRADED = {"kpi", "chart", "smartart", "image", "toc"}
_PPTX_DEGRADED = {"kpi", "chart", "smartart", "image", "divider", "toc"}


def _degraded_kinds(findings) -> set[str]:
    kinds = set()
    for f in findings:
        if f.check == "block_degraded":
            # message form: "'<kind>' block ..."
            parts = f.message.split("'")
            if len(parts) > 1:
                kinds.add(parts[1])
    return kinds


class _Base(unittest.TestCase):
    KIND = ""
    EXTRACT = None
    GENERATE = None

    @classmethod
    def setUpClass(cls):
        cls.template = TEMPLATES / f"branddocs_template.{cls.KIND}"
        if not cls.template.exists():
            raise unittest.SkipTest(f"missing example template {cls.template}")

    def _extract(self, td: Path):
        old = os.getcwd()
        os.chdir(td)
        try:
            self.EXTRACT.extract(self.template, "ks", scope="project", cwd=td)
            return store.load_profile("ks", "project")
        finally:
            os.chdir(old)


class DocxKitchenSink(_Base):
    KIND = "docx"
    EXTRACT = docx_extract
    GENERATE = docx_generate

    def test_kitchen_sink_generates_with_inline_and_native_divider(self):
        import copy

        from PIL import Image as PILImage

        with tempfile.TemporaryDirectory() as t:
            td = Path(t)
            loaded = self._extract(td)
            # A real image source so the native image writer places it (the shared
            # KITCHEN points at a missing file to exercise graceful degradation).
            png = td / "fig.png"
            PILImage.new("RGB", (40, 20), (10, 30, 60)).save(png)
            data = copy.deepcopy(KITCHEN)
            for b in data["blocks"]:
                if b.get("type") == "image":
                    b["src"] = str(png)

            out = td / "out.docx"
            sink = []
            docx_generate.generate(
                loaded.profile, loaded.shell_path, parse_idoc(data), out, findings=sink
            )
            self.assertTrue(out.is_file())
            report = run_qa(
                out,
                loaded.profile,
                shell=loaded.shell_path,
                extra_findings=list(sink),
                qa="fast",
            )
            self.assertNotEqual(
                report.verdict,
                "failed",
                [f.message for f in report.findings if f.severity == "ERROR"],
            )

            from docx import Document
            import lxml.etree as ET

            doc = Document(out)
            runs = [r for p in doc.paragraphs for r in p.runs]
            # Inline emphasis survives (was silently flattened before the run-aware writer).
            self.assertTrue(any(r.bold for r in runs), "bold run lost")
            self.assertTrue(any(r.italic for r in runs), "italic run lost")
            self.assertTrue(any(r.underline for r in runs), "underline run lost")
            body_xml = ET.tostring(doc.element.body).decode().lower()
            self.assertIn("hyperlink", body_xml, "hyperlink lost")
            # Divider is native (a paragraph border), image is native (a:blip drawing),
            # KPI is native (a brand table) - none should be a degraded warning.
            self.assertIn("pbdr", body_xml, "native divider rule missing")
            self.assertIn("a:blip", body_xml, "native image not placed")
            degraded = _degraded_kinds(sink)
            self.assertEqual(
                degraded & {"divider", "image", "kpi"},
                set(),
                f"these should be native now: {degraded}",
            )
            # Genuinely-deferred native writers still degrade loudly (never silently).
            self.assertTrue({"chart", "smartart"} <= degraded)


class PptxKitchenSink(_Base):
    KIND = "pptx"
    EXTRACT = pptx_extract
    GENERATE = pptx_generate

    def test_kitchen_sink_generates_and_degrades_loudly(self):
        with tempfile.TemporaryDirectory() as t:
            td = Path(t)
            loaded = self._extract(td)
            out = td / "out.pptx"
            sink = []
            pptx_generate.generate(
                loaded.profile,
                loaded.shell_path,
                parse_idoc(KITCHEN),
                out,
                findings=sink,
            )
            self.assertTrue(out.is_file())
            report = run_qa(
                out,
                loaded.profile,
                shell=loaded.shell_path,
                extra_findings=list(sink),
                qa="fast",
            )
            self.assertNotEqual(
                report.verdict,
                "failed",
                [f.message for f in report.findings if f.severity == "ERROR"],
            )
            degraded = _degraded_kinds(sink)
            # Deferred native writers degrade loudly, never silently.
            self.assertTrue({"kpi", "chart", "smartart", "image"} <= degraded)


class XlsxKitchenSink(_Base):
    KIND = "xlsx"
    EXTRACT = xlsx_extract
    GENERATE = xlsx_generate

    def test_kitchen_sink_fills_named_regions_clean(self):
        with tempfile.TemporaryDirectory() as t:
            td = Path(t)
            loaded = self._extract(td)
            prof = loaded.profile
            regions = ((prof.get("surface") or {}).get("xlsx") or {}).get(
                "named_regions"
            ) or {}
            self.assertTrue(regions, "example xlsx has no named regions to fill")
            cells = {
                n: "X"
                for n, g in regions.items()
                if g.get("cardinality") == "single_cell"
            }
            regs = {
                n: [
                    ["a"] * max(g.get("cols", 1), 1)
                    for _ in range(min(max(g.get("rows", 1), 1), 2))
                ]
                for n, g in regions.items()
                if g.get("cardinality") != "single_cell"
            }
            out = td / "out.xlsx"
            sink = []
            xlsx_generate.generate(
                prof,
                loaded.shell_path,
                GridDocument(cells=cells, regions=regs),
                out,
                findings=sink,
            )
            self.assertTrue(out.is_file())
            report = run_qa(
                out, prof, shell=loaded.shell_path, extra_findings=list(sink), qa="fast"
            )
            errors = [f.message for f in report.findings if f.severity == "ERROR"]
            self.assertEqual(errors, [], errors)


if __name__ == "__main__":  # pragma: no cover
    unittest.main()
