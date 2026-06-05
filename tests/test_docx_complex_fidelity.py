# SPDX-License-Identifier: MIT
"""DOCX fidelity regression tests on the committed COMPLEX fixture.

These reproduce, on ``tests/fixtures/complex/acme_complex.docx`` (a 100% synthetic
Acme template), the confirmed DOCX faithfulness fixes and assert the brand
artifacts now win over the generic/builtin floor:

  - D1  list ROLE NOMINATION from ``word/numbering.xml`` (a paragraph style whose
        definition carries a ``w:numPr`` binds ``list.<family>.<level>`` to the
        custom brand style, family read from the abstractNum ``w:numFmt``), AND
        the generator writes a real ``w:numPr`` (numId + per-item ilvl) so the
        list renders bulleted/numbered - not flat ``Normal`` paragraphs.
  - D2  ``table.default`` NOMINATION from the custom ``w:type="table"`` brand
        style (``Acme Table``), not the builtin ``Normal Table`` / ``Table Grid``.
  - D6  ``callout.info`` nominated from the brand boxed style (``Acme Callout``,
        ``w:shd`` + ``w:pBdr``), NOT the builtin ``Footnote Text`` (the ``note``
        lexicon token must not let a builtin win).
  - D4  a filled cover PARAGRAPH slot (not an SDT) re-applies its bound role
        style (``cover.title``), mirroring the SDT branch.

All assertions inspect the produced role registry and the real output OOXML.
"""
from __future__ import annotations

import hashlib
import json
import os
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from brandkit.formats.docx import extract as docx_extract
from brandkit.formats.docx import generate as docx_generate
from brandkit.formats.docx import roles as docx_roles
from brandkit.formats.docx import structure as docx_structure
from brandkit.formats.docx.structure import w
from brandkit.ir import model as ir
from brandkit.profile import schema

_COMPLEX_DOCX = (
    Path(__file__).resolve().parents[0] / "fixtures" / "complex" / "acme_complex.docx"
)


def _extract_profile(td: str) -> dict:
    """Extract the complex docx fixture into ``td`` and return the loaded profile."""
    old = Path.cwd()
    os.chdir(td)
    try:
        pj = docx_extract.extract(_COMPLEX_DOCX, "acme", scope="project")
        return json.loads(Path(pj).read_text())
    finally:
        os.chdir(old)


def _num_pr(p):
    """Return ``(numId, ilvl)`` of a paragraph's ``w:numPr``, or ``(None, None)``."""
    pPr = p._p.find(w("pPr"))
    if pPr is None:
        return None, None
    numPr = pPr.find(w("numPr"))
    if numPr is None:
        return None, None
    n = numPr.find(w("numId"))
    il = numPr.find(w("ilvl"))
    return (
        n.get(w("val")) if n is not None else None,
        il.get(w("val")) if il is not None else None,
    )


@unittest.skipUnless(_COMPLEX_DOCX.exists(), "complex docx fixture missing")
class DocxRoleNominationTest(unittest.TestCase):
    """The role registry promotes the PRESENT brand styles into their roles."""

    def setUp(self):
        self._td = tempfile.TemporaryDirectory()
        self.addCleanup(self._td.cleanup)
        self.prof = _extract_profile(self._td.name)
        self.roles = self.prof["roles"]

    def test_callout_is_brand_style_not_builtin_footnote(self):
        # D6: 'Acme Callout' (boxed custom style) wins over builtin 'Footnote Text'.
        callout = self.roles.get("callout.info")
        self.assertIsNotNone(callout, "callout.info not nominated")
        r = callout["resolver"]
        self.assertEqual(r.get("style_id"), "AcmeCallout")
        self.assertEqual(r.get("style_name"), "Acme Callout")
        self.assertNotIn(r.get("style_id"), {"FootnoteText", "EndnoteText"})

    def test_table_default_is_custom_brand_style(self):
        # D2: the custom w:type='table' 'Acme Table' wins over Normal Table/Grid.
        table = self.roles.get("table.default")
        self.assertIsNotNone(table, "table.default not nominated")
        r = table["resolver"]
        self.assertEqual(r.get("style_id"), "AcmeTable")
        self.assertNotIn(r.get("style_id"), {"TableNormal", "TableGrid"})

    def test_list_roles_bind_custom_styles_with_numbering(self):
        # D1: list.bullet.1 / list.number.1 bind the CUSTOM brand styles and carry
        # the verbatim num_id from the numbering part (family from numFmt, not name).
        bullet = self.roles.get("list.bullet.1")
        number = self.roles.get("list.number.1")
        self.assertIsNotNone(bullet, "list.bullet.1 not nominated")
        self.assertIsNotNone(number, "list.number.1 not nominated")
        self.assertEqual(bullet["resolver"].get("style_id"), "AcmeBulletL1")
        self.assertEqual(number["resolver"].get("style_id"), "AcmeNumberL1")
        # The resolver carries the verbatim numbering id the generator re-asserts.
        self.assertEqual(str(bullet["resolver"].get("num_id")), "1")
        self.assertEqual(str(number["resolver"].get("num_id")), "3")

    def test_list_role_is_not_the_normal_floor(self):
        # The old 'Normal as list' floor must NOT be used when a real list style
        # exists - list.bullet.1 must not resolve to the body 'Normal' style.
        bullet = self.roles["list.bullet.1"]["resolver"]
        self.assertNotEqual(bullet.get("style_id"), "Normal")

    def test_builtin_list_style_loses_to_custom(self):
        # CC-2: Word's builtin latent 'List Bullet' (id ListBullet) references the
        # same numId, but the custom 'Acme Bullet L1' is nominated instead.
        self.assertEqual(
            self.roles["list.bullet.1"]["resolver"].get("style_id"), "AcmeBulletL1"
        )


@unittest.skipUnless(_COMPLEX_DOCX.exists(), "complex docx fixture missing")
class DocxListGenerationTest(unittest.TestCase):
    """The generator writes real numbering onto list items keyed by item.level."""

    def _generate(self, idoc, td) -> Document:
        prof = _extract_profile(td)
        out = Path(td) / "out.docx"
        docx_generate.generate(prof, _COMPLEX_DOCX, idoc, out)
        return Document(out)

    def test_bulleted_and_numbered_lists_get_numPr_and_brand_style(self):
        with tempfile.TemporaryDirectory() as td:
            idoc = ir.IntermediateDocument(
                blocks=[
                    ir.ListBlock(
                        ordered=False,
                        items=[
                            ir.ListItem(
                                runs=[{"t": "Top bullet"}],
                                level=0,
                                items=[ir.ListItem(runs=[{"t": "Nested bullet"}], level=1)],
                            ),
                            ir.ListItem(runs=[{"t": "Second top"}], level=0),
                        ],
                    ),
                    ir.ListBlock(
                        ordered=True,
                        items=[
                            ir.ListItem(runs=[{"t": "Step one"}], level=0),
                            ir.ListItem(runs=[{"t": "Step two"}], level=0),
                        ],
                    ),
                ]
            )
            gen = self._generate(idoc, td)
            by_text = {p.text: p for p in gen.paragraphs if p.text}

            # Bullet items: brand style + numId 1; nested item gets ilvl 1.
            top = by_text["Top bullet"]
            self.assertEqual(top.style.name, "Acme Bullet L1")
            nid, ilvl = _num_pr(top)
            self.assertEqual(nid, "1")
            self.assertEqual(ilvl, "0")
            nested = by_text["Nested bullet"]
            nid_n, ilvl_n = _num_pr(nested)
            self.assertEqual(nid_n, "1")
            self.assertEqual(ilvl_n, "1")  # keyed by item.level

            # Numbered items: brand number style + numId 3.
            step = by_text["Step one"]
            self.assertEqual(step.style.name, "Acme Number L1")
            nid_s, _ = _num_pr(step)
            self.assertEqual(nid_s, "3")

    def test_list_items_are_not_flat_normal_paragraphs(self):
        # Regression on D1's failure mode: a list must not render as 'Normal' with
        # no numbering.
        with tempfile.TemporaryDirectory() as td:
            idoc = ir.IntermediateDocument(
                blocks=[
                    ir.ListBlock(
                        ordered=False,
                        items=[ir.ListItem(runs=[{"t": "Has bullets"}], level=0)],
                    )
                ]
            )
            gen = self._generate(idoc, td)
            p = next(p for p in gen.paragraphs if p.text == "Has bullets")
            self.assertNotEqual(p.style.name, "Normal")
            nid, _ = _num_pr(p)
            self.assertIsNotNone(nid, "list paragraph has no numPr")

    def test_table_uses_brand_table_style(self):
        # D2 through full generation: the produced table carries 'Acme Table'.
        with tempfile.TemporaryDirectory() as td:
            idoc = ir.IntermediateDocument(
                blocks=[
                    ir.Table(
                        columns=[{"t": "A"}, {"t": "B"}],
                        rows=[[ir.TableCell(runs=[{"t": "1"}]), ir.TableCell(runs=[{"t": "2"}])]],
                    )
                ]
            )
            gen = self._generate(idoc, td)
            self.assertEqual(gen.tables[-1].style.name, "Acme Table")

    def test_callout_renders_with_brand_callout_style(self):
        with tempfile.TemporaryDirectory() as td:
            idoc = ir.IntermediateDocument(
                blocks=[ir.Callout(intent="info", runs=[{"t": "Important note"}])]
            )
            gen = self._generate(idoc, td)
            p = next(p for p in gen.paragraphs if "Important note" in p.text)
            self.assertEqual(p.style.name, "Acme Callout")

    def test_list_generation_is_idempotent(self):
        # The numPr writes must be deterministic (idempotent generation).
        with tempfile.TemporaryDirectory() as td:
            prof = _extract_profile(td)
            idoc = ir.IntermediateDocument(
                blocks=[
                    ir.ListBlock(
                        ordered=False,
                        items=[
                            ir.ListItem(
                                runs=[{"t": "A"}],
                                level=0,
                                items=[ir.ListItem(runs=[{"t": "B"}], level=1)],
                            )
                        ],
                    )
                ]
            )
            o1 = Path(td) / "o1.docx"
            o2 = Path(td) / "o2.docx"
            docx_generate.generate(prof, _COMPLEX_DOCX, idoc, o1)
            docx_generate.generate(prof, _COMPLEX_DOCX, idoc, o2)
            self.assertEqual(
                hashlib.sha256(o1.read_bytes()).hexdigest(),
                hashlib.sha256(o2.read_bytes()).hexdigest(),
            )


@unittest.skipUnless(_COMPLEX_DOCX.exists(), "complex docx fixture missing")
class DocxCoverParagraphStyleTest(unittest.TestCase):
    """D4: a filled cover PARAGRAPH slot re-applies its bound role style."""

    def _build_paragraph_cover_shell(self, td) -> Path:
        """A shell whose cover title is a plain paragraph slot (no SDT), with a
        custom 'Brand Cover Title' style and a TOC after, so the slot is in the
        cover region."""
        shell = Path(td) / "shell.docx"
        doc = Document()
        styles = doc.styles.element
        st = OxmlElement("w:style")
        st.set(qn("w:type"), "paragraph")
        st.set(qn("w:styleId"), "BrandCoverTitle")
        st.set(qn("w:customStyle"), "1")
        nm = OxmlElement("w:name")
        nm.set(qn("w:val"), "Brand Cover Title")
        st.append(nm)
        styles.append(st)
        doc.add_paragraph("{{title}}")  # plain-paragraph cover slot
        # a real TOC field after the slot
        p = doc.add_paragraph()
        r = p.add_run()
        fb = OxmlElement("w:fldChar"); fb.set(qn("w:fldCharType"), "begin"); r._r.append(fb)
        r2 = p.add_run()
        it = OxmlElement("w:instrText"); it.text = 'TOC \\o "1-3" \\h'; r2._r.append(it)
        r3 = p.add_run()
        fs = OxmlElement("w:fldChar"); fs.set(qn("w:fldCharType"), "separate"); r3._r.append(fs)
        p.add_run("entry 1")
        r4 = p.add_run()
        fe = OxmlElement("w:fldChar"); fe.set(qn("w:fldCharType"), "end"); r4._r.append(fe)
        doc.add_paragraph("Body Heading", style="Heading 1")
        doc.save(shell)
        return shell

    def test_filled_paragraph_slot_carries_cover_title_style(self):
        with tempfile.TemporaryDirectory() as td:
            shell = self._build_paragraph_cover_shell(td)
            # extract a real profile, then bind cover.title to the brand style.
            old = Path.cwd(); os.chdir(td)
            try:
                pj = docx_extract.extract(shell, "cv", scope="project")
                prof = json.loads(Path(pj).read_text())
            finally:
                os.chdir(old)
            if "cover.title" not in prof["roles"]:
                prof["roles"].setdefault("_index", []).append("cover.title")
            prof["roles"]["cover.title"] = {
                "resolver": {
                    "type": "named_style",
                    "style_id": "BrandCoverTitle",
                    "style_name": "Brand Cover Title",
                },
                "status": "robust",
                "confidence": 1.0,
            }
            para_anchor = next(
                a["id"] for a in prof["surface"]["docx"]["cover_anchors"]
                if a["id"].startswith("para.")
            )
            prof.setdefault("provenance", {}).setdefault("shell", {})["sha256"] = "sh"
            block = schema.empty_comprehension()
            block["status"] = schema.ComprehensionStatus.PRESENT.value
            block["source_shell_sha256"] = "sh"
            block["confidence"] = 0.9
            block["cover_slots"] = {
                para_anchor: {"binds_to": "title", "fill_rule": "in_place", "demo_value": "{{title}}"}
            }
            prof["comprehension"] = block

            out = Path(td) / "out.docx"
            docx_generate.generate(
                prof,
                shell,
                ir.IntermediateDocument(
                    blocks=[ir.Heading(level=1, runs=[{"t": "Sec"}])],
                    cover=ir.Cover(title=[{"t": "My Real Title"}]),
                ),
                out,
            )
            gen = Document(out)
            filled = next(p for p in gen.paragraphs if "My Real Title" in p.text)
            self.assertEqual(filled.style.name, "Brand Cover Title")


@unittest.skipUnless(_COMPLEX_DOCX.exists(), "complex docx fixture missing")
class DocxNumberingHelpersTest(unittest.TestCase):
    """Unit coverage for the structural numbering helpers behind D1."""

    def test_num_family_for_reads_numFmt_not_name(self):
        doc = Document(_COMPLEX_DOCX)
        # numId 1 -> abstractNum 0 (bullet); numId 3 -> abstractNum 1 (decimal).
        self.assertEqual(docx_structure.num_family_for(doc, "1", 0), "bullet")
        self.assertEqual(docx_structure.num_family_for(doc, "3", 0), "number")
        # an unknown id yields None (skip rather than guess).
        self.assertIsNone(docx_structure.num_family_for(doc, "9999", 0))

    def test_style_num_binding_reads_style_definition_numPr(self):
        doc = Document(_COMPLEX_DOCX)
        from docx.enum.style import WD_STYLE_TYPE

        acme_bullet = next(
            s for s in doc.styles
            if s.type == WD_STYLE_TYPE.PARAGRAPH and getattr(s, "style_id", None) == "AcmeBulletL1"
        )
        binding = docx_structure.style_num_binding(acme_bullet)
        self.assertEqual(binding, ("1", 0))

    def test_no_numbering_part_is_not_an_error(self):
        # A document with no list numbering must not crash the helpers.
        doc = Document()
        # default template has a numbering part, but no nominated list style with
        # a custom numPr binding maps to a family resolvable via num_family_for for
        # an absent id; the helper must simply return None.
        self.assertIsNone(docx_structure.num_family_for(doc, "424242", 0))


if __name__ == "__main__":
    unittest.main()
