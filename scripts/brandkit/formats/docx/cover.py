# SPDX-License-Identifier: MIT
"""DOCX cover-anchor discovery and composition.

Real company covers seldom expose their title as a plain paragraph: the title
usually lives in a block-level ``w:sdt`` (a content control, with an ``alias`` /
``dataBinding`` / ``docPartGallery`` / placeholder prompt). python-docx's
``doc.paragraphs`` cannot see block-level SDTs, so discovery and composition both
work on the lxml tree here. The cover title is filled IN PLACE — only the inner
placeholder run's text is overwritten so the run-level formatting (``w:rPr``) the
brand defined survives — and a brand-new title paragraph is appended only when the
shell genuinely has no cover region (and then it is inserted before the first
toc/body child so it lands on the cover, never after the TOC).
"""
from __future__ import annotations

from typing import Optional

from brandkit.common import text as textutil
from brandkit.formats.docx.styles import lookup_style
from brandkit.formats.docx.structure import (
    _element_holds_strong_toc,
    _local_name,
    classify_body_children,
    w,
)
from brandkit.ir.model import Cover
from brandkit.qa.model import Finding
from brandkit.profile import schema


PLACEHOLDER_TITLE = "{{title}}"
# Substrings (case-insensitive) that mark a paragraph/SDT as the cover title
# placeholder when no structural SDT metadata is present. Brand-agnostic prompts.
_TITLE_PROMPT_TOKENS: tuple[str, ...] = ("insert title", "title", "titolo", "titel", "titre", "titulo")


def _sdt_props(sdt):
    """Return the ``w:sdtPr`` element of a ``w:sdt``, or None."""
    return sdt.find(w("sdtPr"))


def _sdt_is_title(sdt) -> bool:
    """Heuristically decide whether a block-level ``w:sdt`` is the cover title.

    Evidence, in order of strength:
      - ``w:sdtPr/w:alias/@w:val`` or ``w:dataBinding`` xpath mentions a title
        token (``title``/``titolo``/``titel``/``titre``/``titulo``);
      - ``w:docPartGallery/@w:val`` mentions a title token;
      - the inner text is/contains a placeholder prompt (``{{title}}`` / "insert
        title" / a title token). Brand-agnostic and multilingual.

    A Table-of-Contents content control is **never** the cover title, no matter
    what title-like words its rendered entries contain. A TOC SDT
    (``docPartGallery='Table of Contents'`` or an inner ``w:instrText`` ``TOC``
    field) is excluded up front so the weak inner-text token match below can never
    misclassify the TOC as the title — which would dump the user's title into the
    TOC content control and blank every TOC entry. The author's own title slot
    (an ``alias``/``dataBinding`` SDT) still matches via the strong checks.
    """
    if _element_holds_strong_toc(sdt):
        return False
    props = _sdt_props(sdt)
    if props is not None:
        alias = props.find(w("alias"))
        if alias is not None:
            val = (alias.get(w("val")) or "").lower()
            if any(tok in val for tok in _TITLE_PROMPT_TOKENS):
                return True
        binding = props.find(w("dataBinding"))
        if binding is not None:
            xpath = (binding.get(w("xpath")) or "").lower()
            if any(tok in xpath for tok in _TITLE_PROMPT_TOKENS):
                return True
        for gallery in props.iter(w("docPartGallery")):
            val = (gallery.get(w("val")) or "").lower()
            if any(tok in val for tok in _TITLE_PROMPT_TOKENS):
                return True
    text = _sdt_text(sdt).strip().lower()
    if PLACEHOLDER_TITLE in text:
        return True
    if any(tok in text for tok in _TITLE_PROMPT_TOKENS):
        return True
    return False


def _sdt_text(sdt) -> str:
    return "".join(t.text for t in sdt.iter(w("t")) if t.text)


def _iter_block_sdts(doc):
    """Yield top-level block-level ``w:sdt`` elements in document order."""
    for child in doc.element.body:
        if _local_name(child.tag) == "sdt":
            yield child


def _fill_sdt_title(sdt, title: str) -> bool:
    """Overwrite the SDT's first text run with ``title`` IN PLACE.

    Writes into the first ``w:t`` inside the SDT content (``w:sdtContent``),
    preserving its run formatting, and clears any other ``w:t`` in the content so
    the placeholder prompt does not linger. Returns True on success.
    """
    content = sdt.find(w("sdtContent"))
    scope = content if content is not None else sdt
    texts = list(scope.iter(w("t")))
    if not texts:
        return False
    texts[0].text = title
    # xml:space=preserve so leading/trailing spaces in the title are not trimmed.
    texts[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    for extra in texts[1:]:
        extra.text = ""
    # Drop the "showingPlcHdr" flag so Word treats the content as real, not prompt.
    props = _sdt_props(sdt)
    if props is not None:
        for plc in props.findall(w("showingPlcHdr")):
            props.remove(plc)
    return True


def discover_cover(doc) -> tuple[list[dict], dict]:
    """Discover the cover title anchor (block-level SDT first, then paragraph)."""
    anchors: list[dict] = []
    for sdt in _iter_block_sdts(doc):
        if _sdt_is_title(sdt):
            anchors.append(
                {
                    "id": "title",
                    "container": "sdt",
                    "placeholder": _sdt_text(sdt)[:200],
                    "branches": None,
                }
            )
            break
    if not anchors:
        for idx, para in enumerate(doc.paragraphs[:8]):
            if PLACEHOLDER_TITLE in para.text or "Insert title" in para.text:
                anchors.append(
                    {
                        "id": "title",
                        "container": "paragraph",
                        "paragraph_index": idx,
                        "placeholder": para.text,
                        "branches": None,
                    }
                )
                break
    anchor_block = {
        "cover": {
            "kind": "sdt_anchored" if anchors else "NONE",
            "slots_found": len(anchors),
        }
    }
    return anchors, anchor_block


def compose_cover(doc, cover: Cover | None, profile: dict, *, findings: Optional[list[Finding]] = None) -> None:
    """Fill the PRESERVED cover title anchor in place.

    Resolution order:
      1. a block-level ``w:sdt`` cover title (the real-template case) — fill the
         inner run in place, preserving run formatting;
      2. a ``{{title}}`` / "Insert title" placeholder paragraph — overwrite only
         that run's text (never ``para.text =``, which would destroy run rPr);
      3. no cover anchor at all — append a title paragraph, inserted BEFORE the
         first toc/body child so it lands on the cover, and record a degraded
         finding. The title must NEVER land after the TOC.
    """
    sink = findings if findings is not None else []
    if cover is None:
        return
    title = textutil.runs_to_text(cover.title or []) or str(cover.fields.get("title", ""))
    if not title:
        return

    # 1) Block-level SDT cover title.
    for sdt in _iter_block_sdts(doc):
        if _sdt_is_title(sdt) and _fill_sdt_title(sdt, title):
            _apply_role_style_sdt(doc, sdt, profile, "cover.title")
            return

    # 2) Placeholder paragraph — overwrite only the matching run's text in place.
    for para in doc.paragraphs[:8]:
        if PLACEHOLDER_TITLE in para.text or "Insert title" in para.text:
            _fill_paragraph_in_place(para, title)
            _apply_role_style(doc, para, profile, "cover.title")
            return

    # 3) No cover anchor: append a title paragraph but place it on the cover, i.e.
    # BEFORE the first toc/body child, never after the TOC.
    para = doc.add_paragraph(title)
    _apply_role_style(doc, para, profile, "cover.title")
    _move_before_first_toc_or_body(doc, para)
    sink.append(
        Finding(
            "cover_degraded",
            schema.Severity.WARNING.value,
            "no cover anchor in shell; title paragraph appended before the first toc/body child",
        )
    )


def _fill_paragraph_in_place(para, title: str) -> None:
    """Set the title on a placeholder paragraph without destroying run rPr.

    Writes ``title`` into the first run and clears the remaining runs, so the
    run-level formatting the brand defined on the first run is preserved (unlike
    ``para.text = title``, which rebuilds the runs from scratch).
    """
    runs = para.runs
    if not runs:
        para.add_run(title)
        return
    runs[0].text = title
    for r in runs[1:]:
        r.text = ""


def _move_before_first_toc_or_body(doc, para) -> None:
    """Move ``para``'s ``w:p`` element so it precedes the first toc/body child."""
    body = doc.element.body
    p_el = para._p
    classes = classify_body_children(doc)
    children = list(body)
    target = None
    for c in classes:
        if c["index"] >= len(children):
            continue
        el = children[c["index"]]
        if el is p_el:
            continue
        if c["region"] in ("toc", "body"):
            target = el
            break
    if target is not None:
        body.remove(p_el)
        target.addprevious(p_el)


def _apply_role_style(doc, para, profile: dict, role_id: str) -> None:
    entry = (profile.get("roles") or {}).get(role_id) or {}
    resolver = entry.get("resolver") or {}
    style = lookup_style(doc, resolver)
    if style is not None:
        para.style = style


def _apply_role_style_sdt(doc, sdt, profile: dict, role_id: str) -> None:
    """Apply the cover.title paragraph style to the paragraph(s) inside the SDT."""
    entry = (profile.get("roles") or {}).get(role_id) or {}
    resolver = entry.get("resolver") or {}
    style = lookup_style(doc, resolver)
    if style is None:
        return
    style_id = getattr(style, "style_id", None)
    if not style_id:
        return
    content = sdt.find(w("sdtContent"))
    scope = content if content is not None else sdt
    for p in scope.iter(w("p")):
        pPr = p.find(w("pPr"))
        if pPr is None:
            from lxml import etree

            pPr = etree.SubElement(p, w("pPr"))
            p.insert(0, pPr)
        pStyle = pPr.find(w("pStyle"))
        if pStyle is None:
            from lxml import etree

            pStyle = etree.SubElement(pPr, w("pStyle"))
        pStyle.set(w("val"), style_id)
        break
