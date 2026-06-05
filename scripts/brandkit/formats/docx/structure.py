# SPDX-License-Identifier: MIT
"""DOCX structure helpers — ordered skeleton detection, TOC detection, and
structure-aware body clearing.

The template's *ordered document structure* (cover -> toc -> body) is first-class:
generation must preserve the cover region and the TOC region and only replace the
freeform body region. Everything here works on the lxml element tree
(``doc.element`` / ``doc.element.body``) because python-docx does not reliably
expose block-level SDTs (the TOC and the cover title both commonly live inside
block-level ``w:sdt`` elements) and does not expose drawing text boxes.

Detection is grounded in evidence, never in brand-specific names:

- **Cover region** — the body-level content *before* the first TOC region or first
  Heading-1 paragraph. ``cover.discover_cover`` locates the cover anchors (SDTs /
  placeholders / logos); here we mark which top-level body children belong to it.
- **TOC region** — a block-level ``w:sdt`` whose ``w:docPartGallery`` is
  ``Table of Contents``, OR a paragraph using a TOC/TOCHeading style, OR a
  ``w:instrText`` starting with ``TOC``, OR a heading whose text is a known
  contents word in any of EN/IT/FR/DE/ES (multilingual).
- **Body region** — everything after the TOC (or after the cover if no TOC) up to
  the final body-level ``w:sectPr``.

Brand-agnostic TOC detection: rather than hardcoding one template's ``TOCHeading`` /
``Sommario`` literals, any ``*toc*``-named paragraph style and any multilingual
contents word (EN/IT/FR/DE/ES) counts, so it works on any company template in any
language.
"""
from __future__ import annotations

from typing import Optional

# ---------------------------------------------------------------------------
# OOXML namespaces. python-docx registers ``w`` but the literal URI is kept here
# so element matching is robust regardless of prefix registration.
# ---------------------------------------------------------------------------
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def w(tag: str) -> str:
    """Return the Clark-notation qualified name for a ``w:`` local name."""
    return f"{{{W_NS}}}{tag}"


def _local_name(tag) -> str:
    """Return the local name of an lxml tag, robust to non-string tags."""
    if not isinstance(tag, str):
        return ""
    return tag.rsplit("}", 1)[-1]


# ---------------------------------------------------------------------------
# Demo markers (kept for back-compat with the M1 fixture-driven detector).
# ---------------------------------------------------------------------------
DEMO_MARKERS = ("Example first-level title", "General instructions")

# Multilingual "contents" words (EN/IT/FR/DE/ES). Lowercased; matched against the
# stripped, lowercased text of a heading paragraph. Brand-agnostic.
TOC_HEADING_WORDS: frozenset[str] = frozenset(
    {
        "contents",
        "table of contents",
        "sommario",
        "indice",
        "indice generale",
        "inhalt",
        "inhaltsverzeichnis",
        "table des matieres",
        "table des matières",
        "índice",
        "indice de contenido",
        "índice de contenido",
        "contenido",
        "contenidos",
    }
)

# Token that, when contained (case-insensitively) in a paragraph style id/name,
# marks the paragraph as part of a TOC (covers ``TOC``, ``TOCHeading``, ``TOC 1``,
# localized ``Sommario``/``Indice`` style names, etc.).
TOC_STYLE_TOKENS: frozenset[str] = frozenset({"toc", "sommario", "indice", "inhalt", "contenido"})

# Leading token of a ``w:instrText`` TOC field instruction.
TOC_INSTR_PREFIX = "TOC"

# Builtin Heading-1 style ids/names that mark the start of body content when no
# TOC is present (matched case-insensitively, with spaces stripped).
_HEADING1_KEYS = frozenset({"heading1", "heading 1"})


# ---------------------------------------------------------------------------
# lxml paragraph helpers
# ---------------------------------------------------------------------------
def _p_style_val(p) -> Optional[str]:
    """Return the ``w:pStyle/@w:val`` of a ``w:p`` element, or None."""
    pPr = p.find(w("pPr"))
    if pPr is None:
        return None
    pStyle = pPr.find(w("pStyle"))
    if pStyle is None:
        return None
    return pStyle.get(w("val"))


def _p_text(p) -> str:
    """Concatenate all ``w:t`` text nodes inside a ``w:p`` element."""
    parts = [t.text for t in p.iter(w("t")) if t.text]
    return "".join(parts)


def _style_is_toc(style_val: Optional[str]) -> bool:
    if not style_val:
        return False
    low = style_val.replace(" ", "").lower()
    return any(tok in low for tok in TOC_STYLE_TOKENS)


def _style_is_heading1(style_val: Optional[str]) -> bool:
    if not style_val:
        return False
    return style_val.replace(" ", "").lower() in _HEADING1_KEYS


def _text_is_toc_word(text: str) -> bool:
    return text.strip().lower() in TOC_HEADING_WORDS


# ---------------------------------------------------------------------------
# TOC detection (generic, brand-agnostic)
# ---------------------------------------------------------------------------
def _element_holds_toc(el) -> bool:
    """Return True if a top-level body child ``el`` is (or contains) a TOC region.

    Generic detection (any of):
      - a block-level ``w:sdt`` whose ``w:docPartGallery/@w:val`` is
        ``Table of Contents`` (matched case-insensitively), or
      - a descendant paragraph whose style is a TOC/TOCHeading style, or
      - a descendant paragraph whose text is a known multilingual contents word, or
      - a descendant ``w:instrText`` starting with ``TOC``.
    """
    ln = _local_name(el.tag)
    if ln == "sdt":
        for gallery in el.iter(w("docPartGallery")):
            val = (gallery.get(w("val")) or "").strip().lower()
            if val == "table of contents":
                return True
    # Scan descendants for TOC paragraphs / instruction fields.
    for d in el.iter():
        dln = _local_name(d.tag)
        if dln == "instrText" and d.text and d.text.strip().startswith(TOC_INSTR_PREFIX):
            return True
        if dln == "p":
            if _style_is_toc(_p_style_val(d)):
                return True
            if _text_is_toc_word(_p_text(d)):
                return True
    return False


def is_toc_present(doc) -> bool:
    """Return True if the document body contains a TOC region anywhere."""
    body = doc.element.body
    return any(_element_holds_toc(child) for child in body)


# ---------------------------------------------------------------------------
# Top-level region classification
# ---------------------------------------------------------------------------
def _is_sectpr(el) -> bool:
    return _local_name(el.tag) == "sectPr"


def _child_starts_body(el) -> bool:
    """Return True if a top-level body child is a Heading-1 paragraph (the body
    start when there is no TOC)."""
    if _local_name(el.tag) != "p":
        return False
    if _style_is_heading1(_p_style_val(el)):
        return True
    return False


def classify_body_children(doc) -> list[dict]:
    """Classify each top-level body child into a region.

    Returns one descriptor per top-level child (in document order):
    ``{"index", "tag", "region", "is_sectpr"}`` where ``region`` is one of
    ``"cover" | "toc" | "body"`` (``None`` for the final ``sectPr`` sentinel).

    Boundaries (evidence-based, never positional):
      - the TOC region is the *contiguous front-matter span* from the first child
        that holds a TOC marker through the last one. Real templates often stack
        several indexes (e.g. a main table-of-contents SDT, then a table-of-tables
        and a table-of-figures) separated by heading/blank paragraphs that do not
        each individually carry a TOC marker; those interleaved paragraphs are part
        of the same front matter and must travel with it. Anchoring on the first and
        last TOC markers keeps the region as one block (still evidence-based — the
        span is bounded by actual TOC markers, never by a hardcoded index);
      - the cover region is everything before the first TOC child (or, if there is
        no TOC, everything before the first Heading-1 paragraph);
      - the body region is everything after the TOC span (or after the cover) up to
        the final body-level ``sectPr``.
    """
    body = doc.element.body
    children = list(body)

    toc_markers = {i for i, el in enumerate(children) if not _is_sectpr(el) and _element_holds_toc(el)}
    first_toc = min(toc_markers) if toc_markers else None
    last_toc = max(toc_markers) if toc_markers else None
    # The TOC region spans the whole contiguous front matter [first_toc .. last_toc],
    # not just the children that each carry a TOC marker, so index headings / blank
    # separator paragraphs between stacked indexes are kept with the front matter.
    toc_indices = (
        {i for i in range(first_toc, last_toc + 1) if not _is_sectpr(children[i])}
        if first_toc is not None
        else set()
    )
    first_h1 = next(
        (i for i, el in enumerate(children) if not _is_sectpr(el) and _child_starts_body(el)),
        None,
    )

    # The cover ends at the FIRST of {first TOC child, first Heading-1}: both are
    # valid "end of cover" signals (a Heading-1 before the TOC means the document
    # has no real cover, only body). The body begins after the TOC when the TOC is
    # the cover boundary, else at the Heading-1.
    if first_toc is not None and (first_h1 is None or first_toc <= first_h1):
        cover_end = first_toc
        body_start = last_toc + 1
    elif first_h1 is not None:
        # A Heading-1 precedes the TOC (or there is no TOC): the cover ends and the
        # body begins at that Heading-1. (TOC children, if any, are still tagged
        # 'toc' below by membership in ``toc_indices``.)
        cover_end = first_h1
        body_start = first_h1
    else:
        # Neither a TOC nor a Heading-1: no cover boundary signal -> all body.
        cover_end = 0
        body_start = 0

    out: list[dict] = []
    for i, el in enumerate(children):
        if _is_sectpr(el):
            out.append({"index": i, "tag": _local_name(el.tag), "region": None, "is_sectpr": True})
            continue
        if i in toc_indices:
            region = "toc"
        elif i < cover_end:
            region = "cover"
        elif i >= body_start:
            region = "body"
        else:
            # Between cover_end and body_start with no TOC shouldn't happen, but
            # default to body to stay safe.
            region = "body"
        out.append({"index": i, "tag": _local_name(el.tag), "region": region, "is_sectpr": False})
    return out


# ---------------------------------------------------------------------------
# Ordered skeleton (the profile["structure"] payload)
# ---------------------------------------------------------------------------
def detect_skeleton(doc, cover_anchors: Optional[list] = None) -> dict:
    """Detect the template's ordered top-level skeleton.

    Returns the ``structure`` section for ``profile.json``::

        {"ordered": True,
         "skeleton": [ {"region","order","role","required","repeatable",...}, ... ]}

    Only regions actually present in the template are included. ``ordered`` means
    the top-level region order must be respected on generation. The ``body`` region
    is ``freeform`` (element order inside it is NOT prescribed).
    """
    classes = classify_body_children(doc)
    has_cover = any(c["region"] == "cover" for c in classes) or bool(cover_anchors)
    has_toc = any(c["region"] == "toc" for c in classes)
    has_body = any(c["region"] == "body" for c in classes)

    skeleton: list[dict] = []
    order = 0
    if has_cover:
        skeleton.append(
            {
                "region": "cover",
                "order": order,
                "role": "section.cover",
                "required": True,
                "repeatable": False,
                "evidence": "body content before first TOC/Heading-1; cover anchors discovered"
                if cover_anchors
                else "body content before first TOC/Heading-1",
            }
        )
        order += 1
    if has_toc:
        skeleton.append(
            {
                "region": "toc",
                "order": order,
                "role": "section.toc",
                "required": True,
                "repeatable": False,
                "evidence": "block-level w:sdt docPartGallery 'Table of Contents', "
                "TOC-styled paragraph, w:instrText 'TOC', or multilingual contents heading",
            }
        )
        order += 1
    # The body region always exists conceptually (a document has content); include
    # it whenever there is post-cover/post-TOC content OR there is no other region.
    if has_body or not skeleton:
        skeleton.append(
            {
                "region": "body",
                "order": order,
                "role": "section.body",
                "required": True,
                "repeatable": True,
                "freeform": True,
                "evidence": "everything after the TOC (or after the cover) up to the final body sectPr",
            }
        )

    return {"ordered": True, "skeleton": skeleton}


# ---------------------------------------------------------------------------
# Demo-region detection (kept; made less fixture-specific while still anchoring
# on evidence). Backwards-compatible shape.
# ---------------------------------------------------------------------------
def detect_demo_region(doc) -> dict:
    """Detect the template's demo / instruction body region.

    Anchors on evidence, not on a hardcoded index. The region begins at the first
    body-region Heading-1 paragraph (the demo body start) and the presence of any
    known demo/instruction marker phrase is recorded. The legacy fixture markers
    are still honoured so the M1 synthetic template keeps working, but detection no
    longer *requires* them: any body-region content after the cover/TOC is treated
    as a candidate demo region.
    """
    classes = classify_body_children(doc)
    body_children = [c for c in classes if c["region"] == "body"]
    body = doc.element.body
    children = list(body)

    # First body-region Heading-1 paragraph -> the demo start anchor.
    start_style_id = None
    start_text = None
    for c in body_children:
        el = children[c["index"]]
        if _local_name(el.tag) == "p" and _style_is_heading1(_p_style_val(el)):
            start_style_id = _p_style_val(el)
            start_text = _p_text(el)[:200]
            break

    marker_hits = [m for m in DEMO_MARKERS if any(m in p.text for p in doc.paragraphs)]
    present = bool(body_children) or bool(marker_hits)

    return {
        "present": present,
        "start_style_id": start_style_id or ("Heading1" if marker_hits else None),
        "start_text": start_text or (marker_hits[0] if marker_hits else None),
        "instruction_markers": marker_hits,
    }


# ---------------------------------------------------------------------------
# Body clearing
# ---------------------------------------------------------------------------
def clear_body(doc) -> None:
    """Remove every top-level body child except the final ``sectPr``.

    Back-compat entry point (M1 behaviour). Prefer :func:`clear_body_region`, which
    preserves the cover and TOC regions. This is now routed through the structural
    clear with both preservation flags off, so it behaves identically to the old
    "wipe everything" clear.
    """
    clear_body_region(doc, preserve_cover=False, preserve_toc=False)


def clear_body_region(doc, structure: Optional[dict] = None, *, preserve_cover: bool = True, preserve_toc: bool = True) -> None:
    """Remove ONLY the body region, preserving the cover and TOC regions.

    Keeps the cover block(s) and the TOC sdt/region and the final ``sectPr``;
    removes only the in-between/after demo body content. python-docx ``add_*`` then
    appends new content into the (now-empty) body region, immediately before the
    ``sectPr`` — which is exactly the right place.

    Args:
        doc: a python-docx ``Document``.
        structure: the profile ``structure`` section (currently informational; the
            actual region boundaries are recomputed from the live tree so the call
            is robust even if the profile is stale or absent).
        preserve_cover: keep the cover region (default True).
        preserve_toc: keep the TOC region (default True).

    Idempotent: clearing an already-clear body is a no-op. Never crashes when the
    template has no cover and/or no TOC.
    """
    body = doc.element.body
    classes = classify_body_children(doc)
    children = list(body)
    keep_regions = set()
    if preserve_cover:
        keep_regions.add("cover")
    if preserve_toc:
        keep_regions.add("toc")

    for c in classes:
        if c["is_sectpr"]:
            continue  # always keep the final sectPr
        if c["region"] in keep_regions:
            continue
        body.remove(children[c["index"]])


# ---------------------------------------------------------------------------
# TOC refresh
# ---------------------------------------------------------------------------
def refresh_toc(doc) -> bool:
    """Mark TOC fields dirty so Word recomputes them on open.

    Sets ``w:updateFields val="true"`` in ``w:settings`` AND marks every
    ``w:fldChar fldCharType="begin"`` of a TOC complex field with ``w:dirty="true"``.
    Returns True if a TOC was found and marked. Safe (no-op + returns False) when
    the template has no TOC field. Never duplicates the TOC.
    """
    if not is_toc_present(doc):
        return False

    # 1) Global updateFields in settings.xml so Word offers to refresh on open.
    try:
        settings = doc.settings.element
        existing = settings.find(w("updateFields"))
        if existing is None:
            from lxml import etree

            existing = etree.SubElement(settings, w("updateFields"))
        existing.set(w("val"), "true")
    except Exception:
        pass

    # 2) Per-field dirty flag on the TOC field begin char(s).
    marked = False
    body = doc.element.body
    for fld in body.iter(w("fldChar")):
        if fld.get(w("fldCharType")) == "begin":
            fld.set(w("dirty"), "true")
            marked = True
    return marked or True
