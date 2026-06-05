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
# marks the paragraph as part of a TOC entry (covers ``TOC``, ``TOC 1``, the
# table-of-figures ``TableOfFigures`` entry style, and localized ``Sommario``/
# ``Indice`` *entry* style names). These are STRONG markers: a paragraph carrying
# one of these styles is an actual TOC/index ENTRY, never ordinary body text.
TOC_STYLE_TOKENS: frozenset[str] = frozenset(
    {"toc", "tableoffigures", "sommario", "indice", "inhalt", "contenido"}
)

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
def _element_holds_strong_toc(el) -> bool:
    """Return True if ``el`` carries a STRONG, unambiguous TOC marker.

    A strong marker is a real structural proof that this child is part of a
    table-of-contents / index region, not merely a heading that happens to be
    named like one. Any of:
      - a block-level ``w:sdt`` whose ``w:docPartGallery/@w:val`` is
        ``Table of Contents`` (case-insensitive), or
      - a descendant ``w:instrText`` whose instruction starts with ``TOC`` (the
        TOC field code; this also covers Table-of-Figures via ``TOC \\c``), or
      - a descendant paragraph using a TOC-entry / Table-of-Figures style
        (``TOC 1``…``TOC 9``, ``TableOfFigures``, localized index-entry styles).

    A bare contents-word heading (``Contents`` / ``Sommario``) is *not* strong on
    its own — it is handled separately as the optional TOC *heading* that may
    immediately precede a strong anchor.
    """
    ln = _local_name(el.tag)
    if ln == "sdt":
        for gallery in el.iter(w("docPartGallery")):
            val = (gallery.get(w("val")) or "").strip().lower()
            if val == "table of contents":
                return True
    for d in el.iter():
        dln = _local_name(d.tag)
        if dln == "instrText" and d.text and d.text.strip().startswith(TOC_INSTR_PREFIX):
            return True
        if dln == "p" and _style_is_toc(_p_style_val(d)):
            return True
    return False


def _is_lone_contents_heading(el) -> bool:
    """Return True if ``el`` is a paragraph whose text is *only* a contents word.

    This is the WEAK signal: a heading literally named ``Contents`` / ``Indice`` /
    ``Sommario`` with no strong TOC structure of its own. It counts as the TOC
    *heading* only when it immediately precedes a strong anchor (see
    :func:`classify_body_children`); on its own it never extends the TOC span,
    because a body section legitimately titled "Contents" must stay body content.
    """
    if _local_name(el.tag) != "p":
        return False
    if _element_holds_strong_toc(el):
        return False
    return _text_is_toc_word(_p_text(el))


def _element_holds_toc(el) -> bool:
    """Back-compat: True if ``el`` carries any TOC signal (strong OR a lone
    contents-word heading). Used only by detection summaries, never to *extend*
    the preserved span (that is anchored on strong markers in
    :func:`classify_body_children`)."""
    return _element_holds_strong_toc(el) or _is_lone_contents_heading(el)


def is_toc_present(doc) -> bool:
    """Return True if the document body contains a real (strong) TOC region."""
    body = doc.element.body
    return any(_element_holds_strong_toc(child) for child in body)


# ---------------------------------------------------------------------------
# Top-level region classification
# ---------------------------------------------------------------------------
def _is_sectpr(el) -> bool:
    """True for the final body-level ``w:sectPr`` sentinel (the last section)."""
    return _local_name(el.tag) == "sectPr"


def _holds_sectpr(el) -> bool:
    """True if ``el`` is a ``w:p`` carrying ``w:pPr/w:sectPr`` (an *intermediate*
    section break). Such a paragraph defines section geometry / header-footer
    references for the section that ends at it and must never be deleted by a body
    clear, even when it otherwise falls in the body region."""
    if _local_name(el.tag) != "p":
        return False
    pPr = el.find(w("pPr"))
    if pPr is None:
        return False
    return pPr.find(w("sectPr")) is not None


def _child_starts_body(el) -> bool:
    """Return True if a top-level body child is a Heading-1 paragraph (a body
    start signal: the first such paragraph after the TOC breaks the TOC span and,
    when there is no TOC, marks the cover/body boundary)."""
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
      - the TOC region anchors on **strong** markers only (a block-level TOC sdt,
        a ``w:instrText`` ``TOC`` field, or a TOC/Table-of-Figures *entry* style).
        It starts at the first strong anchor — extended back by one paragraph if a
        lone contents-word heading (``Contents``/``Sommario``) immediately precedes
        it (that is the TOC heading) — and ends at the **last** strong anchor.
        Stacked index front matter (a table-of-contents, then a table-of-tables and
        a table-of-figures, each a real TOC field) is preserved as one block:
        their headings/blank separators sit *between* strong anchors and travel
        with the span. The span **breaks at the first body Heading-1 after the last
        strong anchor**: a heading literally named "Contents"/"Sommario" appearing
        *after* the real TOC is body content and is cleared, never preserved;
      - the cover region is everything before the TOC heading (or, if there is no
        TOC, everything before the first Heading-1 paragraph);
      - the body region is everything after the TOC span (or after the cover) up to
        the final body-level ``sectPr``. A paragraph holding an intermediate
        ``w:sectPr`` is body-region but flagged ``holds_sectpr`` so a body clear
        skips it (preserving multi-section geometry / header-footer references).
    """
    body = doc.element.body
    children = list(body)

    # Strong TOC anchors only — these are the structural proof of a TOC/index.
    strong = [
        i
        for i, el in enumerate(children)
        if not _is_sectpr(el) and _element_holds_strong_toc(el)
    ]

    toc_indices: set[int] = set()
    first_toc: Optional[int] = None
    if strong:
        first_anchor = strong[0]
        last_anchor = strong[-1]
        toc_start = first_anchor
        # Fold in an immediately-preceding lone contents-word heading (the TOC
        # heading) and any blank separator paragraphs between it and the anchor.
        j = first_anchor - 1
        heading_at: Optional[int] = None
        while j >= 0 and not _is_sectpr(children[j]):
            if _is_lone_contents_heading(children[j]):
                heading_at = j
                break
            if _p_text(children[j]).strip() == "" and _local_name(children[j].tag) == "p":
                j -= 1
                continue
            break
        if heading_at is not None:
            toc_start = heading_at
        first_toc = toc_start

        # The span ends at the last strong anchor, EXCEPT it breaks early at the
        # first body Heading-1 that appears after the last strong anchor would
        # never trigger (Heading-1 is past the anchor). The real risk is a span
        # that, without strong anchors past ``last_anchor``, should still stop at a
        # body Heading-1 *between* anchors only if that heading is not itself index
        # front matter. Strong anchors bound the front matter, so we keep
        # [toc_start .. last_anchor] and then, defensively, never let the span run
        # past the first Heading-1 that follows ``last_anchor``.
        toc_end = last_anchor
        toc_indices = {
            i for i in range(toc_start, toc_end + 1) if not _is_sectpr(children[i])
        }

    # First Heading-1 anywhere (the no-TOC cover/body boundary).
    first_h1 = next(
        (i for i, el in enumerate(children) if not _is_sectpr(el) and _child_starts_body(el)),
        None,
    )

    if first_toc is not None and (first_h1 is None or first_toc <= first_h1):
        # The TOC is the cover boundary. Body starts right after the TOC span.
        cover_end = first_toc
        body_start = max(toc_indices) + 1 if toc_indices else first_toc + 1
    elif first_h1 is not None:
        # A Heading-1 precedes the TOC (or there is no TOC): cover ends / body
        # begins at that Heading-1.
        cover_end = first_h1
        body_start = first_h1
    else:
        # No boundary signal at all -> everything is body.
        cover_end = 0
        body_start = 0

    out: list[dict] = []
    for i, el in enumerate(children):
        if _is_sectpr(el):
            out.append(
                {"index": i, "tag": _local_name(el.tag), "region": None, "is_sectpr": True, "holds_sectpr": False}
            )
            continue
        if i in toc_indices:
            region = "toc"
        elif i < cover_end:
            region = "cover"
        elif i >= body_start:
            region = "body"
        else:
            region = "body"
        out.append(
            {
                "index": i,
                "tag": _local_name(el.tag),
                "region": region,
                "is_sectpr": False,
                "holds_sectpr": _holds_sectpr(el),
            }
        )
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
        if c.get("holds_sectpr"):
            continue  # preserve intermediate section breaks (geometry / hdr-ftr)
        if c["region"] in keep_regions:
            continue
        body.remove(children[c["index"]])


# ---------------------------------------------------------------------------
# TOC refresh
# ---------------------------------------------------------------------------
# Local names that, per the ECMA-376 ``CT_Settings`` sequence, legally come AFTER
# ``w:updateFields``. ``updateFields`` must be inserted before the first of these
# that already exists in settings.xml so the part stays schema-valid (Word and
# strict validators reject out-of-order children).
_SETTINGS_AFTER_UPDATEFIELDS: tuple[str, ...] = (
    "hdrShapeDefaults",
    "footnotePr",
    "endnotePr",
    "compat",
    "rsids",
    "mathPr",
    "uiCompat97To2003",
    "attachedSchema",
    "themeFontLang",
    "clrSchemeMapping",
    "doNotIncludeSubdocsInStats",
    "doNotAutoCompressPictures",
    "forceUpgrade",
    "captions",
    "readModeInkLockDown",
    "smartTagType",
    "schemaLibrary",
    "shapeDefaults",
    "doNotEmbedSmartTags",
    "decimalSymbol",
    "listSeparator",
    "docId",
    "discardImageEditingData",
    "defaultImageDpi",
    "conflictMode",
    "chartTrackingRefBased",
)


def _set_update_fields(doc) -> None:
    """Set ``w:updateFields val="true"`` in settings.xml, in schema position."""
    from lxml import etree

    settings = doc.settings.element
    existing = settings.find(w("updateFields"))
    if existing is None:
        existing = etree.SubElement(settings, w("updateFields"))
        # Move it into the correct ECMA-376 position: before the first child that
        # must legally follow ``updateFields``.
        for child in list(settings):
            if child is existing:
                continue
            if _local_name(child.tag) in _SETTINGS_AFTER_UPDATEFIELDS:
                child.addprevious(existing)
                break
    existing.set(w("val"), "true")


def refresh_toc(doc) -> int:
    """Mark only TOC fields dirty so Word recomputes them on open.

    Sets ``w:updateFields val="true"`` in ``w:settings`` (in schema position) AND
    marks the ``w:fldChar fldCharType="begin"`` of each *TOC* complex field with
    ``w:dirty="true"``. A complex field is delimited by ``begin``/``separate``/
    ``end`` ``w:fldChar`` chars and identified by its ``w:instrText``; only the
    begin char of the field whose own instruction starts with ``TOC`` is marked.
    Nested ``PAGEREF``/``HYPERLINK`` fields inside a TOC (and every non-TOC field)
    are left untouched, so a doc with 59 fields but 3 TOCs marks exactly 3.

    Returns the number of TOC fields marked dirty (0 when the template has no TOC
    field — in which case nothing is written). Never duplicates the TOC.
    """
    if not is_toc_present(doc):
        return 0

    # 1) Global updateFields in settings.xml so Word offers to refresh on open.
    try:
        _set_update_fields(doc)
    except Exception:
        pass

    # 2) Walk fldChar/instrText in document order, tracking field nesting with a
    # stack so each instrText is attributed to its OWN enclosing field's begin.
    marked = 0
    body = doc.element.body
    stack: list = []  # stack of begin fldChar elements (innermost on top)
    for el in body.iter(w("fldChar"), w("instrText")):
        ln = _local_name(el.tag)
        if ln == "fldChar":
            ctype = el.get(w("fldCharType"))
            if ctype == "begin":
                stack.append(el)
            elif ctype == "end":
                if stack:
                    stack.pop()
        elif ln == "instrText":
            text = (el.text or "").strip()
            if stack and text.startswith(TOC_INSTR_PREFIX):
                begin = stack[-1]
                if begin.get(w("dirty")) != "true":
                    begin.set(w("dirty"), "true")
                    marked += 1
    return marked
