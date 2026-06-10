# SPDX-License-Identifier: MIT
"""DOCX structure helpers - ordered skeleton detection, TOC detection, and
structure-aware body clearing.

The template's *ordered document structure* (cover -> toc -> body) is first-class:
generation must preserve the cover region and the TOC region and only replace the
freeform body region. Everything here works on the lxml element tree
(``doc.element`` / ``doc.element.body``) because python-docx does not reliably
expose block-level SDTs (the TOC and the cover title both commonly live inside
block-level ``w:sdt`` elements) and does not expose drawing text boxes.

Detection is grounded in evidence, never in brand-specific names:

- **Cover region** - the body-level content *before* the first TOC region or first
  Heading-1 paragraph. ``cover.discover_cover`` locates the cover anchors (SDTs /
  placeholders / logos); here we mark which top-level body children belong to it.
- **TOC region** - a block-level ``w:sdt`` whose ``w:docPartGallery`` is
  ``Table of Contents``, OR a paragraph using a TOC/TOCHeading style, OR a
  ``w:instrText`` starting with ``TOC``, OR a heading whose text is a known
  contents word in any of EN/IT/FR/DE/ES (multilingual).
- **Body region** - everything after the TOC (or after the cover if no TOC) up to
  the final body-level ``w:sectPr``.

Brand-agnostic TOC detection: rather than hardcoding one template's ``TOCHeading`` /
``Sommario`` literals, any ``*toc*``-named paragraph style and any multilingual
contents word (EN/IT/FR/DE/ES) counts, so it works on any company template in any
language.
"""

from __future__ import annotations

import copy
import re
from typing import Optional

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from brandkit.ooxml.fields import iter_complex_field_events
from brandkit.ooxml.names import local_name as _local_name, make_qn
from brandkit.ooxml.pack import NAMESPACES

# ---------------------------------------------------------------------------
# OOXML namespaces. python-docx registers ``w`` but the literal URI is kept here
# so element matching is robust regardless of prefix registration. ``w`` and
# ``_local_name`` come from the shared :mod:`brandkit.ooxml.names` layer; they are
# re-exported here (and ``W_NS`` is kept) so existing importers keep working.
# ---------------------------------------------------------------------------
W_NS = NAMESPACES["w"]

w = make_qn("w")


# 1 twip = 635 EMU. Used only to recover a section length when python-docx refuses
# to parse a malformed twips attribute: some editors emit non-integer twips such as
# '1440.0000000000002', which makes python-docx raise ValueError the moment the
# section property is accessed.
_TWIPS_TO_EMU = 635
_PGSZ_ATTR = {"page_width": "w", "page_height": "h"}
_PGMAR_ATTR = {
    "top_margin": "top",
    "right_margin": "right",
    "bottom_margin": "bottom",
    "left_margin": "left",
}


def section_length_emu(section, attr: str) -> int:
    """Read a section length (``page_width`` / ``page_height`` / ``*_margin``) in
    EMU, tolerating malformed measure attributes.

    python-docx raises ``ValueError`` on access when a twips attribute is a
    non-integer decimal. In that case parse the raw twips off the ``w:pgSz`` /
    ``w:pgMar`` element and convert to EMU ourselves; fall back to 0 only when the
    element/attribute is genuinely absent or unparseable.
    """
    try:
        value = getattr(section, attr)
        return int(value) if value is not None else 0
    except (ValueError, TypeError):
        pass
    if attr in _PGSZ_ATTR:
        element = section._sectPr.find(w("pgSz"))
        local = _PGSZ_ATTR[attr]
    else:
        element = section._sectPr.find(w("pgMar"))
        local = _PGMAR_ATTR[attr]
    if element is None:
        return 0
    raw = element.get(w(local))
    try:
        return round(float(raw) * _TWIPS_TO_EMU)
    except (TypeError, ValueError):
        return 0


def section_content_width_emu(section) -> int:
    """Printable content width in EMU: page width minus left/right margins, robust
    to malformed margin attributes (see :func:`section_length_emu`)."""
    return (
        section_length_emu(section, "page_width")
        - section_length_emu(section, "left_margin")
        - section_length_emu(section, "right_margin")
    )


_PGMAR_LOCALS = ("top", "right", "bottom", "left", "header", "footer", "gutter")
_PGSZ_LOCALS = ("w", "h")


def sanitize_section_measures(doc) -> int:
    """Round any non-integer twips on every section's ``w:pgMar`` / ``w:pgSz`` to an
    integer, in place. Returns the count of attributes repaired.

    Some editors emit non-integer twips such as ``'1440.0000000000002'``. python-docx
    parses these measures with ``int()`` and raises ``ValueError`` the moment *any*
    code touches the section measure - including its own internals (e.g.
    ``Document.add_table`` derives the table width from ``section.left_margin``). A
    defensive read at our own call sites is therefore not enough; the value has to be
    repaired on the element so the whole python-docx surface stays usable. The
    rounding is sub-twip (< 1/1440 inch), so the page geometry is visually identical;
    it only removes a malformed value python-docx refuses to parse.
    """
    repaired = 0
    for sectPr in doc.element.body.iter(w("sectPr")):
        for tag, locals_ in ((w("pgMar"), _PGMAR_LOCALS), (w("pgSz"), _PGSZ_LOCALS)):
            element = sectPr.find(tag)
            if element is None:
                continue
            for local in locals_:
                key = w(local)
                raw = element.get(key)
                if raw is None:
                    continue
                try:
                    int(raw)
                    continue  # already a clean integer twips value
                except ValueError:
                    pass
                try:
                    element.set(key, str(round(float(raw))))
                    repaired += 1
                except (TypeError, ValueError):
                    pass
    return repaired


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

# Leading token of a ``w:instrText`` TOC field instruction. This is a FIELD CODE
# (ECMA-376), not a brand word: ``TOC`` is the literal keyword of the table-of-
# contents/figures/tables field instruction in every language Word writes, so
# matching it is language-invariant (the PRIMARY structural signal).
TOC_INSTR_PREFIX = "TOC"

# Regex that pulls the ``\c "<seq>"`` switch argument out of a TOC field
# instruction OPAQUELY. A caption index (``TOC \c "<seq>"``) exposes its caption
# SEQ identifier here; the argument is captured verbatim and never interpreted
# (the SEQ name is data carried into the profile, NEVER a matching rule, so the
# recognition is language-invariant whatever the template's SEQ name happens to
# be). A bare ``TOC`` (an outline table of contents, e.g. ``TOC \o "1-3"``) has
# no ``\c`` and yields ``None``.
_TOC_C_SWITCH_RE = re.compile(r'\\c\s+"([^"]*)"')

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
    its own - it is handled separately as the optional TOC *heading* that may
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
        if (
            dln == "instrText"
            and d.text
            and d.text.strip().startswith(TOC_INSTR_PREFIX)
        ):
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


def is_toc_present(doc) -> bool:
    """Return True if the document body contains a real (strong) TOC region.

    Counts ANY strong TOC field, including a caption index / table-of-figures
    (``TOC \\c``). Used by :func:`refresh_toc` (which marks every TOC field dirty).
    """
    body = doc.element.body
    return any(_element_holds_strong_toc(child) for child in body)


def is_outline_toc_present(doc) -> bool:
    """Return True if the body carries an OUTLINE TOC field (``TOC \\o``).

    Stricter than :func:`is_toc_present`: a caption index / table-of-figures
    (``TOC \\c``) does NOT count. An authored ``toc`` block defers only to an
    outline TOC of the same kind it would itself author, so a shell that ships only
    a table-of-figures still gets its requested table of contents.
    """
    body = doc.element.body
    return any(
        not _is_sectpr(child) and _outline_toc_instruction(child) is not None
        for child in body
    )


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
        It starts at the first strong anchor - extended back by one paragraph if a
        lone contents-word heading (``Contents``/``Sommario``) immediately precedes
        it (that is the TOC heading) - and ends at the **last** strong anchor.
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

    # Strong TOC anchors only - these are the structural proof of a TOC/index.
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
            if (
                _p_text(children[j]).strip() == ""
                and _local_name(children[j].tag) == "p"
            ):
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
        # Extend the span to cover the closing ``end`` fldChar of any TOC/index
        # complex field whose ``begin`` falls inside the span. A table-of-figures /
        # table-of-tables field commonly ends one paragraph PAST its last styled
        # entry (the lone ``w:fldChar end`` sits in an otherwise-empty paragraph);
        # without this the closing paragraph is misclassified as body and a body
        # clear would sever the field, leaving an unremovable orphan index. The
        # span is anchored on the field code (``TOC``), never on a brand word.
        for f in _toc_field_begins(children):
            bi, ei = f["begin_index"], f["end_index"]
            if bi is None or ei is None:
                continue
            if toc_start <= bi <= toc_end and ei > toc_end:
                toc_end = ei
        toc_indices = {
            i for i in range(toc_start, toc_end + 1) if not _is_sectpr(children[i])
        }

    # First Heading-1 anywhere (the no-TOC cover/body boundary).
    first_h1 = next(
        (
            i
            for i, el in enumerate(children)
            if not _is_sectpr(el) and _child_starts_body(el)
        ),
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
                {
                    "index": i,
                    "tag": _local_name(el.tag),
                    "region": None,
                    "is_sectpr": True,
                    "holds_sectpr": False,
                }
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
# Field / index inventory (the deterministic facts the model reasons over and
# the validator binds to - plan §4 "fields/indexes inventory")
# ---------------------------------------------------------------------------
def _toc_seq_id(instr: str) -> Optional[str]:
    """Return the ``\\c "<seq>"`` switch argument of a TOC field, or None.

    Captured OPAQUELY (verbatim) from the field instruction; never interpreted.
    A caption index (``TOC \\c "<seq>"``) yields its SEQ name; an outline
    ``TOC \\o`` yields ``None``. Language-invariant: the argument is whatever SEQ
    identifier the template author chose, carried into the profile as data, not a
    matching rule.
    """
    m = _TOC_C_SWITCH_RE.search(instr or "")
    return m.group(1) if m else None


def _toc_field_begins(children: list) -> list[dict]:
    """Find every top-level TOC/index complex field and where its span begins/ends.

    Walks ``fldChar``/``instrText`` across the whole body in document order with a
    nesting stack so each TOC instruction is attributed to its OWN enclosing
    field, then maps the field's ``begin``/``end`` fldChars back to the top-level
    body child that contains them. A stacked index region (a table-of-tables, then
    a table-of-figures, each its own ``TOC \\c`` field) yields one entry per field.

    Returns a list of ``{"begin_index", "end_index", "instr", "seq_id"}`` (top-
    level body-child indices), in document order. Robust to a field whose begin and
    end live in different paragraphs (the common case for a multi-line index).
    """
    out: list[dict] = []
    # Stack frames: {"begin_owner", "instr"} for the innermost open field. The
    # walk is done one top-level child at a time so each fldChar/instrText is
    # attributed to the child currently being iterated (lxml element identity is
    # not stable across separate ``iter`` passes, so a global element->owner map
    # cannot be keyed on ``id`` - the owner index is taken from the iteration).
    stack: list[dict] = []
    for i, ch in enumerate(children):
        if _is_sectpr(ch):
            continue
        for kind, el in iter_complex_field_events(ch):
            if kind == "begin":
                stack.append({"begin_owner": i, "instr": ""})
            elif kind == "end":
                if stack:
                    frame = stack.pop()
                    instr = frame.get("instr", "").strip()
                    if instr.startswith(TOC_INSTR_PREFIX):
                        out.append(
                            {
                                "begin_index": frame.get("begin_owner"),
                                "end_index": i,
                                "instr": instr,
                                "seq_id": _toc_seq_id(instr),
                            }
                        )
            elif kind == "instr":
                if stack:
                    stack[-1]["instr"] += el.text or ""
    # Order by the begin child index so the surfaced ids are stable.
    out.sort(
        key=lambda f: f["begin_index"] if f["begin_index"] is not None else 1 << 30
    )
    return out


def inventory_fields(doc) -> list[dict]:
    """Surface every top-level TOC/index complex field as a stable-id inventory.

    One entry per TOC/index field (outline table-of-contents AND each ``\\c``
    caption index), keyed by ``field.<begin-child-index>`` - a deterministic id
    the generator can recompute from the live tree because the cover/index front
    matter keeps its body-child positions across generation (only the freeform
    body region is rewritten). Each entry::

        {"id": "field.18", "seq_id": "<seq>" | None, "instr": "TOC \\h \\z \\c ...",
         "begin_index": 18, "end_index": 26}

    ``seq_id`` is the OPAQUE ``\\c`` switch argument (None for an outline TOC). The
    model binds an ``index_ref`` to one of these ids; the validator checks
    membership; the generator reconciles by ``seq_id``. Brand- and language-
    agnostic: the only signal is the ``TOC`` field code.
    """
    children = list(doc.element.body)
    out: list[dict] = []
    for f in _toc_field_begins(children):
        bi = f["begin_index"]
        if bi is None:
            continue
        out.append(
            {
                "id": f"field.{bi}",
                "seq_id": f["seq_id"],
                "instr": f["instr"][:200],
                "begin_index": bi,
                "end_index": f["end_index"],
            }
        )
    return out


def _field_span_indices(doc, field_id: str) -> Optional[list[int]]:
    """Resolve a ``field.<begin-index>`` id to the live span of body-child indices.

    Recomputes the field inventory against the live tree and returns the
    ``[begin .. end]`` top-level body-child index range for the matching field,
    or None when the id no longer maps. Used by index REMOVE so a stale caption
    index can be deleted as one block.
    """
    for f in inventory_fields(doc):
        if f["id"] == field_id:
            begin = f["begin_index"]
            end = f.get("end_index")
            if begin is None:
                return None
            if end is None or end < begin:
                end = begin
            return list(range(begin, end + 1))
    return None


def _index_field_remove_indices(doc, field_id: str) -> set[int]:
    """Return the body-child indices a REMOVE of ``field_id`` would delete.

    The field's ``[begin .. end]`` span PLUS its **introducing heading** - the
    orphan-index-heading fix. When a caption index is removed (e.g. a stale
    table-of-tables / table-of-figures), the heading paragraph that introduces it
    (e.g. an ``Indice delle Tabelle`` / ``Index of Figures`` line) would otherwise
    be left behind as an orphan heading pointing at nothing. So this also removes
    the introducing heading and any blank separators between it and the span.

    The heading is identified **STRUCTURALLY, never by literal**: it is the
    paragraph immediately preceding the field span that

      - is a non-field paragraph (carries no strong TOC marker of its own), and
      - is itself classified in the **toc region** (the stacked-index front matter)
        by :func:`classify_body_children`, and
      - is non-empty.

    The toc-region membership is the structural proof that this preceding line is
    index front matter (an index title) rather than ordinary body content; it is
    language- and brand-invariant (no word list, no length cutoff). Resolved
    against the LIVE tree; never includes the final ``sectPr``.
    """
    span = _field_span_indices(doc, field_id)
    if not span:
        return set()
    children = list(doc.element.body)
    # Structural region map: only a paragraph the classifier puts in the toc region
    # qualifies as the introducing index heading.
    toc_indices = {
        c["index"] for c in classify_body_children(doc) if c.get("region") == "toc"
    }
    to_remove = set(span)
    j = min(span) - 1
    heading_at: Optional[int] = None
    while j >= 0 and not _is_sectpr(children[j]):
        if _local_name(children[j].tag) != "p":
            break
        if j not in toc_indices:
            break  # crossed out of the index front matter -> body content, stop
        txt = _p_text(children[j]).strip()
        if txt == "":
            j -= 1
            continue
        if not _element_holds_strong_toc(children[j]):
            heading_at = j
        break
    if heading_at is not None:
        to_remove.add(heading_at)
        for k in range(heading_at + 1, min(span)):
            if (
                _local_name(children[k].tag) == "p"
                and _p_text(children[k]).strip() == ""
            ):
                to_remove.add(k)
    return {i for i in to_remove if i < len(children) and not _is_sectpr(children[i])}


def remove_index_fields(doc, field_ids: list[str]) -> set[str]:
    """REMOVE several orphan caption-index blocks in one shift-safe pass.

    Each ``field.<begin-index>`` id is position-based, so deleting one index would
    invalidate the ids of those after it. This resolves every target's child
    elements against the ORIGINAL live tree FIRST (lxml element references survive
    sibling removal), then deletes them all. Returns the set of field ids that
    actually removed something.
    """
    body = doc.element.body
    children = list(body)
    removed_ids: set[str] = set()
    elements: list = []
    for fid in field_ids:
        idxs = _index_field_remove_indices(doc, fid)
        if not idxs:
            continue
        removed_ids.add(fid)
        elements.extend(children[i] for i in sorted(idxs))
    # De-dup element references while preserving order, then remove.
    seen: set[int] = set()
    for el in elements:
        if id(el) in seen:
            continue
        seen.add(id(el))
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)
    return removed_ids


def inventory_regions(doc) -> list[dict]:
    """Surface the template's region inventory (stable ids the model binds to).

    Combines the ordered top-level skeleton regions (``region.cover`` /
    ``region.toc`` / ``region.body``, present only when the template has them) with
    each derived-index block (``region.field.<i>`` for a ``\\c`` caption index, so
    a ``demo_classification`` / ``sections`` ref can target a specific index span).
    Each entry is ``{"id": <region_ref>, "kind": <open advisory token>}``.

    Deterministic and recomputable at generate time; the ids never encode brand
    words, only structural positions.
    """
    out: list[dict] = []
    classes = classify_body_children(doc)
    seen_regions: set[str] = set()
    for c in classes:
        region = c.get("region")
        if region in ("cover", "toc", "body") and region not in seen_regions:
            seen_regions.add(region)
            out.append({"id": f"region.{region}", "kind": region})
    for f in inventory_fields(doc):
        if f.get("seq_id"):
            out.append({"id": f"region.{f['id']}", "kind": "caption_index"})
    return out


# ---------------------------------------------------------------------------
# Numbering inventory (the list-definition facts list-role nomination binds to).
#
# A real bulleted / numbered list in OOXML is a paragraph carrying a
# ``w:pPr/w:numPr`` that references a ``w:num`` in ``word/numbering.xml`` (which
# in turn references a ``w:abstractNum`` whose per-level ``w:numFmt`` decides the
# *family*: ``bullet`` vs decimal/roman/etc. = ``number``). The brand's list
# PARAGRAPH styles carry their own ``w:numPr`` in the style definition, so a list
# is authored by applying the style alone - but python-docx's ``add_paragraph``
# does NOT inherit a style's ``w:numPr`` onto the paragraph, so generation must
# re-assert ``w:numPr`` explicitly (see ``generate._write_list_items``).
#
# These helpers are purely STRUCTURAL and brand-agnostic: the family is read from
# the ``w:numFmt`` field code, never from a style name. They never raise on a
# document with no numbering part (a template legitimately may have none).
# ---------------------------------------------------------------------------
def _numbering_root(doc):
    """Return the ``w:numbering`` element of the doc's numbering part, or None.

    A template with no list numbering has no numbering part; that is a legitimate
    absence (returns None), never an error.
    """
    try:
        part = doc.part.numbering_part
    except (KeyError, AttributeError, ValueError):
        return None
    if part is None:
        return None
    return getattr(part, "element", None)


def _num_to_abstract(root) -> dict[str, str]:
    """Map ``w:num/@w:numId`` -> ``w:abstractNumId/@w:val`` for the numbering part."""
    out: dict[str, str] = {}
    if root is None:
        return out
    for num in root.findall(w("num")):
        nid = num.get(w("numId"))
        an = num.find(w("abstractNumId"))
        if nid is not None and an is not None:
            val = an.get(w("val"))
            if val is not None:
                out[nid] = val
    return out


def _abstract_level_formats(root) -> dict[str, dict[int, str]]:
    """Map ``abstractNumId`` -> ``{ilvl: numFmt}`` for the numbering part.

    ``numFmt`` is the OOXML field code (``bullet`` / ``decimal`` / ``upperRoman``
    / ...), language-invariant; it is the structural proof of the list *family*.
    """
    out: dict[str, dict[int, str]] = {}
    if root is None:
        return out
    for an in root.findall(w("abstractNum")):
        aid = an.get(w("abstractNumId"))
        if aid is None:
            continue
        levels: dict[int, str] = {}
        for lvl in an.findall(w("lvl")):
            try:
                ilvl = int(lvl.get(w("ilvl")) or 0)
            except (TypeError, ValueError):
                continue
            nf = lvl.find(w("numFmt"))
            if nf is not None and nf.get(w("val")):
                levels[ilvl] = nf.get(w("val"))
        out[aid] = levels
    return out


def num_family_for(doc, num_id: str, ilvl: int = 0) -> Optional[str]:
    """Return ``"bullet"`` or ``"number"`` for a ``w:numId`` at level ``ilvl``.

    Resolves ``numId -> abstractNumId -> abstractNum/lvl[ilvl]/numFmt`` and maps
    the OOXML ``numFmt`` to the engine's two list families: ``bullet`` for the
    ``bullet`` format, ``number`` for every counted format (decimal, roman,
    letter, ...). Returns None when the numbering part has no such id/level (so a
    caller can skip rather than guess). Falls back to level 0's format when the
    requested ``ilvl`` is not explicitly defined.
    """
    root = _numbering_root(doc)
    if root is None:
        return None
    aid = _num_to_abstract(root).get(str(num_id))
    if aid is None:
        return None
    levels = _abstract_level_formats(root).get(aid) or {}
    fmt = levels.get(ilvl) or levels.get(0)
    if not fmt:
        return None
    return "bullet" if fmt == "bullet" else "number"


def style_num_binding(style) -> Optional[tuple[str, int]]:
    """Return ``(numId, ilvl)`` if a paragraph style's definition carries a
    ``w:pPr/w:numPr``, else None.

    The ``w:ilvl`` defaults to 0 when the style omits it (the common case for a
    per-level list style that pins only its ``w:numId``). Reads the style's own
    XML element; never raises on a style without a ``w:pPr``.
    """
    el = getattr(style, "element", None)
    if el is None:
        return None
    pPr = el.find(w("pPr"))
    if pPr is None:
        return None
    numPr = pPr.find(w("numPr"))
    if numPr is None:
        return None
    num = numPr.find(w("numId"))
    num_id = num.get(w("val")) if num is not None else None
    if not num_id:
        return None
    il = numPr.find(w("ilvl"))
    try:
        ilvl = int(il.get(w("ilvl")) or il.get(w("val")) or 0) if il is not None else 0
    except (TypeError, ValueError):
        ilvl = 0
    return (num_id, ilvl)


# ---------------------------------------------------------------------------
# Per-level numbering FACTS capture (Cluster D3, DOCX-ONLY).
#
# These readers surface the per-level facts a list role re-applies so a generated
# list looks like the template's lists: the ``w:numFmt`` field code (the closed enum
# decimal / bullet / lowerLetter / ...), the ``w:lvlText`` level format string (e.g.
# ``"%1."`` or a bullet glyph), and the level's ``w:ind`` indentation (twips). Every
# value is read VERBATIM off the shell's OWN ``w:abstractNum/w:lvl``: the engine never
# synthesizes a numFmt/lvlText/indent - it can only REFERENCE the shell's numbering by
# id and at most CLONE an existing ``w:abstractNum`` (see ``generate``). A level the
# abstractNum does not declare is simply absent (it inherits per OOXML at render time);
# this reader records ONLY the declared levels. Brand-agnostic and crash-safe.
# ---------------------------------------------------------------------------
# The four ``w:ind`` indentation attributes captured per level (twips). Each is an
# INDEPENDENT key, absent when the level does not declare it (so apply re-asserts only
# the indents the template actually carried).
_NUM_INDENT_ATTRS: tuple[str, ...] = ("left", "right", "firstLine", "hanging")


def _lvl_indent_facts(lvl) -> dict:
    """``{attr: int}`` for every ``w:lvl/w:pPr/w:ind`` twips attribute the level declares.

    Reads the level's own ``w:pPr/w:ind`` (``left`` / ``right`` / ``firstLine`` /
    ``hanging``) as twips. A missing element/attribute or a malformed (non-integer)
    value contributes nothing to that axis (fail-soft, like geometry capture)."""
    out: dict = {}
    if lvl is None:
        return out
    ppr = lvl.find(w("pPr"))
    if ppr is None:
        return out
    ind = ppr.find(w("ind"))
    if ind is None:
        return out
    for attr in _NUM_INDENT_ATTRS:
        val = ind.get(w(attr))
        if val is None:
            continue
        try:
            out[attr] = int(val)
        except (TypeError, ValueError):
            continue
    return out


def _abstract_num_per_level_facts(root, abstract_num_id: str) -> dict[int, dict]:
    """Map ``ilvl -> {numFmt, lvlText, indent}`` for one ``w:abstractNum``, or ``{}``.

    Reads ``w:abstractNum[@w:abstractNumId=abstract_num_id]/w:lvl[@w:ilvl]`` and, per
    declared level, surfaces:

      - ``numFmt``: the ``w:numFmt@w:val`` field code (the closed enum; absent when the
        level declares none);
      - ``lvlText``: the ``w:lvlText@w:val`` level format string VERBATIM (kept
        byte-for-byte, including an empty bullet string - no normalization);
      - ``indent``: the per-level ``w:ind`` twips (:func:`_lvl_indent_facts`).

    Each fact is recorded only when the level declares it; an absent fact is simply
    omitted. Returns ``{}`` when ``abstract_num_id`` does not exist in the numbering
    part (so a caller skips rather than guesses). Crash-safe / brand-agnostic."""
    out: dict[int, dict] = {}
    if root is None or abstract_num_id is None:
        return out
    target = None
    for an in root.findall(w("abstractNum")):
        if an.get(w("abstractNumId")) == str(abstract_num_id):
            target = an
            break
    if target is None:
        return out
    for lvl in target.findall(w("lvl")):
        try:
            ilvl = int(lvl.get(w("ilvl")) or 0)
        except (TypeError, ValueError):
            continue
        facts: dict = {}
        nf = lvl.find(w("numFmt"))
        if nf is not None and nf.get(w("val")) is not None:
            facts["numFmt"] = nf.get(w("val"))
        lt = lvl.find(w("lvlText"))
        if lt is not None and lt.get(w("val")) is not None:
            facts["lvlText"] = lt.get(w("val"))
        indent = _lvl_indent_facts(lvl)
        if indent:
            facts["indent"] = indent
        if facts:
            out[ilvl] = facts
    return out


def num_per_level_facts(doc, num_id: str) -> Optional[dict]:
    """Resolve a ``w:numId`` to its captured per-level numbering facts, or ``None``.

    Resolves ``numId -> abstractNumId`` (via :func:`_num_to_abstract`) then reads the
    abstractNum's per-level facts (:func:`_abstract_num_per_level_facts`). Returns a dict
    with:

      - ``num_id``: the referenced ``w:numId`` (str, a SYMBOLIC reference the generator
        re-asserts verbatim - never invented);
      - ``abstract_num_id``: the resolved ``w:abstractNumId`` (str, the def the generator
        clones by id when the output needs it present);
      - ``per_level_facts``: ``{ilvl -> {numFmt, lvlText, indent}}`` (the declared levels
        only).

    Returns ``None`` when the numbering part is absent, the id is undefined, or the
    abstractNum declares no per-level fact at all (so a role with no captured numbering
    leaves ``appearance.numbering`` absent - the byte-identical no-capture path)."""
    root = _numbering_root(doc)
    if root is None or num_id is None:
        return None
    aid = _num_to_abstract(root).get(str(num_id))
    if aid is None:
        return None
    per_level = _abstract_num_per_level_facts(root, aid)
    if not per_level:
        return None
    return {
        "num_id": str(num_id),
        "abstract_num_id": str(aid),
        "per_level_facts": per_level,
    }


def clone_abstract_num(root, abstract_num_id: str):
    """A DEEP COPY of the shell's ``w:abstractNum[@w:abstractNumId=abstract_num_id]``
    element, or ``None`` when the numbering part has no such definition.

    The copy is VERBATIM (``copy.deepcopy`` of the live element): the engine clones the
    shell's own definition by id, never minting a new one. Used by the generator to
    ensure a referenced abstractNum is present in the output's numbering part."""
    if root is None or abstract_num_id is None:
        return None
    for an in root.findall(w("abstractNum")):
        if an.get(w("abstractNumId")) == str(abstract_num_id):
            return copy.deepcopy(an)
    return None


# ---------------------------------------------------------------------------
# Demo-region detection (purely structural; NO hardcoded marker phrases).
# ---------------------------------------------------------------------------
def detect_demo_region(doc) -> dict:
    """Detect the template's demo / instruction body region structurally.

    No literal marker phrases (the old ``DEMO_MARKERS`` are gone, plan §5): demo
    handling is owned by comprehension ``demo_classification`` plus the structural
    freeform-body clear, both of which are language-invariant. This function only
    surfaces the structural facts:

      - ``present`` - whether the template has any freeform body region at all
        (everything after the cover/TOC), which is the candidate demo region;
      - ``start_style_id`` / ``start_text`` - the first body-region Heading-1
        paragraph's style id and its *own* captured text (THIS template's own
        placeholder text, never a fixed phrase). ``no_residual_template_text``
        compares the produced document against this captured text, so the
        anti-residual guard stays per-template and language-agnostic.

    ``instruction_markers`` is retained for shape back-compat but is always empty:
    detection no longer matches any global phrase list.
    """
    classes = classify_body_children(doc)
    body_children = [c for c in classes if c["region"] == "body"]
    body = doc.element.body
    children = list(body)

    # First body-region Heading-1 paragraph -> the demo start anchor (structural).
    start_style_id = None
    start_text = None
    for c in body_children:
        el = children[c["index"]]
        if _local_name(el.tag) == "p" and _style_is_heading1(_p_style_val(el)):
            start_style_id = _p_style_val(el)
            start_text = _p_text(el)[:200]
            break

    return {
        "present": bool(body_children),
        "start_style_id": start_style_id,
        "start_text": start_text,
        "instruction_markers": [],
    }


# ---------------------------------------------------------------------------
# Body clearing
# ---------------------------------------------------------------------------
def clear_body_region(
    doc,
    structure: Optional[dict] = None,
    *,
    preserve_cover: bool = True,
    preserve_toc: bool = True,
) -> None:
    """Remove ONLY the body region, preserving the cover and TOC regions.

    Keeps the cover block(s) and the TOC sdt/region and the final ``sectPr``;
    removes only the in-between/after demo body content. python-docx ``add_*`` then
    appends new content into the (now-empty) body region, immediately before the
    ``sectPr`` - which is exactly the right place.

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


def prune_leading_empty_body_artifacts(doc) -> int:
    """Remove empty body-region artifacts before generated body content is written.

    ``clear_body_region`` intentionally preserves section-break paragraphs because
    they may carry page geometry/header-footer data. In real front matter, caption
    indexes can later be reconciled away and leave behind blank TOC separators plus
    demo-body section breaks immediately before the new content insertion point.
    If kept there, Word/LibreOffice renders blank pages before the generated body.

    This cleanup runs after index reconciliation and before body writing, so any
    leading empty body paragraph is still template/demo scaffolding, not authored
    content. Non-empty body content stops the prune.
    """
    body = doc.element.body
    children = list(body)
    to_remove = []
    for c in classify_body_children(doc):
        if c["is_sectpr"]:
            continue
        if c["region"] != "body":
            continue
        el = children[c["index"]]
        if _local_name(el.tag) == "p" and _p_text(el).strip() == "":
            to_remove.append(el)
            continue
        break

    removed = 0
    for el in to_remove:
        if el.getparent() is body:
            body.remove(el)
            removed += 1
    return removed


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
    field - in which case nothing is written). Never duplicates the TOC.
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
    for kind, el in iter_complex_field_events(body):
        if kind == "begin":
            stack.append(el)
        elif kind == "end":
            if stack:
                stack.pop()
        elif kind == "instr":
            text = (el.text or "").strip()
            if stack and text.startswith(TOC_INSTR_PREFIX):
                begin = stack[-1]
                if begin.get(w("dirty")) != "true":
                    begin.set(w("dirty"), "true")
                    marked += 1
    return marked


def append_outline_toc_field(doc, *, max_level: int = 3):
    """Append a native, updateable outline TOC field as ONE body paragraph.

    For an authored ``toc`` block when the shell carries NO structural TOC to defer
    to. The single-paragraph shape (begin[dirty] / instrText / separate / cache /
    end) is exactly what :func:`_outline_toc_instruction` recognizes, so
    :func:`refresh_visible_outline_toc_cache` fills its visible cache from the
    generated headings and :func:`refresh_toc` marks it dirty and sets
    ``updateFields`` (both run after the body is built). The paragraph is inserted
    before the final ``w:sectPr`` so it lands in body-flow order. Returns the
    appended paragraph element.
    """
    level = max(1, int(max_level or 3))
    instr = f' TOC \\o "1-{level}" \\h \\z \\u '
    p = OxmlElement("w:p")
    p.append(_field_run("begin", dirty=True))
    p.append(_instr_run(instr))
    p.append(_field_run("separate"))
    p.append(_text_run("Right-click to update the table of contents."))
    p.append(_field_run("end"))
    body = doc.element.body
    sectpr = body.find(w("sectPr"))
    if sectpr is not None:
        sectpr.addprevious(p)
    else:
        body.append(p)
    return p


def refresh_visible_outline_toc_cache(
    doc,
    headings: "list[tuple[int, str]] | list[tuple[int, str, object]]",
) -> int:
    """Rewrite the visible cache of outline TOC fields from generated headings.

    Word/LibreOffice store a field *result* alongside the TOC field code. Marking
    the field dirty is enough for Word-on-open, but headless LibreOffice often
    renders the stale cached result when exporting to PDF. This function keeps the
    outline TOC field updateable while replacing the visible cache with the
    generated document's current headings, so visual audit does not show template
    sample entries.

    Each heading is ``(level, text)`` or ``(level, text, w:p element)``. When every
    heading carries its live generated paragraph element AND at least one outline
    TOC field's old cache is well-formed enough to harvest, the cache is rewritten
    in the *Word-faithful* shape: bookmarks are authored around the generated
    heading paragraphs and each cache entry becomes a paragraph whose ``pPr`` is
    deep-copied from the template's own old cached entry of the SAME level,
    carrying a ``w:hyperlink`` to the heading's bookmark plus a nested dirty
    ``PAGEREF`` field (no cached page number). Otherwise - legacy 2-tuple callers,
    a malformed cache, or detached heading elements - the function falls back, per
    field and all-or-nothing, to the simple plain-text rewrite, and authors ZERO
    bookmarks. Documents without an outline TOC are never mutated.

    An EMPTY ``headings`` list rebuilds every outline TOC cache empty (the field
    code survives, dirty, for Word to recompute on open) so the template's demo
    entries never outlive a heading-less generation - the same "stale derived
    index" contract as the caption-index sibling.
    """
    headings3 = _normalize_outline_headings(headings)
    plain_headings = [(level, text) for level, text, _ in headings3]

    body = doc.element.body
    children = list(body)
    spans = _toc_field_begins(children)

    # PHASE 1 - PLAN (strictly read-only). One plan per outline TOC field;
    # ``rich`` is decided per field, once, before any mutation. Fail-closed: a
    # field is rewritten either in the FULL Word-faithful shape or by the exact
    # current plain writer - no hybrid output exists by construction.
    plans: list[dict] = []
    for i, child in enumerate(children):
        if _is_sectpr(child):
            continue
        instr = _outline_toc_instruction(child)
        if instr is None:
            continue
        tag = _local_name(child.tag)
        if tag == "sdt":
            plan = _plan_sdt_outline_rewrite(child, instr, i)
            if plan is not None:
                plans.append(plan)
        elif tag == "p":
            plans.append(
                _plan_paragraph_outline_rewrite(child, instr, i, children, spans)
            )
    if not plans:
        return 0

    # PHASE 2 - bookmark authoring (the only doc-level mutation outside the
    # fields). Bookmarks exist ONLY when an outline TOC is present (a plan exists)
    # AND at least one field takes the rich shape AND every heading carries its
    # live generated paragraph; otherwise every plan is demoted to the plain
    # writer and zero bookmarks are authored (today's bytes exactly).
    if headings3:
        rich_mode = any(p["rich"] for p in plans) and all(
            _is_live_paragraph(para, doc) for _, _, para in headings3
        )
        if rich_mode:
            bookmark_names = _author_heading_bookmarks(doc, headings3)
        else:
            bookmark_names = []
            for plan in plans:
                plan["rich"] = False
    else:
        # Heading-less generation: rebuild every RESOLVABLE cache empty. A
        # span-resolved bare-paragraph plan keeps its rich splice so the FULL
        # span collapses to [field-start, field-end] (zero entries, zero
        # bookmarks); an UNRESOLVED bare span is left untouched, because the
        # begin-only plain rewrite would close the field in the begin
        # paragraph and orphan the original end fldChar (malformed XML is
        # strictly worse than a stale cache). SDT plans rebuild empty via
        # either writer, so they always stay in.
        bookmark_names = []
        plans = [p for p in plans if p["rich"] or p["kind"] == "sdt"]

    # PHASE 3 - MUTATE, last-first (descending top-level position) so a
    # bare-paragraph span splice never shifts an earlier plan's body position.
    # Plans hold element handles captured in phase 1; indices are never
    # re-derived mid-mutation.
    rewritten = 0
    for plan in sorted(plans, key=lambda p: p["top_index"], reverse=True):
        if plan["kind"] == "sdt":
            if plan["rich"]:
                _apply_rich_sdt_outline_plan(plan, headings3, bookmark_names)
                rewritten += 1
            elif _rewrite_sdt_outline_toc_cache(
                plan["el"], plan["instr"], plain_headings
            ):
                rewritten += 1
        else:
            if plan["rich"]:
                _apply_rich_paragraph_outline_plan(
                    plan, body, headings3, bookmark_names
                )
            else:
                _rewrite_paragraph_outline_toc_cache(
                    plan["el"], plan["instr"], plain_headings
                )
            rewritten += 1
    return rewritten


def _normalize_outline_headings(headings) -> list[tuple[int, str, object]]:
    """Normalize heading items to ``(level, text, paragraph-or-None)`` 3-tuples.

    Items of length 2 carry no paragraph element (legacy callers); items of any
    other arity, or with an unusable level, are dropped (fail-closed - never
    raise). Empty-text items are dropped exactly as before.
    """
    out: list[tuple[int, str, object]] = []
    for item in headings or []:
        try:
            arity = len(item)
        except TypeError:
            continue
        if arity == 3:
            level, text, para = item
        elif arity == 2:
            level, text = item
            para = None
        else:
            continue
        try:
            lvl = max(1, int(level or 1))
        except (TypeError, ValueError):
            continue
        txt = str(text).strip()
        if txt:
            out.append((lvl, txt, para))
    return out


def _is_live_paragraph(para, doc) -> bool:
    """True when ``para`` is an element attached to ``doc``'s main document part."""
    if para is None:
        return False
    try:
        return para.getroottree().getroot() is doc.element
    except Exception:
        return False


def _plan_sdt_outline_rewrite(sdt, instr: str, top_index: int) -> Optional[dict]:
    """Plan (read-only) the rewrite of an SDT-wrapped outline TOC field.

    Returns None when the SDT carries no ``sdtContent`` or no field paragraph -
    exactly the cases today's writer bails on without counting. Otherwise the
    plan is ``rich`` when the old cache could be harvested, else plain.
    """
    content = sdt.find(w("sdtContent"))
    if content is None:
        return None
    paras = [el for el in list(content) if _local_name(el.tag) == "p"]
    field_idx = next(
        (i for i, p in enumerate(paras) if _outline_toc_instruction(p) is not None),
        None,
    )
    if field_idx is None:
        return None
    plan: dict = {
        "kind": "sdt",
        "el": sdt,
        "instr": instr,
        "top_index": top_index,
        "rich": False,
        "content": content,
        "old_field_para": paras[field_idx],
        "stale_paras": paras[field_idx:],
    }
    try:
        level_pprs, first_ppr = _harvest_outline_entry_pprs(paras[field_idx:])
        plan.update(rich=True, level_pprs=level_pprs, first_ppr=first_ppr)
    except Exception:
        plan["rich"] = False
    return plan


def _plan_paragraph_outline_rewrite(
    p, instr: str, top_index: int, children: list, spans: list[dict]
) -> dict:
    """Plan (read-only) the rewrite of a bare-paragraph outline TOC field.

    The field's full top-level span is resolved from the ``_toc_field_begins``
    inventory (the begin paragraph is ``children[top_index]``). An unresolvable
    or malformed span demotes the plan to the plain writer - which rewrites only
    the begin paragraph, exactly what runs today on such input.
    """
    plan: dict = {
        "kind": "p",
        "el": p,
        "instr": instr,
        "top_index": top_index,
        "rich": False,
    }
    try:
        span = next(
            (
                f
                for f in spans
                if f.get("seq_id") is None and f.get("begin_index") == top_index
            ),
            None,
        )
        if span is None:
            return plan
        begin_i, end_i = span.get("begin_index"), span.get("end_index")
        if (
            begin_i is None
            or end_i is None
            or end_i < begin_i
            or begin_i < 0
            or end_i >= len(children)
        ):
            return plan
        span_els = children[begin_i : end_i + 1]
        candidates = [el for el in span_els if _local_name(el.tag) == "p"]
        level_pprs, first_ppr = _harvest_outline_entry_pprs(candidates)
        plan.update(
            rich=True, level_pprs=level_pprs, first_ppr=first_ppr, span_els=span_els
        )
    except Exception:
        plan["rich"] = False
    return plan


def _harvest_outline_entry_pprs(candidate_paras) -> "tuple[dict[int, object], object]":
    """Harvest a ``{level: pPr}`` map from the OLD cached TOC entries.

    Candidate ENTRY paragraphs are those carrying at least one non-empty ``w:t``
    (the template-shape begin paragraph has none; a Word-native begin+entry
    paragraph qualifies; a bare end-fldChar paragraph never does). Per candidate
    the level is the trailing digit of its TOC style id (an ECMA structural
    convention, gated by ``_style_is_toc`` so e.g. ``Heading1`` never parses) or,
    failing that, the ascending rank of its ``w:ind/@w:left`` indent among the
    style-less candidates. The FIRST candidate seen for a level (with a real
    ``pPr``) wins - deterministic in document order. ``pPr`` values are live
    element references; callers deepcopy at write time. Returns
    ``(level_pprs, first_candidate_ppr)`` - both possibly empty/None (zero
    candidates is NOT malformed: a freshly authored placeholder field has none).
    """
    entries = [
        p
        for p in candidate_paras
        if any((t.text or "").strip() for t in p.iter(w("t")))
    ]
    if not entries:
        return {}, None
    unstyled_indents = sorted(
        {
            _entry_indent_left(p)
            for p in entries
            if _toc_style_level(_p_style_val(p)) is None
        }
    )
    level_pprs: dict[int, object] = {}
    for p in entries:
        level = _toc_style_level(_p_style_val(p))
        if level is None:
            level = 1 + unstyled_indents.index(_entry_indent_left(p))
        ppr = p.find(w("pPr"))
        if ppr is not None and level not in level_pprs:
            level_pprs[level] = ppr
    return level_pprs, entries[0].find(w("pPr"))


def _toc_style_level(style_val: Optional[str]) -> Optional[int]:
    """Level from a TOC entry style id's trailing digit (``TOC1``..``TOC9``,
    ``Sommario2``, ...), or None. The ``_style_is_toc`` gate keeps non-TOC styles
    (``Heading1``) from parsing; digits outside 1..9 are rejected."""
    if not _style_is_toc(style_val):
        return None
    m = re.search(r"(\d+)$", style_val or "")
    if m is None:
        return None
    level = int(m.group(1))
    return level if 1 <= level <= 9 else None


def _entry_indent_left(p) -> int:
    """``w:pPr/w:ind/@w:left`` as int twips; absent or unparseable -> 0."""
    pPr = p.find(w("pPr"))
    if pPr is None:
        return 0
    ind = pPr.find(w("ind"))
    if ind is None:
        return 0
    try:
        return int(ind.get(w("left")) or 0)
    except (TypeError, ValueError):
        return 0


def _entry_ppr_for_level(level: int, level_pprs: dict, first_ppr, begin_ppr):
    """Resolve the harvested ``pPr`` for an entry level, per the fallback chain:
    exact level -> nearest harvested LOWER level -> the first old entry's pPr ->
    the old field-begin paragraph's pPr (may be None: entry gets no pPr)."""
    if level in level_pprs:
        return level_pprs[level]
    lower = [k for k in level_pprs if k < level]
    if lower:
        return level_pprs[max(lower)]
    if first_ppr is not None:
        return first_ppr
    return begin_ppr


_TOC_BOOKMARK_PREFIX = "_TocBD"


def _author_heading_bookmarks(doc, headings3) -> list[str]:
    """Author ``w:bookmarkStart``/``w:bookmarkEnd`` around each generated heading
    paragraph; return the bookmark names in heading order.

    Names are ``_TocBD000001``-style with a sequential counter skipping any name
    already in the document; numeric ids continue past the document's current
    maximum. Both derive solely from the input document state and heading order -
    never random/time-based - so regenerating the same input reproduces them
    exactly. The collision scan covers the MAIN document part only
    (``doc.element``); bookmarks in header/footer parts are separate stories this
    engine never writes into. ``bookmarkStart`` goes immediately after ``w:pPr``
    (or first when the paragraph has no pPr); ``bookmarkEnd`` is the paragraph's
    last child.
    """
    existing_names: set[str] = set()
    max_id = -1
    for bs in doc.element.iter(w("bookmarkStart")):
        name = bs.get(w("name"))
        if name:
            existing_names.add(name)
        try:
            max_id = max(max_id, int(bs.get(w("id"))))
        except (TypeError, ValueError):
            continue
    names: list[str] = []
    counter = 1
    next_id = max_id + 1
    for _, _, para in headings3:
        name = f"{_TOC_BOOKMARK_PREFIX}{counter:06d}"
        while name in existing_names:
            counter += 1
            name = f"{_TOC_BOOKMARK_PREFIX}{counter:06d}"
        existing_names.add(name)
        counter += 1
        start = OxmlElement("w:bookmarkStart")
        start.set(w("id"), str(next_id))
        start.set(w("name"), name)
        end = OxmlElement("w:bookmarkEnd")
        end.set(w("id"), str(next_id))
        next_id += 1
        pPr = para.find(w("pPr"))
        if pPr is not None:
            pPr.addnext(start)
        else:
            para.insert(0, start)
        para.append(end)
        names.append(name)
    return names


def _apply_rich_sdt_outline_plan(plan: dict, headings3, names: list[str]) -> None:
    """Rewrite an SDT-wrapped outline TOC cache in the Word-faithful shape.

    Same skeleton as the plain SDT writer (remove the old field-and-cache
    paragraphs, append field-start + entries + field-end), but each entry carries
    its per-level harvested ``pPr``, a bookmark hyperlink and a nested PAGEREF.
    """
    content = plan["content"]
    old_field_para = plan["old_field_para"]
    for p in plan["stale_paras"]:
        if p.getparent() is content:
            content.remove(p)
    content.append(_toc_field_start_paragraph(plan["instr"], old_field_para))
    begin_ppr = old_field_para.find(w("pPr"))
    for (level, text, _), name in zip(headings3, names):
        ppr = _entry_ppr_for_level(
            level, plan["level_pprs"], plan["first_ppr"], begin_ppr
        )
        content.append(_outline_toc_entry_paragraph(text, ppr, name))
    content.append(_toc_field_end_paragraph(old_field_para))


def _apply_rich_paragraph_outline_plan(
    plan: dict, body, headings3, names: list[str]
) -> None:
    """Splice a bare-paragraph outline TOC's FULL top-level span with the
    Word-faithful multi-paragraph shape (field-start, entries, field-end).

    The proven caption-index splice pattern: insert the new paragraphs before the
    old begin paragraph, then remove every old span child - skipping the final
    ``w:sectPr``, any paragraph HOLDING an intermediate ``w:pPr/w:sectPr``
    (section geometry must never be deleted), and anything already detached. A
    true single-paragraph field has span length 1 and converges to the same shape.
    """
    begin_el = plan["el"]
    begin_ppr = begin_el.find(w("pPr"))
    new_paras = [_toc_field_start_paragraph(plan["instr"], begin_el)]
    for (level, text, _), name in zip(headings3, names):
        ppr = _entry_ppr_for_level(
            level, plan["level_pprs"], plan["first_ppr"], begin_ppr
        )
        new_paras.append(_outline_toc_entry_paragraph(text, ppr, name))
    new_paras.append(_toc_field_end_paragraph(begin_el))
    for np in new_paras:
        begin_el.addprevious(np)
    for old in plan["span_els"]:
        if not _is_sectpr(old) and not _holds_sectpr(old) and old.getparent() is body:
            body.remove(old)


def kept_caption_index_seq_ids(doc) -> list[str]:
    r"""Distinct ``\c`` sequence ids of every caption index present in the body.

    The deterministic complement to the emitted-captions map: a kept caption
    index whose sequence received NO captions this run must still have its
    visible cache cleared, otherwise the template's demo entries survive into
    the generated document (the "stale derived index" defect class). Document
    order, first occurrence wins; malformed spans are skipped (same validity
    filter as :func:`refresh_visible_caption_index_cache`).
    """
    seen: list[str] = []
    for f in _toc_field_begins(list(doc.element.body)):
        seq = f.get("seq_id")
        if (
            seq
            and seq not in seen
            and f.get("begin_index") is not None
            and f.get("end_index") is not None
            and f["end_index"] >= f["begin_index"]
        ):
            seen.append(seq)
    return seen


def refresh_visible_caption_index_cache(doc, entries_by_seq: dict) -> int:
    r"""Rewrite the visible cache of every KEPT caption index from the emitted captions.

    The sibling of :func:`refresh_visible_outline_toc_cache` for ``TOC \c "<seq>"``
    caption indexes (a table-of-tables / table-of-figures). Marking the field dirty is
    enough for Word-on-open, but headless LibreOffice renders the stale cached result
    on export, so this replaces each kept caption index's cached entry paragraphs with
    one entry per caption the generator emitted for that index's ``seq_id``. The field
    code (the ``TOC \c`` instruction) is preserved and re-marked dirty, so Word still
    recomputes it on open. ``entries_by_seq`` maps a ``seq_id`` to its ordered visible
    entry strings (e.g. ``{"Tabella": ["Tabella 1. ...", "Tabella 2. ..."]}``).

    A caption index spans several top-level body children (begin/instr/separate plus
    entry-1 in the first paragraph, more entries in following paragraphs, the closing
    ``end`` fldChar in the last). Each kept index is rebuilt as
    ``[field-start, entry*, field-end]`` cloned from its first paragraph's style.
    Processed LAST-first so a splice never invalidates an earlier index's body-child
    span. Returns the number of indexes rebuilt.
    """
    if not entries_by_seq:
        return 0
    body = doc.element.body
    fields = [
        f
        for f in _toc_field_begins(list(body))
        if f.get("seq_id") in entries_by_seq
        and f.get("begin_index") is not None
        and f.get("end_index") is not None
        and f["end_index"] >= f["begin_index"]
    ]
    rebuilt = 0
    for f in sorted(fields, key=lambda f: f["begin_index"], reverse=True):
        # An empty entry list is a REBUILD-EMPTY, not a skip: the field survives
        # (dirty, Word recomputes on open) with zero cached entries, so the
        # template's demo entries never outlive a caption-less generation.
        entries = entries_by_seq.get(f["seq_id"]) or []
        children = list(body)
        begin_i, end_i = f["begin_index"], f["end_index"]
        # Both bounds are validated BEFORE any children[...] access: a corrupted
        # field inventory (or a body mutated between inventory and splice) must
        # skip the index, never raise IndexError mid-rebuild.
        if (
            begin_i < 0
            or begin_i >= len(children)
            or end_i < 0
            or end_i >= len(children)
        ):
            continue
        template_p = children[begin_i]
        instr = f" {f['instr'].strip()} "  # restore the field-code surrounding spaces
        new_paras = [_toc_field_start_paragraph(instr, template_p)]
        new_paras.extend(_toc_entry_paragraph(1, text, template_p) for text in entries)
        new_paras.append(_toc_field_end_paragraph(template_p))
        for np in new_paras:
            template_p.addprevious(np)
        for k in range(begin_i, min(end_i + 1, len(children))):
            old = children[k]
            if not _is_sectpr(old) and old.getparent() is body:
                body.remove(old)
        rebuilt += 1
    return rebuilt


def _outline_toc_instruction(el) -> Optional[str]:
    """Return the full instruction of an outline ``TOC`` field inside ``el``, or None.

    Word routinely splits a single field instruction across several consecutive
    ``w:instrText`` runs (e.g. ``TOC \\o `` + ``"1-3" \\h \\z \\u``), so each run
    must be concatenated PER ENCLOSING FIELD before classifying -- the same
    ``fldChar`` begin/end stack walk ``_toc_field_begins`` uses. Inspecting a
    single ``instrText`` node in isolation truncates the field code (losing the
    ``\\o`` range and switches) and can misread a split caption index
    (``TOC \\c "<seq>"`` whose ``\\c`` landed in a later run) as an outline TOC.

    The stack attributes each ``instrText`` to its innermost open field, so a
    nested field inside a rendered TOC entry (e.g. a ``PAGEREF``) accumulates into
    its own frame and never pollutes the outer TOC instruction. A field whose
    ``begin`` sits in ``el`` but whose ``end`` lives in a later body child (a
    multi-paragraph TOC result) leaves an open frame whose instruction is still
    complete, because ``instrText`` always precedes the ``separate`` fldChar.
    """
    stack: list[str] = []
    candidates: list[str] = []
    for kind, node in iter_complex_field_events(el):
        if kind == "begin":
            stack.append("")
        elif kind == "end" and stack:
            candidates.append(stack.pop())
        elif kind == "instr" and stack:
            stack[-1] += node.text or ""
    # Outermost-first for any field still open (end in a later body child).
    candidates.extend(stack)
    for instr in candidates:
        text = instr.strip()
        if text.startswith(TOC_INSTR_PREFIX) and _toc_seq_id(text) is None:
            return text
    return None


def _rewrite_sdt_outline_toc_cache(
    sdt, instr: str, headings: list[tuple[int, str]]
) -> bool:
    content = sdt.find(w("sdtContent"))
    if content is None:
        return False
    paras = [el for el in list(content) if _local_name(el.tag) == "p"]
    field_idx = next(
        (i for i, p in enumerate(paras) if _outline_toc_instruction(p) is not None),
        None,
    )
    if field_idx is None:
        return False

    old_field_para = paras[field_idx]
    entry_template = (
        paras[field_idx + 1] if field_idx + 1 < len(paras) else old_field_para
    )
    for p in paras[field_idx:]:
        if p.getparent() is content:
            content.remove(p)

    content.append(_toc_field_start_paragraph(instr, old_field_para))
    for level, text in headings:
        content.append(_toc_entry_paragraph(level, text, entry_template))
    content.append(_toc_field_end_paragraph(old_field_para))
    return True


def _rewrite_paragraph_outline_toc_cache(
    p, instr: str, headings: list[tuple[int, str]]
) -> None:
    _clear_paragraph_content(p)
    p.append(_field_run("begin", dirty=True))
    p.append(_instr_run(instr))
    p.append(_field_run("separate"))
    p.append(
        _text_run(" | ".join(_toc_entry_text(level, text) for level, text in headings))
    )
    p.append(_field_run("end"))


def _toc_field_start_paragraph(instr: str, template_p):
    p = _new_paragraph_like(template_p)
    p.append(_field_run("begin", dirty=True))
    p.append(_instr_run(instr))
    p.append(_field_run("separate"))
    return p


def _toc_field_end_paragraph(template_p):
    p = _new_paragraph_like(template_p)
    p.append(_field_run("end"))
    return p


def _toc_entry_paragraph(level: int, text: str, template_p):
    p = _new_paragraph_like(template_p)
    p.append(_text_run(_toc_entry_text(level, text)))
    return p


def _outline_toc_entry_paragraph(text: str, ppr_source, bookmark_name: str):
    r"""Build one Word-faithful rich TOC cache entry paragraph.

    Child order is normative: deep-copied ``pPr`` (omitted entirely when
    ``ppr_source`` is None), a ``w:hyperlink w:anchor=<bookmark>`` wrapping ONLY
    the entry text run (no ``rStyle`` - no style literals), a ``w:tab`` run, then
    a nested complex PAGEREF field authored dirty with NO cached result between
    ``separate`` and ``end`` (Word computes the page number on update). The
    PAGEREF instruction keeps its surrounding spaces via ``xml:space=preserve``,
    matching the template's own nested PAGEREFs. An internal anchor needs no
    relationship entry.
    """
    p = OxmlElement("w:p")
    if ppr_source is not None:
        p.append(copy.deepcopy(ppr_source))
    link = OxmlElement("w:hyperlink")
    link.set(w("anchor"), bookmark_name)
    link.set(w("history"), "1")
    link.append(_text_run(text))
    p.append(link)
    p.append(_tab_run())
    p.append(_field_run("begin", dirty=True))
    p.append(_instr_run(f" PAGEREF {bookmark_name} \\h ", preserve=True))
    p.append(_field_run("separate"))
    p.append(_field_run("end"))
    return p


def _toc_entry_text(level: int, text: str) -> str:
    return f"{'  ' * max(0, level - 1)}{text}"


def _new_paragraph_like(template_p):
    p = OxmlElement("w:p")
    pPr = template_p.find(w("pPr")) if template_p is not None else None
    if pPr is not None:
        p.append(copy.deepcopy(pPr))
    return p


def _clear_paragraph_content(p) -> None:
    for child in list(p):
        if _local_name(child.tag) != "pPr":
            p.remove(child)


def _field_run(kind: str, *, dirty: bool = False):
    r = OxmlElement("w:r")
    fld = OxmlElement("w:fldChar")
    fld.set(w("fldCharType"), kind)
    if dirty:
        fld.set(w("dirty"), "true")
    r.append(fld)
    return r


def _tab_run():
    r = OxmlElement("w:r")
    r.append(OxmlElement("w:tab"))
    return r


def _instr_run(instr: str, *, preserve: bool = False):
    # ``preserve=False`` keeps every existing call site (authored TOC, caption
    # index, plain fallbacks) byte-identical; only the rich PAGEREF construction
    # passes ``preserve=True`` to keep its surrounding spaces through Word.
    r = OxmlElement("w:r")
    it = OxmlElement("w:instrText")
    if preserve:
        it.set(qn("xml:space"), "preserve")
    it.text = instr
    r.append(it)
    return r


def _text_run(text: str):
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = text
    r.append(t)
    return r
