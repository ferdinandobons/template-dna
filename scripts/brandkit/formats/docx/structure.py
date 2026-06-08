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


def remove_index_field(doc, field_id: str) -> bool:
    """REMOVE one orphan caption-index block (its whole field span) in place.

    Returns True if anything was removed. Never touches the final ``sectPr``. To
    remove several indexes at once use :func:`remove_index_fields` - removing one
    shifts the body-child indices the position-based ids encode, so the multi
    variant resolves every target's elements against the ORIGINAL tree before
    deleting any.
    """
    return remove_index_fields(doc, [field_id])


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
def clear_body(doc) -> None:
    """Remove every top-level body child except the final ``sectPr``.

    Back-compat entry point (M1 behaviour). Prefer :func:`clear_body_region`, which
    preserves the cover and TOC regions. This is now routed through the structural
    clear with both preservation flags off, so it behaves identically to the old
    "wipe everything" clear.
    """
    clear_body_region(doc, preserve_cover=False, preserve_toc=False)


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


def refresh_visible_outline_toc_cache(doc, headings: list[tuple[int, str]]) -> int:
    """Rewrite the visible cache of outline TOC fields from generated headings.

    Word/LibreOffice store a field *result* alongside the TOC field code. Marking
    the field dirty is enough for Word-on-open, but headless LibreOffice often
    renders the stale cached result when exporting to PDF. This function keeps the
    outline TOC field updateable while replacing the visible cache with the
    generated document's current headings, so visual audit does not show template
    sample entries.
    """
    clean_headings = [
        (max(1, int(level or 1)), str(text).strip())
        for level, text in headings
        if str(text).strip()
    ]
    if not clean_headings:
        return 0

    rewritten = 0
    for child in list(doc.element.body):
        instr = _outline_toc_instruction(child)
        if instr is None:
            continue
        if _local_name(child.tag) == "sdt":
            if _rewrite_sdt_outline_toc_cache(child, instr, clean_headings):
                rewritten += 1
        elif _local_name(child.tag) == "p":
            _rewrite_paragraph_outline_toc_cache(child, instr, clean_headings)
            rewritten += 1
    return rewritten


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


def _instr_run(instr: str):
    r = OxmlElement("w:r")
    it = OxmlElement("w:instrText")
    it.text = instr
    r.append(it)
    return r


def _text_run(text: str):
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = text
    r.append(t)
    return r
