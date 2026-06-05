# SPDX-License-Identifier: MIT
"""DOCX cover-anchor discovery and composition.

Real company covers seldom expose their title as a plain paragraph: the title
usually lives in a block-level ``w:sdt`` (a content control, with an ``alias`` /
``dataBinding`` / ``docPartGallery`` / placeholder prompt). python-docx's
``doc.paragraphs`` cannot see block-level SDTs, so discovery and composition both
work on the lxml tree here. The cover title is filled IN PLACE - only the inner
placeholder run's text is overwritten so the run-level formatting (``w:rPr``) the
brand defined survives - and a brand-new title paragraph is appended only when the
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
    _p_style_val,
    classify_body_children,
    w,
)
from brandkit.ir.model import Cover
from brandkit.qa.model import Finding
from brandkit.profile import schema


PLACEHOLDER_TITLE = "{{title}}"
# WEAK PRIOR ONLY (plan §5). These multilingual title prompts are a LAST-RESORT
# tiebreaker for naming the title slot when no structural SDT metadata is present;
# they are never the primary signal and never a matching rule that gates output.
# Cover-slot DISCOVERY is structural (block-level SDTs + the cover-region
# placeholder paragraphs); the model names each slot via ``comprehension``.
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
    misclassify the TOC as the title - which would dump the user's title into the
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


def _sdt_alias(sdt) -> Optional[str]:
    props = _sdt_props(sdt)
    if props is None:
        return None
    alias = props.find(w("alias"))
    return alias.get(w("val")) if alias is not None else None


def _sdt_databinding(sdt) -> Optional[str]:
    props = _sdt_props(sdt)
    if props is None:
        return None
    binding = props.find(w("dataBinding"))
    return binding.get(w("xpath")) if binding is not None else None


def _sdt_gallery(sdt) -> Optional[str]:
    props = _sdt_props(sdt)
    if props is None:
        return None
    for gallery in props.iter(w("docPartGallery")):
        return gallery.get(w("val"))
    return None


def _sdt_showing_placeholder(sdt) -> bool:
    props = _sdt_props(sdt)
    return props is not None and props.find(w("showingPlcHdr")) is not None


def _iter_block_sdts(doc):
    """Yield top-level block-level ``w:sdt`` elements in document order."""
    for child in doc.element.body:
        if _local_name(child.tag) == "sdt":
            yield child


def _cover_child_indices(doc) -> list[int]:
    """Return the top-level body-child indices that belong to the cover region."""
    return [c["index"] for c in classify_body_children(doc) if c.get("region") == "cover"]


def _paragraph_is_placeholder_slot(p) -> bool:
    """Decide STRUCTURALLY whether a cover-region paragraph is a fillable slot.

    Evidence (language-invariant, no brand words):
      - it carries a literal ``{{...}}`` placeholder token, OR
      - it is a single short text run on the default (no explicit ``pStyle``, i.e.
        the run carries the cover's own direct formatting) - the shape a hand-typed
        cover prompt takes (a title line, an identifier line, a date line).

    A blank paragraph (empty text) is NOT a slot. The weak title-prompt tokens are
    consulted only to break ties when several short paragraphs qualify (see
    :func:`discover_cover`), never as the primary gate here.
    """
    text = _p_text_local(p).strip()
    if not text:
        return False
    if "{{" in text and "}}" in text:
        return True
    # A short, single-line cover prompt with no list/heading structure.
    if "\n" in text or len(text) > 120:
        return False
    return True


def _p_text_local(p) -> str:
    return "".join(t.text for t in p.iter(w("t")) if t.text)


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
    """Enumerate EVERY fillable cover slot as one anchor (plan §4 fact inventory).

    A real company cover is multi-slot: a title, a subtitle / object, a document
    id, a date, each its own fillable element. Discovery surfaces one anchor per
    slot - never just the title - so the comprehension can bind each slot and
    generation can FILL/CLEAR each in place. The two structural slot containers:

      - **block-level ``w:sdt``** (a content control) that is NOT a TOC/index field
        - keyed ``sdt.<body-child-index>``. The structural metadata (alias /
          dataBinding / docPartGallery / placeholder text / ``showingPlcHdr``) is
          captured as evidence; the model assigns the semantic role.
      - **cover-region placeholder paragraph** - a non-blank, short, single-line
        paragraph in the cover region (the body content before the first TOC /
        Heading-1) - keyed ``para.<body-child-index>``.

    The body-child index is a STABLE id: generation rewrites only the freeform body
    region, so the cover front matter keeps its positions, and the generator
    recomputes the same ids from the live tree. No anchor carries a brand word; the
    captured ``placeholder`` text is data (the template's own demo value), used for
    in-place fill and the residual-text check, never as a matching rule.

    Returns ``(anchors, anchor_block)`` where each anchor is::

        {"id": "sdt.8", "container": "sdt", "child_index": 8,
         "placeholder": "Descrizione", "alias": "Oggetto", "data_binding": None,
         "gallery": None, "showing_placeholder": False, "branches": None}

    and ``anchor_block`` is the legacy ``{"cover": {"kind","slots_found"}}`` summary.
    """
    anchors: list[dict] = []
    children = list(doc.element.body)
    cover_indices = set(_cover_child_indices(doc))

    for i, child in enumerate(children):
        ln = _local_name(child.tag)
        if ln == "sdt":
            # A TOC / index content control is never a cover slot.
            if _element_holds_strong_toc(child):
                continue
            anchors.append(
                {
                    "id": f"sdt.{i}",
                    "container": "sdt",
                    "child_index": i,
                    "placeholder": _sdt_text(child)[:200],
                    "alias": _sdt_alias(child),
                    "data_binding": _sdt_databinding(child),
                    "gallery": _sdt_gallery(child),
                    "showing_placeholder": _sdt_showing_placeholder(child),
                    "branches": None,
                }
            )
        elif ln == "p" and i in cover_indices and _paragraph_is_placeholder_slot(child):
            anchors.append(
                {
                    "id": f"para.{i}",
                    "container": "paragraph",
                    "child_index": i,
                    "placeholder": _p_text_local(child)[:200],
                    "style_id": _p_style_val(child),
                    "branches": None,
                }
            )

    anchor_block = {
        "cover": {
            "kind": schema.AnchorKind.SDT_ANCHORED.value if anchors else schema.AnchorKind.NONE.value,
            "slots_found": len(anchors),
        }
    }
    return anchors, anchor_block


# Below this confidence the model's DESTRUCTIVE verdicts (a ``clear`` on a slot
# determinism does not also corroborate) are downgraded to KEEP+WARNING. Additive
# FILL is never gated on confidence (a wrong fill is recoverable; a wrong delete
# is not - the destructive-action floor, plan §6).
_DESTRUCTIVE_CONFIDENCE_FLOOR: float = 0.5


def compose_cover(
    doc,
    cover: Cover | None,
    profile: dict,
    *,
    findings: Optional[list[Finding]] = None,
) -> set[str]:
    """Reconcile the PRESERVED cover slots with the new content (plan §6).

    When the profile carries a present comprehension, every cover slot is
    reconciled in place by its frozen ``fill_rule`` (multi-slot, never just the
    title): ``in_place`` FILLs the bound ``idoc.cover`` content preserving run
    formatting; ``clear`` empties / re-arms the slot, but ONLY when determinism
    corroborates the slot is a placeholder (``showingPlcHdr`` set, run==style
    default) and ``comprehension.confidence`` clears the floor - otherwise it is
    downgraded to KEEP + WARNING (the destructive-action floor); ``leave`` keeps
    the slot untouched. The append-a-new-title fallback fires ONLY when no
    title-bearing slot exists at all, so a duplicate title is never appended.

    When comprehension is ABSENT this falls back to today's deterministic single-
    title behavior (SDT title, then ``{{title}}``/"Insert title" placeholder, then
    append-before-TOC).

    Returns the set of cover anchor refs the reconciliation actually CLEARED
    (emptied/re-armed), so the caller can feed ``no_net_structure_loss``.
    """
    sink = findings if findings is not None else []
    if cover is None:
        return set()

    comp = _present_comprehension(profile)
    if comp is not None and comp.get("cover_slots"):
        return _compose_cover_comprehended(doc, cover, profile, comp, sink)
    return _compose_cover_deterministic(doc, cover, profile, sink)


def _present_comprehension(profile: dict) -> Optional[dict]:
    """Return the comprehension block only when it is present AND sha-current."""
    from brandkit.profile import store

    if not store.comprehension_is_present(profile):
        return None
    return profile.get("comprehension")


def _cover_content_for(cover: Cover, binds_to: Optional[str]) -> Optional[str]:
    """Resolve the new content string for a slot's ``binds_to`` key (advisory).

    ``binds_to`` is the model's advisory hint for WHICH content slot fills this
    anchor. ``title`` / ``subtitle`` map to the structured ``ir.Cover`` fields;
    anything else is looked up in ``cover.fields`` (the free-form ``{slot: value}``
    map). Returns None when the content has nothing for this slot (⇒ CLEAR if the
    model so ruled, else LEAVE).
    """
    if not binds_to:
        return None
    if binds_to == "title":
        return textutil.runs_to_text(cover.title or []) or str(cover.fields.get("title", "")) or None
    if binds_to == "subtitle":
        return textutil.runs_to_text(cover.subtitle or []) or str(cover.fields.get("subtitle", "")) or None
    val = cover.fields.get(binds_to)
    return str(val) if val not in (None, "") else None


def _compose_cover_comprehended(doc, cover: Cover, profile: dict, comp: dict, sink: list) -> set[str]:
    """Multi-slot, comprehension-steered cover reconciliation."""
    confidence = float(comp.get("confidence") or 0.0)
    slots = comp.get("cover_slots") or {}
    cleared: set[str] = set()
    title_slot_filled = False
    has_title_slot = False

    # Resolve each slot's live element by its stable id, then act on fill_rule.
    for anchor_ref in sorted(slots):
        slot = slots[anchor_ref]
        if not isinstance(slot, dict):
            continue
        el = _resolve_anchor_element(doc, anchor_ref)
        if el is None:
            continue  # the slot is gone (already cleared / never existed); skip.
        fill_rule = slot.get("fill_rule")
        binds_to = slot.get("binds_to")
        if binds_to == "title":
            has_title_slot = True
        content = _cover_content_for(cover, binds_to)

        if fill_rule == schema.FillRule.IN_PLACE.value and content:
            _fill_anchor_in_place(doc, el, anchor_ref, content, profile, binds_to)
            if binds_to == "title":
                title_slot_filled = True
            continue

        if fill_rule == schema.FillRule.CLEAR.value or (
            fill_rule == schema.FillRule.IN_PLACE.value and not content
        ):
            # A CLEAR (or an in_place slot the content does not fill) is honored
            # only when determinism corroborates the slot is a placeholder AND the
            # confidence clears the floor; otherwise KEEP + WARNING.
            if _clear_is_corroborated(el, confidence):
                _clear_anchor(doc, el)
                cleared.add(anchor_ref)
            else:
                sink.append(
                    Finding(
                        "cover_clear_downgraded",
                        schema.Severity.WARNING.value,
                        f"cover slot {anchor_ref!r} clear not corroborated "
                        f"(confidence {confidence:.2f}); kept",
                    )
                )
            continue
        # fill_rule == leave (or unknown): leave the slot untouched.

    # Append a brand-new title ONLY when there is no title-bearing slot at all.
    title = textutil.runs_to_text(cover.title or []) or str(cover.fields.get("title", ""))
    if title and not has_title_slot and not title_slot_filled:
        para = doc.add_paragraph(title)
        _apply_role_style(doc, para, profile, "cover.title")
        _move_before_first_toc_or_body(doc, para)
        sink.append(
            Finding(
                "cover_degraded",
                schema.Severity.WARNING.value,
                "no title-bearing cover slot in shell; title paragraph appended "
                "before the first toc/body child",
            )
        )
    return cleared


def _compose_cover_deterministic(doc, cover: Cover, profile: dict, sink: list) -> set[str]:
    """Today's deterministic single-title cover fill (comprehension absent)."""
    title = textutil.runs_to_text(cover.title or []) or str(cover.fields.get("title", ""))
    if not title:
        return set()

    # 1) Block-level SDT cover title.
    for sdt in _iter_block_sdts(doc):
        if _sdt_is_title(sdt) and _fill_sdt_title(sdt, title):
            _apply_role_style_sdt(doc, sdt, profile, "cover.title")
            return set()

    # 2) Placeholder paragraph - overwrite only the matching run's text in place.
    for para in doc.paragraphs[:8]:
        if PLACEHOLDER_TITLE in para.text or "Insert title" in para.text:
            _fill_paragraph_in_place(para, title)
            _apply_role_style(doc, para, profile, "cover.title")
            return set()

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
    return set()


def _resolve_anchor_element(doc, anchor_ref: str):
    """Resolve a stable anchor id (``sdt.<i>`` / ``para.<i>``) to its live element.

    The body-child index in the id is recomputed against the live tree; cover
    front matter keeps its positions across generation (only the freeform body is
    rewritten), so the same id resolves to the same element at generate time.
    Returns the element, or None when the id no longer maps (defensive).
    """
    kind, _, idx_s = anchor_ref.partition(".")
    try:
        idx = int(idx_s)
    except ValueError:
        return None
    children = list(doc.element.body)
    if idx < 0 or idx >= len(children):
        return None
    el = children[idx]
    ln = _local_name(el.tag)
    if kind == "sdt" and ln == "sdt":
        return el
    if kind == "para" and ln == "p":
        return el
    return None


def _fill_anchor_in_place(
    doc, el, anchor_ref: str, content: str, profile: dict, binds_to: Optional[str] = None
) -> None:
    """FILL a cover anchor element in place, preserving run formatting.

    After the in-place fill, the slot's bound role style is re-asserted so a filled
    cover slot is GUARANTEED brand-styled, not merely whatever incidental style the
    prompt carried (D4). Both the SDT branch and the plain-paragraph branch now
    re-apply the role style; the role is derived from the slot's ``binds_to``
    (``title``->``cover.title``, ``subtitle``->``cover.subtitle``), a verbatim
    resolver target only - no literal.
    """
    ln = _local_name(el.tag)
    role_id = _cover_role_for(binds_to)
    if ln == "sdt":
        _fill_sdt_title(el, content)
        if role_id:
            _apply_role_style_sdt(doc, el, profile, role_id)
    elif ln == "p":
        _fill_p_element_in_place(el, content)
        if role_id:
            _apply_role_style_p_element(doc, el, profile, role_id)


def _cover_role_for(binds_to: Optional[str]) -> Optional[str]:
    """Map a slot's ``binds_to`` hint to its cover role id (verbatim, no literal).

    ``title`` -> ``cover.title``; ``subtitle`` -> ``cover.subtitle``. Any other (or
    absent) binding defaults to ``cover.title`` so a filled, un-annotated slot
    still re-asserts the cover title style rather than keeping a prompt's incidental
    style. Returns the role id, or None only when there is nothing to bind.
    """
    if binds_to == "subtitle":
        return "cover.subtitle"
    return "cover.title"


def _clear_is_corroborated(el, confidence: float) -> bool:
    """Destructive-action floor for a cover CLEAR (plan §6).

    A CLEAR is corroborated only when determinism agrees the slot is a placeholder
    - an SDT still ``showingPlcHdr``, or a paragraph whose runs carry no explicit
    paragraph style (the cover's own default-styled prompt) - AND the model's
    confidence clears the floor. Both conditions are required; otherwise the slot
    is kept.
    """
    if confidence < _DESTRUCTIVE_CONFIDENCE_FLOOR:
        return False
    ln = _local_name(el.tag)
    if ln == "sdt":
        return _sdt_showing_placeholder(el)
    if ln == "p":
        # A cover prompt paragraph carries no heading/list style of its own.
        return _p_style_val(el) is None
    return False


def _clear_anchor(doc, el) -> None:
    """CLEAR a cover anchor in place: empty its text, re-arming an SDT placeholder."""
    ln = _local_name(el.tag)
    if ln == "sdt":
        content = el.find(w("sdtContent"))
        scope = content if content is not None else el
        for t in scope.iter(w("t")):
            t.text = ""
    elif ln == "p":
        for t in el.iter(w("t")):
            t.text = ""


def _fill_p_element_in_place(p_el, content: str) -> None:
    """Set ``content`` on a lxml ``w:p`` placeholder, preserving the first run rPr."""
    texts = list(p_el.iter(w("t")))
    if not texts:
        # No run/text node: create a minimal run so the content is not lost.
        from lxml import etree

        r = etree.SubElement(p_el, w("r"))
        t = etree.SubElement(r, w("t"))
        t.text = content
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        return
    texts[0].text = content
    texts[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    for extra in texts[1:]:
        extra.text = ""


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
    """Apply a role's paragraph style to the first paragraph inside the SDT."""
    content = sdt.find(w("sdtContent"))
    scope = content if content is not None else sdt
    p = next(iter(scope.iter(w("p"))), None)
    if p is not None:
        _apply_role_style_p_element(doc, p, profile, role_id)


def _apply_role_style_p_element(doc, p_el, profile: dict, role_id: str) -> None:
    """Stamp a role's ``w:pStyle`` onto a bare lxml ``w:p`` element.

    The single place cover style re-assertion writes a paragraph style id onto an
    lxml element (used by both the SDT branch and the plain-paragraph branch, D4).
    No-op when the role resolves to no shell style or the style has no id, so a
    missing/absent role never crashes the in-place fill.
    """
    if _local_name(p_el.tag) != "p":
        return
    entry = (profile.get("roles") or {}).get(role_id) or {}
    resolver = entry.get("resolver") or {}
    style = lookup_style(doc, resolver)
    if style is None:
        return
    style_id = getattr(style, "style_id", None)
    if not style_id:
        return
    pPr = p_el.find(w("pPr"))
    if pPr is None:
        from lxml import etree

        pPr = etree.SubElement(p_el, w("pPr"))
        p_el.insert(0, pPr)
    pStyle = pPr.find(w("pStyle"))
    if pStyle is None:
        from lxml import etree

        pStyle = etree.SubElement(pPr, w("pStyle"))
        pPr.insert(0, pStyle)
    pStyle.set(w("val"), style_id)
