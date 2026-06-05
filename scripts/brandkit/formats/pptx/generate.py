# SPDX-License-Identifier: MIT
"""Generate an on-brand ``.pptx`` from the IR block stream and a Brand Profile.

Design (off-brand-by-construction, §C3/M6/M7, plan §6 reconcile-not-rebuild):

- Layouts are resolved through the SHARED resolver spine
  (:class:`brandkit.profile.resolver.ProfileResolver`), the same one the docx and
  xlsx generators route through, so the brand guarantee - "apply only a layout the
  profile proved exists" - is enforced in ONE place for all three formats. The
  resolver returns a ``placeholder`` op (``{layout, ph_idx, ph_type}``) read
  verbatim from the profile's real ``roles`` (``cover.title`` / ``heading.1`` /
  ``paragraph``); a stub role (no layout) degrades to a real shell layout, never a
  fabricated name.

- Two paths, chosen by whether a CURRENT comprehension is present (sha-bound):

  * **Comprehension ABSENT (the deterministic path CI exercises):** today's
    behavior, unchanged and byte-identical - clear every existing slide, then build
    a cover slide (when the IR carries a cover title) plus one content slide per
    heading from the IR block stream.

  * **Comprehension PRESENT (reconcile-not-rebuild):** the deck is RECONCILED, not
    blind-rebuilt. Structural slides the template ships (the ones the model did not
    tag ``verdict=demo``) are KEPT; only demo slides (a slide whose body text equals
    a layout placeholder prompt) are cleared. The multi-placeholder cover is filled
    IN PLACE on the existing/created cover slide (each ``cover_slots`` entry by its
    ``fill_rule``), never recreated. The agenda / section-list index is regenerated
    from the NEW headings. New body content is appended after the kept slides.

- Slides are built from the IR block stream, not a flattened string: each heading
  opens a new section/slide (its own runs become the title); the blocks that follow
  become that slide's body, preserving lists / tables / quotes / captions /
  callouts as distinct lines. Heading text is never duplicated into the body.
"""
from __future__ import annotations

from pathlib import Path
from typing import Optional

from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER

from brandkit.common import text as textutil
from brandkit.formats.pptx import structure
from brandkit.ir import model as ir
from brandkit.profile import schema, store
from brandkit.profile.resolver import ProfileResolver
from brandkit.qa.checks_deterministic import check_no_net_structure_loss
from brandkit.qa.model import Finding


def generate(
    profile: dict,
    shell_path: str | Path,
    idoc: ir.IntermediateDocument,
    output: str | Path,
    *,
    findings: Optional[list[Finding]] = None,
) -> Path:
    """Generate a PPTX from ``idoc`` onto the brand ``shell_path``.

    ``findings`` (optional out-param) is appended with reconciliation /
    destructive-floor findings so the QA gate can surface them. When comprehension
    is absent this is the proven deterministic rebuild; when present it reconciles
    (keeps structural slides, fills the multi-placeholder cover in place,
    regenerates the agenda from the new headings).
    """
    sink: list[Finding] = findings if findings is not None else []
    prs = Presentation(shell_path)
    resolver = ProfileResolver(profile)

    cover_layout = _layout_for_role(prs, resolver, "cover.title")
    content_layout = _layout_for_role(prs, resolver, "heading.1") or _layout_for_role(prs, resolver, "paragraph")

    if store.comprehension_is_present(profile):
        _generate_reconciled(prs, profile, idoc, cover_layout, content_layout, sink)
    else:
        _generate_deterministic(prs, profile, idoc, cover_layout, content_layout)

    out = Path(output)
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(out)
    return out


# ---------------------------------------------------------------------------
# Deterministic path (comprehension ABSENT) - today's behavior, byte-identical
# ---------------------------------------------------------------------------
def _generate_deterministic(prs, profile: dict, idoc, cover_layout, content_layout) -> None:
    """Blind rebuild: clear all slides, build cover + one content slide per heading.

    Unchanged from the pre-comprehension behavior so the model-free CI path stays
    the ground truth and generate-twice is byte-identical.
    """
    _clear_existing_slides(prs)

    # Cover slide: only emitted when the IR actually carries a cover title and the
    # shell offers a layout with a title placeholder.
    if idoc.cover and idoc.cover.title:
        cover_slide = prs.slides.add_slide(cover_layout or content_layout or prs.slide_layouts[0])
        if cover_slide.shapes.title is not None:
            cover_slide.shapes.title.text = textutil.runs_to_text(idoc.cover.title)
        if idoc.cover.subtitle:
            sub = _subtitle_placeholder(cover_slide)
            if sub is not None:
                sub.text = textutil.runs_to_text(idoc.cover.subtitle)

    _append_content_slides(prs, profile, idoc, content_layout)


# ---------------------------------------------------------------------------
# Reconcile path (comprehension PRESENT) - keep structural, fill cover, regen agenda
# ---------------------------------------------------------------------------
def _generate_reconciled(prs, profile: dict, idoc, cover_layout, content_layout, sink: list) -> None:
    """Reconcile the preserved deck with the new content (plan §6).

    KEEP structural slides (everything the model did NOT tag ``verdict=demo``),
    CLEAR demo slides only, FILL the multi-placeholder cover in place, regenerate
    the agenda/section-list from the new headings, then append the new body slides.
    """
    comp = profile.get("comprehension") or {}

    # 1) Cover fill IN PLACE on the existing cover slide, or a freshly-added cover
    # slide when the deck ships none. Multi-placeholder, by fill_rule. Returns the
    # set of cover anchor refs the reconciliation actually CLEARED.
    cleared_anchors = _fill_cover_in_place(prs, profile, comp, idoc, cover_layout, content_layout, sink)

    # 2) CLEAR demo slides only (the destructive floor: a slide is removed only when
    # the model tagged its region demo AND determinism corroborates it is demo - its
    # body text equals a layout prompt). Structural slides are KEPT.
    removed_region_refs = _clear_demo_slides(prs, comp, sink)

    # 3) Append the new body content after the kept slides.
    _append_content_slides(prs, profile, idoc, content_layout)

    # 4) Regenerate the agenda / section-list index from the NEW headings (the PPTX
    # peer of refreshing the docx outline TOC). No-op when the deck has no section
    # list / the convention is preserve.
    _regenerate_agenda(prs, comp, idoc, content_layout, sink)

    # 5) Destructive-action floor (plan §6): every cover anchor / region the
    # reconciliation removed must carry a corroborated destructive verdict, else
    # ERROR. Model-free; reads the frozen verdicts.
    sink.extend(check_no_net_structure_loss(cleared_anchors | removed_region_refs, profile))


def _fill_cover_in_place(prs, profile: dict, comp: dict, idoc, cover_layout, content_layout, sink: list) -> set[str]:
    """Fill the multi-placeholder cover IN PLACE by each slot's ``fill_rule``.

    Resolves each ``cover_slots`` anchor (``ph.<layout-idx>.<ph-idx>``) to a live
    placeholder on the cover slide: an existing cover slide is reused, else one is
    added on the cover layout (only when the IR carries cover content to place). For
    each slot: ``in_place`` FILLs the bound content; ``clear`` empties the
    placeholder, but ONLY when determinism corroborates it still shows its prompt
    AND confidence clears the floor (else KEEP + WARNING); ``leave`` is untouched.

    Returns the set of anchor refs actually CLEARED (for ``no_net_structure_loss``).
    """
    cover = idoc.cover
    slots = comp.get("cover_slots") or {}
    if not slots or cover is None:
        return set()

    cover_slide = _existing_cover_slide(prs, profile)
    if cover_slide is None:
        # No cover slide in the deck: add one on the cover layout so the slots have a
        # surface to fill. Only worth doing when there is cover content to place.
        if not (cover.title or cover.subtitle or cover.fields):
            return set()
        cover_slide = prs.slides.add_slide(cover_layout or content_layout or prs.slide_layouts[0])

    confidence = float(comp.get("confidence") or 0.0)
    cleared: set[str] = set()
    for anchor_ref in sorted(slots):
        slot = slots[anchor_ref]
        if not isinstance(slot, dict):
            continue
        ph = _placeholder_for_anchor(cover_slide, anchor_ref)
        if ph is None:
            continue  # the slot's placeholder is not on this slide; skip.
        fill_rule = slot.get("fill_rule")
        binds_to = slot.get("binds_to")
        content = _cover_content_for(cover, binds_to)

        if fill_rule == schema.FillRule.IN_PLACE.value and content:
            _set_placeholder_text(ph, content)
            continue
        if fill_rule == schema.FillRule.CLEAR.value or (
            fill_rule == schema.FillRule.IN_PLACE.value and not content
        ):
            if _clear_is_corroborated(ph, slot, confidence):
                _set_placeholder_text(ph, "")
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
        # fill_rule == leave (or unknown): leave untouched.
    return cleared


def _clear_demo_slides(prs, comp: dict, sink: list) -> set[str]:
    """Remove demo slides the model tagged ``verdict=demo`` AND determinism agrees.

    The destructive floor: a slide is removed only when (i) the model tagged its
    ``region.slide.<i>`` ref ``verdict=demo`` and (ii) determinism corroborates it
    is demo (its body text equals a layout placeholder prompt - recomputed live).
    Otherwise the slide is KEPT (a wrong delete is not recoverable). Returns the set
    of region refs actually removed.
    """
    demo_refs = {
        r.get("region_ref")
        for r in (comp.get("demo_classification") or {}).get("regions") or []
        if isinstance(r, dict) and r.get("verdict") == schema.Verdict.DEMO.value
    }
    if not demo_refs:
        return set()
    det_demo = set(structure.demo_slide_indices(prs))
    removed: set[str] = set()
    # Resolve refs to slide elements FIRST (removing one shifts indices), then drop.
    sld_id_lst = prs.slides._sldIdLst
    sld_ids = list(sld_id_lst)
    to_remove: list = []
    for ref in demo_refs:
        idx = _slide_index_from_ref(ref)
        if idx is None or idx < 0 or idx >= len(sld_ids):
            continue
        if idx not in det_demo:
            sink.append(
                Finding(
                    "demo_clear_downgraded",
                    schema.Severity.WARNING.value,
                    f"slide region {ref!r} clear not corroborated "
                    f"(slide carries authored text, not a layout prompt); kept",
                )
            )
            continue
        to_remove.append((ref, sld_ids[idx]))
    for ref, sld_id in to_remove:
        r_id = sld_id.rId
        prs.part.drop_rel(r_id)
        sld_id_lst.remove(sld_id)
        removed.add(ref)
    return removed


def _regenerate_agenda(prs, comp: dict, idoc, content_layout, sink: list) -> None:
    """Regenerate the agenda / section-list index from the NEW headings.

    For each ``conventions.indexes`` entry whose ``reconcile == regenerate`` and
    which is the section-list field (``field.sections``), refresh the agenda's body
    to list the NEW headings (the PPTX peer of refreshing the docx outline TOC from
    the new body headings). ``preserve`` keeps the existing agenda untouched;
    ``clear`` is handled by the demo-slide path (an agenda slide tagged demo). No-op
    when the deck has no section-list convention.

    When the deck SHIPS an agenda / section-list slide, it is refreshed IN PLACE -
    its own title (the author's word, e.g. "Sommario" / "Übersicht", in the
    template's language) and run formatting are preserved and only the list body is
    rewritten - so no English literal is injected and no stale duplicate is left
    behind. This is the faithful PPTX peer of ``refresh_toc`` (which refreshes the
    docx TOC field in place and never re-emits the heading). Only when the deck ships
    NO agenda slide is a fresh one appended, with the layout's default (empty) title
    rather than a fabricated literal.
    """
    indexes = (comp.get("conventions") or {}).get("indexes") or []
    headings = [
        textutil.runs_to_text(b.runs)
        for b in idoc.blocks
        if isinstance(b, ir.Heading) and textutil.runs_to_text(b.runs)
    ]
    for idx in indexes:
        if not isinstance(idx, dict):
            continue
        if idx.get("index_ref") != "field.sections":
            continue
        if idx.get("reconcile") != schema.Reconcile.REGENERATE.value:
            continue  # preserve / clear: not regenerated here
        if not headings:
            continue
        new_body = "\n".join(headings)
        existing = _existing_agenda_slide(prs)
        if existing is not None:
            # Refresh IN PLACE: keep the title (author's word, language) + run
            # formatting; rewrite only the section-list body to the new headings.
            body = _first_body_placeholder(existing)
            if body is not None:
                _set_placeholder_text(body, new_body)
        else:
            layout = content_layout or prs.slide_layouts[0]
            slide = prs.slides.add_slide(layout)
            # No agenda slide shipped: leave the layout's default (empty) title rather
            # than inject a literal; fill the body with the new section list.
            body = _first_body_placeholder(slide)
            if body is not None:
                body.text = new_body
        sink.append(
            Finding(
                "agenda_regenerated",
                schema.Severity.INFO.value,
                f"agenda/section-list regenerated from {len(headings)} new heading(s)",
            )
        )


def _existing_agenda_slide(prs):
    """Return the deck's existing agenda / section-list slide, or None.

    The agenda slide is identified STRUCTURALLY, never by a fixed word: it is the
    structural slide whose body text lists the deck's OWN section names (the visible
    list page that mirrors the ``p14:sectionLst``). A slide qualifies when each of
    its body lines matches a section name from ``detect_sections`` (and it lists at
    least two), so the page is the section index in whatever language the template is
    authored in. Its title (e.g. "Sommario" / "Übersicht" / "Agenda") and formatting
    are preserved when the agenda is refreshed - the PPTX peer of ``refresh_toc``
    preserving the docx TOC heading.

    Returns None when the deck ships no such slide (caller appends a fresh one with
    no fabricated title rather than carrying a literal).
    """
    section_names = {
        (sec.get("name") or "").strip()
        for sec in structure.detect_sections(prs)
        if (sec.get("name") or "").strip()
    }
    if len(section_names) < 2:
        return None
    for slide in prs.slides:
        if slide.shapes.title is None:
            continue
        body = _first_body_placeholder(slide)
        if body is None or not getattr(body, "has_text_frame", False) or not body.text:
            continue
        lines = [ln.strip() for ln in body.text.splitlines() if ln.strip()]
        if len(lines) >= 2 and all(ln in section_names for ln in lines):
            return slide
    return None


# ---------------------------------------------------------------------------
# Shared content-slide builder (both paths append body content the same way)
# ---------------------------------------------------------------------------
def _append_content_slides(prs, profile: dict, idoc, content_layout) -> None:
    """Append one content slide per IR heading-section (capacity-split)."""
    capacity = _body_capacity(profile)
    layout = content_layout or prs.slide_layouts[0]
    for section in _sections(idoc.blocks):
        body_lines = _body_lines(section["body"])
        for page, chunk in enumerate(_split_lines(body_lines, capacity)):
            slide = prs.slides.add_slide(layout)
            title = section["title"]
            if page:
                title = f"{title} ({page + 1})"
            if slide.shapes.title is not None:
                slide.shapes.title.text = title
            body = _first_body_placeholder(slide)
            if body is not None and chunk:
                body.text = "\n".join(chunk)


# ---------------------------------------------------------------------------
# Layout resolution - through the SHARED resolver spine (M-i-8)
# ---------------------------------------------------------------------------
def _layout_for_role(prs, resolver: ProfileResolver, role_id: str):
    """Return the real shell layout the profile assigns to ``role_id``.

    Routes through the shared :class:`ProfileResolver`, which returns the
    ``placeholder`` op verbatim from ``profile['roles']`` (and refuses any resolver
    type illegal for kind ``pptx``). The op's ``layout`` is a name the extractor
    proved exists; we look it up by name. Returns ``None`` when the role is a stub
    (no layout) or the named layout is absent from the shell (callers fall back to a
    real layout, never a fiction).
    """
    op = resolver.resolve_role(role_id, fallback=None)
    name = op.resolver.get("layout") if op.resolver else None
    if not name:
        return None
    return _layout_by_name(prs, name)


def _layout_by_name(prs, name: str):
    for layout in prs.slide_layouts:
        if layout.name == name:
            return layout
    return None


# ---------------------------------------------------------------------------
# Cover-anchor / placeholder helpers (reconcile path)
# ---------------------------------------------------------------------------
def _existing_cover_slide(prs, profile: dict):
    """Return the first slide built on the cover layout, or None.

    The cover layout name is the one the cover-anchor inventory was built on
    (``surface.pptx.cover_anchors[*].layout``), so the cover slide is the slide
    whose layout matches - the anchor ids (``ph.<layout-idx>.<ph-idx>``) resolve to
    placeholders on it.
    """
    cover_layout_name = _cover_layout_name(profile)
    if cover_layout_name is None:
        return None
    for slide in prs.slides:
        if slide.slide_layout.name == cover_layout_name:
            return slide
    return None


def _cover_layout_name(profile: dict) -> Optional[str]:
    anchors = ((profile.get("surface") or {}).get("pptx") or {}).get("cover_anchors") or []
    for a in anchors:
        if isinstance(a, dict) and a.get("layout"):
            return a["layout"]
    return None


def _placeholder_for_anchor(slide, anchor_ref: str):
    """Resolve an anchor id ``ph.<layout-idx>.<ph-idx>`` to a live placeholder.

    The ``ph_idx`` is the stable placeholder index; we find the placeholder on
    ``slide`` whose ``placeholder_format.idx`` matches. Returns None when the slide
    has no such placeholder (defensive).
    """
    ph_idx = _anchor_ph_idx(anchor_ref)
    if ph_idx is None:
        return None
    for shape in slide.placeholders:
        if shape.placeholder_format.idx == ph_idx:
            return shape
    return None


def _anchor_ph_idx(anchor_ref: str) -> Optional[int]:
    """Return the ``<ph-idx>`` of a ``ph.<layout-idx>.<ph-idx>`` anchor id."""
    parts = (anchor_ref or "").split(".")
    if len(parts) != 3 or parts[0] != "ph":
        return None
    try:
        return int(parts[2])
    except ValueError:
        return None


def _slide_index_from_ref(region_ref: str) -> Optional[int]:
    """Return the ``<i>`` of a ``region.slide.<i>`` ref, or None."""
    parts = (region_ref or "").split(".")
    if len(parts) != 3 or parts[0] != "region" or parts[1] != "slide":
        return None
    try:
        return int(parts[2])
    except ValueError:
        return None


def _cover_content_for(cover: ir.Cover, binds_to: Optional[str]) -> Optional[str]:
    """Resolve the new content string for a slot's ``binds_to`` key (advisory).

    ``title`` / ``subtitle`` map to the structured ``ir.Cover`` fields; anything
    else is looked up in ``cover.fields``. Returns None when the content has nothing
    for this slot (⇒ CLEAR if so ruled, else LEAVE).
    """
    if not binds_to:
        return None
    if binds_to == "title":
        return textutil.runs_to_text(cover.title or []) or str(cover.fields.get("title", "")) or None
    if binds_to == "subtitle":
        return textutil.runs_to_text(cover.subtitle or []) or str(cover.fields.get("subtitle", "")) or None
    val = cover.fields.get(binds_to)
    return str(val) if val not in (None, "") else None


# Below this confidence a DESTRUCTIVE cover CLEAR is downgraded to KEEP + WARNING.
# Additive FILL is never gated on confidence (a wrong fill is recoverable; a wrong
# delete is not - the destructive-action floor, plan §6). Mirrors the docx floor.
_DESTRUCTIVE_CONFIDENCE_FLOOR: float = 0.5


def _clear_is_corroborated(ph, slot: dict, confidence: float) -> bool:
    """Destructive-action floor for a cover CLEAR (plan §6).

    A CLEAR is corroborated only when the model's confidence clears the floor AND
    determinism agrees the slot still shows its placeholder prompt: the live
    placeholder text equals the captured ``demo_value`` (or the placeholder is
    empty). Both conditions required; otherwise the slot is kept.
    """
    if confidence < _DESTRUCTIVE_CONFIDENCE_FLOOR:
        return False
    live = ph.text.strip() if getattr(ph, "has_text_frame", False) and ph.text else ""
    if not live:
        return True
    demo_value = str(slot.get("demo_value") or "").strip()
    return bool(demo_value) and live == demo_value


def _set_placeholder_text(ph, text: str) -> None:
    """Set a placeholder's text, preserving its run formatting where possible.

    Writes into the first run of the first paragraph (keeping its run properties)
    and clears the rest, so the brand run formatting survives - the python-pptx peer
    of the docx rPr-preserving in-place fill.
    """
    if not getattr(ph, "has_text_frame", False):
        return
    tf = ph.text_frame
    paras = tf.paragraphs
    if paras and paras[0].runs:
        paras[0].runs[0].text = text
        for extra in paras[0].runs[1:]:
            extra.text = ""
        for p in paras[1:]:
            for r in p.runs:
                r.text = ""
    else:
        tf.text = text


# ---------------------------------------------------------------------------
# IR block stream -> sections -> slides (no flattening, nothing dropped)
# ---------------------------------------------------------------------------
def _sections(blocks: list) -> list[dict]:
    """Split the block stream into sections at each heading.

    One section per heading: ``{"title": <heading text>, "body": [blocks...]}``.
    Blocks before the first heading form a leading section with an empty title.
    Heading runs become the title and are NOT echoed into the body.
    """
    sections: list[dict] = []
    current: Optional[dict] = None
    for block in blocks:
        if isinstance(block, ir.Heading):
            current = {"title": textutil.runs_to_text(block.runs) or "Content", "body": []}
            sections.append(current)
        elif isinstance(block, ir.PageBreak):
            current = None
        else:
            if current is None:
                current = {"title": "Content", "body": []}
                sections.append(current)
            current["body"].append(block)
    return [s for s in sections if s["title"] or s["body"]]


def _body_lines(blocks: list) -> list[str]:
    """Render the non-heading body blocks of a section to display lines.

    Preserves block structure: list items, table rows (tab-joined cells), quotes
    (with attribution), captions and callouts each become their own line(s). KPI /
    chart / smartart / image carry a short textual stand-in so they are never
    silently dropped (full visual fidelity is later-milestone work).
    """
    lines: list[str] = []
    for block in blocks:
        if isinstance(block, ir.Paragraph):
            _append(lines, textutil.runs_to_text(block.runs))
        elif isinstance(block, ir.Callout):
            if block.title:
                _append(lines, textutil.runs_to_text(block.title))
            _append(lines, textutil.runs_to_text(block.runs))
        elif isinstance(block, ir.Quote):
            quote = textutil.runs_to_text(block.runs)
            if block.attribution:
                attribution = textutil.runs_to_text(block.attribution)
                if attribution:
                    quote = f"{quote} - {attribution}" if quote else attribution
            _append(lines, quote)
        elif isinstance(block, ir.Caption):
            _append(lines, textutil.runs_to_text(block.runs))
        elif isinstance(block, ir.ListBlock):
            for item in block.items:
                _append_list_item(lines, item)
        elif isinstance(block, ir.Table):
            _append_table(lines, block)
        elif isinstance(block, ir.Kpi):
            for kpi in block.items:
                parts = [p for p in (kpi.label, kpi.value, kpi.delta) if p]
                _append(lines, ": ".join(parts) if len(parts) > 1 else (parts[0] if parts else ""))
        elif isinstance(block, ir.Chart):
            _append(lines, block.title or "")
        elif isinstance(block, ir.Image):
            _append(lines, textutil.runs_to_text(block.caption) if block.caption else (block.alt or ""))
        # Divider / SmartArt / Component / Section / Toc carry no body text here.
    return lines


def _append(lines: list[str], text: str) -> None:
    if text:
        lines.append(text)


def _append_list_item(lines: list[str], item) -> None:
    text = textutil.runs_to_text(item.runs)
    if text:
        lines.append(("    " * max(item.level, 0)) + "• " + text)
    for sub in item.items:
        _append_list_item(lines, sub)


def _append_table(lines: list[str], table) -> None:
    header = [textutil.runs_to_text([c]) if isinstance(c, dict) else str(c) for c in table.columns]
    if any(header):
        lines.append("\t".join(header))
    for row in table.rows:
        lines.append("\t".join(textutil.runs_to_text(cell.runs) for cell in row))
    if table.caption:
        _append(lines, textutil.runs_to_text(table.caption))


# ---------------------------------------------------------------------------
# Capacity split (within a section only)
# ---------------------------------------------------------------------------
def _split_lines(lines: list[str], capacity: int) -> list[list[str]]:
    """Pack body lines into slide-sized chunks, preserving structure."""
    if not lines:
        return [[]]
    capacity = max(capacity, 1)
    chunks: list[list[str]] = []
    cur: list[str] = []
    cur_len = 0

    def flush() -> None:
        nonlocal cur, cur_len
        if cur:
            chunks.append(cur)
            cur = []
            cur_len = 0

    for line in lines:
        if len(line) > capacity:
            flush()
            for piece in _wrap_words(line, capacity):
                chunks.append([piece])
            continue
        add = len(line) + (1 if cur else 0)
        if cur and cur_len + add > capacity:
            flush()
            cur = [line]
            cur_len = len(line)
        else:
            cur.append(line)
            cur_len += add
    flush()
    return chunks or [[]]


def _wrap_words(line: str, capacity: int) -> list[str]:
    """Wrap a single over-capacity line into word-bounded pieces."""
    pieces: list[str] = []
    cur: list[str] = []
    cur_len = 0
    for word in line.split():
        add = len(word) + (1 if cur else 0)
        if cur and cur_len + add > capacity:
            pieces.append(" ".join(cur))
            cur = [word]
            cur_len = len(word)
        else:
            cur.append(word)
            cur_len += add
    if cur:
        pieces.append(" ".join(cur))
    return pieces or [line]


def _body_capacity(profile: dict) -> int:
    # Conservative L0 estimate: approximately one medium paragraph per slide.
    return int((((profile.get("qa") or {}).get("pptx_text_capacity_chars")) or 1200))


# ---------------------------------------------------------------------------
# Shell helpers
# ---------------------------------------------------------------------------
def _clear_existing_slides(prs) -> None:
    sld_id_lst = prs.slides._sldIdLst
    for sld_id in list(sld_id_lst):
        r_id = sld_id.rId
        prs.part.drop_rel(r_id)
        sld_id_lst.remove(sld_id)


def _subtitle_placeholder(slide):
    for shape in slide.placeholders:
        if shape == slide.shapes.title:
            continue
        if shape.placeholder_format.type == PP_PLACEHOLDER.SUBTITLE:
            return shape
    return None


def _first_body_placeholder(slide):
    for shape in slide.placeholders:
        if shape != slide.shapes.title:
            return shape
    return None
