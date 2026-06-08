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

  * List items are written as REAL body-placeholder paragraphs carrying their
    ``paragraph.level`` (so the layout's own list formatting supplies the bullets and
    indentation), never a string-joined ``"    • text"`` stand-in. A line's
    ``indent`` (the IR list level) survives the capacity split and is applied to the
    written paragraph.

Native PPTX objects: tables (``graphicFrame``/``a:tbl`` via ``shapes.add_table``),
KPI groups (authored as a native brand table, one row per metric), and images
(``shapes.add_picture`` from an external ``src``, sized to the body placeholder; an
unresolved source degrades loudly, never crashes), and charts (a native
``graphicFrame``/``c:chart`` via ``add_chart``, inheriting the deck theme's accent
colors so it is on-brand by construction; an unknown ``chart_type`` falls back to a
clustered column chart with an INFO note), and SmartArt (native brand-themed
autoshapes - a chevron row for a process/flow, a stacked rounded-box list otherwise,
inheriting the deck theme accent fill). A component-survival check (shell-vs-output
native counts) backs the same guarantee from the QA side.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE, PP_PLACEHOLDER

from brandkit.common import text as textutil
from brandkit.common.links import is_safe_link_url
from brandkit.formats.pptx import structure
from brandkit.ir import components
from brandkit.ir import model as ir
from brandkit.ooxml import chart as chartlib
from brandkit.ooxml.idempotency import repack_fixed_timestamps
from brandkit.profile import schema, store
from brandkit.profile.reconcile import confidence_clears_floor
from brandkit.profile.resolver import ProfileResolver
from brandkit.qa.checks_deterministic import check_no_net_structure_loss
from brandkit.qa.model import Finding

# A fixed package "modified" timestamp pinned before save so two identical
# generations are byte-identical (python-pptx stamps ``dcterms:modified`` at save
# time otherwise - the pptx peer of the xlsx pinned-timestamp idempotency fix). A
# FORMAT constant, never a brand value.
from datetime import datetime, timezone

_PINNED_MODIFIED = datetime(2001, 1, 1, tzinfo=timezone.utc)


@dataclass
class BodyLine:
    """One rendered body line carrying its list ``indent`` (paragraph level).

    Body content is rendered to a list of these (not a flat ``list[str]``) so a
    list item's level survives the capacity split and reaches the written
    paragraph as a real ``paragraph.level`` - the layout then supplies the bullet
    and indentation rather than a string-joined ``"    • "`` prefix.
    """

    text: str
    indent: int = 0
    # The rich IR runs behind ``text`` (bold/italic/underline/link). When present the
    # writer emits real runs preserving inline emphasis; ``text`` stays the flattened
    # form so the capacity split keeps measuring display width by length.
    runs: list = field(default_factory=list)

    def __len__(self) -> int:  # capacity split measures display width by text length
        return len(self.text)


@dataclass
class SlideChunk:
    """One generated content slide payload.

    Text lines and native table blocks are kept separate so table blocks can be
    authored as real PowerPoint tables instead of being flattened into body text.
    """

    lines: list[BodyLine]
    table: Optional[ir.Table] = None
    image: Optional[ir.Image] = None
    chart: Optional[ir.Chart] = None
    smartart: Optional[ir.SmartArt] = None


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
    # Expand reusable-fragment refs to primitives BEFORE resolution, mirroring the
    # docx leg (docx/generate.py): a component/section block names a profile
    # ``components``/``sections`` registry entry and is replaced in place by that
    # entry's primitive sub-blocks. An undefined ref RAISES ``ComponentExpansionError``
    # (loud, fail-closed) - symmetric with docx, so missing content is never silently
    # dropped. NOTE: registry POPULATION (auto-detect / comprehend-time fragment
    # detection) and ``slots`` PARAMETERIZATION remain DEFERRED milestones; this wires
    # only the static, profile-defined expansion plumbing.
    idoc = components.expand_components(idoc, profile)
    prs = Presentation(shell_path)
    resolver = ProfileResolver(profile)

    cover_layout = _layout_for_role(prs, resolver, "cover.title")
    content_layout = _layout_for_role(prs, resolver, "heading.1") or _layout_for_role(
        prs, resolver, "paragraph"
    )
    # The body placeholder index the profile resolved for body content (the
    # ``paragraph`` role's ``ph_idx``, ph_type ``body``). Body-placeholder
    # selection PREFERS this named idx over the positional first body placeholder,
    # so the schema's resolved ``ph_idx`` is honored (it only differs from "first"
    # on a multi-body layout). None when the role is a stub / names no idx, in
    # which case the positional fallback applies - the brand guarantee never
    # fabricates a placeholder.
    body_ph_idx = _body_ph_idx(resolver)

    if store.comprehension_is_present(profile):
        _generate_reconciled(
            prs, profile, idoc, cover_layout, content_layout, body_ph_idx, sink
        )
    else:
        _generate_deterministic(
            prs, profile, idoc, cover_layout, content_layout, body_ph_idx, sink
        )

    # Component survival (a native table/chart/picture in the shell with no
    # counterpart in the output) is checked in ONE place: the QA gate's
    # ``check_component_survival`` (checks_deterministic.py), which re-reads shell
    # and output independently for all three formats. The generator no longer runs
    # its own pre-reconcile, drop-to-zero variant - that produced duplicate
    # ``component_survival`` findings with different semantics on the same output.

    # Pin the package modified time so two identical generations are byte-identical
    # (python-pptx stamps ``dcterms:modified`` at save otherwise).
    prs.core_properties.modified = _PINNED_MODIFIED

    out = Path(output)
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(out)
    # ``dcterms:modified`` is already pinned above, so only the ZIP entry
    # timestamps need normalizing for byte-idempotent re-runs.
    repack_fixed_timestamps(out)
    return out


# ---------------------------------------------------------------------------
# Deterministic path (comprehension ABSENT) - today's behavior, byte-identical
# ---------------------------------------------------------------------------
def _generate_deterministic(
    prs, profile: dict, idoc, cover_layout, content_layout, body_ph_idx, sink: list
) -> None:
    """Blind rebuild: clear all slides, build cover + one content slide per heading.

    Unchanged from the pre-comprehension behavior so the model-free CI path stays
    the ground truth and generate-twice is byte-identical.
    """
    _clear_existing_slides(prs)

    # Cover slide: only emitted when the IR actually carries a cover title and the
    # shell offers a layout with a title placeholder.
    if idoc.cover and idoc.cover.title:
        cover_slide = prs.slides.add_slide(
            cover_layout or content_layout or prs.slide_layouts[0]
        )
        title_ph = cover_slide.shapes.title
        if title_ph is not None:
            # Re-assert the cover title placeholder's brand run formatting after the
            # value write (the pptx peer of the docx/xlsx cover-style re-assertion):
            # write into the first run keeping its rPr rather than clobbering it.
            _set_placeholder_text(title_ph, textutil.runs_to_text(idoc.cover.title))
        if idoc.cover.subtitle:
            sub = _subtitle_placeholder(cover_slide)
            if sub is not None:
                _set_placeholder_text(sub, textutil.runs_to_text(idoc.cover.subtitle))

    _append_content_slides(prs, profile, idoc, content_layout, body_ph_idx, sink)


# ---------------------------------------------------------------------------
# Reconcile path (comprehension PRESENT) - keep structural, fill cover, regen agenda
# ---------------------------------------------------------------------------
def _generate_reconciled(
    prs, profile: dict, idoc, cover_layout, content_layout, body_ph_idx, sink: list
) -> None:
    """Reconcile the preserved deck with the new content (plan §6).

    KEEP structural slides (everything the model did NOT tag ``verdict=demo``),
    CLEAR demo slides only, FILL the multi-placeholder cover in place, regenerate
    the agenda/section-list from the new headings, then append the new body slides.
    """
    comp = profile.get("comprehension") or {}

    # 1) Cover fill IN PLACE on the existing cover slide, or a freshly-added cover
    # slide when the deck ships none. Multi-placeholder, by fill_rule. Returns the
    # set of cover anchor refs the reconciliation actually CLEARED.
    cleared_anchors = _fill_cover_in_place(
        prs, profile, comp, idoc, cover_layout, content_layout, sink
    )

    # 2) CLEAR demo slides only (the destructive floor: a slide is removed only when
    # the model tagged its region demo AND determinism corroborates it is demo - its
    # body text equals a layout prompt). Structural slides are KEPT.
    removed_region_refs = _clear_demo_slides(prs, comp, sink)

    # 3) Append the new body content after the kept slides.
    _append_content_slides(prs, profile, idoc, content_layout, body_ph_idx, sink)

    # 4) Regenerate the agenda / section-list index from the NEW headings (the PPTX
    # peer of refreshing the docx outline TOC). No-op when the deck has no section
    # list / the convention is preserve.
    _regenerate_agenda(prs, comp, idoc, content_layout, body_ph_idx, sink)

    # 5) Destructive-action floor (plan §6): every cover anchor / region the
    # reconciliation removed must carry a corroborated destructive verdict AND clear
    # the confidence floor, else ERROR. Model-free; reads the frozen verdicts. Both
    # pptx reconcile sites now gate on the floor (_fill_cover_in_place and
    # _clear_demo_slides), uniform with the docx/xlsx sites, so the confidence is
    # threaded into the backstop to re-verify every sanctioned removal also cleared
    # it. confidence is the model's single comprehension value.
    confidence = (
        float(comp.get("confidence") or 0.0) if isinstance(comp, dict) else None
    )
    sink.extend(
        check_no_net_structure_loss(
            cleared_anchors | removed_region_refs, profile, confidence=confidence
        )
    )


def _fill_cover_in_place(
    prs, profile: dict, comp: dict, idoc, cover_layout, content_layout, sink: list
) -> set[str]:
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
        cover_slide = prs.slides.add_slide(
            cover_layout or content_layout or prs.slide_layouts[0]
        )

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
    # Destructive confidence floor (uniform with the docx/xlsx reconcile sites and
    # the cover sites): a demo-slide removal is honored only when the model's
    # confidence clears the floor. Below it, KEEP + WARNING (a wrong delete is not
    # recoverable). confidence is the single top-level comprehension value.
    confidence = float(comp.get("confidence") or 0.0)
    floor_ok = confidence_clears_floor(confidence)
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
        if not floor_ok:
            sink.append(
                Finding(
                    "demo_clear_downgraded",
                    schema.Severity.WARNING.value,
                    f"slide region {ref!r} clear below the destructive confidence "
                    f"floor (confidence {confidence:.2f}); kept",
                )
            )
            continue
        to_remove.append((ref, sld_ids[idx]))
    for ref, sld_id in to_remove:
        r_id = sld_id.rId
        prs.part.drop_rel(r_id)
        sld_id_lst.remove(sld_id)
        removed.add(ref)
    if removed:
        # Renumber the surviving slide parts to a CONTIGUOUS slide1..slideN sequence.
        # python-pptx allocates a new slide's partname as ``slide{len(sldIdLst)+1}.xml``
        # (count-based, see PresentationPart._next_slide_partname). Removing a demo
        # slide that is NOT the highest-indexed one leaves a gap (e.g. slide10 gone,
        # slide11 alive) so the NEXT add_slide reuses slide11.xml and collides with the
        # live part -> a duplicate ZIP part name and a corrupt OPC package. Compacting
        # partnames here (the same operation python-pptx uses on reorder) makes the
        # subsequent append/agenda add_slide land on a free name.
        prs.part.rename_slide_parts([sld_id.rId for sld_id in sld_id_lst])
    return removed


def _regenerate_agenda(
    prs, comp: dict, idoc, content_layout, body_ph_idx, sink: list
) -> None:
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
        existing = _existing_agenda_slide(prs, body_ph_idx)
        if existing is not None:
            # Refresh IN PLACE: keep the title (author's word, language) + run
            # formatting; rewrite only the section-list body to the new headings.
            body = _body_placeholder(existing, body_ph_idx)
            if body is not None:
                _set_placeholder_text(body, new_body)
        else:
            layout = content_layout or prs.slide_layouts[0]
            slide = prs.slides.add_slide(layout)
            # No agenda slide shipped: leave the layout's default (empty) title rather
            # than inject a literal; fill the body with the new section list.
            body = _body_placeholder(slide, body_ph_idx)
            if body is not None:
                body.text = new_body
        sink.append(
            Finding(
                "agenda_regenerated",
                schema.Severity.INFO.value,
                f"agenda/section-list regenerated from {len(headings)} new heading(s)",
            )
        )


def _existing_agenda_slide(prs, body_ph_idx: Optional[int] = None):
    """Return the deck's existing agenda / section-list slide, or None.

    The agenda slide is identified STRUCTURALLY, never by a fixed word: it is the
    structural slide whose body text lists the deck's OWN section names (the visible
    list page that mirrors the ``p14:sectionLst``). A slide qualifies when each of
    its body lines matches a section name from ``detect_sections`` (and it lists at
    least two), so the page is the section index in whatever language the template is
    authored in. Its title (e.g. "Sommario" / "Übersicht" / "Agenda") and formatting
    are preserved when the agenda is refreshed - the PPTX peer of ``refresh_toc``
    preserving the docx TOC heading.

    The body is read through the SAME ``body_ph_idx``-aware selection the refresh
    write path uses (:func:`_body_placeholder`), so detection and rewrite target the
    identical placeholder on a multi-body layout (falls back to the first body
    placeholder when no idx is named - identical to before on single-body layouts).

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
        body = _body_placeholder(slide, body_ph_idx)
        if body is None or not getattr(body, "has_text_frame", False) or not body.text:
            continue
        lines = [ln.strip() for ln in body.text.splitlines() if ln.strip()]
        if len(lines) >= 2 and all(ln in section_names for ln in lines):
            return slide
    return None


# ---------------------------------------------------------------------------
# Shared content-slide builder (both paths append body content the same way)
# ---------------------------------------------------------------------------
def _append_content_slides(
    prs, profile: dict, idoc, content_layout, body_ph_idx, sink: list
) -> None:
    """Append one content slide per IR heading-section (capacity-split).

    Body lines are written as REAL body-placeholder paragraphs (one per line), each
    carrying its ``BodyLine.indent`` as ``paragraph.level`` so the layout's own list
    formatting supplies the bullet/indentation - never a string-joined body blob.

    The body placeholder is chosen by the profile-resolved ``body_ph_idx`` when the
    slide carries that index (the schema's source of truth), falling back to the
    positional first body placeholder only when no named idx is present.
    """
    capacity = _body_capacity(profile)
    layout = content_layout or prs.slide_layouts[0]
    for section in _sections(idoc.blocks):
        chunks = _content_chunks(section["body"], capacity, sink)
        for page, chunk in enumerate(chunks):
            slide = prs.slides.add_slide(layout)
            title = section["title"]
            if page:
                title = f"{title} ({page + 1})"
            if slide.shapes.title is not None:
                slide.shapes.title.text = title
            body = _body_placeholder(slide, body_ph_idx)
            if chunk.table is not None:
                _clear_body_placeholder(body)
                _add_native_table(slide, prs, chunk.table, body)
            elif chunk.image is not None:
                _clear_body_placeholder(body)
                _add_native_picture(slide, prs, chunk.image, body, sink)
            elif chunk.chart is not None:
                _clear_body_placeholder(body)
                _add_native_chart(slide, prs, chunk.chart, body, sink)
            elif chunk.smartart is not None:
                _clear_body_placeholder(body)
                _add_native_smartart(slide, prs, chunk.smartart, body, sink)
            elif chunk.lines:
                if body is not None:
                    _write_body_lines(body, chunk.lines)
                else:
                    # No body placeholder on this layout: surface the unplaced body
                    # text rather than dropping it silently (the engine invariant).
                    _degrade(
                        sink,
                        "paragraph",
                        note="not placed in pptx (layout has no body placeholder)",
                    )


def _content_chunks(blocks: list, capacity: int, sink: list) -> list[SlideChunk]:
    """Split a section into slide payloads while preserving native table blocks."""
    chunks: list[SlideChunk] = []
    pending: list = []

    def flush_pending() -> None:
        nonlocal pending
        if not pending:
            return
        lines = _body_lines(pending, sink)
        for line_chunk in _split_lines(lines, capacity):
            chunks.append(SlideChunk(lines=line_chunk))
        pending = []

    for block in blocks:
        if isinstance(block, ir.Table):
            flush_pending()
            chunks.append(SlideChunk(lines=[], table=block))
        elif isinstance(block, ir.Kpi):
            # KPI -> a native brand table (label / bold value / optional delta),
            # the same faithful rendering as the docx vertical.
            flush_pending()
            kpi_table = _kpi_to_table(block)
            if kpi_table is not None:
                chunks.append(SlideChunk(lines=[], table=kpi_table))
            else:
                _degrade(sink, "kpi", note="had no items; skipped")
        elif isinstance(block, ir.Image):
            flush_pending()
            chunks.append(SlideChunk(lines=[], image=block))
        elif isinstance(block, ir.Chart):
            # Chart -> a real PowerPoint chart shape (native graphicFrame/c:chart),
            # inheriting the deck theme's accent colors (on-brand by construction).
            flush_pending()
            if chartlib.has_plottable_data(block):
                chunks.append(SlideChunk(lines=[], chart=block))
            else:
                _degrade(sink, "chart", note="had no series/categories; skipped")
        elif isinstance(block, ir.SmartArt):
            # SmartArt -> native brand-themed shapes (a chevron row for a process, a
            # stacked box list otherwise), not flattened to body text.
            flush_pending()
            if _smartart_labels(block):
                chunks.append(SlideChunk(lines=[], smartart=block))
            else:
                _degrade(sink, "smartart", note="had no nodes; skipped")
        else:
            pending.append(block)
    flush_pending()
    return chunks or [SlideChunk(lines=[])]


def _kpi_to_table(block: ir.Kpi) -> Optional[ir.Table]:
    """Build a brand table (one row per metric) from a KPI block, value bolded.

    ``columns`` is intentionally empty: KPI metrics render as label/value(/delta)
    rows with NO header row (the table writer treats empty ``columns`` as headerless)."""
    items = block.items or []
    if not items:
        return None
    has_delta = any(getattr(k, "delta", None) for k in items)
    rows = []
    for k in items:
        cells = [
            ir.TableCell(runs=textutil.normalize_runs(k.label or "")),
            ir.TableCell(runs=[{"t": str(k.value or ""), "b": True}]),
        ]
        if has_delta:
            cells.append(ir.TableCell(runs=textutil.normalize_runs(k.delta or "")))
        rows.append(cells)
    return ir.Table(columns=[], rows=rows, role="default")


def _add_native_picture(slide, prs, image: ir.Image, body_placeholder, sink) -> None:
    """Place an ``ir.Image`` as a real picture shape in the body placeholder bounds.

    Only an external ``src`` file is placed natively; an unresolved ``src``/``asset``
    degrades to a loud ``block_degraded`` WARNING (never a crash or a silent drop).
    Sized to the body width (python-pptx scales height to preserve aspect); placement
    is the layout's body affordance, not a fabricated coordinate. No brand literal."""
    src = image.src
    if not (src and Path(src).is_file()):
        _degrade(
            sink, "image", note="not placed in pptx (src/asset source unavailable)"
        )
        return
    left, top, width, _height = _table_bounds(prs, body_placeholder)
    try:
        slide.shapes.add_picture(src, left, top, width=width)
    except Exception:
        _degrade(
            sink, "image", note="not placed in pptx (image decode/placement failed)"
        )


# IR ``chart_type`` -> python-pptx XL_CHART_TYPE. A FORMAT mapping, not a brand
# value. "bar" maps to clustered COLUMNS (the common business "bar chart" is
# vertical); "barh" is the true horizontal bar. An unknown type falls back to a
# clustered column chart with an INFO note (a chart of the data is better than a
# dropped block, and the substitution is surfaced, never silent).
_CHART_TYPE_MAP = {
    "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "barh": XL_CHART_TYPE.BAR_CLUSTERED,
    "line": XL_CHART_TYPE.LINE,
    "line_markers": XL_CHART_TYPE.LINE_MARKERS,
    "area": XL_CHART_TYPE.AREA,
    "pie": XL_CHART_TYPE.PIE,
    "doughnut": XL_CHART_TYPE.DOUGHNUT,
}


def _add_native_chart(slide, prs, chart: ir.Chart, body_placeholder, sink) -> None:
    """Author an ``ir.Chart`` as a REAL PowerPoint chart (a ``graphicFrame``/
    ``c:chart`` via python-pptx ``add_chart``), sized to the body placeholder.

    The chart inherits the DECK THEME's accent colors, so it is on-brand by
    construction - no literal colors are injected. ``chart.title`` is authoritative
    when present. An unmappable type falls back to a clustered column chart (INFO);
    a build failure degrades to a loud ``block_degraded`` WARNING, never a crash or
    a silent drop. The embedded data workbook python-pptx creates is timestamp-
    normalized by ``repack_fixed_timestamps`` so generation stays byte-idempotent.
    Series coercion / single-series-truncation are the SHARED ``ooxml.chart`` gates
    (the docx vertical uses the same), so both formats agree on what is plottable.
    """
    chart_type = _CHART_TYPE_MAP.get((chart.chart_type or "bar").lower())
    if chart_type is None:
        chart_type = XL_CHART_TYPE.COLUMN_CLUSTERED
        sink.append(
            Finding(
                "chart_type_fallback",
                schema.Severity.INFO.value,
                f"chart_type {chart.chart_type!r} unknown; "
                "rendered as a clustered column chart",
            )
        )

    plottable = chartlib.coerce_series(chart)
    if not plottable:
        _degrade(sink, "chart", note="had no usable numeric series; skipped")
        return
    # A pie/doughnut shows only its first series; surface the dropped ones rather
    # than hiding them silently (the chart is still authored from the first series).
    if chartlib.is_single_series_type(chart.chart_type) and len(plottable) > 1:
        sink.append(
            Finding(
                "chart_series_truncated",
                schema.Severity.WARNING.value,
                f"a {chart.chart_type!r} chart renders only the first of "
                f"{len(plottable)} series; the others are not shown",
            )
        )
        plottable = plottable[:1]

    data = CategoryChartData()
    data.categories = [str(c) for c in chart.categories]
    for name, values in plottable:
        data.add_series(name, values)

    left, top, width, height = _table_bounds(prs, body_placeholder)
    try:
        frame = slide.shapes.add_chart(chart_type, left, top, width, height, data)
    except Exception:
        _degrade(sink, "chart", note="not placed in pptx (chart build failed)")
        return
    native = frame.chart
    title = chart.title or ""
    native.has_title = bool(title)
    if title:
        native.chart_title.text_frame.text = title


# SmartArt diagram families rendered as a stacked box LIST (vs the default chevron
# ROW for a process/flow). A FORMAT layout choice, not a brand value.
_SMARTART_LIST_DIAGRAMS = frozenset(
    {"list", "hierarchy", "pyramid", "table", "vertical_list", "bullet_list"}
)


def _smartart_node_text(node) -> str:
    """One node's label: its text, plus any child texts as sub-lines (``\\n`` makes
    real paragraphs in a pptx shape), so a nested node loses nothing."""
    if not isinstance(node, dict):
        return str(node or "").strip()
    text = str(node.get("text") or "").strip()
    kids = [
        str(c.get("text") if isinstance(c, dict) else c or "").strip()
        for c in (node.get("children") or [])
    ]
    kids = [k for k in kids if k]
    if kids:
        text = (text + "\n" + "\n".join(kids)).strip()
    return text


def _smartart_labels(smartart: ir.SmartArt) -> list[str]:
    """Non-empty node labels (text + flattened children) for the diagram."""
    return [t for t in (_smartart_node_text(n) for n in (smartart.nodes or [])) if t]


def _add_native_smartart(
    slide, prs, smartart: ir.SmartArt, body_placeholder, sink
) -> None:
    """Author an ``ir.SmartArt`` as NATIVE brand-themed autoshapes (a chevron ROW for
    a process/flow, a stacked rounded-box LIST otherwise), sized to the body
    placeholder. The shapes inherit the deck theme's accent fill, so the diagram is
    on-brand by construction - no literal color. A node's children become sub-lines
    inside its shape (never dropped). An empty diagram degrades loudly upstream."""
    labels = _smartart_labels(smartart)
    if not labels:
        _degrade(sink, "smartart", note="had no nodes; skipped")
        return
    left, top, width, height = _table_bounds(prs, body_placeholder)
    n = len(labels)
    is_list = (smartart.diagram or "process").lower() in _SMARTART_LIST_DIAGRAMS
    gap = 90000
    if is_list:  # stacked rounded-rectangle boxes
        box_h = max(360000, int((height - gap * (n - 1)) / n))
        y = top
        for label in labels:
            shp = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, left, y, width, box_h
            )
            shp.text_frame.text = label
            y += box_h + gap
    else:  # chevron row (process / flow)
        box_w = max(360000, int((width - gap * (n - 1)) / n))
        box_h = min(height, 1100000)
        x = left
        for label in labels:
            shp = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, x, top, box_w, box_h)
            shp.text_frame.text = label
            x += box_w + gap


def _clear_body_placeholder(body) -> None:
    if body is not None and getattr(body, "has_text_frame", False):
        _set_placeholder_text(body, "")


def _add_native_table(slide, prs, table: ir.Table, body_placeholder=None) -> None:
    """Author an ``ir.Table`` as a real PowerPoint table shape.

    Geometry is derived from the layout's body placeholder when available; this
    keeps placement tied to the template affordance rather than to a fabricated
    slide coordinate system. The table uses PowerPoint's native table object and
    theme/default styling, avoiding brand-specific literal colors or fonts.
    """
    col_count = _table_column_count(table)
    if col_count <= 0:
        return
    has_header = bool(table.columns)
    row_count = len(table.rows) + (1 if has_header else 0)
    if row_count <= 0:
        return

    left, top, width, height = _table_bounds(prs, body_placeholder)
    caption = textutil.runs_to_text(table.caption or []) if table.caption else ""
    caption_height = 300000 if caption else 0
    gap = 90000 if caption else 0
    usable_height = max(300000, height - caption_height - gap)
    table_height = min(usable_height, max(360000, row_count * 360000))

    gtable = slide.shapes.add_table(
        row_count, col_count, left, top, width, table_height
    )
    ppt_table = gtable.table
    row_offset = 0
    if has_header:
        for c_idx in range(col_count):
            ppt_table.cell(0, c_idx).text = _table_header_text(table, c_idx)
        row_offset = 1
    # Honor colspan/rowspan (parity with the docx table writer): place each logical
    # cell at the next free grid column, merge its spanned rectangle, and mark the
    # covered grid cells occupied so a later row's cursor skips a cell a rowspan
    # above already claimed. Without this a banner that spans columns rendered as a
    # full unmerged grid on pptx while docx merged it.
    occupied: set[tuple[int, int]] = set()
    for r_idx, row in enumerate(table.rows):
        r = r_idx + row_offset
        c_cursor = 0
        for cell in row:
            while c_cursor < col_count and (r, c_cursor) in occupied:
                c_cursor += 1
            if c_cursor >= col_count:
                break
            cspan = max(1, getattr(cell, "colspan", 1))
            rspan = max(1, getattr(cell, "rowspan", 1))
            end_c = min(c_cursor + cspan - 1, col_count - 1)
            end_r = min(r + rspan - 1, row_count - 1)
            origin = ppt_table.cell(r, c_cursor)
            if end_c > c_cursor or end_r > r:
                origin.merge(ppt_table.cell(end_r, end_c))
            origin.text = textutil.runs_to_text(getattr(cell, "runs", []) or [])
            for rr in range(r, end_r + 1):
                for cc in range(c_cursor, end_c + 1):
                    occupied.add((rr, cc))
            c_cursor = end_c + 1

    if caption:
        cap_top = top + table_height + gap
        cap = slide.shapes.add_textbox(left, cap_top, width, caption_height)
        cap.text_frame.text = caption


def _table_column_count(table: ir.Table) -> int:
    """Grid width = max(header cell count, span-EXPANDED body row width). Counting
    colspan keeps the grid wide enough to merge a multi-column banner into."""
    header = len(table.columns)
    body = max(
        (sum(max(1, getattr(c, "colspan", 1)) for c in row) for row in table.rows),
        default=0,
    )
    return max(header, body, 0)


def _table_header_text(table: ir.Table, index: int) -> str:
    if index >= len(table.columns):
        return ""
    col = table.columns[index]
    if isinstance(col, list):  # a run-list (the from_dict shape) - keep all runs
        return textutil.runs_to_text(col)
    if isinstance(col, dict):  # a bare single run (direct construction)
        return textutil.runs_to_text([col])
    return str(col)


def _table_bounds(prs, body_placeholder=None) -> tuple[int, int, int, int]:
    if body_placeholder is not None:
        left = int(getattr(body_placeholder, "left", 0) or 0)
        top = int(getattr(body_placeholder, "top", 0) or 0)
        width = int(getattr(body_placeholder, "width", 0) or 0)
        height = int(getattr(body_placeholder, "height", 0) or 0)
        if left >= 0 and top >= 0 and width > 0 and height > 0:
            return left, top, width, height

    slide_w = int(prs.slide_width)
    slide_h = int(prs.slide_height)
    return (
        int(slide_w * 0.08),
        int(slide_h * 0.24),
        int(slide_w * 0.84),
        int(slide_h * 0.62),
    )


def _write_body_lines(body, lines: list[BodyLine]) -> None:
    """Write ``lines`` as one body-placeholder paragraph each, applying levels.

    The first line reuses the placeholder's existing first paragraph (keeping its
    run formatting where present); each subsequent line is a fresh
    ``text_frame.add_paragraph``. Every paragraph's ``level`` is set from the line's
    ``indent`` so a list item renders at its real depth and the layout supplies the
    bullet glyph and indentation.
    """
    if not getattr(body, "has_text_frame", False) or not lines:
        return
    tf = body.text_frame
    # Drop any surplus prompt paragraphs so we start from a single clean paragraph.
    for extra in list(tf.paragraphs[1:]):
        extra._p.getparent().remove(extra._p)
    first = tf.paragraphs[0]
    if lines[0].runs:
        _set_para_runs(first, lines[0].runs)
    elif first.runs:
        first.runs[0].text = lines[0].text
        for r in first.runs[1:]:
            r.text = ""
    else:
        first.text = lines[0].text
    first.level = max(lines[0].indent, 0)
    for line in lines[1:]:
        para = tf.add_paragraph()
        if line.runs:
            _set_para_runs(para, line.runs)
        else:
            para.text = line.text
        para.level = max(line.indent, 0)


def _set_para_runs(para, runs) -> None:
    """Write IR ``runs`` into a pptx paragraph as real runs, preserving inline
    emphasis (bold/italic/underline) and hyperlinks, instead of a flat string.

    Existing runs are cleared first; the fresh runs inherit the placeholder's brand
    font and color from the layout's level formatting, and only the author's emphasis
    toggles + link target are set - never a literal brand font/color (guarantee holds).
    """
    for r in list(para.runs):
        r._r.getparent().remove(r._r)
    for ir_run in runs or []:
        text = str(ir_run.get("t", ""))
        link = ir_run.get("link")
        if not text and not link:
            continue
        run = para.add_run()
        run.text = text
        if ir_run.get("b"):
            run.font.bold = True
        if ir_run.get("i"):
            run.font.italic = True
        if ir_run.get("u"):
            run.font.underline = True
        # Only wire a link with a SAFE scheme; an unsafe target (file:/smb:/
        # javascript:/data:/...) is dropped while the run's text is kept, so
        # untrusted content cannot smuggle a hostile link into the deck.
        if link and is_safe_link_url(link):
            run.hyperlink.address = link


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


def _body_ph_idx(resolver: ProfileResolver) -> Optional[int]:
    """Return the placeholder idx the profile resolved for body content.

    Routes through the shared :class:`ProfileResolver`: the ``paragraph`` role's
    ``placeholder`` op carries ``ph_idx`` (ph_type ``body``) read verbatim from the
    profile's real roles. This is the named index body content must target.
    Returns None when the role is a stub (no idx) or the resolver carries none, in
    which case body-placeholder selection falls back to the positional first body
    placeholder - the brand guarantee never fabricates an index.
    """
    op = resolver.resolve_role("paragraph", fallback=None)
    if not op.resolver:
        return None
    ph_idx = op.resolver.get("ph_idx")
    return ph_idx if isinstance(ph_idx, int) else None


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
    anchors = ((profile.get("surface") or {}).get("pptx") or {}).get(
        "cover_anchors"
    ) or []
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
    return _placeholder_by_idx(slide, _anchor_ph_idx(anchor_ref))


def _placeholder_by_idx(slide, ph_idx: Optional[int]):
    """Return the placeholder on ``slide`` whose ``placeholder_format.idx`` matches.

    The stable, idx-based selection shared by the cover-anchor path
    (:func:`_placeholder_for_anchor`) and the body-content path
    (:func:`_body_placeholder`). Returns None when ``ph_idx`` is None or the slide
    carries no placeholder with that index (defensive - never fabricates one).
    """
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
        return (
            textutil.runs_to_text(cover.title or [])
            or str(cover.fields.get("title", ""))
            or None
        )
    if binds_to == "subtitle":
        return (
            textutil.runs_to_text(cover.subtitle or [])
            or str(cover.fields.get("subtitle", ""))
            or None
        )
    val = cover.fields.get(binds_to)
    return str(val) if val not in (None, "") else None


def _clear_is_corroborated(ph, slot: dict, confidence: float) -> bool:
    """Destructive-action floor for a cover CLEAR (plan §6).

    A CLEAR is corroborated only when the model's confidence clears the floor AND
    determinism agrees the slot still shows its placeholder prompt: the live
    placeholder text equals the captured ``demo_value`` (or the placeholder is
    empty). Both conditions required; otherwise the slot is kept.
    """
    if not confidence_clears_floor(confidence):
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
        # Remove surplus paragraphs entirely (not just empty their runs) so a longer
        # previous body (e.g. a section list being refreshed to fewer headings) leaves
        # no trailing empty paragraph behind.
        for p in list(paras[1:]):
            p._p.getparent().remove(p._p)
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
            current = {
                "title": textutil.runs_to_text(block.runs) or "Content",
                "body": [],
            }
            sections.append(current)
        elif isinstance(block, ir.PageBreak):
            current = None
        else:
            if current is None:
                current = {"title": "Content", "body": []}
                sections.append(current)
            current["body"].append(block)
    return [s for s in sections if s["title"] or s["body"]]


def _body_lines(blocks: list, sink: list) -> list[BodyLine]:
    """Render the non-heading body blocks of a section to structured body lines.

    Preserves block structure: list items become real ``BodyLine``s carrying their
    ``indent`` (level) so the layout supplies the bullet; quotes (with attribution),
    captions and callouts each become their own line(s).

    Tables, KPI groups, images, charts and SmartArt are intercepted by
    ``_content_chunks`` and authored as native shapes BEFORE reaching here; the
    matching branches below are a defensive fallback that only fires when
    ``_body_lines`` is called directly (e.g. a unit test). Each such fallback records
    a ``block_degraded`` WARNING on ``sink`` so the down-render is visible in QA
    (symmetric with the docx vertical), never silent.
    """
    lines: list[BodyLine] = []
    for block in blocks:
        if isinstance(block, ir.Paragraph):
            _append_runs(lines, block.runs)
        elif isinstance(block, ir.Callout):
            if block.title:
                _append_runs(lines, block.title)
            _append_runs(lines, block.runs)
        elif isinstance(block, ir.Quote):
            runs = list(block.runs or [])
            if block.attribution:
                attribution = list(block.attribution or [])
                if attribution:
                    runs = (
                        (runs + [{"t": " - "}] + attribution) if runs else attribution
                    )
            _append_runs(lines, runs)
        elif isinstance(block, ir.Caption):
            _append_runs(lines, block.runs)
        elif isinstance(block, ir.ListBlock):
            for item in block.items:
                _append_list_item(lines, item)
        elif isinstance(block, ir.Kpi):
            for kpi in block.items:
                parts = [p for p in (kpi.label, kpi.value, kpi.delta) if p]
                _append(
                    lines,
                    ": ".join(parts) if len(parts) > 1 else (parts[0] if parts else ""),
                )
            _degrade(sink, "kpi")
        elif isinstance(block, ir.Chart):
            _append(lines, block.title or "")
            _degrade(sink, "chart")
        elif isinstance(block, ir.SmartArt):
            for node in block.nodes:
                _append(
                    lines,
                    str(node.get("text") or "")
                    if isinstance(node, dict)
                    else str(node),
                )
            _degrade(sink, "smartart")
        elif isinstance(block, ir.Image):
            _append(
                lines,
                textutil.runs_to_text(block.caption)
                if block.caption
                else (block.alt or ""),
            )
            _degrade(sink, "image")
        elif isinstance(block, (ir.Component, ir.Section)):
            # Defensive / unreachable from ``generate()``: reusable-fragment refs are
            # now expanded pre-resolve on BOTH legs (``components.expand_components``
            # is wired at the top of pptx ``generate`` exactly as in docx), so a
            # defined ref becomes primitives and an undefined ref RAISES upstream -
            # neither reaches here. The branch is kept so a direct ``_body_lines``
            # caller still surfaces the unexpanded ref rather than dropping it.
            _degrade(sink, block.TYPE, note=_NOT_RENDERED)
        elif isinstance(block, (ir.Toc, ir.Divider)):
            # Structural markers with no pptx primitive in this engine. A deck uses
            # an agenda/section-list slide for navigation, not a Word-style TOC;
            # surface the unrendered block rather than dropping it silently.
            _degrade(sink, block.TYPE, note=_NOT_RENDERED)
    return lines


_NOT_RENDERED = "not rendered in pptx (no native writer)"


def _degrade(
    sink: list,
    kind: str,
    *,
    note: str = "flattened to body text in pptx (native writer deferred)",
) -> None:
    """Record a ``block_degraded`` WARNING for a block pptx cannot author natively.

    Mirrors the docx vertical's loud degradation so a deck that down-renders a
    native chart/SmartArt/KPI/image to a textual stand-in - or drops a block type
    pptx has no writer for at all (``note=_NOT_RENDERED``) - is visible in QA
    rather than silently lost. The native writers themselves are DEFERRED.
    """
    sink.append(
        Finding(
            "block_degraded",
            schema.Severity.WARNING.value,
            f"{kind!r} block {note}",
        )
    )


def _append(lines: list[BodyLine], text: str, indent: int = 0) -> None:
    if text:
        lines.append(BodyLine(text=text, indent=indent))


def _append_runs(lines: list[BodyLine], runs, indent: int = 0) -> None:
    """Append a body line carrying its rich runs (inline emphasis/links preserved)."""
    text = textutil.runs_to_text(runs)
    if text:
        lines.append(BodyLine(text=text, indent=indent, runs=list(runs or [])))


def _append_list_item(lines: list[BodyLine], item) -> None:
    text = textutil.runs_to_text(item.runs)
    if text:
        # Real paragraph level (not a string-joined "    • " prefix): the layout's
        # own list formatting supplies the bullet glyph and indentation. Carry the
        # rich runs so inline emphasis/links survive on list items too.
        lines.append(
            BodyLine(text=text, indent=max(item.level, 0), runs=list(item.runs or []))
        )
    for sub in item.items:
        _append_list_item(lines, sub)


# ---------------------------------------------------------------------------
# Capacity split (within a section only)
# ---------------------------------------------------------------------------
def _split_lines(lines: list[BodyLine], capacity: int) -> list[list[BodyLine]]:
    """Pack body lines into slide-sized chunks, preserving structure + indent.

    Each chunk is a list of :class:`BodyLine`, so a list item's ``indent`` (level)
    survives the split and reaches the written paragraph. An over-capacity single
    line is word-wrapped into pieces that all keep the original line's ``indent``.
    """
    if not lines:
        return [[]]
    capacity = max(capacity, 1)
    chunks: list[list[BodyLine]] = []
    cur: list[BodyLine] = []
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


def _wrap_words(line: BodyLine, capacity: int) -> list[BodyLine]:
    """Wrap a single over-capacity line into word-bounded pieces (keeping indent)."""
    indent = line.indent
    pieces: list[BodyLine] = []
    cur: list[str] = []
    cur_len = 0
    for word in line.text.split():
        add = len(word) + (1 if cur else 0)
        if cur and cur_len + add > capacity:
            pieces.append(BodyLine(text=" ".join(cur), indent=indent))
            cur = [word]
            cur_len = len(word)
        else:
            cur.append(word)
            cur_len += add
    if cur:
        pieces.append(BodyLine(text=" ".join(cur), indent=indent))
    return pieces or [BodyLine(text=line.text, indent=indent)]


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


def _body_placeholder(slide, ph_idx: Optional[int]):
    """Return the body placeholder, PREFERRING the profile-resolved ``ph_idx``.

    Honors the schema's source of truth: when the slide carries the placeholder the
    profile resolved for body content (its ``placeholder_format.idx`` equals
    ``ph_idx``), that one is used. Falls back to the positional first body
    placeholder when no idx is named OR the named idx is absent from this slide -
    so the brand guarantee holds (never fabricate a placeholder) while a multi-body
    layout writes into the NAMED body, not merely the first one. For a single-body
    layout the named idx IS the first body placeholder, so behavior is unchanged.
    """
    named = _placeholder_by_idx(slide, ph_idx)
    if named is not None and named != slide.shapes.title:
        return named
    return _first_body_placeholder(slide)
