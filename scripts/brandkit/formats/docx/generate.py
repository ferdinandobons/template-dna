# SPDX-License-Identifier: MIT
"""DOCX generation from an IntermediateDocument."""

from __future__ import annotations

from pathlib import Path
from typing import Optional

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.shared import Emu

from brandkit.common import text as textutil
from brandkit.common.links import is_safe_link_url
from brandkit.formats.docx import cover, structure
from brandkit.formats.docx.styles import lookup_style
from brandkit.ir import components
from brandkit.ir import model as ir
from brandkit.ooxml import chart as chartlib
from brandkit.ooxml.idempotency import repack_fixed_timestamps
from brandkit.profile import schema, store
from brandkit.profile.reconcile import confidence_clears_floor
from brandkit.profile.resolver import ProfileResolver
from brandkit.qa.checks_deterministic import (
    check_index_matches_content,
    check_no_net_structure_loss,
)
from brandkit.qa.model import Finding

# Block types that carry no DOCX writer. ``component``/``section`` are EXPANDED to
# primitives before this writer runs (``expand_components``), so they should never
# reach it; this set is the defensive fail-loud backstop if one ever survives. A
# block in this set is skipped cleanly (no blank paragraph) and recorded as a
# degradation finding rather than dropped silently. ``toc`` is now a native writer
# (``_write_toc``), so it is no longer here.
_UNHANDLED_BLOCK_TYPES: frozenset[str] = frozenset({"component", "section"})


class GenerationError(ValueError):
    """Raised when a block cannot be written and must not be silently dropped."""


def generate(
    profile: dict,
    shell_path: str | Path,
    idoc: ir.IntermediateDocument,
    output: str | Path,
    *,
    findings: Optional[list[Finding]] = None,
) -> Path:
    """Generate a DOCX from ``idoc`` onto the brand ``shell_path``.

    ``findings`` (optional out-param) is appended with any degradation / style-miss
    findings recorded during writing so the QA gate can surface them. Off-brand
    output stays impossible by construction: every style comes from the profile via
    :func:`brandkit.formats.docx.styles.lookup_style`, and a role that resolves to
    no matching shell style is recorded LOUDLY (ERROR) instead of silently falling
    back to ``Normal``.
    """
    sink: list[Finding] = findings if findings is not None else []
    idoc = components.expand_components(idoc, profile)
    doc = Document(shell_path)

    # Order-aware body replacement: remove ONLY the freeform body region, keeping
    # the ordered cover and TOC regions (and the final sectPr) in place. New
    # content is appended into the now-empty body region - immediately before the
    # sectPr - which is exactly the right slot.
    struct = profile.get("structure")
    structure.clear_body_region(doc, struct, preserve_cover=True, preserve_toc=True)

    # Reconcile the preserved derived indexes (TOC/TOF/TOT) and demo regions with
    # the new content FIRST, comprehension-steered when present (no-op when absent).
    # This MUST run before the cover fill: the surfaced index ids (and the orphan-
    # index-heading lookup) are position-based over the top-level body children, and
    # ``compose_cover`` inserts a paragraph into the cover region, shifting every
    # subsequent child index. The body clear above leaves the cover/TOC positions
    # untouched (the body region is last), so resolving the index refs here - before
    # any cover insert - keeps the position ids valid. Removing a stale caption index
    # also removes its introducing heading (the orphan-index-heading fix).
    removed_refs = _reconcile_indexes_and_demo(doc, profile, idoc, sink)

    # Caption-index reconciliation can expose blank TOC separators and preserved
    # demo-body section breaks at the start of the body slot. Prune them only when
    # this generation actually removed such structure; otherwise keep historical
    # multi-section geometry byte-for-byte.
    if removed_refs:
        structure.prune_leading_empty_body_artifacts(doc)

    # Fill the preserved cover anchors in place (never recreate the cover). Returns
    # the set of cover anchor refs the reconciliation CLEARED, for the destructive
    # floor below.
    cleared_anchors = cover.compose_cover(doc, idoc.cover, profile, findings=sink)

    # Write the body blocks in the order given (the body region is freeform).
    resolver = ProfileResolver(profile)
    for block in idoc.blocks:
        _write_block(doc, resolver, block, sink)

    # Keep the visible TOC field cache aligned for renderers that do not update
    # fields in headless export. The field itself remains dirty/updateable.
    structure.refresh_visible_outline_toc_cache(doc, _outline_headings(idoc.blocks))

    # Refresh the preserved TOC (if any) so Word recomputes it on open - the new
    # headings written into the body will be picked up. No-op when there is no TOC.
    structure.refresh_toc(doc)

    # Destructive-action floor (plan §6): every cover anchor / index block the
    # reconciliation removed must carry a corroborated destructive verdict AND clear
    # the confidence floor, else ERROR (a wrong delete is not recoverable). Model-free;
    # reads frozen verdicts. The confidence threaded here is the model's single
    # comprehension confidence (the same value the reconcile sites gate on).
    if store.comprehension_is_present(profile):
        comp = profile.get("comprehension")
        confidence = (
            float(comp.get("confidence") or 0.0) if isinstance(comp, dict) else None
        )
        sink.extend(
            check_no_net_structure_loss(
                cleared_anchors | removed_refs, profile, confidence=confidence
            )
        )

    out = Path(output)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out)
    # python-docx's core.xml is already stable, so only the ZIP entry timestamps
    # need normalizing for byte-idempotent re-runs (no modified<-created pin).
    repack_fixed_timestamps(out)
    return out


def _outline_headings(blocks: list) -> list[tuple[int, str]]:
    headings: list[tuple[int, str]] = []
    for block in blocks:
        if isinstance(block, ir.Heading):
            text = textutil.runs_to_text(block.runs).strip()
            if text:
                headings.append((block.level, text))
    return headings


def _content_has_captionables(idoc: ir.IntermediateDocument) -> bool:
    """True when the new content carries any captionable item (caption/table/image).

    A derived caption index (table-of-tables / table-of-figures) has something to
    point at only when the content contains captionable blocks. This is the
    deterministic corroboration for KEEPING such an index; an empty result
    corroborates REMOVING it (plan §6 caption/index reconciliation).
    """
    for block in idoc.blocks:
        if isinstance(block, ir.Caption):
            return True
        if isinstance(block, ir.Table) and getattr(block, "caption", None):
            return True
        if isinstance(block, ir.Image) and getattr(block, "caption", None):
            return True
    return False


def _reconcile_indexes_and_demo(
    doc, profile: dict, idoc: ir.IntermediateDocument, findings: list[Finding]
) -> set[str]:
    """Reconcile preserved derived indexes + demo regions with the new content.

    Comprehension-steered (no-op when comprehension is absent - the deterministic
    body clear already removed stale body content, and the outline TOC is refreshed
    separately). For each ``conventions.indexes`` entry:

      - ``reconcile == clear``  -> REMOVE the orphan index block, but ONLY when the
        destructive floor corroborates it: the model also tagged the index's own
        region (``region.<field-id>``) ``verdict=demo``, OR the content carries no
        captionable item the index could point at. Otherwise KEEP + WARNING.
      - ``reconcile in {regenerate, preserve}``  -> KEEP (the outline TOC is
        refreshed by ``refresh_toc``; a caption index is left for Word to recompute).

    Demo regions tagged ``verdict=demo`` that map to a derived-index region are
    removed via the same index path; demo regions over the freeform body are
    already gone (the body clear). Returns the set of index refs actually removed
    (for ``no_net_structure_loss``).
    """
    comp = profile.get("comprehension")
    if not store.comprehension_is_present(profile) or not isinstance(comp, dict):
        return set()

    has_captionables = _content_has_captionables(idoc)
    # The model's single comprehension confidence - the SAME value the cover
    # reconciler gates on - so a CLEAR is honored uniformly across formats.
    confidence = float(comp.get("confidence") or 0.0)
    demo_region_refs = {
        r.get("region_ref")
        for r in (comp.get("demo_classification") or {}).get("regions") or []
        if isinstance(r, dict) and r.get("verdict") == schema.Verdict.DEMO.value
    }

    # Collect every index ruled CLEAR whose deletion the destructive floor
    # corroborates, then remove them all in one shift-safe pass (a position-based
    # field id would otherwise be invalidated by a prior removal).
    to_clear: list[str] = []
    for idx in (comp.get("conventions") or {}).get("indexes") or []:
        if not isinstance(idx, dict):
            continue
        index_ref = idx.get("index_ref")
        if not index_ref:
            continue
        if idx.get("reconcile") != schema.Reconcile.CLEAR.value:
            continue  # regenerate / preserve: keep (refresh handles the outline TOC)
        # Destructive floor: a CLEAR is honored only when (i) the model's confidence
        # clears the floor AND (ii) determinism corroborates the removal - the model
        # also tagged the index's region demo, or the content has no captionable item
        # the index could point at. Either gate failing downgrades to KEEP + WARNING
        # (a wrong delete is unrecoverable). The confidence gate mirrors the cover
        # reconcilers so the SAME confidence yields the SAME behavior per format.
        if not confidence_clears_floor(confidence):
            findings.append(
                Finding(
                    "index_clear_downgraded",
                    schema.Severity.WARNING.value,
                    f"index {index_ref!r} clear not corroborated "
                    f"(confidence {confidence:.2f}); kept",
                )
            )
            continue
        region_ref = f"region.{index_ref}"
        corroborated = (region_ref in demo_region_refs) or (not has_captionables)
        if not corroborated:
            findings.append(
                Finding(
                    "index_clear_downgraded",
                    schema.Severity.WARNING.value,
                    f"index {index_ref!r} clear not corroborated "
                    f"(no demo verdict and content has captionable items); kept",
                )
            )
            continue
        to_clear.append(index_ref)
    removed = structure.remove_index_fields(doc, to_clear)

    # Self-awareness check (plan §7): a preserved CAPTION index (one carrying a
    # deterministic ``seq_id`` from its ``\c`` switch) that is KEPT yet has no
    # captionable content to point at is flagged WARNING. The set of emitted SEQ
    # classes is keyed on the opaque ``seq_id`` (never the advisory ``kind``): a
    # caption index's seq is "emitted" only when the content carries captionable
    # items. Empty when none, so every kept-but-stale caption index is surfaced.
    present_seq_ids: set[str] = set()
    if has_captionables:
        for idx in (comp.get("conventions") or {}).get("indexes") or []:
            if not isinstance(idx, dict):
                continue
            if idx.get("index_ref") in removed:
                continue  # already cleared; not a kept index
            seq = idx.get("seq_id")
            if seq:
                present_seq_ids.add(seq)
    findings.extend(check_index_matches_content(present_seq_ids, profile))
    return removed


def _apply_run_toggles(run, r: dict) -> None:
    """Apply inline emphasis from an IR run to a docx run.

    Only character-level toggles (``w:b``/``w:i``/``w:u``/strike/super-/subscript) -
    author intent, NOT a brand style/font/color - are applied, so the brand
    guarantee (generators never write a literal style/hex/font) is untouched. The
    run still inherits the paragraph style's brand font and color. A ``code`` toggle
    carries no brand-safe typeface (a monospace family would be a literal font), so
    it is rendered as plain text rather than fabricating one.
    """
    if r.get("b"):
        run.bold = True
    if r.get("i"):
        run.italic = True
    if r.get("u"):
        run.underline = True
    if r.get("strike"):
        run.font.strike = True
    if r.get("sup"):
        run.font.superscript = True
    if r.get("sub"):
        run.font.subscript = True


def _add_hyperlink(para, url: str, text: str, r: dict) -> None:
    """Append a real ``w:hyperlink`` (external relationship) carrying ``text``.

    The link target is the author's URL (content, not brand), wired through a
    package relationship. The visible run keeps the author's inline emphasis; we do
    not inject a literal ``Hyperlink`` character style or color (brand guarantee).

    An UNSAFE scheme (``file:``/``smb:``/``javascript:``/``data:``/...) is neutralized:
    the author's TEXT is kept (emphasis preserved) but the dangerous target is not
    wired, so untrusted content cannot smuggle a hostile link into the document.
    """
    if not is_safe_link_url(url):
        _apply_run_toggles(para.add_run(text), r)
        return
    r_id = para.part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    run = OxmlElement("w:r")
    if any(r.get(k) for k in ("b", "i", "u", "strike", "sup", "sub")):
        rpr = OxmlElement("w:rPr")
        for key, tag in (("b", "w:b"), ("i", "w:i"), ("strike", "w:strike")):
            if r.get(key):
                rpr.append(OxmlElement(tag))
        if r.get("u"):
            u = OxmlElement("w:u")
            u.set(qn("w:val"), "single")
            rpr.append(u)
        if r.get("sup") or r.get("sub"):
            va = OxmlElement("w:vertAlign")
            va.set(qn("w:val"), "superscript" if r.get("sup") else "subscript")
            rpr.append(va)
        run.append(rpr)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    run.append(t)
    hyperlink.append(run)
    para._p.append(hyperlink)


def _add_runs(para, runs) -> None:
    """Write IR ``runs`` into ``para`` as real docx runs, preserving inline emphasis
    and hyperlinks, instead of flattening to a single plain run."""
    for r in runs or []:
        text = str(r.get("t", ""))
        link = r.get("link")
        if link and text:
            _add_hyperlink(para, link, text, r)
        elif text:
            _apply_run_toggles(para.add_run(text), r)


def _para_with_runs(doc, runs):
    """Add a paragraph carrying ``runs`` as real, formatting-preserving docx runs."""
    para = doc.add_paragraph()
    _add_runs(para, runs)
    return para


def _write_block(
    doc,
    resolver: ProfileResolver,
    block: ir.Block,
    findings: list[Finding],
) -> None:
    if isinstance(block, ir.Heading):
        para = _para_with_runs(doc, block.runs)
        _apply_resolved_style(doc, para, resolver.resolve_block(block), findings)
    elif isinstance(block, ir.Paragraph):
        para = _para_with_runs(doc, block.runs)
        _apply_resolved_style(doc, para, resolver.resolve_block(block), findings)
    elif isinstance(block, ir.Callout):
        para = doc.add_paragraph()
        if block.title:
            _add_runs(para, block.title)
            para.add_run().add_break()  # title above body, same callout paragraph
        _add_runs(para, block.runs)
        _apply_resolved_style(doc, para, resolver.resolve_block(block), findings)
    elif isinstance(block, ir.ListBlock):
        _write_list_items(doc, resolver, block, block.items, findings)
    elif isinstance(block, ir.Table):
        _write_table(doc, resolver, block, findings)
    elif isinstance(block, ir.Caption):
        para = _para_with_runs(doc, block.runs)
        _apply_resolved_style(doc, para, resolver.resolve_block(block), findings)
    elif isinstance(block, ir.Quote):
        para = _para_with_runs(doc, block.runs)
        _apply_resolved_style(doc, para, resolver.resolve_block(block), findings)
        if block.attribution:
            attr = _para_with_runs(doc, block.attribution)
            _apply_resolved_style(
                doc, attr, resolver.resolve_role("paragraph"), findings
            )
    elif isinstance(block, ir.PageBreak):
        doc.add_page_break()
    elif isinstance(block, ir.Divider):
        # Native horizontal rule: an empty paragraph carrying a bottom border. Color
        # is OOXML "auto" (resolves to the theme's text color) - never a literal hex,
        # so the brand guarantee holds without a profile-defined divider artifact.
        para = doc.add_paragraph()
        pbdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "6")
        bottom.set(qn("w:space"), "1")
        bottom.set(qn("w:color"), "auto")
        pbdr.append(bottom)
        para._p.get_or_add_pPr().append(pbdr)
    elif isinstance(block, ir.Image):
        _write_image(doc, resolver, block, findings)
    elif isinstance(block, ir.Kpi):
        _write_kpi(doc, resolver, block, findings)
    elif isinstance(block, ir.Chart):
        _write_chart(doc, block, findings)
    elif isinstance(block, ir.SmartArt):
        _write_smartart(doc, resolver, block, findings)
    elif isinstance(block, ir.Toc):
        _write_toc(doc, resolver, block, findings)
    elif block.TYPE in _UNHANDLED_BLOCK_TYPES:
        # No writer for this block. Skip cleanly - NEVER emit a blank ``Normal``
        # paragraph - and record a degradation finding so the skipped content is
        # visible in QA instead of silently lost. (component/section are normally
        # expanded away before this writer; reaching here is the defensive path.)
        findings.append(
            Finding(
                "block_degraded",
                schema.Severity.WARNING.value,
                f"{block.TYPE!r} block not rendered in docx vertical (skipped, no placeholder emitted)",
            )
        )
    else:
        # A genuinely unknown block type is a programming error, not a degradation:
        # fail loudly rather than dropping authored content or injecting a blank.
        raise GenerationError(f"unhandled block type {block.TYPE!r}")


def _write_toc(doc, resolver, block, findings) -> None:
    """Render an authored ``toc`` block as a native, updateable table of contents.

    If the shell already carries a structural TOC (preserved through the body
    clear), DEFER to it: that canonical TOC is refreshed by ``refresh_toc`` and a
    second one would duplicate it (the interplay this writer is careful to avoid).
    Otherwise author a native outline TOC field at the block's flow position; its
    visible cache is filled from the generated headings by
    ``refresh_visible_outline_toc_cache`` and it is marked dirty + ``updateFields``
    by ``refresh_toc`` (both run after the body is built), so Word rebuilds it on
    open. An optional title heading uses the resolved ``toc`` role style.
    """
    if structure.is_outline_toc_present(doc):
        findings.append(
            Finding(
                "toc_deferred",
                schema.Severity.INFO.value,
                "toc block deferred to the template's preserved outline table of "
                "contents (already present; no duplicate emitted)",
            )
        )
        return
    if block.title:
        para = doc.add_paragraph()
        para.add_run(str(block.title))
        _apply_resolved_style(doc, para, resolver.resolve_role("toc"), findings)
    structure.append_outline_toc_field(doc, max_level=block.max_level)


def _write_kpi(doc, resolver, block, findings) -> None:
    """Render a KPI / metric-card group as a brand-styled table (one row per metric:
    label, value, optional delta), value bolded for prominence.

    Reuses the table writer so the metric grid carries the profile's table style (no
    fabricated KPI box style/color): a faithful native rendering instead of a dropped
    block. The delta column is emitted only when at least one metric carries a delta.
    """
    items = block.items or []
    if not items:
        findings.append(
            Finding(
                "block_degraded",
                schema.Severity.WARNING.value,
                "'kpi' block had no items; skipped",
            )
        )
        return
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
    _write_table(
        doc, resolver, ir.Table(columns=[], rows=rows, role="default"), findings
    )


def _write_image(doc, resolver, block, findings) -> None:
    """Place an ``Image`` block as an inline picture, sized to the section's content
    width when no explicit size hint is given, with its caption below.

    Only an external ``src`` file is placed natively; an unresolved ``src``/``asset``
    (file missing, or only a profile asset id that cannot be loaded here) degrades to
    a loud ``block_degraded`` WARNING rather than crashing, so authored figures are
    realized when available and never silently lost otherwise. No brand literal is
    written: the picture is the author's asset, sized by layout geometry.
    """
    src = block.src
    placed = False
    if src and Path(src).is_file():
        sec = doc.sections[-1]
        width = (
            Emu(int(block.width_emu))
            if block.width_emu
            else sec.page_width - sec.left_margin - sec.right_margin
        )
        height = Emu(int(block.height_emu)) if block.height_emu else None
        try:
            run = doc.add_paragraph().add_run()
            if height is not None:
                run.add_picture(src, width=width, height=height)
            else:
                run.add_picture(src, width=width)
            placed = True
        except Exception:
            # Any image-decode / placement failure degrades rather than crashing.
            placed = False
    if not placed:
        findings.append(
            Finding(
                "block_degraded",
                schema.Severity.WARNING.value,
                "'image' block not placed in docx (src/asset source unavailable); skipped",
            )
        )
        return
    if block.caption:
        para = _para_with_runs(doc, block.caption)
        _apply_resolved_style(doc, para, resolver.resolve_role("caption"), findings)


_CHART_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
)
_CHART_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
)
_NS_CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"
_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
_NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _write_chart(doc, block, findings) -> None:
    """Author an ``ir.Chart`` as a NATIVE Word chart (a real DrawingML
    ``c:chartSpace`` part referenced by an inline ``w:drawing``), not flattened to
    text. The chart carries INLINE cached data (no embedded workbook), so it stays
    byte-idempotent, and writes no literal colors, so it inherits the document
    theme's accent colors (on-brand by construction). An unknown ``chart_type``
    falls back to a clustered column chart (surfaced as INFO); an empty/all-non-
    numeric chart degrades to a loud ``block_degraded`` WARNING - never a crash or a
    silent drop. ``title`` is authoritative when present.
    """
    if not chartlib.has_plottable_data(block):
        findings.append(
            Finding(
                "block_degraded",
                schema.Severity.WARNING.value,
                "'chart' block had no plottable numeric data; skipped",
            )
        )
        return
    if not chartlib.is_known_chart_type(block.chart_type):
        findings.append(
            Finding(
                "chart_type_fallback",
                schema.Severity.INFO.value,
                f"chart_type {block.chart_type!r} unknown; "
                "rendered as a clustered column chart",
            )
        )

    series = chartlib.coerce_series(block)
    if chartlib.is_single_series_type(block.chart_type) and len(series) > 1:
        findings.append(
            Finding(
                "chart_series_truncated",
                schema.Severity.WARNING.value,
                f"a {block.chart_type!r} chart renders only the first of "
                f"{len(series)} series; the others are not shown",
            )
        )
    xml = chartlib.build_chart_xml(
        block.chart_type, series, block.categories, block.title
    )

    # New chart part with a STABLE partname (derived from the count of existing
    # chart parts, so a re-run produces the same name) + the relationship from the
    # document part. python-docx auto-registers the content-type override at save.
    existing = sum(
        1
        for p in doc.part.package.iter_parts()
        if str(p.partname).startswith("/word/charts/chart")
    )
    partname = PackURI(f"/word/charts/chart{existing + 1}.xml")
    part = Part(partname, _CHART_CONTENT_TYPE, xml, doc.part.package)
    r_id = doc.part.relate_to(part, _CHART_REL_TYPE)

    # ``wp:docPr@id`` must be UNIQUE across every drawing object in the document
    # (images use python-docx's global max+1; a chart-only counter would collide
    # with an image's id and make Word refuse the file). Derive the next id from the
    # max of all existing ``wp:docPr`` ids - deterministic given the fixed
    # generation order, so byte-idempotency holds.
    doc_pr_ids = [
        int(el.get("id"))
        for el in doc.element.iter(qn("wp:docPr"))
        if (el.get("id") or "").isdigit()
    ]
    doc_pr_id = max(doc_pr_ids, default=0) + 1

    # Inline drawing sized to the section content width with a 3:2 aspect (matches
    # the image writer's layout-driven sizing; no fabricated coordinate). The
    # ``c:chart`` graphicData references the chart part by rId. The drawing goes into
    # a run of a paragraph added via ``add_paragraph`` so it lands BEFORE the body's
    # final ``w:sectPr`` (appending to the body directly would place it after, which
    # is invalid).
    sec = doc.sections[-1]
    cx = int(sec.page_width - sec.left_margin - sec.right_margin)
    cy = int(cx * 2 / 3)
    drawing = (
        f'<w:drawing xmlns:w="{_NS_W}" xmlns:wp="{_NS_WP}" xmlns:a="{_NS_A}" '
        f'xmlns:r="{_NS_R}" xmlns:c="{_NS_CHART}">'
        f'<wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/>'
        f'<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{doc_pr_id}" name="Chart {existing + 1}"/>'
        f"<wp:cNvGraphicFramePr/>"
        f'<a:graphic><a:graphicData uri="{_NS_CHART}">'
        f'<c:chart r:id="{r_id}"/>'
        f"</a:graphicData></a:graphic></wp:inline></w:drawing>"
    )
    run = doc.add_paragraph().add_run()
    run._r.append(parse_xml(drawing))


# SmartArt diagram families rendered as a single-COLUMN table (one row per node) vs
# the default single-ROW process strip. A FORMAT layout choice, not a brand value.
_SMARTART_LIST_DIAGRAMS = frozenset(
    {"list", "hierarchy", "pyramid", "table", "vertical_list", "bullet_list"}
)


def _smartart_node_text(node) -> str:
    """One node's label: its text plus any child texts inline (a docx table cell has
    no soft line break, so children are joined with ' - '), so nesting is not lost."""
    if not isinstance(node, dict):
        return str(node or "").strip()
    text = str(node.get("text") or "").strip()
    kids = [
        str(c.get("text") if isinstance(c, dict) else c or "").strip()
        for c in (node.get("children") or [])
    ]
    kids = [k for k in kids if k]
    if kids:
        text = f"{text} - {'; '.join(kids)}".strip(" -")
    return text


def _write_smartart(doc, resolver, block, findings) -> None:
    """Author an ``ir.SmartArt`` as a NATIVE brand-styled table, reusing the table
    writer: a process/flow becomes a single ROW (one cell per step), a list/
    hierarchy a single COLUMN (one row per node). On-brand via the profile's table
    style - never a fabricated diagram color. An empty diagram degrades to a loud
    ``block_degraded`` WARNING, never a silent drop. (The pptx vertical renders the
    same SmartArt as native chevron/box shapes.)"""
    labels = [t for t in (_smartart_node_text(n) for n in (block.nodes or [])) if t]
    if not labels:
        findings.append(
            Finding(
                "block_degraded",
                schema.Severity.WARNING.value,
                "'smartart' block had no nodes; skipped",
            )
        )
        return
    is_list = (block.diagram or "process").lower() in _SMARTART_LIST_DIAGRAMS
    if is_list:
        rows = [[ir.TableCell(runs=[{"t": label}])] for label in labels]
    else:
        rows = [[ir.TableCell(runs=[{"t": label}]) for label in labels]]
    _write_table(
        doc, resolver, ir.Table(columns=[], rows=rows, role="default"), findings
    )


def _write_list_items(doc, resolver, block, items, findings) -> None:
    """Write a list depth-first, one paragraph per item, threading each item's
    level into the resolver so nested items get their level-specific list role.

    Applies the resolved list PARAGRAPH style AND re-asserts the numbering
    (``w:pPr/w:numPr``) on each item (D1): a style's own ``w:numPr`` is not
    inherited onto a python-docx ``add_paragraph``, so without this the list would
    render as flat un-numbered paragraphs. The ``numId`` is the verbatim id the
    role nominated from the document's numbering part; the ``ilvl`` follows the
    item's actual nesting level so deeper items indent under the same definition.
    """
    for item in items:
        op = resolver.resolve_list_item(block, item)
        para = _para_with_runs(doc, item.runs)
        _apply_resolved_style(doc, para, op, findings)
        _apply_list_numbering(para, op, item)
        if item.items:
            _write_list_items(doc, resolver, block, item.items, findings)


def _apply_list_numbering(para, op, item) -> None:
    """Write ``w:pPr/w:numPr`` (numId + ilvl) on a list paragraph from a resolved
    list role that carries a verbatim ``num_id``.

    No-op when the role carries no ``num_id`` (e.g. a ``Normal``-floor list with no
    real numbering definition, or a non-list role) - the paragraph then keeps only
    its style, exactly as before. The ``ilvl`` is the item's own nesting level so a
    nested item indents correctly even when its role fell back to the level-1
    list role.
    """
    resolver = op.resolver or {}
    num_id = resolver.get("num_id")
    if not num_id:
        return
    try:
        num_id_int = int(num_id)
    except (TypeError, ValueError):
        return
    ilvl = max(0, int(getattr(item, "level", 0) or 0))
    pPr = para._p.get_or_add_pPr()
    # Remove any inherited/duplicate numPr first so re-runs stay idempotent.
    for existing in pPr.findall(qn("w:numPr")):
        pPr.remove(existing)
    numPr = OxmlElement("w:numPr")
    ilvl_el = OxmlElement("w:ilvl")
    ilvl_el.set(qn("w:val"), str(ilvl))
    num_el = OxmlElement("w:numId")
    num_el.set(qn("w:val"), str(num_id_int))
    numPr.append(ilvl_el)
    numPr.append(num_el)
    pPr.append(numPr)


def _write_table(doc, resolver, block, findings) -> None:
    # Honor colspan/rowspan: size the grid by span-expanded width, then merge.
    # Header cells are a flat run list (one column each), so the header width is
    # simply the column count.
    body_widths = [sum(max(1, c.colspan) for c in row) for row in block.rows]
    header_width = len(block.columns)
    cols = max([header_width] + body_widths + [1])
    rows = len(block.rows) + (1 if block.columns else 0)
    table = doc.add_table(rows=rows, cols=cols)
    _apply_table_style(doc, table, resolver.resolve_block(block), findings)

    header_op = resolver.resolve_role(
        schema.role_id("table", (block.role or "default"), "header"), fallback=None
    )
    row_offset = 0
    if block.columns:
        _fill_row(
            doc,
            table,
            0,
            [_as_cell(c) for c in block.columns],
            header_op,
            findings,
            force_header=True,
        )
        row_offset = 1
    for r_idx, row in enumerate(block.rows):
        _fill_row(doc, table, r_idx + row_offset, row, header_op, findings)
    if block.caption:
        para = _para_with_runs(doc, block.caption)
        _apply_resolved_style(doc, para, resolver.resolve_role("caption"), findings)


def _fill_row(
    doc, table, r_idx, cells, header_op, findings, *, force_header=False
) -> None:
    """Fill one logical row, honoring colspan/rowspan by merging grid cells."""
    c_cursor = 0
    ncols = len(table.columns)
    for cell in cells:
        if c_cursor >= ncols:
            break
        cspan = max(1, getattr(cell, "colspan", 1))
        rspan = max(1, getattr(cell, "rowspan", 1))
        anchor = table.cell(r_idx, c_cursor)
        # Merge the spanned rectangle so the column grid stays aligned.
        if cspan > 1 or rspan > 1:
            end_c = min(c_cursor + cspan - 1, ncols - 1)
            end_r = min(r_idx + rspan - 1, len(table.rows) - 1)
            anchor = anchor.merge(table.cell(end_r, end_c))
        # Run-aware cell text: a fresh cell has one empty paragraph; write the IR
        # runs into it so inline emphasis/links survive (not a flattened cell.text).
        _add_runs(anchor.paragraphs[0], cell.runs)
        if (
            (force_header or getattr(cell, "header", False))
            and header_op is not None
            and header_op.resolver
        ):
            for para in anchor.paragraphs:
                _apply_resolved_style(doc, para, header_op, findings)
        c_cursor += cspan


def _as_cell(col):
    """Wrap a header column in a ``TableCell`` shim. ``col`` is a run-list (the
    shape ``Table.from_dict`` produces, so multi-run header emphasis survives); a
    bare run dict / string from a direct construction is tolerated."""
    runs = col if isinstance(col, list) else [col]
    return ir.TableCell(runs=runs)


def _apply_style(doc, target_obj, op, findings: list[Finding], *, label: str) -> None:
    """Apply a resolved role's style to a paragraph OR a table (both expose
    ``.style``). A role that names a resolver but resolves to no shell style is a
    brand breach: be LOUD (ERROR) instead of silently leaving the default style.
    The single implementation behind ``_apply_resolved_style`` / ``_apply_table_style``."""
    style = lookup_style(doc, op.resolver)
    if style is not None:
        target_obj.style = style
        return
    if op.resolver:
        target = op.resolver.get("style_id") or op.resolver.get("style_name")
        findings.append(
            Finding(
                "resolver_targets_exist",
                schema.Severity.ERROR.value,
                f"{label}role {op.role_id!r} resolves to style {target!r} "
                "which is not in the shell",
            )
        )


def _apply_resolved_style(doc, para, op, findings: list[Finding]) -> None:
    _apply_style(doc, para, op, findings, label="")


def _apply_table_style(doc, table, op, findings: list[Finding]) -> None:
    _apply_style(doc, table, op, findings, label="table ")
