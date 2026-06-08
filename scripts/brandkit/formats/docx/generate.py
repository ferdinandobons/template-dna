# SPDX-License-Identifier: MIT
"""DOCX generation from an IntermediateDocument."""

from __future__ import annotations

from pathlib import Path
from typing import Optional

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from brandkit.common import text as textutil
from brandkit.formats.docx import cover, structure
from brandkit.formats.docx.styles import lookup_style
from brandkit.ir import components
from brandkit.ir import model as ir
from brandkit.ooxml.idempotency import repack_fixed_timestamps
from brandkit.profile import schema, store
from brandkit.profile.reconcile import confidence_clears_floor
from brandkit.profile.resolver import ProfileResolver
from brandkit.qa.checks_deterministic import (
    check_index_matches_content,
    check_no_net_structure_loss,
)
from brandkit.qa.model import Finding

# Block types that carry no DOCX writer in the M1 vertical. They are skipped
# cleanly (no blank paragraph) and recorded as a degradation finding rather than
# dropped silently or rendered as an empty ``Normal`` paragraph. ``toc`` is a true
# no-op in the body (the live TOC field is refreshed separately by ``refresh_toc``)
# so it is degraded as INFO, not WARNING.
_UNHANDLED_BLOCK_TYPES: frozenset[str] = frozenset(
    {"image", "kpi", "chart", "smartart", "divider", "component", "section", "toc"}
)


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
        _write_block(doc, profile, resolver, block, sink)

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


def _write_block(
    doc,
    profile: dict,
    resolver: ProfileResolver,
    block: ir.Block,
    findings: list[Finding],
) -> None:
    if isinstance(block, ir.Heading):
        para = doc.add_paragraph(textutil.runs_to_text(block.runs))
        _apply_resolved_style(doc, para, resolver.resolve_block(block), findings)
    elif isinstance(block, ir.Paragraph):
        para = doc.add_paragraph(textutil.runs_to_text(block.runs))
        _apply_resolved_style(doc, para, resolver.resolve_block(block), findings)
    elif isinstance(block, ir.Callout):
        text = textutil.runs_to_text(block.runs)
        if block.title:
            text = textutil.runs_to_text(block.title) + "\n" + text
        para = doc.add_paragraph(text)
        _apply_resolved_style(doc, para, resolver.resolve_block(block), findings)
    elif isinstance(block, ir.ListBlock):
        _write_list_items(doc, resolver, block, block.items, findings)
    elif isinstance(block, ir.Table):
        _write_table(doc, resolver, block, findings)
    elif isinstance(block, ir.Caption):
        para = doc.add_paragraph(textutil.runs_to_text(block.runs))
        _apply_resolved_style(doc, para, resolver.resolve_block(block), findings)
    elif isinstance(block, ir.Quote):
        para = doc.add_paragraph(textutil.runs_to_text(block.runs))
        _apply_resolved_style(doc, para, resolver.resolve_block(block), findings)
        if block.attribution:
            attr = doc.add_paragraph(textutil.runs_to_text(block.attribution))
            _apply_resolved_style(
                doc, attr, resolver.resolve_role("paragraph"), findings
            )
    elif isinstance(block, ir.PageBreak):
        doc.add_page_break()
    elif block.TYPE in _UNHANDLED_BLOCK_TYPES:
        # No writer for this block in the M1 docx vertical. Skip cleanly - NEVER
        # emit a blank ``Normal`` paragraph - and record a degradation finding so
        # the dropped content is visible in QA instead of silently lost.
        sev = (
            schema.Severity.INFO.value
            if block.TYPE == "toc"
            else schema.Severity.WARNING.value
        )
        findings.append(
            Finding(
                "block_degraded",
                sev,
                f"{block.TYPE!r} block not rendered in docx M1 vertical (skipped, no placeholder emitted)",
            )
        )
    else:
        # A genuinely unknown block type is a programming error, not a degradation:
        # fail loudly rather than dropping authored content or injecting a blank.
        raise GenerationError(f"unhandled block type {block.TYPE!r}")


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
        para = doc.add_paragraph(textutil.runs_to_text(item.runs))
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
        para = doc.add_paragraph(textutil.runs_to_text(block.caption))
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
        anchor.text = textutil.runs_to_text(cell.runs)
        if (
            (force_header or getattr(cell, "header", False))
            and header_op is not None
            and header_op.resolver
        ):
            for para in anchor.paragraphs:
                _apply_resolved_style(doc, para, header_op, findings)
        c_cursor += cspan


def _as_cell(run):
    """Wrap a header-row run (a plain run dict) in a TableCell-like shim."""
    return ir.TableCell(runs=[run])


def _apply_resolved_style(doc, para, op, findings: list[Finding]) -> None:
    style = lookup_style(doc, op.resolver)
    if style is not None:
        para.style = style
        return
    # A role that names a resolver but resolves to no shell style is a brand
    # breach: be LOUD (ERROR) instead of silently leaving the paragraph as Normal.
    if op.resolver:
        target = op.resolver.get("style_id") or op.resolver.get("style_name")
        findings.append(
            Finding(
                "resolver_targets_exist",
                schema.Severity.ERROR.value,
                f"role {op.role_id!r} resolves to style {target!r} which is not in the shell",
            )
        )


def _apply_table_style(doc, table, op, findings: list[Finding]) -> None:
    style = lookup_style(doc, op.resolver)
    if style is not None:
        table.style = style
        return
    if op.resolver:
        target = op.resolver.get("style_id") or op.resolver.get("style_name")
        findings.append(
            Finding(
                "resolver_targets_exist",
                schema.Severity.ERROR.value,
                f"table role {op.role_id!r} resolves to style {target!r} which is not in the shell",
            )
        )
