# SPDX-License-Identifier: MIT
"""DOCX generation from an IntermediateDocument."""
from __future__ import annotations

from pathlib import Path
from typing import Optional

from docx import Document

from brandkit.common import text as textutil
from brandkit.formats.docx import cover, structure
from brandkit.formats.docx.styles import lookup_style
from brandkit.ir import components
from brandkit.ir import model as ir
from brandkit.profile import schema
from brandkit.profile.resolver import ProfileResolver
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
    # content is appended into the now-empty body region — immediately before the
    # sectPr — which is exactly the right slot.
    struct = profile.get("structure")
    structure.clear_body_region(doc, struct, preserve_cover=True, preserve_toc=True)

    # Fill the preserved cover anchors in place (never recreate the cover).
    cover.compose_cover(doc, idoc.cover, profile, findings=sink)

    # Write the body blocks in the order given (the body region is freeform).
    resolver = ProfileResolver(profile)
    for block in idoc.blocks:
        _write_block(doc, profile, resolver, block, sink)

    # Refresh the preserved TOC (if any) so Word recomputes it on open — the new
    # headings written into the body will be picked up. No-op when there is no TOC.
    structure.refresh_toc(doc)

    out = Path(output)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out)
    return out


def _write_block(doc, profile: dict, resolver: ProfileResolver, block: ir.Block, findings: list[Finding]) -> None:
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
            _apply_resolved_style(doc, attr, resolver.resolve_role("paragraph"), findings)
    elif isinstance(block, ir.PageBreak):
        doc.add_page_break()
    elif block.TYPE in _UNHANDLED_BLOCK_TYPES:
        # No writer for this block in the M1 docx vertical. Skip cleanly — NEVER
        # emit a blank ``Normal`` paragraph — and record a degradation finding so
        # the dropped content is visible in QA instead of silently lost.
        sev = schema.Severity.INFO.value if block.TYPE == "toc" else schema.Severity.WARNING.value
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
    level into the resolver so nested items get their level-specific list role."""
    for item in items:
        op = resolver.resolve_list_item(block, item)
        para = doc.add_paragraph(textutil.runs_to_text(item.runs))
        _apply_resolved_style(doc, para, op, findings)
        if item.items:
            _write_list_items(doc, resolver, block, item.items, findings)


def _write_table(doc, resolver, block, findings) -> None:
    # Honor colspan/rowspan: size the grid by span-expanded width, then merge.
    body_widths = [sum(max(1, c.colspan) for c in row) for row in block.rows]
    header_width = sum(max(1, _col_span(c)) for c in block.columns)
    cols = max([header_width] + body_widths + [1])
    rows = len(block.rows) + (1 if block.columns else 0)
    table = doc.add_table(rows=rows, cols=cols)
    _apply_table_style(doc, table, resolver.resolve_block(block), findings)

    header_op = resolver.resolve_role(schema.role_id("table", (block.role or "default"), "header"), fallback=None)
    row_offset = 0
    if block.columns:
        _fill_row(doc, table, 0, [_as_cell(c) for c in block.columns], header_op, findings, force_header=True)
        row_offset = 1
    for r_idx, row in enumerate(block.rows):
        _fill_row(doc, table, r_idx + row_offset, row, header_op, findings)
    if block.caption:
        para = doc.add_paragraph(textutil.runs_to_text(block.caption))
        _apply_resolved_style(doc, para, resolver.resolve_role("caption"), findings)


def _fill_row(doc, table, r_idx, cells, header_op, findings, *, force_header=False) -> None:
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
        if (force_header or getattr(cell, "header", False)) and header_op is not None and header_op.resolver:
            for para in anchor.paragraphs:
                _apply_resolved_style(doc, para, header_op, findings)
        c_cursor += cspan


def _as_cell(run):
    """Wrap a header-row run (a plain run dict) in a TableCell-like shim."""
    return ir.TableCell(runs=[run])


def _col_span(run) -> int:
    return 1


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
