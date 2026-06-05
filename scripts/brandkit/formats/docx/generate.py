# SPDX-License-Identifier: MIT
"""DOCX generation from an IntermediateDocument."""
from __future__ import annotations

from pathlib import Path

from docx import Document

from brandkit.common import text as textutil
from brandkit.formats.docx import cover, structure
from brandkit.ir import components
from brandkit.ir import model as ir
from brandkit.profile.resolver import ProfileResolver


def generate(profile: dict, shell_path: str | Path, idoc: ir.IntermediateDocument, output: str | Path) -> Path:
    idoc = components.expand_components(idoc, profile)
    doc = Document(shell_path)

    # Order-aware body replacement: remove ONLY the freeform body region, keeping
    # the ordered cover and TOC regions (and the final sectPr) in place. New
    # content is appended into the now-empty body region — immediately before the
    # sectPr — which is exactly the right slot.
    struct = profile.get("structure")
    structure.clear_body_region(doc, struct, preserve_cover=True, preserve_toc=True)

    # Fill the preserved cover anchors in place (never recreate the cover).
    cover.compose_cover(doc, idoc.cover, profile)

    # Write the body blocks in the order given (the body region is freeform).
    resolver = ProfileResolver(profile)
    for block in idoc.blocks:
        _write_block(doc, profile, resolver, block)

    # Refresh the preserved TOC (if any) so Word recomputes it on open — the new
    # headings written into the body will be picked up. No-op when there is no TOC.
    structure.refresh_toc(doc)

    out = Path(output)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out)
    return out


def _write_block(doc, profile: dict, resolver: ProfileResolver, block: ir.Block) -> None:
    if isinstance(block, ir.Heading):
        para = doc.add_paragraph(textutil.runs_to_text(block.runs))
        _apply_resolved_style(doc, para, resolver.resolve_block(block))
    elif isinstance(block, ir.Paragraph):
        para = doc.add_paragraph(textutil.runs_to_text(block.runs))
        _apply_resolved_style(doc, para, resolver.resolve_block(block))
    elif isinstance(block, ir.Callout):
        text = textutil.runs_to_text(block.runs)
        if block.title:
            text = textutil.runs_to_text(block.title) + "\n" + text
        para = doc.add_paragraph(text)
        _apply_resolved_style(doc, para, resolver.resolve_block(block))
    elif isinstance(block, ir.ListBlock):
        op = resolver.resolve_block(block)
        for item in block.items:
            para = doc.add_paragraph(textutil.runs_to_text(item.runs))
            _apply_resolved_style(doc, para, op)
    elif isinstance(block, ir.Table):
        rows = len(block.rows) + (1 if block.columns else 0)
        cols = max(len(block.columns), max((len(r) for r in block.rows), default=0), 1)
        table = doc.add_table(rows=rows, cols=cols)
        op = resolver.resolve_block(block)
        _apply_table_style(doc, table, op)
        row_offset = 0
        if block.columns:
            for idx, run in enumerate(block.columns):
                table.cell(0, idx).text = textutil.runs_to_text([run])
            row_offset = 1
        for r_idx, row in enumerate(block.rows):
            for c_idx, cell in enumerate(row):
                table.cell(r_idx + row_offset, c_idx).text = textutil.runs_to_text(cell.runs)
        if block.caption:
            doc.add_paragraph(textutil.runs_to_text(block.caption))
    elif isinstance(block, ir.Caption):
        para = doc.add_paragraph(textutil.runs_to_text(block.runs))
        _apply_resolved_style(doc, para, resolver.resolve_block(block))
    elif isinstance(block, ir.Quote):
        para = doc.add_paragraph(textutil.runs_to_text(block.runs))
        _apply_resolved_style(doc, para, resolver.resolve_block(block))
        if block.attribution:
            attr = doc.add_paragraph(textutil.runs_to_text(block.attribution))
            _apply_resolved_style(doc, attr, resolver.resolve_role("paragraph"))
    elif isinstance(block, ir.PageBreak):
        doc.add_page_break()
    else:
        para = doc.add_paragraph(getattr(block, "text", ""))
        _apply_resolved_style(doc, para, resolver.resolve_role("paragraph"))


def _apply_resolved_style(doc, para, op) -> None:
    style = _lookup_style(doc, op.resolver)
    if style is not None:
        para.style = style


def _apply_table_style(doc, table, op) -> None:
    style = _lookup_style(doc, op.resolver)
    if style is not None:
        table.style = style


def _lookup_style(doc, resolver: dict):
    style_id = resolver.get("style_id")
    style_name = resolver.get("style_name")
    for style in doc.styles:
        if getattr(style, "style_id", None) == style_id or style.name == style_name:
            return style
    return None
