# SPDX-License-Identifier: MIT
from __future__ import annotations

from pathlib import Path

from pptx import Presentation

from brandkit.common import text as textutil
from brandkit.ir import model as ir


def generate(profile: dict, shell_path: str | Path, idoc: ir.IntermediateDocument, output: str | Path) -> Path:
    prs = Presentation(shell_path)
    _clear_existing_slides(prs)
    cover_slide = prs.slides.add_slide(_layout_by_name(prs, "Title Slide") or prs.slide_layouts[0])
    if cover_slide.shapes.title is not None and idoc.cover and idoc.cover.title:
        cover_slide.shapes.title.text = textutil.runs_to_text(idoc.cover.title)
    chunks = _content_chunks(idoc.blocks, _body_capacity(profile))
    for idx, content in enumerate(chunks):
        slide = prs.slides.add_slide(_layout_by_name(prs, "Title and Content") or prs.slide_layouts[1])
        title = _first_heading(idoc.blocks) or "Content"
        if idx:
            title = f"{title} ({idx + 1})"
        if slide.shapes.title is not None:
            slide.shapes.title.text = title
        body = _first_body_placeholder(slide)
        if body is not None:
            body.text = content
    out = Path(output)
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(out)
    return out


def _clear_existing_slides(prs: Presentation) -> None:
    sld_id_lst = prs.slides._sldIdLst
    for sld_id in list(sld_id_lst):
        r_id = sld_id.rId
        prs.part.drop_rel(r_id)
        sld_id_lst.remove(sld_id)


def _layout_by_name(prs: Presentation, name: str):
    for layout in prs.slide_layouts:
        if layout.name == name:
            return layout
    return None


def _first_heading(blocks: list[ir.Block]) -> str | None:
    for block in blocks:
        if isinstance(block, ir.Heading):
            return textutil.runs_to_text(block.runs)
    return None


def _body_text(blocks: list[ir.Block]) -> str:
    lines = []
    for block in blocks:
        if isinstance(block, ir.Heading):
            lines.append(textutil.runs_to_text(block.runs))
        elif isinstance(block, ir.Paragraph):
            lines.append(textutil.runs_to_text(block.runs))
        elif isinstance(block, ir.Callout):
            lines.append(textutil.runs_to_text(block.runs))
        elif isinstance(block, ir.ListBlock):
            lines.extend(textutil.runs_to_text(item.runs) for item in block.items)
    return "\n".join(line for line in lines if line)


def _content_chunks(blocks: list[ir.Block], capacity: int) -> list[str]:
    text = _body_text(blocks)
    if not text:
        return []
    words = text.split()
    chunks: list[str] = []
    cur: list[str] = []
    cur_len = 0
    for word in words:
        add = len(word) + (1 if cur else 0)
        if cur and cur_len + add > capacity:
            chunks.append(" ".join(cur))
            cur = [word]
            cur_len = len(word)
        else:
            cur.append(word)
            cur_len += add
    if cur:
        chunks.append(" ".join(cur))
    return chunks


def _body_capacity(profile: dict) -> int:
    # Conservative L0 estimate: approximately one medium paragraph per slide.
    # M4 will replace this with geometry/font-aware estimator from the profile.
    return int((((profile.get("qa") or {}).get("pptx_text_capacity_chars")) or 1200))


def _first_body_placeholder(slide):
    for shape in slide.placeholders:
        if shape != slide.shapes.title:
            return shape
    return None
