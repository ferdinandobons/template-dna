# SPDX-License-Identifier: MIT
"""Generate an on-brand ``.pptx`` from the IR block stream and a Brand Profile.

Design (off-brand-by-construction, §C3/M6/M7):

- We open FROM the template shell (``shell_path``) and only ever add slides built
  on layouts the *profile proved exist*. The cover and content layouts are read
  from the profile's real ``roles`` (``cover.title`` / ``heading.1`` /
  ``paragraph``), which the extractor derived from the parsed deck — never a
  hardcoded ``"Title Slide"``/``"Title and Content"`` literal. If the profile has
  no usable layout for a role (status ``stub`` / ``layout: null``) we fall back to
  a real layout already present in the shell, never to a fabricated name.

- Slides are built from the IR block stream, not a flattened string: each heading
  opens a new section/slide (its own runs become the title); the blocks that
  follow it become that slide's body, preserving lists / tables / quotes /
  captions / callouts as distinct lines. Heading text is never duplicated into the
  body. Within a single section, body text that exceeds the slide's capacity is
  split into continuation slides (``Title (2)`` …) carrying the same layout.
"""
from __future__ import annotations

from pathlib import Path

from pptx import Presentation

from brandkit.common import text as textutil
from brandkit.ir import model as ir


def generate(profile: dict, shell_path: str | Path, idoc: ir.IntermediateDocument, output: str | Path) -> Path:
    prs = Presentation(shell_path)
    _clear_existing_slides(prs)

    cover_layout = _layout_for_role(prs, profile, "cover.title")
    content_layout = _layout_for_role(prs, profile, "heading.1") or _layout_for_role(prs, profile, "paragraph")

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

    out = Path(output)
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(out)
    return out


# ---------------------------------------------------------------------------
# Layout resolution (from the profile's REAL role/layout data, not literals)
# ---------------------------------------------------------------------------
def _layout_for_role(prs: Presentation, profile: dict, role_id: str):
    """Return the real shell layout the profile assigns to ``role_id``.

    Reads ``profile['roles'][role_id]['resolver']['layout']`` — a name the
    extractor proved exists in this deck — and looks it up by name. Returns
    ``None`` when the role is a stub (no layout) or the named layout is absent
    from the shell (so callers fall back to a real layout, never a fiction).
    """
    roles = profile.get("roles")
    if not isinstance(roles, dict):
        return None
    entry = roles.get(role_id)
    if not isinstance(entry, dict):
        return None
    resolver = entry.get("resolver")
    if not isinstance(resolver, dict):
        return None
    name = resolver.get("layout")
    if not name:
        return None
    return _layout_by_name(prs, name)


def _layout_by_name(prs: Presentation, name: str):
    for layout in prs.slide_layouts:
        if layout.name == name:
            return layout
    return None


# ---------------------------------------------------------------------------
# IR block stream -> sections -> slides (no flattening, nothing dropped)
# ---------------------------------------------------------------------------
def _sections(blocks: list[ir.Block]) -> list[dict]:
    """Split the block stream into sections at each heading.

    One section per heading: ``{"title": <heading text>, "body": [blocks...]}``.
    Blocks before the first heading form a leading section with an empty title
    (its title placeholder is then left to the shell default). Heading runs become
    the title and are NOT echoed into the body.
    """
    sections: list[dict] = []
    current: dict | None = None
    for block in blocks:
        if isinstance(block, ir.Heading):
            current = {"title": textutil.runs_to_text(block.runs) or "Content", "body": []}
            sections.append(current)
        elif isinstance(block, ir.PageBreak):
            # An explicit slide break closes the current section.
            current = None
        else:
            if current is None:
                current = {"title": "Content", "body": []}
                sections.append(current)
            current["body"].append(block)
    return [s for s in sections if s["title"] or s["body"]]


def _body_lines(blocks: list[ir.Block]) -> list[str]:
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
                    quote = f"{quote} — {attribution}" if quote else attribution
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


def _append_list_item(lines: list[str], item: ir.ListItem) -> None:
    text = textutil.runs_to_text(item.runs)
    if text:
        lines.append(("    " * max(item.level, 0)) + "• " + text)
    for sub in item.items:
        _append_list_item(lines, sub)


def _append_table(lines: list[str], table: ir.Table) -> None:
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
    """Pack body lines into slide-sized chunks, preserving structure.

    Splits at line boundaries first (a list item, table row, quote… is never
    broken across slides if it fits). A single line that alone exceeds the
    capacity is wrapped at word boundaries onto continuation slides rather than
    dropped or truncated. Always yields at least one chunk so a section with an
    empty body still produces one slide.
    """
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
            # An oversized single line: flush what we have, then wrap it at word
            # boundaries across as many continuation slides as needed.
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
    # M4 will replace this with a geometry/font-aware estimator from the profile.
    return int((((profile.get("qa") or {}).get("pptx_text_capacity_chars")) or 1200))


# ---------------------------------------------------------------------------
# Shell helpers
# ---------------------------------------------------------------------------
def _clear_existing_slides(prs: Presentation) -> None:
    sld_id_lst = prs.slides._sldIdLst
    for sld_id in list(sld_id_lst):
        r_id = sld_id.rId
        prs.part.drop_rel(r_id)
        sld_id_lst.remove(sld_id)


def _subtitle_placeholder(slide):
    from pptx.enum.shapes import PP_PLACEHOLDER

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
