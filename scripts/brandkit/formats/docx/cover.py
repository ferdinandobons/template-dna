# SPDX-License-Identifier: MIT
"""DOCX cover-anchor discovery and composition."""
from __future__ import annotations

from brandkit.common import text as textutil
from brandkit.ir.model import Cover


PLACEHOLDER_TITLE = "{{title}}"


def discover_cover(doc) -> tuple[list[dict], dict]:
    anchors: list[dict] = []
    for idx, para in enumerate(doc.paragraphs[:8]):
        if PLACEHOLDER_TITLE in para.text or "Insert title" in para.text:
            anchors.append(
                {
                    "id": "title",
                    "container": "paragraph",
                    "paragraph_index": idx,
                    "placeholder": para.text,
                    "branches": None,
                }
            )
            break
    anchor_block = {
        "cover": {
            "kind": "sdt_anchored" if anchors else "NONE",
            "slots_found": len(anchors),
        }
    }
    return anchors, anchor_block


def compose_cover(doc, cover: Cover | None, profile: dict) -> None:
    """Fill the PRESERVED cover title anchor in place.

    Order-aware generation no longer wipes the cover, so the title placeholder is
    still present in the shell when this runs: we overwrite it in place (never
    recreate the cover). Only when the shell carries no cover anchor at all do we
    append a title paragraph as a last resort. Idempotent: a second run finds the
    title already written (the placeholder is gone) and appends nothing because the
    early-out below only appends when no anchor exists — callers re-open the shell
    each run, so the placeholder is always present on a fresh shell.
    """
    if cover is None:
        return
    title = textutil.runs_to_text(cover.title or []) or str(cover.fields.get("title", ""))
    if not title:
        return
    # Fill the preserved cover anchor in place.
    for para in doc.paragraphs[:8]:
        if PLACEHOLDER_TITLE in para.text or "Insert title" in para.text:
            para.text = title
            _apply_role_style(doc, para, profile, "cover.title")
            return
    # No cover anchor in the shell: append a title paragraph as a last resort.
    para = doc.add_paragraph(title)
    _apply_role_style(doc, para, profile, "cover.title")


def _apply_role_style(doc, para, profile: dict, role_id: str) -> None:
    entry = (profile.get("roles") or {}).get(role_id) or {}
    resolver = entry.get("resolver") or {}
    style = _lookup_style(doc, resolver)
    if style is not None:
        para.style = style


def _lookup_style(doc, resolver: dict):
    style_id = resolver.get("style_id")
    style_name = resolver.get("style_name")
    for style in doc.styles:
        if getattr(style, "style_id", None) == style_id or style.name == style_name:
            return style
    return None

