# SPDX-License-Identifier: MIT
"""Format-neutral artifact catalog helpers.

The Brand Profile's roles describe what the generator can resolve today. The
artifact catalog is broader: it records every reusable thing the extractor can
observe so an AI agent can reason about the template, even when generation for a
given artifact is still staged.
"""
from __future__ import annotations

from collections import Counter
from pathlib import Path

from brandkit.ooxml import pack


def part_catalog(path: str | Path) -> dict:
    parts = pack.list_parts(path)
    kinds = Counter(_part_kind(part) for part in parts)
    return {
        "ooxml_parts": parts,
        "part_counts": dict(sorted(kinds.items())),
        "media_parts": [p for p in parts if _part_kind(p) == "media"],
        "theme_parts": [p for p in parts if "theme" in p.lower()],
        "relationship_parts": [p for p in parts if p.endswith(".rels")],
    }


def _part_kind(part: str) -> str:
    lower = part.lower()
    if lower.endswith(".rels"):
        return "relationships"
    if "/media/" in lower or lower.endswith((".png", ".jpg", ".jpeg", ".gif", ".bmp", ".svg")):
        return "media"
    if "theme" in lower:
        return "theme"
    if lower.endswith(".xml"):
        return "xml"
    return "binary"

