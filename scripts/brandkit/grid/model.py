# SPDX-License-Identifier: MIT
"""GridDocument for XLSX generation.

M1/M2 grid input is intentionally thin: named cells and named regions are filled
into the workbook shell. Formulas and workbook topology remain in the shell.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class GridDocument:
    cells: dict[str, Any] = field(default_factory=dict)
    regions: dict[str, list[list[Any]]] = field(default_factory=dict)
    # Optional per-target NUMBER-FORMAT intent, keyed by the SAME name as ``cells`` /
    # ``regions`` (the author's brand vocabulary). The value is a brand-agnostic
    # semantic family (``currency`` | ``percent`` | ``date`` | ...); generation
    # resolves ``number.<family>`` through the profile to the template's own mask
    # and applies it to the filled cell(s). An intent the profile does not carry
    # degrades loudly (the cell keeps its existing format) - a format is never
    # fabricated. ``{"price": "currency", "share": "percent"}``.
    formats: dict[str, str] = field(default_factory=dict)
    # Native charts to author over the workbook's OWN cell data. Each spec is a
    # plain dict describing a chart that REFERENCES existing cell ranges (the xlsx
    # peer of the docx/pptx inline-data chart - here the data lives in the sheet, the
    # chart's strength): ``{"sheet"?, "type", "title"?, "anchor", "data",
    # "categories"?, "data_titles"?}``. ``data``/``categories`` are A1 ranges on
    # ``sheet`` (default: the active sheet); ``anchor`` is the top-left cell.
    charts: list[dict] = field(default_factory=list)

    @classmethod
    def from_dict(cls, data: dict) -> "GridDocument":
        if not isinstance(data, dict):
            raise ValueError("grid document must be a JSON object")
        return cls(
            cells=dict(data.get("cells") or {}),
            regions=dict(data.get("regions") or {}),
            charts=[dict(c) for c in (data.get("charts") or []) if isinstance(c, dict)],
            formats={
                str(k): str(v)
                for k, v in (data.get("formats") or {}).items()
                if isinstance(v, str) and v
            },
        )


def parse_grid(data: dict) -> GridDocument:
    return GridDocument.from_dict(data)
