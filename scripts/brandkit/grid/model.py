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

    @classmethod
    def from_dict(cls, data: dict) -> "GridDocument":
        if not isinstance(data, dict):
            raise ValueError("grid document must be a JSON object")
        return cls(
            cells=dict(data.get("cells") or {}),
            regions=dict(data.get("regions") or {}),
        )


def parse_grid(data: dict) -> GridDocument:
    return GridDocument.from_dict(data)

