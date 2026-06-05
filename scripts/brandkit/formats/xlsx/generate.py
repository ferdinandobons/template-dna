# SPDX-License-Identifier: MIT
from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries

from brandkit.grid.model import GridDocument


def generate(profile: dict, shell_path: str | Path, grid: GridDocument, output: str | Path) -> Path:
    wb = load_workbook(shell_path, data_only=False)
    regions = ((profile.get("surface") or {}).get("xlsx") or {}).get("named_regions") or {}
    for name, value in grid.cells.items():
        target = regions.get(name)
        if not target:
            raise ValueError(f"unknown named cell/range {name!r}")
        sheet, coord = target["sheet"], target["range"]
        min_col, min_row, _, _ = range_boundaries(coord)
        wb[sheet].cell(row=min_row, column=min_col).value = value
    for name, values in grid.regions.items():
        target = regions.get(name)
        if not target:
            raise ValueError(f"unknown named region {name!r}")
        sheet, coord = target["sheet"], target["range"]
        min_col, min_row, max_col, max_row = range_boundaries(coord)
        _check_region_bounds(name, values, max_rows=max_row - min_row + 1, max_cols=max_col - min_col + 1)
        ws = wb[sheet]
        for r_idx, row in enumerate(values):
            for c_idx, value in enumerate(row):
                ws.cell(row=min_row + r_idx, column=min_col + c_idx).value = value
    out = Path(output)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return out


def _check_region_bounds(name: str, values: list[list], *, max_rows: int, max_cols: int) -> None:
    if len(values) > max_rows:
        raise ValueError(f"region {name!r} has {len(values)} rows; named range allows {max_rows}")
    for idx, row in enumerate(values, start=1):
        if len(row) > max_cols:
            raise ValueError(f"region {name!r} row {idx} has {len(row)} columns; named range allows {max_cols}")
