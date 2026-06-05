# SPDX-License-Identifier: MIT
from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path

from openpyxl import load_workbook

from brandkit.common.text import slugify
from brandkit.formats import catalog
from brandkit.ooxml import pack
from brandkit.profile import schema, store


def extract(template: str | Path, name: str, *, scope: str = "project", cwd: str | Path | None = None) -> Path:
    template_path = Path(template)
    wb = load_workbook(template_path, data_only=False)
    named_regions = _defined_names(wb)
    roles = _roles(named_regions)
    surface = {
        "xlsx": {
            "sheets": wb.sheetnames,
            "named_regions": named_regions,
            "named_styles": [style if isinstance(style, str) else style.name for style in wb.named_styles],
            "number_formats": [],
            "table_styles": [],
        }
    }
    profile = schema.build_envelope(
        "xlsx",
        {"name": name, "display_name": name},
        extracted_at=datetime.now(timezone.utc).isoformat(),
        source_template={"filename": template_path.name, "sha256": store.sha256_file(template_path)},
        theme=_theme(),
        roles=roles,
        surface=surface,
    )
    profile["anchors"] = {
        "cover": {"kind": schema.AnchorKind.NAMED_RANGE.value if "title_cell" in named_regions else "NONE", "slots_found": 1 if "title_cell" in named_regions else 0},
        "demo_region": {"present": "data_region" in named_regions},
        "toc": {"present": False},
    }
    profile["provenance"]["ooxml_parts_seen"] = pack.list_parts(template_path)
    profile["artifact_catalog"] = _artifact_catalog(template_path, wb, named_regions, profile["provenance"]["ooxml_parts_seen"])
    profile["capabilities"] = _capabilities()
    target = store.target_dir_for_save(name, scope, cwd=cwd)
    return store.save_profile(target, profile, template_path.read_bytes(), extra_files={"PROFILE.md": _profile_md(profile)}, overwrite=True)


def _defined_names(wb) -> dict:
    out = {}
    for name, defined_name in wb.defined_names.items():
        destinations = list(defined_name.destinations)
        if destinations:
            sheet, coord = destinations[0]
            out[name] = {"sheet": sheet, "range": coord}
    return out


def _roles(named_regions: dict) -> dict:
    roles = {"_index": []}

    def add(rid: str, resolver: dict, signal: str) -> None:
        roles[rid] = {
            "resolver": resolver,
            "appearance": {},
            "verified": True,
            "confidence": 0.85,
            "status": schema.Status.ROBUST.value,
            "evidence": {"signal": signal},
        }
        roles["_index"].append(rid)

    if "title_cell" in named_regions:
        add("title", {"type": schema.ResolverType.NAMED_RANGE.value, "name": "title_cell"}, "named range title_cell")
    for name in named_regions:
        if name != "title_cell":
            add(f"region.{slugify(name).replace('-', '')}", {"type": schema.ResolverType.NAMED_RANGE.value, "name": name}, f"named range {name}")
    if not roles["_index"]:
        add("cell.default", {"type": schema.ResolverType.CELL_STYLE.value, "style_name": "Normal"}, "default style")
    return roles


def _artifact_catalog(path: Path, wb, named_regions: dict, parts: list[str]) -> dict:
    out = catalog.part_catalog(path)
    out["ooxml_parts"] = parts
    out["named_ranges"] = named_regions
    out["named_styles"] = [style if isinstance(style, str) else style.name for style in wb.named_styles]
    out["sheets"] = {}
    formulas = {}
    for ws in wb.worksheets:
        sheet_info = {
            "max_row": ws.max_row,
            "max_column": ws.max_column,
            "freeze_panes": str(ws.freeze_panes) if ws.freeze_panes else None,
            "tables": list(ws.tables.keys()),
            "merged_cells": [str(rng) for rng in ws.merged_cells.ranges],
            "dimensions": {
                "column_widths": {key: dim.width for key, dim in ws.column_dimensions.items() if dim.width},
                "row_heights": {str(key): dim.height for key, dim in ws.row_dimensions.items() if dim.height},
            },
            "non_empty_cells": [],
        }
        # Iterate materialized cells only; large corporate models often have
        # broad dimensions with sparse content.
        for cell in ws._cells.values():
            if cell.value is None:
                continue
            address = f"{ws.title}!{cell.coordinate}"
            value = cell.value
            if isinstance(value, str) and value.startswith("="):
                formulas[address] = value
            sheet_info["non_empty_cells"].append(
                {
                    "address": address,
                    "data_type": cell.data_type,
                    "style": cell.style,
                    "number_format": cell.number_format,
                }
            )
        out["sheets"][ws.title] = sheet_info
    out["formulas"] = formulas
    return out


def _capabilities() -> dict:
    return {
        "extracts_all_ooxml_parts": True,
        "extracts_named_ranges": True,
        "extracts_formula_catalog": True,
        "preserves_formulas_in_shell": True,
        "region_bounds_guard": True,
        "generates_from_shell": True,
    }


def _theme() -> dict:
    return {
        "colors": {},
        "palette_roles": {"primary": {"theme": "accent1"}, "text": {"theme": "dk1"}},
        "fonts": {"major": {"latin": None, "fallback": "Arial"}, "minor": {"latin": None, "fallback": "Calibri"}},
        "embedded_fonts": [],
    }


def _profile_md(profile: dict) -> str:
    return "# Brand Profile: " + profile["identity"]["display_name"] + "\n\n- kind: xlsx\n"
