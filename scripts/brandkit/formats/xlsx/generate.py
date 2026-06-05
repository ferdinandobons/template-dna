# SPDX-License-Identifier: MIT
"""XLSX generation from a GridDocument.

Routed through the SHARED resolver spine (``ProfileResolver``): every named-cell /
named-region fill is resolved to its ``named_range`` resolver op by the same
kind-aware spine the docx/pptx generators use, so the brand guarantee (only fill
a region the profile PROVED exists, with a resolver type legal for ``xlsx``) is
enforced in one place. Formulas are NEVER authored - the shell's formulas are
preserved verbatim and the workbook is marked for a full recalc on open. When
comprehension is present it steers the cover (fill / clear / leave named regions)
and demo (clear sample-data regions ruled ``verdict=demo``); when absent the
deterministic path fills exactly the named cells/regions the grid names.
"""
from __future__ import annotations

from pathlib import Path
from typing import Optional

from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries

from brandkit.grid.model import GridDocument
from brandkit.profile import schema, store
from brandkit.profile.resolver import ProfileResolver
from brandkit.qa.checks_deterministic import check_no_net_structure_loss
from brandkit.qa.model import Finding


def generate(
    profile: dict,
    shell_path: str | Path,
    grid: GridDocument,
    output: str | Path,
    *,
    findings: Optional[list[Finding]] = None,
) -> Path:
    """Generate an XLSX from ``grid`` onto the brand ``shell_path``.

    ``findings`` (optional out-param) is appended with any degradation findings so
    the QA gate can surface them. Off-brand / fabricated fills stay impossible by
    construction: every fill target is resolved through :class:`ProfileResolver`
    against ``profile['roles']``, so a grid key that does not name a region the
    profile proved exists is rejected (never invented).
    """
    sink: list[Finding] = findings if findings is not None else []
    wb = load_workbook(shell_path, data_only=False)
    resolver = ProfileResolver(profile)

    # The named-range geometry the resolver's ``named_range`` ops point at. This is
    # the profile's OWN surfaced map (the author's range names), not a code literal.
    regions = ((profile.get("surface") or {}).get("xlsx") or {}).get("named_regions") or {}

    # Reverse index: named-range NAME -> the role id whose resolver targets it. The
    # grid is keyed by the author's range names (its vocabulary); we resolve each
    # through the spine so the brand guarantee gates every fill.
    name_to_role = _name_to_role(profile)

    # Fill single named cells.
    for name, value in grid.cells.items():
        target = _resolve_named_target(resolver, name_to_role, regions, name, sink)
        sheet, coord = target["sheet"], target["range"]
        min_col, min_row, _, _ = range_boundaries(coord)
        wb[sheet].cell(row=min_row, column=min_col).value = value

    # Fill multi-cell named regions (bounds-guarded; never overruns the range).
    for name, values in grid.regions.items():
        target = _resolve_named_target(resolver, name_to_role, regions, name, sink)
        sheet, coord = target["sheet"], target["range"]
        min_col, min_row, max_col, max_row = range_boundaries(coord)
        _check_region_bounds(name, values, max_rows=max_row - min_row + 1, max_cols=max_col - min_col + 1)
        ws = wb[sheet]
        for r_idx, row in enumerate(values):
            for c_idx, value in enumerate(row):
                ws.cell(row=min_row + r_idx, column=min_col + c_idx).value = value

    # Comprehension-steered reconciliation (no-op when comprehension is absent):
    # CLEAR cover anchors / demo sample-data regions the model ruled destructive,
    # with the destructive floor corroborating each removal.
    removed_refs = _reconcile_cover_and_demo(wb, profile, grid, regions, sink)

    # Recalc: mark the workbook for a full recompute on open so preserved formulas
    # pick up the new inputs. The analogue of the docx TOC refresh; formulas are
    # NEVER authored here, only the shell's own formulas are recalculated.
    try:
        wb.calculation.fullCalcOnLoad = True
    except Exception:
        pass

    # Destructive-action floor (plan §6): every region the reconciliation cleared
    # must carry a corroborated destructive verdict, else ERROR. Model-free.
    if store.comprehension_is_present(profile):
        sink.extend(check_no_net_structure_loss(removed_refs, profile))

    out = Path(output)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return out


def _name_to_role(profile: dict) -> dict[str, str]:
    """Map each named-range NAME to the role id whose resolver targets it.

    Built from ``profile['roles']`` so a grid key is resolved through the spine
    (the brand guarantee) rather than read straight off ``surface``. A role whose
    resolver is not a ``named_range`` (e.g. the default ``cell_style`` fallback) is
    skipped - it names no range.
    """
    out: dict[str, str] = {}
    for rid, entry in (profile.get("roles") or {}).items():
        if rid == "_index" or not isinstance(entry, dict):
            continue
        resolver = entry.get("resolver") or {}
        if resolver.get("type") != schema.ResolverType.NAMED_RANGE.value:
            continue
        name = resolver.get("name")
        if isinstance(name, str) and name:
            out[name] = rid
    return out


def _resolve_named_target(resolver, name_to_role, regions, name, findings) -> dict:
    """Resolve a grid key (a named-range name) to its ``{sheet, range}`` target.

    Routes through the shared spine: the name maps to a role id, the resolver op
    must be a legal ``named_range`` for the profile kind, and its target range must
    be one the profile surfaced. A key naming no proven region is an ERROR (the
    fill is never invented).
    """
    rid = name_to_role.get(name)
    if rid is not None:
        op = resolver.resolve_role(rid, fallback=None)
        if op.resolver.get("type") == schema.ResolverType.NAMED_RANGE.value:
            resolved_name = op.resolver.get("name", name)
            target = regions.get(resolved_name)
            if target:
                return target
    # Fail closed: a grid key that resolves to no surfaced named range is a brand
    # breach (a fabricated target), not a silent no-op.
    raise ValueError(f"unknown named cell/range {name!r}")


def _reconcile_cover_and_demo(wb, profile: dict, grid: GridDocument, regions: dict, findings: list[Finding]) -> set[str]:
    """Reconcile preserved cover anchors + demo sample-data regions with new content.

    Comprehension-steered; a no-op when comprehension is absent (the deterministic
    fills above already wrote the named regions the grid named). For each
    ``cover_slots`` entry ruled ``fill_rule=clear`` and each ``demo_classification``
    region ruled ``verdict=demo`` that maps to a surfaced named range, the cells are
    CLEARED in place. Returns the set of refs actually cleared (anchor_ref /
    region_ref) for the destructive floor. Formulas are never cleared - only the
    region's own input cells.
    """
    if not store.comprehension_is_present(profile):
        return set()
    comp = profile.get("comprehension")
    if not isinstance(comp, dict):
        return set()

    cover_anchors = ((profile.get("surface") or {}).get("xlsx") or {}).get("cover_anchors") or []
    region_inv = ((profile.get("surface") or {}).get("xlsx") or {}).get("regions") or []
    anchor_to_name = {a.get("id"): a.get("name") for a in cover_anchors if isinstance(a, dict)}
    region_to_name = {r.get("id"): r.get("name") for r in region_inv if isinstance(r, dict)}

    removed: set[str] = set()

    # Cover anchors ruled CLEAR.
    for anchor_ref, slot in (comp.get("cover_slots") or {}).items():
        if not isinstance(slot, dict) or slot.get("fill_rule") != schema.FillRule.CLEAR.value:
            continue
        name = anchor_to_name.get(anchor_ref)
        target = regions.get(name) if name else None
        if not target:
            continue
        if _clear_region(wb, target):
            removed.add(anchor_ref)

    # Demo sample-data regions ruled DEMO -> clear in place (the demo rows go away;
    # the new content is written into the same named region by the grid above).
    for reg in (comp.get("demo_classification") or {}).get("regions") or []:
        if not isinstance(reg, dict) or reg.get("verdict") != schema.Verdict.DEMO.value:
            continue
        region_ref = reg.get("region_ref")
        name = region_to_name.get(region_ref)
        target = regions.get(name) if name else None
        if not target:
            continue
        # If the grid already refilled this region, the demo rows are overwritten;
        # only clear the trailing cells the new content did NOT cover so no stale
        # demo value survives. Still counts as a destructive act for the floor.
        if _clear_region(wb, target, skip_rows=len(grid.regions.get(name) or [])):
            removed.add(region_ref)

    return removed


def _clear_region(wb, target: dict, *, skip_rows: int = 0) -> bool:
    """Empty a named region's NON-FORMULA cells in place. Returns True if it ran.

    A cell holding a formula (``=...``) is never cleared (formulas are load-bearing
    shell content); only literal demo/placeholder cells are emptied. ``skip_rows``
    leaves the first N region rows intact (already refilled by the grid).
    """
    sheet, coord = target.get("sheet"), target.get("range")
    if not sheet or not coord or sheet not in wb.sheetnames:
        return False
    try:
        min_col, min_row, max_col, max_row = range_boundaries(coord)
    except Exception:
        return False
    ws = wb[sheet]
    for r in range(min_row + skip_rows, max_row + 1):
        for c in range(min_col, max_col + 1):
            cell = ws.cell(row=r, column=c)
            if isinstance(cell.value, str) and cell.value.startswith("="):
                continue  # preserve formulas
            cell.value = None
    return True


def _check_region_bounds(name: str, values: list[list], *, max_rows: int, max_cols: int) -> None:
    if len(values) > max_rows:
        raise ValueError(f"region {name!r} has {len(values)} rows; named range allows {max_rows}")
    for idx, row in enumerate(values, start=1):
        if len(row) > max_cols:
            raise ValueError(f"region {name!r} row {idx} has {len(row)} columns; named range allows {max_cols}")
