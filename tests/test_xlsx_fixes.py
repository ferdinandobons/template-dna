# SPDX-License-Identifier: MIT
"""Regression tests for XLSX correctness fixes.

Covers the confirmed finding from the strategic review:

  - MERGED-CELL CRASH: filling or clearing a named region whose first row straddles
    a merged banner used to write ``.value`` onto a merge SLAVE cell, which openpyxl
    makes read-only -> ``AttributeError``, so ``generate`` crashed with no document
    produced. The extractor surfaces ``merged_header`` as valid region evidence, so
    the engine invited exactly the input it crashed on (merged banners/titles are
    near-universal in corporate financial templates). The fill/clear choke points
    now skip merge slaves; a dropped NON-``None`` fill is surfaced as a
    ``block_degraded`` WARNING (the merge anchor keeps its value) so the skip is
    visible in QA rather than a silent loss.

All shells are synthesized in a temp dir with openpyxl and never committed; the
range names used as ids are the synthetic author's OWN vocabulary carried as data.
"""

from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.workbook.defined_name import DefinedName

from brandkit.formats.xlsx import extract as xlsx_extract
from brandkit.formats.xlsx import generate as xlsx_generate
from brandkit.grid.model import GridDocument
from brandkit.profile import store
from brandkit.qa.model import Finding


class FillClearChokePointMergedCellGuard(unittest.TestCase):
    """Unit-level guard at the exact write choke points (no profile machinery)."""

    def _merged_ws(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "DEMO"
        ws.merge_cells("A1:C1")  # B1, C1 become read-only MergedCell slaves
        return wb, ws

    def test_fill_cell_anchor_is_written(self) -> None:
        _wb, ws = self._merged_ws()
        sink: list[Finding] = []
        xlsx_generate._fill_cell(ws["A1"], "NEW", sink=sink, where="banner")
        self.assertEqual(ws["A1"].value, "NEW")
        self.assertEqual(sink, [])  # writing the anchor is normal, no finding

    def test_fill_cell_merged_slave_is_skipped_and_surfaced(self) -> None:
        _wb, ws = self._merged_ws()
        slave = ws["B1"]
        self.assertIsInstance(slave, MergedCell)  # precondition: it IS a slave
        sink: list[Finding] = []
        # Before the fix this raised AttributeError ('... is read-only').
        xlsx_generate._fill_cell(slave, "x", sink=sink, where="banner")
        self.assertEqual(len(sink), 1)
        self.assertEqual(sink[0].check, "block_degraded")
        self.assertEqual(sink[0].severity, "WARNING")
        self.assertIn("B1", sink[0].message)

    def test_fill_cell_none_value_on_slave_is_silent(self) -> None:
        # A sparse / ragged grid cell (None) over a slave is a no-op with no noise.
        _wb, ws = self._merged_ws()
        sink: list[Finding] = []
        xlsx_generate._fill_cell(ws["B1"], None, sink=sink, where="banner")
        self.assertEqual(sink, [])

    def test_clear_region_skips_merged_slaves_without_crashing(self) -> None:
        wb, ws = self._merged_ws()
        target = {"sheet": ws.title, "range": "A1:C1"}
        # Before the fix this raised AttributeError on the B1/C1 slaves.
        ran = xlsx_generate._clear_region(wb, target)
        self.assertTrue(ran)

    def test_reassert_cover_style_skips_merged_slave(self) -> None:
        _wb, ws = self._merged_ws()
        # Setting .style on a MergedCell raises; the guard must skip it (no crash).
        xlsx_generate._reassert_cover_style(ws["B1"], "Normal")


class MergedBannerRegionEndToEnd(unittest.TestCase):
    """End-to-end: a named region straddling a merged banner generates cleanly."""

    def _build_shell(self, td: Path) -> Path:
        wb = Workbook()
        ws = wb.active
        ws.title = "S"
        ws["A1"] = "DEMO BANNER"
        ws.merge_cells("A1:C1")  # the named region's first row straddles this merge
        ws["A3"] = "label"
        wb.defined_names.add(DefinedName("banner", attr_text="S!$A$1:$C$1"))
        wb.defined_names.add(DefinedName("label_cell", attr_text="S!$A$3"))
        shell = td / "shell.xlsx"
        wb.save(shell)
        return shell

    def _extract(self, td: Path, shell: Path):
        old = os.getcwd()
        os.chdir(td)
        try:
            xlsx_extract.extract(shell, "syn", scope="project", cwd=td)
            return store.load_profile("syn", "project")
        finally:
            os.chdir(old)

    def test_fill_region_over_merged_banner_does_not_crash(self) -> None:
        with tempfile.TemporaryDirectory() as t:
            td = Path(t)
            shell = self._build_shell(td)
            loaded = self._extract(td, shell)
            grid = GridDocument(
                regions={"banner": [["NEW BANNER", "x", "y"]]},
                cells={"label_cell": "Hello"},
            )
            out = td / "out.xlsx"
            sink: list[Finding] = []
            # Before the fix: AttributeError on the B1/C1 slaves, no file produced.
            xlsx_generate.generate(
                loaded.profile, loaded.shell_path, grid, out, findings=sink
            )
            self.assertTrue(out.is_file())

            wb = load_workbook(out)
            # The merge ANCHOR took the new value; the unrelated cell filled normally.
            self.assertEqual(wb["S"]["A1"].value, "NEW BANNER")
            self.assertEqual(wb["S"]["A3"].value, "Hello")
            # The dropped slave writes are surfaced, never silent.
            degraded = [f for f in sink if f.check == "block_degraded"]
            self.assertTrue(degraded)
            self.assertTrue(all(f.severity == "WARNING" for f in degraded))


if __name__ == "__main__":  # pragma: no cover
    unittest.main()
