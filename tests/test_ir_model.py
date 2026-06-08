# SPDX-License-Identifier: MIT
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from brandkit.common import text as textutil
from brandkit.ir import model as ir


def test_list_items_accept_plain_string_shortcut() -> None:
    doc = ir.parse_idoc(
        {
            "blocks": [
                {
                    "type": "list",
                    "items": [
                        "Plain bullet",
                        {"text": "Nested parent", "items": ["Nested child"]},
                    ],
                }
            ]
        }
    )

    block = doc.blocks[0]
    assert isinstance(block, ir.ListBlock)
    assert textutil.runs_to_text(block.items[0].runs) == "Plain bullet"
    assert block.items[0].level == 0
    assert textutil.runs_to_text(block.items[1].items[0].runs) == "Nested child"
    assert block.items[1].items[0].level == 1


def test_table_multi_run_column_header_survives_round_trip() -> None:
    # A multi-run column header (plain + bold unit) must keep BOTH runs through
    # from_dict (it used to collapse to the first run, silently dropping the rest).
    doc = ir.parse_idoc(
        {
            "blocks": [
                {
                    "type": "table",
                    "columns": [
                        {"runs": [{"t": "Sales "}, {"t": "(M)", "b": True}]},
                        "Region",
                    ],
                    "rows": [["100", "North"]],
                }
            ]
        }
    )
    table = doc.blocks[0]
    assert isinstance(table, ir.Table)
    # Column 0 keeps both runs (bold preserved); column 1 normalizes the string.
    assert textutil.runs_to_text(table.columns[0]) == "Sales (M)"
    assert table.columns[0][1].get("b") is True
    assert textutil.runs_to_text(table.columns[1]) == "Region"
    # to_dict -> from_dict is content-stable (no run lost on the second trip).
    again = ir.Table.from_dict(table.to_dict())
    assert textutil.runs_to_text(again.columns[0]) == "Sales (M)"
    assert again.columns[0][1].get("b") is True


def test_deeply_nested_list_raises_iidparse_error_not_recursion():
    # A pathologically deep ``items`` chain must raise the contracted IIDParseError,
    # never an uncaught RecursionError (fail-closed on hostile/runaway input).
    node = {"text": "leaf"}
    for _ in range(200):
        node = {"text": "n", "items": [node]}
    raised = None
    try:
        ir.ListItem.from_dict(node)
    except ir.IIDParseError:
        raised = "iidparse"
    except RecursionError:
        raised = "recursion"
    assert raised == "iidparse", f"expected IIDParseError, got {raised}"


def test_modest_list_nesting_still_parses():
    node = {"text": "leaf"}
    for _ in range(10):
        node = {"text": "n", "items": [node]}
    item = ir.ListItem.from_dict(node)
    assert item.items  # parsed fine within the cap
