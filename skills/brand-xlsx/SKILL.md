---
name: brand-xlsx
description: Brand-aware Excel engine. Use to extract reusable named ranges and workbook shell structure from a .xlsx template, verify it, and generate a new on-brand .xlsx from a GridDocument fill manifest.
---

# brand-xlsx

Use this skill when the user wants reusable branded Excel/workbook generation
from a company `.xlsx` template and variable user-provided data.

This is an AI-agent skill for Codex and Claude Code. The user should describe
the workbook/model they want filled; the agent maps that request to named cells
and named regions, invokes the internal Python engine, verifies the output, and
returns the generated `.xlsx`.

## Agent Workflow

1. Determine the brand name and locate the user-provided `.xlsx` template.
2. If no matching `brand-kit/<name>` exists, extract one.
3. Convert the user's tabular/model data into `GridDocument` JSON.
4. Generate the `.xlsx` with the internal engine.
5. Run QA and report any warnings honestly.

Before generation, inspect `profile.json.artifact_catalog` when the user asks
to mimic a specific workbook piece. It records OOXML parts, named ranges,
formulas, sheet dimensions, table names, merged cells, row/column sizing, cell
styles, and number formats.

## Internal Extract

```bash
python scripts/brandkit/cli.py extract --name <brand> --template <template.xlsx> --scope project
```

## Internal Verify

```bash
python scripts/brandkit/cli.py verify --name <brand> --scope auto --qa auto
```

## Internal Generate

```bash
python scripts/brandkit/cli.py generate --name <brand> --input <grid-document.json> --output <output.xlsx> --scope auto --qa auto
```

M2 fills named cells and named regions while preserving formulas and workbook
topology in the shell. Region fills that exceed the named range are refused
before saving.
