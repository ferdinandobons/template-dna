---
name: brand-xlsx
description: >-
  Brand-aware Excel engine. Use to (1) EXTRACT a company's brand from a .xlsx template into a
  reusable "Brand Profile", (2) COMPREHEND the template with the model (optional), (3) VERIFY it,
  (4) GENERATE a new on-brand .xlsx from a GridDocument fill manifest. Trigger on "extract our
  brand", "use our workbook template", "generate a branded workbook from our profile", or when a
  ./brand-kit exists. For one-off spreadsheet edits with no saved brand profile, use the normal
  xlsx skill instead. NOT for .docx (brand-docx), .pptx (brand-pptx), or PDFs.
---

# brand-xlsx

Use this skill when the user wants reusable branded Excel/workbook generation
from a company `.xlsx` template and variable user-provided data.

This is an AI-agent skill for Codex and Claude Code. The user should describe
the workbook/model they want filled; the agent maps that request to named cells
and named regions, invokes the internal Python engine, verifies the output, and
returns the generated `.xlsx`.

## The four verbs

Every brand skill (`brand-docx`, `brand-pptx`, `brand-xlsx`) implements the same
contract: **extract / comprehend / verify / generate**.

| Verb | Input | Output |
|---|---|---|
| **extract** | a company `.xlsx` template | a reusable Brand Profile |
| **comprehend** *(optional, model-driven)* | a saved profile + a model-authored `comprehension.json` | the profile with a validated, cached `comprehension` block |
| **verify** | a saved Brand Profile | QA findings + a verdict |
| **generate** | data (a GridDocument) + a profile | a new on-brand `.xlsx` |

`comprehend` is **optional**: `generate` works on the deterministic profile alone.
See [reference/comprehension.md](reference/comprehension.md) for the full step.

## Hard Rules

- Treat `python scripts/brandkit/cli.py ...` as an internal engine command, not the user-facing workflow.
- Run the dependency preflight before starting extract / comprehend / verify / generate, and report missing or unusable dependencies before proceeding.
- Extract opens the source template read-only and saves `brand-kit/<name>/template/shell.xlsx` byte-for-byte.
- Generate opens the saved shell and resolves every named cell/region through `profile.json`.
- Do not put style names, colors, fonts, or brand identifiers in a GridDocument.
- If the user did not provide a template or enough data, ask for the missing input.
- Return the generated file path plus a QA summary.
- Consult `profile.json.artifact_catalog` before generation when the user asks to mimic a specific piece of the template.

## Preflight (always first)

Before doing any work, run:

```bash
python scripts/brandkit/cli.py doctor
```

Use its output to decide the run mode:

- If a required Python dependency is missing, install/repair it before extraction
  or generation; the core engine is not ready.
- If only visual renderers are missing or unusable (`soffice` plus `pdftoppm` or
  optional PyMuPDF/`fitz`), the
  core L0 workflow can still run, but a full visual audit cannot be claimed.
  Tell the user what is missing, include the install/repair hint printed by
  `doctor`, and either proceed with degraded QA or install the renderer first.
- If optional OCR (`tesseract`) is missing, the visual audit can still run, but
  rendered residual-text proof is incomplete. Report that limitation when
  judging stale placeholders or field caches.
- For `--qa deep` or `--qa strict`, prefer repairing/installing renderers before
  generation. If the environment cannot run them, `deep` generates a degraded
  manifest and `strict` fails with a visual proof blocker.

## Agent Workflow

1. Run the dependency preflight above and report any degraded capability.
2. Determine the brand name and locate the user-provided `.xlsx` template.
3. If no matching `brand-kit/<name>` exists, **extract** one.
4. **Comprehend** the template (optional, model-driven; see below). Skip when a
   current comprehension is already cached or no model is available.
5. Convert the user's tabular/model data into `GridDocument` JSON.
6. **Generate** the `.xlsx` with the internal engine.
7. Run **QA** and report any warnings honestly.
8. **Feedback** (only after returning the file): invite a refinement of the
   understanding for future workbooks (see below).

Before generation, inspect `profile.json.artifact_catalog` when the user asks
to mimic a specific workbook piece. It records OOXML parts, named ranges,
formulas, sheet dimensions, table names, merged cells, row/column sizing, cell
styles, and number formats.

## Feedback (end of generation)

Ask for feedback **only after** you have returned the generated `.xlsx` and its
QA summary - never before or during generation. Invite the user to reply with
**text or a screenshot** of the workbook, and name the roles, palette colors, and
sections you actually used so the answer is concrete. A screenshot is your own
multimodal read; the engine only ever ingests the structured JSON delta you
distil from it.

Turn the answer into a small refinement delta of verbatim ids and merge it with
the `refine` verb (see [reference/comprehension.md](reference/comprehension.md)):

```bash
python scripts/brandkit/cli.py refine --name <brand> --input refinement.json --accept
```

A refinement improves **FUTURE** generations of this brand only - it mutates the
saved profile, never the `.xlsx` you just produced. To apply it, generate again.

## Internal Extract

```bash
python scripts/brandkit/cli.py extract --name <brand> --template <template.xlsx> --scope project
```

## Internal Comprehend (optional, model-driven)

Read [reference/comprehension.md](reference/comprehension.md) for the full
guidance, the five questions, and the anti-overfitting directive. In short:

```bash
python scripts/brandkit/cli.py comprehend-input --name <brand>   # prints {facts, excerpt} for the model
python scripts/brandkit/cli.py comprehend --name <brand> --input comprehension.json  # the ONLY writer
```

Skip this verb when `comprehension.status` is `present` **and** its
`source_shell_sha256` equals the live `provenance.shell.sha256`. Never re-run it
at generate time.

> **xlsx readiness.** The Excel extractor surfaces named-region cover anchors and
> sample-data regions, while `fields` is intentionally empty because workbooks do
> not have a TOC-style derived index. A current comprehension can therefore steer
> cover fills/clears and demo-region cleanup, but it must not invent an index ref;
> a ref into the empty field inventory is fail-closed and will be rejected.

## Internal Verify

```bash
python scripts/brandkit/cli.py verify --name <brand> --scope auto --qa auto
```

`--qa` selects the QA depth (see [reference/visual-audit.md](reference/visual-audit.md)):

- `fast`: deterministic **L0** only.
- `auto`: L0 **+ L1** visual pixel proxies when renderers (`soffice` plus `pdftoppm` or optional PyMuPDF/`fitz`) are present; otherwise L0 plus a single INFO `visual.unavailable`.
- `deep`: L0 + L1 **+ a `visual_manifest.json`** and per-page PNGs; if `tesseract` is installed the manifest also includes OCR text/hits. The orchestrator must then run the **L2** step (see below).
- `strict`: deep visual audit plus gate errors when full render proof is unavailable or L1/OCR evidence is not clean.

Verify has no output to render, so all modes behave as L0 at verify time; the visual stages run at **generate** time.

## Internal Generate

```bash
python scripts/brandkit/cli.py generate --name <brand> --input <grid-document.json> --output <output.xlsx> --scope auto --qa auto
```

See `reference/comprehension.md` and `reference/visual-audit.md`.

## Visual audit (two-stage)

The engine renders the output and runs deterministic pixel proxies, but the
**qualitative visual judgement is yours (the orchestrator), never the engine's** -
the Python engine never calls a model. To run the full two-stage audit:

1. Generate with `--qa deep`. The engine renders each printed page to a PNG, runs
   the L1 proxies, and writes `visual_manifest.json` next to the output in an
   `<output-file>.visual/` dir, such as `workbook.xlsx.visual/` (a side artifact;
   the `.xlsx` bytes never change).
2. Read the manifest path from stdout (`visual manifest: <path>`).
3. Open the PNGs listed in `pages[*].png`. For every entry in `checklist`, judge
   PASS/FAIL against the rendered pages, taking `l1_findings` and `ocr.hits` into
   account.
4. If any checklist item FAILS (or an L1 WARNING is confirmed visually as a real
   defect, or a `visual.ocr_residual_text` hit is confirmed as stale visible
   template text): **repair** the grid/content or the generated composition,
   **regenerate**, then **re-run the audit**. Loop until the checklist is clean,
   or until no further targeted repair can be justified without user input.

L1 findings are WARNING-only and never fail the gate by themselves; the real
qualitative gate is your L2 judgement.

During repair, treat the template as a source of reusable workbook affordances,
not a rule to preserve blindly. If inherited print areas, hidden rows/columns,
frozen regions, named-region geometry, or other template structures create blank
printed pages, overflow, clipped tables, or stale visible content, diagnose the
structure as the cause and make the smallest targeted composition change. It is
acceptable to adjust or collapse inherited scaffolding when preserving it damages
the final workbook. After every repair, regenerate and rerun `--qa deep` or
`--qa strict`.

## Current Guarantees and Limits

M2 fills named cells and named regions while preserving formulas and workbook
topology in the shell. Region fills that exceed the named range are refused
before saving. When a current comprehension block is present, generation can
clear corroborated cover/demo regions while preserving formulas; derived indexes
remain out of scope for XLSX because the field inventory is intentionally empty.

The two-stage visual audit closes the "L0-only" gap: L1 deterministic pixel
proxies catch rendered-layout defects L0 cannot see (blank/broken printed pages,
content bleeding past the printable margins), and the L2 manifest drives the
orchestrator's qualitative judgement and repair loop. See
[reference/visual-audit.md](reference/visual-audit.md). When `soffice` and both
PDF rasterizers (`pdftoppm`, optional PyMuPDF/`fitz`) are absent (e.g. CI), the
audit degrades cleanly to L0 plus a single INFO
`visual.unavailable`; exit codes are unchanged.
