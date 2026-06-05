---
name: brand-docx
description: >-
  Brand-aware Word engine. Use to (1) EXTRACT a company's brand from a Word template into a
  reusable "Brand Profile", (2) COMPREHEND the template with the model (optional), (3) VERIFY it,
  (4) GENERATE new on-brand .docx documents FROM a saved profile. Trigger on "extract our brand",
  "learn/match this template", "use our brand kit", "generate a branded report from our profile",
  or when a ./brand-kit exists. For one-off Word edits with NO saved brand profile, use the docx
  skill instead. NOT for .pptx (brand-pptx), .xlsx (brand-xlsx), or PDFs.
---

# brand-docx

Use this skill when the user wants a reusable Word brand kit or wants to create
a new on-brand `.docx` from a company template and variable content.

This is an AI-agent skill for Codex and Claude Code. The user should not need to
write JSON or run shell commands. The agent converts the user's content into an
IntermediateDocument, invokes the internal engine, verifies the output, and
returns the generated `.docx`.

## The four verbs

Every brand skill (`brand-docx`, `brand-pptx`, `brand-xlsx`) implements the same
contract: **extract / comprehend / verify / generate**.

| Verb | Input | Output |
|---|---|---|
| **extract** | a company `.docx` template | a reusable Brand Profile |
| **comprehend** *(optional, model-driven)* | a saved profile + a model-authored `comprehension.json` | the profile with a validated, cached `comprehension` block |
| **verify** | a saved Brand Profile | QA findings + a verdict |
| **generate** | content (an IntermediateDocument) + a profile | a new on-brand `.docx` |

`comprehend` is **optional**: `generate` works on the deterministic profile alone.
When a current comprehension is present, `generate` additionally reconciles the
template's preserved cover/index structures with the new content. See
[reference/comprehension.md](reference/comprehension.md) for the full step.

## Hard Rules

- Treat `python scripts/brandkit/cli.py ...` as an internal engine command, not the user-facing workflow.
- Extract opens the source template read-only and saves `brand-kit/<name>/template/shell.docx` byte-for-byte.
- Generate opens the saved shell and resolves every semantic block through `profile.json`.
- Do not put style names, colors, fonts, or brand identifiers in an IntermediateDocument.
- If the user did not provide a template or enough content, ask for the missing input.
- Return the generated file path plus a QA summary.
- Consult `profile.json.artifact_catalog` before generation when the user asks to mimic a specific piece of the template.

## Agent Workflow

1. Determine the brand name and locate the user-provided `.docx` template.
2. If no matching `brand-kit/<name>` exists, **extract** one.
3. **Comprehend** the template (optional, model-driven) — see below. Skip when a
   current comprehension is already cached or no model is available.
4. Convert the user's requested content into `IntermediateDocument` JSON.
5. **Generate** the `.docx` with the internal engine.
6. Run **QA** and report any warnings honestly.

## Internal Extract

```bash
python scripts/brandkit/cli.py extract --name <brand> --template <template.docx> --scope project
```

## Internal Comprehend (optional, model-driven)

Read [reference/comprehension.md](reference/comprehension.md) for the full
guidance, the four questions, and the anti-overfitting directive. In short:

```bash
python scripts/brandkit/cli.py comprehend-input --name <brand>   # prints {facts, excerpt} for the model
python scripts/brandkit/cli.py comprehend --name <brand> --input comprehension.json  # the ONLY writer
```

Skip this verb when `comprehension.status` is `present` **and** its
`source_shell_sha256` equals the live `provenance.shell.sha256` (a current
comprehension is already cached). A re-extract resets it to `absent`; re-run
`comprehend` only then. Never re-run it at generate time.

## Internal Verify

```bash
python scripts/brandkit/cli.py verify --name <brand> --scope auto --qa auto
```

`--qa` selects the QA depth (see [reference/visual-audit.md](reference/visual-audit.md)):

- `fast` — deterministic **L0** only (schema, resolver targets, residual text, structural diffs).
- `auto` — L0 **+ L1** visual pixel proxies when renderers (`soffice` + `pdftoppm`) are present; otherwise L0 plus a single INFO `visual.unavailable`.
- `deep` — L0 + L1 **+ a `visual_manifest.json`** and per-page PNGs; the orchestrator must then run the **L2** step (see below).

Verify has no output to render, so all three modes behave as L0 at verify time; the visual stages run at **generate** time.

## Internal Generate

```bash
python scripts/brandkit/cli.py generate --name <brand> --input <intermediate-document.json> --output <output.docx> --scope auto --qa auto
```

See `reference/comprehension.md`, `reference/profile-schema.md`,
`reference/generation.md`, `reference/visual-audit.md`, and
`examples/intermediate-document.example.json`.

## Visual audit (two-stage)

The engine renders the output and runs deterministic pixel proxies, but the
**qualitative visual judgement is yours (the orchestrator), never the engine's** —
the Python engine never calls a model. To run the full two-stage audit:

1. Generate with `--qa deep`. The engine renders each page to a PNG, runs the L1
   proxies, and writes `visual_manifest.json` next to the output in an
   `<output>.visual/` dir (a side artifact; the `.docx` bytes never change).
2. Read the manifest path from stdout (`visual manifest: <path>`).
3. Open the PNGs listed in `pages[*].png`. For every entry in `checklist`, judge
   PASS/FAIL against the rendered pages, taking `l1_findings` into account.
4. If any checklist item FAILS (or an L1 `visual.blank_page` / `visual.edge_bleed`
   WARNING is confirmed visually as a real defect): **repair** the
   IntermediateDocument/content, **regenerate**, then **re-run the audit**. Loop
   until the checklist is clean (max 3 iterations by default).

L1 findings are WARNING-only and never fail the gate by themselves; the real
qualitative gate is your L2 judgement.

## Current Guarantees and Limits

Generation opens the saved `.docx` shell, clears detected demo text, and applies
only styles resolved from `profile.json`. L0 QA catches schema problems,
unresolved roles, markdown literals, and residual demo text.

When a current comprehension is present, generation also fills the cover slots in
place (no duplicate title) and reconciles preserved indexes (a table of contents,
a list of tables/figures) against the new content — regenerating, preserving, or
purging stale entries — instead of carrying demo entries forward. Destructive
reconciliation is bounded: a clear/remove is honored only when determinism
corroborates and confidence clears a threshold, else the structure is kept with a
warning.

Extraction also records a broad `artifact_catalog`: OOXML parts, media parts,
paragraph/table styles, style details, sections/margins, paragraph samples, and
table counts. Use it to understand and describe template conventions beyond the
roles that are directly generatable today.

The two-stage visual audit closes the "L0-only" gap: L1 deterministic pixel
proxies catch rendered-layout defects L0 cannot see (blank/broken pages, content
bleeding past the printable margins), and the L2 manifest drives the
orchestrator's qualitative judgement and repair loop. See
[reference/visual-audit.md](reference/visual-audit.md).

DOCX visual overflow requires render-time QA with LibreOffice because Word
layout is not deterministic from OOXML alone. When `soffice`/`pdftoppm` are
absent (e.g. CI), the visual audit degrades cleanly to L0 plus a single INFO
`visual.unavailable`; exit codes are unchanged and the skill does not claim a
full no-overflow visual proof.
