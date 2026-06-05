---
name: brand-pptx
description: >-
  Brand-aware PowerPoint engine. Use to (1) EXTRACT a company's brand from a .pptx template into a
  reusable "Brand Profile", (2) COMPREHEND the template with the model (optional), (3) VERIFY it,
  (4) GENERATE a new on-brand .pptx from an IntermediateDocument. Trigger on "extract our brand",
  "use our deck template", "generate a branded deck from our profile", or when a ./brand-kit
  exists. For one-off slide edits with no saved brand profile, use the normal pptx skill instead.
  NOT for .docx (brand-docx), .xlsx (brand-xlsx), or PDFs.
---

# brand-pptx

Use this skill when the user wants reusable branded PowerPoint generation from
a company `.pptx` template and variable user-provided content.

This is an AI-agent skill for Codex and Claude Code. The user should describe
the deck they want; the agent converts that request into an IntermediateDocument,
uses the internal Python engine, verifies the output, and returns the generated
`.pptx`.

## The four verbs

Every brand skill (`brand-docx`, `brand-pptx`, `brand-xlsx`) implements the same
contract: **extract / comprehend / verify / generate**.

| Verb | Input | Output |
|---|---|---|
| **extract** | a company `.pptx` template | a reusable Brand Profile |
| **comprehend** *(optional, model-driven)* | a saved profile + a model-authored `comprehension.json` | the profile with a validated, cached `comprehension` block |
| **verify** | a saved Brand Profile | QA findings + a verdict |
| **generate** | content (an IntermediateDocument) + a profile | a new on-brand `.pptx` |

`comprehend` is **optional**: `generate` works on the deterministic profile alone.
See [reference/comprehension.md](reference/comprehension.md) for the full step.

## Hard Rules

- Treat `python scripts/brandkit/cli.py ...` as an internal engine command, not the user-facing workflow.
- Extract opens the source template read-only and saves `brand-kit/<name>/template/shell.pptx` byte-for-byte.
- Generate opens the saved shell and resolves every semantic block through `profile.json`.
- Do not put style names, colors, fonts, or brand identifiers in an IntermediateDocument.
- If the user did not provide a template or enough content, ask for the missing input.
- Return the generated file path plus a QA summary.
- Consult `profile.json.artifact_catalog` before generation when the user asks to mimic a specific piece of the template.

## Agent Workflow

1. Determine the brand name and locate the user-provided `.pptx` template.
2. If no matching `brand-kit/<name>` exists, **extract** one.
3. **Comprehend** the template (optional, model-driven) — see below. Skip when a
   current comprehension is already cached or no model is available.
4. Convert the user's outline/content into `IntermediateDocument` JSON.
5. **Generate** the `.pptx` with the internal engine.
6. Run **QA** and report any warnings honestly.

Before generation, inspect `profile.json.artifact_catalog` when the user asks
to mimic a specific template piece. It records OOXML parts, media parts, slide
layouts, masters, placeholder geometry, slide texts, and slide size.

## Internal Extract

```bash
python scripts/brandkit/cli.py extract --name <brand> --template <template.pptx> --scope project
```

## Internal Comprehend (optional, model-driven)

Read [reference/comprehension.md](reference/comprehension.md) for the full
guidance, the four questions, and the anti-overfitting directive. In short:

```bash
python scripts/brandkit/cli.py comprehend-input --name <brand>   # prints {facts, excerpt} for the model
python scripts/brandkit/cli.py comprehend --name <brand> --input comprehension.json  # the ONLY writer
```

Skip this verb when `comprehension.status` is `present` **and** its
`source_shell_sha256` equals the live `provenance.shell.sha256`. Never re-run it
at generate time.

> **pptx readiness.** The PowerPoint extractor does not yet surface cover-anchor
> or derived-index inventories, so those questions have no ids to bind to and
> `comprehension.status` stays `absent` — `generate` runs the deterministic path.
> Do not force a cover or index shape onto an empty inventory; a ref into an empty
> inventory is fail-closed and will be rejected. Deeper pptx comprehension lands on
> the pptx fact-enrichment milestone.

## Internal Verify

```bash
python scripts/brandkit/cli.py verify --name <brand> --scope auto --qa auto
```

`--qa` selects the QA depth (see [reference/visual-audit.md](reference/visual-audit.md)):

- `fast` — deterministic **L0** only.
- `auto` — L0 **+ L1** visual pixel proxies when renderers (`soffice` + `pdftoppm`) are present; otherwise L0 plus a single INFO `visual.unavailable`.
- `deep` — L0 + L1 **+ a `visual_manifest.json`** and per-page PNGs; the orchestrator must then run the **L2** step (see below).

Verify has no output to render, so all three modes behave as L0 at verify time; the visual stages run at **generate** time.

## Internal Generate

```bash
python scripts/brandkit/cli.py generate --name <brand> --input <intermediate-document.json> --output <output.pptx> --scope auto --qa auto
```

See `reference/comprehension.md` and `reference/visual-audit.md`.

## Visual audit (two-stage)

The engine renders the output and runs deterministic pixel proxies, but the
**qualitative visual judgement is yours (the orchestrator), never the engine's** —
the Python engine never calls a model. To run the full two-stage audit:

1. Generate with `--qa deep`. The engine renders each slide to a PNG, runs the L1
   proxies, and writes `visual_manifest.json` next to the output in an
   `<output>.visual/` dir (a side artifact; the `.pptx` bytes never change).
2. Read the manifest path from stdout (`visual manifest: <path>`).
3. Open the PNGs listed in `pages[*].png`. For every entry in `checklist`, judge
   PASS/FAIL against the rendered pages, taking `l1_findings` into account.
4. If any checklist item FAILS (or an L1 WARNING is confirmed visually as a real
   defect): **repair** the IntermediateDocument/content, **regenerate**, then
   **re-run the audit**. Loop until the checklist is clean (max 3 iterations by
   default).

L1 findings are WARNING-only and never fail the gate by themselves; the real
qualitative gate is your L2 judgement.

## Current Guarantees and Limits

M2 supports title/content deck generation from the saved shell. Long content is
split across multiple content slides with a conservative capacity guard. Layout
and placeholder extraction are intentionally basic and will be deepened on the
pptx fact-enrichment milestone; until then comprehension stays `absent` and
generation uses the proven deterministic path.

The two-stage visual audit closes the "L0-only" gap: L1 deterministic pixel
proxies catch rendered-layout defects L0 cannot see (blank/broken slides, content
bleeding past the slide edges), and the L2 manifest drives the orchestrator's
qualitative judgement and repair loop. See
[reference/visual-audit.md](reference/visual-audit.md). When `soffice`/`pdftoppm`
are absent (e.g. CI), the audit degrades cleanly to L0 plus a single INFO
`visual.unavailable`; exit codes are unchanged.
