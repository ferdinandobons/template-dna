---
name: brand-docx
description: >-
  Brand-aware Word engine. Use to (1) EXTRACT a company's brand from a Word template into a
  reusable "Brand Profile", (2) VERIFY it, (3) GENERATE new on-brand .docx documents FROM a
  saved profile. Trigger on "extract our brand", "learn/match this template", "use our brand
  kit", "generate a branded report from our profile", or when a ./brand-kit exists. For one-off
  Word edits with NO saved brand profile, use the docx skill instead. NOT for .pptx (brand-pptx),
  .xlsx (brand-xlsx), or PDFs.
---

# brand-docx

Use this skill when the user wants a reusable Word brand kit or wants to create
a new on-brand `.docx` from a company template and variable content.

This is an AI-agent skill for Codex and Claude Code. The user should not need to
write JSON or run shell commands. The agent converts the user's content into an
IntermediateDocument, invokes the internal engine, verifies the output, and
returns the generated `.docx`.

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
2. If no matching `brand-kit/<name>` exists, extract one.
3. Convert the user's requested content into `IntermediateDocument` JSON.
4. Generate the `.docx` with the internal engine.
5. Run QA and report any warnings honestly.

## Internal Extract

```bash
python scripts/brandkit/cli.py extract --name <brand> --template <template.docx> --scope project
```

## Internal Verify

```bash
python scripts/brandkit/cli.py verify --name <brand> --scope auto --qa auto
```

Use `--qa fast` for deterministic L0 only. M1 degrades visual QA gracefully when render tools are absent.

## Internal Generate

```bash
python scripts/brandkit/cli.py generate --name <brand> --input <intermediate-document.json> --output <output.docx> --scope auto --qa auto
```

See `reference/profile-schema.md`, `reference/generation.md`, and `examples/intermediate-document.example.json`.

## Current Guarantees and Limits

Generation opens the saved `.docx` shell, clears detected demo text, and applies
only styles resolved from `profile.json`. L0 QA catches schema problems,
unresolved roles, markdown literals, and residual demo text.

Extraction also records a broad `artifact_catalog`: OOXML parts, media parts,
paragraph/table styles, style details, sections/margins, paragraph samples, and
table counts. Use it to understand and describe template conventions beyond the
roles that are directly generatable today.

DOCX visual overflow requires render-time QA with LibreOffice because Word
layout is not deterministic from OOXML alone. When `soffice` is absent, the skill
does not claim a full no-overflow visual proof.
