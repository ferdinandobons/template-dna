---
name: brand-pptx
description: Brand-aware PowerPoint engine. Use to extract a reusable Brand Profile from a .pptx template, verify it, and generate a new on-brand .pptx from an IntermediateDocument. For one-off slide edits with no saved brand profile, use the normal pptx skill instead.
---

# brand-pptx

Use this skill when the user wants reusable branded PowerPoint generation from
a company `.pptx` template and variable user-provided content.

This is an AI-agent skill for Codex and Claude Code. The user should describe
the deck they want; the agent converts that request into an IntermediateDocument,
uses the internal Python engine, verifies the output, and returns the generated
`.pptx`.

## Agent Workflow

1. Determine the brand name and locate the user-provided `.pptx` template.
2. If no matching `brand-kit/<name>` exists, extract one.
3. Convert the user's outline/content into `IntermediateDocument` JSON.
4. Generate the `.pptx` with the internal engine.
5. Run QA and report any warnings honestly.

Before generation, inspect `profile.json.artifact_catalog` when the user asks
to mimic a specific template piece. It records OOXML parts, media parts, slide
layouts, masters, placeholder geometry, slide texts, and slide size.

## Internal Extract

```bash
python scripts/cli.py extract --name <brand> --template <template.pptx> --scope project
```

## Internal Verify

```bash
python scripts/cli.py verify --name <brand> --scope auto --qa auto
```

## Internal Generate

```bash
python scripts/cli.py generate --name <brand> --input <intermediate-document.json> --output <output.pptx> --scope auto --qa auto
```

M2 supports title/content deck generation from the saved shell. Long content is
split across multiple content slides with a conservative capacity guard. Layout
and placeholder extraction are intentionally basic and will be deepened later.
