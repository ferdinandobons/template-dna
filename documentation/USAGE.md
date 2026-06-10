# Usage details

The [README quick start](../README.md#quick-start) covers the agent flow and the
three `extract → verify → generate` CLI commands. This page collects the common
use cases and the structured input format the generator expects.

## Common use cases

- **Consulting and operations reports**: generate branded Word reports, memos,
  briefs and status updates from the approved corporate template.
- **Sales and marketing decks**: create PowerPoint presentations from real masters
  and layouts instead of asking an AI to invent approximate slides.
- **Finance and planning workbooks**: fill named Excel inputs and regions while
  preserving formulas and workbook structure.
- **Repeatable agent workflows**: give Claude Code, Codex or another agent a
  reusable Brand Profile instead of re-explaining the brand for every document.

## Input format: the IntermediateDocument

The content you pass to `generate` (`idoc.json`) is an **IntermediateDocument** -
brand-agnostic typed blocks. Notice there is **no style, color or font anywhere**:
the profile resolves all of that.

```json
{
  "cover": {
    "title": "Quarterly Review",
    "subtitle": "Q2 - Revenue and delivery",
    "fields": { "doc_id": "RPT-001", "date": "2026-06-10" }
  },
  "blocks": [
    { "type": "toc", "title": "Contents", "max_level": 3 },
    { "type": "heading", "level": 1, "text": "Highlights" },
    { "type": "paragraph", "text": "This paragraph resolves to the brand body style." },
    { "type": "callout", "intent": "info", "text": "The profile chooses the callout style." },
    { "type": "list", "items": [{ "text": "List styling comes from the profile." }] },
    { "type": "table", "columns": ["Area", "Status"], "rows": [["Pipeline", "Healthy"], ["Delivery", "Green"]] }
  ]
}
```

Structural notes:

- The optional top-level **`cover`** object carries named semantic slots only:
  `title`, an optional `subtitle`, and an optional `fields` mapping of
  key-to-value pairs (document id, date, author, ...). The profile decides where
  and how each slot renders on the template's own cover.
- A **`toc`** block is a table-of-contents placeholder, not content: it carries an
  optional `title` and a `max_level` (default 3), and the engine emits a native
  field or defers to a preserved template outline.
- Block order in `blocks` is the author's reading order; every block resolves
  through semantic roles, so nothing in the JSON names a style, color or font.

A comprehensive worked example lives at
[`skills/brand-docx/examples/intermediate-document.example.json`](../skills/brand-docx/examples/intermediate-document.example.json).
PowerPoint uses the same `IntermediateDocument`; Excel uses a `GridDocument`
(named-region fills, formulas preserved).

QA depth is explicit via `--qa fast|auto|deep|strict`; `deep`/`strict` write a
visual manifest for render-based review and targeted repair. Run
`python scripts/brandkit/cli.py doctor` to preflight dependencies.
