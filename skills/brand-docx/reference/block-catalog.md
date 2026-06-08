# Block Catalog

The IntermediateDocument (IID) is an ordered flow of typed **blocks** that carry
**intent**, never presentation. No block ever names a style, hex, font, or layout;
the resolver maps each block to a concrete brand artifact in `profile.json`. The
authoritative model lives in `scripts/brandkit/ir/model.py` (the `BLOCK_TYPES`
registry is the closed set of `type` discriminators).

Inline text is a **rich-run array** (`runs: [{t, b?, i?, u?, code?, link?}]`); a
bare `text: "..."` is accepted as sugar and normalized to runs on parse.

M1 DOCX generation renders `heading`, `paragraph`, `callout`, `list`, `table`, and
`pagebreak` fully; the remaining types are present in the model and degrade
gracefully where a format's generator does not yet realize them natively.

## Block types

The 16 block `type` values (the keys of `BLOCK_TYPES`):

- **`heading`** - a section heading. Carries `level` (int) and `runs`. Resolves to
  `heading.{level}`.
- **`paragraph`** - a body paragraph. Carries `runs` and an optional `variant`.
  Resolves to `paragraph.{variant or default}`.
- **`list`** - an ordered or unordered list. Carries `ordered` (bool) and nested
  `items` (each a `ListItem` with `runs`, `level`, and optional sub-`items`).
  Resolves to `list.{bullet|number}.{level}`.
- **`table`** - a table. Carries `columns` (header-row rich cells), `rows` (lists
  of `TableCell` with optional `header`/`colspan`/`rowspan`), an optional
  `caption`, and a `role` (default `"default"`). Resolves to `table.{role}` plus a
  header style.
- **`callout`** - a highlighted box. Carries a semantic `intent`
  (`info` | `warning` | `danger` | `success` | `note`), `runs`, and an optional
  `title`. The intent is semantic; the brand picks the color. Resolves to
  `callout.{intent}`.
- **`kpi`** - a KPI / metric card group. Carries `items` (each a `KpiItem` with
  `label`, `value`, optional `delta`) and an optional `layout`. Resolves to
  `component:kpi.{layout}`.
- **`chart`** - a chart. Carries `chart_type` (default `"bar"`), `series`
  (`[{name, values}]`), `categories`, and an optional `title`. Authored as a
  NATIVE chart on **both docx and pptx** (a real DrawingML `c:chart`: an inline
  `w:drawing` on docx, a `graphicFrame` on pptx) that inherits the document/deck
  theme's accent colors, so it is on-brand by construction.
  `bar`/`column`/`barh`/`line`/`area`/`pie`/`doughnut` map to the matching chart
  type; an unknown type falls back to a clustered column chart (surfaced as INFO,
  never silent), an empty/all-non-numeric chart degrades loudly, and a multi-series
  pie/doughnut surfaces a truncation WARNING (only its first series renders). The
  docx chart uses inline cached data (no embedded workbook), so generation stays
  byte-idempotent. On **xlsx** charts are native too but authored differently: the
  grid model is range-based, so a chart is a `GridDocument.charts` entry
  (`{sheet?, type, title?, anchor, data, categories?, data_titles?}`) that
  REFERENCES the workbook's own cell ranges via openpyxl - the data lives in the
  sheet, which is the spreadsheet's strength.
- **`smartart`** - a diagram. Carries `diagram` (default `"process"`) and `nodes`
  (`[{text, children}]`). Resolves to `smartart.{diagram}` (clone-fill or rendered
  image).
- **`component`** - a reference to a reusable single fragment. Carries `ref` (names
  a `components/<ref>` profile entry) and `slots` (fills its render contract).
  Expanded into primitive sub-blocks before resolution.
- **`section`** - a reference to a multi-block reusable unit. Carries `ref` (names a
  `sections/<ref>` profile entry) and `slots`. Like `component`, expanded to
  primitives before resolution.
- **`caption`** - a figure/table caption line. Carries `runs` and an optional
  `target` (`"figure"` | `"table"`). Resolves to `caption`.
- **`toc`** - a table-of-contents marker. Carries an optional `title` and
  `max_level` (default `3`). The generator refreshes the live TOC field.
- **`image`** - an image reference. Carries `asset` (a profile asset id) or `src`
  (an external path) - exactly one - plus optional `alt`, `caption`, and
  `width_emu`/`height_emu` sizing hints. Resolves to an image placement.
- **`quote`** - a block quotation. Carries `runs` and an optional `attribution`.
  Resolves to `quote` (falls back to body).
- **`divider`** - a horizontal rule / separator. No payload. Resolves to a brand
  divider.
- **`pagebreak`** - an explicit page (docx) or slide (pptx) break. No payload.

Every block also accepts the base fields `id` (optional author-supplied
identifier, stable across regenerations) and `meta` (free-form annotations the
resolver ignores).

## Cover (not a flow block)

The document's `cover` is **not** in `BLOCK_TYPES`; it sits in
`IntermediateDocument.cover` as semantic slots only: `title`, `subtitle`, and a
free-form `fields` map (e.g. `{"doc_id": "RPT-2026-014", "date": "2026-06-04"}`).
The resolver maps these to whatever cover anchors the shell has. An absent cover
(`None`) means "leave the shell cover as-is".
