# DOCX Extraction

Extraction copies the source `.docx` as `template/shell.docx`, parses theme
colors when available, records paragraph/table styles, discovers a simple cover
title placeholder, detects demo instruction text, and infers the core DOCX roles.

## Ordered document structure (schema 1.1.0+)

Extraction also detects the template's **ordered top-level skeleton** — the
sequence of regions a generated document of this brand must respect — and records
it in `profile.structure`:

```jsonc
"structure": {
  "ordered": true,                 // top-level region order must be respected
  "skeleton": [
    {"region": "cover", "order": 0, "role": "section.cover",
     "required": true,  "repeatable": false, "evidence": "..."},
    {"region": "toc",   "order": 1, "role": "section.toc",
     "required": true,  "repeatable": false, "evidence": "..."},
    {"region": "body",  "order": 2, "role": "section.body",
     "required": true,  "repeatable": true,  "freeform": true, "evidence": "..."}
  ]
}
```

Only regions actually present in the template are included. The `body` region is
**freeform**: element order *inside* it is not prescribed.

### How regions are detected (brand-agnostic, multilingual)

Detection runs on the lxml body element (`doc.element.body`), never on
python-docx paragraph indices, because the cover title and the TOC commonly live
inside block-level `w:sdt` elements that python-docx does not expose.

- **Cover region** — body-level content *before* the first TOC region or first
  Heading-1 paragraph. Cover anchors (SDTs / placeholders / logos) come from
  `cover.discover_cover()`.
- **TOC region** — any of:
  - a block-level `w:sdt` whose `w:docPartGallery/@w:val` is `Table of Contents`,
  - a paragraph using a TOC / TOCHeading style (style id/name contains
    `toc`/`sommario`/`indice`/`inhalt`/`contenido`),
  - a `w:instrText` whose text starts with `TOC`, or
  - a heading whose text is a known contents word in EN/IT/FR/DE/ES
    (`Contents`, `Sommario`, `Indice`, `Inhalt`, `Table des matières`, `Índice`,
    `Contenido`, …).
- **Body region** — everything after the TOC (or after the cover when there is no
  TOC) up to the final body-level `w:sectPr`.

The detector is brand-agnostic by design: instead of hardcoding one template's
`TOCHeading`/`Sommario` literals, it matches any `*toc*`-named style token and the
multilingual contents words above, so it works on any company template in any language.

`anchors.toc.present` is a **real** detection result (no longer hardcoded
`false`).

## Per-artifact usage

Every role entry in `roles` (and every component, when present) carries a `usage`
object recording where and how the artifact is used. See
[profile-schema.md](profile-schema.md#per-artifact-usage) for the vocabulary. In
short: `cover.*` roles are `structural`/`required` at the front of the skeleton,
`toc` is `structural`/`required` after the cover, and body roles (heading,
paragraph, list, callout, table, quote, caption) are `freeform` and used on
demand. The `usage` is derived from the role *family* (already inferred from style
placement / OOXML), never from a brand-specific name.

`PROFILE.md` prints a `## Structure` section (the ordered skeleton) and, next to
each role, its usage (`scope · placement · required`), so a human or agent reading
it sees exactly which parts to respect in order versus use on demand.
