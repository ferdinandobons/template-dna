# Brand Profile Schema

The schema discriminator is `kind`. A docx profile writes `kind: "docx"` with a
`surface.docx` block and a role registry named `roles`.

Concrete style ids and names live only in `profile.json`. IntermediateDocument
content uses semantic role intent such as `heading`, `paragraph`, `callout`, and
`table`.

`schema_version` is **`1.1.0`** (`SCHEMA_VERSION`). The 1.1.0 minor bump is
**additive**: it introduces the optional `structure` section and the optional
per-role `usage` object. Profiles written by older extractors (1.0.0, lacking
both) remain valid — `validate()` never rejects a profile for missing either.

## The `structure` section

```jsonc
"structure": {
  "ordered": true,
  "skeleton": [
    {"region": "cover", "order": 0, "role": "section.cover",
     "required": true, "repeatable": false, "evidence": "..."},
    {"region": "toc",   "order": 1, "role": "section.toc",
     "required": true, "repeatable": false, "evidence": "..."},
    {"region": "body",  "order": 2, "role": "section.body",
     "required": true, "repeatable": true, "freeform": true, "evidence": "..."}
  ]
}
```

- `ordered` (bool) — when `true`, the top-level region order must be respected on
  generation.
- `skeleton` — an ordered list of the regions **actually present** in the
  template. Each region carries:
  - `region` — one of `cover` | `toc` | `body`.
  - `order` (int) — the region's position in the skeleton.
  - `role` — the section role id (`section.cover` | `section.toc` |
    `section.body`).
  - `required` (bool) — must appear in every document of this brand.
  - `repeatable` (bool) — the region may occur more than once (true for `body`).
  - `freeform` (bool, body only) — element order *inside* the region is not
    prescribed.
  - `evidence` — the human-readable signal the region was detected from.

An empty/absent `structure` (or `skeleton: []`) means no ordered skeleton was
detected; generation falls back to the legacy "append into body" behaviour.

## Per-artifact `usage`

Every role entry in `roles` (and every component, when present) may carry a
`usage` object:

```jsonc
"usage": {
  "scope": "cover" | "toc" | "body" | "anywhere",
  "placement": "structural" | "freeform",
  "required": true | false,
  "order": <int or null>
}
```

- `scope` — which region the artifact belongs to.
- `placement` — `structural` = part of the ordered skeleton (must appear in its
  slot); `freeform` = used on demand inside the freeform body region.
- `required` — must appear in every document of this brand.
- `order` — the skeleton position when `structural`, else `null`.

Derivation (from the role *family*, never a brand-specific name):

| role family | scope | placement | required | order |
|---|---|---|---|---|
| `cover.*` | cover | structural | true | 0 |
| `toc` | toc | structural | true | 1 |
| `heading` / `paragraph` / `list` / `callout` / `table` / `quote` / `caption` … | body | freeform | false | null |

`validate()` accepts a profile with **or without** `usage` on its roles; when
present, it checks the `scope`/`placement` enums and that `order` is an int or
null.
