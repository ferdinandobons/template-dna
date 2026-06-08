# Brand Profile Schema

The schema discriminator is `kind`. A docx profile writes `kind: "docx"` with a
`surface.docx` block and a role registry named `roles`.

Concrete style ids and names live only in `profile.json`. IntermediateDocument
content uses semantic role intent such as `heading`, `paragraph`, `callout`, and
`table`.

`schema_version` is **`1.2.0`** (`SCHEMA_VERSION`). The minor bumps are
**additive**, so every profile written by an older extractor stays valid -
`validate()` never rejects a profile for a section it lacks:

- **1.1.0** introduced the optional `structure` section and the optional per-role
  `usage` object. Profiles written by 1.0.0 extractors (lacking both) remain valid.
- **1.2.0** introduced the optional top-level `comprehension` block - the single
  canonical sink for the model's understanding of the template (cover slots,
  derived-index/section conventions, role annotations, demo-vs-real
  classification). Every extractor now stamps an empty `comprehension` block with
  `status: "absent"` by default, so the deterministic path is the ground truth
  until `comprehend` runs. 1.2.0 also relaxes region names (`structure.skeleton[].region`
  and the comprehension refs) from the frozen `cover`/`toc`/`body` trio to **open
  tokens** validated for syntax only - the generator branches on the boolean
  region attributes, never on the name.

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

- `ordered` (bool) - when `true`, the top-level region order must be respected on
  generation.
- `skeleton` - an ordered list of the regions **actually present** in the
  template. Each region carries:
  - `region` - an **open token** (since 1.2.0) validated for syntax only (dotted
    lowercase, like a role id). docx conventionally uses the `cover` | `toc` |
    `body` trio shown here; pptx/xlsx may use their own honest names (`agenda`,
    `appendix`, `sheet`, …). The generator branches on the boolean attributes
    below, never on the name.
  - `order` (int) - the region's position in the skeleton.
  - `role` - the section role id (`section.cover` | `section.toc` |
    `section.body`).
  - `required` (bool) - must appear in every document of this brand.
  - `repeatable` (bool) - the region may occur more than once (true for `body`).
  - `freeform` (bool, body only) - element order *inside* the region is not
    prescribed.
  - `evidence` - the human-readable signal the region was detected from.

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

- `scope` - which region the artifact belongs to.
- `placement` - `structural` = part of the ordered skeleton (must appear in its
  slot); `freeform` = used on demand inside the freeform body region.
- `required` - must appear in every document of this brand.
- `order` - the skeleton position when `structural`, else `null`.

Derivation (from the role *family*, never a brand-specific name):

| role family | scope | placement | required | order |
|---|---|---|---|---|
| `cover.*` | cover | structural | true | 0 |
| `toc` | toc | structural | true | 1 |
| `heading` / `paragraph` / `list` / `callout` / `table` / `quote` / `caption` … | body | freeform | false | null |

`validate()` accepts a profile with **or without** `usage` on its roles; when
present, it checks the `scope`/`placement` enums and that `order` is an int or
null.
