<!-- SPDX-License-Identifier: MIT -->
# template-dna — Frozen Conventions

This is the **single source of truth** for the vocabulary of `template-dna`. It is
owned by `scripts/brandkit/profile/schema.py`; this document is its human-readable
mirror. Every `SKILL.md`, command, and reference doc quotes these names **verbatim**.
If a name here disagrees with code, the code (the schema module) wins and this file
is the bug.

> **The one hard rule.** Brand-specific identifiers — style names, theme color
> tokens, font names, cover aliases, layout ids, chart-template ids — live in the
> Brand Profile and **nowhere else**. No IntermediateDocument block, and no writer,
> ever contains a literal style name, hex color, or font. The resolver is the only
> code that reads them. Off-brand output is impossible by construction.

---

## 1. The three verbs

Every skill (`brand-docx`, `brand-pptx`, `brand-xlsx`) implements the same contract:

| Verb | Input | Output |
|---|---|---|
| **extract** | a company `.docx`/`.pptx`/`.xlsx` template | a reusable **Brand Profile** |
| **verify** | a saved Brand Profile | a QA report + a rendered proof + a role-mapping table |
| **generate** | content (free text or an IntermediateDocument) + a profile | a new on-brand document |

Invocation is by **bare relative path**: `python scripts/cli.py extract --name <brand> --template <t>`.
`${CLAUDE_PLUGIN_ROOT}` never appears in author-facing SKILL.md.

---

## 2. The Brand Profile directory

A profile is a **self-contained, copyable directory**. All internal paths are
relative to the profile root, so it can move between stores unchanged.

```
brand-kit/<name>/
├─ profile.json          # the index: schema-versioned envelope
├─ PROFILE.md            # human-readable palette + role table
├─ template/
│  └─ shell.<ext>        # the byte-for-byte shell the generator opens FROM
├─ assets/               # logos + manifest.json
├─ components/           # reusable fragments + index.json
├─ sections/             # multi-page/slide units + index.json
├─ specimens/            # opaque chart/SmartArt captures
├─ samples/              # source.png (template) + smoke.png (generated proof)
├─ provenance.sha256     # the shell hash (drift detection)
└─ .cache/               # resolved-hex / render caches (safe to delete)
```

The binary shell is **byte-for-byte** (never round-tripped from JSON). The JSON
**describes and points; the shell IS the brand.**

### Dual store (project wins)

| Store | Path |
|---|---|
| project (wins) | `./brand-kit/<name>` |
| global | `~/.claude/brand-kit/<name>` |

`scope` is `auto` (project then global), `project`, or `global`.

---

## 3. Frozen vocabulary (schema-owned)

These constants and enums are defined in `schema.py`. Do not invent synonyms.

### Discriminator
- **`kind`** — the format discriminator. Values: **`docx` | `pptx` | `xlsx`**.
  (Never `doc_type`.) Selects which `surface.*` block and which resolver types are legal.

### Schema version
- **`schema_version`** — semver string. Current: **`1.1.0`** (`SCHEMA_VERSION`).
  The 1.1.0 minor bump is **additive**: the optional `structure` section and the
  optional per-role `usage` object (see §12). 1.0.0 profiles lacking either stay
  valid — `validate()` never rejects a profile for missing them.
- `$schema` id: `https://template-dna/schema/profile-1.json` (`SCHEMA_ID`).

### Role registry
- **`roles`** — the join table from semantic role id → concrete resolver descriptor.
  (Never `bindings`.) `roles._index` lists every concrete role id.

### Resolver types (`resolver.type`)
| Value | Used by `kind` | Meaning |
|---|---|---|
| **`named_style`** | docx, pptx | a named OOXML style (keyed on `style_id` + `style_type`) |
| **`placeholder`** | pptx | a layout placeholder (`layout` + `ph_idx`/`ph_type`) |
| **`cell_style`** | xlsx | a named cell style |
| `number_format` | xlsx | a number-format vocabulary entry (staged) |
| `named_range` | xlsx | a named range fill target (staged) |
| `chart_template` | all | a captured chart specimen part (staged) |

(The redundant `layout_placeholder` is **dropped**.)

### Status (capability maturity) — `Status`
**`robust` | `best_effort` | `stub`**. Maturity is *data, not structure*: the
schema holds charts/SmartArt from day one at `stub`; generation degrades gracefully
below `robust`.

### Severity (QA findings) — `Severity`
**`INFO` | `WARNING` | `ERROR`**.

### Verification status — `VerificationStatus`
**`passed` | `passed_with_warnings` | `failed` | `unverified`**.

### Anchor kind — `AnchorKind`
**`sdt_anchored` | `placeholder` | `named_range` | `NONE`**. Anchor presence is
first-class: when no cover anchors exist, the kind is `NONE` (never a silently
empty list).

### Overflow capability — `OverflowCapability`
**`estimator` (pptx) | `cellfit` (xlsx) | `render` (docx) | `none`**. DOCX has **no
deterministic overflow estimator** — Word reflows; overflow is a render-time (L1/L2)
concern only.

---

## 4. Role ids

A role id is a dotted, lowercase path: `family[.qualifier[.qualifier...]]`.
Compose with `schema.role_id(...)`, never by hand-concatenation.

Canonical role ids (the `_index` for a typical docx profile):

```
heading.1  heading.2  heading.3  …
paragraph  paragraph.<variant>
list.bullet.1  list.bullet.2  …  list.number.1  …
table.default  table.<role>
callout.info  callout.warning  callout.danger  callout.success  callout.note
caption
cover.title  cover.subtitle
toc
quote  divider  image
chart.bar  chart.line  chart.pie  …
smartart.process  …
kpi.<layout>
```

Each role entry carries: `resolver`, `appearance` (transforms kept, **not** flattened),
`verified`, `confidence`, `status`, `evidence`.

---

## 5. Role inference scorer (frozen weights)

`score(role R, candidate style S) = Σ wᵢ · signalᵢ`:

| Signal | Weight | Measures |
|---|---|---|
| **inheritance** | **0.40** | S resolves via `basedOn` to the builtin defining R |
| **fingerprint** | **0.30** | OOXML markers (`w:shd` fill + border → callout; `numPr`/`numFmt=bullet` → bullet) |
| **name-token** | **0.20** | multilingual lexicon (`box`/`riquadro`/`encadré`/`kasten`; `bullet`/`puce`…) |
| **placement** | **0.10** | observed only in cover region → cover_*; only in tables → table_* |

**Accept floor: `0.45`.** Below-floor roles are `unresolved` (reported), never
mis-bound. Concretes are keyed on **`style_id` + `style_type`**, never the localized
display name. `w:link` paragraph/character pairs collapse to one role.

The multilingual name-token lexicon (EN/IT/FR/DE/ES) is `text.NAME_TOKEN_LEXICON`;
the name-token is the **weakest** signal and never overrides structure.

---

## 6. The IntermediateDocument (flow content model)

Used by **docx** and **pptx**. (xlsx uses a separate **GridDocument**.) The author
supplies an ordered list of typed blocks carrying **intent**, never presentation.

### Inline rich runs
A run is `{"t": str, "b"?, "i"?, "u"?, "strike"?, "code"?, "sup"?, "sub"?, "link"?}`.
`"b"/"i"/…` are booleans; `"link"` is a URL string. A bare `"text": "..."` on a block
is **sugar** normalized to `[{"t": "..."}]` on parse.

### Block catalog (16 flow types + the cover)

| `type` | key fields | resolves to (via role) |
|---|---|---|
| `heading` | `level`, `runs`/`text` | `heading.{level}` |
| `paragraph` | `runs`, `variant?` | `paragraph.{variant or default}` |
| `list` | `ordered`, `items:[{runs,level,items?}]` | `list.{bullet\|number}.{level}` |
| `table` | `columns`, `rows`, `caption?`, `role?` | `table.{role}` + header style |
| `callout` | `intent`, `runs`, `title?` | `callout.{intent}` (brand picks color) |
| `kpi` | `items:[{label,value,delta?}]`, `layout?` | `component:kpi.{layout}` |
| `chart` | `chart_type`, `series`, `categories`, `title?` | `chart.{type}` (clone-fill) |
| `smartart` | `diagram`, `nodes` | `smartart.{diagram}` (clone-fill) |
| `component` | `ref`, `slots` | expands pre-resolve into primitives |
| `section` | `ref`, `slots` | expands pre-resolve into primitives |
| `caption` | `runs`, `target?` | `caption` |
| `toc` | `title?`, `max_level` | refreshes the live TOC field |
| `image` | `asset?`/`src?`, `alt?`, `caption?`, `width_emu?`, `height_emu?` | image placement |
| `quote` | `runs`, `attribution?` | `quote` (falls back to body) |
| `divider` | — | brand divider |
| `pagebreak` | — | page (docx) / slide (pptx) break |

`callout.intent` ∈ `info | warning | danger | success | note` (semantic; the **brand**
decides the color). The author **never** picks a style.

### Cover (not a flow block)
The document's `cover` carries semantic slots only: `title`, `subtitle`, `fields{}`.
`fields` keys are matched to cover anchor ids by the resolver. An absent cover means
"leave the shell cover as-is".

### Document shape
```jsonc
{ "cover": { "title": "...", "subtitle": [...], "fields": {"doc_id": "..."} },
  "meta":  { "lang": "en-US" },
  "blocks": [ { "type": "heading", "level": 1, "text": "..." }, ... ] }
```

---

## 7. Generate — mandatory order

Enforced by one orchestrator; the author cannot reorder steps.

1. `expand_components` — component/section refs → primitives.
2. `validate_iid` — capability pre-flight (degradations reported **before** writing).
3. **Open the shell** (`Document(shell)`, never blank).
4. `clear_body_region` — remove **only** the freeform body region, **preserving**
   the ordered cover region and the TOC region (and the final `sectPr`). Boundary
   by region evidence, never index (§11). Never wipes the whole body.
5. `compose_cover` — fill the **preserved** cover anchors in place (sole author of
   the cover; never recreates it after a wipe).
6. `compose_body` — resolver-driven flat dispatch; new blocks are appended into the
   now-empty body region, immediately before `sectPr` (after the cover and TOC).
7. `refresh_toc` — refresh the preserved live TOC if present (mark fields dirty /
   `w:updateFields`) so Word recomputes it on open; never duplicate it. No-op when
   there is no TOC.
8. `finalize` — set `w:updateFields` + per-field `w:dirty`; save.

The order-aware clear (steps 4–7) replaces the old "wipe the whole body" clear:
the cover/TOC skeleton survives generation, so a generated document keeps the
brand's front matter and only its body is replaced. Robust on templates with no
cover and/or no TOC, and idempotent (re-open shell each run).

**Idempotency is a hard requirement:** `generate()` twice yields a byte-identical file
(re-open shell each run; cover overwrites, never appends; clear-demo no-ops when clean;
`clone_part` uses deterministic sorted/content-hash rel-id + media-name allocation).

**Cloning structural/layout parts is forbidden** (layout deepcopy corrupts the package);
sections re-instantiate from existing layouts.

---

## 8. Unmapped-block policy

`policy.unmapped_block ∈ { strict → fail | degrade → nearest-role | passthrough → paragraph.default }`,
with `degraded=True` surfaced to QA — a block is **never** invented off-brand.

---

## 9. QA gate

One library (`qa/`), shared by all skills, one entrypoint `run_qa(target, profile, plan)`.
The gate carries an explicit **mode**: `gate_generated` | `verify_foreign`.

- **L0 deterministic** — always runs, no external binaries. Style allowlist, resolved-color
  palette adherence, residual placeholder/demo text, markdown literals, logo presence,
  language rules, WCAG contrast (solid resolved pairs only), table integrity, duplicate
  structure. `no_residual_template_text` runs **even when the profile had no demo region**.
- **L1 visual** — the Claude orchestrator (or a Task subagent) reads rendered PNGs. **No
  second model, no Bedrock, no boto3.**
- **L2 autonomous loop** — headless; audit → repair → re-render, bounded by a cost ceiling;
  on exhaustion with residual errors → `NEEDS_HUMAN` (never reports clean on exhaustion).

`--qa = fast | auto | deep`. `auto` is L0+L1 interactive (1 pass) or L0+L1+L2 headless.
`soffice`/`pdftoppm` may be absent → `doctor.py` degrades gracefully; **L0 always works**.

### Off-brand color rule (resolved-comparison)
Resolve every color through the theme + transform chain to a final hex **before**
comparing. Accept any hex reachable from a palette slot via a stored tint/shade. Treat
unresolved/`None` as **inherited — OK**. `ERROR` is reserved for a literal sRGB the
resolver itself wrote; in `verify_foreign` it downgrades to `WARNING` unless the palette
is declared closed.

---

## 10. Units

Geometry is stored as **EMU integers** with a mandatory `_emu` suffix on field names.
Conversions live in `common/units.py`:

`914400 EMU = 1 inch`, `360000 = 1 cm`, `12700 = 1 pt`, `635 = 1 dxa`. Font sizes in
`w:sz` are **half-points** (`sz="36"` = 18 pt); DrawingML `a:rPr sz` is **centipoints**
(`sz="1800"` = 18 pt).

---

## 11. Ordered structure & per-artifact usage (schema 1.1.0)

The profile records the template's **ordered top-level skeleton** in
`profile.structure`, and annotates every reusable artifact with where/how it is
used. Both are **additive** (a 1.0.0 profile may omit them; `validate()` never
rejects for their absence). Detection is brand-agnostic and multilingual —
grounded in style placement / OOXML evidence, never in brand-specific names.

### `structure`
```jsonc
"structure": {
  "ordered": true,                 // top-level region order must be respected
  "skeleton": [
    {"region": "cover", "order": 0, "role": "section.cover", "required": true,  "repeatable": false, "evidence": "..."},
    {"region": "toc",   "order": 1, "role": "section.toc",   "required": true,  "repeatable": false, "evidence": "..."},
    {"region": "body",  "order": 2, "role": "section.body",  "required": true,  "repeatable": true,  "freeform": true, "evidence": "..."}
  ]
}
```
Only regions actually present are listed. `region` ∈ `cover | toc | body`. The
`body` region is **freeform** (element order inside it is not prescribed).

**Region detection (evidence, on the lxml body):**
- **cover** — body content before the first TOC region or first Heading-1.
- **toc** — a block-level `w:sdt` with `w:docPartGallery='Table of Contents'`, a
  TOC/TOCHeading-styled paragraph, a `w:instrText` starting with `TOC`, or a
  heading whose text is a contents word in EN/IT/FR/DE/ES (`Contents`, `Sommario`,
  `Indice`, `Inhalt`, `Table des matières`, `Índice`, `Contenido`, …).
- **body** — everything after the TOC (or after the cover) up to the final body
  `sectPr`.

`anchors.toc.present` is a **real** detection result (never hardcoded).

### `usage` (on every role; on components when present)
```jsonc
"usage": {
  "scope": "cover" | "toc" | "body" | "anywhere",
  "placement": "structural" | "freeform",   // structural = part of the ordered skeleton
  "required": true | false,
  "order": <int or null>                     // skeleton position if structural, else null
}
```
Derived from the role **family**: `cover.*` → `cover`/`structural`/`required`/`0`;
`toc` → `toc`/`structural`/`required`/`1`; `heading`/`paragraph`/`list`/`callout`/
`table`/`quote`/`caption` → `body`/`freeform`/not-required/`null`. `structural`
artifacts must appear in their skeleton slot; `freeform` artifacts are used on
demand inside the body. `PROFILE.md` prints the `## Structure` skeleton and each
role's usage so a reader sees what to respect in order vs use on demand.

---

## 12. Licensing

- template-dna original code: **MIT** (every engine file carries an
  `SPDX-License-Identifier: MIT` header).
- Third-party proprietary Office helper scripts: **never vendored** — the OOXML
  engine is re-implemented from scratch (CI guard `tests/test_no_proprietary.py`).

Every file in the engine carries an `SPDX-License-Identifier` header.
