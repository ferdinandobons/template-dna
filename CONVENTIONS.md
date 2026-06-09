<!-- SPDX-License-Identifier: MIT -->
# brand-docs - Frozen Conventions

This is the **single source of truth** for the vocabulary of `brand-docs`. It is
owned by `scripts/brandkit/profile/schema.py`; this document is its human-readable
mirror. Every `SKILL.md`, command, and reference doc quotes these names **verbatim**.
If a name here disagrees with code, the code (the schema module) wins and this file
is the bug.

> **The one hard rule.** Brand-specific identifiers - style names, theme color
> tokens, font names, cover aliases, layout ids, chart-template ids - live in the
> Brand Profile and **nowhere else**. No IntermediateDocument block, and no writer,
> ever contains a literal style name, hex color, or font. The resolver is the only
> code that reads them. Off-brand output is impossible by construction.

---

## 1. The five verbs

Every skill (`brand-docx`, `brand-pptx`, `brand-xlsx`) implements the same contract:

| Verb | Input | Output |
|---|---|---|
| **extract** | a company `.docx`/`.pptx`/`.xlsx` template | a reusable **Brand Profile** |
| **comprehend** *(optional, model-driven)* | a saved profile + a model-authored `comprehension.json` | the profile with a validated, cached `comprehension` block |
| **verify** | a saved Brand Profile | QA findings + a verdict (the role map lives in `PROFILE.md`) |
| **generate** | content (free text or an IntermediateDocument) + a profile | a new on-brand document |
| **learn** *(optional, deterministic)* | a saved profile + its cross-run `generation_report.json` history | the profile with a distilled, ADVISORY `rules.overrides` lesson (live only with `--accept`; see §14) |

`verify` reports deterministic QA findings and a verdict; it does **not** render a
proof image, and the role-mapping table is written to `PROFILE.md` at extract time
(not re-emitted by verify). Pass `verify --accept` to mark a passing profile as
accepted (`verification.accepted = true`).

**`comprehend` is optional and model-driven** (schema 1.2.0). `generate` works on
the deterministic profile alone (CI/no-model); when a current comprehension is
present it additionally reconciles preserved cover/index structures with the new
content. The verb is realized as two CLI steps:

```bash
python scripts/brandkit/cli.py comprehend-input --name <brand>    # prints {facts, excerpt} for the model
python scripts/brandkit/cli.py comprehend --name <brand> --input comprehension.json  # the ONLY writer
```

`comprehend-input` surfaces the bounded, format-uniform bundle the model reasons
over (deterministic facts + a length-capped text excerpt; never raw OOXML).
`comprehend` is the **single writer** of the `comprehension` block: it
merge-validates the model's JSON **fail-closed** (schema shape + verbatim
membership of every load-bearing ref against the surfaced inventories), and on a
clean pass freezes it into `profile.json` with `status='present'`, stamping
`source_shell_sha256` from the live shell hash. On any finding it writes
`status='rejected'` with the findings and exits non-zero; the model retries. The
merge is **idempotent**: comprehend-twice yields a byte-identical `profile.json`.

The canonical engine entrypoint is `scripts/brandkit/cli.py`, run from the plugin
root (the directory containing `.claude-plugin/`):

```bash
python scripts/brandkit/cli.py extract --name <brand> --template <t>
```

Set `BRAND_DOCS_ROOT` to the plugin root to invoke the CLI from any working
directory (the skill `cli.py` shims honor it; if unset they walk up to the nearest
`.claude-plugin/`). `${CLAUDE_PLUGIN_ROOT}` never appears in author-facing SKILL.md.

---

## 2. The Brand Profile directory

A profile is a **self-contained, copyable directory**. All internal paths are
relative to the profile root, so it can move between stores unchanged.

```
brand-kit/<name>/
├─ profile.json          # the index: schema-versioned envelope          [always]
├─ PROFILE.md            # human-readable palette + role table            [always]
├─ template/
│  └─ shell.<ext>        # the byte-for-byte shell the generator opens FROM [always]
├─ provenance.sha256     # the shell hash (drift detection)               [always]
├─ assets/               # logos + manifest.json                          [optional]
├─ components/           # reusable fragments + index.json                [optional]
├─ sections/             # multi-page/slide units + index.json            [optional]
├─ specimens/            # opaque chart/SmartArt captures                 [optional]
├─ samples/              # source.png (template) + smoke.png (proof)      [optional]
└─ .cache/               # resolved-hex / render caches (safe to delete)  [optional]
```

The four `[always]` entries are written by every extract. The `[optional]` subdirs
are produced **only when the template supplies the corresponding artifacts** (e.g.
`assets/` only when logos are captured, `samples/` only when render tools rendered a
proof); a plain extract of a logo-free template yields just the four always-present
entries.

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
- **`kind`** - the format discriminator. Values: **`docx` | `pptx` | `xlsx`**.
  (Never `doc_type`.) Selects which `surface.*` block and which resolver types are legal.

### Schema version
- **`schema_version`** - semver string. Current: **`1.2.0`** (`SCHEMA_VERSION`).
  Minor bumps are **additive**: 1.1.0 added the optional `structure` section and
  the optional per-role `usage` object (see §11); 1.2.0 adds the optional
  top-level `comprehension` block (see §12) and opens region tokens (§12). Older
  profiles lacking any of these stay valid - `validate()` never rejects a profile
  for missing them.
- `$schema` id: `https://brand-docs/schema/profile-1.json` (`SCHEMA_ID`).
- The `comprehension` sub-block carries its own tag `COMPREHENSION_SCHEMA_VERSION`
  = **`comprehension-1`** so the model-facing contract can evolve independently.

### Role registry
- **`roles`** - the join table from semantic role id → concrete resolver descriptor.
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

### Status (capability maturity) - `Status`
**`robust` | `best_effort` | `stub`**. Maturity is *data, not structure*: the
schema holds charts/SmartArt from day one at `stub`; generation degrades gracefully
below `robust`.

### Severity (QA findings) - `Severity`
**`INFO` | `WARNING` | `ERROR`**.

### Verification status - `VerificationStatus`
**`passed` | `passed_with_warnings` | `failed` | `unverified`**. Stamped into
`verification.status`. `verify --accept` additionally sets `verification.accepted =
true` when the run passed (and is a no-op otherwise).

### Optional extractor-added top-level keys
Beyond the required envelope keys (`REQUIRED_TOP_KEYS` in `schema.py`), each
extractor stamps two **optional** top-level objects (not validated by `validate()`,
so 1.0.0 profiles may omit them):
- **`artifact_catalog`** - a broad, descriptive inventory of the template beyond the
  directly-generatable roles: OOXML/media parts, styles, sections/margins, and
  format-specific extras (docx: paragraph samples + table counts; pptx: layouts,
  masters, placeholder geometry, slide texts, slide size; xlsx: named ranges,
  formulas, sheet dimensions, tables, merged cells, row/column sizing, number
  formats). Read it to mimic a specific template piece; it is **not** a resolver.
- **`capabilities`** - what this profile can generate today, per kind. Descriptive
  only; the resolver and QA gate are the enforcement, not this key.

### Anchor kind - `AnchorKind`
**`sdt_anchored` | `placeholder` | `named_range` | `NONE`**. Anchor presence is
first-class: when no cover anchors exist, the kind is `NONE` (never a silently
empty list).

### Comprehension executor enums (the ONLY closed value sets the model may write)
Per **Ruling A**, only these four fields are closed enums - each value maps to a
real generator code branch. Every other comprehension field (`semantic_role`,
region names, `purpose`, `kind`, `evidence`, `generation_rules`) is an **open
advisory token** the generator never pattern-matches on.

| Field | Enum | Values |
|---|---|---|
| `comprehension.status` | `ComprehensionStatus` | **`present` \| `absent` \| `rejected`** |
| `cover_slots[*].fill_rule` | `FillRule` | **`in_place` \| `clear` \| `leave`** |
| `conventions.indexes[*].reconcile` | `Reconcile` | **`regenerate` \| `preserve` \| `clear`** |
| `demo_classification.regions[*].verdict` | `Verdict` | **`demo` \| `real` \| `mixed`** |

### Overflow capability - `OverflowCapability`
**`estimator` (pptx) | `cellfit` (xlsx) | `render` (docx) | `none`**. DOCX has **no
deterministic overflow estimator** - Word reflows; overflow is a render-time (L1/L2)
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
A run is `{"t": str, "b"?, "i"?, "u"?, "strike"?, "code"?, "sup"?, "sub"?, "link"?,
"color"?}`. `"b"/"i"/…` are booleans; `"link"` is a URL string; `"color"` is an
optional **palette token** (model-driven color, see §13). A bare `"text": "..."` on a
block is **sugar** normalized to `[{"t": "..."}]` on parse.

`"color"` is a verbatim key of `theme.palette` (a theme slot like `accent1`, or
`hex:RRGGBB`) - **never a literal color**. `normalize_runs` drops any hex-shaped or
`#`-bearing value structurally, so a literal color can never enter the
IntermediateDocument through a run (the one hard rule, enforced by construction).
The resolver maps the token to the captured color `ref`; an unknown token leaves the
run inherited (`color_token_unresolved` INFO).

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
| `divider` | - | brand divider |
| `pagebreak` | - | page (docx) / slide (pptx) break |

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

## 7. Generate - mandatory order

Enforced by one orchestrator; the author cannot reorder steps.

1. `expand_components` - component/section refs → primitives.
2. `validate_iid` - capability pre-flight (degradations reported **before** writing).
3. **Open the shell** (`Document(shell)`, never blank).
4. `clear_body_region` - remove **only** the freeform body region, **preserving**
   the ordered cover region and the TOC region (and the final `sectPr`). Boundary
   by region evidence, never index (§11). Never wipes the whole body.
5. `compose_cover` - fill the **preserved** cover anchors in place (sole author of
   the cover; never recreates it after a wipe).
6. `compose_body` - resolver-driven flat dispatch; new blocks are appended into the
   now-empty body region, immediately before `sectPr` (after the cover and TOC).
7. `refresh_toc` - refresh the preserved live TOC if present (mark fields dirty /
   `w:updateFields`) so Word recomputes it on open; never duplicate it. No-op when
   there is no TOC.
8. `finalize` - set `w:updateFields` + per-field `w:dirty`; save.

The order-aware clear (steps 4-7) replaces the old "wipe the whole body" clear:
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
with `degraded=True` surfaced to QA - a block is **never** invented off-brand.

---

## 9. QA gate

One library (`qa/`), shared by all skills, one entrypoint `run_qa(target, profile, plan)`.
The gate carries an explicit **mode**: `gate_generated` | `verify_foreign`.

- **L0 deterministic** - always runs, no external binaries. Style allowlist, resolved-color
  palette adherence, residual placeholder/demo text, markdown literals, logo presence,
  language rules, WCAG contrast (solid resolved pairs only), table integrity, duplicate
  structure. `no_residual_template_text` runs **even when the profile had no demo region**.
- **L1 visual** - the Claude orchestrator (or a Task subagent) reads rendered PNGs. **No
  second model, no Bedrock, no boto3.**
- **L2 autonomous loop** - headless; audit → repair → re-render, bounded by a cost ceiling;
  on exhaustion with residual errors → `NEEDS_HUMAN` (never reports clean on exhaustion).

`--qa = fast | auto | deep`. `auto` is L0+L1 interactive (1 pass) or L0+L1+L2 headless.
`soffice`/`pdftoppm` may be absent → `doctor.py` degrades gracefully; **L0 always works**.

### Off-brand color rule (resolved-comparison)
Resolve every color through the theme + transform chain to a final hex **before**
comparing. Accept any hex reachable from a palette slot via a stored tint/shade. Treat
unresolved/`None` as **inherited - OK**. `ERROR` is reserved for a literal sRGB the
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
rejects for their absence). Detection is brand-agnostic and multilingual -
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
- **cover** - body content before the first TOC region or first Heading-1.
- **toc** - a block-level `w:sdt` with `w:docPartGallery='Table of Contents'`, a
  TOC/TOCHeading-styled paragraph, a `w:instrText` starting with `TOC`, or a
  heading whose text is a contents word in EN/IT/FR/DE/ES (`Contents`, `Sommario`,
  `Indice`, `Inhalt`, `Table des matières`, `Índice`, `Contenido`, …).
- **body** - everything after the TOC (or after the cover) up to the final body
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

## 12. Comprehension block (schema 1.2.0, additive + optional)

The optional top-level **`comprehension`** block is the **single canonical sink**
for model output (**Ruling B**): one writer (the `comprehend` verb /
`profile/comprehension.py::merge`), one home. The pre-existing additive sinks
(`roles[*].usage`, `structure.skeleton` attrs, `anchors.*`) are **derived** from it
at merge time, never written independently. The block is always present and
`absent` by default (every extract stamps `empty_comprehension()`); `absent` ⇒
today's deterministic path. `validate()` never requires it.

```jsonc
"comprehension": {
  "schema_version": "comprehension-1",
  "status": "present",                 // present | absent | rejected  (closed enum)
  "generated_by": { "model": "...", "prompt_version": "...", "generated_at": "ISO-8601" },
  "source_shell_sha256": "<hex>",      // BOUND to provenance.shell.sha256 (cache key)
  "confidence": 0.0,                   // advisory; in [0,1]; gates DESTRUCTIVE acts

  "cover_slots": {                     // KEY = verbatim id from surface.<kind>.cover_anchors
    "<anchor_ref>": { "semantic_role": "<open token>", "purpose": "<free text>",
                      "binds_to": "<content slot key>", "demo_value": "<captured text>",
                      "fill_rule": "in_place" } },

  "conventions": {
    "indexes": [ { "index_ref": "<verbatim field id>", "kind": "<open token>",
                   "seq_id": "<\\c switch arg or null>",
                   "feeds_from_role_id": "<role id or null>",
                   "reconcile": "regenerate" } ],
    "sections": [ { "region_ref": "<verbatim region id>", "required": false, "repeatable": false } ] },

  "role_annotations": { "<role_id>": { "purpose": "<free text>", "generation_rules": "<free text>" } },

  "demo_classification": { "regions": [ { "region_ref": "<verbatim region id>",
                                          "verdict": "demo", "evidence": "<free text>" } ] }
}
```

### The fail-closed validation contract (`comprehension_targets_exist`)
Every **load-bearing reference** must be a verbatim id from the surfaced
deterministic inventories: `cover_slots[*]` keys ∈ cover-anchor inventory;
`indexes[*].index_ref` ∈ field inventory and `feeds_from_role_id` ∈ `roles`;
`sections[*].region_ref` / `demo_classification[*].region_ref` ∈ region inventory;
`role_annotations` keys ∈ `roles`. Unlike resolver-consistency (which no-ops on an
empty namespace), this check is **fail-closed**: a ref into an empty/absent
inventory is itself an **ERROR** (`status=rejected`), because it is the **sole**
gate for anchor/index/region refs. Enforcement is by **wiring the check into
`run_qa`** (severity-driven verdict); its id and `no_net_structure_loss` are listed
in `DEFAULT_L0_INVARIANTS`.

### Surfaced inventories (format-uniform)
`profile/comprehension.py::surface_inventories(profile)` returns the SINGLE
definition both `comprehend-input` and the gate use:

```jsonc
{ "cover_anchors": [ {"id": "<anchor_ref>", ...}, ... ],
  "fields":        [ {"id": "<index_ref>", ...}, ... ],
  "regions":       [ {"id": "<region_ref>", ...}, ... ],
  "roles":         [ "<role_id>", ... ] }
```

Read from `surface.<kind>.{cover_anchors,fields,regions}` (enriched per format on
its own fact milestone; legally empty until then) and the concrete role-id list.

### Cache binding (sha-bound, in code)
A comprehension counts as **present** only when `status='present'` **and**
`source_shell_sha256 == provenance.shell.sha256` (`store.comprehension_is_present`).
A re-extract rebuilds the profile with a fresh `absent` block, so a drifted shell
never reuses a stale comprehension. Generation is **idempotent**: comprehension is
frozen at merge and never re-invoked at generate time.

### Open region tokens (replaces frozen per-format region word-lists)
`structure.skeleton[*].region` is an **open token** (validated for syntax like a
role id, never against a frozen word-list - that would re-commit the lexicon sin).
The only things generation branches on are the boolean attributes
`STRUCTURE_REGION_ATTRS = (freeform, demo, ordered, required)` - the same four for
every format.

---

## 13. Model-driven color (`theme.palette` + `palette_annotations`)

Color follows the same **"the model proposes, the deterministic disposes"** pattern
as comprehension, layered additively on the deterministic appearance capture
(`theme.text.body.color`, `role.appearance.color` - both untouched and
byte-identical for existing profiles). Three pieces:

### `theme.palette` (deterministic UNDERSTAND - additive, optional)
A template-derived map keyed by a **theme slot token** (`accent1`..`accent6` /
`dk1` / `dk2` / `lt1` / `lt2` / `hlink` / `folHlink`) or `hex:RRGGBB` for an
observed off-theme run color. Each entry:

```jsonc
"theme": { "palette": {
  "<slot or hex:RRGGBB>": {
    "ref": { "kind": "theme", "theme": "accent1" },   // OR { "kind": "hex", "hex": "RRGGBB" }
    "provenance": [ { "where": "<closed>", "detail": "<str>" } ],  // sorted by (where, detail)
    "frequency": "dominant" | "accent" | "rare",      // COARSE bucket, never raw counts
    "name": null, "purpose": null, "use_when": null   // model-only; null in the deterministic path
  } } }
```

The closed provenance `where` vocabulary is **exactly**
`palette_role | role.appearance | run.color | link.color` (`palette_role` is the
only NON-authoritative source). Capture is **model-free, deterministic, and
byte-identical on re-extract**; a template with no observed color leaves an empty
`{}` palette. The schema validates the shape only (`_validate_palette`); there is
**no `SCHEMA_VERSION` bump** (it is an additive optional key, its documented default
`{}`). All three formats capture the same way (format-uniform): seed the palette
from the parsed theme slots, then fold the DIRECT run/cell colors the template
carries (low-floor accent aggregation + per-role + link colors) on top - docx walks
runs, pptx walks slide/shape runs, xlsx walks styled cell fonts. A template with no
direct color keeps just the seeded slots.

### `comprehension.palette_annotations` (model-driven, derived sink)
The model **NAMES** each palette color, keyed by its palette id:

```jsonc
"comprehension": { "palette_annotations": {
  "<palette_key>": { "name": "...", "purpose": "...", "use_when": "...",
                     "semantic_role": "..." } } }
```

Every key is **fail-closed** against the surfaced `palette` inventory (a key into an
empty/absent inventory is an **ERROR**, the same rule as anchor/index/region refs),
enforced by `check_membership` and the gate's `check_color_token_targets` (wired
right after `check_comprehension_targets`). On a clean merge `_derive_palette_annotations`
mirrors `name`/`purpose`/`use_when`/`semantic_role` onto `theme.palette[key]`. The
model **NEVER** writes `ref`/a color - only names.

### The run `color` token (APPLY)
A run's optional `color` (§6) is a palette token the resolver maps to
`theme.palette[token]['ref']` and applies. Every applied palette `ref` is
re-validated against the shell **fail-closed** by `check_appearance_targets`, which
is **format-neutral**: a per-kind collector reduces each shell (docx/pptx/xlsx) to
the same fact set, then membership is checked uniformly (hex vs the shell's theme
palette UNION its own observed run/cell hexes; theme token vs the parsed `clrScheme`
slot). Guarantees: no literal hex in any IDoc/writer (structural rejection in
`normalize_runs`); the model can NAME but never author a real color; the
deterministic no-model path is byte-identical; re-runs are idempotent; everything is
template-derived (no hardcoded colors/names/language). Applied on all three formats
(the appearance/color apply layer is shared via `common/appearance.py`).

---

## 14. Learn-from-errors (`generation_report.json` + `rules.overrides` + `learn`)

The deterministic feedback loop (Cluster B core; the model-proposed phase is a later
increment). Three additive layers, format-uniform, schema stays **1.2.0**:

### The persisted run report (`qa/report.py`)
Every `generate` writes `generation_report.json` into the same `<output>.visual`
side-artifact dir as the visual manifest: the QA `verdict` + `findings` **verbatim**
(never reordered/deduped) + `shell_sha256` / `content_sha256` (canonical `to_dict()`
JSON of the parsed input) / `output_sha256` + `generated_at` (the ONLY volatile
field, read by no generator path - document bytes stay identical across runs).
Degrade-to-no-op: a failed write never raises into the gate and never flips a
verdict. GENERATE-only (`verify` writes no report - a hash-less row would pollute
the history). QA producers carry the structured `location` pointer (a role id for
`resolver_targets_exist`/`style_fallback`, the matched marker for
`no_residual_template_text`) - the binding pointer the loop keys on, never the
brand-bearing `message`.

### Cross-run regression findings
Before writing the new report, prior SAME-shell reports are discovered (partitioned
strictly by `shell_sha256` - a re-extract starts a fresh history) and the run's
findings are diffed against them on the `(check, location)` key:
`regression.recurred` (INFO, carries the `recurred_runs` count the `learn`
threshold gates on) and `regression.reintroduced` (WARNING - it came back after a
clean run). ADVISORY only: neither is in `DEFAULT_L0_INVARIANTS`, neither can flip
a verdict or the CLI return code.

### The `learn` verb + the `rules.overrides` block (`profile/overrides.py`)
`learn` distills UNAMBIGUOUS findings (`LEARNABLE_CHECKS`) that recurred across
`>= 2` same-shell runs into a closed-vocab lesson - `reroute_role` (a stub role to
a healthy SAME-family sibling), `number_format` (swap to a mask in
`surface.<kind>.number_formats`), `register_demo_clear` (a captured demo string) -
and routes it through the SINGLE `merge_overrides` sink: shape validation +
fail-closed membership (reject-never-skip, a pointer into an empty inventory is an
error) + an acyclic reroute-graph proof, ALL-OR-NOTHING (one unbound pointer
rejects the whole proposal). The block mirrors the comprehension freeze contract:
`status` / `source_shell_sha256` (== `provenance.shell.sha256`, reset by
re-extract) / `confidence` / flat per-entry `provenance`.

Consume + guarantees: the resolver applies a lesson only as a **LAST-RESORT on a
genuine stub** (never on a healthy typed resolve), single-hop with the role id
pinned to the REQUESTED role (the heading body-default exclusion still keys on the
original); a reroute reuses the target role's existing shell-proven resolver
verbatim through the legal-type gate, so an override can never inject a
style/font/hex. `check_override_targets` (gate-wired, `override_targets_exist`)
re-proves every lesson against the live shell at verify. Lessons are ADVISORY
until an explicit `learn --accept` (mirroring `verify --accept`); with no accepted
lesson the resolver takes zero new branches and generation is **byte-identical**.

---

## 15. Licensing

- brand-docs original code: **MIT** (every engine file carries an
  `SPDX-License-Identifier: MIT` header).
- Third-party proprietary Office helper scripts: **never vendored** - the OOXML
  engine is re-implemented from scratch (CI guard `tests/test_no_proprietary.py`).

Every file in the engine carries an `SPDX-License-Identifier` header.
