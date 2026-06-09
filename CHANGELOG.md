# Changelog

All notable changes to BrandDocs are documented in this file.

## [Unreleased]

### Added

- **Cross-format appearance vertical (pptx + xlsx now capture, apply, and verify
  brand typography/color).** The role typography (font/size/color) and model-driven
  run color that shipped for docx in 0.7.0 now work on PowerPoint and Excel too. The
  docx capture/apply helpers were extracted into shared `common/typography.py` +
  `common/appearance.py` (docx output stays byte-identical, guarded by a frozen-hash
  anchor test); pptx and xlsx gained capture adapters that populate the same
  `role.appearance` + `theme.palette` shape and a per-format apply backend that
  consumes `op.appearance` + the resolver's color, set-only-when-unset (an inherited
  theme value is never clobbered). `check_appearance_targets` is now format-neutral:
  a per-kind shell-fact collector re-validates every applied font/size/color against
  each shell's own facts, fail-closed. The model's `palette_annotations` naming is
  now load-bearing on all three formats. Schema stays 1.2.0 (additive); a profile
  with no captured appearance generates byte-identically as before.
  - **Universal, not tuned:** capture is a dominance statistic over the template's
    own runs/cells; nothing is hardcoded. Excel's cell theme-color index (which swaps
    the first dark/light pairs vs clrScheme order) is mapped correctly so a default
    `theme=1` text color is captured as `dk1`, not `lt1`.
  - **Still docx-only:** caption-index regeneration (the `SEQ` field machinery).
- **Learn-from-errors deterministic core (the profile learns from its own QA
  findings).** Three additive layers, format-uniform across docx/pptx/xlsx:
  - **Persisted `generation_report.json`** (new `qa/report.py`): every `generate`
    writes a side artifact next to the output (the QA verdict + findings verbatim +
    shell/content/output sha256 + a timestamp). Degrade-to-no-op (a failed write can
    never flip a verdict), generate-only, and the timestamp lives only in the JSON -
    generated document bytes stay identical across runs.
  - **Cross-run regression findings**: a new run's findings are diffed against prior
    SAME-shell reports and `regression.recurred` / `regression.reintroduced` are
    folded into the QA report - keyed strictly on `(check, location)`, never the
    brand-bearing message; advisory (INFO/WARNING), never flips a verdict.
  - **A `learn` verb + shell-bound `rules.overrides`** (new `profile/overrides.py`):
    deterministic distillation of unambiguous recurring findings into a closed-vocab
    lesson (`reroute_role` to a healthy same-family sibling, `number_format` swap to
    a shell-backed mask, `register_demo_clear` of a captured demo string), written
    through a single all-or-nothing `merge_overrides` sink (fail-closed membership +
    an acyclic reroute-graph proof). The resolver consumes a lesson only as a
    LAST-RESORT on a genuine stub (never on a healthy resolve), pinned to the
    requested role id; `check_override_targets` re-proves every lesson against the
    live shell at verify (ERROR on a missing target, reject-on-empty). Lessons are
    frozen to the shell sha (a re-extract resets them) and stay ADVISORY until an
    explicit `learn --accept` - with no accepted lesson, generation is
    byte-identical to today. QA producers now carry the structured `location`
    pointer (role id / demo marker) so lessons distill from real run history.
    The model-proposed phase (B4) is a later increment.

## [0.7.0] - 2026-06-09

Brand-fidelity release: the engine now follows a template's real per-run typography
(font, size, color), applies model-named brand colors, fills content-control covers
faithfully, and regenerates caption indexes from the new content. Schema stays 1.2.0
(additive); profiles from earlier versions keep working unchanged.

### Added

- **Brand typography capture (font family + size + color).** Extraction now captures
  the template's dominant DIRECT run typography - the real visible font, font size
  and run color a designed template applies per-run (e.g. Roboto 16pt) rather than
  via the named styles or theme - into the reserved `role.appearance` field and the
  additive `theme.fonts.body` (font/size) and `theme.text.body` (color). Generation
  applies all three as direct run formatting through the resolver (paragraphs, list
  items, captions, quotes, callouts, table cells, and hyperlink runs), so a generated
  document renders in the template's real font, size and color instead of falling
  back to the `docDefaults` (typically Arial 11pt black). `_extract_theme` also reads
  the real `theme1.xml` major/minor fonts and `docDefaults`.
  - **Universal and deterministic.** Each axis is captured by dominant sampling over
    ALL runs (a run with no explicit value votes "inherit"), so a value is applied
    only when it dominates the whole document - an accent color present on a few runs
    never becomes the body color. Nothing is hardcoded to any template; a document
    with no dominant value (and every pre-capture profile) behaves exactly as before.
  - **Brand guarantee, fail-closed.** Every applied font/size/color is re-validated
    against what the shell proves it contains (`appearance_targets_exist`): fonts vs
    the fontTable + theme, sizes vs the template's own `w:sz` values, colors vs the
    theme palette + the template's own `w:color` values. The body size/color default
    never flows onto heading roles (their style's intrinsic size/color is preserved).
  - Deferred (see [ROADMAP](documentation/ROADMAP.md) section 1): per-word accents,
    cover-layout reconstruction, and heading typography when the template fakes
    headings in the body style.
- **Model-driven brand color (run color palette tokens).** The profile now carries a
  `theme.palette` of the template's brand colors (theme slots like `accent1`, plus
  off-theme `hex:RRGGBB`) with usage provenance; an IntermediateDocument run can carry
  a semantic `color` token (a palette key, never a literal hex) that the resolver maps
  to the brand color and the generator applies - rendering a theme color as
  `w:color@w:val=<hex>` + `w:themeColor=<slot>` so it shows in both LibreOffice and
  Word. The model NAMES each color's purpose via `comprehension.palette_annotations`;
  it never authors a color (fail-closed: a token into an empty/absent palette errors).
- **Caption-index regeneration (list of tables / figures).** A kept caption index
  (`TOC \c "<seq>"`) is now repopulated from the new content: each captioned table or
  figure emits a real Word `SEQ` field, and the index's visible cache is rebuilt from
  the emitted captions (so a headless render shows the new entries, not the template's
  stale ones). The model maps a caption kind to its index via the new closed
  `comprehension.conventions.indexes[*].caption_target` (`table|figure`); brand- and
  language-agnostic (the seq label is the template's own, read from the profile).
- **Cover fidelity for content-control covers.** A cover content control bound to a
  core property (e.g. the subject repeated in the page header) now has its cached run
  text refreshed across the body and every header/footer, so a headless render shows
  the filled value, not the template prompt. The cover role-style re-assertion is
  gated to bare placeholder controls, so a real author-formatted slot keeps its own
  formatting (a builtin `Subtitle` whose color is a near-white tint no longer blanks a
  filled cover subtitle).

### Fixed

- Colored hyperlink runs now write `w:color` at the schema-correct `CT_RPr` position
  (before `w:u` / `w:sz` / `w:vertAlign`) instead of appending it last, so an
  underlined colored link is conformant OOXML; a theme-token link with no resolved hex
  emits `themeColor` only (no stray `w:val="000000"` that would paint it black).

## [0.6.2] - 2026-06-09

A robustness fix for real-world templates. No schema, profile or output change for
well-formed templates; profiles from earlier versions keep working unchanged.

### Fixed

- DOCX `extract` and `generate` no longer crash on templates whose section
  measures are non-integer twips (e.g. `1440.0000000000002` in `w:pgMar` /
  `w:pgSz`, emitted by some editors). python-docx parses measures with `int()`
  and raises the moment any code touches the value, including its own internals
  (`Document.add_table` derives the table width from `section.left_margin`). The
  shell's section measures are now sanitized in place (sub-twip rounding, so the
  page geometry is visually identical) and every section-length read tolerates the
  malformed value via a raw-twips fallback.

### Tests

- `MalformedSectionMeasureTest`: the section-length helpers tolerate the bad
  value, `sanitize_section_measures` repairs it in place, and a generate with a
  table block (the exact crash path) survives a malformed margin.

## [0.6.1] - 2026-06-09

A clean finished checkpoint: the remaining LOW items from the v0.6.0 review are
closed and dead parameters removed. No behavior change beyond the fixes; profiles
from earlier versions keep working unchanged.

### Fixed

- The QA gate now verifies a `number_format` role's mask against the SHELL's
  actually-used formats (the shell-backed peer of the schema's intra-profile
  check), so a fabricated/hand-edited mask is caught at verify time, never applied.
- A `table`/`KPI`/`SmartArt` block rendered against a profile with no matching role
  no longer falls back to the default style SILENTLY: it surfaces an INFO
  `style_fallback` so the missing brand style is visible in QA (the content still
  renders).

### Changed

- Removed dead parameters (no behavior change): `run_qa`'s never-read `mode`, the
  docx `_write_block` `profile` arg, and the xlsx `_resolve_named_target` `findings`
  arg.

### Tests

- Closed the review's remaining coverage gaps: strict-mode promotion of
  `visual.ocr_degraded` to an error, `_ocr_png` timeout/OSError degradation, the
  Quick-Look-absent render degrade, the QA-gate fabricated-mask rejection, and the
  `style_fallback` path.

## [0.6.0] - 2026-06-09

Visual QA works by default, a whole-project code + quality review (multi-agent,
adversarially verified) hardens security/correctness/determinism, and the docs are
re-synced to the shipped engine. Brand Profiles from 0.1.x-0.5.0 keep working
unchanged.

### Added

- **Visual QA runs by default and installs in one step.** The renderer probe is now
  FUNCTIONAL-first: a LibreOffice that actually renders is usable even if its macOS
  code signature was knocked loose by an update/quarantine removal (no re-sign
  needed), so the visual gate is no longer falsely disabled. `scripts/setup_visual_qa.sh`
  auto-detects the platform package manager and installs LibreOffice + Poppler
  (+ optional Tesseract); the README/INSTALLATION reframe visual QA as a standard,
  on-by-default part of the QA step.
- An aggregate OCR time budget so a many-page document cannot turn a deep/strict
  gate into an `N x timeout` hang.

### Fixed

- **Security - Excel formula injection:** author content starting with `=` became a
  LIVE formula (`=WEBSERVICE`/`=HYPERLINK`/DDE) in the generated workbook, breaking
  the "formulas live only in the shell" invariant. It is now neutralized to a TEXT
  cell (verbatim, never executed) and surfaced; the QA gate also fails closed on any
  output formula the shell did not have.
- **Security - hyperlink scheme allowlist:** docx/pptx now refuse `file:`/
  `javascript:`/`data:`/`smb:` link targets (the text is kept), so untrusted content
  cannot wire a hostile link into an on-brand file.
- **Robustness:** component/section expansion has a node-count budget (runaway
  fan-out fails loud, never hangs/OOMs); a pathologically deep list IID raises the
  contracted `IIDParseError` instead of `RecursionError`; the OCR step tolerates
  non-UTF-8 tesseract output instead of crashing the whole run.
- **Determinism:** xlsx generation is byte-idempotent even when the shell's
  `core.xml` lacks `dcterms:created` (openpyxl would otherwise fabricate a
  wall-clock one); a number-format mask is no longer applied to a preserved-formula
  / merged cell; `profile.json` is written with sorted keys.
- **Number-format classification:** elapsed-time masks (`[h]:mm`) classify as time;
  accounting masks with bracketed currency + padding idioms classify as accounting.
- **Visual QA:** the Quick Look fallback accepts only the expected thumbnail name
  (never stages a stale/unrelated PNG as a bogus render).

### Changed

- Removed verified-dead helpers and corrected stale docstrings/docs: the docx TOC
  and the pptx charts/SmartArt/KPI/images are native now (not "deferred"), and
  `number_format`/`named_range` are first-class resolver types.

## [0.5.0] - 2026-06-08

Completes the block-type matrix: Word now renders **every** block type natively
(the `toc` block was the last to degrade), and Excel resolves a cell's semantic
number format to the template's own mask. Brand Profiles from 0.1.x-0.4.0 keep
working unchanged; an absent comprehension is still the deterministic path.

### Added

- **Word `toc` blocks render natively.** The `toc` block was the last block type
  that degraded instead of rendering; it now authors a real, updateable outline
  table-of-contents field. If the shell already carries an outline TOC the block
  defers to it (refreshed in place, never a duplicate); otherwise it authors a
  native outline TOC field at the block's position, with its visible cache filled
  from the generated headings and the field marked dirty + `updateFields` so Word
  rebuilds it on open. It defers only to an OUTLINE TOC, so a shell shipping only a
  table-of-figures still gets its requested table of contents. Byte-idempotent.
  Every Word block type now renders natively.
- **Excel resolves a cell's semantic number format.** The `number_format` resolver
  type is now wired end-to-end (previously reserved/staged). `extract` classifies
  the template's own number-format masks into brand-agnostic families
  (`currency`/`percent`/`date`/`datetime`/`time`/`accounting`/`decimal`/`integer`/
  `text`/`scientific`) and emits a `number.<family>` role bound to the template's
  VERBATIM mask. A `GridDocument` names the intent via `formats: {name: family}`
  (keyed by the same named-range/cell vocabulary as `cells`/`regions`); generation
  resolves `number.<family>` and applies the template's own mask to the filled
  cell(s). A format is never fabricated: an intent the template does not carry
  degrades loudly (`number_format_degraded`) and leaves the existing format, and a
  resolver whose mask is not in the template's surfaced formats is rejected at
  validation.

## [0.4.0] - 2026-06-08

Model-driven reusable-fragment population completes the comprehension vertical,
plus native SmartArt and PowerPoint cell merges. Brand Profiles from 0.1.x-0.3.0
keep working unchanged; an absent comprehension is still the deterministic path.

### Added

- The reusable-fragment registries (`components` / `sections`) are now
  **populated by the model through the fail-closed `comprehend` boundary** - no
  hardcoded catalog. A `comprehend` proposal may carry a `fragments` list (each
  `{ref, kind: component|section, purpose?, blocks}`); `merge` validates it
  fail-closed (every block parses as a known IID primitive; a nested
  `component`/`section` ref must resolve to another fragment proposed in the same
  comprehension; `(kind, ref)` is unique; cyclic references are rejected) and, on
  a clean pass, DERIVES the entries into `profile['components']`/`['sections']` -
  the registries `expand_components` already inlines. The registries are rebuilt
  deterministically from the (single-source) comprehension on every clean merge,
  so re-comprehending the same proposal is byte-idempotent and a fragment-less
  comprehension leaves them empty. A fragment is presentation-free IID, so a
  validated proposal resolves through the same brand chokepoint as inline content
  and cannot widen the brand guarantee. The `comprehend-input` bundle now surfaces
  advisory `fragment_candidates` (recurring layouts/styles derived cheaply from
  the existing `artifact_catalog`, possibly empty) as a hint; the model does the
  semantic detection. The shared comprehension prompt gains a fifth question + a
  worked example, byte-identical across the three skills.
- Fragment **`slots` are now substituted**: a `{{name}}` token in a fragment's
  template text is filled from the referencing `component`/`section` block's
  `slots` at expansion time (an unfilled or null token resolves to the empty
  string and is never leaked). Substitution deep-copies, so it never mutates the
  shared profile registry.
- PowerPoint tables now honor `colspan`/`rowspan`: a spanning cell (e.g. a banner
  across columns) merges the covered grid cells in the native table, matching the
  docx table writer. Previously pptx rendered a full ungrouped grid while docx
  merged.
- `smartart` blocks are now authored as a NATIVE, on-brand diagram instead of
  degrading to text - completing native rendering for every block type. On **pptx**
  the diagram is real autoshapes (a chevron row for a process/flow, a stacked
  rounded-box list otherwise), inheriting the deck theme's accent fill. On **docx**
  it is a brand-styled table (a process is a single row, one cell per step; a list a
  single column, one row per node). A node's `children` are preserved (sub-lines in
  the pptx shape, inlined in the docx cell); an empty diagram degrades loudly. The
  diagram is rendered (not editable SmartArt with the layout-engine parts).

## [0.3.0] - 2026-06-08

Native charts across all three formats, a deterministic-cover subtitle fill, a
single source of truth for component-survival, and a round of correctness +
quality fixes from a multi-agent code review. Brand Profiles from 0.1.x/0.2.0 keep
working unchanged.

### Added

- A `chart` block is now authored as a **native chart on both Word and PowerPoint**
  (a real DrawingML `c:chart`: an inline `w:drawing` on docx, a `graphicFrame` on
  pptx), no longer flattened to body text. `bar`/`column`/`barh`/`line`/`area`/`pie`/
  `doughnut` map to the matching chart type (unknown -> clustered column, surfaced as
  INFO); series/categories/title come from the block, and the chart inherits the
  document/deck theme's accent colors so it is on-brand by construction. A
  multi-series pie/doughnut surfaces a truncation WARNING; an empty/all-non-numeric
  chart degrades loudly. A shared `ooxml.chart` builds the docx chart with INLINE
  cached data (no embedded workbook) and is the single data gate both formats use;
  the pptx data workbook's wall-clock timestamps are normalized by
  `repack_fixed_timestamps` (now recursive over nested OOXML packages), so generation
  stays byte-idempotent on both.
- Excel charts complete the set: a `GridDocument.charts` entry (`{sheet?, type,
  title?, anchor, data, categories?, data_titles?}`) is authored as a NATIVE
  openpyxl chart that REFERENCES the workbook's own cell ranges (the grid model is
  range-based, so the data lives in the sheet). Same type map / unknown-type INFO
  fallback / loud degrade contract; the chart inherits the workbook theme, and
  generation stays byte-idempotent. All three formats now author native charts.
- Word: the deterministic cover fill (comprehension absent) now also places
  the authored **subtitle** into the cover slot identified by its resolved
  `cover.subtitle` style - correct-by-style, never guessed from the template's
  placeholder text - so the output no longer shows the template's stale demo
  subtitle. Role inference resolves `cover.subtitle` from a custom
  subtitle-named style (preferred) or the builtin `Subtitle`; a multilingual
  `subtitle` name-token family backs it. Templates whose subtitle is a databound
  SDT keep working via core-property sync; extra cover fields (date/id/author)
  remain the comprehension path's job and are still surfaced as unplaced.

### Fixed

- General correctness + quality review (multi-agent, adversarially verified):
  - **Word/PPTX tables no longer drop multi-run column headers.** A header cell
    authored with rich runs (e.g. plain text + a bold unit) kept only its first
    run through `Table.from_dict`; every run is now preserved (and the loose
    `{"runs": [...]}` / `{"text": "..."}` / run-list / string shapes a body cell
    accepts are accepted for columns too).
  - **PowerPoint body text on a placeholderless layout degrades loudly** instead
    of vanishing silently (a `block_degraded` WARNING is now recorded).
  - Word hyperlink runs with empty text no longer emit an empty `w:hyperlink`.
  - `comprehend`'s skeleton demo/required annotation matched the wrong key
    (always None); it now keys on the region id, so the annotation actually applies.
  - `extract` wraps its work in error handling (clean `ERROR extract: ...` + exit 1,
    matching `generate`) and defaults `--scope` to `auto` like the other commands.
  - Idempotency: a non-UTF-8 `core.xml` no longer crashes the timestamp pin; the
    nested-package dcterms regex uses `[^<]*` + `count=1` so it cannot cross a tag
    boundary on malformed XML.
- Quality cleanups: consolidated the duplicate docx `_apply_*_style` helpers,
  removed the unused `safe_filename` utility, dropped a redundant `set()` in
  `has_part`, removed a duplicate `visual.no_pages` finding, and stopped running
  `check_profile` twice in the QA gate.

### Changed

- `component_survival` now has a single source of truth. The pptx generator's
  own pre-reconcile, drop-to-zero variant was removed; the QA gate's
  `check_component_survival` (which re-reads the shell and output independently,
  for all three formats, on any count decrease) is the sole emitter. This ends the
  duplicate, differently-worded `component_survival` findings a pptx run produced.

## [0.2.0] - 2026-06-08

Hardening release: correctness fixes across all three formats, a real
`doctor` preflight gate, a uniform destructive-action policy, and a
3.10+ / CI matrix toolchain. Brand Profiles from 0.1.0 keep working
unchanged (schema 1.x is read as-is).

### Added

- `doctor` is now a preflight gate, not just a report: it exits nonzero
  when a required Python dependency is missing, so a broken environment
  fails before generation instead of mid-run.
- `doctor --json` prints the machine-readable dependency/renderer probe;
  `doctor --fast` skips the slow LibreOffice render probes (visual QA is
  marked "not probed") for a quick check.
- PowerPoint now expands reusable `components` / `sections` fragments,
  matching the Word vertical: a profile-defined fragment inlines
  identically across formats, fail-closed (an undefined reference raises
  rather than silently dropping content).
- Schema: a profile whose major version is newer than this build supports
  now fails with a single clear "upgrade brand-docs or re-extract" message
  instead of a pile of confusing per-field errors, and a forward
  `migrate()` seam is in place for future schema growth.

### Changed

- The destructive-action confidence floor (>= 0.5) is now applied
  uniformly at every reconcile site across Word, PowerPoint, and Excel
  from a single source of truth, so a low-confidence delete is downgraded
  to keep-with-warning consistently (previously only some sites enforced
  it).
- PowerPoint now honors the profile-resolved body placeholder index
  (`ph_idx`) when filling body content, instead of always assuming the
  first placeholder.
- Python floor is now 3.10+ (declared in `pyproject.toml`
  `requires-python`). CI runs a 3.10 / 3.11 / 3.12 matrix, a `ruff` lint +
  format gate, and a real-render lane that installs LibreOffice + poppler
  and actually executes the visual end-to-end tests (a broken renderer
  fails the lane).
- Reference docs (block catalog, profile schema) are synced to the engine
  and guarded by a freshness test; a fixture-determinism guard keeps the
  test fixtures byte-stable.
- Internal: a shared `ooxml/` layer now centralizes qualified-name
  helpers, the complex-field walk, and ZIP byte-idempotency, removing
  duplicated boilerplate from the format generators.
- README slimmed and repositioned; the GitHub Pages site is split from the
  human-facing `documentation/`; em dashes removed across skills and
  metadata.

### Fixed

- Excel no longer crashes when a named region's first row straddles a
  merged banner: writes to a merged-slave cell are guarded (only the merge
  anchor is writable) and the skipped value surfaces as a `block_degraded`
  warning instead of raising or being lost.
- PowerPoint no longer silently drops Component / Section / Toc / Divider
  blocks it has no native writer for: each now emits a visible
  `block_degraded` warning, honoring the "never drop content silently"
  invariant.
- Word generation is now byte-deterministic: re-running the generator on
  the same inputs produces an identical file (ZIP entry timestamps are
  pinned), matching the existing PowerPoint and Excel behavior.

## [0.1.0] - 2026-06-07

Initial public alpha release.

### Added

- `brand-docx`, `brand-pptx`, and `brand-xlsx` skills for same-format
  generation from Word, PowerPoint, and Excel templates.
- Shared Brand Profile engine for extracting template styles, structure,
  layouts, named ranges, formulas, media, and reusable OOXML artifacts.
- Deterministic QA for resolver targets, allowed styles/layouts/ranges,
  residual template text, table integrity, formula preservation, language
  checks, and artifact drift.
- Optional visual QA with renderer dependency preflight, visual manifests,
  degraded-mode reporting, OCR support, and strict QA mode.
- Example templates and regression evals covering DOCX, PPTX, and XLSX flows.
- GitHub Pages SEO entry point, `llms.txt`, `robots.txt`, `sitemap.xml`, and
  directory-submission guide.

### Notes

- BrandDocs is alpha software. DOCX is the reference vertical; PPTX and XLSX
  share the same engine and are intentionally catching up through the eval
  suite and visual repair workflow.

[0.7.0]: https://github.com/ferdinandobons/brand-docs/releases/tag/v0.7.0
[0.6.2]: https://github.com/ferdinandobons/brand-docs/releases/tag/v0.6.2
[0.6.1]: https://github.com/ferdinandobons/brand-docs/releases/tag/v0.6.1
[0.6.0]: https://github.com/ferdinandobons/brand-docs/releases/tag/v0.6.0
[0.5.0]: https://github.com/ferdinandobons/brand-docs/releases/tag/v0.5.0
[0.4.0]: https://github.com/ferdinandobons/brand-docs/releases/tag/v0.4.0
[0.3.0]: https://github.com/ferdinandobons/brand-docs/releases/tag/v0.3.0
[0.2.0]: https://github.com/ferdinandobons/brand-docs/releases/tag/v0.2.0
[0.1.0]: https://github.com/ferdinandobons/brand-docs/releases/tag/v0.1.0
