# Changelog

All notable changes to BrandDocs are documented in this file.

## [Unreleased]

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

[0.2.0]: https://github.com/ferdinandobons/brand-docs/releases/tag/v0.2.0
[0.1.0]: https://github.com/ferdinandobons/brand-docs/releases/tag/v0.1.0
