# Changelog

All notable changes to BrandDocs are documented in this file.

## [0.9.0] - 2026-06-10

The brand-intelligence release: derived indexes that never lie (rich TOC
cache, empty rebuilds), multi-template profiles (`extract --blend`), the
cross-template drift report (`compare-profiles`), a guided authoring surface,
a local-corpus fidelity benchmark, and a 31-finding general-review hardening
wave on top.

### Fixed

- **General-review wave (multi-agent, 31 confirmed findings).** Engine: the
  heading-less generation now collapses a bare multi-paragraph outline TOC's
  FULL span to an empty field (the plain writer used to leave the template's
  demo entries and orphan the end fldChar - a regression of the rich-TOC
  delta); pptx and xlsx extraction now parse the theme `a:fontScheme` latin
  typefaces instead of hardcoding `latin=null` (the docx extractor already
  did); `schema.validate` returns problems instead of raising on a
  hand-edited blend ledger with mixed-type entries; profile JSON/sidecar
  writes are atomic (same-directory temp + `os.replace`), making the blend's
  byte-identity-on-failure claim true under ENOSPC/kill; OOXML `unpack`/
  `read_part` reject decompression bombs (per-part and per-package declared
  inflated-size floors); the corpus runner exits 1 on verify exceptions, not
  just clean failures. Examples: the three builders now consume ONE complete
  12-slot brand clrScheme from `_brandlib` (docx/pptx/xlsx had drifted apart
  on five supporting slots - caught live by `compare-profiles` - and the
  shipped pair now reports zero theme drift, pinned by a strengthened CLI
  test that replaced a tautological assertion). Docs: README test count and
  blending/compare teaching, CONVENTIONS gains the engine verb surface and a
  §16 blend/compare vocabulary, the Pages site's broken README anchor and
  stale install command fixed, a real brand word in ROADMAP replaced with a
  placeholder.

- **The refreshed outline TOC cache is now a real Word-shaped TOC (REFLECTIONS
  P3).** The visible cache of a preserved or authored outline table of contents
  was rewritten as flat plain-text entries: one style for every level, no
  links, no page-number fields. The rewrite now authors deterministic
  bookmarks on the generated headings (`_TocBD000001` sequential,
  collision-scanned ids), harvests a per-level paragraph-style map from the
  template's OWN cached entries (trailing-digit TOC style convention,
  indentation-rank fallback, nearest-lower-level chain), and emits each entry
  as hyperlink-to-bookmark + tab + nested dirty `PAGEREF`, so previews show
  the template's real per-level TOC styling and Word fills page numbers on
  open. The bare-paragraph TOC path converges to the same multi-paragraph
  shape. Fail-closed: a malformed cache falls back to the previous simple
  rewrite atomically; documents without an outline TOC author zero bookmarks
  and stay byte-identical (now pinned by a dedicated TOC-free frozen anchor;
  the historical anchor was deliberately recomputed because the committed
  fixture itself contains an outline TOC, so its cached-TOC bytes necessarily
  changed shape).

- **Template demo entries no longer survive in derived indexes the content did
  not feed.** A kept caption index (table-of-tables / table-of-figures) whose
  sequence received no captions, and an outline TOC in a heading-less
  generation, used to keep the template's cached demo entries (the "stale
  derived index" defect class, visible in every preview renderer). Both caches
  are now rebuilt EMPTY: the field code survives, dirty, so Word recomputes it
  on open, but fabricated entries never reach the generated document. The
  destructive-action floor is untouched: structure is still never removed
  without corroboration; only the visible cache is honest now.

### Added

- **Local-corpus fidelity benchmark (REFLECTIONS P3).** New
  `scripts/corpus_benchmark.py` walks a corpus of REAL templates that lives
  outside the repository (and refuses one inside it), runs
  extract -> verify -> brand-agnostic probe generate -> QA per template in a
  throwaway directory, and writes a dated `report.{md,json}` next to the
  corpus with the LibreOffice-vs-Word caveat stated in the header. The probe
  input is typed blocks only (no styles, colors, fonts), so the corpus
  measures the engine without ever tuning it; `documentation/DEVELOPMENT.md`
  documents the layout. CI smokes the runner on the synthetic example
  template.
- **Authoring intelligence for the IntermediateDocument/GridDocument
  (REFLECTIONS P3).** The least-guarded stage (what the authoring agent puts
  IN the input) now has a guided surface: PROFILE.md gains three
  format-uniform sections rendered by the new shared `common/profilemd.py`
  (the role table on all three formats - pptx/xlsx previously shipped a
  two-line stub - a "Brand palette roles" table naming the semantic color
  tokens an author may reference, and "Authoring hints" that advertise
  comprehended fragments when present); each skill's SKILL.md gains an
  "Authoring the IntermediateDocument/GridDocument" section with role-first
  composition rules (skeleton order, heading hierarchy the TOC regenerates
  from, captions so derived indexes stay real, native objects on pptx, region
  bounds and formula preservation on xlsx, color discipline, fragments before
  re-derivation). Advisory by construction: nothing here adds engine
  authority, the resolver remains the only author of values.
- **Multi-template profile blending (REFLECTIONS P3, the quality-ceiling
  lift).** `extract --name X --template second.docx --blend` folds a SECOND
  same-format template into an existing profile at the VALUE-fact level: it
  fills captured facts the primary left unset, corroborates agreeing facts
  with a bounded deterministic confidence boost, and keeps the primary's value
  on every conflict. Artifact POINTERS (style ids, layouts, anchors,
  numbering) never cross shells: generation still opens the PRIMARY shell and
  the resolver still membership-validates everything against it, so the brand
  guarantee is untouched. Fail-closed all-or-nothing transaction (a rejected
  blend leaves profile.json byte-identical), sha-deduped and idempotent
  (re-blending the same file is a no-op), secondary shells stored
  content-addressed next to the primary with provenance, guarded by the new
  `blend_shell_provenance` QA check (donor tamper/drift = ERROR). Single-
  template profiles serialize without one new key and both frozen generation
  anchors are unmoved. Same-format only by design; cross-format stays the job
  of `compare-profiles`.
- **`compare-profiles`: the cross-template drift report (REFLECTIONS P3).**
  New read-only CLI verb that compares the BRAND-level facts of two saved
  profiles: theme colors per slot, theme/captured fonts, semantic palette
  roles, and off-theme usage (a raw hex in one profile that is a theme slot in
  the other - the sharpest "one template is off-theme" signal). Structural
  facts (styles, layouts, anchors) are never compared as drift: they are
  per-shell by design; role coverage is reported as information only. Writes
  nothing; exits 1 on brand-level drift so it can gate brand coherence in CI.
  First live run immediately caught a real inconsistency between the shipped
  docx and pptx example templates (five theme slots disagree) - queued for the
  next example-template refresh.
- **Contributor onramp (REFLECTIONS P1).** `CONTRIBUTING.md` expanded from a
  10-line stub to the full gate (dev setup, the three test lanes, the seven
  non-negotiable rules, PR checklist, sized first contributions);
  `scripts/brandkit/README.md` created as the engine internals map (module
  layout, the ten verbs' call paths, the two invariants, the recipe for a new
  capture axis); the release checklist written down in
  `documentation/DEVELOPMENT.md`; three sized "good first issue"s opened.
- **PLUGIN_WORKFLOW drift guard.** `documentation/PLUGIN_WORKFLOW.md` now
  documents the full ten-verb CLI surface and the learning loop
  (`generation_report.json` -> `learn`/`propose-overrides` -> `--accept`) in the
  end-to-end diagram; `tests/test_doc_drift.py` fails the build if a CLI verb
  or a `commands/` slash command is missing from the document.

## [0.8.1] - 2026-06-10

A post-0.8.0 hardening patch. Generated bytes are unchanged for well-formed
templates and profiles (frozen anchor verified); the fixes close
crash/corruption windows on malformed or adversarial input, pin the QA
check-id vocabulary, repair the plugin onboarding path, and make the example
pptx import cleanly in strict readers (Keynote).

### Fixed

- **Engine commands now work from the skill folder (plugin onboarding).** Every
  skill/command doc taught `python scripts/brandkit/cli.py ...`, which resolves
  only from the repo/plugin ROOT, while a skill's natural working directory is
  `skills/<skill>/` - so every fresh plugin session tripped once on
  "No such file or directory" before recovering. All 13 docs now teach
  `python scripts/cli.py ...` (the LAUNCHER each skill already shipped, which
  resolves the engine root by itself; `BRAND_DOCS_ROOT` overrides), a root-level
  launcher copy makes the one documented invocation work from both locations,
  and a Hard Rule states the behavior explicitly.
- **The example pptx now imports in Keynote.** python-pptx creates the
  notesMaster part + relationship when speaker notes are authored but never
  declares it in `p:notesMasterIdLst` (upstream omission); strict importers
  reject the package as invalid. The pptx example builder gains a
  deterministic post-save repair (`_fix_notes_master_id_lst`) that adds the
  spec-ordered id list; PowerPoint/LibreOffice behavior is unchanged.
- `schema.role_id` now validates its own composition and raises `ValueError`
  when the composed id would fail `is_valid_role_id` (empty family,
  non-lowercase or special characters in a qualifier), so a malformed id can
  never silently fail a later role lookup.
- `refresh_visible_caption_index_cache` (docx) validates BOTH field-inventory
  bounds before any body-children access: a corrupted inventory with a negative
  `begin_index` used to corrupt the body via python negative indexing, and the
  splice loop is additionally clamped to the children length.
- `_apply_borders` (docx) refuses a parsed border side whose qualified tag is
  not the expected `w:<side>` (a mislabeled capture can no longer land under
  `w:pBdr` as the wrong side) and narrows its parse-failure catch from bare
  `Exception` to the two failures a profile string can cause
  (`XMLSyntaxError`, the unicode-declaration `ValueError`).
- The docx set-only-when-unset guards (`_set_twips_if_unset`,
  `_reassert_level_facts`) treat an explicitly EMPTY attribute (`w:val=""`,
  not valid OOXML) as unset, so a malformed template cannot pin a captured
  fact behind an empty string; authored values are still never clobbered.

### Added

- Frozen `CHECK_REGISTRY` in `qa/model.py`: the closed vocabulary of every
  check id the deterministic L0 library may stamp on a `Finding`, with the
  naming convention documented in CONVENTIONS.md section 9.
  `tests/test_check_registry.py` AST-scans the module and fails on any
  unregistered (or stale) id, catching silent typos fail-closed in CI.

### Tests

- `tests/test_review_hardening.py`: regression coverage for the four fixes
  above (composition validation, caption-index bounds, border-side refusal and
  narrowed catches, empty-string-counts-as-unset guards).

### Changed

- **Example docx template: visual-audit fixes** (rebuilt byte-deterministically
  from its builder after a page-by-page render audit): the decorative cover
  band is now a flush solid navy bar (the rounded amber-outlined pill rendered
  as a broken-looking box); the landscape appendix explicitly resets to ONE
  column (python-docx copies the previous sectPr, so it inherited the
  two-column synopsis layout - matrix squeezed into half the page); the
  rollout-matrix header runs are white over the navy fill (they rendered
  navy-on-navy, unreadable); page-spanning branded tables repeat their header
  row (`w:trPr/w:tblHeader`); and the two-column synopsis carries enough copy
  to fill BOTH columns.

### Docs

- The three `SKILL.md` now say "the six questions" (matching the shared
  `reference/comprehension.md`); `brand-docx/SKILL.md` gains the
  `artifact_catalog` workflow paragraph and a "docx readiness" caveat for
  parity with the pptx/xlsx skills; the three `reference/visual-audit.md`
  drop an en-dash for a hyphen (house style); `NOTICE` states which
  components delegate to python-docx / python-pptx / openpyxl / lxml / Pillow
  and which are custom-implemented.
- README sets the speed expectation honestly: the FIRST run on a new template
  can take up to ~15 minutes end to end with an AI agent (extract + model
  comprehension + visual QA + repair rounds); every later document from the
  saved profile takes seconds - still a fraction of formatting by hand, with a
  faithful file instead of an approximate one.

## [0.8.0] - 2026-06-10

### Performance

Seven behavior-preserving quick wins from a 3-way profiling pass (engine / QA /
test suite). Generated document bytes, extracted profiles, and QA findings are
byte-identical before and after (verified by byte-diffing extracted profiles and
full QA reports for all 3 formats, plus the frozen anchor and idempotency
tests). Numbers measured on the reference machine (M-series macOS, soffice +
pdftoppm installed). Structural follow-ups (probe caching, render daemon, L2
short-circuit extension) are documented - not applied - in
`documentation/ROADMAP.md` section 6.

- **Single-launch 3-format doctor render probe.** `doctor._probe_visual_pipeline`
  now converts all three probe documents (distinct basenames) with ONE
  `soffice --convert-to pdf` launch and one fresh `UserInstallation` instead of
  three serial launches, keeping the fail-closed per-format assertions (each
  format's PDF must exist AND rasterize, same format-attributed errors).
  `doctor.probe()` 5.81s -> 3.32s; `generate --qa auto` 8.11/7.77/8.28s ->
  4.74/4.61/5.00s (docx/pptx/xlsx); real-render test lane 41.4s -> 25.2s;
  `doctor` CLI output bit-identical on healthy and broken environments.
- **Per-`run_qa` artifact-load memo.** One QA pass used to reopen the same
  shell/output via python-docx/openpyxl/python-pptx up to 9x across independent
  checks; `gate.run_qa` now scopes a load memo (`checks_deterministic.load_memo`)
  to a single pass so each path is loaded at most once. The memo dies with the
  pass - nothing is cached across invocations and the shell-sha check still
  reads bytes from disk. Post-generate QA pass: xlsx 119.1ms -> 31.9ms, docx
  24.6ms -> 17.0ms, pptx 19.5ms -> 13.4ms.
- **Lazy Office-lib imports in `qa/checks_deterministic.py` and `cli.py`.** The
  three Office libs import inside the (memoized) loaders, and the six per-format
  extract/generate modules import inside their CLI dispatch branches, so an
  invocation pays only for the active format. `import brandkit.cli` 191ms ->
  43ms; `cli.py list` 0.29s -> 0.13s.
- **One materialized docx run-facts pass per extract.** `capture_fonts` and
  `capture_pseudo_headings` (which scanned twice) now share a single
  `collect_font_run_facts(doc)` list instead of three full document walks.
  Warm docx extract (complex fixture) 81.1ms -> 73.1ms.
- **Shared pptx layout/slide classification per extract.** `classify_layouts`
  (computed 5x) and `classify_slides` (2x) are computed once in
  `pptx/extract.py` and threaded through `_roles` / cover anchors / regions /
  skeleton via new optional parameters (pure functions of the unmutated deck).
  Warm pptx extract (example template) 78.3ms -> 62.9ms.
- **Test suite: injected visual seam + class-scoped fixture extraction.** The 4
  only suite call sites that ran `gate.run_qa` with the default `qa='auto'`
  (and so spawned real soffice work) now pass the documented test seam
  `visual=(False, [])`, keeping the visual-audit branch executing in degraded
  mode (8.50/7.64/7.46/7.44s -> 0.10/0.03/0.04/0.03s); the ~17 identical
  per-test extracts of the committed complex docx fixture are class-scoped
  (`setUpClass` + per-test `deepcopy`, the pattern
  `test_xlsx_complex_fidelity.py` already uses). Full suite 59.5s -> 24.0s,
  green twice in a row; renderer-present coverage stays in `test_visual_qa.py`
  and the opt-in `BRANDDOCS_RUN_REAL_RENDER=1` lane.

### Added

- **Example docx template now exercises every new capture axis.** The synthetic
  `examples/templates/branddocs_template.docx` (rebuilt byte-deterministically from
  its builder) now carries: a dominant DIRECT paragraph-spacing convention
  (`w:spacing w:after="160"` on every body paragraph - D1 geometry capture), declared
  `w:tblCellMar` on all five tables alongside the existing `w:tblLook` (D2 table
  capture: `cell_margins` + `style_id` + `tblLook` all captured), an explicit dominant
  body run convention (Calibri/22hp/navy, set-only-when-unset so headers/accents keep
  their look) that gives `theme.fonts.body` / `theme.text.body` their baseline, an
  off-theme coral accent run (`hex:C24D2B` palette entry - the E1 alias channel's raw
  material), and a FAKED heading (36hp teal body-style paragraph) that the E2
  detector surfaces as a `pseudo_heading` candidate (together with the coral run, a
  realistic adjudication case for the model). The pptx/xlsx examples already carry
  off-theme `hex:` palette entries, so E1 is exercisable on all three formats.
- **Universal cover synthesis for `AnchorKind.NONE` (Cluster E4).** A template whose
  cover-anchor detection recorded the structural fact `anchors.cover.kind == NONE` (no SDT
  title slot, no cover-layout placeholder, no anchorable cover structure) previously hit a
  dead-end: the in-place cover machinery had nothing to fill, so the authored title/subtitle
  was placed only by the docx last-resort append (title alone, subtitle lost) and dropped
  entirely on the pptx reconcile path. E4 closes it: when (and ONLY when) the profile
  records `kind == NONE`, the generator BUILDS a minimal cover from the profile's own
  RESOLVABLE `cover.*` roles through `resolve_role` (the single brand chokepoint) - never
  from literals - and no-ops byte-identically when nothing resolves.
  - **DOCX:** a new synthesis branch in `compose_cover` (completely DISJOINT from the
    anchored reconciliation): each authored slot (title, then subtitle - the canonical order)
    is resolved via `ProfileResolver.resolve_role`; a slot is synthesized only when its role
    resolves to a style the live shell actually carries, as an ordinary role-styled paragraph
    moved before the first toc/body child (the same position the append fallback used). A
    stub/unresolvable role contributes nothing; an authored title is NEVER dropped (title
    content + unresolvable `cover.title` declines synthesis entirely, falling through to the
    unchanged deterministic fill). The synthesized paragraphs are never touched by
    `_sync_core_properties` / the SDT `showingPlcHdr` machinery (existing-anchor-only).
  - **PPTX:** the reconcile path gains the synthesis step: a `kind == NONE` deck surfaces an
    empty `cover_anchors` inventory, so the model can bind no `cover_slots` and the cover
    fill was a guaranteed no-op - the authored cover silently never appeared. Now, when the
    `cover.title` role still resolves to a real shell layout (e.g. a learned reroute), the
    cover slide is built on that role-resolved layout via the same run-preserving placeholder
    fill the deterministic rebuild uses. The deterministic path's bytes are UNCHANGED (it
    already built the cover on the role-resolved layout); it now records the audit finding
    when that happened on a `kind == NONE` deck.
  - **XLSX: N/A by design.** Xlsx cover anchors are named ranges - data geometry with no
    page/slide cover concept and no dedicated cover writer; generation steers them as
    ordinary named-range fills, so there is nothing to synthesize (pinned by test).
  - **Audit:** every synthesized cover records an INFO `cover_synthesized` finding (the
    `override_applied` mirror) naming the structural fact (`kind=NONE`) and the role ids
    used - never brand text. Unplaced extra authored fields surface as the existing INFO
    `cover_degraded` note, never silently lost.
  - **Authorization decision: deterministic-only (no new comprehension sink).** The ROADMAP's
    optional closed `cover_layout` enum was evaluated and deferred: the trigger is a recorded
    structural fact, the slot order is canonical (title, subtitle), and every value comes
    from already-inferred+validated roles, so there is nothing for the model to name that
    `resolve_role` does not already gate - and the no-sink design keeps merge/validation
    surface and the model-facing reference docs byte-identical. If a future template needs
    model-ordered slots, the enum can land later as an additive opt-in.
  - **Byte-identity:** every `kind != NONE` profile (and every profile that never recorded
    the fact, `anchors == {}`) takes the existing paths untouched; a `kind == NONE` profile
    without a resolvable `cover.*` role or without cover content is byte-identical to
    before. Frozen anchor, full cover regression suite and the real-render lane stay green;
    schema stays 1.2.0 (no new keys).
- **Uniform parity ledger: `appearance_apply_degraded` (Cluster E3).** Whenever a
  resolved op carries a captured appearance axis the format's writer cannot realize,
  the shared apply orchestration now emits ONE stable INFO finding per (role, axis)
  (`location='<role_id>:<axis>'`, message names only role id + axis, never a brand
  value) instead of silently dropping the captured intent. Each backend DECLARES its
  capability (`realized_axes`: docx all six axes, pptx/xlsx the run-typography trio;
  an undeclared backend keeps its exact prior inferred behavior), so future
  capture-before-apply increments and hand-edited profiles surface measurably - the
  finding feeds the cross-run `generation_history` (B4) and the L2 model. INFO-only:
  not in `DEFAULT_L0_INVARIANTS` or `LEARNABLE_CHECKS`, can never flip a verdict;
  every existing real profile emits nothing (byte-identical paths).
- **Faked-heading-in-body-style detection (DOCX-first, Cluster E2).** Some templates FAKE a
  heading: a line that LOOKS like a heading (visibly larger and/or in a brand color) is authored
  with the BODY paragraph style - no heading role, no named heading style - so the deterministic
  engine treats it as body and the brand heading look is lost. E2 closes this in three
  deterministic-plus-model steps, reusing the existing resolver/appearance seam, with NO new
  appearance check and NO schema bump (additive within 1.2.0).
  - **Detect (pure-deterministic):** a new `common.typography.detect_pseudo_headings` is a PURE
    STATISTIC vs the captured DOMINANT body appearance - a body-style run whose EXPLICIT
    `size_hp` is a clear OUTLIER (>=1.5x or <=0.67x the dominant body size) OR whose EXPLICIT
    color is OFF-BODY (a different bucket than the dominant body color) is a candidate fake
    heading. Nothing is hardcoded to any template (the ratio/equality test is against the
    template's OWN observed dominant). The docx extractor's `capture_pseudo_headings` runs after
    `capture_fonts` and stores each candidate additively under `theme.pseudo_headings` as
    `{ref, size_hp?, color?, evidence}` - a stable structural ref, the run's OWN captured outlier
    size/color (a FACT about the template, never synthesized), and coarse, brand-text-free
    evidence. A uniform body leaves the key ABSENT (no-op).
  - **Surface:** `comprehend_input_bundle` adds `facts.pseudo_headings` ONLY when the detector
    found candidates (the B4 `generation_history` pattern), so the no-candidate bundle is
    byte-identical. The facts are READ-ONLY; they never change generation on their own.
  - **Adjudicate (model NAMES only):** a new closed `comprehension.promote_appearance` sink (a
    peer of `triage`) - a list of `{pseudo_heading_ref, target_role_id}`. The model NAMES a
    surfaced ref + a declared heading role; it authors NO size/color. Shape is validated by
    `schema._validate_comp_promote_appearance` (which also REJECTS a model-authored
    `size_hp`/`color`); membership is validated fail-closed by a new
    `comprehension.check_promote_appearance` (ref must be a surfaced `pseudo_heading`, target
    must be a DECLARED `heading.*` role, `(ref, target)` unique), wired into `merge()` in the
    SAME all-or-nothing transaction as `check_triage`. An empty/absent palette of candidates
    fails closed (like every other load-bearing ref).
  - **Dispose (engine AUTHORS the value):** on a clean merge, a new
    `comprehension._derive_promote_appearance` COPIES the captured outlier `size_hp`/`color`
    from the detector fact onto `roles[target_role_id].appearance` - the SAME role-specific dict
    `resolver._merge_appearance` reads (a role-specific size/color applies to ANY role WITHOUT
    the body-default family gate, so a promoted value on `heading.1` WILL apply). The two axes
    are independent (a size-only outlier writes no color).
  - **Verify (existing check, no new one):** because the promoted size/color now live in
    `roles[*].appearance`, the EXISTING `check_appearance_targets` re-validates them
    SHELL-BACKED at QA time - a promoted size the shell carries on no run, or a promoted color
    absent from the palette/runs, is an ERROR. The engine never injects a value the template
    does not contain.
  - **Universal + additive:** schema stays at 1.2.0 (`promote_appearance` is a new closed sink
    and `theme.pseudo_headings` is additive). A profile that adjudicates no promotion leaves the
    heading appearance unchanged, so generation stays byte-identical (the frozen anchor stays
    green). pptx/xlsx carry no body-style runs, so they surface no candidates and stay
    byte-identical. The model bundle's `reference/comprehension.md` (all three byte-identical
    copies) documents the `promote_appearance` adjudication.
- **Off-theme accent reachability via palette aliases (all three formats, Cluster E1).** An
  off-theme brand accent the template uses in content is captured as a `hex:RRGGBB`
  `theme.palette` entry, which is harder to ADDRESS as a named run color than a clrScheme
  theme slot (`accent1` ...). E1 lets the model propose a human/semantic **alias** for such a
  captured entry, and the engine mints a syntactically-legal dotted token in `theme.palette`
  whose `ref` is COPIED BYTE-IDENTICAL from the captured entry's `ref` - so the accent becomes
  addressable as a clean run-color token, WITHOUT the model ever authoring a hex.
  - **Annotation:** `comprehension.palette_annotations[<palette-id>]` gains an optional
    `alias` directive (a dotted token name). It is additive and absent by default; the model
    NAMES the alias, never a color value.
  - **Mint:** a new `comprehension._derive_palette_aliases` runs in `merge()` immediately after
    `_derive_palette_annotations`, in the SAME all-or-nothing transaction. For each annotation
    carrying an `alias`, it mints `theme.palette[alias]` with the captured `ref` deep-copied, a
    `palette.alias` provenance fact recording the source key, the source `frequency`, and all
    advisory fields null (a pure BRIDGE token). The mint is idempotent and refuses to shadow a
    non-alias key.
  - **Resolve (zero resolver change):** `resolve_color(token)` still reads
    `theme.palette[token]['ref']` verbatim - an alias is just another palette key, so it is
    addressable identically on docx / pptx / xlsx through the shared resolver and
    `appearance.apply_run_color`.
  - **Verify:** `comprehension.check_membership` rejects an alias whose token is syntactically
    illegal, collides with an existing palette key / role id, or is proposed twice (fail-closed
    at merge); a new fail-closed L0 check `check_palette_alias_targets`
    (gate-wired as `palette_alias_targets_exist`, listed in `DEFAULT_L0_INVARIANTS`) verifies at
    QA time that every minted alias token is a legal dotted token whose minted `ref` is
    byte-identical to its declared source entry's `ref` (an invented or diverged ref is an
    ERROR). An alias to a non-captured palette key is rejected by the existing
    `palette_annotations` membership check; an empty palette fails closed.
  - **Universal + additive:** schema stays at 1.2.0 (the `alias` field and the minted tokens
    are additive within the existing 1.2.0 `comprehension` / `theme.palette`). A profile that
    carries no alias mints NO new palette key, so generation stays byte-identical (the frozen
    anchor stays green). The model bundle's `reference/comprehension.md` (all three byte-identical
    copies) documents the `alias` annotation.
- **List / numbering-definition fidelity is now a captured brand axis (DOCX only, Cluster
  D3).** The same dominant-capture + appearance-seam + fail-closed-check pattern that ships
  for font/size/color/geometry/table now also covers a generated list's NUMBERING facts, so
  a generated list looks like the template's lists: per level, the `w:numFmt` field code
  (decimal / bullet / lowerLetter / ...), the `w:lvlText` level format string (e.g. `"%1."`
  or a bullet glyph), and the level `w:ind` indentation. This completes Cluster D
  (D1 + D2 + D3).
  - **Capture** reads the per-level facts VERBATIM off the shell's OWN `w:abstractNum` for
    the numbering id each list role already references (resolved `numId -> abstractNumId`),
    storing them additively under `role.appearance.numbering` (the referenced `num_id` /
    `abstract_num_id` plus a `per_level_facts` map). Only the levels the template declares
    are recorded; nothing is hardcoded to any template.
  - **Apply** flows through the single resolver/appearance seam (`op.appearance.numbering`):
    the generator REFERENCES the shell's numbering by id (the existing `w:numPr`), ensures
    the referenced `w:abstractNum` is present (CLONING the shell's own definition by id when
    needed - never minting a new one), then re-asserts each per-level numFmt / lvlText /
    indent SET-ONLY-WHEN-UNSET so an authored value is never clobbered and re-runs stay
    byte-identical. The engine NEVER synthesizes a numbering definition from JSON. A
    no-numbering profile generates byte-identically (the frozen anchor stays green).
  - **Verify** adds `check_numbering_targets` (gate-wired as `appearance_numbering_targets`),
    the honest fail-closed peer: every referenced `num_id` / `abstract_num_id` must be one
    the shell's numbering part DEFINES (SYMBOLIC membership, like the table-style ref); every
    per-level numFmt must be a valid OOXML field code (SHAPE, a closed enum); and every
    per-level lvlText / indent must be byte-identical to the shell's OWN abstractNum for that
    level (OBSERVED-FLOOR, like geometry). An undefined id / malformed numFmt / un-observed
    lvlText / out-of-range or un-observed indent is an ERROR.
  - **Still DOCX-only:** pptx/xlsx never capture or apply numbering (`w:abstractNum` /
    `w:num` are WordprocessingML constructs with no peer in scope here); their output is
    untouched. Schema stays 1.2.0 (additive).
- **Table-style fidelity is now a captured brand axis (DOCX only, Cluster D2).** The
  same dominant-capture + appearance-seam + fail-closed-check pattern that ships for
  font/size/color/geometry now also covers a generated table's CONDITIONAL-FORMAT facts,
  so a generated table looks like the template's tables: the `w:tblLook` bitmask (which
  of the table style's `w:tblStylePr` emphases - first/last row/column, row/column
  banding - apply), the referenced table STYLE id, and the `w:tblCellMar` cell margins.
  KPI-as-table inherits this for free (the KPI writer routes through the same table
  path).
  - **Capture** is a dominance statistic over the template's OWN tables, identical floor
    to the other axes (`MIN_RUNS` + `MIN_DOMINANCE`): each fact is an INDEPENDENT axis,
    recorded only when an explicit value dominates the template's tables. Nothing is
    hardcoded to any template. Stored additively under `role.appearance.table` and
    `theme.table.body`.
  - **Apply** flows through the single resolver/appearance seam (`op.appearance.table`)
    and writes each fact set-only-when-unset: an authored `tblLook` / style / margin is
    never clobbered (python-docx's synthetic-default `tblLook` is treated as unset so the
    captured bitmask can enable the shell style's own emphases). The engine NEVER authors
    a `w:tblStylePr`, a fill, or a border - the band fills stay in the SHELL's style part
    and are only TOGGLED via the bitmask + the style reference. A no-table profile
    generates byte-identically (the frozen anchor stays green).
  - **Verify** adds `check_table_targets` (gate-wired as `appearance_table_targets`), the
    honest fail-closed peer that validates THREE INDEPENDENT dimensions: the `tblLook`
    bitmask is WELL-FORMED (an int in the 16-bit OOXML range whose only set bits are the
    spec-fixed flags - SHAPE, not membership); the referenced table STYLE is one the
    shell's styles part actually defines (SYMBOLIC name-membership, like fonts); and each
    cell margin is in the OOXML twips range AND byte-identical to a margin the template's
    OWN tables carried (OBSERVED-FLOOR, like geometry). A malformed bitmask / undefined
    style / un-observed margin is an ERROR.
  - **Still DOCX-only:** pptx/xlsx never capture or apply table appearance
    (`w:tblLook` / `w:tblStyle` / `w:tblCellMar` are WordprocessingML table constructs
    with no peer in scope here); their output is untouched. Schema stays 1.2.0 (additive).
- **Paragraph geometry is now a captured brand axis (DOCX only, Cluster D1).** The
  proven dominant-capture + appearance-seam + fail-closed-check pattern that ships for
  font family/size/color now also covers paragraph GEOMETRY: spacing (before/after/
  line + line rule), indentation (left/right/first-line/hanging), paragraph borders
  (`w:pBdr`), and shading (`w:shd`), all read from the template's own `w:pPr`.
  - **Capture** is a dominance statistic, identical floor to the typographic axes
    (`MIN_RUNS` + `MIN_DOMINANCE`): every geometry property is an INDEPENDENT axis, and
    a value is recorded for a role only when it dominates that role's own paragraphs
    (the document body geometry when it dominates all of them). Nothing is hardcoded to
    any template's twips. Stored additively under `role.appearance.geometry` and
    `theme.geometry.body`; borders are byte-copied as serialized elements.
  - **Apply** flows through the single resolver/appearance seam (`op.appearance.geometry`)
    and writes each property set-only-when-unset, so an authored/inherited value is
    never clobbered and a no-geometry profile generates byte-identically (the frozen
    anchor stays green). Geometry has NO family gate (unlike body size/color): a
    heading's own captured indentation/spacing is intentional branding and applies.
  - **Verify** adds `check_geometry_targets` (gate-wired), the honest fail-closed peer
    for an intrinsic-measurement axis: it does NOT do name-membership against a shell
    inventory (geometry is captured numbers/elements, not symbolic refs). It proves
    every applied value is WELL-FORMED (twips in the OOXML range, valid border element,
    real `RRGGBB` shading) AND byte-identical to a value the template's OWN paragraphs
    carried (the captured floor) - a synthesized or out-of-range value is an ERROR.
  - **Still DOCX-only:** pptx/xlsx never capture or apply geometry (WordprocessingML
    `w:pPr` has no shape-geometry peer in scope here); their output is untouched.
    Schema stays 1.2.0 (additive).
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
- **Model-proposed overrides corrections (the model proposes a fix for the AMBIGUOUS
  recurring remainder; Cluster B4).** The deterministic `learn` binds only the
  findings it can resolve on its own; the remainder (a stub role with no healthy
  sibling, a finding whose right re-point needs judgement) is now surfaced to the
  model in the `comprehend-input` bundle as a bounded, MESSAGE-FREE
  `facts.generation_history` slice (`{check, location, severity, recurred_runs}`,
  keyed on `(check, location)` only - no brand text leaks). A new `propose-overrides`
  verb takes a model-authored proposal, OVERLAYS it onto any existing lesson
  (`overlay_overrides`, additive: a deterministic lesson and a model proposal coexist)
  and routes the WHOLE block through the SAME single `merge_overrides` sink, so every
  pointer is re-bound fail-closed and the model can only ever NAME a shell-backed
  re-point (never author a style/font/color). Advisory until `--accept` (mirrors
  `learn`); byte-identical generation until a correction goes live. Every LIVE
  override is now auditable: a gate-wired `check_overrides_applied` emits an INFO
  `override_applied` finding per live entry (in `generate` and `verify`), so a learned
  re-point is never silent (INFO-only, never flips a verdict). Schema stays 1.2.0.
- **Model-in-the-loop (the model's judgements persist and re-validate fail-closed).**
  Three additive touchpoints, all routed through the comprehension block (the single
  model-writable, fail-closed, shell-frozen home); model PROPOSES, deterministic
  DISPOSES; byte-identical when absent:
  - **Persisted L2 visual-audit verdict**: the orchestrator's per-checklist
    PASS/FAIL/NA judgement (written looking at the rendered PNGs) is cached as a
    `comprehension.audit` sub-block keyed by the structural checklist id
    (`check_audit_targets`, fail-closed). `generate` then SHORT-CIRCUITS the L2
    render round only when every CURRENT checklist id is PASS at an exact
    `(shell_sha, content_sha)` match - any change defeats the gate, so it can never
    mask a regression; disabled under `--qa strict` and at verify.
  - **Model-assisted QA triage**: the model adjudicates AMBIGUOUS WARNINGs
    (blank-page / edge-bleed / component-survival) as expected (demote to INFO) or
    defect (keep), via a `comprehension.triage` map bound to the closed
    `AMBIGUOUS_TRIAGE_CHECKS` set (`check_triage_targets`). The eligible set is
    WARNING-only and the consume path guards on `severity == WARNING`, so a triage
    entry can NEVER demote an ERROR.
  - **Interactive `refine` verb**: turns a user's qualitative end-of-generation
    feedback into a comprehension delta over the EXISTING annotation sinks
    (`palette_annotations` / `role_annotations` / `demo_classification`), routed
    through the same fail-closed `merge` (every binding a verbatim surfaced id),
    shown as a confirm-as-diff and advisory until `--accept`. The skills now ask for
    feedback ONLY at the end of generation, invite it as text OR a screenshot, and
    apply it to FUTURE generations (never the just-produced file).

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

[0.9.0]: https://github.com/ferdinandobons/brand-docs/releases/tag/v0.9.0
[0.8.1]: https://github.com/ferdinandobons/brand-docs/releases/tag/v0.8.1
[0.8.0]: https://github.com/ferdinandobons/brand-docs/releases/tag/v0.8.0
[0.7.0]: https://github.com/ferdinandobons/brand-docs/releases/tag/v0.7.0
[0.6.2]: https://github.com/ferdinandobons/brand-docs/releases/tag/v0.6.2
[0.6.1]: https://github.com/ferdinandobons/brand-docs/releases/tag/v0.6.1
[0.6.0]: https://github.com/ferdinandobons/brand-docs/releases/tag/v0.6.0
[0.5.0]: https://github.com/ferdinandobons/brand-docs/releases/tag/v0.5.0
[0.4.0]: https://github.com/ferdinandobons/brand-docs/releases/tag/v0.4.0
[0.3.0]: https://github.com/ferdinandobons/brand-docs/releases/tag/v0.3.0
[0.2.0]: https://github.com/ferdinandobons/brand-docs/releases/tag/v0.2.0
[0.1.0]: https://github.com/ferdinandobons/brand-docs/releases/tag/v0.1.0
