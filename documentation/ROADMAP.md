<!-- SPDX-License-Identifier: MIT -->
# Roadmap and planned features

Planned, designed-but-not-yet-built work for BrandDocs. Each entry records the
problem, the root cause (with code references), and a design that is **compatible
with the architecture as it exists today** so the work can start without a rewrite.

The non-negotiable constraint for every item below: **off-brand output stays
impossible by construction.** Brand-specific values (style names, fonts, hex
colors) live only in the Brand Profile, only the resolver reads them, and `verify`
refuses a profile that points at anything the template does not actually contain
(`check_resolver_targets`, fail-closed). No design here may weaken that.

---

## 1. Brand typography capture (direct-formatting -> role `appearance`)

> **Status: shipped (font family + size + color).** The capture/verify/apply
> pipeline below is implemented for THREE independent axes - font family, font size,
> and run color. Extraction records the dominant direct run value (sampled over ALL
> runs, so a minority accent never wins) into `role.appearance` and the document
> defaults `theme.fonts.body` (font/size) + `theme.text.body` (color); `verify`
> re-validates each applied value against the shell (`appearance_targets_exist`:
> fonts vs fontTable+theme, sizes vs the template's `w:sz` set, colors vs the theme
> palette + the template's `w:color` set); and generation applies them as direct run
> formatting through the resolver, with the body size/color default gated off heading
> roles. Still future: **per-word accents**, **cover-layout reconstruction** (the
> separate `cover.kind = NONE` gap), and **heading typography** when a template fakes
> headings in the body style (route via `comprehension.role_annotations`). The design
> below documents the full feature.

### Problem
A generated document does not always match the template's **real visible
typography**. Observed on a real template (`Come funziona la Sotto-community.docx`):
the original renders its title in **Montserrat** and its body in **Roboto**, with a
brand-blue accent, but the generated document came out in **Arial**.

### Root cause (verified)
The brand's real fonts live **only as direct run-level formatting** on the
template's content, not in any reusable artifact the extractor reads:

- Extraction maps roles from **named paragraph/table styles only**. It never reads
  run-level `w:rFonts` / `w:color` / `w:sz` (`formats/docx/roles.py:infer_roles`,
  `formats/docx/extract.py`).
- `_extract_theme` does not even capture the theme's real major/minor latin fonts
  (it stores `latin: None` with hardcoded `Arial` / `Calibri` fallbacks) and does
  not read `docDefaults` (`formats/docx/extract.py:_extract_theme`).
- Generation applies a role as a **named style only** (`target_obj.style = ...`);
  there is no code path that applies font/color directly from the profile. Run
  toggles are semantic (bold/italic/underline) and explicitly never carry a brand
  font or color (`formats/docx/generate.py:_apply_style` / `_apply_run_toggles`).

So the only fonts the styles resolve to are `docDefaults = Arial`, and the
template's actual Roboto/Montserrat were never captured anywhere.

### Proposed design (compatible)
Use the **already-reserved, currently-empty `appearance: {}` field on each role**
(`formats/docx/roles.py`; present in the envelope, populated by nothing today) as
the home for captured typography. Three additive layers:

1. **Capture (extract).**
   - Fix `_extract_theme` to read the real major/minor latin typefaces from
     `theme1.xml` and the `docDefaults` `rFonts`, so `theme.fonts` is truthful.
   - Add a deterministic typography sampler: for each role, sample the **dominant**
     direct run formatting actually applied to that role's runs (font, run color,
     size). When a clear convention dominates (not noise), record it into the
     role's `appearance` (e.g. `appearance: {font: {latin: "Roboto"}, color:
     {hex: "..."}, size_pt: 11}`) with a confidence. This is the
     "deterministic disposes" half of the existing pattern.

2. **Verify (keep the guarantee).**
   - Extend the resolver-target check with an `appearance` validator: a captured
     font is legitimate only if the **shell itself uses/embeds it** (its
     `fontTable` / referenced fonts); a captured color must be a theme color or a
     hex actually present in the template. A profile naming a font the template
     never uses is an ERROR, exactly like a missing style today
     (`qa/checks_deterministic.py:check_resolver_targets`).

3. **Apply (generate).**
   - Add one new code path in `_apply_resolved_style` that, after setting the named
     style, applies the role's `appearance` as direct run/paragraph formatting,
     reading the values **only from the profile via the resolver** (no literal in
     the writer). Roles without `appearance` (every profile that exists today)
     behave exactly as now.

### Compatibility notes
- Schema is **additive** (`appearance` already exists): bump `schema_version`
  minor; old profiles keep working unchanged.
- The guarantee holds: fonts/colors are captured **facts about the template**,
  stored in the profile, applied only by the resolver, and re-validated against the
  shell by `verify`. You can only ever apply typography the template itself uses.
- Deterministic; no model required. The optional `comprehend` step could later
  confirm which captured appearance is intentional vs incidental.

### Scope / non-goals (first cut)
- Role-level **dominant** typography (font + dominant color + size). Per-word
  accents (e.g. only the word "Collabfor.it" in blue) are an `IntermediateDocument`
  authoring concern, out of scope.
- Cover-layout reconstruction (the separate `cover.kind = NONE` detection gap) is
  tracked independently; typography capture alone already makes a `heading.1` title
  render in the brand font.

---

## 2. Profile learns from generation QA findings (a feedback loop)

> **Status: B1-B4 SHIPPED (deterministic core + model-proposed phase).**
> Implemented: every `generate` persists `generation_report.json`
> (verdict + findings verbatim + shell/content/output sha256) next to the output;
> cross-run `regression.recurred`/`regression.reintroduced` findings (advisory,
> keyed on `(check, location)`, same-shell only); a deterministic `learn` verb
> that distills unambiguous recurring findings into the shell-frozen
> `rules.overrides` block (closed vocab: reroute_role / number_format /
> register_demo_clear) through a single all-or-nothing fail-closed sink, consumed
> by the resolver only as a LAST-RESORT on a stub and re-proven at verify by
> `check_override_targets`; and **(B4)** a model-proposed phase - the
> `comprehend-input` bundle surfaces a message-free `generation_history` slice of the
> AMBIGUOUS recurring remainder, and a `propose-overrides` verb overlays a
> model-authored correction onto any existing lesson and routes it through the SAME
> `merge_overrides` sink, with a gate-wired `override_applied` INFO audit finding for
> every live correction. Lessons stay ADVISORY until `--accept`. See
> CONVENTIONS §14. The design below documents the full feature.

### Problem
If extracting a profile leads to repeatable problems at generation time (a role
that resolves to a weak/missing style, a fragment that degrades, a rejected
number format, a demo value that keeps reappearing), the **next** generation from
the same profile hits the same problems and needs the same repair loop. We want
those profile-driven problems to be learned once and not recur, so a subsequent
generation is faster (fewer repair/regenerate rounds).

### Root cause (verified)
Generation findings are **ephemeral**: collected into a `findings` out-param,
folded into the `QAReport`, and printed to stdout, but **not persisted per run**
and never fed back into the profile (`cli.py` generate subcommand;
`formats/docx/generate.py`). Only `visual_manifest.json` is persisted, and it
records L1/OCR findings, not the generation/L0 findings. Nothing closes the loop.

### Proposed design (compatible)
Reuse the existing seams rather than inventing a parallel system:

1. **Persist the signal.** Write a per-run `generation_report.json` next to the
   output (mirrors the existing `visual_manifest.json` side artifact): the
   `QAReport` findings (`check`, `severity`, `message`, `location`) plus the shell
   sha256, a content hash, and a timestamp. Additive; no behavior change.

2. **A shell-bound learned-corrections registry.** Record distilled corrections on
   the profile under the **already-present `rules.overrides`** seam (or a sibling
   `lessons` block), frozen to the shell exactly like comprehension
   (`source_shell_sha256 == provenance.shell.sha256`;
   `profile/store.py:comprehension_is_present`). A correction may only ever
   **re-point within artifacts the shell already defines** (e.g. "role X's weak
   style -> fall back to role Y", "fragment Z degrades repeatedly -> deprioritize",
   "this number_format mask was rejected -> use shell-backed mask M", "demo value V
   keeps surviving -> register it for demo-clearing"). It can never invent a style,
   font, or color.

3. **Who writes lessons.** Two phases:
   - Phase A (deterministic, no model): a `learn` step distills corrections from
     the **unambiguous** findings (`resolver_targets_exist`, `style_fallback`,
     number-format rejection, residual-demo hits). Deterministic finding ->
     deterministic correction.
   - Phase B (model-assisted, later): ambiguous findings are proposed by the model
     in a comprehend-like step ("model proposes, deterministic disposes",
     all-or-nothing fail-closed merge, frozen) using the existing comprehension
     writer mechanism.

4. **Freeze and invalidate** lessons on `source_shell_sha256`, mirroring
   comprehension: a re-extract (shell changes) resets them, since they were learned
   about a specific shell.

5. **Consume at generate time.** The generator consults the lessons registry after
   the base resolver and before writing, applying learned re-points. No lessons ->
   identical to today's behavior.

### Why it speeds up the next run
The first generation may need a repair loop; the distilled lessons are cached on
the profile and reused across sessions, so the second generation starts from the
already-corrected profile and tends to pass QA on the first attempt.

### Compatibility notes
- Reuses `rules.overrides`, the comprehension freeze pattern, the `Finding` /
  `QAReport` structures, and `write_profile_json` (sorted keys). The new
  `generation_report.json` mirrors `visual_manifest.json`.
- Fail-closed guarantee preserved: every lesson is re-validated against the shell
  by `verify` (a lesson pointing at a now-missing artifact is dropped); lessons
  reference only shell-backed artifacts.
- Fully backward compatible (no lessons -> today's behavior).

### Risk to manage
Auto-applied corrections could entrench a wrong fix or mask a real authoring
problem. Mitigation: lessons carry provenance (which runs/findings produced them)
and confidence; they stay advisory until a threshold or an explicit accept
(mirroring `verify --accept`).

---

## 3. Interactive profile refinement from user feedback (human-in-the-loop)

### Problem
The Brand Profile only ever learns from what the engine can *measure*. After a
generation, the engine cannot tell whether a heading that **should** be the brand
teal came out black, whether a section it preserved is actually boilerplate the
user wants dropped, or whether it misread *what an extracted element is for*. Those
are qualitative judgements only the **user** can make. Today there is no moment in
the workflow where the skill asks, and no channel for that answer to improve the
profile, so the same off-brand deviation recurs on every future generation.

### Proposed design (compatible)
Add an explicit **ask-and-refine** step to the agent workflow, complementary to
the machine-driven feedback loop in section 2 (which learns from QA findings):

1. **Ask (agent guidance, `SKILL.md`).** The ask happens **only at the END of
   generation** - after the generated file and the QA summary are returned, never
   before or during. The skill explicitly tells the user that **feedback given as
   text OR as a screenshot image** can help better define the final result for
   **future** generations (it does not change the document just produced). A
   screenshot is a first-class multimodal input: the model can *see* a deviation
   from the template that deterministic QA cannot measure. The question is grounded
   in what was actually extracted (it can name the roles / palette entries /
   sections it used) and is concrete: *"Does this match your template? Is anything
   off-brand or deviating from the original - a heading that should be colored, a
   section that should not repeat, an element whose purpose I read wrong? Reply in
   text or attach a screenshot, and I'll fold it into the profile for next time."*

2. **Refine (model proposes).** The user's qualitative answer is turned by the model
   into structured refinements written **through the same fail-closed channels that
   already exist** - `comprehension.role_annotations` / `palette_annotations` /
   cover-slot and index conventions (the *purpose* of each extracted element), plus
   the shell-bound lessons/overrides registry from section 2 (a re-point within
   artifacts the shell already defines). The model can sharpen *what an element is
   for* and re-point within shell-backed artifacts; it can **never** invent a style,
   font, color, or layout the template does not contain.

3. **Confirm and freeze (deterministic disposes).** The proposed profile refinement
   is shown as a diff, the user confirms (mirroring `verify --accept`), and it is
   frozen to the shell (`source_shell_sha256`, like comprehension). `verify`
   re-validates every refinement against the shell, so a confirmed-but-impossible
   change is still refused fail-closed.

### Why it matters
This is how the skill **progressively learns to manage the extracted elements and
their scope**: each round enriches the comprehension's understanding of what each
role / color / section is *for*, so subsequent generations are more faithful with
less correction. Section 2 learns from what the machine can check; this learns from
what only the user can judge - together they close the loop from both sides.

### Compatibility notes
- Reuses the comprehension writer (the one model-writable, fail-closed, shell-frozen
  block), `role_annotations` / `palette_annotations`, and the section-2 lessons
  registry. No new brand-value channel: the model annotates purpose and re-points
  within shell-backed artifacts only.
- Fully optional and backward compatible: with no feedback the profile is unchanged
  and generation behaves exactly as today.
- Off-brand stays impossible by construction: the user's confirmation never lets a
  refinement reference an artifact the shell lacks - `verify` is still the floor.

### Risk to manage
A confident-but-wrong user answer (or an over-eager model interpretation) could
entrench a bad refinement. Mitigation: always present the change as a confirmable
diff, keep refinements advisory until accepted, and let a later re-extract reset
them (they are shell-bound).

---

## 4. Visual audit (salvaged future ideas)

Carried over from the now-removed `VISUAL_AUDIT_IMPROVEMENTS.md` (most of which
shipped: LibreOffice+Poppler+Pillow path, doctor preflight, PyMuPDF fallback,
optional Tesseract OCR, deep/strict modes, manifest diagnostics, the L2 repair
loop). Genuinely still-future items:

- **Renderer-disagreement cross-check.** ~~Compare `pdftoppm` vs PyMuPDF renders to
  flag rasterizer-specific artifacts.~~ **Deprioritized (vetted 2026-06-09):** the
  engine has only ONE layout engine (`soffice`); both `pdftoppm` and PyMuPDF
  rasterize the *same* PDF, so the only detectable disagreement is rasterization
  noise (anti-aliasing/hinting/DPI), not layout-fidelity defects. A real signal
  needs a second independent layout engine (native Word/PowerPoint), a heavy new
  dependency. Not worth the false-positive cost.
- **Richer image analysis.** Add `numpy` + `opencv-python`/`scikit-image` to move
  L1 from pixel proxies to bounding-box / connected-component / overlap heuristics
  and template-vs-output diff heatmaps.
- **OCR confidence scoring** and deeper stale TOC/field-cache detection
  (nested/multi-column TOCs, page-number-aware static entries).

---

## 5. Next-wave universal improvements (ideation synthesis, 2026-06-09)

A vetted brainstorm (run after v0.7.0) of how to make the skill smarter, more
faithful to a user's template, **universal across the 3 formats**, able to **learn
from its own errors**, and to use the **model** for more evaluation/fixing. Every
item below passed an adversarial gate on the two hard rules - **no fine-tuning on a
single template (universal mechanism only)** and **off-brand impossible by
construction** - and reuses an existing seam (resolver chokepoint, comprehension
writer, `rules.overrides`, QA `Finding`/`QAReport`, `visual_manifest`). 19 of 20
ideas survived; the 11 distinct buildable items cluster as below.

**Two foundations gate almost everything else:** a **shared cross-format apply
layer** (Cluster A) must precede per-format color/typography, and a **persisted
generation report** (B1) must precede every learning move. Land those two, and the
rest become small additive PRs.

### Build order (top 5) and load-bearing dependencies

1. **A1+A2+A3 - the cross-format appearance vertical** (capture -> shared apply core
   -> format-neutral verify). Closes the stated KEY GAP (appearance/color/caption
   apply is docx-only) and makes model-driven color universal for free. Ship A2+A3
   in the **same** release (apply-parity without verify-parity breaks fail-closed).
2. **B1 - persist `generation_report.json`** (feasibility S, purely additive):
   unblocks the entire learning cluster (B2/B3/B4), C3, and E3's downstream value.
3. **B3 - deterministic `learn` verb** writing `rules.overrides`: the first time the
   engine acts on its own errors. Needs B1+B2.
4. **C1 - persist the L2 visual-audit verdict**: the model's richest signal (it looks
   at the rendered PNGs) currently evaporates each run. Independent of A/B.
5. **D1 - paragraph-geometry appearance axis**: the biggest visible fidelity gain
   below typography, reusing the `_dominant` + `_merge_appearance` + family-gate
   machinery the A-cluster hardens.

Arrows: `A2 -> per-format color, E2, cross-format halves of D1/D2/D3`;
`A2 ships with A3`; `B1 -> B2 -> B3 -> B4` (strict chain); `B1 -> C3` and
`B1 -> E3`; `E1 -> off-theme accents reachable by D1/E2`. C1/C2/C3 are a parallel
model-in-the-loop track.

### Cluster A - Cross-format universality (the unblocker)

> **Status: SHIPPED (appearance + model-driven color).** A1+A2+A3 below are
> implemented: the docx capture/apply helpers were extracted into shared
> `common/typography.py` + `common/appearance.py` (docx stays byte-identical, proven
> by a frozen-hash anchor), pptx/xlsx now CAPTURE + APPLY appearance/color through a
> per-format backend protocol, and `check_appearance_targets` is format-neutral
> (per-kind shell-fact collectors). The model's `palette_annotations` naming is now
> load-bearing on all 3 formats. **Still docx-only: caption-index regeneration** (the
> `SEQ`/`refresh_visible_caption_index_cache` machinery lives only in
> `formats/docx/structure.py`), so that half of the original gap remains open.

Originally fixed the KEY GAP: `appearance` v2, model-driven color, and caption-index
regen APPLYed on docx only (`check_appearance_targets` hard-gated on `Kind.DOCX`;
`resolver.resolve_color`/`_merge_appearance` were already kind-agnostic but unconsumed
by pptx/xlsx).

| # | Item | Value | Seam reused | Feas | Why universal |
|---|---|---|---|---|---|
| **A1** | Per-format typography **capture** adapters | pptx/xlsx populate the same `role.appearance` + `theme.palette` shape docx fills, by walking placeholder runs / cell fonts under the existing `_dominant` floor | extract `_dominant`/`_color_obj`/`_palette_key` into `common/typography.py`; per-format `extract()` | M | pure dominance statistic over the template's own runs/cells; nothing privileged |
| **A2** | Shared run-branding **apply** core `common/appearance.py` | pptx/xlsx finally consume `op.appearance` + `resolve_color`, set-only-when-unset | extract `_brand_run_*` to a neutral core | M | core reads zero template specifics; applies only what the resolver hands it |
| **A3** | Format-neutral `check_appearance_targets` | one fail-closed verify proves every applied font/size/color against *each shell's own* facts, for all 3 kinds | lift the `Kind.DOCX` gate; per-kind shell-fact collectors | M | allowed set = whatever *this* shell proves, computed identically per kind |

A1 -> A2 -> A3; **A2 and A3 in the same release**. Side effect: the model's
`palette_annotations` naming (already inventory-parity on all 3 formats) becomes
load-bearing on pptx/xlsx for the first time - model-driven color goes universal
with no new model code.

### Cluster B - Learn-from-errors (new capability; refines section 2)

> **Status: B1+B2+B3+B4 SHIPPED** (see the section-2 banner; `qa/report.py` +
> `profile/overrides.py` + the `learn` verb + `check_override_targets`, plus the B4
> `propose-overrides` verb + `overlay_overrides` + the `generation_history` bundle
> slice + the `override_applied` audit finding). The whole B chain B1 -> B2 -> B3 ->
> B4 reuses the single `merge_overrides` sink.

When this was written nothing here existed: `generation_report` was absent and
`rules.overrides` was reserved (`schema.py`) with no reader/writer. Strict chain
B1 -> B2 -> B3 -> B4.

| # | Item | Value | Seam reused | Feas | Why universal |
|---|---|---|---|---|---|
| **B1** | Persist `generation_report.json` | serialize the `QAReport` (every `Finding` + shell sha + output/idoc hashes + ts) next to output | `visual_manifest` side-artifact pattern; generate verb | **S** | records check ids + shas, never template words |
| **B2** | Cross-run regression findings (`regression.recurred`/`reintroduced`) | the recurrence signal the `learn` verb thresholds on | prior `generation_report.json`; `run_qa`; `source_shell_sha256` freeze | **S** | compares `(check, location)` multisets across same-shell runs only |
| **B3** | Deterministic `learn` verb (Phase A) | distill unambiguous recurring findings into shell-bound `rules.overrides` re-points (reroute role / number_format / register_demo_clear) | `rules.overrides`; `resolve_role` last-resort; new `check_override_targets` mirroring `check_resolver_targets` | M | keys on stable check ids; re-points only to shell-defined artifacts (membership-checked) |
| **B4** *(shipped)* | Model-proposed corrections (Phase B) | model proposes corrections for the ambiguous remainder; `propose-overrides` overlays them onto any existing lesson (`overlay_overrides`) and routes the whole block through the SAME `merge_overrides` sink; a gate-wired `override_applied` INFO finding audits every live correction | bounded message-free `generation_history` bundle slice; single `merge_overrides` writer + `check_membership` | M-L | model only NAMES a shell-backed pointer; merge binds every proposal fail-closed, all-or-nothing |

### Cluster C - Model-in-the-loop (widen the 2 model touchpoints to 3-4)

Each persists a model judgement that is ephemeral today and re-validates it
fail-closed via the comprehension `merge` writer or a sibling check.

| # | Item | Value | Seam reused | Feas | Why universal |
|---|---|---|---|---|---|
| **C1** | L2 visual-audit verdict as a comprehension sub-block | persist per-checklist PASS/FAIL/NA + evidence; generate short-circuits the L2 round when all PASS at the current shell+content sha | additive `audit` sub-collection + `comprehension.merge`; new `check_audit_targets` | M | checklist ids derived purely from profile structure; model writes only a verdict against a structural id |
| **C2** | Model-assisted QA triage | model adjudicates ambiguous WARNINGs (blank_page/edge_bleed/component_survival) as EXPECTED->INFO or DEFECT; never silences an ERROR | additive `triage` map keyed by `Finding.check`+location; `run_qa` severity-fold | M | binds to the engine's closed set of check ids; disposition is a closed enum |
| **C3** | Interactive `refine` verb (see section 3) | turn a user's qualitative answer into a comprehension delta over existing sinks, confirm-as-diff | `comprehension.merge` (no schema change); `surface_inventories` binding | M | every binding is a verbatim `surface_inventories` id; a value the template never used cannot be named |

**C3 constraint (per the user, see section 3):** the feedback ask happens **only at
the end of generation**, invites **text or screenshot** feedback, and improves
**future** generations - never the just-produced file.

### Cluster D - Deeper template-following (fidelity below typography)

> **Status: D1 + D2 + D3 SHIPPED (DOCX-only).** Paragraph geometry, table-style
> conditional formats, and list/numbering definitions are now the 4th/5th/6th captured
> appearance axes, each riding the single resolver/appearance seam (NO family gate),
> applied set-only-when-unset, and verified by an honest fail-closed check
> (`check_geometry_targets` / `check_table_targets` / `check_numbering_targets`, gate-wired).
> Schema stays 1.2.0 (additive); pptx/xlsx are untouched (these are WordprocessingML-only
> constructs). The frozen anchor stays green (no-capture profiles generate byte-identically).

Extend the proven dominant-capture + resolver-chokepoint + fail-closed-check pattern
to new axes. Cross-format only once Cluster A's apply layer exists; on docx alone
they can ship independently.

| # | Item | Value | Seam reused | Feas | Why universal |
|---|---|---|---|---|---|
| **D1** *(shipped; tabs deferred)* | Paragraph geometry as a new appearance axis | capture+apply spacing / indent / `w:pBdr` / `w:shd` per role under the same `_dominant` floor (tabs `w:tabs` deferred - same independent-axis pattern when wanted) | `capture_fonts` model; `role.appearance.geometry` + `theme.geometry.body`; `_merge_appearance` (NO family gate - role geometry is intentional) ; `check_geometry_targets` (shape + observed-floor membership, not name-membership) | M | reads only this template's `w:pPr` twips/borders; keeps a value only when it dominates |
| **D2** *(shipped)* | Table-style fidelity (banding / `tblLook` / cell margins) | enable the template's own `w:tblStylePr` banding + first/last emphasis via the captured `tblLook` bitmask; KPI-as-table inherits free | docx table writer; `role.appearance.table` + `theme.table.body`; `_merge_appearance` (NO family gate) ; `check_table_targets` (gate-wired `appearance_table_targets`: `tblLook` shape + table-style name-membership + cell-margin observed-floor) | **S** | reads conditional-format facts the template declares; fills stay in the shell's style part |
| **D3** *(shipped)* | List/numbering definition fidelity | capture per-level numFmt / lvlText / indent; clone the shell's own `w:abstractNum` by id | `structure.py` numbering readers; `role.appearance.numbering` + `theme.numbering.body`; `_merge_appearance` (NO family gate) ; `check_numbering_targets` (gate-wired `appearance_numbering_targets`: num-id / abstract-num-id membership + numFmt shape + lvlText/indent observed-floor) | M | clones the numbering def this template references, by id; never synthesized from JSON |

### Cluster E - Faithfulness robustness (close known dead-ends)

> **Status: E1+E2+E3+E4 ALL SHIPPED - cluster complete.** Palette aliases, faked-heading
> promotion, the `appearance_apply_degraded` parity ledger, and universal cover synthesis
> for `kind==NONE` are all live. E4 landed DETERMINISTIC-only (no new comprehension sink):
> the synthesis triggers off the recorded `anchors.cover.kind == NONE` fact and builds
> only from `cover.*` roles that resolve through `resolve_role` (docx paragraphs before
> the first toc/body child; pptx cover slide on the role-resolved layout, including the
> reconcile path; xlsx N/A - named ranges carry no cover-page concept). Every synthesized
> cover is audited by an INFO `cover_synthesized` finding; the optional closed
> `cover_layout` model authorization was evaluated and deferred (nothing for the model to
> name that `resolve_role` does not already gate) and can land later as an additive
> opt-in. Byte-identical whenever the fact is absent or nothing resolves; schema stays
> 1.2.0.

| # | Item | Value | Seam reused | Feas | Why universal |
|---|---|---|---|---|---|
| **E1** *(shipped)* | Off-theme accent reachability via palette aliases | mint a syntactically-legal dotted token aliasing a model-named `hex:RRGGBB` palette entry, so off-theme brand accents become addressable run colors | `palette_annotations` -> `theme.palette` (new `_derive_palette_aliases` mints the byte-copied bridge token); `resolve_color` reads `ref` verbatim (ZERO resolver change); new L0 alias check `check_palette_alias_targets` (gate-wired `palette_alias_targets_exist`, in `DEFAULT_L0_INVARIANTS`) + `check_membership` syntax/collision gate | M | alias derived from the template's own captured palette entry; model proposes a name, engine copies the captured ref byte-identical |
| **E2** *(shipped)* | Faked-heading-in-body-style detection | surface body-style runs that are size/color outliers as a `pseudo_heading` fact (`common.typography.detect_pseudo_headings`, docx `capture_pseudo_headings` -> `theme.pseudo_headings`); model adjudicates a new closed `comprehension.promote_appearance` sink (`{pseudo_heading_ref, target_role_id}`) onto a real heading role | `comprehend_input_bundle` facts (`pseudo_headings`, present-only); new `promote_appearance` sink validated fail-closed by `check_promote_appearance` (ref surfaced + target a declared heading role + pair-unique) and derived onto the heading role appearance (`_derive_promote_appearance`); `_merge_appearance` (role-specific size/color, no family gate); EXISTING `check_appearance_targets` re-validates the promoted value shell-backed (no new check) | M | detector is a pure outlier test vs the captured dominant body appearance; model NAMES a surfaced ref + a declared heading role, engine copies the captured size/color; no promotion ⇒ byte-identical |
| **E3** *(shipped)* | Uniform `appearance_apply_degraded` finding | one stable INFO per (role, axis) whenever a format cannot realize a captured axis - makes parity gaps *measurable*, feeds the learning loop (B4 `generation_history`) and the L2 model | shared apply orchestration (`common/appearance.py`); per-backend `realized_axes` declaration; `Finding`/`QAReport` | **S** | fires on the structural fact that a captured axis was not realized; names only role id + axis; INFO-only, never flips a verdict |
| **E4** *(shipped; deterministic-only, `cover_layout` enum deferred)* | Universal cover synthesis for `AnchorKind.NONE` | build a cover from resolvable `cover.*` roles through `resolve_role` when no anchor exists; optional closed `cover_layout` enum lets the model authorize/order | `resolve_role` chokepoint; docx/pptx/xlsx cover write paths; QA finding | L | triggers off a structural fact (`kind==NONE`); builds only from inferred+validated roles; no-ops if none resolve |

### Vetted and cut (recorded)

- **Renderer-disagreement cross-check** - cut on feasibility (see section 4): one
  layout engine means `pdftoppm` vs PyMuPDF differ only in rasterization noise, not
  layout fidelity. Would need a second independent layout engine.

---

## 6. Performance (structural paths; profiled 2026-06-10)

A profiling pass (3 independent profilers: engine, QA, test-suite) landed seven
behavior-preserving quick wins (see CHANGELOG `[Unreleased]` / Performance:
injected visual seam in 4 tests, single-launch 3-format doctor probe, lazy
Office-lib imports, per-`run_qa` artifact-load memo, shared docx run-facts pass,
shared pptx layout/slide classification, class-scoped test fixture extraction).
Measured on the reference machine (M-series macOS, soffice + pdftoppm
installed): full suite **59.5s -> 24.0s**, real-render lane **41.4s -> 25.2s**,
`generate --qa auto` **8.1-8.3s -> 4.6-5.0s** per format, xlsx post-generate QA
pass **119ms -> 32ms**, `cli.py list` **0.29s -> 0.13s**.

What remains is dominated by ONE cost: every `--qa auto/deep` run pays a full
`doctor.probe()` render smoke-test (~2.5-3.0s post-quick-wins, was ~5.6s) plus
the real render (~1.1s, soffice startup-bound). The paths below would remove it,
but each one is caching / semantics-adjacent, so they are documented here -
NOT applied - until a maintainer signs off on the failure-mode tradeoffs.

- **Per-process memoization of `doctor.probe()` in
  `qa/visual.py:renderers_available()`.** Today it re-probes unconditionally on
  every in-process `run_qa(auto)`; `_LAST_RENDERER_STATUS` and
  `_reset_renderer_cache` already exist as the natural cache slot, and
  `RendererAvailabilityTest.setUp` already resets it. Measured: saved ~5.9s per
  additional in-process `run_qa(auto)` pre-quick-wins; the real-render lane runs
  5 identical full probes in one pytest process (~26s of its 41.4s before;
  ~10s residual after the single-launch probe). Why not applied: a renderer
  that vanishes mid-process would be reported as `render_failed` instead of
  `visual.unavailable` - an availability-semantics change.
- **Cross-CLI-invocation probe cache** keyed on soffice/pdftoppm path+mtime:
  would remove the residual probe cost from every `generate --qa auto`.
  Motivating number: a repeat byte-identical generate still paid 8.1s
  pre-quick-wins (~4.7s now) because `gate._l2_short_circuit` requires a
  persisted PASS audit verdict that plain model-free CLI runs never produce.
  Why not applied: a cache that outlives the process is exactly the class of
  state the fail-closed rules exclude; needs an explicit invalidation design.
- **Extend the L2 short-circuit to the model-free CLI flow** so byte-identical
  repeat generates skip probe+render (today it only fires with persisted
  `comprehension.audit` PASS rows). Numbers: repeat run ~4.7s vs ~0.4s of
  actual engine work. Why not applied: widens a carefully-gated skip path.
- **Skip the 3-format gate-time smoke probe entirely**, trusting the binary
  version probes (0.24s) plus the degrade-clean real render: the deep-generate
  floor would drop to ~2.7s. Why not applied: changes failure-mode findings on
  partially-broken environments (`visual.unavailable` vs
  `render_failed`+quicklook) - a semantics change, not an optimization.
- **Persistent soffice listener / render daemon, parallel probe+render.** Cost
  structure measured on this machine: each soffice launch = ~0.94s process
  startup + ~0.80s fresh `UserInstallation` creation + ~0.15s actual conversion
  of the real 3-page docx; one deep generate spawned 14 external processes of
  which only ~3 did real work. A daemon violates the no-persistent-state /
  no-platform-tricks bar today.
- **Reuse a warm `UserInstallation` for the REAL render in `qa/visual.py`**
  (the fresh per-render profile is a deliberate isolation choice): ~0.8s per
  render, at the cost of shared LibreOffice state across renders.
- **`sha256_file` recomputed 4x on the same shell per generate invocation**
  (`store.py` load-time shell check, `qa/report.py` shell+output hashes,
  `checks_deterministic.check_shell_provenance`): negligible at the 76KB
  example shell (~0.1ms each) but linear in template size - worth a per-pass
  memo (same scoping as the `run_qa` load memo) if multi-MB enterprise
  templates become the norm.
