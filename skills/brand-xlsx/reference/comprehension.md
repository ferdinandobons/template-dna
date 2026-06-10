<!-- SPDX-License-Identifier: MIT -->
<!-- SHARED BODY: skills/{brand-docx,brand-pptx,brand-xlsx}/reference/comprehension.md
     are BYTE-IDENTICAL. The comprehension guidance is format-neutral by design (a
     cover slot is a cover slot whether the file is a Word document, a PowerPoint
     deck, or an Excel workbook), so the same prompt drives all three formats and
     cannot drift. A CI guard (tests/test_reference_sync.py) asserts byte-identity.
     Edit all three together, or run the guard's --write helper. -->

# Comprehension (optional, model-driven)

`comprehend` is the third of the four verbs (**extract / comprehend / verify /
generate**). It is **optional** and **model-driven**: it is the one step where the
model (you) reads a bounded bundle of deterministic facts about the template and
writes down what each structure is **for**, so that `generate` can reconcile the
template's preserved cover/index structures with the new content instead of
leaving a cover unfilled, appending a duplicate title, or carrying stale index
entries from removed demo content.

Everything load-bearing the model writes is a **verbatim id** copied from the
facts bundle; the deterministic validator rejects anything else. The model
**proposes**, the engine **disposes**. `generate` still works with no
comprehension at all (the deterministic path), so skip this verb whenever a
current comprehension already exists or no model is available.

## When to run it

Run `comprehend` once per template shell, right after `extract`, **before**
`generate`. Skip it when a current comprehension is already cached, that is,
when the profile's `comprehension.status` is `present` **and** its
`source_shell_sha256` equals the live `provenance.shell.sha256`. A re-extract
produces a new shell and resets `comprehension.status` to `absent`, so re-run
`comprehend` only then. Never re-run it at `generate` time: the understanding is
frozen into `profile.json` at merge and reused byte-for-byte.

## The two CLI steps

`comprehend` is realized as two internal engine steps. Treat them as engine
plumbing, not a user-facing workflow.

```bash
# 1. Print the bounded {facts, excerpt} bundle for the model to read.
python scripts/cli.py comprehend-input --name <brand>

# 2. Merge + validate + cache the model-authored comprehension.json (THE ONLY WRITER).
python scripts/cli.py comprehend --name <brand> --input comprehension.json
```

`comprehend-input` prints `{facts, excerpt}`: `facts` is the surfaced,
format-uniform inventory the model reasons over and binds to (`inventories`:
`cover_anchors`, `fields`, `regions`, `roles`, each entry an `{"id": ...}`),
plus `structure`/`anchors`/`styles` slices; `excerpt` is an ordered,
length-capped sample of text the extractor already collected. **Read this
bundle. Never parse the raw template or its OOXML.**

`comprehend` is the single writer of the `comprehension` block. It re-runs the
full schema validation **and** a fail-closed membership check, and only on a
clean pass freezes the block into `profile.json` with `status='present'`,
stamping `source_shell_sha256` from the live shell hash. On any finding it writes
`status='rejected'` with the findings and exits non-zero; read the findings, fix
the offending refs, and retry. A clean merge is idempotent: running `comprehend`
twice yields a byte-identical `profile.json`.

## The six questions (the SAME for every format)

Reasoning over the **structure** in the bundle, answer exactly these six
questions. They are format-neutral: a cover slot is a cover slot, a derived index
is a derived index, demo content is demo content, a brand color is a brand color,
whether the file is a Word document, a PowerPoint deck, or an Excel workbook.

1. **What is each element FOR?** For each surfaced role and cover anchor, what is
   its purpose in this template? Annotate roles in `role_annotations` and name
   each cover slot's `semantic_role`/`purpose`.
2. **What is demo vs real?** Which surfaced regions hold placeholder/sample
   content the author is meant to replace, and which hold real, structural
   content to keep? Record each in `demo_classification` with a `verdict`.
3. **Which artifacts feed which index?** For each surfaced derived index (a table
   of contents, a list of tables/figures, an agenda/section list, anything the
   document regenerates from its own content), which role's items feed it, and
   should it be regenerated, preserved, or cleared? Record it in
   `conventions.indexes`. For a **caption index** (a list of tables or figures,
   identified by a non-null `seq_id`), also set `caption_target` (`table` or
   `figure`): it tells `generate` which captionable kind feeds the index, so it can
   emit the matching numbered `SEQ` field on each caption and regenerate the index
   from the new content. Without it a caption index can only be preserved or cleared
   (never repopulated), so its cache would keep the template's stale entries.
4. **Which cover slot is which?** For each surfaced cover anchor, which content
   slot fills it (`binds_to`), and should it be filled in place, cleared, or left
   alone (`fill_rule`)? Record it in `cover_slots`.
5. **Which content shapes recur?** Do the same few primitive blocks repeat as a
   unit across the template (a callout pattern, a titled card, a standard section
   opener)? `fragment_candidates` in the bundle hints at recurrences; trust your
   own reading of the `excerpt` too. For each genuine recurrence, propose a
   reusable fragment in `fragments`: a `ref`, a `kind` (`component` for a single
   inline fragment, `section` for a multi-block unit), a `purpose`, and a `blocks`
   template built ONLY from primitive block types. Put a `{{slot}}` token wherever
   the text varies per use; the author fills it via the referencing block's
   `slots`. Propose nothing when no shape genuinely recurs.
6. **What is each brand color FOR / where does the template use it?** The bundle's
   `palette` lists every brand color the extractor observed (a theme slot like
   `accent1`, or an off-theme `hex:RRGGBB`), each with its captured `ref`,
   `provenance` (where it was seen), and a coarse `frequency`. For each color, what
   role does it play in this template (a `name`, a `purpose`, a `use_when`, an
   optional `semantic_role`)? Record it in `palette_annotations`, keyed by the
   palette id. You **name** a color; you never author one - the deterministic
   capture owns the actual `ref`/hex, and a key into an empty/absent `palette`
   inventory is fail-closed. Annotate nothing when the template carried no color.
   An optional `alias` field on an annotation names a dotted-token alias for the
   captured color entry, minted at merge time with the captured `ref` byte-copied -
   so an off-theme `hex:RRGGBB` accent becomes addressable as a clean run-color
   token. The alias must be a syntactically-legal dotted token (the same lowercase
   `family.qualifier` grammar as a role id) that does not collide with an existing
   palette key or role id; you name it, the engine copies the captured `ref`.

## The anti-overfitting directive (state and obey verbatim)

> A title slot is a title slot whether its placeholder reads "Titolo", "Title",
> or "Titre". Quote a literal **only as evidence**, never as a matching rule.
> Reason over structure (builtin style ids, field codes, placeholder types, named
> regions, geometry), not over the words a particular template happens to use.

Concretely, in the comprehension JSON:

- Every **load-bearing reference** - a `cover_slots` key (`anchor_ref`), an
  `index_ref`, a `region_ref`, a `feeds_from_role_id`, a `role_annotations` key, a
  `palette_annotations` key (a palette id) - **must be a verbatim id copied from
  the facts bundle**. If an id is not in the surfaced inventory, do not invent it;
  the merge is fail-closed and will reject it (a ref into an empty inventory is
  itself an error).
- Eight fields are closed enums, and each maps to a real engine branch:
  `status` (`present|absent|rejected`), `fill_rule` (`in_place|clear|leave`),
  `reconcile` (`regenerate|preserve|clear`), `verdict` (`demo|real|mixed`), a
  `fragments` entry's `kind` (`component|section`), a caption index's
  `caption_target` (`table|figure`), an `audit` row's `verdict`
  (`PASS|FAIL|NA`), and a `triage` entry's `disposition` (`expected|defect`). Use
  exactly those values.
- Every other field (`semantic_role`, an index's `kind`, `purpose`,
  `generation_rules`, `evidence`, region names) is an **open advisory token**. The
  generator never pattern-matches on it, so write it honestly for a human reader;
  never bend it to fit a fixed vocabulary.

## Worked example (frozen role-ids + non-language placeholders only)

The example uses frozen role-ids (`heading.1`, the kind of id that appears in
`feeds_from_role_id` and `role_annotations`) and non-language placeholder ids
(`<slot-1>`, `<index-A>`, `<region-1>`, standing in for the surfaced
cover-anchor / field / region ids) on purpose: it must carry **zero** plausible
word from any human language, so it can never become a matching rule. Replace
each `<...>` with a verbatim id from your own bundle.

```jsonc
{
  "confidence": 0.86,
  "cover_slots": {
    "<slot-1>": { "semantic_role": "title", "purpose": "main cover title",
                  "binds_to": "title", "demo_value": "<captured placeholder text>",
                  "fill_rule": "in_place" }
  },
  "conventions": {
    "indexes": [
      { "index_ref": "<index-A>", "kind": "outline", "seq_id": null,
        "feeds_from_role_id": "heading.1", "reconcile": "regenerate" },
      { "index_ref": "<index-B>", "kind": "caption", "seq_id": "<seq>",
        "feeds_from_role_id": "caption", "caption_target": "table",
        "reconcile": "regenerate" }
    ],
    "sections": [ { "region_ref": "<region-1>", "required": true, "repeatable": false } ]
  },
  "role_annotations": {
    "heading.1": { "purpose": "top-level section heading",
                   "generation_rules": "one per major section" }
  },
  "demo_classification": {
    "regions": [ { "region_ref": "<region-1>", "verdict": "real",
                   "evidence": "structural content region, not sample text" } ]
  },
  "fragments": [
    { "ref": "<fragment-1>", "kind": "component", "purpose": "recurring titled note",
      "blocks": [ { "type": "callout", "intent": "note",
                    "runs": [ { "t": "{{body}}" } ] } ] }
  ],
  "palette_annotations": {
    "<color-1>": { "name": "primary brand color", "semantic_role": "accent",
                   "purpose": "headings and key emphasis",
                   "use_when": "drawing attention to a section title",
                   "alias": "accent.primary" }
  }
}
```

The `<color-1>` key is a verbatim **palette id** from the bundle's `palette`
inventory (a theme slot like `accent1`, or `hex:RRGGBB`). The `name` / `purpose` /
`use_when` / `semantic_role` are advisory free text you write for a human; you
**never** author a color value - the deterministic capture owns the actual
`ref`/hex, and the merge mirrors your names onto `theme.palette` without ever
touching the captured color. The optional `alias` (here `accent.primary`) is a
dotted-token alias for the captured color entry, minted at merge time with the
captured `ref` byte-copied for reachability as a run-color token; it must be a
syntactically-legal dotted token that collides with no existing palette key or role
id, and again you only name it - the engine copies the captured `ref`.

## Format readiness

Comprehension binds only to the ids a format's extractor actually surfaces. All
three Office extractors expose cover anchors and regions when the template
contains them. Derived-index fields are format-specific: DOCX can expose TOC/list
fields, PPTX can expose an agenda/section-list field when present, and XLSX keeps
`fields` intentionally empty because workbooks have no TOC-style field code. The
model writes only the parts it has ids for; never force a cover, index, or region
shape onto an empty inventory; a ref into an empty inventory is fail-closed and
will be rejected.

Reusable fragments are format-neutral too: any format may carry a `fragments`
proposal, and each lands in the profile's `components` / `sections` registry for
`generate` to inline. A fragment's `blocks` must be primitive block types (never a
style, color, or layout), so a proposed fragment resolves through the same brand
chokepoint as inline content and cannot be off-brand. Propose fragments only when
a shape genuinely recurs; an empty `fragments` list is the norm.

## Persisting the visual-audit verdict (the `audit` map)

After a `generate --qa deep|strict` run, the engine writes a `visual_manifest.json`
side artifact listing the rendered PNG pages, the exact-artifact `shell_sha256` /
`content_sha256`, and a profile-derived `checklist` (each item has an `id`). You
(the model) open the pages and judge each checklist item. You MAY persist that
judgement back into the comprehension as an `audit` map, keyed by the manifest's
`checklist[*].id`, then merge it via `comprehend`:

```jsonc
"audit": {
  "<checklist-id>": {
    "verdict": "PASS",                       // PASS|FAIL|NA (closed enum)
    "evidence": "cover title is bound, no residual demo prompt",
    "shell_sha256": "<manifest.shell_sha256>",
    "content_sha256": "<manifest.content_sha256>"
  }
}
```

The `audit` key MUST be a verbatim `checklist[*].id` from the manifest; a key that
is not a current checklist id is fail-closed and rejected at merge (same rule as
every other load-bearing ref). `verdict` is the closed `PASS|FAIL|NA`; `evidence`
is advisory free text for a human; the two shas are copied verbatim from the
manifest so the verdict is scoped to that exact artifact. You never write a brand
value here, only a disposition against a structural checklist id the engine derived.

This makes the next `generate` cheaper: a same-shell / same-content `generate`
**short-circuits** the L2 render+manifest round when every current checklist item
PASSes at the matching `(shell_sha256, content_sha256)`. Any `FAIL`, any `NA`, a
newly-derived checklist id with no row, or a sha mismatch forces a full L2 round,
so the short-circuit can never hide a regression. It is disabled under `--qa
strict` (which always re-renders) and never fires at `verify`. An empty `audit`
map is the norm until you have judged a render.

## Triaging an ambiguous QA warning (the `triage` list)

Some QA checks emit a **WARNING** because the signal is genuinely ambiguous: a
full-bleed cover that the edge-bleed proxy flags, a deliberately blank section
page, or a native component family the new content legitimately has fewer of. When
you (the model) have judged that such a WARNING is an EXPECTED property of this
template - not a defect - you MAY persist that judgement as a `triage` list,
merged via `comprehend`:

```jsonc
"triage": [
  {
    "check": "visual.edge_bleed",            // expected|defect maps the disposition
    "location": "page:1:bottom",             // the finding's exact location (or null)
    "disposition": "expected",               // expected -> demote that WARNING to INFO
    "evidence": "the cover is intentionally full-bleed in this brand"
  }
]
```

Only three checks are triage-eligible, and they are ALL WARNING-only:
`visual.blank_page`, `visual.edge_bleed`, and `component_survival`. A `triage`
entry naming any other check is fail-closed and rejected at merge, and two entries
for the same `(check, location)` pair are rejected as ambiguous. `disposition` is
the closed `expected|defect`; `location` is copied verbatim from the QA finding (it
may be null); `evidence` is advisory free text for a human.

The effect is deliberately narrow and one-directional: an `expected` entry demotes
exactly the matched WARNING to INFO (the message keeps a `(triaged EXPECTED: ...)`
note); a `defect` entry, or any entry that does not match a finding, leaves the
gate verbatim. **A triage entry can NEVER turn an ERROR into a warning, nor raise a
severity** - the eligible checks are WARNING-only and the demotion only ever lowers
a WARNING to INFO, so a real failure can never be silenced. An empty `triage` list
is the norm.

## Promoting a faked heading onto a real heading role (the `promote_appearance` list)

Some templates **fake** a heading: a line that LOOKS like a heading - visibly larger
and/or in a brand color - is authored with the **body** paragraph style (no heading
role, no named heading style), so the deterministic engine treats it as body and the
brand heading look is lost. The extractor's pure-deterministic detector surfaces each
such body-style run as a `pseudo_heading` fact in the bundle when its captured size or
color is a clear **outlier** vs the dominant body appearance. Each fact carries a
stable structural `ref`, the run's OWN captured outlier `size_hp` (half-points) and/or
`color`, and coarse, brand-text-free `evidence` (e.g. `"size 44hp vs dominant body
24hp"`). The list is absent when the detector found no outlier.

When you judge that a candidate is a real heading the template faked, you MAY persist a
`promote_appearance` list, merged via `comprehend`:

```jsonc
"promote_appearance": [
  {
    "pseudo_heading_ref": "body_run_3",     // a verbatim ref from facts.pseudo_headings
    "target_role_id": "heading.1"           // a declared heading.* role to promote onto
  }
]
```

You **name** only the two ids; you **never** author a size or color. On a clean merge
the engine COPIES the captured outlier `size_hp`/`color` from the detector fact onto
`roles[target_role_id].appearance`, so the generated heading carries the size/color the
template proved it uses. The `pseudo_heading_ref` MUST be a ref the detector surfaced
(an empty/absent `pseudo_headings` inventory fails closed), `target_role_id` MUST be a
declared `heading.*` role (a non-heading target is rejected), and a `(ref, target)` pair
must be unique. The promoted size/color is re-validated **shell-backed** at QA exactly
like any applied appearance: a value the shell does not carry is an ERROR, so the engine
can never inject a size/color the template lacks. An empty `promote_appearance` list is
the norm.

## Refining the understanding from user feedback (the `refine` verb)

After a generation, the user may give qualitative feedback - in **text** or as a
**screenshot** of the produced file. You turn that answer into a small,
**structured refinement delta** of verbatim ids and merge it with the `refine`
verb:

```bash
# Overlay the delta onto the EXISTING comprehension, then re-validate the whole block.
python scripts/cli.py refine --name <brand> --input refinement.json
# Add --accept to persist the refined comprehension (else the diff is previewed).
python scripts/cli.py refine --name <brand> --input refinement.json --accept
```

A refinement delta touches ONLY the existing qualitative-understanding sinks -
`role_annotations`, `palette_annotations`, `demo_classification`, `cover_slots`,
and `conventions` (`indexes` / `sections`). It is **not** a schema change and it
never writes a brand value: a palette annotation carries only `name` / `purpose` /
`use_when` / `semantic_role` (the `ref` / hex stay the deterministic capture's),
and every ref is still a **verbatim id** from the surfaced inventories.

```jsonc
"palette_annotations": {
  "<color-1>": { "name": "primary brand", "use_when": "section headings" }
},
"role_annotations": {
  "heading.1": { "purpose": "section titles", "generation_rules": "keep terse" }
}
```

The verb **overlays** the delta onto the present block (replacing or adding the
named sink entries; list sinks merge by their ref, never naive concat), then routes
the WHOLE combined block back through `comprehend`'s single fail-closed writer: the
full schema + membership validation re-runs and every ref re-binds to
`surface_inventories`, so a delta naming an id that is not surfaced (or one in an
empty inventory) is rejected and nothing is written. Without `--accept` the
post-overlay diff is previewed and the prior block stays authoritative; `--accept`
persists it and re-stamps `source_shell_sha256`.

A **screenshot** is YOUR multimodal read of the produced file; the engine only ever
ingests the resulting structured JSON delta of verbatim ids - it never sees the
image. The feedback ask happens **only after** the file and its QA summary are
returned, never before or during generation, and a refinement improves **FUTURE**
generations - it never re-emits or edits the file you just produced.

## Proposing overrides corrections from recurring findings (the `propose-overrides` verb)

`refine` sharpens the qualitative *understanding*; `propose-overrides` is its peer
for **shell-bound corrections** - the model-assisted sibling of the deterministic
`learn` verb. When the same QA finding keeps recurring across runs, the engine's
deterministic `learn` distils the ones it can bind to a brand-safe target on its own
(a stub role with a healthy same-family sibling, a captured demo string). The
**ambiguous remainder** - a stub role with no sibling, a finding whose right re-point
needs judgement - is surfaced to you in the `comprehend-input` bundle under
`facts.generation_history`: a bounded, **message-free** list of
`{check, location, severity, recurred_runs}` (the universal `(check, location)`
identity only, never the finding's message text). You reason over it and author a
small overrides proposal, then merge it:

```bash
# Overlay the proposal onto any existing lesson, then re-validate the whole block.
python scripts/cli.py propose-overrides --name <brand> --input overrides.json
# Add --accept to make the correction LIVE (else it is written advisory-'absent').
python scripts/cli.py propose-overrides --name <brand> --input overrides.json --accept
```

A proposal may ONLY **NAME a shell-backed pointer** - never author a style, font, or
color:

```jsonc
// reroute a role whose resolver is a dead stub to an EXISTING healthy role:
"reroute_roles": { "heading.9": "heading.1" },
// swap a number_format MASK the shell already uses (xlsx):
"number_format_swaps": { "<role-1>": "#,##0" },
// register a CAPTURED demo string for clearing:
"demo_clears": ["<captured-demo-value>"]
```

The verb **overlays** the proposal onto any present lesson (additive: a deterministic
`learn` lesson and your proposal coexist, your entry winning a key collision) and
routes the WHOLE combined block through the single `merge_overrides` writer: the
target role must be a **declared, shell-backed** role, the mask must be one the shell
uses, and the demo value must have been captured - any unbound pointer rejects the
WHOLE proposal (all-or-nothing), so an off-brand correction is impossible by
construction. Without `--accept` the correction is written but kept OUT of the live
resolver (`status='absent'`, byte-identical generation); `--accept` makes it LIVE and
re-stamps `source_shell_sha256`, so a re-extract (new shell) invalidates it.

Like the feedback ask, this happens **only after** generation and improves **FUTURE**
generations only. Every LIVE correction is auditable: the gate emits an INFO
`override_applied` finding for each one (in `generate` and `verify` alike), so a
learned re-point is never silent.
