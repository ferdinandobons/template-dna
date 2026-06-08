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
python scripts/brandkit/cli.py comprehend-input --name <brand>

# 2. Merge + validate + cache the model-authored comprehension.json (THE ONLY WRITER).
python scripts/brandkit/cli.py comprehend --name <brand> --input comprehension.json
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

## The four questions (the SAME for every format)

Reasoning over the **structure** in the bundle, answer exactly these four
questions. They are format-neutral: a cover slot is a cover slot, a derived index
is a derived index, demo content is demo content, whether the file is a Word
document, a PowerPoint deck, or an Excel workbook.

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
   `conventions.indexes`.
4. **Which cover slot is which?** For each surfaced cover anchor, which content
   slot fills it (`binds_to`), and should it be filled in place, cleared, or left
   alone (`fill_rule`)? Record it in `cover_slots`.

## The anti-overfitting directive (state and obey verbatim)

> A title slot is a title slot whether its placeholder reads "Titolo", "Title",
> or "Titre". Quote a literal **only as evidence**, never as a matching rule.
> Reason over structure (builtin style ids, field codes, placeholder types, named
> regions, geometry), not over the words a particular template happens to use.

Concretely, in the comprehension JSON:

- Every **load-bearing reference** - a `cover_slots` key (`anchor_ref`), an
  `index_ref`, a `region_ref`, a `feeds_from_role_id`, a `role_annotations` key -
  **must be a verbatim id copied from the facts bundle**. If an id is not in the
  surfaced inventory, do not invent it; the merge is fail-closed and will reject
  it (a ref into an empty inventory is itself an error).
- Only four fields are closed enums, and each maps to a real engine branch:
  `status` (`present|absent|rejected`), `fill_rule` (`in_place|clear|leave`),
  `reconcile` (`regenerate|preserve|clear`), `verdict` (`demo|real|mixed`). Use
  exactly those values.
- Every other field (`semantic_role`, `kind`, `purpose`, `generation_rules`,
  `evidence`, region names) is an **open advisory token**. The generator never
  pattern-matches on it, so write it honestly for a human reader; never bend it
  to fit a fixed vocabulary.

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
        "feeds_from_role_id": "heading.1", "reconcile": "regenerate" }
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
  }
}
```

## Format readiness

Comprehension binds only to the ids a format's extractor actually surfaces. All
three Office extractors expose cover anchors and regions when the template
contains them. Derived-index fields are format-specific: DOCX can expose TOC/list
fields, PPTX can expose an agenda/section-list field when present, and XLSX keeps
`fields` intentionally empty because workbooks have no TOC-style field code. The
model writes only the parts it has ids for; never force a cover, index, or region
shape onto an empty inventory; a ref into an empty inventory is fail-closed and
will be rejected.
