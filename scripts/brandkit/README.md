<!-- SPDX-License-Identifier: MIT -->
# brandkit: the shared engine

One Python package powers all three skills (`brand-docx`, `brand-pptx`,
`brand-xlsx`). This is the internals map for contributors; the user-facing
contracts live in [CONVENTIONS.md](../../CONVENTIONS.md).

## Module layout

```text
brandkit/
|- cli.py            argparse entrypoint; the eleven verbs (see below)
|- doctor.py         environment probe (deps, renderers, OCR) + repair hints
|- profile/          the Brand Profile: the only home of brand facts
|  |- schema.py      FROZEN vocabulary owner (verbs, role ids, fields); 1.2.0 additive
|  |- store.py       profile directories (./brand-kit/<name> and ~/.claude/brand-kit/<name>)
|  |- resolver.py    THE chokepoint: the only reader of brand literals
|  |- blend.py       multi-template blending: same-format VALUE-fact merge (fail-closed)
|  |- comprehension.py  fail-closed single writer of the model's comprehension block
|  |- overrides.py   learn / propose-overrides / refine (ADVISORY until --accept)
|  `- reconcile.py   reconciles preserved cover/index structures with new content
|- ir/               IntermediateDocument: brand-agnostic typed flow blocks (docx/pptx)
|- grid/             GridDocument: named-region tabular input (xlsx)
|- formats/          per-format extract/generate, one package per kind
|  |- catalog.py     format registry: kind -> backend
|  |- docx/          reference implementation (+ cover, roles, styles helpers)
|  |- pptx/          masters/layouts/placeholders path
|  `- xlsx/          named-range fills, formula preservation
|- ooxml/            shared OOXML layer: pack (OPC zip), names, fields, chart, idempotency
|- common/           cross-format helpers: appearance merge (6 axes), color,
|                    typography, text, links
`- qa/               the gate: deterministic checks (L0), visual render (L1),
                     manifest for orchestrator judgement (L2), CHECK_REGISTRY
```

## The verbs and their call paths

`cli.py` exposes eleven verbs; each is a thin dispatcher over the packages above.

| Verb | Path through the engine |
|---|---|
| `extract` | `formats/<kind>/extract` reads the template -> `profile/store` writes profile + frozen shell + `PROFILE.md`; with `--blend`, folds a SECOND same-format template's value-facts into an existing profile via `profile/blend` (pointers never cross shells) |
| `comprehend-input` | `profile/comprehension` surfaces bounded facts + excerpt for the model (read-only) |
| `comprehend` | `profile/comprehension` validates the model JSON fail-closed and freezes it into the profile (single writer) |
| `verify` | `qa/gate` runs deterministic checks against profile + shell |
| `generate` | `ir`/`grid` parse input -> `profile/resolver` maps roles to captured facts -> `formats/<kind>/generate` writes from the shell -> `qa/gate` |
| `learn` | `profile/overrides` distills cross-run `generation_report.json` history into an ADVISORY lesson (live only with `--accept`) |
| `propose-overrides` | `profile/overrides` validates a model-authored proposal fail-closed (ADVISORY until `--accept`) |
| `refine` | `profile/overrides` folds a model-authored refinement into the comprehension block (diff preview; `--accept` persists) |
| `compare-profiles` | `profile/compare` diffs the BRAND-level facts of two saved profiles (read-only; exit 1 on drift) |
| `list` | `profile/store` enumerates saved profiles per scope |
| `doctor` | `doctor.probe()` reports what the environment can do |

## The two invariants everything hangs on

1. **The inversion.** The model only ever NAMES facts the extractor captured;
   the deterministic engine is the only author of values. Every model-writable
   sink validates shape + verbatim membership against surfaced inventories in
   one all-or-nothing transaction, then stamps `source_shell_sha256` so a stale
   cache can never apply to a different shell.
2. **Byte-identity as regression armor.** Generation is deterministic, and the
   no-capture path is frozen by a hash anchor test. Any new feature must prove
   it moved nothing when absent. This is what makes fast iteration safe.

Practical corollary for new code: if you find yourself writing a style name,
hex color, or font literal anywhere outside `profile/`, stop; that fact must be
captured at extract time and resolved through `profile/resolver.py`.

## QA in one paragraph

`qa/gate.py` orchestrates three levels: L0 deterministic checks (membership,
floors, idempotency, formula preservation; ids frozen in `qa/model.py`'s
`CHECK_REGISTRY`), L1 rendering via LibreOffice + Poppler into page PNGs, and
L2, a `visual_manifest.json` the orchestrating agent judges. Checks are honest:
a check only claims what it can prove (name membership where names exist,
sanity + observed floors where they do not), and missing renderers degrade the
verdict instead of faking a pass (`--qa strict` refuses to run degraded).

## Adding a capture axis (the proven recipe)

1. Capture the fact in `formats/<kind>/extract` (dominance floors, not single
   samples) and store it under the existing schema (additive only).
2. Merge it in `common/appearance.py` as its own axis; declare it in each
   backend's `realized_axes`.
3. Apply it in `formats/<kind>/generate` set-only-when-unset, through the
   resolver.
4. Add the honest fail-closed check in `qa/` and register it.
5. Prove byte-identity when the template lacks the fact (the anchor test must
   not move), then add fixtures that exercise it.
