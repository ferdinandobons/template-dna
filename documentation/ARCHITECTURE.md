# Architecture & how it works

> **The core guarantee: off-brand output is impossible by construction.** No
> generator ever writes a literal style name, hex color, or font: those live only
> in the Brand Profile, and `verify` refuses a profile that points at anything the
> template doesn't contain. There is no "creative" path that drifts from the brand.

---

## Highlights

- 🎯 **Brand-faithful by construction**: generation opens from the real template shell and applies only its named styles, theme colors, fonts and layouts. The content model is brand-agnostic; the **Brand Profile** is the single source of brand truth.
- 🧠 **Extract once, reuse forever**: a portable `brand-kit/<name>/` is the template's memory; every later document reads it. No re-explaining the brand.
- 🏛️ **Structure-aware**: captures the template's *ordered skeleton* (e.g. **cover → table of contents → body**) and tags each component as a **fixed structure to keep in order** or a **style to use on demand**. ([details](#structure-aware-not-just-style-aware))
- ✅ **Enforced, not just promised**: `verify` opens the shell and **fails** if a role resolves to a style/layout/named-range that doesn't exist. Deterministic checks also cover allowed styles, palette adherence, residual template text, broken tables, **formula preservation** (Excel), native-component survival and language rules.
- 🧪 **Auditable generation**: `doctor` preflights dependencies, `--qa fast|auto|deep|strict` makes QA depth explicit, and `--qa deep`/`strict` writes a visual manifest for render-based review and targeted repair.
- 🧩 **One shared engine**: a single profile schema, resolver, OOXML layer and QA gate underpin all three formats. The Word vertical is the reference implementation; PowerPoint and Excel build on the same foundation.
- 🗂️ **Full artifact catalog**: records OOXML parts, styles, media, layouts, formulas and named ranges, so an agent can reason about anything the template exposes, even artifacts it can't yet regenerate.
- 🔓 **Self-contained & MIT**: pure `python-docx` / `python-pptx` / `openpyxl` + OOXML. No cloud, no external services, no vendor lock-in.

---

## Why not just ask an AI to "use this template"?

General-purpose document skills generate *freely* and only loosely imitate a
reference file: fonts drift, the palette wanders, the corporate structure is
lost. BrandDocs is the opposite: narrow and faithful. It learns the template as a
set of **rules and reusable parts**, remembers them in a `brand-kit/`, and
**respects them** across an unlimited number of documents.

|  | General-purpose Office skills | **BrandDocs** |
|---|---|---|
| Mental model | "create a nice document" | "fill the company's template" |
| Brand fidelity | best-effort imitation | **by construction**: opens from the shell, applies only its artifacts |
| Reusability | re-explain the brand every time | **extract once**, reuse forever via `brand-kit/` |
| Structure | free-form | **respects the template's cover → contents → body order** |
| Guardrails | none | `verify` fails if a profile references a missing style/layout/range |

---

## The pipeline

```
 company template ──▶ ① EXTRACT ──▶ brand-kit/<name>/ ──▶ ② GENERATE ──▶ on-brand document
   .docx/.pptx/.xlsx      │          (profile + shell)         │
                          │                                    ├─ opens FROM the template shell
                          ├─ theme colors & fonts              ├─ resolves semantic blocks → brand styles
                          ├─ named styles → roles              ├─ keeps the template's structure order
                          ├─ document structure (skeleton)     └─ runs the QA gate
                          ├─ layouts / cover anchors
                          ├─ logos, media, tables, formulas
                          └─ full artifact catalog
```

> **Same format throughout.** A `.docx` template yields `.docx` documents, a
> `.pptx` yields `.pptx`, a `.xlsx` yields `.xlsx`. Each skill is its own lane:
> there is no cross-format conversion.

1. **Extract** unpacks the template's OOXML and records its brand: theme, named
   styles mapped to semantic **roles**, the **document structure** (the ordered
   skeleton plus which parts are fixed vs free), layouts, cover anchors, logos, and
   a complete artifact catalog. The original file is kept **byte-for-byte** as the
   *shell*.
2. **Generate** turns your content into an **IntermediateDocument** of
   brand-agnostic typed blocks (heading, paragraph, callout, list, table, …). A
   **pure resolver** maps each block to the concrete brand artifact from the
   profile, fills the shell **in the template's structural order**, and saves.
3. **Verify / QA** runs deterministic checks (every role resolves to a real
   artifact, only allowed styles are used, the palette holds, no residual template
   text remains, tables are intact, **Excel formulas survive every region fill**)
   and, when LibreOffice is available, a render-based visual pass. When renderers
   are unavailable, the audit degrades explicitly instead of pretending visual
   proof happened.

The full end-to-end agent workflow (skill selection, `doctor` preflight,
extract/comprehend/generate/QA, visual manifests, autonomous repair rounds, and
the final clean-output criteria) is documented in
[`PLUGIN_WORKFLOW.md`](PLUGIN_WORKFLOW.md).

---

## Structure-aware, not just style-aware

Most "use my template" tools copy *styling*. BrandDocs also learns the template's
**document structure** and reproduces it. During extraction it detects the ordered
skeleton (typically **cover → table of contents → body**) and annotates every
captured component with **how it is used**:

- **Structural** parts (cover, table of contents) are kept **in order** in every
  generated document: the cover is filled in place, the TOC is preserved and
  refreshed.
- **Freeform** parts (headings, callouts, lists, tables, quotes, captions) are
  styles to **use on demand**, in whatever order your content needs.

So a generated report opens with the company cover, keeps a live table of
contents, and fills **only the body** with your content, exactly like a person
starting from the corporate template, rather than a bare wall of text.

---

## Reliability & repair loop

BrandDocs is designed so reliability is visible, testable, and improvable. The
agent should not just produce a file; it should know which guarantees were proven,
which were degraded, and what to repair next.

| Layer | What it proves | What happens on failure |
|---|---|---|
| **Preflight** | `doctor` checks required Python packages, optional renderers (`soffice`, `pdftoppm`, PyMuPDF/`fitz`), and optional OCR (`tesseract`) before work starts. | Missing required packages must be installed/repaired. Missing visual/OCR tools downgrade only that proof layer. |
| **L0 deterministic QA** | Schema validity, resolver targets, allowed styles/layouts/ranges, residual demo text, markdown leaks, structural diffs, formula preservation. | The gate fails or emits explicit findings before the output is treated as clean. |
| **L1 visual proxies** | Rendered-page signals such as blank pages, zero pages, content near page/slide edges, and optional OCR hits for visible residual template text. | Findings are warnings because the engine can detect symptoms, not intent. |
| **L2 visual judgement** | The orchestrator opens the PNGs from `visual_manifest.json`, judges checklist items, and decides whether the result is visually acceptable. `strict` turns unclean visual evidence into gate errors. | Apply a targeted repair, regenerate, and rerun `--qa deep` or `--qa strict` until clean or honestly blocked. |

The template is treated as a source of reusable brand affordances, not a script to
preserve blindly. If an inherited section break, slide scaffold, print area, field
cache, or named-region geometry creates blank pages, stale entries, overlap, or
clipped output, the right move is to diagnose the cause and make the smallest
targeted composition change. Preserving a broken structure is less important than
producing a clean branded document.

The most valuable next reliability improvements are:

1. **Deeper PPTX/XLSX fidelity**: native object authoring already covers tables
   (with merges), charts, SmartArt, KPIs and images on-brand; the open work is
   richer placeholder/overflow handling, multi-master selection, and closing the
   remaining fidelity gap to the docx reference vertical.
2. **Richer visual analysis**: build on the PyMuPDF fallback with optional
   `numpy`/`opencv-python` or `scikit-image` for overlap, clipping, and
   large-empty-region detection.
3. **Broader skill evals**: expand the current template-based eval set with more
   corporate templates, visual-repair traces, and with/without-skill comparisons.

---

## The Brand Kit

Each extracted template produces a self-contained, copyable directory:

```text
brand-kit/<name>/
├─ profile.json          # the brand rules: theme, roles, structure, anchors, catalog
├─ PROFILE.md            # human-readable summary (role map + structure)
├─ template/shell.docx   # the original template, kept byte-for-byte (the shell)
└─ provenance.sha256     # source hash for drift detection
```

`brand-kit/` lives either in your **project** (`./brand-kit/`, versionable, wins)
or **globally** (`~/.claude/brand-kit/`, reusable across projects). It is the
template's portable memory: copy the folder and the brand travels with it.
