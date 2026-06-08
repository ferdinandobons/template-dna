# The three skills & project status

## The three skills

| Skill | Format | Generates |
|---|---|---|
| **`brand-docx`** | Word `.docx` | reports, letters, memos: cover, headings, paragraphs, callouts, quotes, captions, lists, tables, in the template's structural order |
| **`brand-pptx`** | PowerPoint `.pptx` | decks: title / section / content slides from the template's real masters & layouts, with real bullet levels and long-text splitting |
| **`brand-xlsx`** | Excel `.xlsx` | workbooks: fills named cells & regions while **preserving formulas** and workbook structure |

All three expose the same three verbs: **`extract` → `verify` → `generate`**. Each skill is self-contained and **same-format** (a Word template makes Word
documents, never a deck or a sheet). They share one engine: a single profile
schema, resolver, OOXML layer and QA gate underpin all three formats.

An optional fourth verb, **`comprehend`**, lets the model read a bounded summary
of the template and record what each structure is **for** before generating: cover
slots, demo-vs-real regions, derived-index conventions, and recurring layouts it
proposes as **reusable component/section fragments**. Every proposal is validated
fail-closed and frozen into the Brand Profile; generation works with no
comprehension at all (the deterministic path).

---

## Project status

**Alpha.** The Word vertical (`brand-docx`) is the reference implementation,
verified end-to-end on real templates; PowerPoint and Excel share the engine and
are catching up.

| Area | Status |
|---|---|
| Shared engine (profile schema, resolver, OOXML, CLI, dual store) | ✅ working |
| `brand-docx`: extract → verify → generate | ✅ working |
| Document **structure** extraction & order-aware generation | ✅ working |
| Brand-guarantee enforcement (`verify` fails on missing artifacts) | ✅ working |
| Deterministic QA (L0: styles, palette, residual text, tables, formula preservation, language) | ✅ working |
| `brand-pptx`: roles from real layouts, native charts / SmartArt / merged tables | ✅ working (fidelity still catching up to docx) |
| `brand-xlsx`: named-region fills, formula-preserving, native charts | ✅ working (fidelity still catching up to docx) |
| Visual QA (LibreOffice render + manifest-driven repair loop) | 🚧 implemented with graceful degraded mode |
| Native charts (DOCX / PPTX / XLSX), SmartArt diagrams (DOCX / PPTX), merged tables | ✅ working |
| Native Word `toc` (authored field, or deferral to a preserved outline TOC) | ✅ working |
| Excel semantic number formats (`number.<family>` resolved to the template's mask) | ✅ working |
| Model-driven reusable-fragment population (`comprehend` → `components` / `sections`, with `{{slot}}` substitution) | ✅ working |
| PyMuPDF PDF raster fallback | ✅ working |
| Optional OCR rendered-text residual scan | ✅ working when Tesseract is installed |
| Template-based skill eval set (DOCX/PPTX/XLSX) | ✅ working in CI |
| Strict visual mode (`--qa strict`) | ✅ working |
| Richer image analysis | 🔭 planned |

Visual Word overflow needs LibreOffice, since Word lays out at render time.
