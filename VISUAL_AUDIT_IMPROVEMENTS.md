# Visual Audit Improvements

## Current baseline

The current visual audit stack is a good V1 foundation:

- `LibreOffice` / `soffice` renders DOCX, PPTX, and XLSX to PDF using a real
  office layout engine.
- `Poppler` / `pdftoppm` rasterizes the generated PDF into per-page PNG files.
- `Pillow` runs deterministic pixel checks such as blank-page and edge-bleed
  detection.
- `doctor` probes Python packages, external binaries, and the real
  DOCX/PPTX/XLSX-to-PDF-to-PNG pipelines before a visual audit is trusted.

This is the right default architecture because pure OOXML inspection cannot prove
rendered layout correctness. The engine needs a real renderer, then image-level
checks, then an orchestrator-level qualitative review.

## Dependency preflight contract

Every brand skill (`brand-docx`, `brand-pptx`, `brand-xlsx`) should run this
before extract, verify, or generate:

```bash
python scripts/brandkit/cli.py doctor
```

Expected behavior:

- If required Python packages are missing, stop and install/repair them before
  running the core engine.
- If only visual renderers are missing or unusable, extraction and deterministic
  L0 QA can still run, but the skill must not claim a full visual audit.
- If `--qa deep` is requested and renderers are unavailable, the engine should
  write a degraded manifest and clearly report which visual proof is incomplete.
- `doctor` should always print actionable `install:` or `repair:` hints for
  missing/unusable dependencies.

## Implemented improvements

- Visual manifests now include environment diagnostics: platform, Python version,
  renderer availability, renderer paths, degraded status, and install/repair
  hints when available.
- The visual preflight now smoke-tests the full render chain for DOCX, PPTX, and
  XLSX instead of treating DOCX success as proof for every Office format.
- PyMuPDF (`fitz`) is now detected as an optional PDF raster fallback when
  `pdftoppm` is unavailable or fails after LibreOffice has produced a PDF.
- Gated real-render E2E tests now cover `--qa deep` manifest/PNG generation for
  DOCX, PPTX, and XLSX when `BRANDDOCS_RUN_REAL_RENDER=1` is enabled.
- DOCX generation now rewrites the visible cache of outline TOC fields from the
  generated headings before marking the field dirty. This prevents headless
  LibreOffice renders from showing stale template TOC entries while keeping the
  TOC field updateable in Word/LibreOffice.
- The DOCX generator already prunes leading empty body artifacts exposed after
  caption-index reconciliation, so inherited section breaks that cause blank
  pre-body pages are not preserved blindly.

## Recommended improvements

### 1. Add composition intelligence, not literal preservation

The template should be treated as a source of reusable affordances, constraints,
and brand conventions, not as a byte-for-byte layout script that must be followed
even when the rendered result is broken.

Important example:

- If a source template contains section breaks that are useful for its original
  long document but cause two blank pages in the generated result, the skill
  should not preserve those section breaks blindly.
- It should record them as available structural elements, understand what they do
  (page geometry, header/footer transition, orientation, spacing), then decide
  whether keeping, moving, collapsing, or removing them best serves the generated
  document.
- Visual findings such as `visual.blank_page`, stale TOC pages, or large empty
  regions after front matter should trigger targeted composition repair, not just
  a warning in the final report.

Target behavior:

1. Render and audit the generated artifact.
2. Classify each visual defect by likely cause: inherited section break, stale
   field cache, oversized block, unsupported component, bad pagination, residual
   placeholder, or style mismatch.
3. Apply the smallest justified repair to the generated artifact or
   IntermediateDocument.
4. Regenerate and rerun the visual audit.
5. Repeat until the rendered output is clean, or until no further targeted repair
   can be justified without user input.

The skill should preserve template structures only while they help the final
document. A visually broken result is stronger evidence than a structural
preservation preference.

### 2. Add a second PDF rasterizer

Add `PyMuPDF` as a fallback or cross-check after PDF generation.

Status: fallback is implemented; renderer-disagreement cross-checking remains a
future enhancement.

Why:

- It can render PDF pages directly from Python.
- It reduces dependence on a single `pdftoppm` rasterization path.
- It enables side-by-side renderer disagreement checks.

Suggested use:

- Keep `pdftoppm` as default.
- Use `PyMuPDF` when `pdftoppm` is missing or when L1 results look suspicious.
- Optionally compare page count, dimensions, and coarse ink maps between both
  renderers.

### 3. Add stronger image analysis

Add `numpy` plus either `opencv-python` or `scikit-image`.

Useful checks:

- text/image bounding boxes near page edges;
- connected-component analysis for tiny clipped fragments;
- large empty regions after expected content starts;
- overlap heuristics based on dense connected regions;
- diff heatmaps between template render and generated render.

This would upgrade L1 from simple pixel proxies to richer deterministic layout
proxies.

### 4. Add optional OCR

Add OCR only as an optional capability, not a hard dependency.

Candidate stack:

- `tesseract` binary;
- `pytesseract` Python wrapper.

Useful cases:

- visible template placeholders rendered inside text boxes or shapes;
- residual demo text that OOXML extraction misses;
- table-of-contents caches that render old text even when fields are marked
  dirty.

OCR should be used as an advisory signal because it can be noisy across fonts,
languages, and image quality.

### 5. Deepen TOC/cache handling

The current DOCX flow marks TOC fields dirty and rewrites the visible outline
TOC cache from generated headings, but deeper field-cache cases still need work.

Potential fixes:

- handle nested/multi-column TOCs and more complex field-result structures;
- generate page-number-aware static TOC entries when the template allows it;
- add an L1/L2 checklist item for stale TOC/demo entries;
- use OCR or text extraction from rendered PDF to detect stale visible TOC text.

This matters because the visual audit can be technically successful while the
render still shows stale cached entries.

### 6. Store environment diagnostics in the manifest

Extend `visual_manifest.json` with:

- renderer binary paths;
- renderer versions;
- OS/platform;
- DPI;
- fallback mode;
- `doctor` status summary;
- install/repair hints when degraded.

This makes audit failures easier to reproduce and compare across machines.

### 7. Add renderer policy per QA mode

Suggested policy:

- `fast`: L0 only, never touches renderers.
- `auto`: L0 plus visual render when the full pipeline is available.
- `deep`: preflight first; prefer full render; if unavailable, produce degraded
  manifest and say exactly what is unproven.
- future `strict`: fail if full render is unavailable or if L1/L2 checks are not
  clean.

### 8. Keep Playwright out of the Office core path

`Playwright` is useful for HTML/web preview audits, but it should not become the
primary Office renderer. Browser screenshots vary by OS/browser environment and
do not model Word/PowerPoint/Excel layout. It is better as a future companion for
HTML exports or dashboards, not as the main Office audit engine.

## Priority order

1. Keep the current LibreOffice + Poppler + Pillow path stable.
2. Make preflight mandatory in every skill workflow.
3. Add composition intelligence for inherited structures such as section breaks.
4. Add richer manifest diagnostics for degraded visual audits.
5. Deepen stale TOC/cache rendering beyond outline DOCX fields.
6. Add renderer disagreement checks on top of the PyMuPDF fallback.
7. Add richer image analysis with `numpy` and `opencv-python` or `scikit-image`.
8. Add optional OCR for residual visible text.
