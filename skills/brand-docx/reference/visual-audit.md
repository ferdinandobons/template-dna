<!-- SPDX-License-Identifier: MIT -->
# Visual audit (two-stage)

The visual audit sits **on top of** the L0 deterministic gate (schema, resolver
targets, residual text, structural diffs). L0 stays unchanged and authoritative.
The audit adds two stages that see what L0 cannot, the *rendered layout*:

- **L1: deterministic pixel proxies (engine).** The engine renders the output to
  per-page PNGs and runs pixel checks that flag rendered-layout defects. During
  `--qa deep`, optional Tesseract OCR also scans the rendered pages for captured
  template/demo text. Each defect is a `Finding(check="visual.<name>")`, severity
  **WARNING** (never ERROR, so the audit never fails a gate that L0 passes).
- **L2: qualitative judgement (orchestrator, i.e. you).** The engine emits a
  structured `visual_manifest.json` (PNG paths + a profile-derived checklist + the
  L1 findings). **You** open the PNGs, judge each checklist item PASS/FAIL, and
  drive a repair loop. The engine **never** calls a model.

The PNGs and the manifest are **side artifacts** written to an
`<output-file>.visual/` dir next to the output, such as `report.docx.visual/`.
The bytes of the generated document never change because of the audit.

## When it runs (`--qa`)

| `--qa` | renderers present (`soffice` + `pdftoppm`, or optional PyMuPDF/`fitz` fallback) | renderers absent (e.g. CI) |
|---|---|---|
| `fast` | L0 only | L0 only (identical) |
| `auto` | L0 + L1 | L0 + one INFO `visual.unavailable` |
| `deep` | L0 + L1 + **manifest** (INFO `visual.manifest` with the path) -> triggers your L2 step | L0 + INFO `visual.unavailable` + a **degraded manifest** with the checklist; on macOS a first-page Quick Look fallback may be included |
| `strict` | L0 + L1 + OCR/manifest + ERROR `visual.strict` if render/L1/OCR evidence is unclean | L0 + degraded manifest + ERROR `visual.strict_unavailable` |

Notes:
- At **verify** time there is no output to render, so every mode behaves as L0.
- Renderer absence degrades cleanly: INFO/WARNING findings only, no ERROR, **exit
  code unchanged**. `auto` may create no `.visual` dir; `deep` writes a degraded
  `.visual/visual_manifest.json` so the orchestrator can still inspect what was
  skipped and which checklist items remain unproven.
- `strict` is the release-quality mode: it writes the manifest like `deep`, but
  fails when full render proof is unavailable or when L1/OCR findings need repair.
- The renderer is env-detected via `doctor.probe()["visual_qa"]`, which
  smoke-tests the DOCX, PPTX, and XLSX render paths end to end; run
  `python scripts/brandkit/cli.py doctor` before starting the workflow and report
  missing/unusable dependencies before claiming a full visual audit.
- `pdftoppm` remains the baseline PDF rasterizer. If it is missing or fails after
  LibreOffice produced a PDF, optional PyMuPDF (`fitz`) can rasterize the PDF as
  a degraded fallback; the manifest records this in `environment.optional_python`.
- Optional OCR uses the external `tesseract` binary. Missing OCR does not block
  the visual audit; it only means rendered residual-text proof is incomplete. The
  manifest records OCR availability in `environment.ocr` and results in `ocr`.

## L1 proxies

| check id | catches (that L0 cannot see) | severity |
|---|---|---|
| `visual.blank_page` | a blank/near-blank page: broken page, content not rendered, or overflow that pushed everything off the page | WARNING |
| `visual.edge_bleed` | content touching/exceeding the printable margins (text/image clipping or overflow); the defect `OverflowCapability.RENDER` exists to intercept | WARNING |
| `visual.no_pages` | an existing output that rendered zero pages | WARNING |
| `visual.ocr_residual_text` | optional OCR saw captured template/demo text still visible in the render | WARNING |

The proxies are deterministic (aggregate luma/ink fractions, no randomness, no
single-pixel font-hinting sensitivity) and conservative (tuned to render data:
content pages mean-luma ~240–252, a blank page ~255) to minimize false positives.
They accept either a PIL image or a PNG path and never raise on an unreadable
image (they return no findings instead).

A WARNING is a *signal for your L2 judgement*, not a standalone failure: a
legitimately near-empty page (a minimal cover/separator) or a deliberate
full-bleed background exists, and only you can tell those apart from a real
defect.

## The manifest (`visual_manifest.json`)

Top-level fields:

```json
{
  "schema_version": "visual-manifest-1",
  "kind": "docx",
  "profile_name": "<identity.name>",
  "document": "out.docx",
  "renderers_available": true,
  "qa_mode": "deep",
  "dpi": 100,
  "pages": [
    {"index": 1, "png": "page-1.png", "width": 850, "height": 1100, "orientation": "portrait"}
  ],
  "l1_findings": [
    {"check": "visual.blank_page", "severity": "WARNING", "message": "...", "location": "page:2"}
  ],
  "ocr": {
    "engine": "tesseract",
    "available": true,
    "status": "ok",
    "terms_checked": ["<captured template/demo text>"],
    "pages": [{"index": 1, "text": "...", "text_truncated": false}],
    "hits": [],
    "errors": []
  },
  "environment": {
    "platform": "macOS-...",
    "python": "3.x",
    "visual_qa": true,
    "degraded": false,
    "renderers": {
      "soffice": {"available": true, "path": "/path/to/soffice"},
      "pdftoppm": {"available": true, "path": "/path/to/pdftoppm"}
    },
    "optional_python": {
      "fitz": {"available": false, "purpose": "PyMuPDF PDF raster fallback"}
    },
    "ocr": {
      "tesseract": {"available": true, "path": "/path/to/tesseract", "purpose": "optional OCR visible-text audit"}
    },
    "install_hints": []
  },
  "checklist": [
    {"id": "regions_present", "what": "...", "derived_from": "structure.skeleton[*].region + order", "severity_hint": "WARNING"}
  ],
  "instructions": "Open each PNG. For each checklist item, judge PASS/FAIL ..."
}
```

`pages` may be populated even when `renderers_available` is `false`: this means a
degraded fallback image (for example a macOS Quick Look thumbnail) is available
for inspection, not that full page-by-page render proof succeeded. Treat those
PNGs as advisory evidence only and do not claim a clean visual audit from them.

`png` paths are **relative to the `.visual` dir**. When renderers are absent the
manifest carries `"degraded": true`; `pages` is populated only if an advisory
fallback image exists, `l1_findings` still records visual failure/degradation
signals, and the `checklist` stays populated so you know what would have been
inspected with full render proof.

### Checklist items (derived from the profile, model-free)

Each item is included only when its backing data exists in the profile, so the
checklist is tailored to the template:

| item id | inspect | derived from |
|---|---|---|
| `regions_present` | expected regions appear in order | `structure.skeleton[*].region + order` |
| `cover_correct` | bound title, no duplicate, no demo prompt | `anchors.cover` + `comprehension.cover_slots` |
| `no_residual_placeholder` | no template demo text visible | `surface.<kind>.demo_region` + cover demo values + optional `ocr.hits` |
| `palette_on_brand` | on-screen colors are brand colors | `theme.colors` + `theme.palette_roles` |
| `roles_styled` | semantic blocks carry brand styles, not "Normal" | `roles._index` |
| `no_overlap` | no overlapping/clipped text or shapes | constant; reinforced by `visual.edge_bleed` |
| `no_blank_pages` | no unexpected blank/broken pages | constant; reinforced by `visual.blank_page` |
| `charts_rendered` | charts drawn correctly, not empty boxes | present only when a chart role/component exists |
| `overflow_clean` | nothing past the printable margins | `qa.overflow_capability` ∈ {render, estimator, cellfit} |

## L2 repair loop (your protocol)

1. Generate with `--qa deep`.
2. Read the manifest path from stdout (`visual manifest: <path>`).
3. Open every `pages[*].png`. For each `checklist` item, judge PASS/FAIL against
   the rendered pages, taking `l1_findings` and `ocr.hits` into account (an L1
   WARNING is a pointer to a page/side worth looking at, not a verdict).
4. If any item FAILS (or an L1/OCR WARNING is a real defect on inspection):
   repair the IntermediateDocument/content, regenerate, and re-run the audit.
5. Repeat until the checklist is clean (**max 3 iterations** by default), then
   return the file with an honest QA summary.
6. For release-quality validation, rerun with `--qa strict`; any `visual.strict`
   ERROR identifies the exact visual finding that still needs targeted repair.

The engine produces the evidence; the qualitative judgement and the decision to
regenerate are yours.
