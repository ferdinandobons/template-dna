# Changelog

All notable changes to BrandDocs are documented in this file.

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

[0.1.0]: https://github.com/ferdinandobons/brand-docs/releases/tag/v0.1.0
