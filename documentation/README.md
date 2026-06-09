# BrandDocs documentation

Long-form documentation for [BrandDocs](../README.md). The top-level
[`README.md`](../README.md) stays intentionally short; everything else lives here.

> Not to be confused with [`docs/`](../docs), which is the published
> [GitHub Pages site](https://ferdinandobons.github.io/brand-docs/) (HTML, sitemap,
> robots, `llms.txt`), not human documentation.

| Document | What's inside |
|---|---|
| [INSTALLATION.md](INSTALLATION.md) | Prerequisites (required + optional visual QA) and every install method: Claude Code plugin, Git submodule, Codex clone, updating. |
| [USAGE.md](USAGE.md) | Common use cases and the structured `IntermediateDocument` / `GridDocument` input format. (The quick start lives in the [README](../README.md#quick-start).) |
| [ARCHITECTURE.md](ARCHITECTURE.md) | How it works (extract → generate → verify), structure-awareness, the reliability/repair loop, the Brand Kit, and why this beats "just ask an AI". |
| [SKILLS.md](SKILLS.md) | The three skills (`brand-docx` / `brand-pptx` / `brand-xlsx`) and the project status table. |
| [PLUGIN_WORKFLOW.md](PLUGIN_WORKFLOW.md) | The full end-to-end agent workflow: skill selection, preflight, extract/comprehend/generate/QA, visual manifests, repair rounds. |
| [DEVELOPMENT.md](DEVELOPMENT.md) | Local dev setup, test suites, and contribution conventions. |
| [FAQ.md](FAQ.md) | Frequently asked questions and discovery keywords. |
| [ROADMAP.md](ROADMAP.md) | Planned, designed-but-not-yet-built features (typography capture, profile learning from QA findings) and salvaged visual-audit ideas. |
| [DIRECTORY_SUBMISSIONS.md](DIRECTORY_SUBMISSIONS.md) | Reusable copy, categories and tags for listing BrandDocs in directories. |
