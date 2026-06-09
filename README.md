<div align="center">

<img src="assets/hero.png" alt="BrandDocs workflow: a real Word, PowerPoint or Excel template is extracted into a reusable Brand Profile, then composed and verified into on-brand DOCX, PPTX and XLSX output" width="100%" />

<br/>

# BrandDocs: AI On-Brand Document Generator for Word, PowerPoint & Excel

**BrandDocs turns existing Word, PowerPoint and Excel templates into reusable AI document-generation skills.** Unlike generic AI document generators, it preserves **brand, structure, styles and formulas by construction**. Built for Claude Code, Codex and compatible AI agents.

[![License: MIT](https://img.shields.io/badge/License-MIT-3B82F6.svg)](LICENSE)
[![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-3776AB.svg)](https://www.python.org/)
[![CI](https://github.com/ferdinandobons/brand-docs/actions/workflows/ci.yml/badge.svg)](https://github.com/ferdinandobons/brand-docs/actions/workflows/ci.yml)
[![Website](https://img.shields.io/badge/website-GitHub%20Pages-16A34A.svg)](https://ferdinandobons.github.io/brand-docs/)
[![Latest release](https://img.shields.io/github/v/release/ferdinandobons/brand-docs?label=latest%20release)](https://github.com/ferdinandobons/brand-docs/releases/latest)
[![Skills](https://img.shields.io/badge/skills-docx%20·%20pptx%20·%20xlsx-6EA8FE.svg)](#the-three-skills)
[![Status: alpha](https://img.shields.io/badge/status-alpha-F59E0B.svg)](#project-status)

</div>

---

## What is BrandDocs?

**BrandDocs** is an open-source **agent-skill bundle** that turns a company's existing Office templates into reusable AI document-generation skills. Point it at one branded `.docx`, `.pptx`, or `.xlsx`; it **extracts** the brand (theme colors and fonts, named styles, the document's *structure*, layouts, cover anchors, logos and tables) into a portable **Brand Profile**. From then on, every document it **generates** is built *from the original template shell* and uses *only* the artifacts the template actually defines. Each format stays in its own lane: there is no cross-format conversion.

> **The core guarantee: off-brand output is impossible by construction.** No generator ever writes a literal style name, hex color, or font: those live only in the Brand Profile, and `verify` refuses a profile that points at anything the template doesn't contain.

### At a glance

| Question | Answer |
|---|---|
| **Input** | Existing company `.docx`, `.pptx`, or `.xlsx` templates |
| **Output** | Same-format on-brand Word documents, PowerPoint decks, and Excel workbooks |
| **Works with** | Claude Code, Codex, compatible AI agents, or the direct Python CLI |
| **Best for** | Repeatable reports, decks, workbooks, proposals, memos, briefs, and internal document workflows |
| **Privacy model** | Local-first; no cloud service is required, and real templates are git-ignored |
| **Current release** | [v0.7.0](https://github.com/ferdinandobons/brand-docs/releases/tag/v0.7.0) alpha |

---

## The three skills

| Skill | Format | Generates |
|---|---|---|
| **`brand-docx`** | Word `.docx` | reports, letters, memos, in the template's structural order |
| **`brand-pptx`** | PowerPoint `.pptx` | decks from the template's real masters & layouts, with native charts, diagrams & merged tables |
| **`brand-xlsx`** | Excel `.xlsx` | workbooks: named-region fills with **formulas preserved** and brand number formats |

All three share one engine and expose the same verbs: **`extract` → `verify` → `generate`**. Details → [documentation/SKILLS.md](documentation/SKILLS.md).

---

## Prerequisites

BrandDocs runs locally and needs a few things installed **before first use**:

- **Python ≥ 3.10** plus the packages in [`requirements.txt`](requirements.txt) (`python-docx`, `python-pptx`, `openpyxl`, `lxml`, `Pillow`):
  ```bash
  pip install -r requirements.txt
  ```
- **Visual QA tools (keep the visual gate on):** the render-based **visual QA gate runs by default** and catches layout problems the deterministic checks can't (text overflow, blank pages, clipping, stale demo text). It needs LibreOffice + Poppler (Tesseract is optional, for OCR). Install them with one auto-detecting command:
  ```bash
  bash scripts/setup_visual_qa.sh
  ```
  Generation still runs without them (it degrades gracefully to L0), but install them so the visual gate stays on.

Check what's present at any time with `python scripts/brandkit/cli.py doctor`. Per-OS commands and the full setup → **[documentation/INSTALLATION.md](documentation/INSTALLATION.md)**.

---

## Installation

The three skills share one Python engine (`scripts/brandkit/`), so install the **whole repository** (copying a single skill folder on its own won't work). After either install below, set up the [prerequisites](#prerequisites) so the engine can run, then verify with `python scripts/brandkit/cli.py doctor`.

### Claude Code

The plugin loads all three skills plus the shared engine together:

```text
/plugin marketplace add ferdinandobons/brand-docs
/plugin install brand-docs@brand-docs
```

### Codex (and other agents)

Clone the repo and symlink the three skills into your Codex skills directory, so each skill's engine in `scripts/brandkit/` travels with it:

```bash
git clone https://github.com/ferdinandobons/brand-docs.git ~/.codex/brand-docs
cd ~/.codex/brand-docs && python3 -m venv .venv && . .venv/bin/activate && pip install -r requirements.txt
mkdir -p ~/.codex/skills
for s in brand-docx brand-pptx brand-xlsx; do ln -s ~/.codex/brand-docs/skills/$s ~/.codex/skills/$s; done
```

Restart or reload the agent if the skills don't appear immediately.

> Git-submodule install and updating instructions are in **[documentation/INSTALLATION.md](documentation/INSTALLATION.md)**.

---

## Quick start

**With an AI agent** (the intended experience). Describe what you want and attach a template:

> "Use this company Word template and write a report on the history of Napoleon."

The agent activates `brand-docx`, extracts (or reuses) a Brand Profile, fills the template shell in its structural order, runs QA, and returns the file. PowerPoint (`brand-pptx`) and Excel (`brand-xlsx`) work the same way.

**Direct CLI** (the engine, for tests & debugging):

```bash
# 1) Extract the brand from a template into a reusable kit
python scripts/brandkit/cli.py extract --name <your_company> --template template.docx --scope project

# 2) Verify the profile (fails if a role points at a missing artifact)
python scripts/brandkit/cli.py verify --name <your_company> --scope auto --qa auto

# 3) Generate a new on-brand document from structured content
python scripts/brandkit/cli.py generate --name <your_company> --input idoc.json --output out.docx --scope auto --qa auto
```

The input (`idoc.json`) is an **IntermediateDocument** of brand-agnostic typed blocks (no styles, colors or fonts); the Brand Profile resolves all of that.

---

## Project status

**Alpha.** The Word vertical (`brand-docx`) is the reference implementation, verified end-to-end on real templates; PowerPoint and Excel share the engine and are catching up. Full status table → [documentation/SKILLS.md](documentation/SKILLS.md#project-status).

## Website & discovery

- Public website: [ferdinandobons.github.io/brand-docs](https://ferdinandobons.github.io/brand-docs/)
- Full documentation: [`documentation/`](documentation/)

## Changelog

Latest release: [v0.7.0](https://github.com/ferdinandobons/brand-docs/releases/tag/v0.7.0). See [CHANGELOG.md](CHANGELOG.md).

## License & acknowledgements

- This project's own code is **[MIT](LICENSE)** © 2026 Ferdinando Bonsegna.
- Self-contained: the OOXML engine is re-implemented from scratch; it does **not** vendor any proprietary or third-party Office tooling. See [`NOTICE`](NOTICE).
