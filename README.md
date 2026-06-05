<div align="center">

<img src="assets/hero.svg" alt="TemplateDNA — learn a company's template once, generate on-brand Word, PowerPoint and Excel forever" width="100%" />

<br/>

**Learn a company's Office template once. Generate on-brand Word, PowerPoint & Excel documents forever.**

[![License: MIT](https://img.shields.io/badge/License-MIT-3B82F6.svg)](LICENSE)
[![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-3776AB.svg)](https://www.python.org/)
[![Skills](https://img.shields.io/badge/skills-docx%20·%20pptx%20·%20xlsx-6EA8FE.svg)](#the-three-skills)
[![Status: alpha](https://img.shields.io/badge/status-alpha-F59E0B.svg)](#project-status)

</div>

---

## What is TemplateDNA?

`TemplateDNA` is an **agent-skill bundle** (for Claude Code, Codex, and compatible AI agents) that turns a company's existing Office template into a **reusable brand memory**, then writes new documents that stay faithful to it.

You give it one branded `.docx`, `.pptx`, or `.xlsx`. It **extracts** the brand — colors, fonts, named styles, layouts, cover anchors, the document's *structure*, logos, tables — into a portable **Brand Profile**. From then on, every document it **generates** is built *from the original template shell* and uses *only* the artifacts the template actually defines.

> **The core guarantee: off-brand output is impossible by construction.** Generation never writes a literal style name, hex color, or font — those live only in the Brand Profile. There is no "creative" path that drifts from the brand.

### Why not just ask an AI to "use this template"?

General-purpose document skills generate *freely* and only loosely imitate a reference file — fonts drift, the palette wanders, the corporate structure is lost. `TemplateDNA` is the opposite: narrow and faithful. It learns the template as a set of **rules and reusable parts**, remembers them in a `brand-kit/`, and **respects them** across an unlimited number of documents.

|  | General-purpose Office skills | **TemplateDNA** |
|---|---|---|
| Mental model | "create a nice document" | "fill the company's template" |
| Brand fidelity | best-effort imitation | **by construction** (opens from the shell, uses only its styles) |
| Reusability | re-explain the brand every time | **extract once**, reuse forever via `brand-kit/` |
| Structure | free-form | **respects the template's cover → contents → body order** |

---

## Highlights

- 🎯 **Brand-faithful by construction** — opens from the real template shell; applies only its named styles, theme colors, fonts and layouts.
- 🧠 **Extract once, reuse forever** — a portable `brand-kit/<name>/` is the template's memory; every later document reads it.
- 🏛️ **Structure-aware** — extracts the template's *ordered skeleton* (e.g. **cover → table of contents → body**) and knows which parts are a **fixed structure to respect** versus **styles to use on demand**.
- 🧩 **Three formats, one engine** — Word, PowerPoint and Excel share a single resolver, profile schema and QA gate, so they behave consistently.
- 🗂️ **Full artifact catalog** — records OOXML parts, styles, media, layouts, formulas and named ranges so an agent can reason about anything the template exposes.
- ✅ **Built-in QA** — deterministic checks (allowed styles, palette adherence, residual template text, broken tables) with an optional visual pass.
- 🔓 **Self-contained & MIT** — pure `python-docx` / `python-pptx` / `openpyxl` + OOXML. No cloud, no vendor lock-in.

---

## The three skills

| Skill | Format | Generates |
|---|---|---|
| **`brand-docx`** | Word `.docx` | reports, letters, memos: cover, headings, paragraphs, callouts, quotes, captions, lists, tables — in the template's structural order |
| **`brand-pptx`** | PowerPoint `.pptx` | decks: title/section/content slides from the template's masters & layouts, with long-text splitting |
| **`brand-xlsx`** | Excel `.xlsx` | workbooks: fills named cells & regions while **preserving formulas** and workbook structure |

All three expose the same three verbs: **`extract` → `verify` → `generate`**.

---

## Prerequisites & installation

### Required (core extract / generate / deterministic QA)

- **Python ≥ 3.10**
- Python packages (installed via `requirements.txt`):
  `python-docx>=1.1`, `python-pptx>=1.0`, `openpyxl>=3.1`, `lxml>=5.0`, `Pillow>=10.0`

```bash
git clone https://github.com/ferdinandobons/template-dna.git
cd template-dna
python3 -m venv .venv && . .venv/bin/activate     # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

### Optional (visual QA — render-based checks)

Needed only for the **visual** verification pass; their absence degrades gracefully and never blocks extraction, generation, or deterministic QA.

- **LibreOffice** (`soffice`) — headless render to PDF
- **Poppler** (`pdftoppm`) — PDF → PNG

```bash
# macOS (Homebrew)
brew install --cask libreoffice && brew install poppler
# Debian / Ubuntu
sudo apt-get install -y libreoffice poppler-utils
# Fedora
sudo dnf install -y libreoffice poppler-utils
# Windows: winget install TheDocumentFoundation.LibreOffice
#          + Poppler via conda-forge or a prebuilt binary on PATH
```

Check what's available at any time:

```bash
python scripts/brandkit/cli.py doctor
```

`doctor` reports your capability level and prints the exact install command for anything missing — it never fails the run.

### Install as an agent skill

**Claude Code** (loads all three skills + the shared engine together):

```text
/plugin marketplace add ferdinandobons/template-dna
/plugin install template-dna
```

**Codex** (clone + symlink the skills):

```bash
git clone https://github.com/ferdinandobons/template-dna.git ~/.codex/template-dna
cd ~/.codex/template-dna && python3 -m venv .venv && . .venv/bin/activate && pip install -r requirements.txt
mkdir -p ~/.codex/skills
for s in brand-docx brand-pptx brand-xlsx; do ln -s ~/.codex/template-dna/skills/$s ~/.codex/skills/$s; done
```

> Restart/reload the agent after installing if the skills don't appear immediately.

---

## Quick start

### With an AI agent (the intended experience)

Just describe what you want and attach a template:

> "Use this company Word template and write a report on the history of Napoleon."

The agent activates `brand-docx`, extracts a Brand Profile from the template (or reuses an existing one), converts your request into the template's structure, generates the `.docx` from the original shell, runs QA, and returns the file.

### Direct CLI (engine — for tests & debugging)

```bash
# 1) Extract the brand from a template into a reusable kit
python scripts/brandkit/cli.py extract --name acme --template template.docx --scope project

# 2) Verify the profile (role mapping + QA)
python scripts/brandkit/cli.py verify --name acme --scope auto --qa auto

# 3) Generate a new on-brand document from structured content
python scripts/brandkit/cli.py generate --name acme --input idoc.json --output out.docx --scope auto --qa auto
```

The same three commands work for `.pptx` (with an `IntermediateDocument`) and `.xlsx` (with a `GridDocument`).

---

## How it works

```
 company template ──▶ EXTRACT ──▶ brand-kit/<name>/ ──▶ GENERATE ──▶ on-brand document
   .docx/.pptx/.xlsx     │         (profile + shell)        │
                         │                                  ├─ opens FROM the template shell
                         ├─ theme colors & fonts            ├─ resolves semantic blocks → brand styles
                         ├─ named styles & roles            ├─ respects the template's structure order
                         ├─ document structure (skeleton)   └─ runs the QA gate
                         ├─ layouts / cover anchors
                         ├─ logos, media, tables, formulas
                         └─ full artifact catalog
```

1. **Extract** unpacks the template's OOXML and records its brand: theme, named styles (mapped to semantic **roles**), the **document structure** (the ordered skeleton + which parts are fixed vs free), layouts, cover anchors, logos, and a complete artifact catalog. The original file is kept **byte-for-byte** as the *shell*.
2. **Generate** turns your content into an **IntermediateDocument** of brand-agnostic typed blocks (heading, paragraph, callout, table, …). A **pure resolver** maps each block to the concrete brand artifact from the profile, fills the shell **in the template's structural order**, and saves.
3. **Verify / QA** runs deterministic checks (only allowed styles, palette adherence, no residual template text, table integrity) and, when LibreOffice is available, a visual pass.

### The Brand Kit

Each extracted template produces a self-contained, copyable directory:

```text
brand-kit/<name>/
├─ profile.json          # the brand rules: theme, roles, structure, anchors, catalog
├─ PROFILE.md            # human-readable summary (role map + structure)
├─ template/shell.docx   # the original template, kept byte-for-byte (the shell)
└─ provenance.sha256     # source hash for drift detection
```

`brand-kit/` lives either in your **project** (`./brand-kit/`, versionable, wins) or **globally** (`~/.claude/brand-kit/`, reusable across projects). It is the template's portable memory.

---

## Project status

**Alpha.** The Word vertical (`brand-docx`) is the reference implementation; PowerPoint and Excel share the engine and are progressing behind it.

| Area | Status |
|---|---|
| Shared engine (profile schema, resolver, OOXML, CLI) | ✅ working |
| `brand-docx` extract → generate | ✅ working |
| Document **structure** extraction & order-aware generation | 🚧 in progress |
| `brand-pptx` / `brand-xlsx` | 🚧 early |
| Adaptive QA gate (visual + auto-repair) | 🚧 deterministic L0 first |
| Charts / SmartArt / section templates | 🔭 catalogued, regeneration staged |

Visual Word overflow needs LibreOffice, since Word lays out at render time.

---

## Development

```bash
PYTHONPATH=scripts python -m unittest tests.test_smoke tests.test_no_proprietary
```

> **Never commit real templates or company assets.** `brand-kit/` and `generated/` are intentionally git-ignored, and `tests/test_no_proprietary.py` fails the build if a proprietary asset leaks in. See [`CONTRIBUTING.md`](CONTRIBUTING.md).

---

## License & acknowledgements

- This project's own code is **[MIT](LICENSE)** © 2026 Ferdinando Bonsegna.
- Self-contained: the OOXML engine is re-implemented from scratch; it does **not** vendor any proprietary or third-party Office tooling. See [`NOTICE`](NOTICE).
