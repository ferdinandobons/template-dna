<div align="center">

<img src="assets/hero.svg" alt="OfficeSkills - extract a company's Word, PowerPoint or Excel template into a reusable Brand Profile, then generate unlimited on-brand documents of that same format" width="100%" />

<br/>

**Turn a company's Word, PowerPoint or Excel template into unlimited on-brand documents of the same type.**

[![License: MIT](https://img.shields.io/badge/License-MIT-3B82F6.svg)](LICENSE)
[![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-3776AB.svg)](https://www.python.org/)
[![CI](https://github.com/ferdinandobons/office-skills/actions/workflows/ci.yml/badge.svg)](https://github.com/ferdinandobons/office-skills/actions/workflows/ci.yml)
[![Skills](https://img.shields.io/badge/skills-docx%20·%20pptx%20·%20xlsx-6EA8FE.svg)](#the-three-skills)
[![Status: alpha](https://img.shields.io/badge/status-alpha-F59E0B.svg)](#project-status)

</div>

---

## Contents

- [What is OfficeSkills?](#what-is-officeskills)
- [Why not just ask an AI to "use this template"?](#why-not-just-ask-an-ai-to-use-this-template)
- [Highlights](#highlights)
- [How it works](#how-it-works)
- [Structure-aware, not just style-aware](#structure-aware-not-just-style-aware)
- [The three skills](#the-three-skills)
- [The Brand Kit](#the-brand-kit)
- [Prerequisites & installation](#prerequisites--installation)
- [Quick start](#quick-start)
- [Project status](#project-status)
- [Development](#development)
- [License & acknowledgements](#license--acknowledgements)

---

## What is OfficeSkills?

**OfficeSkills** is an **agent-skill bundle** - for Claude Code, Codex, and compatible AI agents - that turns a company's existing Office template into a **reusable brand memory**, then writes new documents that stay faithful to it.

You give it one branded `.docx`, `.pptx`, or `.xlsx`. It **extracts** the brand - theme colors and fonts, named styles, the document's *structure*, layouts, cover anchors, logos and tables - into a portable **Brand Profile**. From then on, every document it **generates** is built *from the original template shell* and uses *only* the artifacts the template actually defines. Each format stays in its own lane: a Word template makes Word documents, a deck makes decks, a workbook makes workbooks - there is no cross-format conversion.

> **The core guarantee: off-brand output is impossible by construction.** No generator ever writes a literal style name, hex color, or font - those live only in the Brand Profile, and `verify` refuses a profile that points at anything the template doesn't contain. There is no "creative" path that drifts from the brand.

### Why not just ask an AI to "use this template"?

General-purpose document skills generate *freely* and only loosely imitate a reference file - fonts drift, the palette wanders, the corporate structure is lost. OfficeSkills is the opposite: narrow and faithful. It learns the template as a set of **rules and reusable parts**, remembers them in a `brand-kit/`, and **respects them** across an unlimited number of documents.

|  | General-purpose Office skills | **OfficeSkills** |
|---|---|---|
| Mental model | "create a nice document" | "fill the company's template" |
| Brand fidelity | best-effort imitation | **by construction** - opens from the shell, applies only its artifacts |
| Reusability | re-explain the brand every time | **extract once**, reuse forever via `brand-kit/` |
| Structure | free-form | **respects the template's cover → contents → body order** |
| Guardrails | none | `verify` fails if a profile references a missing style/layout/range |

---

## Highlights

- 🎯 **Brand-faithful by construction** - generation opens from the real template shell and applies only its named styles, theme colors, fonts and layouts. The content model is brand-agnostic; the **Brand Profile** is the single source of brand truth.
- 🧠 **Extract once, reuse forever** - a portable `brand-kit/<name>/` is the template's memory; every later document reads it. No re-explaining the brand.
- 🏛️ **Structure-aware** - captures the template's *ordered skeleton* (e.g. **cover → table of contents → body**) and tags each component as a **fixed structure to keep in order** or a **style to use on demand**. ([details](#structure-aware-not-just-style-aware))
- ✅ **Enforced, not just promised** - `verify` opens the shell and **fails** if a role resolves to a style/layout/named-range that doesn't exist. Deterministic checks also cover allowed styles, palette adherence, residual template text, broken tables and language rules.
- 🧩 **One shared engine** - a single profile schema, resolver, OOXML layer and QA gate underpin all three formats. The Word vertical is the reference implementation; PowerPoint and Excel build on the same foundation.
- 🗂️ **Full artifact catalog** - records OOXML parts, styles, media, layouts, formulas and named ranges, so an agent can reason about anything the template exposes - even artifacts it can't yet regenerate.
- 🔓 **Self-contained & MIT** - pure `python-docx` / `python-pptx` / `openpyxl` + OOXML. No cloud, no external services, no vendor lock-in.

---

## How it works

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

> **Same format throughout.** A `.docx` template yields `.docx` documents, a `.pptx` yields `.pptx`, a `.xlsx` yields `.xlsx`. Each skill is its own lane: there is no cross-format conversion.

1. **Extract** unpacks the template's OOXML and records its brand: theme, named styles mapped to semantic **roles**, the **document structure** (the ordered skeleton plus which parts are fixed vs free), layouts, cover anchors, logos, and a complete artifact catalog. The original file is kept **byte-for-byte** as the *shell*.
2. **Generate** turns your content into an **IntermediateDocument** of brand-agnostic typed blocks (heading, paragraph, callout, list, table, …). A **pure resolver** maps each block to the concrete brand artifact from the profile, fills the shell **in the template's structural order**, and saves.
3. **Verify / QA** runs deterministic checks - every role resolves to a real artifact, only allowed styles are used, the palette holds, no residual template text remains, tables are intact - and, when LibreOffice is available, a visual pass.

---

## Structure-aware, not just style-aware

Most "use my template" tools copy *styling*. OfficeSkills also learns the template's **document structure** and reproduces it. During extraction it detects the ordered skeleton - typically **cover → table of contents → body** - and annotates every captured component with **how it is used**:

- **Structural** parts (cover, table of contents) are kept **in order** in every generated document - the cover is filled in place, the TOC is preserved and refreshed.
- **Freeform** parts (headings, callouts, lists, tables, quotes, captions) are styles to **use on demand**, in whatever order your content needs.

So a generated report opens with the company cover, keeps a live table of contents, and fills **only the body** with your content - exactly like a person starting from the corporate template, rather than a bare wall of text.

---

## The three skills

| Skill | Format | Generates |
|---|---|---|
| **`brand-docx`** | Word `.docx` | reports, letters, memos: cover, headings, paragraphs, callouts, quotes, captions, lists, tables - in the template's structural order |
| **`brand-pptx`** | PowerPoint `.pptx` | decks: title / section / content slides from the template's real masters & layouts, with long-text splitting |
| **`brand-xlsx`** | Excel `.xlsx` | workbooks: fills named cells & regions while **preserving formulas** and workbook structure |

All three expose the same three verbs: **`extract` → `verify` → `generate`** - each skill is self-contained and **same-format** (a Word template makes Word documents, never a deck or a sheet).

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

`brand-kit/` lives either in your **project** (`./brand-kit/`, versionable, wins) or **globally** (`~/.claude/brand-kit/`, reusable across projects). It is the template's portable memory - copy the folder and the brand travels with it.

---

## Prerequisites & installation

### Required (core extract / generate / deterministic QA)

- **Python ≥ 3.10**
- Python packages (installed via `requirements.txt`): `python-docx>=1.1`, `python-pptx>=1.0`, `openpyxl>=3.1`, `lxml>=5.0`, `Pillow>=10.0`

```bash
git clone https://github.com/ferdinandobons/office-skills.git
cd office-skills
python3 -m venv .venv && . .venv/bin/activate     # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

### Optional (visual QA - render-based checks)

Needed only for the **visual** verification pass; their absence degrades gracefully and never blocks extraction, generation, or deterministic QA.

- **LibreOffice** (`soffice`) - headless render to PDF
- **Poppler** (`pdftoppm`) - PDF → PNG

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

`doctor` lists each dependency (present or missing) and prints the exact install command for anything missing - it never fails the run.

### Install as an agent skill

**Claude Code** (loads all three skills + the shared engine together):

```text
/plugin marketplace add ferdinandobons/office-skills
/plugin install office-skills
```

**Codex** (clone + symlink the skills):

```bash
git clone https://github.com/ferdinandobons/office-skills.git ~/.codex/office-skills
cd ~/.codex/office-skills && python3 -m venv .venv && . .venv/bin/activate && pip install -r requirements.txt
mkdir -p ~/.codex/skills
for s in brand-docx brand-pptx brand-xlsx; do ln -s ~/.codex/office-skills/skills/$s ~/.codex/skills/$s; done
```

> Restart/reload the agent after installing if the skills don't appear immediately.

---

## Quick start

### With an AI agent (the intended experience)

Just describe what you want and attach a template:

> "Use this company Word template and write a report on the history of Napoleon."

The agent activates `brand-docx`, extracts a Brand Profile from the template (or reuses an existing one), turns your request into the template's structure, generates the `.docx` from the original shell, runs QA, and returns the file.

### Direct CLI (the engine - for tests & debugging)

```bash
# 1) Extract the brand from a template into a reusable kit
python scripts/brandkit/cli.py extract --name acme --template template.docx --scope project

# 2) Verify the profile (role mapping + QA; fails if a role points at a missing artifact)
python scripts/brandkit/cli.py verify --name acme --scope auto --qa auto

# 3) Generate a new on-brand document from structured content
python scripts/brandkit/cli.py generate --name acme --input idoc.json --output out.docx --scope auto --qa auto
```

The content you pass in (`idoc.json`) is an **IntermediateDocument** - brand-agnostic typed blocks. Notice there is **no style, color or font anywhere**: the profile resolves all of that.

```json
{
  "cover": { "title": "Quarterly Review", "fields": { "doc_id": "RPT-001" } },
  "blocks": [
    { "type": "heading", "level": 1, "text": "Highlights" },
    { "type": "paragraph", "text": "This paragraph resolves to the brand body style." },
    { "type": "callout", "intent": "info", "text": "The profile chooses the callout style." },
    { "type": "list", "items": [{ "text": "List styling comes from the profile." }] },
    { "type": "table", "columns": ["Area", "Status"], "rows": [["Pipeline", "Healthy"], ["Delivery", "Green"]] }
  ]
}
```

PowerPoint uses the same `IntermediateDocument`; Excel uses a `GridDocument` (named-region fills, formulas preserved).

---

## Project status

**Alpha.** The Word vertical (`brand-docx`) is the reference implementation, verified end-to-end on real templates; PowerPoint and Excel share the engine and are catching up.

| Area | Status |
|---|---|
| Shared engine (profile schema, resolver, OOXML, CLI, dual store) | ✅ working |
| `brand-docx` - extract → verify → generate | ✅ working |
| Document **structure** extraction & order-aware generation | ✅ working |
| Brand-guarantee enforcement (`verify` fails on missing artifacts) | ✅ working |
| Deterministic QA (L0: styles, palette, residual text, tables, language) | ✅ working |
| `brand-pptx` - roles from real layouts, basic generation | 🚧 early |
| `brand-xlsx` - named-region fills, formula-preserving | 🚧 early |
| Visual QA (LibreOffice render + auto-repair loop) | 🔭 planned |
| Charts / SmartArt / multi-page section templates | 🔭 catalogued, regeneration staged |

Visual Word overflow needs LibreOffice, since Word lays out at render time.

---

## Development

```bash
python3 -m venv .venv && . .venv/bin/activate
pip install -r requirements.txt pytest
PYTHONPATH=scripts pytest -q        # docx / pptx / security / integration / smoke suites
```

> **Never commit real templates or company assets.** `brand-kit/` and `generated/` are intentionally git-ignored, and `tests/test_no_proprietary.py` fails the build if any Office binary is tracked outside `tests/fixtures` (or a vendored proprietary import sneaks in). See [`CONTRIBUTING.md`](CONTRIBUTING.md) and the frozen vocabulary in [`CONVENTIONS.md`](CONVENTIONS.md).

---

## License & acknowledgements

- This project's own code is **[MIT](LICENSE)** © 2026 Ferdinando Bonsegna.
- Self-contained: the OOXML engine is re-implemented from scratch; it does **not** vendor any proprietary or third-party Office tooling. See [`NOTICE`](NOTICE).
