# Installation & prerequisites

BrandDocs is **local-first**: it runs entirely on your machine, with no cloud
service. Before first use you install a small set of prerequisites, then add the
skills to your agent.

> **One important difference from pure-markdown skill bundles:** the three skills
> (`brand-docx`, `brand-pptx`, `brand-xlsx`) share a single Python engine in
> [`scripts/brandkit/`](../scripts/brandkit). Each skill is a thin `SKILL.md` plus
> a shim that calls that engine. **Copying a single skill folder on its own will
> not work**: install the whole repository (plugin, submodule, or clone), so the
> engine travels with the skills.

---

## 1. Prerequisites

### Required: core extract / generate / deterministic QA

- **Python ≥ 3.10**: Python packages (from [`requirements.txt`](../requirements.txt)):
  `python-docx>=1.1`, `python-pptx>=1.0`, `openpyxl>=3.1`, `lxml>=5.0`, `Pillow>=10.0`

```bash
git clone https://github.com/ferdinandobons/brand-docs.git
cd brand-docs
python3 -m venv .venv && . .venv/bin/activate     # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

### Visual QA dependencies (keep the visual gate on)

The **visual QA gate is a core part of the QA step** and runs by default whenever
its tools are present: it renders the output and catches layout problems the
deterministic checks can't (text overflow, blank pages, clipping, stale demo text)
which in Word only surface at render time. Generation still runs without them (the
engine degrades gracefully to L0), but **install them so the visual gate stays
on** - it is treated as fundamental, not optional.

Simplest path - one command auto-detects your package manager and installs them:

```bash
bash scripts/setup_visual_qa.sh
```

What it installs:

- **LibreOffice** (`soffice`): headless render to PDF (the only hard system dep)
- **Poppler** (`pdftoppm`): PDF → PNG (or `python -m pip install PyMuPDF` instead)
- **Tesseract** (`tesseract`): optional OCR for rendered residual placeholder/demo text

Or install manually:

```bash
# macOS (Homebrew)
brew install --cask libreoffice && brew install poppler tesseract

# Debian / Ubuntu
sudo apt-get install -y libreoffice poppler-utils tesseract-ocr

# Fedora
sudo dnf install -y libreoffice poppler-utils tesseract

# Windows
# winget install TheDocumentFoundation.LibreOffice
#   + Poppler via conda-forge or a prebuilt binary on PATH (or: pip install PyMuPDF)
#   + optional: winget install UB-Mannheim.TesseractOCR
```

> **macOS note:** the gate considers `soffice` usable as long as it actually
> renders - a LibreOffice whose code signature was knocked loose by an update or a
> quarantine removal still works (the functional render is authoritative, not a
> pristine signature). If `doctor` ever reports `soffice` unusable but it does run,
> a clean reinstall (`brew reinstall --cask libreoffice`) restores the signature.

### Check what's available

```bash
python scripts/brandkit/cli.py doctor
```

`doctor` lists each dependency (present or missing) and prints the exact install
command for anything missing. It is a preflight gate: it exits non-zero if any
*required* Python package is missing, and exits zero when all required packages are
present (missing optional renderers/OCR only downgrade visual QA, they never fail
the run). Two flags refine it:

```bash
python scripts/brandkit/cli.py doctor --json   # machine-readable probe dict; skips the human report
python scripts/brandkit/cli.py doctor --fast   # skip the slow LibreOffice render smoke-test; marks visual QA "not probed"
```

---

## 2. Install as an agent skill

### Claude Code plugin (recommended)

Loads all three skills plus the shared engine together:

```text
/plugin marketplace add ferdinandobons/brand-docs
/plugin install brand-docs@brand-docs
```

After installing the plugin, install the Python prerequisites above so the engine
can run.

### Git submodule (whole repo)

```bash
git submodule add https://github.com/ferdinandobons/brand-docs.git .agents/brand-docs
cd .agents/brand-docs && python3 -m venv .venv && . .venv/bin/activate && pip install -r requirements.txt
```

### Codex / other agents (clone + symlink)

```bash
git clone https://github.com/ferdinandobons/brand-docs.git ~/.codex/brand-docs
cd ~/.codex/brand-docs && python3 -m venv .venv && . .venv/bin/activate && pip install -r requirements.txt
mkdir -p ~/.codex/skills
for s in brand-docx brand-pptx brand-xlsx; do ln -s ~/.codex/brand-docs/skills/$s ~/.codex/skills/$s; done
```

The symlink keeps each skill's engine reachable, since `scripts/brandkit/` stays
in the cloned repo.

> Restart/reload the agent after installing if the skills don't appear immediately.

---

## 3. Updating

**Claude Code plugin:**

```text
/plugin update brand-docs@brand-docs
```

If the update doesn't pick up new changes, refresh the marketplace source and reinstall:

```text
/plugin uninstall brand-docs@brand-docs
/plugin marketplace remove brand-docs
/plugin marketplace add ferdinandobons/brand-docs
/plugin install brand-docs@brand-docs
```

**Git submodule:**

```bash
git submodule update --remote .agents/brand-docs
```

**Clone:** `git pull` in the cloned repo; the symlinks keep pointing at the updated skills.
