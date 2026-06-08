#!/usr/bin/env bash
# SPDX-License-Identifier: MIT
#
# One-command setup for the BrandDocs visual QA gate's SYSTEM dependencies
# (LibreOffice + Poppler, plus optional Tesseract for OCR residual-text checks).
# These are not pip-installable, so this script auto-detects the platform package
# manager and installs them. It is idempotent: anything already present is skipped.
#
# The Python packages still come from requirements.txt:
#   python -m pip install -r requirements.txt
#
# After running this, verify everything with:
#   python scripts/brandkit/cli.py doctor
set -euo pipefail

have() { command -v "$1" >/dev/null 2>&1; }

echo "BrandDocs: installing visual QA system dependencies (LibreOffice, Poppler, Tesseract)"

if have soffice && have pdftoppm; then
  echo "  soffice + pdftoppm already on PATH; nothing to install."
elif have brew; then
  # macOS / Linuxbrew
  have soffice || brew install --cask libreoffice
  have pdftoppm || brew install poppler
  have tesseract || brew install tesseract
elif have apt-get; then
  sudo apt-get update
  sudo apt-get install -y libreoffice poppler-utils tesseract-ocr
elif have dnf; then
  sudo dnf install -y libreoffice poppler-utils tesseract
elif have pacman; then
  sudo pacman -S --needed --noconfirm libreoffice-still poppler tesseract
else
  cat >&2 <<'EOF'
No supported package manager found (brew/apt-get/dnf/pacman).
Install these manually, then re-run `python scripts/brandkit/cli.py doctor`:
  - LibreOffice (provides `soffice`)   https://www.libreoffice.org/download/
  - Poppler (provides `pdftoppm`)      (or: python -m pip install PyMuPDF)
  - Tesseract (optional OCR)           https://github.com/tesseract-ocr/tesseract
Windows: winget install TheDocumentFoundation.LibreOffice
EOF
  exit 1
fi

echo
echo "Done. Verify with:"
echo "  python scripts/brandkit/cli.py doctor"
