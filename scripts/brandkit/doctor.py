# SPDX-License-Identifier: MIT
from __future__ import annotations

import importlib.util
import shutil


REQUIRED = ("docx", "pptx", "openpyxl", "lxml", "PIL")
OPTIONAL_BINARIES = {"soffice": "visual DOCX/PPTX/XLSX render", "pdftoppm": "PDF to PNG visual proof"}


def probe() -> dict:
    deps = {name: importlib.util.find_spec(name) is not None for name in REQUIRED}
    bins = {name: shutil.which(name) is not None for name in OPTIONAL_BINARIES}
    return {"python_deps": deps, "binaries": bins, "visual_qa": all(bins.values())}


def print_report() -> None:
    status = probe()
    for name, ok in status["python_deps"].items():
        print(f"python:{name}: {'ok' if ok else 'missing'}")
    for name, ok in status["binaries"].items():
        print(f"binary:{name}: {'ok' if ok else 'missing'} ({OPTIONAL_BINARIES[name]})")
    if not status["visual_qa"]:
        print("visual QA disabled; L0 deterministic QA remains available")

