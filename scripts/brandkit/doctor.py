# SPDX-License-Identifier: MIT
from __future__ import annotations

import importlib.util
import shutil
import subprocess
import tempfile
from pathlib import Path


REQUIRED = ("docx", "pptx", "openpyxl", "lxml", "PIL")
OPTIONAL_BINARIES = {"soffice": "visual DOCX/PPTX/XLSX render", "pdftoppm": "PDF to PNG visual proof"}
OPTIONAL_BINARY_PROBES = {
    "soffice": ("--headless", "--version"),
    "pdftoppm": ("-v",),
}
BINARY_PROBE_TIMEOUT_S = 10
VISUAL_PIPELINE_TIMEOUT_S = 45


def probe() -> dict:
    deps = {name: importlib.util.find_spec(name) is not None for name in REQUIRED}
    bins: dict[str, bool] = {}
    paths: dict[str, str | None] = {}
    errors: dict[str, str] = {}
    for name in OPTIONAL_BINARIES:
        ok, path, error = _probe_binary(name)
        bins[name] = ok
        paths[name] = path
        if error:
            errors[name] = error
    visual_ok = all(bins.values())
    visual_error = None
    if visual_ok:
        visual_ok, visual_error = _probe_visual_pipeline(paths)
        if visual_error:
            errors["visual_qa"] = visual_error
    return {
        "python_deps": deps,
        "binaries": bins,
        "binary_paths": paths,
        "binary_errors": errors,
        "visual_qa": visual_ok,
    }


def _probe_binary(name: str) -> tuple[bool, str | None, str | None]:
    path = shutil.which(name)
    if path is None:
        return False, None, "not found on PATH"
    args = [path, *OPTIONAL_BINARY_PROBES.get(name, ("--version",))]
    try:
        proc = subprocess.run(
            args,
            capture_output=True,
            timeout=BINARY_PROBE_TIMEOUT_S,
            check=False,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        return False, path, str(exc)
    if proc.returncode != 0:
        stderr = _short_output(proc.stderr)
        stdout = _short_output(proc.stdout)
        detail = stderr or stdout or f"exit code {proc.returncode}"
        return False, path, detail
    return True, path, None


def _short_output(data) -> str:
    if data is None:
        return ""
    if isinstance(data, bytes):
        text = data.decode("utf-8", errors="replace")
    else:
        text = str(data)
    return " ".join(text.strip().split())[:240]


def _probe_visual_pipeline(paths: dict[str, str | None]) -> tuple[bool, str | None]:
    """Smoke-test the actual DOCX -> PDF -> PNG render pipeline.

    Version probes catch missing executables, but they do not prove LibreOffice can
    run headless conversion in the current environment. This tiny render keeps
    ``doctor`` aligned with what ``visual.render_to_pngs`` needs.
    """
    try:
        from docx import Document
    except Exception as exc:
        return False, f"cannot create probe docx: {exc}"

    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td)
        docx_path = tmp / "probe.docx"
        pdf_dir = tmp / "pdf"
        png_dir = tmp / "png"
        pdf_dir.mkdir()
        png_dir.mkdir()
        doc = Document()
        doc.add_paragraph("BrandDocs visual QA probe")
        doc.save(docx_path)

        soffice_path = paths.get("soffice") or "soffice"
        soffice = subprocess.run(
            [
                soffice_path,
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(pdf_dir),
                str(docx_path),
            ],
            capture_output=True,
            timeout=VISUAL_PIPELINE_TIMEOUT_S,
            check=False,
        )
        if soffice.returncode != 0:
            return False, "soffice convert failed: " + (
                _short_output(soffice.stderr)
                or _short_output(soffice.stdout)
                or f"exit code {soffice.returncode}"
            )

        pdfs = list(pdf_dir.glob("*.pdf"))
        if not pdfs:
            return False, "soffice convert produced no PDF"

        pdftoppm_path = paths.get("pdftoppm") or "pdftoppm"
        toppm = subprocess.run(
            [
                pdftoppm_path,
                "-png",
                "-r",
                "50",
                str(pdfs[0]),
                str(png_dir / "page"),
            ],
            capture_output=True,
            timeout=VISUAL_PIPELINE_TIMEOUT_S,
            check=False,
        )
        if toppm.returncode != 0:
            return False, "pdftoppm failed: " + (
                _short_output(toppm.stderr)
                or _short_output(toppm.stdout)
                or f"exit code {toppm.returncode}"
            )
        if not list(png_dir.glob("page-*.png")):
            return False, "pdftoppm produced no PNG"
    return True, None


def print_report() -> None:
    status = probe()
    for name, ok in status["python_deps"].items():
        print(f"python:{name}: {'ok' if ok else 'missing'}")
    for name, ok in status["binaries"].items():
        if ok:
            label = "ok"
        elif status.get("binary_paths", {}).get(name):
            label = "unusable"
        else:
            label = "missing"
        msg = f"binary:{name}: {label} ({OPTIONAL_BINARIES[name]})"
        error = status.get("binary_errors", {}).get(name)
        if error and label == "unusable":
            msg += f" - {error}"
        print(msg)
    if status["visual_qa"]:
        print("visual QA: L1 proxy + L2 manifest available")
    else:
        suffix = ""
        if status.get("binary_errors", {}).get("visual_qa"):
            suffix = f" ({status['binary_errors']['visual_qa']})"
        print(f"visual QA disabled; L0 deterministic QA remains available{suffix}")
