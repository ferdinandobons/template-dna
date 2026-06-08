# SPDX-License-Identifier: MIT
"""Two-stage visual audit (model-free engine half).

This module adds the *engine* side of a two-stage visual audit on top of the
existing L0 deterministic gate (``qa/checks_deterministic.py``). It never calls a
model: the qualitative L2 judgement is the orchestrator's job (driven by
``SKILL.md``). The engine only:

  * **renders** an output (``.docx``/``.pptx``/``.xlsx``) to per-page PNGs using
    the env-detected tools (``soffice`` + ``pdftoppm`` by default, with optional
    PyMuPDF/``fitz`` PDF raster fallback) -- env-aware and degrading cleanly to
    ``[]`` when they are absent;
  * runs **L1 deterministic pixel proxies** that catch defects L0 cannot see
    because they depend on the *rendered* layout (blank pages, content bleeding
    past the printable margins, zero rendered pages). Optional OCR adds an
    advisory rendered-text signal for residual placeholders when ``tesseract`` is
    installed. Each defect becomes one
    :class:`~brandkit.qa.model.Finding` ``check="visual.<name>"``;
  * emits a structured **L2 manifest** (``visual_manifest.json``): the PNG paths
    per page plus a checklist derived from the profile (expected regions/roles,
    on-brand palette, no-overlap, no residual placeholder, correct cover, charts
    rendered) plus the L1 findings, for the orchestrator to judge and drive a
    repair loop.

The PNGs and the manifest are **side artifacts** written to a working/out dir
next to the output; the bytes of the generated document never change because of
the audit.

Design constraints (deliberate):
  * The L1 proxies accept ``PIL.Image.Image | str | Path`` so they are
    unit-testable without ``soffice`` (tests feed synthetic PIL images).
  * Nothing here raises on a render/IO failure: a side artifact's failure must
    never break the gate. Proxies on an unreadable image return ``[]``.
  * Thresholds are module constants, motivated by measured render data
    (US-Letter @100 DPI = 850x1100; content pages mean-luma ~240-252; a blank
    page ~255).
"""

from __future__ import annotations

import json
import platform
import re
import shutil
import subprocess
import tempfile
import time
from pathlib import Path
from typing import Union

from PIL import Image

from brandkit import doctor
from brandkit.profile import schema
from brandkit.qa import checks_deterministic
from brandkit.qa.model import Finding

# A proxy input is either an already-opened PIL image or a path to a PNG.
ImageInput = Union["Image.Image", str, Path]

# ---------------------------------------------------------------------------
# Render configuration
# ---------------------------------------------------------------------------
DEFAULT_DPI: int = 100  # 850x1100 for US-Letter portrait; enough for proxies
RENDER_TIMEOUT_S: int = 90
OCR_TIMEOUT_S: int = 30
# Aggregate ceiling across ALL pages so a many-page document cannot turn a deep/
# strict gate into an N x OCR_TIMEOUT_S hang; remaining pages are skipped (partial).
OCR_TOTAL_BUDGET_S: int = 120
OCR_TEXT_LIMIT: int = 4000
OCR_TERM_LIMIT: int = 80

# ---------------------------------------------------------------------------
# L1 proxy thresholds (module constants, motivated by measured render data).
# Content pages render to a mean luma of ~240-252 and a blank page to ~255, so
# the blank threshold sits just under pure white; the ink fraction guards against
# a near-white page that still carries a faint legitimate mark.
# ---------------------------------------------------------------------------
BLANK_LUMA_MIN: float = 252.0  # mean luma at/above which a page is "near blank"
BLANK_INK_FRAC_MAX: float = 0.004  # max fraction of "ink" pixels to still call it blank
INK_LUMA_THRESHOLD: int = 180  # a pixel below this counts as "ink"
EDGE_MARGIN_FRAC: float = 0.012  # edge band width = 1.2% of the side (printable margin)
EDGE_INK_FRAC_MAX: float = 0.004  # ink allowed in an edge band before flagging

# ---------------------------------------------------------------------------
# Manifest constants
# ---------------------------------------------------------------------------
MANIFEST_FILENAME: str = "visual_manifest.json"
MANIFEST_SCHEMA_VERSION: str = "visual-manifest-1"

_PAGE_RE = re.compile(r"page-(\d+)\.png$")
_LAST_RENDERER_STATUS: dict | None = None


# ---------------------------------------------------------------------------
# Render: env-aware, isolated, mockable, degrades to []
# ---------------------------------------------------------------------------
def renderers_available() -> bool:
    """Return True iff both ``soffice`` and ``pdftoppm`` are present.

    Delegates to :func:`brandkit.doctor.probe` so there is a single source of
    truth for renderer detection (the same flag ``doctor`` reports).
    """
    global _LAST_RENDERER_STATUS
    _LAST_RENDERER_STATUS = doctor.probe()
    return bool(_LAST_RENDERER_STATUS.get("visual_qa"))


def last_renderer_status() -> dict | None:
    """Return the last ``doctor.probe()`` result captured by renderer probing."""
    return _LAST_RENDERER_STATUS


def _page_sort_key(p: Path) -> int:
    """Numeric sort key from a ``page-<n>.png`` name (so page-10 > page-2)."""
    m = _PAGE_RE.search(p.name)
    return int(m.group(1)) if m else 0


def render_to_pngs(
    document: str | Path,
    out_dir: str | Path,
    *,
    dpi: int = DEFAULT_DPI,
    timeout_s: int = RENDER_TIMEOUT_S,
    check_available: bool = True,
    quicklook_only: bool = False,
    render_errors: list[str] | None = None,
    render_warnings: list[str] | None = None,
) -> list[Path]:
    """Render a ``.docx``/``.pptx``/``.xlsx`` to an ordered list of per-page PNGs.

    Pipeline: ``soffice --headless --convert-to pdf --outdir <tmp> <document>``,
    then ``pdftoppm -png -r <dpi> <pdf> <out_dir>/page``. If ``pdftoppm`` is
    missing or fails, optional PyMuPDF (``fitz``) is used as a PDF raster fallback
    when installed. Returns the PNG paths ordered numerically (``page-1.png``,
    ``page-2.png``, ..., ``page-10.png``).

    Clean degrade: returns ``[]`` if :func:`renderers_available` is False, if
    ``soffice``/``pdftoppm`` fail (non-zero rc), time out, or produce no
    PDF/PNG. NEVER raises -- the render is a side artifact whose failure must not
    break the gate. ``check_available=False`` is used by the QA gate after it has
    already performed the availability probe, avoiding a second ``soffice``
    launch on macOS. ``quicklook_only=True`` skips LibreOffice entirely and emits
    a single macOS Quick Look thumbnail, used when the primary pipeline is known
    to be broken. ``render_errors`` / ``render_warnings`` are optional out-params
    for diagnostics.

    The PDF is written to an internal ``TemporaryDirectory``; the PNGs go to
    ``out_dir`` (the working/out dir the caller passes), never inside the
    generated document. ``out_dir`` is created if missing and any pre-existing
    ``page-*.png`` are removed first so a repair-loop re-run is not confused by
    stale frames.
    """
    if check_available and not renderers_available():
        _append_render_error(render_errors, "visual renderers unavailable")
        return []

    document = Path(document)
    out_dir = Path(out_dir)
    if not document.is_file():
        return []

    try:
        out_dir.mkdir(parents=True, exist_ok=True)
        # Idempotent per call: clear stale frames from an earlier repair iteration.
        for stale in out_dir.glob("page-*.png"):
            try:
                stale.unlink()
            except OSError:
                pass

        if quicklook_only:
            return _render_quicklook_thumbnail(
                document,
                out_dir,
                timeout_s=timeout_s,
                render_errors=render_errors,
                render_warnings=render_warnings,
                reason="because LibreOffice visual pipeline is unavailable",
            )

        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            pdf_dir = tmp / "pdf"
            pdf_dir.mkdir()
            lo_profile = tmp / "lo-profile"
            soffice = subprocess.run(
                _soffice_convert_cmd(document, pdf_dir, lo_profile),
                capture_output=True,
                timeout=timeout_s,
                check=False,
            )
            if soffice.returncode != 0:
                _append_render_error(
                    render_errors,
                    "soffice convert failed: "
                    + (
                        _short_output(soffice.stderr)
                        or _short_output(soffice.stdout)
                        or f"exit code {soffice.returncode}"
                    ),
                )
                return _render_quicklook_thumbnail(
                    document,
                    out_dir,
                    timeout_s=timeout_s,
                    render_errors=render_errors,
                    render_warnings=render_warnings,
                    reason="after soffice failure",
                )
            pdfs = list(pdf_dir.glob("*.pdf"))
            if not pdfs:
                _append_render_error(render_errors, "soffice convert produced no PDF")
                return _render_quicklook_thumbnail(
                    document,
                    out_dir,
                    timeout_s=timeout_s,
                    render_errors=render_errors,
                    render_warnings=render_warnings,
                    reason="after soffice produced no PDF",
                )
            pdf = pdfs[0]
            pngs = _rasterize_pdf_to_pngs(
                pdf,
                out_dir,
                dpi=dpi,
                timeout_s=timeout_s,
                render_errors=render_errors,
                render_warnings=render_warnings,
            )
            if pngs:
                return pngs
            else:
                return _render_quicklook_thumbnail(
                    document,
                    out_dir,
                    timeout_s=timeout_s,
                    render_errors=render_errors,
                    render_warnings=render_warnings,
                    reason="after PDF rasterization failure",
                )
    except subprocess.TimeoutExpired as exc:
        _append_render_error(render_errors, f"render timed out: {exc}")
        return []
    except OSError as exc:
        _append_render_error(render_errors, f"render failed: {exc}")
        return []


def _rasterize_pdf_to_pngs(
    pdf: Path,
    out_dir: Path,
    *,
    dpi: int,
    timeout_s: int,
    render_errors: list[str] | None,
    render_warnings: list[str] | None,
) -> list[Path]:
    """Rasterize an existing PDF via pdftoppm, then optional PyMuPDF fallback."""
    pdftoppm = shutil.which("pdftoppm")
    if pdftoppm:
        try:
            toppm = subprocess.run(
                [pdftoppm, "-png", "-r", str(dpi), str(pdf), str(out_dir / "page")],
                capture_output=True,
                timeout=timeout_s,
                check=False,
            )
        except subprocess.TimeoutExpired as exc:
            _append_render_error(render_errors, f"pdftoppm timed out: {exc}")
        except OSError as exc:
            _append_render_error(render_errors, f"pdftoppm failed: {exc}")
        else:
            if toppm.returncode == 0:
                pngs = sorted(out_dir.glob("page-*.png"), key=_page_sort_key)
                if pngs:
                    return pngs
                _append_render_error(render_errors, "pdftoppm produced no PNG")
            else:
                _append_render_error(
                    render_errors,
                    "pdftoppm failed: "
                    + (
                        _short_output(toppm.stderr)
                        or _short_output(toppm.stdout)
                        or f"exit code {toppm.returncode}"
                    ),
                )
    else:
        _append_render_error(render_errors, "pdftoppm unavailable")

    pngs, error = _rasterize_pdf_with_pymupdf(pdf, out_dir, dpi=dpi)
    if pngs:
        _append_render_warning(
            render_warnings,
            "PyMuPDF PDF raster fallback used; audit is degraded relative to pdftoppm baseline",
        )
        return pngs
    if error:
        _append_render_error(render_errors, f"PyMuPDF fallback failed: {error}")
    return []


def _rasterize_pdf_with_pymupdf(
    pdf: Path, out_dir: Path, *, dpi: int
) -> tuple[list[Path], str | None]:
    try:
        import fitz  # type: ignore[import-not-found]
    except Exception as exc:
        return [], f"fitz import failed: {exc}"
    doc = None
    try:
        doc = fitz.open(str(pdf))
        zoom = dpi / 72.0
        matrix = fitz.Matrix(zoom, zoom)
        for index, page in enumerate(doc, start=1):
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            pix.save(str(out_dir / f"page-{index}.png"))
    except Exception as exc:
        return [], str(exc)
    finally:
        if doc is not None and hasattr(doc, "close"):
            doc.close()
    pngs = sorted(out_dir.glob("page-*.png"), key=_page_sort_key)
    if not pngs:
        return [], "produced no PNG"
    return pngs, None


def _render_quicklook_thumbnail(
    document: Path,
    out_dir: Path,
    *,
    timeout_s: int,
    render_errors: list[str] | None,
    render_warnings: list[str] | None,
    reason: str,
) -> list[Path]:
    """Fallback renderer for macOS when LibreOffice crashes.

    Quick Look returns a single thumbnail, not a faithful multi-page render. It is
    therefore a degraded audit input: good enough to run pixel proxies and give
    the L2 checklist something visual to inspect, but never a replacement for the
    full LibreOffice PDF pipeline.
    """
    qlmanage = shutil.which("qlmanage")
    if qlmanage is None:
        _append_render_error(
            render_errors, "Quick Look fallback unavailable: qlmanage not found"
        )
        return []
    try:
        proc = subprocess.run(
            [qlmanage, "-t", "-s", "1200", "-o", str(out_dir), str(document)],
            capture_output=True,
            timeout=timeout_s,
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        _append_render_error(render_errors, f"Quick Look fallback timed out: {exc}")
        return []
    except OSError as exc:
        _append_render_error(render_errors, f"Quick Look fallback failed: {exc}")
        return []
    if proc.returncode != 0:
        _append_render_error(
            render_errors,
            "Quick Look fallback failed: "
            + (
                _short_output(proc.stderr)
                or _short_output(proc.stdout)
                or f"exit code {proc.returncode}"
            ),
        )
        return []

    # qlmanage names its thumbnail after the input file: ``<document.name>.png``.
    # Accept ONLY that exact name - never glob-fall-back to an arbitrary pre-existing
    # PNG in out_dir, which could stage a stale/unrelated frame as a bogus "render".
    produced = out_dir / f"{document.name}.png"
    if not produced.is_file():
        _append_render_error(
            render_errors, "Quick Look fallback produced no matching PNG"
        )
        return []

    target = out_dir / "page-1.png"
    if produced != target:
        try:
            if target.exists():
                target.unlink()
            produced.replace(target)
        except OSError as exc:
            _append_render_error(
                render_errors, f"Quick Look fallback could not stage PNG: {exc}"
            )
            return []

    _append_render_warning(
        render_warnings,
        f"Quick Look thumbnail fallback used {reason}; "
        "audit is first-page-only and degraded",
    )
    return [target]


def _soffice_convert_cmd(document: Path, pdf_dir: Path, lo_profile: Path) -> list[str]:
    """Build a headless LibreOffice conversion command with an isolated profile.

    LibreOffice on macOS may try to recover or lock the user's normal profile even
    for headless conversion. A per-render ``UserInstallation`` keeps the visual
    audit isolated and reduces crash/recovery-dialog coupling with the desktop app.
    """
    return [
        "soffice",
        f"-env:UserInstallation={lo_profile.as_uri()}",
        "--headless",
        "--nologo",
        "--nodefault",
        "--nolockcheck",
        "--norestore",
        "--nofirststartwizard",
        "--convert-to",
        "pdf",
        "--outdir",
        str(pdf_dir),
        str(document),
    ]


def _append_render_error(errors: list[str] | None, message: str) -> None:
    if errors is not None and message:
        errors.append(message)


def _append_render_warning(warnings: list[str] | None, message: str) -> None:
    if warnings is not None and message:
        warnings.append(message)


def _short_output(data) -> str:
    if data is None:
        return ""
    if isinstance(data, bytes):
        text = data.decode("utf-8", errors="replace")
    else:
        text = str(data)
    return " ".join(text.strip().split())[:240]


# ---------------------------------------------------------------------------
# L1 deterministic pixel proxies (pure; accept Image | path; never raise)
# ---------------------------------------------------------------------------
def _as_luma(img: ImageInput) -> Image.Image | None:
    """Open/convert ``img`` to an 8-bit luminance (``"L"``) image.

    Accepts an already-opened :class:`PIL.Image.Image` or a path to a PNG so the
    proxies are unit-testable without ``soffice``. Returns ``None`` (not an
    exception) on any PIL/IO error so the caller degrades instead of crashing.
    """
    try:
        if isinstance(img, Image.Image):
            return img.convert("L")
        im = Image.open(img)
        im.load()
        return im.convert("L")
    except (OSError, ValueError, TypeError):
        return None


def _ink_fraction(
    luma: Image.Image, box: tuple[int, int, int, int] | None = None
) -> float:
    """Fraction of pixels darker than :data:`INK_LUMA_THRESHOLD` in ``box``.

    ``box`` is an ``(left, top, right, bottom)`` crop or None for the whole image.
    """
    region = luma.crop(box) if box else luma
    total = region.width * region.height
    if total <= 0:
        return 0.0
    # Histogram bin i = count of pixels with luma == i.
    hist = region.histogram()
    ink = sum(hist[:INK_LUMA_THRESHOLD])
    return ink / total


def check_blank_page(
    image: ImageInput,
    *,
    page_index: int,
    luma_min: float = BLANK_LUMA_MIN,
    ink_frac_max: float = BLANK_INK_FRAC_MAX,
) -> list[Finding]:
    """Flag a blank/near-blank page (broken page, lost content, or overflow).

    A page whose mean luma is at/above ``luma_min`` AND whose ink fraction (pixels
    below :data:`INK_LUMA_THRESHOLD`) is at/below ``ink_frac_max`` yields one
    ``Finding(check="visual.blank_page", severity=WARNING,
    location="page:<n>")``. WARNING (never ERROR): a legitimately near-empty page
    exists (a minimal cover/separator), so this is a signal for the L2 judge, not
    a standalone gate failure.
    """
    luma = _as_luma(image)
    if luma is None:
        return []
    total = luma.width * luma.height
    if total <= 0:
        return []
    hist = luma.histogram()
    mean = sum(i * c for i, c in enumerate(hist)) / total
    ink = sum(hist[:INK_LUMA_THRESHOLD]) / total
    if mean >= luma_min and ink <= ink_frac_max:
        return [
            Finding(
                "visual.blank_page",
                schema.Severity.WARNING.value,
                f"page {page_index + 1} renders blank/near-blank "
                f"(mean luma {mean:.1f}, ink {ink:.4f})",
                location=f"page:{page_index + 1}",
            )
        ]
    return []


def check_edge_bleed(
    image: ImageInput,
    *,
    page_index: int,
    margin_frac: float = EDGE_MARGIN_FRAC,
    edge_ink_frac_max: float = EDGE_INK_FRAC_MAX,
) -> list[Finding]:
    """Flag content that touches/exceeds the printable margins (clipping/overflow).

    Defines four bands (top/bottom/left/right) ``margin_frac`` of the side wide.
    For each band the ink fraction is computed; a band over ``edge_ink_frac_max``
    yields one ``Finding(check="visual.edge_bleed", severity=WARNING,
    location="page:<n>:<side>")``. This is exactly the defect docx declares it
    wants to intercept via ``OverflowCapability.RENDER``. WARNING: a deliberate
    full-bleed cover/background is legitimate; the L2 judge distinguishes. The
    band is computed off the correct side, so portrait and landscape both work.
    """
    luma = _as_luma(image)
    if luma is None:
        return []
    w, h = luma.width, luma.height
    if w <= 0 or h <= 0:
        return []
    bw = max(1, int(round(w * margin_frac)))
    bh = max(1, int(round(h * margin_frac)))
    bands = {
        "top": (0, 0, w, bh),
        "bottom": (0, h - bh, w, h),
        "left": (0, 0, bw, h),
        "right": (w - bw, 0, w, h),
    }
    findings: list[Finding] = []
    for side, box in bands.items():
        frac = _ink_fraction(luma, box)
        if frac > edge_ink_frac_max:
            findings.append(
                Finding(
                    "visual.edge_bleed",
                    schema.Severity.WARNING.value,
                    f"ink in {side} margin band on page {page_index + 1} "
                    f"(ink {frac:.4f} > {edge_ink_frac_max})",
                    location=f"page:{page_index + 1}:{side}",
                )
            )
    return findings


def check_page_count_sane(
    images_or_paths: list[ImageInput],
    *,
    max_pages: int | None = None,
) -> list[Finding]:
    """Flag zero rendered pages when an output was expected.

    Robust and deterministic and defensive: an empty render is already covered by
    the clean degrade, but if an existing output produces zero pages that is a
    signal -> ``Finding(check="visual.no_pages", severity=WARNING)``. ``max_pages``
    is accepted for forward-compatibility (a future "too many pages" check) and
    is currently advisory only.
    """
    if not images_or_paths:
        return [
            Finding(
                "visual.no_pages",
                schema.Severity.WARNING.value,
                "output rendered zero pages",
            )
        ]
    return []


def run_visual_l1(png_paths: list[Path]) -> list[Finding]:
    """Run every L1 pixel proxy over an ordered PNG list and concatenate findings.

    Returns ``[]`` on an empty list (the ``no_pages`` signal is only meaningful
    when an output was expected and is raised by the caller via
    :func:`check_page_count_sane`; ``run_visual_l1`` itself stays a no-op so the
    clean-degrade path never adds noise). Each PNG is opened once.
    """
    findings: list[Finding] = []
    for i, path in enumerate(png_paths):
        luma = _as_luma(path)
        if luma is None:
            continue
        findings.extend(check_blank_page(luma, page_index=i))
        findings.extend(check_edge_bleed(luma, page_index=i))
    return findings


# ---------------------------------------------------------------------------
# Optional OCR signal (rendered-text advisory; never a hard dependency)
# ---------------------------------------------------------------------------
def run_visual_ocr(
    png_paths: list[Path],
    profile: dict,
    *,
    timeout_s: int = OCR_TIMEOUT_S,
) -> dict:
    """Run optional OCR over rendered PNGs and compare against captured template text.

    OCR is advisory: it helps the orchestrator spot visible stale placeholders,
    stale TOC/cache text, or demo copy that the OOXML text scan missed. Missing or
    failing OCR never blocks L0/L1. The report is designed to be embedded in the
    visual manifest and converted to WARNING findings only for concrete hits.
    """
    terms = _ocr_terms(profile)
    tesseract = shutil.which("tesseract")
    report: dict = {
        "engine": "tesseract",
        "available": bool(tesseract),
        "status": "ok" if tesseract else "unavailable",
        "terms_checked": terms,
        "pages": [],
        "hits": [],
        "errors": [],
    }
    if not png_paths:
        report["status"] = "not_run"
        report["reason"] = "no rendered pages"
        return report
    if not tesseract:
        report["reason"] = "tesseract not found on PATH"
        return report

    deadline = time.monotonic() + OCR_TOTAL_BUDGET_S
    for i, png in enumerate(png_paths, start=1):
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            report["errors"].append(
                {
                    "page": i,
                    "error": (
                        f"OCR total time budget ({OCR_TOTAL_BUDGET_S}s) exhausted; "
                        "remaining pages skipped"
                    ),
                }
            )
            break
        page_timeout = max(1, int(min(timeout_s, remaining)))
        text, error = _ocr_png(tesseract, Path(png), timeout_s=page_timeout)
        if error:
            report["errors"].append({"page": i, "error": error})
            continue
        page_hits = _ocr_hits(text, terms)
        limited_text, truncated = _limit_ocr_text(text)
        page = {
            "index": i,
            "text": limited_text,
            "text_truncated": truncated,
        }
        if page_hits:
            page["hits"] = page_hits
        report["pages"].append(page)
        for hit in page_hits:
            report["hits"].append({"page": i, **hit})

    if report["errors"] and not report["pages"]:
        report["status"] = "failed"
    elif report["errors"]:
        report["status"] = "partial"
    return report


def ocr_findings(report: dict) -> list[Finding]:
    """Convert OCR report hits into advisory visual findings."""
    findings: list[Finding] = []
    for hit in report.get("hits") or []:
        page = hit.get("page")
        term = hit.get("term")
        findings.append(
            Finding(
                "visual.ocr_residual_text",
                schema.Severity.WARNING.value,
                f"OCR saw captured template text still visible: {term!r}",
                location=f"page:{page}" if page else None,
            )
        )
    if report.get("status") in {"failed", "partial"} and report.get("errors"):
        findings.append(
            Finding(
                "visual.ocr_degraded",
                schema.Severity.INFO.value,
                "OCR visible-text scan did not complete for every rendered page",
            )
        )
    return findings


def _empty_ocr_report(profile: dict, *, reason: str) -> dict:
    return {
        "engine": "tesseract",
        "available": bool(shutil.which("tesseract")),
        "status": "not_run",
        "terms_checked": _ocr_terms(profile),
        "pages": [],
        "hits": [],
        "errors": [],
        "reason": reason,
    }


def _ocr_png(tesseract: str, png: Path, *, timeout_s: int) -> tuple[str, str | None]:
    try:
        # Capture BYTES (not text=True): tesseract output is usually UTF-8 but can
        # carry stray non-UTF-8 bytes, and text=True would hard-crash the whole QA
        # run on a UnicodeDecodeError instead of degrading. Decode tolerantly below.
        proc = subprocess.run(
            [tesseract, str(png), "stdout", "--psm", "6"],
            capture_output=True,
            timeout=timeout_s,
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        return "", f"tesseract timed out: {exc}"
    except OSError as exc:
        return "", f"tesseract failed: {exc}"
    if proc.returncode != 0:
        return "", (
            _short_output(proc.stderr)
            or _short_output(proc.stdout)
            or f"exit code {proc.returncode}"
        )
    return (proc.stdout or b"").decode("utf-8", errors="replace"), None


def _ocr_terms(profile: dict) -> list[str]:
    terms: list[str] = []
    for term in checks_deterministic.captured_template_texts(
        profile, include_surface_prompts=True
    ):
        cleaned = _normalize_ocr_space(str(term))
        if _has_ocr_signal(cleaned):
            terms.append(cleaned)
    seen: set[str] = set()
    out: list[str] = []
    for term in terms:
        key = _normalize_ocr_match(term)
        if key in seen:
            continue
        seen.add(key)
        out.append(term)
        if len(out) >= OCR_TERM_LIMIT:
            break
    return out


def _ocr_hits(text: str, terms: list[str]) -> list[dict]:
    normalized_text = _normalize_ocr_match(text)
    hits: list[dict] = []
    if not normalized_text:
        return hits
    for term in terms:
        needle = _normalize_ocr_match(term)
        if needle and needle in normalized_text:
            hits.append({"term": term})
    return hits


def _has_ocr_signal(term: str) -> bool:
    alnum = sum(1 for ch in term if ch.isalnum())
    return alnum >= 3 and len(term.strip()) >= 4


def _normalize_ocr_space(text: str) -> str:
    return " ".join(str(text).split())


def _normalize_ocr_match(text: str) -> str:
    return _normalize_ocr_space(text).casefold()


def _limit_ocr_text(text: str) -> tuple[str, bool]:
    """Normalize OCR whitespace, then cap to ``OCR_TEXT_LIMIT``.

    Returns the (possibly truncated) text together with whether truncation
    occurred. Both are derived from the same normalized string so the
    advisory ``text_truncated`` flag can never disagree with the stored
    text (the raw tesseract output is dominated by per-line newlines, so a
    flag computed from its length would over-report truncation).
    """
    text = _normalize_ocr_space(text)
    return text[:OCR_TEXT_LIMIT], len(text) > OCR_TEXT_LIMIT


# ---------------------------------------------------------------------------
# L2 manifest (model-free): PNG paths + profile-derived checklist + L1 findings
# ---------------------------------------------------------------------------
def _orientation(width: int, height: int) -> str:
    return "landscape" if width > height else "portrait"


def _png_dimensions(path: Path) -> tuple[int, int]:
    """Return (width, height) for a PNG, or (0, 0) if unreadable."""
    try:
        with Image.open(path) as im:
            return im.width, im.height
    except (OSError, ValueError):
        return 0, 0


def derive_visual_checklist(profile: dict) -> list[dict]:
    """Build the L2 checklist from the profile (model-free).

    Each item is ``{"id","what","derived_from","severity_hint"}``. "Constant"
    items (no_overlap, no_blank_pages) are always present; derived items are
    included only when the backing data exists (e.g. ``charts_rendered`` only when
    a chart role/component is present, ``cover_correct`` only when ``anchors.cover``
    exists). ``derived_from`` keeps every item traceable so the orchestrator knows
    *why* it is checking. Nothing here calls a model.
    """
    items: list[dict] = []
    structure = profile.get("structure") or {}
    skeleton = structure.get("skeleton") or []
    anchors = profile.get("anchors") or {}
    theme = profile.get("theme") or {}
    roles = profile.get("roles") or {}
    role_index = roles.get("_index") or []
    surface = profile.get("surface") or {}
    kind = profile.get("kind")
    sub = surface.get(kind) or {} if isinstance(surface, dict) else {}
    qa = profile.get("qa") or {}

    if skeleton:
        regions = [
            r.get("region") for r in skeleton if isinstance(r, dict) and r.get("region")
        ]
        items.append(
            {
                "id": "regions_present",
                "what": f"Each expected region appears in the expected order: {regions}",
                "derived_from": "structure.skeleton[*].region + order",
                "severity_hint": "WARNING",
            }
        )

    if anchors.get("cover"):
        items.append(
            {
                "id": "cover_correct",
                "what": "The cover shows the bound title, no duplicate title, no residual demo prompt",
                "derived_from": "anchors.cover + comprehension.cover_slots",
                "severity_hint": "WARNING",
            }
        )

    demo_region = sub.get("demo_region") if isinstance(sub, dict) else None
    if anchors.get("demo_region") or (
        isinstance(demo_region, dict) and demo_region.get("present")
    ):
        items.append(
            {
                "id": "no_residual_placeholder",
                "what": "No template placeholder/demo text is visible in the rendered output",
                "derived_from": "surface.<kind>.demo_region + comprehension.cover_slots[*].demo_value",
                "severity_hint": "WARNING",
            }
        )

    if theme.get("colors"):
        items.append(
            {
                "id": "palette_on_brand",
                "what": "On-screen colors belong to the brand palette",
                "derived_from": "theme.colors + theme.palette_roles",
                "severity_hint": "INFO",
            }
        )

    if role_index:
        items.append(
            {
                "id": "roles_styled",
                "what": "Semantic blocks (heading/list/callout/table/quote/caption) "
                "appear with the brand style, not 'Normal'",
                "derived_from": "roles._index",
                "severity_hint": "WARNING",
            }
        )

    # Constant items: always relevant for every kind; reinforced by L1 findings.
    items.append(
        {
            "id": "no_overlap",
            "what": "No overlapping or clipped text/shapes",
            "derived_from": "constant; reinforced by l1_findings visual.edge_bleed",
            "severity_hint": "WARNING",
        }
    )
    items.append(
        {
            "id": "no_blank_pages",
            "what": "No unexpected blank/broken pages",
            "derived_from": "constant; reinforced by l1_findings visual.blank_page",
            "severity_hint": "WARNING",
        }
    )

    if _profile_has_charts(profile):
        items.append(
            {
                "id": "charts_rendered",
                "what": "Every chart is drawn correctly (axes/legend/data), not an empty box",
                "derived_from": "roles._index chart.* / components / artifact_catalog charts",
                "severity_hint": "WARNING",
            }
        )

    oc = qa.get("overflow_capability")
    if oc in (
        schema.OverflowCapability.RENDER.value,
        schema.OverflowCapability.ESTIMATOR.value,
        schema.OverflowCapability.CELLFIT.value,
    ):
        items.append(
            {
                "id": "overflow_clean",
                "what": "No content beyond the printable margins",
                "derived_from": f"qa.overflow_capability={oc}",
                "severity_hint": "WARNING",
            }
        )

    return items


def _profile_has_charts(profile: dict) -> bool:
    """True if the profile evidences any chart (role, component, or catalog)."""
    roles = profile.get("roles") or {}
    if any(
        isinstance(r, str) and r.startswith("chart") for r in roles.get("_index", [])
    ):
        return True
    components = profile.get("components") or {}
    if any("chart" in str(k).lower() for k in components):
        return True
    catalog = profile.get("artifact_catalog") or {}
    if any("chart" in str(k).lower() for k in catalog):
        return True
    return False


def build_visual_manifest(
    *,
    profile: dict,
    document: str | Path,
    png_paths: list[Path],
    l1_findings: list[Finding],
    renderers_ok: bool,
    out_dir: str | Path,
    degraded: bool | None = None,
    environment_status: dict | None = None,
    ocr_report: dict | None = None,
    qa_mode: str = "deep",
) -> Path:
    """Build and write ``<out_dir>/visual_manifest.json`` (a SIDE artifact).

    The checklist is derived from the profile (see :func:`derive_visual_checklist`).
    Returns the manifest path. Deterministic JSON (indent=2, trailing newline);
    PNG paths are stored RELATIVE to ``out_dir`` for portability. When
    ``renderers_ok`` means the full render pipeline produced trustworthy visual
    proof. Degraded fallback images may still be listed in ``pages`` so the
    orchestrator has something concrete to inspect, but the manifest remains
    marked degraded and ``renderers_available`` stays false.
    """
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    document = Path(document)

    pages: list[dict] = []
    for i, png in enumerate(png_paths):
        png = Path(png)
        w, h = _png_dimensions(png)
        try:
            rel = png.relative_to(out_dir).as_posix()
        except ValueError:
            rel = png.name
        pages.append(
            {
                "index": i + 1,
                "png": rel,
                "width": w,
                "height": h,
                "orientation": _orientation(w, h),
            }
        )

    if degraded is None:
        degraded = not renderers_ok

    manifest: dict = {
        "schema_version": MANIFEST_SCHEMA_VERSION,
        "kind": profile.get("kind"),
        "profile_name": (profile.get("identity") or {}).get("name"),
        "document": document.name,
        "renderers_available": bool(renderers_ok),
        "qa_mode": qa_mode,
        "dpi": DEFAULT_DPI,
        "pages": pages,
        "l1_findings": [
            {
                "check": f.check,
                "severity": f.severity,
                "message": f.message,
                "location": f.location,
            }
            for f in l1_findings
        ],
        "ocr": ocr_report
        or _empty_ocr_report(profile, reason="not requested by caller"),
        "checklist": derive_visual_checklist(profile),
        "environment": visual_environment_summary(
            environment_status,
            renderers_ok=bool(renderers_ok),
            degraded=degraded,
        ),
        "instructions": (
            "Open each PNG. For each checklist item, judge PASS/FAIL against the "
            "rendered pages, taking l1_findings into account. If any item FAILS, "
            "repair the IntermediateDocument/content and regenerate, then re-run "
            "the audit. Do NOT call any model from the engine; this judgment is "
            "yours."
        ),
    }
    if degraded:
        manifest["degraded"] = True

    path = out_dir / MANIFEST_FILENAME
    path.write_text(
        json.dumps(manifest, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )
    return path


def visual_environment_summary(
    status: dict | None = None,
    *,
    renderers_ok: bool | None = None,
    degraded: bool | None = None,
) -> dict:
    """Return portable environment diagnostics for the visual manifest."""
    status = status or {}
    binary_paths = status.get("binary_paths") or {}
    binaries = status.get("binaries") or {}
    binary_errors = status.get("binary_errors") or {}
    optional_python = status.get("optional_python_deps") or {}
    ocr_paths = status.get("ocr_binary_paths") or {}
    ocr_binaries = status.get("ocr_binaries") or {}
    ocr_errors = status.get("ocr_binary_errors") or {}
    renderers = {}
    for name in doctor.OPTIONAL_BINARIES:
        path = binary_paths.get(name) or shutil.which(name)
        renderers[name] = {
            "available": bool(binaries.get(name, bool(path))),
            "path": path,
        }
        if binary_errors.get(name):
            renderers[name]["error"] = binary_errors[name]
    ocr = {}
    for name, purpose in doctor.OPTIONAL_OCR_BINARIES.items():
        path = ocr_paths.get(name) or shutil.which(name)
        ocr[name] = {
            "available": bool(ocr_binaries.get(name, bool(path))),
            "path": path,
            "purpose": purpose,
        }
        if ocr_errors.get(name):
            ocr[name]["error"] = ocr_errors[name]

    hints = doctor.install_hints(status) if status else []
    return {
        "platform": platform.platform(),
        "python": platform.python_version(),
        "visual_qa": bool(status.get("visual_qa", renderers_ok)),
        "degraded": bool(degraded),
        "renderers": renderers,
        "optional_python": {
            name: {
                "available": bool(optional_python.get(name)),
                "purpose": purpose,
            }
            for name, purpose in doctor.OPTIONAL_PYTHON.items()
        },
        "ocr": ocr,
        "install_hints": hints,
    }


# ---------------------------------------------------------------------------
# Default side-artifact location helper (used by run_qa / cli when none given)
# ---------------------------------------------------------------------------
def default_out_dir(target: str | Path) -> Path:
    """Return the conventional side-artifact dir next to ``target`` (never inside).

    ``<parent>/<filename>.visual`` -- e.g. ``out.docx`` -> ``out.docx.visual``.
    Keeping the extension prevents side-artifact collisions when different Office
    formats share the same basename in the same directory.
    """
    p = Path(target)
    return p.parent / (p.name + ".visual")
