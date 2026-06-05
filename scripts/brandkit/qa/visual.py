# SPDX-License-Identifier: MIT
"""Two-stage visual audit (model-free engine half).

This module adds the *engine* side of a two-stage visual audit on top of the
existing L0 deterministic gate (``qa/checks_deterministic.py``). It never calls a
model: the qualitative L2 judgement is the orchestrator's job (driven by
``SKILL.md``). The engine only:

  * **renders** an output (``.docx``/``.pptx``/``.xlsx``) to per-page PNGs using
    the env-detected external tools (``soffice`` + ``pdftoppm``) -- env-aware and
    degrading cleanly to ``[]`` when they are absent;
  * runs **L1 deterministic pixel proxies** that catch defects L0 cannot see
    because they depend on the *rendered* layout (blank pages, content bleeding
    past the printable margins, zero rendered pages). Each defect becomes one
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
import re
import subprocess
import tempfile
from pathlib import Path
from typing import Union

from PIL import Image

from brandkit import doctor
from brandkit.profile import schema
from brandkit.qa.model import Finding

# A proxy input is either an already-opened PIL image or a path to a PNG.
ImageInput = Union["Image.Image", str, Path]

# ---------------------------------------------------------------------------
# Render configuration
# ---------------------------------------------------------------------------
DEFAULT_DPI: int = 100        # 850x1100 for US-Letter portrait; enough for proxies
RENDER_TIMEOUT_S: int = 90

# ---------------------------------------------------------------------------
# L1 proxy thresholds (module constants, motivated by measured render data).
# Content pages render to a mean luma of ~240-252 and a blank page to ~255, so
# the blank threshold sits just under pure white; the ink fraction guards against
# a near-white page that still carries a faint legitimate mark.
# ---------------------------------------------------------------------------
BLANK_LUMA_MIN: float = 252.0      # mean luma at/above which a page is "near blank"
BLANK_INK_FRAC_MAX: float = 0.004  # max fraction of "ink" pixels to still call it blank
INK_LUMA_THRESHOLD: int = 180      # a pixel below this counts as "ink"
EDGE_MARGIN_FRAC: float = 0.012    # edge band width = 1.2% of the side (printable margin)
EDGE_INK_FRAC_MAX: float = 0.004   # ink allowed in an edge band before flagging

# ---------------------------------------------------------------------------
# Manifest constants
# ---------------------------------------------------------------------------
MANIFEST_FILENAME: str = "visual_manifest.json"
MANIFEST_SCHEMA_VERSION: str = "visual-manifest-1"

_PAGE_RE = re.compile(r"page-(\d+)\.png$")


# ---------------------------------------------------------------------------
# Render: env-aware, isolated, mockable, degrades to []
# ---------------------------------------------------------------------------
def renderers_available() -> bool:
    """Return True iff both ``soffice`` and ``pdftoppm`` are present.

    Delegates to :func:`brandkit.doctor.probe` so there is a single source of
    truth for renderer detection (the same flag ``doctor`` reports).
    """
    return bool(doctor.probe().get("visual_qa"))


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
) -> list[Path]:
    """Render a ``.docx``/``.pptx``/``.xlsx`` to an ordered list of per-page PNGs.

    Pipeline: ``soffice --headless --convert-to pdf --outdir <tmp> <document>``
    then ``pdftoppm -png -r <dpi> <pdf> <out_dir>/page``. Returns the PNG paths
    ordered numerically (``page-1.png``, ``page-2.png``, ..., ``page-10.png``).

    Clean degrade: returns ``[]`` if :func:`renderers_available` is False, if
    ``soffice``/``pdftoppm`` fail (non-zero rc), time out, or produce no
    PDF/PNG. NEVER raises -- the render is a side artifact whose failure must not
    break the gate.

    The PDF is written to an internal ``TemporaryDirectory``; the PNGs go to
    ``out_dir`` (the working/out dir the caller passes), never inside the
    generated document. ``out_dir`` is created if missing and any pre-existing
    ``page-*.png`` are removed first so a repair-loop re-run is not confused by
    stale frames.
    """
    if not renderers_available():
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

        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            soffice = subprocess.run(
                ["soffice", "--headless", "--convert-to", "pdf",
                 "--outdir", str(tmp), str(document)],
                capture_output=True, timeout=timeout_s, check=False,
            )
            if soffice.returncode != 0:
                return []
            pdfs = list(tmp.glob("*.pdf"))
            if not pdfs:
                return []
            pdf = pdfs[0]
            toppm = subprocess.run(
                ["pdftoppm", "-png", "-r", str(dpi), str(pdf), str(out_dir / "page")],
                capture_output=True, timeout=timeout_s, check=False,
            )
            if toppm.returncode != 0:
                return []
    except (subprocess.TimeoutExpired, OSError):
        return []

    return sorted(out_dir.glob("page-*.png"), key=_page_sort_key)


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


def _ink_fraction(luma: Image.Image, box: tuple[int, int, int, int] | None = None) -> float:
    """Fraction of pixels darker than :data:`INK_LUMA_THRESHOLD` in ``box``.

    ``box`` is an ``(left, top, right, bottom)`` crop or None for the whole image.
    """
    region = luma.crop(box) if box else luma
    total = region.width * region.height
    if total <= 0:
        return 0.0
    # Histogram bin i = count of pixels with luma == i.
    hist = region.histogram()
    ink = sum(hist[: INK_LUMA_THRESHOLD])
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
    ink = sum(hist[: INK_LUMA_THRESHOLD]) / total
    if mean >= luma_min and ink <= ink_frac_max:
        return [Finding(
            "visual.blank_page",
            schema.Severity.WARNING.value,
            f"page {page_index + 1} renders blank/near-blank "
            f"(mean luma {mean:.1f}, ink {ink:.4f})",
            location=f"page:{page_index + 1}",
        )]
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
            findings.append(Finding(
                "visual.edge_bleed",
                schema.Severity.WARNING.value,
                f"ink in {side} margin band on page {page_index + 1} "
                f"(ink {frac:.4f} > {edge_ink_frac_max})",
                location=f"page:{page_index + 1}:{side}",
            ))
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
        return [Finding(
            "visual.no_pages",
            schema.Severity.WARNING.value,
            "output rendered zero pages",
        )]
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
        regions = [r.get("region") for r in skeleton if isinstance(r, dict) and r.get("region")]
        items.append({
            "id": "regions_present",
            "what": f"Each expected region appears in the expected order: {regions}",
            "derived_from": "structure.skeleton[*].region + order",
            "severity_hint": "WARNING",
        })

    if anchors.get("cover"):
        items.append({
            "id": "cover_correct",
            "what": "The cover shows the bound title, no duplicate title, no residual demo prompt",
            "derived_from": "anchors.cover + comprehension.cover_slots",
            "severity_hint": "WARNING",
        })

    demo_region = sub.get("demo_region") if isinstance(sub, dict) else None
    if anchors.get("demo_region") or (isinstance(demo_region, dict) and demo_region.get("present")):
        items.append({
            "id": "no_residual_placeholder",
            "what": "No template placeholder/demo text is visible in the rendered output",
            "derived_from": "surface.<kind>.demo_region + comprehension.cover_slots[*].demo_value",
            "severity_hint": "WARNING",
        })

    if theme.get("colors"):
        items.append({
            "id": "palette_on_brand",
            "what": "On-screen colors belong to the brand palette",
            "derived_from": "theme.colors + theme.palette_roles",
            "severity_hint": "INFO",
        })

    if role_index:
        items.append({
            "id": "roles_styled",
            "what": "Semantic blocks (heading/list/callout/table/quote/caption) "
                    "appear with the brand style, not 'Normal'",
            "derived_from": "roles._index",
            "severity_hint": "WARNING",
        })

    # Constant items: always relevant for every kind; reinforced by L1 findings.
    items.append({
        "id": "no_overlap",
        "what": "No overlapping or clipped text/shapes",
        "derived_from": "constant; reinforced by l1_findings visual.edge_bleed",
        "severity_hint": "WARNING",
    })
    items.append({
        "id": "no_blank_pages",
        "what": "No unexpected blank/broken pages",
        "derived_from": "constant; reinforced by l1_findings visual.blank_page",
        "severity_hint": "WARNING",
    })

    if _profile_has_charts(profile):
        items.append({
            "id": "charts_rendered",
            "what": "Every chart is drawn correctly (axes/legend/data), not an empty box",
            "derived_from": "roles._index chart.* / components / artifact_catalog charts",
            "severity_hint": "WARNING",
        })

    oc = qa.get("overflow_capability")
    if oc in (schema.OverflowCapability.RENDER.value,
              schema.OverflowCapability.ESTIMATOR.value,
              schema.OverflowCapability.CELLFIT.value):
        items.append({
            "id": "overflow_clean",
            "what": "No content beyond the printable margins",
            "derived_from": f"qa.overflow_capability={oc}",
            "severity_hint": "WARNING",
        })

    return items


def _profile_has_charts(profile: dict) -> bool:
    """True if the profile evidences any chart (role, component, or catalog)."""
    roles = profile.get("roles") or {}
    if any(isinstance(r, str) and r.startswith("chart") for r in roles.get("_index", [])):
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
) -> Path:
    """Build and write ``<out_dir>/visual_manifest.json`` (a SIDE artifact).

    The checklist is derived from the profile (see :func:`derive_visual_checklist`).
    Returns the manifest path. Deterministic JSON (indent=2, trailing newline);
    PNG paths are stored RELATIVE to ``out_dir`` for portability. When
    ``renderers_ok`` is False the manifest carries ``"degraded": true``, empty
    ``pages``/``l1_findings``, but a populated ``checklist`` so the orchestrator
    still knows what it would have inspected.
    """
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    document = Path(document)

    pages: list[dict] = []
    if renderers_ok:
        for i, png in enumerate(png_paths):
            png = Path(png)
            w, h = _png_dimensions(png)
            try:
                rel = png.relative_to(out_dir).as_posix()
            except ValueError:
                rel = png.name
            pages.append({
                "index": i + 1,
                "png": rel,
                "width": w,
                "height": h,
                "orientation": _orientation(w, h),
            })

    manifest: dict = {
        "schema_version": MANIFEST_SCHEMA_VERSION,
        "kind": profile.get("kind"),
        "profile_name": (profile.get("identity") or {}).get("name"),
        "document": document.name,
        "renderers_available": bool(renderers_ok),
        "qa_mode": "deep",
        "dpi": DEFAULT_DPI,
        "pages": pages,
        "l1_findings": [
            {
                "check": f.check,
                "severity": f.severity,
                "message": f.message,
                "location": f.location,
            }
            for f in (l1_findings if renderers_ok else [])
        ],
        "checklist": derive_visual_checklist(profile),
        "instructions": (
            "Open each PNG. For each checklist item, judge PASS/FAIL against the "
            "rendered pages, taking l1_findings into account. If any item FAILS, "
            "repair the IntermediateDocument/content and regenerate, then re-run "
            "the audit. Do NOT call any model from the engine; this judgment is "
            "yours."
        ),
    }
    if not renderers_ok:
        manifest["degraded"] = True

    path = out_dir / MANIFEST_FILENAME
    path.write_text(
        json.dumps(manifest, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )
    return path


# ---------------------------------------------------------------------------
# Default side-artifact location helper (used by run_qa / cli when none given)
# ---------------------------------------------------------------------------
def default_out_dir(target: str | Path) -> Path:
    """Return the conventional side-artifact dir next to ``target`` (never inside).

    ``<parent>/<stem>.visual`` -- e.g. ``out.docx`` -> ``out.visual``.
    """
    p = Path(target)
    return p.parent / (p.stem + ".visual")
