# SPDX-License-Identifier: MIT
"""Unit conversions and constants for OOXML geometry and typography.

OOXML mixes several measurement systems. This module is the single source of
truth for converting between them. All public conversion functions are pure and
side-effect free.

Unit primer
-----------
- **EMU** (English Metric Unit): the universal absolute unit in DrawingML and
  the WordprocessingML page/section geometry. ``914400 EMU = 1 inch``,
  ``360000 EMU = 1 cm``, ``12700 EMU = 1 pt``. Geometry in this project is
  stored as **EMU integers** (with an ``_emu`` suffix on the field name).
- **pt** (point): ``1 pt = 1/72 inch``. Font sizes in the OOXML ``w:sz``
  attribute are in **half-points** (``sz="36"`` means 18 pt).
- **dxa** (twentieths of a point, a.k.a. "twips"): the WordprocessingML unit for
  paragraph indents, table widths, tab stops, etc. ``1 pt = 20 dxa``,
  ``1 inch = 1440 dxa``, ``1 dxa = 635 EMU``.
- **inch / cm**: human-facing absolute units. Conversions here go through EMU.

Rounding policy: every ``*_to_emu`` returns an ``int`` (EMU is an integer unit);
every ``emu_to_*`` returns a ``float`` unless the caller asks otherwise. Use the
explicit ``round`` helpers when an integer OOXML attribute is required.
"""
from __future__ import annotations

# ---------------------------------------------------------------------------
# Base constants (exact integers; these are definitional, not approximations).
# ---------------------------------------------------------------------------
EMU_PER_INCH: int = 914400
EMU_PER_CM: int = 360000
EMU_PER_MM: int = 36000
EMU_PER_PT: int = 12700
EMU_PER_DXA: int = 635  # 914400 / 1440

DXA_PER_INCH: int = 1440
DXA_PER_PT: int = 20

PT_PER_INCH: int = 72
HALF_POINTS_PER_PT: int = 2

# ---------------------------------------------------------------------------
# Common page sizes in EMU (width, height) — portrait orientation.
# ---------------------------------------------------------------------------
PAGE_SIZE_A4_EMU: tuple[int, int] = (7560000, 10692000)        # 210 x 297 mm
PAGE_SIZE_LETTER_EMU: tuple[int, int] = (7772400, 10058400)    # 8.5 x 11 in
PAGE_SIZE_A3_EMU: tuple[int, int] = (10692000, 15120000)       # 297 x 420 mm

# Default 16:9 slide size used by PowerPoint (EMU).
SLIDE_SIZE_16_9_EMU: tuple[int, int] = (12192000, 6858000)
SLIDE_SIZE_4_3_EMU: tuple[int, int] = (9144000, 6858000)


# ---------------------------------------------------------------------------
# EMU <-> inch
# ---------------------------------------------------------------------------
def inches_to_emu(inches: float) -> int:
    """Convert inches to EMU (rounded to the nearest integer EMU)."""
    return int(round(inches * EMU_PER_INCH))


def emu_to_inches(emu: int) -> float:
    """Convert EMU to inches (float)."""
    return emu / EMU_PER_INCH


# ---------------------------------------------------------------------------
# EMU <-> cm / mm
# ---------------------------------------------------------------------------
def cm_to_emu(cm: float) -> int:
    """Convert centimetres to EMU (rounded)."""
    return int(round(cm * EMU_PER_CM))


def emu_to_cm(emu: int) -> float:
    """Convert EMU to centimetres (float)."""
    return emu / EMU_PER_CM


def mm_to_emu(mm: float) -> int:
    """Convert millimetres to EMU (rounded)."""
    return int(round(mm * EMU_PER_MM))


def emu_to_mm(emu: int) -> float:
    """Convert EMU to millimetres (float)."""
    return emu / EMU_PER_MM


# ---------------------------------------------------------------------------
# EMU <-> pt
# ---------------------------------------------------------------------------
def pt_to_emu(pt: float) -> int:
    """Convert points to EMU (rounded)."""
    return int(round(pt * EMU_PER_PT))


def emu_to_pt(emu: int) -> float:
    """Convert EMU to points (float)."""
    return emu / EMU_PER_PT


# ---------------------------------------------------------------------------
# EMU <-> dxa (twips)
# ---------------------------------------------------------------------------
def dxa_to_emu(dxa: int) -> int:
    """Convert dxa (twentieths of a point) to EMU (rounded)."""
    return int(round(dxa * EMU_PER_DXA))


def emu_to_dxa(emu: int) -> int:
    """Convert EMU to dxa (rounded to nearest integer dxa)."""
    return int(round(emu / EMU_PER_DXA))


# ---------------------------------------------------------------------------
# pt <-> dxa, pt <-> inch
# ---------------------------------------------------------------------------
def pt_to_dxa(pt: float) -> int:
    """Convert points to dxa (rounded)."""
    return int(round(pt * DXA_PER_PT))


def dxa_to_pt(dxa: int) -> float:
    """Convert dxa to points (float)."""
    return dxa / DXA_PER_PT


def pt_to_inches(pt: float) -> float:
    """Convert points to inches (float)."""
    return pt / PT_PER_INCH


def inches_to_pt(inches: float) -> float:
    """Convert inches to points (float)."""
    return inches * PT_PER_INCH


# ---------------------------------------------------------------------------
# Half-point font sizes (the OOXML w:sz / a:rPr sz conventions)
# ---------------------------------------------------------------------------
def halfpt_to_pt(half_points: int) -> float:
    """Convert OOXML half-points (``w:sz`` value) to points.

    ``<w:sz w:val="36"/>`` -> ``18.0`` pt.
    """
    return half_points / HALF_POINTS_PER_PT


def pt_to_halfpt(pt: float) -> int:
    """Convert points to OOXML half-points (rounded integer ``w:sz`` value)."""
    return int(round(pt * HALF_POINTS_PER_PT))


def centipt_to_pt(centi_points: int) -> float:
    """Convert DrawingML hundredths-of-a-point (``a:rPr sz``) to points.

    DrawingML run sizes use centipoints: ``sz="1800"`` -> ``18.0`` pt.
    """
    return centi_points / 100.0


def pt_to_centipt(pt: float) -> int:
    """Convert points to DrawingML centipoints (rounded integer ``a:rPr sz``)."""
    return int(round(pt * 100.0))
