# SPDX-License-Identifier: MIT
"""Theme color extraction, transform resolution, and WCAG contrast.

This module is the single place where OOXML colors become concrete hex strings.
It does three jobs:

1. **Parse** ``theme1.xml``'s ``a:clrScheme`` into the 12 canonical theme slots,
   resolving ``a:sysClr`` (which carries a ``lastClr`` cached value) to a plain
   hex. (See :func:`parse_theme_colors`.)
2. **Transform** a base hex through the DrawingML modifiers that the rest of the
   pipeline carries on color references — ``lumMod``/``lumOff`` (luminance
   modulation, the modern Office tint/shade), ``tint``/``shade`` (legacy linear
   blends toward white/black), and ``alpha`` (recorded, not blended). (See
   :func:`apply_transforms`.)
3. **Judge** a foreground/background pair with the WCAG 2.x relative-luminance
   contrast ratio. (See :func:`contrast_ratio`.)

All hex values exchanged across this module are **6-digit uppercase RRGGBB with
no leading '#'** — the normal form produced by :func:`normalize_hex`. Inputs may
be ``#rgb`` / ``#rrggbb`` / ``rrggbb`` / ``RRGGBB`` and are normalized.

The 12 theme slots, in OOXML document order, are:

    dk1 lt1 dk2 lt2 accent1 accent2 accent3 accent4 accent5 accent6 hlink folHlink

(``dk1``/``lt1`` are frequently authored as ``a:sysClr`` — windowText / window —
hence the ``lastClr`` resolution.)
"""
from __future__ import annotations

from typing import Optional

try:  # lxml is a hard dependency, but keep parse helpers import-safe.
    from lxml import etree as _etree
except Exception:  # pragma: no cover - exercised only in degraded envs
    _etree = None  # type: ignore[assignment]

# ---------------------------------------------------------------------------
# Canonical slot order and the DrawingML namespace.
# ---------------------------------------------------------------------------
THEME_SLOTS: tuple[str, ...] = (
    "dk1", "lt1", "dk2", "lt2",
    "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink",
)

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _a(tag: str) -> str:
    """Clark-notation qualified name for an ``a:`` (DrawingML) local name."""
    return f"{{{A_NS}}}{tag}"


# ---------------------------------------------------------------------------
# Hex normalization
# ---------------------------------------------------------------------------
def normalize_hex(value: str) -> str:
    """Normalize any accepted hex spelling to 6-digit uppercase ``RRGGBB``.

    Accepts ``#rgb``, ``#rrggbb``, ``rgb``, ``rrggbb`` in any case. A leading
    ``#`` is stripped; a 3-digit shorthand is expanded by doubling each nibble.

    Raises:
        ValueError: if ``value`` is not a 3- or 6-digit hex string.
    """
    s = value.strip().lstrip("#").upper()
    if len(s) == 3 and all(c in "0123456789ABCDEF" for c in s):
        return "".join(c * 2 for c in s)
    if len(s) == 6 and all(c in "0123456789ABCDEF" for c in s):
        return s
    raise ValueError(f"not a valid hex color: {value!r}")


def is_hex(value: str) -> bool:
    """Return True if ``value`` is a parseable hex color (3 or 6 digits)."""
    try:
        normalize_hex(value)
        return True
    except (ValueError, AttributeError):
        return False


def hex_to_rgb(value: str) -> tuple[int, int, int]:
    """Return the ``(r, g, b)`` 0-255 tuple for a hex color."""
    s = normalize_hex(value)
    return int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16)


def rgb_to_hex(r: int, g: int, b: int) -> str:
    """Return the normalized ``RRGGBB`` hex for 0-255 channel values (clamped)."""
    def clamp(x: int) -> int:
        return max(0, min(255, int(round(x))))

    return f"{clamp(r):02X}{clamp(g):02X}{clamp(b):02X}"


# ---------------------------------------------------------------------------
# theme1.xml a:clrScheme parsing
# ---------------------------------------------------------------------------
def parse_theme_colors(theme_xml: bytes) -> dict[str, str]:
    """Parse a ``theme1.xml`` part into the 12 theme slots as hex strings.

    The ``<a:clrScheme>`` element holds one child per slot. Each child wraps
    either an ``<a:srgbClr val="RRGGBB"/>`` (explicit) or an
    ``<a:sysClr val="windowText" lastClr="000000"/>`` (system color whose cached
    resolved value is in ``lastClr``). This resolver returns ``lastClr`` for
    ``sysClr`` (falling back to a sensible default for windowText/window when
    ``lastClr`` is absent).

    Args:
        theme_xml: raw bytes of the ``word/theme/theme1.xml`` (or pptx/xlsx
            equivalent) part.

    Returns:
        A dict keyed by every name in :data:`THEME_SLOTS` -> normalized hex.
        Slots missing from the scheme are omitted (callers should treat a
        missing slot as "inherit / unknown", never as black).

    Raises:
        RuntimeError: if lxml is unavailable.
    """
    if _etree is None:  # pragma: no cover
        raise RuntimeError("lxml is required to parse theme colors")
    root = _etree.fromstring(theme_xml)
    scheme = root.find(f".//{_a('clrScheme')}")
    out: dict[str, str] = {}
    if scheme is None:
        return out
    for slot in THEME_SLOTS:
        node = scheme.find(_a(slot))
        if node is None:
            continue
        hexval = _resolve_color_node(node)
        if hexval is not None:
            out[slot] = hexval
    return out


def _resolve_color_node(node) -> Optional[str]:
    """Resolve a single ``a:clrScheme`` child to a hex, or None if unresolvable."""
    srgb = node.find(_a("srgbClr"))
    if srgb is not None and srgb.get("val"):
        return normalize_hex(srgb.get("val"))
    sysclr = node.find(_a("sysClr"))
    if sysclr is not None:
        last = sysclr.get("lastClr")
        if last:
            return normalize_hex(last)
        # Reasonable defaults for the two common system colors.
        val = (sysclr.get("val") or "").lower()
        if val == "windowtext":
            return "000000"
        if val == "window":
            return "FFFFFF"
    return None


# ---------------------------------------------------------------------------
# DrawingML color transforms
# ---------------------------------------------------------------------------
def _lin(channel: int) -> float:
    """sRGB 8-bit channel -> linear float in [0,1] (no gamma applied here)."""
    return channel / 255.0


def apply_lum_mod(hex_value: str, lum_mod: float) -> str:
    """Apply ``lumMod`` (luminance modulation): scale each channel by a factor.

    ``lum_mod`` is the *fraction* (e.g. ``0.75`` for the OOXML ``lumMod=75000``).
    This multiplies each linear channel, the standard Office darken operation.
    """
    r, g, b = hex_to_rgb(hex_value)
    return rgb_to_hex(r * lum_mod, g * lum_mod, b * lum_mod)


def apply_lum_off(hex_value: str, lum_off: float) -> str:
    """Apply ``lumOff`` (luminance offset): add a fraction of full scale.

    ``lum_off`` is the fraction (e.g. ``0.15`` for ``lumOff=15000``). Office
    pairs ``lumMod`` with ``lumOff`` to lighten: ``out = c*lumMod + 255*lumOff``.
    Apply :func:`apply_lum_mod` first, then this, to reproduce Office tints.
    """
    r, g, b = hex_to_rgb(hex_value)
    add = 255.0 * lum_off
    return rgb_to_hex(r + add, g + add, b + add)


def apply_tint(hex_value: str, tint: float) -> str:
    """Apply legacy ``tint`` (linear blend toward WHITE).

    ``tint`` is the fraction in [0,1] (``tint=40000`` -> ``0.40``). Office:
    ``out = c*tint + 255*(1-tint)`` — higher tint keeps more of the base color.
    """
    r, g, b = hex_to_rgb(hex_value)
    return rgb_to_hex(
        r * tint + 255.0 * (1 - tint),
        g * tint + 255.0 * (1 - tint),
        b * tint + 255.0 * (1 - tint),
    )


def apply_shade(hex_value: str, shade: float) -> str:
    """Apply legacy ``shade`` (linear blend toward BLACK).

    ``shade`` is the fraction in [0,1] (``shade=60000`` -> ``0.60``). Office:
    ``out = c*shade`` — higher shade keeps more of the base color.
    """
    r, g, b = hex_to_rgb(hex_value)
    return rgb_to_hex(r * shade, g * shade, b * shade)


def apply_transforms(hex_value: str, transforms: dict) -> str:
    """Apply an ordered set of DrawingML modifiers to a base hex.

    ``transforms`` is a mapping that may contain any of:
        ``lumMod`` / ``lumOff`` / ``tint`` / ``shade`` — each either a fraction
        (``0.75``) or an OOXML thousandths integer (``75000``); values > 1 are
        interpreted as thousandths and divided by 100000. ``alpha`` is accepted
        and *recorded conceptually* but does not change the opaque hex (alpha is
        a separate channel; callers that need it read it from the source dict).

    Application order matches Office: ``shade`` -> ``tint`` -> ``lumMod`` ->
    ``lumOff``. Unknown keys are ignored. Returns the normalized resulting hex;
    an empty / falsy ``transforms`` returns the normalized input unchanged.
    """
    out = normalize_hex(hex_value)
    if not transforms:
        return out

    def frac(key: str) -> Optional[float]:
        if key not in transforms or transforms[key] is None:
            return None
        v = float(transforms[key])
        return v / 100000.0 if v > 1.0 else v

    shade = frac("shade")
    if shade is not None:
        out = apply_shade(out, shade)
    tint = frac("tint")
    if tint is not None:
        out = apply_tint(out, tint)
    lm = frac("lumMod")
    if lm is not None:
        out = apply_lum_mod(out, lm)
    lo = frac("lumOff")
    if lo is not None:
        out = apply_lum_off(out, lo)
    return out


def resolve_theme_color(
    theme_colors: dict[str, str],
    slot: str,
    transforms: Optional[dict] = None,
) -> Optional[str]:
    """Resolve a theme-slot reference (+optional transforms) to a final hex.

    Args:
        theme_colors: the dict from :func:`parse_theme_colors`.
        slot: one of :data:`THEME_SLOTS` (e.g. ``"accent1"``). Also accepts the
            two WordprocessingML aliases ``tx1``/``bg1`` -> ``dk1``/``lt1`` and
            ``tx2``/``bg2`` -> ``dk2``/``lt2``.
        transforms: optional modifier dict passed to :func:`apply_transforms`.

    Returns:
        The normalized resolved hex, or ``None`` if the slot is unknown/missing
        (caller treats ``None`` as "inherited — not a brand violation").
    """
    alias = {"tx1": "dk1", "bg1": "lt1", "tx2": "dk2", "bg2": "lt2"}
    key = alias.get(slot, slot)
    base = theme_colors.get(key)
    if base is None:
        return None
    return apply_transforms(base, transforms or {})


# ---------------------------------------------------------------------------
# WCAG contrast
# ---------------------------------------------------------------------------
def relative_luminance(hex_value: str) -> float:
    """Return the WCAG 2.x relative luminance (0.0-1.0) of a color.

    Implements the sRGB linearization from WCAG 2.1 (gamma-expanded channels,
    weighted 0.2126 R + 0.7152 G + 0.0722 B).
    """
    def chan(c: int) -> float:
        cs = c / 255.0
        return cs / 12.92 if cs <= 0.03928 else ((cs + 0.055) / 1.055) ** 2.4

    r, g, b = hex_to_rgb(hex_value)
    return 0.2126 * chan(r) + 0.7152 * chan(g) + 0.0722 * chan(b)


def contrast_ratio(fg_hex: str, bg_hex: str) -> float:
    """Return the WCAG contrast ratio between two colors (1.0 .. 21.0).

    Order-independent. ``4.5`` is the AA threshold for normal text, ``3.0`` for
    large text; ``7.0`` is AAA for normal text.
    """
    l1 = relative_luminance(fg_hex)
    l2 = relative_luminance(bg_hex)
    lighter, darker = max(l1, l2), min(l1, l2)
    return (lighter + 0.05) / (darker + 0.05)


def meets_wcag(fg_hex: str, bg_hex: str, *, min_ratio: float = 4.5) -> bool:
    """Return True if the pair meets or exceeds ``min_ratio`` (default AA)."""
    return contrast_ratio(fg_hex, bg_hex) >= min_ratio
