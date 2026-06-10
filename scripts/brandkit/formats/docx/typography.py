# SPDX-License-Identifier: MIT
"""DOCX brand typography capture (font family, size, and color).

The brand's REAL visible typography often lives as DIRECT run-level formatting
(``w:rPr/w:rFonts`` / ``w:sz`` / ``w:color``) on the template's content rather than
in the named styles or the theme: a designed template may put everything in
``Normal`` with a direct Roboto / Montserrat override at 22 half-points in accent1.
Role inference (``roles.py``) and theme extraction read only named styles and
``theme1.xml``, so those direct values are never captured and a generated document
falls back to the ``docDefaults`` font/size/color.

This module captures the DOMINANT direct run typography, deterministically, as
THREE INDEPENDENT axes (font family, size, color) sampled in a SINGLE pass:

  - per role: the dominant explicit value among the runs that use the role's style
    -> ``role['appearance']['font'] = {'latin': <name>}`` /
    ``role['appearance']['size_hp'] = <int>`` /
    ``role['appearance']['color'] = {'kind': ...}``;
  - the document's effective body typography: the dominant explicit value across all
    body runs -> ``theme['fonts']['body']['latin'/'size_hp']`` and
    ``theme['text']['body']['color']`` - the fallbacks the generator applies to a
    paragraph whose role carries no captured value.

Each axis is independent: a role may carry a captured size but no captured font
(or vice versa). Only a clear DOMINANT is recorded per axis (at least
:data:`_MIN_RUNS` explicit values and a winner covering at least
:data:`_MIN_DOMINANCE` of them), with its dominance stored as a per-axis confidence
(``confidence`` for font, ``size_confidence`` for size, ``color_confidence`` for
color). Capture is deterministic (model-free).

The brand guarantee is preserved: every captured value is a FACT observed in the
template, stored in the profile, applied only via the resolver, and re-validated
against what the shell proves it contains by ``check_appearance_targets``
(fail-closed). This module is purely additive - it only populates the already-
reserved ``appearance`` field and additive ``theme.fonts.body`` / ``theme.text``
keys; a template with no dominant direct value leaves all of them untouched, so
behavior is unchanged.
"""

from __future__ import annotations

from collections import Counter
from typing import Optional

from docx.enum.dml import MSO_COLOR_TYPE, MSO_THEME_COLOR

from lxml import etree

from brandkit.common import color as colorutil
from brandkit.common.typography import (
    PALETTE_WHERE,
)
from brandkit.common.typography import (
    capture_appearance as _capture_appearance,
)
from brandkit.common.typography import (
    detect_pseudo_headings as _detect_pseudo_headings,
)
from brandkit.common.typography import (
    capture_palette_facts as _capture_palette_facts,
)
from brandkit.common.typography import (
    capture_paragraph_geometry as _capture_paragraph_geometry,
)
from brandkit.common.typography import (
    dominant as _dominant,
)
from brandkit.ooxml import names

# The generalizable capture engine (``capture_appearance`` / ``capture_palette_facts``
# and the helpers/constants they lean on) now lives in the format-neutral
# ``common.typography`` engine; this module keeps ONLY the docx-specific run readers
# (``_run_size_hp`` / ``_run_color`` / ``_iter_para_runs`` / ``_is_link_run``), wraps
# them into the structural ``RunFacts`` view the engine consumes, and delegates. The
# docx adapter emits its CURRENT WordprocessingML tokens VERBATIM (``_run_color``
# keeps emitting ``'text1'`` / ``'accent1'``; the palette key namespace is unchanged):
# nothing here re-keys to ``THEME_SLOTS`` - that normalization is the pptx/xlsx
# adapters' job only, which is exactly what keeps the docx frozen-hash anchor green.
__all__ = [
    "capture_fonts",
    "capture_palette",
    "capture_geometry",
    "capture_table_appearance",
    "capture_pseudo_headings",
    "collect_font_run_facts",
    "PALETTE_WHERE",
]

# python-docx's sentinel for a theme color that maps to no real slot; its
# ``xml_value`` is the truthy string ``"UNMAPPED"``. It is not a brand token (verify
# has no slot for it and apply cannot realize it), so it is never captured.
_UNMAPPED_THEME_TOKEN = MSO_THEME_COLOR.NOT_THEME_COLOR.xml_value

# WordprocessingML qualified-name builder for the link-color helper (it walks the
# raw ``w:hyperlink`` ancestor / ``w:rStyle`` to detect a link run).
_W = names.make_qn("w")

# The hyperlink character-style ids whose presence on a run marks it as link text
# even when it is NOT physically nested under a ``w:hyperlink`` element (a manually
# styled cross-reference). Closed, spec-fixed style ids (NOT brand literals).
_HYPERLINK_RSTYLES: frozenset[str] = frozenset({"Hyperlink", "FollowedHyperlink"})


def _run_size_hp(run) -> Optional[int]:
    """The run's EXPLICIT size as half-points (``w:sz@w:val``), or ``None``.

    ``run.font.size`` is an explicit-only ``Length`` (``None`` when the size is
    inherited from the style/theme), so a run that inherits its size contributes
    nothing. The half-point bucket ``round(pt * 2)`` matches OOXML's ``w:sz`` unit.
    """
    try:
        size = run.font.size
        if size is None:
            return None
        return round(size.pt * 2)
    except Exception:
        # A malformed measure the OOXML layer refuses to parse contributes nothing
        # to this axis - capture must never crash the extraction.
        return None


def _run_color(run) -> Optional[tuple[str, ...]]:
    """The run's EXPLICIT color as a hashable bucket key, or ``None``.

    ``run.font.color`` is a ``ColorFormat`` whose ``.type`` is ``None`` when the
    color is inherited. An RGB color buckets as ``('hex', <RRGGBB>)``; a THEME color
    buckets as ``('theme', <wordprocessingml token>)`` from the slot's
    ``.theme_color.xml_value`` (e.g. ``'accent1'``, ``'text1'``). AUTO / None / an
    unmapped theme slot contributes nothing (it is not a captured brand value).
    """
    try:
        color = run.font.color
        ctype = color.type
        if ctype == MSO_COLOR_TYPE.RGB and color.rgb is not None:
            return ("hex", str(color.rgb))
        if ctype == MSO_COLOR_TYPE.THEME:
            token = getattr(color.theme_color, "xml_value", None)
            # Drop the UNMAPPED sentinel: it is not a verifiable/appliable brand
            # token, so it must never enter the profile (keeps apply/verify in sync).
            if token and token != _UNMAPPED_THEME_TOKEN:
                return ("theme", token)
    except Exception:
        # A spec-valid-but-unmappable themeColor (e.g. 'none'/'phClr') makes
        # python-docx raise on access; that run contributes nothing to this axis
        # rather than crashing the extraction.
        return None
    return None


class _DocxRunFact:
    """A docx run, reduced to the structural :class:`~brandkit.common.typography.RunFacts`
    view the shared capture engine consumes.

    The three axes are read EXACTLY as the v1 inline capture did (``run.font.name`` or
    ``None`` / :func:`_run_size_hp` / :func:`_run_color`), so the engine sees the
    byte-identical votes. ``color`` keeps the docx WordprocessingML token namespace
    verbatim (``'text1'`` / ``'accent1'`` from ``_run_color``) - NO normalization to
    ``THEME_SLOTS`` - which is what keeps the docx palette keys (and the frozen-hash
    anchor) unchanged."""

    __slots__ = ("style_key", "text", "font_name", "size_hp", "color", "is_link")

    def __init__(self, run, style_key, *, is_link: bool) -> None:
        self.style_key = style_key
        self.text = run.text or ""
        self.font_name = run.font.name or None
        self.size_hp = _run_size_hp(run)
        self.color = _run_color(run)
        self.is_link = is_link


def _para_style_key(para) -> tuple[Optional[str], Optional[str]]:
    """The ``(style_id, style_name)`` of a paragraph's effective style, crash-safe.

    python-docx resolves a paragraph's effective style (a paragraph with no explicit
    ``pStyle`` reports the document's default style), so this is the real bucket key
    the per-role fold matches against. A reader that refuses to resolve the style
    yields ``(None, None)`` - the run then votes only toward the document body."""
    try:
        style = para.style
        sid = getattr(style, "style_id", None) if style is not None else None
        sname = getattr(style, "name", None) if style is not None else None
    except Exception:
        sid = sname = None
    return (sid, sname)


def _font_run_facts(doc):
    """Yield a :class:`_DocxRunFact` for every DIRECT ``w:r`` run in document order
    (``doc.paragraphs`` then ``para.runs``), tagged with the paragraph's style key.

    This is the font/size/color sampling pass: it deliberately does NOT widen to the
    hyperlink runs (matching v1 ``capture_fonts``, which read ``para.runs`` only). A
    single ordered generator keeps the ``Counter`` insertion order deterministic."""
    for para in doc.paragraphs:
        style_key = _para_style_key(para)
        for run in para.runs:
            yield _DocxRunFact(run, style_key, is_link=False)


def collect_font_run_facts(doc) -> list:
    """Materialize the direct-run facts pass ONCE so extract's consumers
    (:func:`capture_fonts` and :func:`capture_pseudo_headings`, which scans
    twice) share a single document walk instead of re-scanning the full tree.

    Pure recomputation removal: :class:`_DocxRunFact` eagerly captures all
    consumed state at construction and the document is not mutated between the
    passes, so iterating the materialized list yields the byte-identical votes
    (same items, same order) as re-running the generator."""
    return list(_font_run_facts(doc))


def capture_fonts(doc, roles: dict, theme: dict, *, facts=None) -> None:
    """Capture dominant direct run typography (font, size, color) into ``roles``
    (per role ``appearance``) and the document defaults (``theme['fonts']['body']``
    for font/size, ``theme['text']['body']`` for color), mutating both in place.

    Reads only the EXPLICIT run value per axis (``run.font.name`` / ``run.font.size``
    / ``run.font.color``); a run that inherits an axis from the style/theme
    contributes nothing to THAT axis (the three axes are sampled independently).
    python-docx resolves a paragraph's effective style (a paragraph with no explicit
    ``pStyle`` reports the document's default style), so runs are bucketed by their
    real style id/name. This now builds a docx ``RunFacts`` generator over the direct
    runs and delegates to the format-neutral
    :func:`~brandkit.common.typography.capture_appearance`; the default
    ``role_style_key`` reproduces the docx ``named_style`` OR-match byte-identically.

    ``facts`` optionally injects a pre-materialized
    :func:`collect_font_run_facts` list so one extract pass can share a single
    document walk across consumers; ``None`` (the default) scans as before.
    """
    _capture_appearance(
        facts if facts is not None else _font_run_facts(doc), roles, theme
    )


# ---------------------------------------------------------------------------
# Faked-heading-in-body-style detection (Cluster E2, DOCX-FIRST).
# ---------------------------------------------------------------------------
def _dominant_body_style_key(
    doc, facts=None
) -> Optional[tuple[Optional[str], Optional[str]]]:
    """The dominant paragraph style key across the document's body runs, or ``None``.

    The body/Normal style is the one the MOST body runs carry: a faked heading lives
    in a paragraph using THIS style (not a named heading style). We pick it as the
    most common style key over the same direct-run pass ``capture_fonts`` samples, so
    the pseudo-heading detector compares runs the engine truly treats as body. A
    document with no styled body run yields ``None`` (the detector then keys only on
    runs with no style of their own). ``facts`` optionally injects the
    pre-materialized pass (same items, same order)."""
    counter: Counter = Counter()
    for fact in facts if facts is not None else _font_run_facts(doc):
        if not (fact.text or "").strip():
            continue
        key = fact.style_key
        if key is not None and (key[0] or key[1]):
            counter[key] += 1
    if not counter:
        return None
    return counter.most_common(1)[0][0]


def capture_pseudo_headings(doc, roles: dict, theme: dict, *, facts=None) -> None:
    """Detect faked-heading-in-body-style candidates and store them additively under
    ``theme['pseudo_headings']`` (mutated in place), for the model to adjudicate via
    the ``comprehension.promote_appearance`` sink. DOCX-FIRST (Cluster E2).

    Runs AFTER :func:`capture_fonts` so the dominant body appearance
    (``theme.fonts.body.size_hp`` / ``theme.text.body.color``) is already captured;
    the detector compares each body-style run's OWN explicit size/color against that
    dominant (a PURE STATISTIC, nothing hardcoded to a template). Each candidate is
    stored as ``{ref, size_hp?, color?, evidence}`` - the run's CAPTURED outlier
    values plus a coarse, brand-text-free evidence string.

    Additive and deterministic: a template with no body-style size/color outlier (a
    uniform body) leaves ``theme['pseudo_headings']`` ABSENT, so the comprehend bundle
    stays byte-identical and generation is unchanged. The facts are READ-ONLY in the
    bundle - they never change generation on their own; only a model-adjudicated
    ``promote_appearance`` (re-validated shell-backed) moves any appearance.

    NOT wired into pptx/xlsx: a faked heading is a WordprocessingML body-paragraph
    phenomenon; pptx/xlsx carry no body-style runs, so they surface no candidates and
    stay byte-identical.

    ``facts`` optionally injects the pre-materialized
    :func:`collect_font_run_facts` list shared with :func:`capture_fonts`, so
    the two passes here (dominant-body vote + detector) re-iterate it instead of
    re-scanning the document; ``None`` (the default) scans as before.
    """
    body_style_key = _dominant_body_style_key(doc, facts)
    detected = _detect_pseudo_headings(
        facts if facts is not None else _font_run_facts(doc),
        theme,
        body_style_key=body_style_key,
    )
    if not detected:
        return
    theme["pseudo_headings"] = [
        {
            "ref": f.ref,
            **({"size_hp": f.size_hp} if f.size_hp is not None else {}),
            **({"color": f.color} if f.color is not None else {}),
            "evidence": f.evidence,
        }
        for f in detected
    ]


# ---------------------------------------------------------------------------
# theme.palette capture (model-free; the UNDERSTAND half of model-driven color)
# ---------------------------------------------------------------------------
def _iter_para_runs(para):
    """Yield every run in ``para``: its direct ``w:r`` runs AND the runs nested under
    its ``w:hyperlink`` elements.

    python-docx's ``para.runs`` exposes only the direct ``w:r`` children, so a link
    run (nested under ``w:hyperlink``) is otherwise invisible to capture. Newer
    python-docx (>= 1.x) surfaces ``para.hyperlinks[*].runs``; this widens the pass
    to include them, crash-safe (a degraded reader yields the direct runs only).
    """
    for run in para.runs:
        yield run
    try:
        for hyperlink in para.hyperlinks:
            for run in hyperlink.runs:
                yield run
    except Exception:
        # An older python-docx without ``para.hyperlinks`` simply contributes no
        # nested link runs - capture must never crash on a missing attribute.
        return


def _is_link_run(run) -> bool:
    """True if ``run`` is hyperlink text: nested under a ``w:hyperlink`` ancestor OR
    carrying a ``Hyperlink``/``FollowedHyperlink`` ``w:rStyle``.

    Wrapped fully crash-safe by the single caller; reads only structural OOXML
    (no brand literals). A run python-docx cannot introspect contributes nothing.
    """
    try:
        rpr = run._r.find(_W("rPr"))
        if rpr is not None:
            rstyle = rpr.find(_W("rStyle"))
            if rstyle is not None and rstyle.get(_W("val")) in _HYPERLINK_RSTYLES:
                return True
        node = run._r.getparent()
        while node is not None:
            if names.local_name(node.tag) == "hyperlink":
                return True
            node = node.getparent()
    except Exception:
        return False
    return False


def capture_palette(doc, roles: dict, theme: dict) -> None:
    """Capture the template's brand PALETTE into ``theme['palette']`` (mutated in
    place), additively and deterministically.

    The palette is a map keyed by a TEMPLATE-DERIVED id - a theme slot token
    (``accent1`` / ``text1`` / ...) for a theme color, or ``hex:RRGGBB`` for an
    observed off-theme run color. Each entry carries:

      - ``ref``: the byte-identical :func:`_color_obj` (``{kind:theme,theme}`` |
        ``{kind:hex,hex}``);
      - ``provenance``: a list of observed ``{where, detail}`` facts from the
        closed :data:`PALETTE_WHERE` vocabulary, sorted ``(where, detail)``;
      - ``frequency``: a COARSE bucket (``dominant`` | ``accent`` | ``rare``),
        never raw counts;
      - ``name`` / ``purpose`` / ``use_when``: ``null`` in this deterministic path
        (``comprehend`` is the only writer that fills them).

    Provenance is built ONLY from observed facts:
      (a) the theme-color slots the template actually carries (seed theme-keyed
          entries; existence, not a where-fact);
      (b) explicit ``w:color`` on runs (a SINGLE pass via :func:`_run_color`),
          INCLUDING a low-floor accent aggregation - a color on at least
          :data:`_MIN_ACCENT_RUNS` runs that is NOT the document-dominant body
          color is an ``accent`` entry (no dominance gate, accents are sparse);
      (c) the per-role ``appearance.color`` already captured (``role.appearance``);
      (d) link-run colors (runs under a ``w:hyperlink`` ancestor / ``Hyperlink``
          style), wrapped crash-safe, falling back to the theme ``hlink`` /
          ``folHlink`` slot when no explicit link color is observed (``link.color``).

    The hardcoded, template-INVARIANT ``theme.palette_roles`` map is NOT trusted as
    brand evidence; it is recorded only as a non-authoritative ``palette_role``
    where-entry on the slot it names. Deterministic and byte-identical on
    re-extract; a template with no observed color leaves an empty ``{}`` palette.

    This now builds a docx ``RunFacts`` generator over the WIDENED run pass (direct
    ``w:r`` runs AND the runs nested under ``w:hyperlink`` - source d) and delegates to
    :func:`~brandkit.common.typography.capture_palette_facts`. The bucket namespace is
    the docx WordprocessingML tokens VERBATIM (``_run_color``); nothing re-keys to
    ``THEME_SLOTS``, so the palette keys are byte-identical to v1.
    """
    _capture_palette_facts(_palette_run_facts(doc), roles, theme)


def _palette_run_facts(doc):
    """Yield a :class:`_DocxRunFact` for every run in the WIDENED palette pass
    (``doc.paragraphs`` then :func:`_iter_para_runs`, which adds the hyperlink runs),
    each tagged with :func:`_is_link_run`.

    The palette pass reads only ``color`` / ``is_link`` (font/size/style_key are
    irrelevant here), but the same ``_DocxRunFact`` view carries them. A single
    ordered generator keeps the dominance ``Counter`` insertion order deterministic,
    identical to v1 ``capture_palette``."""
    for para in doc.paragraphs:
        for run in _iter_para_runs(para):
            yield _DocxRunFact(run, None, is_link=_is_link_run(run))


# ---------------------------------------------------------------------------
# Paragraph GEOMETRY capture (Cluster D1, DOCX-ONLY).
# ---------------------------------------------------------------------------
# These readers are the docx-specific peers of ``_run_size_hp`` / ``_run_color`` for
# the geometry axis: they read EXPLICIT ``w:pPr`` properties off a paragraph and return
# ``None`` for any property the paragraph inherits, so the shared
# ``capture_paragraph_geometry`` engine sees an "inherit" vote it counts but never lets
# win. Everything is read off the live ``w:pPr`` element (no python-docx high-level
# wrapper) so the captured value is byte-faithful to the template. Borders/shading are
# copied as serialized element/attribute facts - never hand-constructed - so apply can
# re-emit the EXACT structure the template carried.
_GEOMETRY_BORDER_SIDES: tuple[str, ...] = ("top", "bottom", "left", "right")


def _ppr_of(para):
    """The paragraph's ``w:pPr`` element, or ``None`` when it carries none.

    Read crash-safe off the raw ``w:p`` (``para._p``); a paragraph with no direct
    properties has no ``w:pPr`` child, so every geometry axis votes ``None`` (inherit).
    """
    try:
        return para._p.find(_W("pPr"))
    except Exception:
        return None


def _twips_attr(el, attr: str) -> Optional[int]:
    """An integer twips attribute off ``el`` (e.g. ``w:spacing@w:before``), or ``None``.

    A missing element/attribute or a non-integer value (a malformed template) votes
    ``None`` (inherit) rather than crashing the extraction - capture is fail-soft."""
    if el is None:
        return None
    val = el.get(_W(attr))
    if val is None:
        return None
    try:
        return int(val)
    except (TypeError, ValueError):
        return None


def _spacing_facts(
    ppr,
) -> tuple[Optional[int], Optional[int], Optional[int], Optional[str]]:
    """``(before, after, line, line_rule)`` from ``w:pPr/w:spacing`` - each ``None``
    when that single attribute is absent (the attributes are INDEPENDENT axes)."""
    if ppr is None:
        return (None, None, None, None)
    spacing = ppr.find(_W("spacing"))
    if spacing is None:
        return (None, None, None, None)
    before = _twips_attr(spacing, "before")
    after = _twips_attr(spacing, "after")
    line = _twips_attr(spacing, "line")
    line_rule = spacing.get(_W("lineRule")) or None
    return (before, after, line, line_rule)


def _indent_facts(
    ppr,
) -> tuple[Optional[int], Optional[int], Optional[int], Optional[int]]:
    """``(left, right, first_line, hanging)`` twips from ``w:pPr/w:ind`` - each ``None``
    when that single attribute is absent (INDEPENDENT axes)."""
    if ppr is None:
        return (None, None, None, None)
    ind = ppr.find(_W("ind"))
    if ind is None:
        return (None, None, None, None)
    return (
        _twips_attr(ind, "left"),
        _twips_attr(ind, "right"),
        _twips_attr(ind, "firstLine"),
        _twips_attr(ind, "hanging"),
    )


def _border_facts(ppr) -> dict[str, str]:
    """``{side: <serialized w:top/.../w:right element>}`` for every ``w:pBdr`` side the
    paragraph carries directly, as an OPAQUE byte-copy (canonical lxml serialization).

    The element is copied VERBATIM (never hand-built), so the apply side re-emits the
    exact structure the template used; a side the paragraph does not carry is absent
    (it votes ``None`` on that side's dominance axis). Crash-safe: an unserializable
    node contributes nothing."""
    out: dict[str, str] = {}
    if ppr is None:
        return out
    pbdr = ppr.find(_W("pBdr"))
    if pbdr is None:
        return out
    for side in _GEOMETRY_BORDER_SIDES:
        el = pbdr.find(_W(side))
        if el is None:
            continue
        try:
            out[side] = etree.tostring(el, encoding="unicode")
        except Exception:
            continue
    return out


def _shading_fill_hex(ppr) -> Optional[str]:
    """The ``w:pPr/w:shd@w:fill`` shading hex, normalized, or ``None``.

    ``auto`` / a missing fill / a malformed hex contributes nothing (it is not a
    captured brand value, mirroring ``_run_color``'s drop of ``auto``)."""
    if ppr is None:
        return None
    shd = ppr.find(_W("shd"))
    if shd is None:
        return None
    fill = shd.get(_W("fill"))
    if not fill or fill.lower() == "auto":
        return None
    try:
        return colorutil.normalize_hex(fill)
    except (ValueError, AttributeError):
        return None


class _DocxParaGeometryFact:
    """A docx paragraph reduced to the structural
    :class:`~brandkit.common.typography.ParaGeometryFacts` view the shared geometry
    engine folds. Every axis is read EXPLICITLY off ``w:pPr`` (``None`` == inherit)."""

    __slots__ = (
        "style_key",
        "spacing_before_twips",
        "spacing_after_twips",
        "spacing_line_twips",
        "spacing_line_rule",
        "indent_left_twips",
        "indent_right_twips",
        "indent_first_line_twips",
        "indent_hanging_twips",
        "borders",
        "shading_fill_hex",
    )

    def __init__(self, para, style_key) -> None:
        self.style_key = style_key
        ppr = _ppr_of(para)
        before, after, line, line_rule = _spacing_facts(ppr)
        self.spacing_before_twips = before
        self.spacing_after_twips = after
        self.spacing_line_twips = line
        self.spacing_line_rule = line_rule
        left, right, first_line, hanging = _indent_facts(ppr)
        self.indent_left_twips = left
        self.indent_right_twips = right
        self.indent_first_line_twips = first_line
        self.indent_hanging_twips = hanging
        self.borders = _border_facts(ppr)
        self.shading_fill_hex = _shading_fill_hex(ppr)


def _geometry_para_facts(doc):
    """Yield a :class:`_DocxParaGeometryFact` per paragraph in document order, tagged
    with the paragraph's style key (the per-role fold key). A single ordered generator
    keeps the dominance ``Counter`` insertion order deterministic."""
    for para in doc.paragraphs:
        yield _DocxParaGeometryFact(para, _para_style_key(para))


def capture_geometry(doc, roles: dict, theme: dict) -> None:
    """Capture dominant paragraph GEOMETRY (spacing, indentation, paragraph borders,
    shading) into ``roles`` (per role ``appearance.geometry``) and the document default
    (``theme['geometry']['body']``), mutating both in place. DOCX-ONLY (Cluster D1).

    Reads only the EXPLICIT ``w:pPr`` value per property; a paragraph that inherits a
    property contributes nothing to THAT axis (every property is sampled
    independently). This builds a docx ``ParaGeometryFacts`` generator over the
    document's paragraphs and delegates to the format-neutral
    :func:`~brandkit.common.typography.capture_paragraph_geometry`; the default
    ``role_style_key`` reproduces the docx ``named_style`` OR-match used by
    :func:`capture_fonts`. Additive and deterministic: a template with no dominant
    direct geometry leaves ``geometry`` absent, so docx output stays byte-identical.

    NOT wired into pptx/xlsx: paragraph geometry is WordprocessingML ``w:pPr`` only.
    """
    _capture_paragraph_geometry(_geometry_para_facts(doc), roles, theme)


# ---------------------------------------------------------------------------
# TABLE conditional-format appearance capture (Cluster D2, DOCX-ONLY).
# ---------------------------------------------------------------------------
# These readers are the table peers of the paragraph-geometry readers: they read the
# EXPLICIT facts off a table's ``w:tblPr`` (the ``w:tblLook`` bitmask, the referenced
# ``w:tblStyle@w:val`` style id, and the four ``w:tblCellMar`` margins) and fold them
# under the SAME ``_dominant`` floor used for every other axis. A table that does not
# declare a fact votes ``None`` (inherit) on that axis, so a value is recorded only
# when an explicit value dominates the template's OWN tables. The engine NEVER reads or
# writes a ``w:tblStylePr`` (the band fills / first-last emphasis live in the shell's
# styles part); ``tblLook`` only ENABLES those shell-defined conditional formats and the
# style id only REFERENCES the shell's table style (membership-checked at verify).
# Additive and deterministic: a template with no dominant table fact leaves
# ``role.appearance.table`` absent, so docx output stays byte-identical.

# The six spec-fixed ``w:tblLook`` flag attributes, each a single bit in the captured
# bitmask. The OOXML bit positions (firstRow=0x0020, lastRow=0x0040, firstColumn=0x0080,
# lastColumn=0x0100, noHBand=0x0200, noVBand=0x0400) are the values Word writes when the
# legacy attribute form is collapsed to ``w:tblLook@w:val``; we reproduce them so a
# captured bitmask round-trips to the same toggles the template declared.
_TBLLOOK_FLAG_BITS: tuple[tuple[str, int], ...] = (
    ("firstRow", 0x0020),
    ("lastRow", 0x0040),
    ("firstColumn", 0x0080),
    ("lastColumn", 0x0100),
    ("noHBand", 0x0200),
    ("noVBand", 0x0400),
)

# The four ``w:tblCellMar`` margin sides, each its own dominance axis (twips).
_TABLE_CELL_MARGIN_SIDES: tuple[tuple[str, str], ...] = (
    ("top_twips", "top"),
    ("bottom_twips", "bottom"),
    ("left_twips", "left"),
    ("right_twips", "right"),
)


def _tbl_pr_of(table):
    """A table's ``w:tblPr`` element, or ``None`` when it carries none. Crash-safe."""
    try:
        return table._tbl.find(_W("tblPr"))
    except Exception:
        return None


def _table_tbllook(tblpr) -> Optional[int]:
    """The table's ``w:tblLook`` bitmask as a single integer, or ``None`` when absent.

    Reads either the modern ``w:tblLook@w:val`` hex/int attribute OR the legacy
    per-flag attribute form (``@w:firstRow`` ...), collapsing the legacy flags into the
    same integer bitmask Word uses for ``@w:val``. A table with no ``w:tblLook`` votes
    ``None`` (inherit) on this axis. Fail-soft: a malformed value contributes nothing.
    """
    if tblpr is None:
        return None
    look = tblpr.find(_W("tblLook"))
    if look is None:
        return None
    val = look.get(_W("val"))
    if val is not None:
        try:
            return int(val, 16)
        except (TypeError, ValueError):
            try:
                return int(val)
            except (TypeError, ValueError):
                return None
    # Legacy per-flag attribute form: collapse the set flags into the bitmask.
    bits = 0
    seen = False
    for attr, bit in _TBLLOOK_FLAG_BITS:
        raw = look.get(_W(attr))
        if raw is None:
            continue
        seen = True
        if raw in ("1", "true", "on"):
            bits |= bit
    return bits if seen else None


def _table_style_id(tblpr) -> Optional[str]:
    """The table's referenced ``w:tblStyle@w:val`` style id, or ``None`` when absent.

    This is a SYMBOLIC reference into the shell's styles part (membership-checked at
    verify); the engine never synthesizes the style's conditional formats. A table with
    no explicit style reference votes ``None`` on this axis."""
    if tblpr is None:
        return None
    style = tblpr.find(_W("tblStyle"))
    if style is None:
        return None
    return style.get(_W("val")) or None


def _table_cell_margins(tblpr) -> dict[str, int]:
    """``{side_twips: int}`` for every ``w:tblCellMar`` side the table declares directly.

    Each margin is read off ``w:tblPr/w:tblCellMar/w:{top,bottom,left,right}@w:w`` as
    twips; a side the table does not carry is absent (it votes ``None`` on that side's
    dominance axis). Fail-soft on a malformed/non-integer width."""
    out: dict[str, int] = {}
    if tblpr is None:
        return out
    cell_mar = tblpr.find(_W("tblCellMar"))
    if cell_mar is None:
        return out
    for field, side in _TABLE_CELL_MARGIN_SIDES:
        el = cell_mar.find(_W(side))
        twips = _twips_attr(el, "w")
        if twips is not None:
            out[field] = twips
    return out


class _DocxTableFact:
    """A docx table reduced to the structural table-appearance facts the D2 fold reads.

    Every axis is read EXPLICITLY off ``w:tblPr`` (``None`` / absent == inherit), so a
    table contributes a vote on an axis only when it declares that fact directly."""

    __slots__ = ("tbllook", "style_id", "cell_margins")

    def __init__(self, table) -> None:
        tblpr = _tbl_pr_of(table)
        self.tbllook = _table_tbllook(tblpr)
        self.style_id = _table_style_id(tblpr)
        self.cell_margins = _table_cell_margins(tblpr)


def _table_facts(doc) -> list:
    """A :class:`_DocxTableFact` per table in the document, in document order. A single
    ordered list keeps the dominance ``Counter`` insertion order deterministic."""
    return [_DocxTableFact(table) for table in doc.tables]


def _fold_table_appearance(facts: list) -> dict:
    """Fold the template's table facts into a captured ``table`` appearance dict, every
    axis (tblLook, style id, each cell margin) gated INDEPENDENTLY by the SAME
    :func:`~brandkit.common.typography.dominant` floor. Returns ``{}`` when nothing
    dominates (so the caller writes no ``table`` key, zero-branch on no-capture)."""
    if not facts:
        return {}
    out: dict = {}

    look_counter: Counter = Counter()
    for f in facts:
        look_counter[f.tbllook] += 1
    look_dom = _dominant(look_counter)
    if look_dom is not None:
        out["tblLook"] = look_dom[0]

    style_counter: Counter = Counter()
    for f in facts:
        style_counter[f.style_id] += 1
    style_dom = _dominant(style_counter)
    if style_dom is not None:
        out["style_id"] = style_dom[0]

    cell_margins: dict[str, int] = {}
    for field, _side in _TABLE_CELL_MARGIN_SIDES:
        counter: Counter = Counter()
        for f in facts:
            counter[f.cell_margins.get(field)] += 1
        dom = _dominant(counter)
        if dom is not None:
            cell_margins[field] = dom[0]
    if cell_margins:
        out["cell_margins"] = cell_margins

    return out


def capture_table_appearance(doc, roles: dict, theme: dict) -> None:
    """Capture the dominant TABLE conditional-format facts (the ``w:tblLook`` bitmask,
    the referenced table-style id, and the ``w:tblCellMar`` cell margins) into the
    ``table.*`` roles (per role ``appearance.table``) and the document default
    (``theme['table']['body']``), mutating both in place. DOCX-ONLY (Cluster D2).

    Reads only the EXPLICIT ``w:tblPr`` facts the template's OWN tables declare; a table
    that inherits a fact contributes nothing to THAT axis (every axis is sampled
    independently under the SAME ``MIN_RUNS`` + ``MIN_DOMINANCE`` floor used for
    font/size/color/geometry). The band fills / first-last emphasis live in the shell's
    styles part (``w:tblStylePr``): this capture NEVER reads or copies them - the bitmask
    only records WHICH of the shell style's own conditional formats the template enables,
    and the style id only RECORDS the symbolic reference (membership-checked at verify).
    Additive and deterministic: a template with no dominant table fact leaves
    ``role.appearance.table`` absent, so docx output stays byte-identical.

    NOT wired into pptx/xlsx: ``w:tblLook`` / ``w:tblStyle`` / ``w:tblCellMar`` are
    WordprocessingML table constructs with no pptx/xlsx peer.
    """
    facts = _table_facts(doc)
    if not facts:
        return
    body = _fold_table_appearance(facts)
    if body:
        theme.setdefault("table", {})["body"] = body
    # The dominant facts are document-wide (a docx has ONE table style convention the
    # captured roles share); the per-role table appearance is the same captured body,
    # recorded on every ``table.*`` role so the resolver's role-specific axis carries it
    # verbatim (the resolver merge then prefers role over body, identical to geometry).
    if not body:
        return
    for rid, entry in roles.items():
        if rid == "_index" or not isinstance(entry, dict):
            continue
        # Only the ``table`` family carries table appearance (``table`` / ``table.*``),
        # mirroring the ``rid.startswith("list.")`` family test used elsewhere here.
        if rid != "table" and not rid.startswith("table."):
            continue
        entry.setdefault("appearance", {})["table"] = body
