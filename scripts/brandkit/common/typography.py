# SPDX-License-Identifier: MIT
"""Format-neutral brand typography helpers (capture-side, model-free).

This is the shared engine the per-format adapters (docx/pptx/xlsx) delegate to so
the dominant-value scoring, the palette-entry shape, and the provenance vocabulary
have exactly ONE writer across all three kinds. It carries NO heavy OOXML stack:
no ``docx`` / ``pptx`` / ``openpyxl`` import (it leans only on the equally light
pure-Python sibling :mod:`brandkit.common.color` for :data:`THEME_SLOTS`), so every
extractor can import it cheaply, exactly like :mod:`brandkit.common.text`.

What lives here (moved VERBATIM out of ``formats/docx/typography.py`` so behavior is
byte-identical):

  - :func:`_dominant` - the dominance gate over a :class:`collections.Counter`;
  - :func:`_color_obj` / :func:`_color_obj_to_bucket` - the captured-color bucket
    <-> stored ``appearance`` object round-trip;
  - :func:`_palette_key` - the template-derived palette key for a color bucket;
  - :func:`_add_provenance` - the de-duplicated, sorted provenance writer;
  - :func:`_palette_entry` - the SINGLE get-or-create writer of the palette entry
    shape (also the writer ``color.seed_theme_palette`` now calls);
  - the capture constants :data:`MIN_RUNS` / :data:`MIN_DOMINANCE` /
    :data:`MIN_ACCENT_RUNS`, the closed :data:`PALETTE_WHERE` vocabulary, and the
    re-exported :data:`THEME_SLOTS`.

The brand guarantee is preserved: every value these helpers record is a FACT the
adapter observed in the shell; nothing here NAMES a color/font/style or injects a
literal. The three model-only fields (``name`` / ``purpose`` / ``use_when``) are
left ``null`` here - ``comprehend`` is the only writer that fills them.
"""

from __future__ import annotations

from collections import Counter
from typing import (
    Any,
    Callable,
    Iterable,
    NamedTuple,
    Optional,
    Protocol,
    runtime_checkable,
)

from brandkit.common import color as colorutil
from brandkit.ooxml import names as ooxml_names
from brandkit.ooxml import pack as ooxml_pack

# The 12 canonical theme slots a palette theme-key may name (single registry).
THEME_SLOTS: frozenset[str] = frozenset(colorutil.THEME_SLOTS)

_A = ooxml_names.make_qn("a")


def theme_font_scheme_latin(
    template_path, theme_part: str
) -> tuple[Optional[str], Optional[str]]:
    """Read the major/minor latin typefaces from a package's theme part.

    The format-uniform sibling of the docx extractor's theme-font read:
    ``(major, minor)`` from ``a:fontScheme/a:majorFont|a:minorFont/a:latin
    @typeface``, each ``None`` when the part, element, or attribute is missing,
    or when Office wrote its ``typeface=""`` "no face here" sentinel. A missing
    theme part is the only swallowed error (KeyError); any other parse error
    propagates - a corrupt theme must not silently blank the captured fonts.
    """
    try:
        xml = ooxml_pack.read_part(template_path, theme_part)
    except KeyError:
        return None, None
    root = ooxml_pack.parse_xml_bytes(xml)
    scheme = root.find(f".//{_A('fontScheme')}")
    if scheme is None:
        return None, None

    def latin_of(font_tag: str) -> Optional[str]:
        font = scheme.find(_A(font_tag))
        if font is None:
            return None
        latin = font.find(_A("latin"))
        if latin is None:
            return None
        return latin.get("typeface") or None

    return latin_of("majorFont"), latin_of("minorFont")


# A capture is only trusted when it is a clear convention, not noise.
MIN_RUNS = 3  # need at least this many explicit values to call a winner
MIN_DOMINANCE = 0.6  # the winner must cover >= 60% of those values

# An accent color is SPARSE by design (a few runs of brand red on a body of black
# text), so the palette accent bucket uses ONLY a low count floor and NOT the
# _dominant MIN_DOMINANCE gate: a color seen on at least this many runs but not the
# document-dominant body color is recorded as an "accent" entry.
MIN_ACCENT_RUNS = 3

# The closed ``where`` vocabulary for a palette entry's provenance (LOCKED). Every
# provenance fact is one of these four observed sources, plus the engine-minted
# ``palette.alias`` (Cluster E1); nothing else may be recorded. ``palette_role`` is
# the only NON-authoritative observed source (the hardcoded, template-invariant
# ``theme.palette_roles`` map), recorded for context but never trusted as brand
# evidence. ``palette.alias`` is stamped by the merge-time alias mint on a token the
# model NAMED for a captured entry; the alias ref is a byte-copy of that entry's ref,
# never an observed capture (kept in sync with ``schema.PALETTE_WHERE``).
PALETTE_WHERE: frozenset[str] = frozenset(
    {"palette_role", "role.appearance", "run.color", "link.color", "palette.alias"}
)


def _dominant(counter: Counter) -> Optional[tuple[Any, float]]:
    """Return ``(value, dominance)`` for the most common EXPLICIT value when it is a
    clear convention over ALL sampled runs, else ``None``.

    A ``None`` key counts runs that carry NO explicit value on this axis (they inherit
    from the style/theme). Those "inherit" votes count toward the denominator but can
    never WIN: a value is a convention only when it dominates EVERY run, not just the
    explicit minority. This is what stops a 2%-of-runs accent color (the 98% inherit
    their color) from being mistaken for the document's body color, while a body that
    really is explicitly Roboto/16pt on (almost) every run still captures."""
    total = sum(counter.values())
    if total < MIN_RUNS:
        return None
    candidates = [(value, n) for value, n in counter.items() if value is not None]
    if not candidates:
        return None
    value, n = max(candidates, key=lambda item: item[1])
    ratio = n / total
    if ratio < MIN_DOMINANCE:
        return None
    return value, ratio


# Public alias: the same dominance gate (``MIN_RUNS`` + ``MIN_DOMINANCE``) the
# typography/geometry axes use, exposed for the docx TABLE-appearance capture (Cluster
# D2) which folds the template's own tables' tblLook / style / cell-margin facts under
# the identical floor without re-implementing the gate.
dominant = _dominant


def _color_obj(bucket: tuple[str, ...]) -> dict:
    """Turn a captured color bucket key back into its stored ``appearance`` object."""
    if bucket[0] == "hex":
        return {"kind": "hex", "hex": bucket[1]}
    return {"kind": "theme", "theme": bucket[1]}


def _color_obj_to_bucket(color: dict) -> Optional[tuple[str, ...]]:
    """Invert :func:`_color_obj`: a stored color object -> its bucket key, or None.

    Used to fold an already-captured ``role.appearance.color`` (which is a
    ``_color_obj``) back into a palette bucket without re-reading the run.
    """
    kind = color.get("kind")
    if kind == "hex" and color.get("hex"):
        return ("hex", str(color["hex"]))
    if kind == "theme" and color.get("theme"):
        return ("theme", str(color["theme"]))
    return None


def _palette_key(bucket: tuple[str, ...]) -> str:
    """The TEMPLATE-DERIVED palette key for a captured color bucket.

    A theme bucket keys by its WML theme token (``accent1`` / ``text1`` / ...);
    an off-theme RGB bucket keys by ``hex:RRGGBB``. The key is the stable id the
    comprehension annotates and the resolver/QA look up - never a brand name.
    """
    if bucket[0] == "hex":
        return f"hex:{bucket[1]}"
    return bucket[1]


def _add_provenance(entry: dict, where: str, detail: str) -> None:
    """Record one observed ``{where, detail}`` provenance fact on a palette entry,
    de-duplicated and kept sorted by ``(where, detail)`` (deterministic).

    ``where`` must be in the closed :data:`PALETTE_WHERE` vocabulary; an unknown
    ``where`` is dropped (capture only records observed facts in the frozen set).
    """
    if where not in PALETTE_WHERE:
        return
    provenance = entry.setdefault("provenance", [])
    fact = {"where": where, "detail": detail}
    if fact in provenance:
        return
    provenance.append(fact)
    provenance.sort(key=lambda p: (p["where"], p["detail"]))


def _palette_entry(palette: dict, bucket: tuple[str, ...]) -> dict:
    """Get-or-create the palette entry for a color bucket, keyed template-derived.

    A new entry carries the byte-identical :func:`_color_obj` ref, an empty
    provenance, a placeholder frequency (set by the caller), and the three
    model-only fields (``name`` / ``purpose`` / ``use_when``) explicitly ``null``
    in the deterministic path - ``comprehend`` is the only writer that fills them.
    """
    key = _palette_key(bucket)
    entry = palette.get(key)
    if entry is None:
        entry = {
            "ref": _color_obj(bucket),
            "provenance": [],
            "frequency": "rare",
            "name": None,
            "purpose": None,
            "use_when": None,
        }
        palette[key] = entry
    return entry


# ---------------------------------------------------------------------------
# Format-neutral capture engine (the shape ALL three kinds fill)
# ---------------------------------------------------------------------------
# A run, reduced to the FACTS the capture engine needs, duck-typed across the three
# kinds. The per-format adapter (docx/pptx/xlsx ``iter_run_facts``) yields one of
# these per styled, non-empty run, reading the format's own run object. The engine
# below never touches an OOXML object - only this structural view - so the
# dominant-value scoring, the body/role fold, and the palette build have exactly ONE
# writer regardless of which backend produced the runs.
@runtime_checkable
class RunFacts(Protocol):
    """The structural (duck-typed) view of one captured run the engine consumes.

    NOT a base class to subclass: any object exposing these attributes satisfies it.
    ``color`` is the hashable bucket key the palette/appearance helpers round-trip
    (``('hex', RRGGBB)`` | ``('theme', <token>)``); the token namespace is the
    ADAPTER'S choice (docx emits its WordprocessingML tokens verbatim; pptx/xlsx
    normalize to :data:`THEME_SLOTS`) - the engine is namespace-agnostic and stores
    whatever bucket it is handed."""

    # (style_id, style_name) | None - the run's owning style identity for the per-role
    # fold; ``None`` means the run votes only toward the document body, never a role.
    style_key: Optional[tuple[Optional[str], Optional[str]]]
    text: str
    font_name: Optional[str]
    size_hp: Optional[int]  # half-points
    color: Optional[tuple[str, ...]]  # ('hex', RRGGBB) | ('theme', <token>)
    is_link: bool


def _role_style_key_default(entry: dict) -> Optional[tuple[str, ...]]:
    """The docx role->style key: a ``named_style`` resolver's ``(style_id, style_name)``.

    A role whose resolver is not a named style (pptx placeholder / xlsx named range /
    cell style / number format) returns ``None`` and is therefore never folded by run
    style key - its per-role appearance simply stays empty (the body fallback still
    reaches it via the resolver at apply time). This is the byte-identical peer of the
    ``ResolverType.NAMED_STYLE`` guard the docx ``capture_fonts`` used inline."""
    resolver = entry.get("resolver") or {}
    if resolver.get("type") != "named_style":
        return None
    return (resolver.get("style_id"), resolver.get("style_name"))


def capture_appearance(
    run_facts: Iterable[RunFacts],
    roles: dict,
    theme: dict,
    *,
    role_style_key: Callable[[dict], Optional[tuple]] = _role_style_key_default,
) -> None:
    """Capture dominant direct run typography (font, size, color) into ``roles`` (per
    role ``appearance``) and the document defaults (``theme['fonts']['body']`` for
    font/size, ``theme['text']['body']`` for color), mutating both in place.

    Format-neutral reproduction of the docx ``capture_fonts`` body, byte-identical:
    a SINGLE ordered pass over ``run_facts`` builds the per-style and overall
    ``Counter``s (insertion order preserved so ``_dominant``'s ``max`` tie-break stays
    deterministic), the three axes are sampled INDEPENDENTLY (a run that inherits an
    axis votes ``None`` for it, counting in the denominator but never winning), then
    the body winners and the per-role winners are recorded only when they clear the
    :func:`_dominant` gate. ``role_style_key`` maps a role entry to its style key (or
    ``None`` to skip it); a role's per-style counter unions every observed run-style
    key matching the role key by ``style_id`` OR ``style_name`` (the docx OR-match)."""
    per_style_font: dict[tuple, Counter] = {}
    per_style_size: dict[tuple, Counter] = {}
    per_style_color: dict[tuple, Counter] = {}
    overall_font: Counter = Counter()
    overall_size: Counter = Counter()
    overall_color: Counter = Counter()

    for fact in run_facts:
        if not (fact.text or "").strip():
            continue
        font = fact.font_name or None
        size_hp = fact.size_hp
        color = fact.color
        overall_font[font] += 1
        overall_size[size_hp] += 1
        overall_color[color] += 1
        key = fact.style_key
        if key is not None and (key[0] or key[1]):
            per_style_font.setdefault(key, Counter())[font] += 1
            per_style_size.setdefault(key, Counter())[size_hp] += 1
            per_style_color.setdefault(key, Counter())[color] += 1

    body_font = _dominant(overall_font)
    body_size = _dominant(overall_size)
    if body_font is not None or body_size is not None:
        fonts = theme.setdefault("fonts", {})
        body = fonts.setdefault("body", {})
        if body_font is not None:
            body["latin"] = body_font[0]
            body["confidence"] = round(body_font[1], 3)
        if body_size is not None:
            body["size_hp"] = int(body_size[0])
            body["size_confidence"] = round(body_size[1], 3)
    body_color = _dominant(overall_color)
    if body_color is not None:
        text = theme.setdefault("text", {}).setdefault("body", {})
        text["color"] = _color_obj(body_color[0])
        text["color_confidence"] = round(body_color[1], 3)

    for rid, entry in roles.items():
        if rid == "_index" or not isinstance(entry, dict):
            continue
        key = role_style_key(entry)
        if key is None:
            continue
        sid, sname = key[0], key[1]

        def _role_counter(per_style: dict, sid=sid, sname=sname) -> Counter:
            counter: Counter = Counter()
            for (k_sid, k_sname), c in per_style.items():
                if (sid and k_sid == sid) or (sname and k_sname == sname):
                    counter.update(c)
            return counter

        dom_font = _dominant(_role_counter(per_style_font))
        dom_size = _dominant(_role_counter(per_style_size))
        dom_color = _dominant(_role_counter(per_style_color))
        if dom_font is None and dom_size is None and dom_color is None:
            continue
        appearance = entry.setdefault("appearance", {})
        if dom_font is not None:
            appearance["font"] = {"latin": dom_font[0]}
            appearance["confidence"] = round(dom_font[1], 3)
        if dom_size is not None:
            appearance["size_hp"] = int(dom_size[0])
            appearance["size_confidence"] = round(dom_size[1], 3)
        if dom_color is not None:
            appearance["color"] = _color_obj(dom_color[0])
            appearance["color_confidence"] = round(dom_color[1], 3)


# ---------------------------------------------------------------------------
# Faked-heading-in-body-style detection (Cluster E2).
# ---------------------------------------------------------------------------
# Some templates FAKE a heading: a line that LOOKS like a heading (visibly larger
# and/or in a brand color) is authored with the BODY paragraph style - no heading
# role, no named heading style - so the deterministic engine treats it as body and
# the brand heading look is lost. E2 SURFACES each such body-style run as a
# ``pseudo_heading`` FACT the model adjudicates: it carries a stable structural ref,
# the run's OWN captured outlier size/color (a FACT about the template, not a brand
# value the engine synthesizes), and coarse, brand-text-free evidence.
#
# The detector is a PURE STATISTIC vs the captured DOMINANT body appearance (the
# SAME ``_dominant`` winners ``capture_appearance`` records into ``theme.fonts.body``
# / ``theme.text.body``): a body-style run is a candidate when its EXPLICIT size is a
# clear OUTLIER (markedly larger or smaller than the body size) OR its EXPLICIT color
# is OFF-BODY (a different bucket than the dominant body color). Nothing is hardcoded
# to any template's points/hex - the comparison is against the template's OWN observed
# dominant, so the test is universal. The detector only SURFACES candidates; it never
# changes generated output by itself (the model must adjudicate a promotion before any
# appearance moves, and the promoted value is re-validated shell-backed at QA).

# The size-outlier ratio bounds: a body-style run is a size outlier when its explicit
# size is at least 1.5x the dominant body size (a markedly larger fake heading) or at
# most 0.67x it (a markedly smaller one). These are FIXED, template-INVARIANT ratios
# applied to the template's OWN dominant body size (never an absolute point bound), so
# the test stays universal - a 12pt body flags a >=18pt run, a 16pt body flags a
# >=24pt run, each relative to its own template.
_PSEUDO_HEADING_SIZE_RATIO_HI = 1.5
_PSEUDO_HEADING_SIZE_RATIO_LO = 0.67


class PseudoHeadingFact(NamedTuple):
    """One detected faked-heading-in-body-style candidate (Cluster E2).

    Emitted by the pure-deterministic outlier detector
    (:func:`detect_pseudo_headings`) and surfaced as a ``pseudo_heading`` fact in the
    ``comprehend_input_bundle``. The model adjudicates a promotion by NAMING the
    ``ref`` + a declared heading role; the engine then copies the CAPTURED
    ``size_hp`` / ``color`` onto that role (never a synthesized value).

      - ``ref`` is a stable STRUCTURAL id from the run enumeration the detector walks
        (``body_run_{n}`` over the body-style runs in document order); it is an
        internal id, NOT a surfaced inventory ref like an anchor/field/region.
      - ``size_hp`` / ``color`` are the run's OWN explicit captured outlier values (a
        half-point int / a stored color object), each ``None`` when that axis is not
        the outlier (the axes are independent).
      - ``evidence`` is coarse, brand-text-free (e.g. ``"size 4400hp vs dominant body
        2400hp"`` / ``"off-body color theme:accent1 vs body theme:text1"``).
    """

    ref: str
    size_hp: Optional[int]
    color: Optional[dict]
    evidence: str


def _is_size_outlier(size_hp: Optional[int], body_size_hp: Optional[int]) -> bool:
    """True when ``size_hp`` is a clear size outlier vs the dominant ``body_size_hp``.

    Pure ratio test against the template's OWN dominant body size (no absolute point
    bound): an explicit run size at least :data:`_PSEUDO_HEADING_SIZE_RATIO_HI` x the
    body size (a markedly larger fake heading) or at most
    :data:`_PSEUDO_HEADING_SIZE_RATIO_LO` x it. ``None`` on either side (an inherited
    run size, or a template with no captured body size) is never an outlier.
    """
    if not size_hp or not body_size_hp:
        return False
    ratio = size_hp / body_size_hp
    return (
        ratio >= _PSEUDO_HEADING_SIZE_RATIO_HI or ratio <= _PSEUDO_HEADING_SIZE_RATIO_LO
    )


def _is_color_outlier(
    color: Optional[tuple[str, ...]], body_bucket: Optional[tuple[str, ...]]
) -> bool:
    """True when the run's explicit ``color`` bucket is OFF the dominant body color.

    A pure equality test against the template's OWN dominant body color bucket: an
    EXPLICIT run color that differs from the body bucket (a different theme token, or
    an off-theme hex where the body is a theme color, or vice versa) is off-body.
    ``None`` on either side (an inherited run color, or a template with no captured
    body color) is never an outlier - the run simply carries the body color.
    """
    if color is None or body_bucket is None:
        return False
    return color != body_bucket


def detect_pseudo_headings(
    run_facts: Iterable[RunFacts],
    theme: dict,
    *,
    body_style_key: Optional[tuple[Optional[str], Optional[str]]] = None,
) -> list[PseudoHeadingFact]:
    """Detect faked-heading-in-body-style candidates (Cluster E2), pure-deterministic.

    Reads the ALREADY-CAPTURED per-run facts (``size_hp`` half-points / ``color``
    bucket) the SAME ``run_facts`` :func:`capture_appearance` consumed, and compares
    each BODY-STYLE run against the captured DOMINANT body appearance (the
    ``theme.fonts.body.size_hp`` / ``theme.text.body.color`` winners). A body-style
    run whose EXPLICIT size is a clear outlier (:func:`_is_size_outlier`) OR whose
    EXPLICIT color is off-body (:func:`_is_color_outlier`) is surfaced as a
    :class:`PseudoHeadingFact`. The size/color carried on the fact are the run's OWN
    captured outlier values (facts about the template, never synthesized).

    A run is "body-style" when it has NO style key of its own (``style_key`` is
    ``None`` / empty) - it votes only toward the document body - OR when its style
    key matches ``body_style_key`` (the explicit body/Normal style). This keeps the
    detector to runs the engine actually treats as body (a run under a named heading
    style already carries its heading role and is never a "faked" heading).

    Returns ``[]`` when nothing is an outlier (a uniform body, or a template with no
    captured body dominant to compare against) - the caller then writes NO
    ``pseudo_headings`` key, so the bundle stays byte-identical. Order-preserving and
    deterministic (one ordered pass over ``run_facts``); the ``ref`` enumerates the
    body-style runs in document order so it is stable across re-extractions.
    """
    body_size_hp = ((theme.get("fonts") or {}).get("body") or {}).get("size_hp")
    body_color_obj = ((theme.get("text") or {}).get("body") or {}).get("color")
    body_bucket = (
        _color_obj_to_bucket(body_color_obj)
        if isinstance(body_color_obj, dict)
        else None
    )
    # Nothing to compare against: a template with no captured body size AND no body
    # color has no dominant to call an outlier (the no-capture path).
    if not body_size_hp and body_bucket is None:
        return []

    facts: list[PseudoHeadingFact] = []
    index = 0
    for fact in run_facts:
        if not (fact.text or "").strip():
            continue
        key = fact.style_key
        is_body = key is None or not (key[0] or key[1]) or key == body_style_key
        if not is_body:
            continue
        ref = f"body_run_{index}"
        index += 1
        size_outlier = _is_size_outlier(fact.size_hp, body_size_hp)
        color_outlier = _is_color_outlier(fact.color, body_bucket)
        if not size_outlier and not color_outlier:
            continue
        size_hp = int(fact.size_hp) if size_outlier else None
        color = _color_obj(fact.color) if color_outlier and fact.color else None
        facts.append(
            PseudoHeadingFact(
                ref=ref,
                size_hp=size_hp,
                color=color,
                evidence=_pseudo_heading_evidence(
                    size_hp, color, body_size_hp, body_bucket
                ),
            )
        )
    return facts


def _pseudo_heading_evidence(
    size_hp: Optional[int],
    color: Optional[dict],
    body_size_hp: Optional[int],
    body_bucket: Optional[tuple[str, ...]],
) -> str:
    """A coarse, brand-text-free evidence string for a detected pseudo-heading.

    Names only the captured NUMERIC size (half-points) and the structural COLOR
    bucket (``theme:<token>`` / ``hex:RRGGBB``) vs the dominant body - never any
    template/brand TEXT. The two halves are joined when both axes are outliers.
    """
    parts: list[str] = []
    if size_hp is not None and body_size_hp:
        parts.append(f"size {size_hp}hp vs dominant body {int(body_size_hp)}hp")
    if color is not None:
        run_key = _color_obj_to_bucket(color)
        run_label = _palette_key(run_key) if run_key is not None else "?"
        body_label = _palette_key(body_bucket) if body_bucket is not None else "?"
        parts.append(f"off-body color {run_label} vs body {body_label}")
    return "; ".join(parts)


# ---------------------------------------------------------------------------
# Paragraph GEOMETRY capture (Cluster D1, DOCX-ONLY).
# ---------------------------------------------------------------------------
# Geometry is the FOURTH appearance axis (after font / size / color), but unlike the
# three typographic axes it is NOT a single value per role: a paragraph carries many
# INDEPENDENT geometry properties (spacing before/after/line, four indents, four
# border sides, a shading fill). Each property is its OWN axis under the SAME
# :func:`_dominant` floor (a paragraph with no explicit value on a property votes
# ``None``, counting in the denominator but never winning), so a role records only the
# properties that DOMINATE its own paragraphs. Nothing is hardcoded to any template's
# twips - the threshold is purely MIN_RUNS + MIN_DOMINANCE on the template's own
# ``w:pPr`` values, exactly like font/size/color. This engine is format-neutral; the
# docx adapter (``formats/docx/typography.py``) is the ONLY producer of
# :class:`ParaGeometryFacts` (it reads ``w:pPr``). pptx/xlsx never call it (D1 is
# WordprocessingML pPr only); a profile that carries no captured geometry leaves
# ``role.appearance.geometry`` absent, so the no-geometry path takes ZERO new branches
# and docx output stays byte-identical.

# The scalar geometry properties, grouped as they are stored under
# ``appearance.geometry`` (a dotted "<group>.<field>" key into the right sub-dict).
# Borders and shading are NOT here (they are element/hex shapes, captured separately).
_GEOMETRY_SCALAR_FIELDS: tuple[tuple[str, str], ...] = (
    ("spacing", "before_twips"),
    ("spacing", "after_twips"),
    ("spacing", "line_twips"),
    ("spacing", "line_rule"),
    ("indentation", "left_twips"),
    ("indentation", "right_twips"),
    ("indentation", "first_line_twips"),
    ("indentation", "hanging_twips"),
)

# The four WordprocessingML paragraph border sides a ``w:pBdr`` can carry, captured
# each as its own dominant axis (the serialized element, an opaque copy).
_GEOMETRY_BORDER_SIDES: tuple[str, ...] = ("top", "bottom", "left", "right")


@runtime_checkable
class ParaGeometryFacts(Protocol):
    """The structural (duck-typed) view of ONE paragraph's geometry the engine folds.

    The docx adapter yields one per paragraph, reading ``w:pPr`` explicitly: a field
    is ``None`` when the paragraph does NOT carry that property directly (it inherits
    from the style), so it votes "inherit". ``style_key`` is the paragraph's owning
    style identity (the per-role fold key, same shape as :class:`RunFacts`).

      - the scalar fields mirror :data:`_GEOMETRY_SCALAR_FIELDS` (twips as ``int``,
        ``line_rule`` as a small ``str`` token like ``'auto'``/``'exact'``);
      - ``borders`` is a ``{side: <opaque serialized element str>}`` map of the sides
        the paragraph carries directly (a byte-copy the apply side re-emits verbatim);
      - ``shading_fill_hex`` is the ``w:shd@w:fill`` normalized hex, or ``None``.
    """

    style_key: Optional[tuple[Optional[str], Optional[str]]]
    spacing_before_twips: Optional[int]
    spacing_after_twips: Optional[int]
    spacing_line_twips: Optional[int]
    spacing_line_rule: Optional[str]
    indent_left_twips: Optional[int]
    indent_right_twips: Optional[int]
    indent_first_line_twips: Optional[int]
    indent_hanging_twips: Optional[int]
    borders: dict[str, str]
    shading_fill_hex: Optional[str]


# The ParaGeometryFacts attribute that feeds each scalar (group, field) axis.
_GEOMETRY_SCALAR_ATTR: dict[tuple[str, str], str] = {
    ("spacing", "before_twips"): "spacing_before_twips",
    ("spacing", "after_twips"): "spacing_after_twips",
    ("spacing", "line_twips"): "spacing_line_twips",
    ("spacing", "line_rule"): "spacing_line_rule",
    ("indentation", "left_twips"): "indent_left_twips",
    ("indentation", "right_twips"): "indent_right_twips",
    ("indentation", "first_line_twips"): "indent_first_line_twips",
    ("indentation", "hanging_twips"): "indent_hanging_twips",
}


def _fold_geometry(geometry_facts: list, key_match) -> dict:
    """Fold a list of :class:`ParaGeometryFacts` (already filtered to a role, or ALL of
    them for the body default) into a captured ``geometry`` dict, every property gated
    by the SAME :func:`_dominant` floor INDEPENDENTLY.

    ``key_match(fact) -> bool`` selects which facts vote (a role unions every paragraph
    whose style matches; the body default takes them all). Returns ``{}`` when nothing
    dominates, so the caller writes no ``geometry`` key (zero-branch on no-capture)."""
    facts = [f for f in geometry_facts if key_match(f)]
    if not facts:
        return {}

    out: dict[str, Any] = {}
    confidence: dict[str, float] = {}

    # (a) scalar spacing / indentation axes - each its own dominance Counter.
    for group, field in _GEOMETRY_SCALAR_FIELDS:
        attr = _GEOMETRY_SCALAR_ATTR[(group, field)]
        counter: Counter = Counter()
        for f in facts:
            counter[getattr(f, attr)] += 1
        dom = _dominant(counter)
        if dom is None:
            continue
        out.setdefault(group, {})[field] = dom[0]
        confidence[f"{group}.{field}"] = round(dom[1], 3)

    # (b) each border side is its own dominance axis over the serialized element; a
    # paragraph with no ``w:pBdr`` side votes ``None`` (inherit) on that side.
    borders: dict[str, str] = {}
    for side in _GEOMETRY_BORDER_SIDES:
        counter = Counter()
        for f in facts:
            counter[f.borders.get(side)] += 1
        dom = _dominant(counter)
        if dom is None:
            continue
        borders[side] = dom[0]
        confidence[f"borders.{side}"] = round(dom[1], 3)
    if borders:
        out["borders"] = borders

    # (c) shading fill hex - one dominance axis.
    counter = Counter()
    for f in facts:
        counter[f.shading_fill_hex] += 1
    dom = _dominant(counter)
    if dom is not None:
        out["shading"] = {"fill_hex": dom[0]}
        confidence["shading.fill_hex"] = round(dom[1], 3)

    if not out:
        return {}
    out["confidence"] = confidence
    return out


def capture_paragraph_geometry(
    geometry_facts: Iterable,
    roles: dict,
    theme: dict,
    *,
    role_style_key: Callable[[dict], Optional[tuple]] = _role_style_key_default,
) -> None:
    """Capture dominant paragraph GEOMETRY into ``roles`` (per role
    ``appearance.geometry``) and the document default (``theme['geometry']['body']``),
    mutating both in place. DOCX-ONLY (the only caller is the docx extractor).

    Each geometry property (spacing before/after/line + line rule, the four indents,
    the four border sides, shading fill) is an INDEPENDENT axis under the SAME
    :func:`_dominant` floor used for font/size/color: a property is recorded for a role
    only when an explicit value DOMINATES that role's own paragraphs (and the body
    default only when it dominates ALL paragraphs). A property with no dominance writes
    NO key, so a template with no convention leaves ``geometry`` absent and the
    no-geometry path is byte-identical. Confidence (the dominance ratio per property)
    is stored under ``geometry['confidence']`` keyed by ``"<group>.<field>"``.

    ``geometry_facts`` is materialized once into a list (the per-role fold replays it),
    so the caller may pass a single ordered generator over the document's paragraphs.
    """
    facts = list(geometry_facts)
    if not facts:
        return

    body = _fold_geometry(facts, lambda f: True)
    if body:
        theme.setdefault("geometry", {})["body"] = body

    for rid, entry in roles.items():
        if rid == "_index" or not isinstance(entry, dict):
            continue
        key = role_style_key(entry)
        if key is None:
            continue
        sid, sname = key[0], key[1]

        def _match(f, sid=sid, sname=sname) -> bool:
            k = f.style_key
            if k is None:
                return False
            k_sid, k_sname = k[0], k[1]
            return bool((sid and k_sid == sid) or (sname and k_sname == sname))

        geom = _fold_geometry(facts, _match)
        if not geom:
            continue
        entry.setdefault("appearance", {})["geometry"] = geom


def capture_palette_facts(
    run_facts: Iterable[RunFacts], roles: dict, theme: dict
) -> None:
    """Capture the template's brand PALETTE into ``theme['palette']`` (mutated in
    place), additively and deterministically.

    Format-neutral reproduction of the docx ``capture_palette`` body, byte-identical:
    a SINGLE ordered pass over ``run_facts`` builds the dominance ``Counter`` (every
    run votes, a ``None`` bucket for an inherited color), the explicit-color frequency
    counts, and the set of buckets seen on a link run; then it seeds the theme slots,
    records each observed run color with its coarse frequency, folds link colors (plus
    the ``hlink``/``folHlink`` theme fallback), folds the per-role ``appearance.color``
    already captured, and records the non-authoritative ``palette_role`` map. The
    ``run_facts`` MUST be a fresh ordered pass (a second iterable, not the one
    :func:`capture_appearance` consumed) so the deterministic insertion order holds."""
    palette: dict = theme.setdefault("palette", {})

    overall_color: Counter = Counter()
    run_color_counts: Counter = Counter()
    link_buckets: set[tuple[str, ...]] = set()
    for fact in run_facts:
        if not (fact.text or "").strip():
            continue
        bucket = fact.color
        overall_color[bucket] += 1
        if bucket is None:
            continue
        run_color_counts[bucket] += 1
        if fact.is_link:
            link_buckets.add(bucket)

    dominant_color = _dominant(overall_color)
    dominant_bucket = dominant_color[0] if dominant_color is not None else None

    # (a) Seed theme-keyed entries for every slot the template's theme carries.
    for slot in theme.get("colors") or {}:
        if slot in THEME_SLOTS:
            _palette_entry(palette, ("theme", slot))

    # (b) record each observed run color, with its coarse frequency.
    for bucket, count in run_color_counts.items():
        entry = _palette_entry(palette, bucket)
        if bucket == dominant_bucket:
            entry["frequency"] = "dominant"
        elif count >= MIN_ACCENT_RUNS:
            entry["frequency"] = "accent"
        else:
            entry["frequency"] = "rare"
        _add_provenance(entry, "run.color", _palette_key(bucket))

    # (d) link-color where-facts plus the hlink/folHlink theme fallback.
    for bucket in link_buckets:
        _add_provenance(
            _palette_entry(palette, bucket), "link.color", _palette_key(bucket)
        )
    for slot in ("hlink", "folHlink"):
        if slot in (theme.get("colors") or {}):
            entry = _palette_entry(palette, ("theme", slot))
            _add_provenance(entry, "link.color", slot)

    # (c) per-role appearance.color already captured (role.appearance where-fact).
    for rid, role_entry in roles.items():
        if rid == "_index" or not isinstance(role_entry, dict):
            continue
        color = (role_entry.get("appearance") or {}).get("color")
        if not isinstance(color, dict):
            continue
        bucket = _color_obj_to_bucket(color)
        if bucket is None:
            continue
        _add_provenance(_palette_entry(palette, bucket), "role.appearance", rid)

    # palette_role: the hardcoded, template-INVARIANT map - recorded NON-authoritatively.
    for prole, ref in (theme.get("palette_roles") or {}).items():
        slot = ref.get("theme") if isinstance(ref, dict) else None
        if slot and slot in THEME_SLOTS:
            _add_provenance(
                _palette_entry(palette, ("theme", slot)), "palette_role", prole
            )
