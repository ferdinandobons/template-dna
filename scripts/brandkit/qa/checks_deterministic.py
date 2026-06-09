# SPDX-License-Identifier: MIT
"""Deterministic L0 checks for M1."""

from __future__ import annotations

import hashlib
import re
from typing import NamedTuple
from pathlib import Path

from docx import Document
from openpyxl import load_workbook
from pptx import Presentation

from brandkit.common import color as colorutil
from brandkit.common import text as textutil
from brandkit.ooxml import names, pack
from brandkit.profile import comprehension as comprehensionmod
from brandkit.profile import schema
from brandkit.profile.reconcile import confidence_clears_floor
from brandkit.qa.model import Finding

# WordprocessingML qualified-name builder (the spec-fixed namespace), so the shell
# readers below never hand-copy the namespace URI.
_W = names.make_qn("w")
# DrawingML qualified-name builder (the pptx run/theme namespace), shared with the
# pptx appearance collector below.
_A = names.make_qn("a")


def check_profile(profile: dict) -> list[Finding]:
    findings: list[Finding] = []
    for problem in schema.validate(profile):
        findings.append(Finding("schema", schema.Severity.ERROR.value, problem))
    for rid in (profile.get("roles") or {}).get("_index", []):
        entry = profile.get("roles", {}).get(rid)
        if not entry or not entry.get("resolver"):
            findings.append(
                Finding(
                    "every_role_resolves",
                    schema.Severity.ERROR.value,
                    f"{rid} has no resolver",
                )
            )
    return findings


def check_no_duplicate_parts(target) -> list[Finding]:
    """ERROR when the output OOXML package has duplicate ZIP part names.

    A valid OPC package is a ZIP whose part names are unique; a duplicate (e.g. a
    slide-partname collision after removing a non-last slide) makes PowerPoint/Word
    show a repair dialog. Cheap, format-agnostic integrity backstop that no
    text/structure scan catches; a missing/garbage file is handled by other checks.
    """
    import zipfile
    from collections import Counter

    if target is None:
        return []
    try:
        with zipfile.ZipFile(target) as z:
            names = z.namelist()
    except (OSError, zipfile.BadZipFile):
        return []
    dups = sorted(n for n, c in Counter(names).items() if c > 1)
    if not dups:
        return []
    shown = ", ".join(dups[:5]) + (" ..." if len(dups) > 5 else "")
    return [
        Finding(
            check="package_integrity",
            severity=schema.Severity.ERROR.value,
            message=f"output package has duplicate part name(s): {shown}",
        )
    ]


def check_shell_provenance(shell, profile: dict) -> list[Finding]:
    """ERROR when the saved shell no longer matches profile provenance.

    The Brand Profile records ``provenance.shell.sha256`` at extraction time. A
    later hand edit, corrupt copy, or tamper must be load-bearing in QA: generation
    from a drifted shell is not the same brand contract the profile verified.
    """
    if shell is None:
        return []
    recorded = ((profile.get("provenance") or {}).get("shell") or {}).get("sha256")
    if not recorded:
        return []

    shell_path = Path(shell)
    if not shell_path.is_file():
        return [
            Finding(
                "shell_provenance",
                schema.Severity.ERROR.value,
                f"recorded shell hash exists but shell is missing: {shell_path}",
                location=str(shell_path),
            )
        ]

    try:
        actual = _sha256_file(shell_path)
    except OSError as exc:
        return [
            Finding(
                "shell_provenance",
                schema.Severity.ERROR.value,
                f"could not hash shell {shell_path}: {exc}",
                location=str(shell_path),
            )
        ]

    if actual != recorded:
        return [
            Finding(
                "shell_provenance",
                schema.Severity.ERROR.value,
                f"shell hash drifted: recorded {recorded}, actual {actual}",
                location=str(shell_path),
            )
        ]
    return []


def _sha256_file(path: Path, *, chunk: int = 1 << 20) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for block in iter(lambda: fh.read(chunk), b""):
            h.update(block)
    return h.hexdigest()


def check_resolver_targets(shell, profile: dict) -> list[Finding]:
    """Verify every role's resolver target actually exists in the ``shell``.

    This is the deterministic backstop for the core promise ("apply only artifacts
    the profile *proved* exist"): it opens the shell once and confirms each role's
    concrete target is present -
      - docx: ``style_id`` or ``style_name`` ∈ ``doc.styles``;
      - pptx: ``layout`` ∈ the presentation's slide-layout names;
      - xlsx: ``name`` ∈ the workbook's defined names.
    A role whose target is missing yields an ERROR ``resolver_targets_exist``
    finding. ``shell`` may be a path or None; when None the check is skipped (the
    caller is responsible for supplying the shell at verify time).
    """
    if shell is None:
        return []
    kind = profile.get("kind")
    roles = profile.get("roles") or {}
    role_ids = [r for r in roles if r != "_index"]
    findings: list[Finding] = []
    try:
        if kind == schema.Kind.DOCX.value:
            present = _docx_style_keys(shell)
            for rid in role_ids:
                resolver = (roles.get(rid) or {}).get("resolver") or {}
                if resolver.get("type") not in (
                    schema.ResolverType.NAMED_STYLE.value,
                    None,
                ):
                    continue
                sid = resolver.get("style_id")
                sname = resolver.get("style_name")
                if not sid and not sname:
                    continue
                if (sid and sid in present) or (sname and sname in present):
                    continue
                findings.append(
                    Finding(
                        "resolver_targets_exist",
                        schema.Severity.ERROR.value,
                        f"role {rid!r} target style {(sid or sname)!r} not found in shell",
                        location=rid,
                    )
                )
        elif kind == schema.Kind.PPTX.value:
            layout_names = _pptx_layout_names(shell)
            for rid in role_ids:
                resolver = (roles.get(rid) or {}).get("resolver") or {}
                if resolver.get("type") != schema.ResolverType.PLACEHOLDER.value:
                    continue
                layout = resolver.get("layout")
                if layout is None or layout in layout_names:
                    continue
                findings.append(
                    Finding(
                        "resolver_targets_exist",
                        schema.Severity.ERROR.value,
                        f"role {rid!r} layout {layout!r} not found in shell (have {sorted(layout_names)})",
                        location=rid,
                    )
                )
        elif kind == schema.Kind.XLSX.value:
            defined = _xlsx_defined_names(shell)
            for rid in role_ids:
                resolver = (roles.get(rid) or {}).get("resolver") or {}
                if resolver.get("type") != schema.ResolverType.NAMED_RANGE.value:
                    continue
                name = resolver.get("name")
                if name is None or name in defined:
                    continue
                findings.append(
                    Finding(
                        "resolver_targets_exist",
                        schema.Severity.ERROR.value,
                        f"role {rid!r} named range {name!r} not found in shell (have {sorted(defined)})",
                        location=rid,
                    )
                )
            # number_format roles must resolve to a mask the shell ACTUALLY uses
            # (the engine never fabricates a format) - the shell-backed peer of the
            # schema's intra-profile number_format consistency check.
            masks = _xlsx_number_format_masks(shell)
            for rid in role_ids:
                resolver = (roles.get(rid) or {}).get("resolver") or {}
                if resolver.get("type") != schema.ResolverType.NUMBER_FORMAT.value:
                    continue
                mask = resolver.get("number_format")
                if mask is None or mask in masks:
                    continue
                findings.append(
                    Finding(
                        "resolver_targets_exist",
                        schema.Severity.ERROR.value,
                        f"role {rid!r} number format {mask!r} not among the shell's "
                        "used formats",
                        location=rid,
                    )
                )
    except Exception as exc:  # opening the shell must never crash the gate
        findings.append(
            Finding(
                "resolver_targets_exist",
                schema.Severity.WARNING.value,
                f"could not verify resolver targets against shell: {exc}",
            )
        )
    return findings


def _docx_style_keys(shell) -> set:
    doc = Document(shell)
    keys: set = set()
    for style in doc.styles:
        sid = getattr(style, "style_id", None)
        if sid:
            keys.add(sid)
        name = getattr(style, "name", None)
        if name:
            keys.add(name)
    return keys


def _docx_shell_fonts(shell, profile: dict) -> set:
    """Font names the shell makes available: every ``w:font`` in its fontTable, the
    theme major/minor latin typefaces, and the Arial ``docDefaults`` baseline."""
    fonts: set = {"Arial"}
    theme_fonts = (profile.get("theme") or {}).get("fonts") or {}
    for key in ("major", "minor"):
        latin = (theme_fonts.get(key) or {}).get("latin")
        if latin:
            fonts.add(latin)
    try:
        xml = pack.read_part(shell, "word/fontTable.xml")
    except KeyError:
        return fonts
    root = pack.parse_xml_bytes(xml)
    for font_el in root.findall(_W("font")):
        name = font_el.get(_W("name"))
        if name:
            fonts.add(name)
    return fonts


def _docx_shell_run_sizes_and_hexes(shell) -> tuple[set[int], set[str]]:
    """The PROVENANCE sets an applied size/color is validated against, read from the
    template's ``word/document.xml`` runs in a SINGLE parse:

      - sizes: every ``w:sz@w:val`` (half-points) the template's OWN runs carry;
      - hexes: every ``w:color@w:val`` (normalized hex), skipping ``auto``.

    A captured size/hex is only trusted when the template itself uses it (not a sanity
    bound). The hardened parser (``parse_xml_bytes``, ``resolve_entities=False``)
    matches every other shell-XML read. A missing/garbage document part yields empty
    sets - the caller then fails closed on any applied size/color."""
    sizes: set[int] = set()
    hexes: set[str] = set()
    try:
        xml = pack.read_part(shell, "word/document.xml")
    except KeyError:
        return sizes, hexes
    root = pack.parse_xml_bytes(xml)
    for sz in root.iter(_W("sz")):
        val = sz.get(_W("val"))
        if val is None:
            continue
        try:
            sizes.add(int(val))
        except (TypeError, ValueError):
            continue
    for color in root.iter(_W("color")):
        val = color.get(_W("val"))
        if not val or val.lower() == "auto":
            continue
        try:
            hexes.add(colorutil.normalize_hex(val))
        except (ValueError, AttributeError):
            continue
    return sizes, hexes


# The WordprocessingML themeColor -> clrScheme slot alias table, the single source of
# truth in brandkit.common.color (shared with the resolver's hex enrichment). A token
# outside it has no palette slot and is rejected (fail-closed).
_WML_THEME_TO_SLOT = colorutil.WML_THEME_TO_SLOT


def _docx_theme_palette(shell) -> dict[str, str]:
    """The template's parsed theme palette (slot -> hex). Reads the shell's own
    ``theme1.xml``; an absent part yields an empty palette (every theme-token color
    then fails closed)."""
    try:
        xml = pack.read_part(shell, "word/theme/theme1.xml")
    except KeyError:
        return {}
    return colorutil.parse_theme_colors(xml)


# ---------------------------------------------------------------------------
# Format-neutral shell appearance facts (A3): the PROVENANCE every applied
# font/size/color is re-validated against, collected per kind into ONE shape.
# ---------------------------------------------------------------------------
class ShellAppearanceFacts(NamedTuple):
    """The typographic provenance a shell proves it carries, format-neutral.

    The single shape every per-kind collector fills so the membership loops in
    :func:`check_appearance_targets` are format-agnostic:

      - ``fonts``: the font names the shell makes AVAILABLE (the allow-set an applied
        font is validated against - the fontTable/theme/Arial baseline, widened, never
        narrowed);
      - ``sizes``: every run size the shell's OWN runs carry, in HALF-POINTS (the unit
        :data:`theme.fonts.body.size_hp` and ``role.appearance.size_hp`` use);
      - ``hexes``: every explicit run hex the shell carries (normalized ``RRGGBB``);
      - ``palette``: the parsed ``clrScheme`` slot -> normalized hex (the theme palette a
        theme-token color resolves against).
    """

    fonts: set[str]
    sizes: set[int]
    hexes: set[str]
    palette: dict[str, str]


def _docx_collect_appearance_facts(shell, profile: dict) -> ShellAppearanceFacts:
    """The docx appearance facts, BYTE-IDENTICALLY composed from the existing docx
    shell readers (``_docx_shell_fonts`` + ``_docx_shell_run_sizes_and_hexes`` +
    ``_docx_theme_palette``). No new docx parsing: this is a pure assembly so the docx
    verify verdict is unchanged."""
    fonts = _docx_shell_fonts(shell, profile)
    sizes, hexes = _docx_shell_run_sizes_and_hexes(shell)
    palette = _docx_theme_palette(shell)
    return ShellAppearanceFacts(fonts=fonts, sizes=sizes, hexes=hexes, palette=palette)


# pptx ``a:rPr@sz`` / ``a:defRPr@sz`` are in CENTIPOINTS (1/100 pt); the engine's
# size unit is HALF-POINTS, so a value divides by 50 (e.g. 1800 -> 36 half-points).
_PPTX_SZ_PER_HALF_POINT = 50
_PPTX_SLIDE_PART = re.compile(r"ppt/slides/slide\d+\.xml")


def _pptx_collect_appearance_facts(shell, profile: dict) -> ShellAppearanceFacts:
    """The pptx appearance facts, read from the deck's own parts.

    - ``fonts``: every ``a:latin@typeface`` in the theme major/minor font scheme
      (``ppt/theme/theme1.xml``), widened by the Arial baseline and the profile's
      captured ``theme.fonts.major/minor.latin`` (the allow-set, never narrowed);
    - ``sizes``: every ``a:rPr@sz`` / ``a:defRPr@sz`` across the slide parts,
      converted from centipoints to half-points (``÷50``);
    - ``hexes``: every ``a:srgbClr@val`` under a run/defRPr ``a:solidFill`` on the
      slides, normalized;
    - ``palette``: the parsed clrScheme of the deck's theme.
    """
    fonts: set[str] = {"Arial"}
    theme_fonts = (profile.get("theme") or {}).get("fonts") or {}
    for key in ("major", "minor"):
        latin = (theme_fonts.get(key) or {}).get("latin")
        if latin:
            fonts.add(latin)
    try:
        theme_xml = pack.read_part(shell, "ppt/theme/theme1.xml")
    except KeyError:
        palette: dict[str, str] = {}
    else:
        palette = colorutil.parse_theme_colors(theme_xml)
        theme_root = pack.parse_xml_bytes(theme_xml)
        for latin in theme_root.iter(_A("latin")):
            face = latin.get("typeface")
            if face:
                fonts.add(face)
    sizes: set[int] = set()
    hexes: set[str] = set()
    for part in pack.list_parts(shell):
        if not _PPTX_SLIDE_PART.fullmatch(part):
            continue
        root = pack.parse_xml_bytes(pack.read_part(shell, part))
        for tag in ("rPr", "defRPr"):
            for rpr in root.iter(_A(tag)):
                val = rpr.get("sz")
                if val is None:
                    continue
                try:
                    # round() (not floor //) to match the capture side
                    # (pptx/typography: round(size.pt * 2) == round(centipoints / 50)),
                    # so a correctly applied fractional size never spuriously fails closed.
                    sizes.add(round(int(val) / _PPTX_SZ_PER_HALF_POINT))
                except (TypeError, ValueError):
                    continue
        for srgb in root.iter(_A("srgbClr")):
            val = srgb.get("val")
            if not val:
                continue
            try:
                hexes.add(colorutil.normalize_hex(val))
            except (ValueError, AttributeError):
                continue
    return ShellAppearanceFacts(fonts=fonts, sizes=sizes, hexes=hexes, palette=palette)


def _xlsx_collect_appearance_facts(shell, profile: dict) -> ShellAppearanceFacts:
    """The xlsx appearance facts, read from the workbook's own cells/styles/theme.

    - ``fonts``: every ``cell.font.name`` across the materialized cells, widened by
      each NamedStyle font name and the Arial baseline (the allow-set);
    - ``sizes``: ``round(font.sz * 2)`` for every explicit cell-font size (half-points,
      the same unit the xlsx capture records);
    - ``hexes``: every ``font.color.rgb`` that is an ``'rgb'`` color, 8-digit ARGB
      alpha stripped, normalized (non-rgb/indexed/auto colors contribute nothing);
    - ``palette``: the parsed clrScheme of the workbook's theme.
    """
    fonts: set[str] = {"Arial"}
    sizes: set[int] = set()
    hexes: set[str] = set()
    wb = load_workbook(shell, data_only=False)
    # ``wb.named_styles`` is a list of style NAMES (strings); the NamedStyle objects
    # (which carry the ``.font``) live on ``wb._named_styles``. Guard both for
    # openpyxl-version robustness - a NamedStyle font widens the allow-set.
    for style in getattr(wb, "_named_styles", []) or []:
        font = getattr(style, "font", None)
        name = getattr(font, "name", None) if font is not None else None
        if name:
            fonts.add(name)
    for ws in wb.worksheets:
        for cell in ws._cells.values():
            font = cell.font
            if font is None:
                continue
            if font.name:
                fonts.add(font.name)
            if font.sz is not None:
                try:
                    sizes.add(round(float(font.sz) * 2))
                except (TypeError, ValueError):
                    pass
            color = getattr(font, "color", None)
            if color is None or getattr(color, "type", None) != "rgb":
                continue
            rgb = color.rgb
            if not isinstance(rgb, str):
                continue
            hexpart = rgb[2:] if len(rgb) == 8 else rgb
            try:
                hexes.add(colorutil.normalize_hex(hexpart))
            except (ValueError, AttributeError):
                continue
    try:
        theme_xml = pack.read_part(shell, "xl/theme/theme1.xml")
    except KeyError:
        palette: dict[str, str] = {}
    else:
        palette = colorutil.parse_theme_colors(theme_xml)
    return ShellAppearanceFacts(fonts=fonts, sizes=sizes, hexes=hexes, palette=palette)


# The per-kind shell appearance collector (peer of :data:`_COMPONENT_COUNTERS`). Each
# fills the format-neutral :class:`ShellAppearanceFacts`; the lifted gate dispatches
# off ``profile['kind']`` and feeds the result into the unchanged membership loops.
_SHELL_APPEARANCE_COLLECTORS = {
    schema.Kind.DOCX.value: _docx_collect_appearance_facts,
    schema.Kind.PPTX.value: _pptx_collect_appearance_facts,
    schema.Kind.XLSX.value: _xlsx_collect_appearance_facts,
}


def _collect_applied_appearance(
    profile: dict,
) -> tuple[list[tuple[str, int]], list[tuple[str, dict]]]:
    """Gather every (where, size_hp) and (where, color-obj) the engine will APPLY.

    Scans per-role ``appearance`` plus the document body defaults
    (``theme.fonts.body.size_hp`` for size, ``theme.text.body.color`` for color),
    plus EVERY ``theme.palette[*].ref`` color object - a run carrying a palette
    color token is realized into that ref by the resolver, so each palette ref is an
    APPLIED color the shell must independently prove it carries. Folding the palette
    refs in here means ``check_appearance_targets`` re-validates them with the exact
    same loop (hex vs the palette UNION the observed ``w:color``; theme token vs the
    mapped ``clrScheme`` slot), fail-closed - the model can NAME a palette color but
    the deterministic shell still has to back it."""
    sizes: list[tuple[str, int]] = []
    colors: list[tuple[str, dict]] = []
    for rid, entry in (profile.get("roles") or {}).items():
        if rid == "_index" or not isinstance(entry, dict):
            continue
        appearance = entry.get("appearance") or {}
        size_hp = appearance.get("size_hp")
        if size_hp:
            sizes.append((f"role {rid!r}", int(size_hp)))
        color = appearance.get("color")
        if isinstance(color, dict):
            colors.append((f"role {rid!r}", color))
    theme = profile.get("theme") or {}
    body_size = ((theme.get("fonts") or {}).get("body") or {}).get("size_hp")
    if body_size:
        sizes.append(("theme.fonts.body", int(body_size)))
    body_color = ((theme.get("text") or {}).get("body") or {}).get("color")
    if isinstance(body_color, dict):
        colors.append(("theme.text.body", body_color))
    # Every palette ref is an applicable run color (a run color token resolves to
    # it), so re-validate it against the shell with the same loop. Sorted for a
    # deterministic finding order.
    for key in sorted((theme.get("palette") or {})):
        entry = (theme.get("palette") or {}).get(key)
        if not isinstance(entry, dict):
            continue
        ref = entry.get("ref")
        if isinstance(ref, dict):
            colors.append((f"theme.palette.{key}", ref))
    return sizes, colors


def check_appearance_targets(shell, profile: dict) -> list[Finding]:
    """Verify every typography value the engine will APPLY is one the ``shell``
    actually proves it carries (font family, run size, run color).

    Typography capture records the template's dominant direct run font/size/color
    into ``role.appearance`` and the document defaults (``theme.fonts.body`` for
    font/size, ``theme.text.body`` for color); this is their shell-backed peer of
    :func:`check_resolver_targets`. Each axis re-validates against true PROVENANCE,
    not a sanity bound, and fails closed (ERROR) on any applied value the shell does
    not prove it contains - off-brand output stays impossible by construction:

      - FONT: must be in the shell's available faces (fontTable/theme latin + the
        Arial baseline; the theme declarations widen the ALLOWED set, never the
        checked set).
      - SIZE: the applied ``size_hp`` must be a run size actually present on the
        template's own runs.
      - COLOR (hex): the applied hex must be a theme-palette hex OR a hex actually
        present on the template's runs.
      - COLOR (theme token): the token's mapped ``clrScheme`` slot
        (:data:`_WML_THEME_TO_SLOT`) must be present in the parsed palette.

    Format-neutral: the per-kind :data:`_SHELL_APPEARANCE_COLLECTORS` reduce the shell
    to one :class:`ShellAppearanceFacts` shape (docx/pptx/xlsx) and the membership loops
    below are kind-agnostic. A collector that cannot parse fails CLOSED - it emits a
    WARNING and yields empty fact sets, so every applied value is then rejected (ERROR).

    A no-op when no appearance value is present (every pre-capture profile), when the
    shell is absent, or when the kind has no registered collector.
    """
    if shell is None:
        return []
    collector = _SHELL_APPEARANCE_COLLECTORS.get(profile.get("kind"))
    if collector is None:
        return []
    findings: list[Finding] = []
    try:
        facts = collector(shell, profile)
    except Exception as exc:  # opening the shell must never crash the gate
        # Fail closed: surface the parse failure and continue with empty fact sets, so
        # any applied font/size/color is rejected below rather than silently passing.
        findings.append(
            Finding(
                "appearance_targets_exist",
                schema.Severity.WARNING.value,
                f"could not verify appearance targets against shell: {exc}",
            )
        )
        facts = ShellAppearanceFacts(fonts=set(), sizes=set(), hexes=set(), palette={})
    available = facts.fonts
    applied: list[tuple[str, str]] = []
    for rid, entry in (profile.get("roles") or {}).items():
        if rid == "_index" or not isinstance(entry, dict):
            continue
        latin = ((entry.get("appearance") or {}).get("font") or {}).get("latin")
        if latin:
            applied.append((f"role {rid!r}", latin))
    body_latin = (
        ((profile.get("theme") or {}).get("fonts") or {}).get("body") or {}
    ).get("latin")
    if body_latin:
        applied.append(("theme.fonts.body", body_latin))

    for where, font in applied:
        if font not in available:
            findings.append(
                Finding(
                    "appearance_targets_exist",
                    schema.Severity.ERROR.value,
                    f"{where} font {font!r} is not available in the shell "
                    f"(have {sorted(available)})",
                )
            )

    applied_sizes, applied_colors = _collect_applied_appearance(profile)
    if applied_sizes or applied_colors:
        shell_sizes, shell_hexes, palette = facts.sizes, facts.hexes, facts.palette
        for where, size_hp in applied_sizes:
            if size_hp not in shell_sizes:
                findings.append(
                    Finding(
                        "appearance_targets_exist",
                        schema.Severity.ERROR.value,
                        f"{where} size {size_hp} (half-points) is not present on any "
                        f"run in the shell (have {sorted(shell_sizes)})",
                    )
                )
        palette_hexes = {colorutil.normalize_hex(h) for h in palette.values()}
        allowed_hexes = palette_hexes | shell_hexes
        for where, color in applied_colors:
            kind = color.get("kind")
            if kind == "hex":
                try:
                    hexval = colorutil.normalize_hex(color.get("hex") or "")
                except (ValueError, AttributeError):
                    hexval = None
                if hexval is None or hexval not in allowed_hexes:
                    findings.append(
                        Finding(
                            "appearance_targets_exist",
                            schema.Severity.ERROR.value,
                            f"{where} color #{color.get('hex')!r} is not in the shell "
                            f"palette or its runs (allowed {sorted(allowed_hexes)})",
                        )
                    )
            elif kind == "theme":
                token = color.get("theme")
                # A theme ref may name the color in EITHER namespace and both must
                # validate against the parsed clrScheme palette:
                #   - a WML ``themeColor`` token (``text1`` / ``background1`` /
                #     ``hyperlink`` ...) captured off a run, mapped to its
                #     clrScheme slot via _WML_THEME_TO_SLOT;
                #   - a clrScheme slot name directly (``dk1`` / ``lt1`` / ``hlink``
                #     ... or ``accent1``), which ``theme.palette`` keys/seeds use
                #     verbatim (it has no WML alias for the dk/lt/hlink slots).
                # The token is present when it resolves to a parsed-palette slot
                # through either route; only a token that does neither fails closed.
                slot = _WML_THEME_TO_SLOT.get(token)
                present = (slot is not None and slot in palette) or (
                    token in colorutil.THEME_SLOTS and token in palette
                )
                if not present:
                    findings.append(
                        Finding(
                            "appearance_targets_exist",
                            schema.Severity.ERROR.value,
                            f"{where} theme color {token!r} is absent from the shell "
                            f"palette (have {sorted(palette)})",
                        )
                    )
    return findings


def _pptx_layout_names(shell) -> set:
    prs = Presentation(shell)
    return {layout.name for layout in prs.slide_layouts}


def _xlsx_defined_names(shell) -> set:
    wb = load_workbook(shell, data_only=False)
    try:
        return set(wb.defined_names.keys())
    except AttributeError:
        # Older openpyxl exposes defined_names as a list-like of DefinedName.
        return {dn.name for dn in wb.defined_names}


def _xlsx_number_format_masks(shell) -> set:
    """The distinct number-format masks the workbook actually uses (``General`` dropped).

    Mirrors ``xlsx_structure.inventory_number_formats`` as a set, kept local so the
    QA layer needs no format-module import; iterates ``_cells`` (sparse-safe).
    """
    wb = load_workbook(shell, data_only=False)
    masks: set = set()
    for ws in wb.worksheets:
        for cell in ws._cells.values():
            if cell.value is None:
                continue
            fmt = cell.number_format
            if fmt and fmt != "General":
                masks.add(fmt)
    return masks


# ---------------------------------------------------------------------------
# Comprehension-aware deterministic checks (model-free; bind to captured facts)
# ---------------------------------------------------------------------------
def check_comprehension_targets(profile: dict) -> list[Finding]:
    """FAIL-CLOSED membership of every load-bearing comprehension ref (§3).

    Every ``anchor_ref`` / ``index_ref`` / ``region_ref`` / ``feeds_from_role_id``
    / ``role_annotations`` key must be a verbatim id from the surfaced inventories;
    a ref whose inventory is empty/absent is itself an ERROR (this is the SOLE gate
    for anchor/index/region refs, so it must reject, never skip). No-ops when the
    comprehension is absent (status != present), keeping the model-free CI path and
    pptx/xlsx green. Delegates to the single membership definition so the gate and
    the merge writer can never disagree.

    ``check_membership`` returns ALL membership problems, including the C1 audit-arm
    ``comprehension.audit`` keys. Those are attributed SOLELY to ``audit_targets_exist``
    by :func:`check_audit_targets`, so they are skipped here - else one bad audit key
    would double-report (one ``comprehension_targets_exist`` + one
    ``audit_targets_exist``). The skip is symmetric to ``check_audit_targets``' keep.
    """
    comp = profile.get("comprehension")
    if (
        not isinstance(comp, dict)
        or comp.get("status") != schema.ComprehensionStatus.PRESENT.value
    ):
        return []
    findings: list[Finding] = []
    for problem in comprehensionmod.check_membership(profile, comp):
        if problem.startswith("comprehension.audit"):
            continue  # attributed to audit_targets_exist (check_audit_targets)
        findings.append(
            Finding("comprehension_targets_exist", schema.Severity.ERROR.value, problem)
        )
    return findings


def check_audit_targets(profile: dict) -> list[Finding]:
    """FAIL-CLOSED membership of every persisted L2 visual-audit verdict key (C1).

    The audit peer of :func:`check_comprehension_targets`: every
    ``comprehension.audit`` key must be a verbatim id from the profile's derived
    visual checklist; a key not in that set (or any key when the derived checklist
    is empty) is an ERROR (this is the SOLE gate for audit keys, so it must reject,
    never skip). No-ops when the comprehension is absent (status != present),
    keeping the model-free CI path and pptx/xlsx green. Delegates to the single
    membership definition (``check_membership``'s audit arm, which itself binds to
    ``qa.visual.visual_checklist_ids``) so the gate and the merge writer can never
    disagree about which checklist ids exist.

    A surfaced finding here emits ``audit_targets_exist`` (the
    :data:`schema.DEFAULT_L0_INVARIANTS` id) rather than the
    ``comprehension_targets_exist`` id that ``check_membership`` problems carry under
    :func:`check_comprehension_targets`, so the verdict-key violation is attributable
    to its own invariant. ``check_membership`` returns ALL membership problems
    (anchor/index/region/role/palette + audit); we keep only the audit ones here so
    the two checks do not double-report the same anchor/index problem.
    """
    comp = profile.get("comprehension")
    if (
        not isinstance(comp, dict)
        or comp.get("status") != schema.ComprehensionStatus.PRESENT.value
    ):
        return []
    findings: list[Finding] = []
    for problem in comprehensionmod.check_membership(profile, comp):
        if problem.startswith("comprehension.audit"):
            findings.append(
                Finding("audit_targets_exist", schema.Severity.ERROR.value, problem)
            )
    return findings


def check_triage_targets(profile: dict) -> list[Finding]:
    """FAIL-CLOSED membership of every model-assisted QA-triage entry (Cluster C2).

    The triage peer of :func:`check_audit_targets`: every ``comprehension.triage``
    entry must name a check in the closed :data:`schema.AMBIGUOUS_TRIAGE_CHECKS` set,
    and each ``(check, location)`` pair must be unique across the proposal. A
    non-eligible check or a duplicate pair is an ERROR (this is the SOLE gate for
    triage entries, so it must reject, never skip). No-ops when the comprehension is
    absent (status != present), keeping the model-free CI path and pptx/xlsx green.
    Delegates to the single membership definition (``check_triage``) so the gate and
    the merge writer can never disagree about which checks are triage-eligible.

    Because the eligible set is WARNING-only, a triage entry can never even be aimed
    at an ERROR-emitting check - it would be rejected here as a non-member. That is
    the belt to ``qa.gate._apply_triage``'s ``severity == WARNING`` suspenders, so an
    ERROR can NEVER be demoted.
    """
    comp = profile.get("comprehension")
    if (
        not isinstance(comp, dict)
        or comp.get("status") != schema.ComprehensionStatus.PRESENT.value
    ):
        return []
    findings: list[Finding] = []
    for problem in comprehensionmod.check_triage(profile, comp):
        findings.append(
            Finding("triage_targets_exist", schema.Severity.ERROR.value, problem)
        )
    return findings


def _role_resolver_target_in_shell(kind, resolver: dict, present: set) -> bool:
    """True iff one role's resolver target is proven by the shell's ``present`` set.

    The membership predicate of ``check_override_targets``' reroute branch. It
    MIRRORS ``check_resolver_targets``' per-kind membership arms but is deliberately
    STRICTER, so it is not literally shared: ``check_resolver_targets`` SKIPS a role
    whose resolver is incomplete (a missing layout/name is another validator's
    problem), while a reroute target with the same incomplete resolver must be
    REJECTED here - a lesson may only re-point at a concrete, shell-backed op, never
    a stub. ``present`` is the kind-appropriate inventory (style keys / layout names
    / defined names+masks).
    """
    rtype = resolver.get("type")
    if kind == schema.Kind.DOCX.value:
        if rtype not in (schema.ResolverType.NAMED_STYLE.value, None):
            return False
        sid = resolver.get("style_id")
        sname = resolver.get("style_name")
        if not sid and not sname:
            return False
        return bool((sid and sid in present) or (sname and sname in present))
    if kind == schema.Kind.PPTX.value:
        if rtype != schema.ResolverType.PLACEHOLDER.value:
            return False
        layout = resolver.get("layout")
        return layout is not None and layout in present
    if kind == schema.Kind.XLSX.value:
        if rtype == schema.ResolverType.NAMED_RANGE.value:
            name = resolver.get("name")
            return name is not None and name in present
        if rtype == schema.ResolverType.NUMBER_FORMAT.value:
            mask = resolver.get("number_format")
            return mask is not None and mask in present
        return False
    return False


def check_override_targets(shell, profile: dict) -> list[Finding]:
    """FAIL-CLOSED membership of every LEARNED override target against the shell.

    The override peer of :func:`check_resolver_targets`, with the **reject-on-empty**
    discipline of :func:`check_comprehension_targets` (NOT the namespace-guarded
    skip-on-empty of ``_validate_resolver_consistency``): when a lesson is present,
    every re-point target must be proven by THIS shell, and a target whose inventory
    is empty/absent is itself an ERROR (this is the SOLE gate for override targets, so
    it must reject, never skip). Emits ERROR ``override_targets_exist`` when:

      - a ``reroute_role`` ``to`` is not a concrete declared role key (``_index`` is
        the order array, not a role), or that role's own resolver target is not
        shell-proven (so a reroute can only ever land on a real, shell-backed op);
      - a ``number_format`` swap mask is not among the shell's used masks
        (``surface.xlsx.number_formats`` via :func:`_xlsx_number_format_masks`);
      - a ``demo_clear`` value was not captured for this template
        (:func:`captured_template_texts`).

    No-ops when overrides are absent (``status != present``), keeping the model-free
    CI path and a pre-B3 profile green - exactly like ``check_comprehension_targets``.
    ``shell`` may be None at verify time only for the reroute/mask branches that need
    to open the package; the demo-clear branch reads the profile's captured set, so it
    fails closed even without a shell.
    """
    overrides = schema.overrides_block(profile)
    if overrides.get("status") != schema.ComprehensionStatus.PRESENT.value:
        return []
    kind = profile.get("kind")
    roles = profile.get("roles") or {}
    role_keys = {r for r in roles if r != "_index"}
    findings: list[Finding] = []

    reroutes = overrides.get("reroute_roles") or {}
    swaps = overrides.get("number_format_swaps") or {}
    clears = overrides.get("demo_clears") or []

    # Pure-profile membership (no shell needed, so it fails closed even at verify time
    # with shell=None or a broken shell): a reroute target must be a CONCRETE declared
    # role (``_index`` is the order array, not a role). Checked unconditionally; the
    # shell-backed resolver proof below only runs the resolver gate on the survivors.
    shell_backed_reroutes: dict[str, str] = {}
    for requested, target in reroutes.items():
        if target not in role_keys:
            findings.append(
                Finding(
                    "override_targets_exist",
                    schema.Severity.ERROR.value,
                    f"reroute target {target!r} (for role {requested!r}) is not a "
                    f"declared role (have {sorted(role_keys)})",
                )
            )
        else:
            shell_backed_reroutes[requested] = target

    # demo-clear values are membership-checked against the profile's OWN captured demo
    # set (no shell open needed) FIRST, so a later shell-open failure can never suppress
    # this fail-closed check. reject-on-empty: a clear value absent from the captured
    # set (including an empty set) is an ERROR.
    if clears:
        captured = set(captured_template_texts(profile))
        for value in clears:
            if value not in captured:
                findings.append(
                    Finding(
                        "override_targets_exist",
                        schema.Severity.ERROR.value,
                        f"demo_clear value {value!r} was not captured for this template",
                    )
                )

    # The resolver-target proof and the mask membership need the shell inventory; open
    # it ONCE. reject-on-empty (NOT skip-on-empty): when a lesson re-points to a target
    # the shell can't prove - including an EMPTY inventory - that is itself an ERROR.
    if (shell_backed_reroutes or swaps) and shell is not None:
        try:
            if kind == schema.Kind.DOCX.value:
                present = _docx_style_keys(shell)
            elif kind == schema.Kind.PPTX.value:
                present = _pptx_layout_names(shell)
            elif kind == schema.Kind.XLSX.value:
                present = _xlsx_defined_names(shell)
            else:
                present = set()
            masks = (
                _xlsx_number_format_masks(shell)
                if kind == schema.Kind.XLSX.value
                else set()
            )
        except Exception as exc:  # opening the shell must never crash the gate
            findings.append(
                Finding(
                    "override_targets_exist",
                    schema.Severity.WARNING.value,
                    f"could not verify override targets against shell: {exc}",
                )
            )
            return findings

        for requested, target in shell_backed_reroutes.items():
            target_resolver = (roles.get(target) or {}).get("resolver") or {}
            target_present = (
                masks
                if (
                    kind == schema.Kind.XLSX.value
                    and target_resolver.get("type")
                    == schema.ResolverType.NUMBER_FORMAT.value
                )
                else present
            )
            if not _role_resolver_target_in_shell(
                kind, target_resolver, target_present
            ):
                findings.append(
                    Finding(
                        "override_targets_exist",
                        schema.Severity.ERROR.value,
                        f"reroute target {target!r} (for role {requested!r}) does not "
                        "resolve to a shell-backed op",
                    )
                )

        for rid, mask in swaps.items():
            # reject-on-empty: a swap into an empty mask inventory is itself an ERROR.
            if mask not in masks:
                findings.append(
                    Finding(
                        "override_targets_exist",
                        schema.Severity.ERROR.value,
                        f"number_format swap mask {mask!r} (for role {rid!r}) is not "
                        f"among the shell's used formats (have {sorted(masks)})",
                    )
                )

    return findings


def check_color_token_targets(profile: dict) -> list[Finding]:
    """FAIL-CLOSED membership of every COLOR token against ``theme.palette``.

    The sibling of :func:`check_comprehension_targets` for model-driven color: a
    color token is only valid when it is a verbatim key of ``theme.palette`` (the
    SOLE namespace a run color resolves off - never ``palette_roles``). Today the
    only place a comprehension references a palette key is
    ``comprehension.palette_annotations`` (the model NAMES a captured color); each
    key must be a real palette entry, mirroring ``check_membership``'s rule for
    anchor/index/region refs - a key into an EMPTY/absent palette is itself an ERROR
    (the model can never invent a color the deterministic capture did not observe).

    Model-free and deterministic. No-ops when the comprehension is absent (status !=
    present), so the model-free CI path and a pre-palette profile stay green.
    """
    comp = _present_comprehension(profile)
    if comp is None:
        return []
    palette = (profile.get("theme") or {}).get("palette") or {}
    palette_keys = set(palette)
    findings: list[Finding] = []
    for key in comp.get("palette_annotations") or {}:
        if key not in palette_keys:
            findings.append(
                Finding(
                    "color_token_targets_exist",
                    schema.Severity.ERROR.value,
                    f"color token {key!r} is not a key of theme.palette "
                    f"(have {sorted(palette_keys)})",
                )
            )
    return findings


def _present_comprehension(profile: dict) -> dict | None:
    comp = profile.get("comprehension")
    if (
        isinstance(comp, dict)
        and comp.get("status") == schema.ComprehensionStatus.PRESENT.value
    ):
        return comp
    return None


def captured_template_texts(
    profile: dict, *, include_surface_prompts: bool = False
) -> list[str]:
    """Collect every demo/placeholder text the extractor captured for THIS template.

    Model-free and language-agnostic: the comparison strings come from the
    template's own captured facts (cover ``demo_value`` s, surfaced demo-region
    markers/text), never a fixed phrase in any language. Surface placeholder
    prompts are opt-in because some formats keep master/layout prompts inside the
    package even when they are not visibly rendered; L0 must not fail on those.
    """
    texts: list[str] = []
    comp = _present_comprehension(profile)
    if comp is not None:
        for slot in (comp.get("cover_slots") or {}).values():
            if isinstance(slot, dict) and slot.get("demo_value"):
                texts.append(str(slot["demo_value"]))
    kind = profile.get("kind")
    sub = ((profile.get("surface") or {}).get(kind) or {}) if kind else {}
    if isinstance(sub, dict):
        demo = sub.get("demo_region") or {}
        if isinstance(demo, dict):
            for m in demo.get("instruction_markers") or []:
                if m:
                    texts.append(str(m))
            if demo.get("start_text"):
                texts.append(str(demo["start_text"]))
        if not include_surface_prompts:
            return _dedup_texts(texts)
        for anchor in sub.get("cover_anchors") or []:
            if not isinstance(anchor, dict):
                continue
            for key in ("placeholder", "demo_value"):
                value = anchor.get(key)
                if value:
                    texts.append(str(value))
    return _dedup_texts(texts)


def _dedup_texts(texts: list[str]) -> list[str]:
    # De-dup preserving order.
    seen: set[str] = set()
    out: list[str] = []
    for t in texts:
        if t and t not in seen:
            seen.add(t)
            out.append(t)
    return out


def _captured_demo_texts(profile: dict) -> list[str]:
    return captured_template_texts(profile)


def check_residual_template_text(text: str, profile: dict) -> list[Finding]:
    """Generalized ``no_residual_template_text`` for ALL kinds (model-free).

    Scans the produced text for any demo/placeholder string captured at extract
    for THIS template (cover ``demo_value`` s + surfaced demo markers/text). No
    hardcoded phrases, language-agnostic, identical across formats.
    """
    findings: list[Finding] = []
    for marker in _captured_demo_texts(profile):
        if marker in text:
            findings.append(
                Finding(
                    "no_residual_template_text",
                    schema.Severity.ERROR.value,
                    f"residual template text: {marker!r}",
                    location=marker,
                )
            )
    return findings


def check_no_orphan_cover_placeholder(text: str, profile: dict) -> list[Finding]:
    """No bound cover slot still shows its captured ``demo_value`` in the output.

    A slot the comprehension marked ``fill_rule='in_place'`` (i.e. it should have
    been filled) whose captured ``demo_value`` still appears in the produced text
    is an ERROR (stale demo prompt left behind). A slot the model intentionally
    re-armed (``fill_rule='clear'``) is INFO. No-ops when comprehension is absent.
    """
    comp = _present_comprehension(profile)
    if comp is None:
        return []
    findings: list[Finding] = []
    for anchor_ref, slot in (comp.get("cover_slots") or {}).items():
        if not isinstance(slot, dict):
            continue
        demo_value = slot.get("demo_value")
        if not demo_value or str(demo_value) not in text:
            continue
        fill_rule = slot.get("fill_rule")
        if fill_rule == schema.FillRule.IN_PLACE.value:
            findings.append(
                Finding(
                    "no_orphan_cover_placeholder",
                    schema.Severity.ERROR.value,
                    f"cover slot {anchor_ref!r} still shows placeholder {demo_value!r}",
                )
            )
        else:
            findings.append(
                Finding(
                    "no_orphan_cover_placeholder",
                    schema.Severity.INFO.value,
                    f"cover slot {anchor_ref!r} retains re-armed prompt {demo_value!r}",
                )
            )
    return findings


def check_index_matches_content(
    present_seq_ids: set[str], profile: dict
) -> list[Finding]:
    """No preserved CAPTION index lacks corresponding captionable content (WARNING).

    Ruling A: the comparison keys on the DETERMINISTIC ``seq_id`` captured opaquely
    from the index's ``\\c`` field switch (plan §7), NEVER on the advisory ``kind``
    token. ``present_seq_ids`` is the set of caption SEQ identifiers the generator
    actually emitted captions for (each is a verbatim ``seq_id`` it matched a
    preserved index against). A preserved CAPTION index (one carrying a non-null
    ``seq_id``) that is kept (``reconcile`` not ``clear``) yet whose ``seq_id`` was
    not emitted is flagged WARNING - a stale caption index the reconciliation should
    have cleared. An outline TOC (``seq_id`` is null) is never flagged: it is always
    refreshed, not content-matched. No-ops when comprehension is absent.
    Generator-driven (the generator knows which SEQ classes it emitted); model-free.
    """
    comp = _present_comprehension(profile)
    if comp is None:
        return []
    findings: list[Finding] = []
    for idx in (comp.get("conventions") or {}).get("indexes") or []:
        if not isinstance(idx, dict):
            continue
        if idx.get("reconcile") == schema.Reconcile.CLEAR.value:
            continue  # the reconciliation already removed it on purpose
        seq_id = idx.get("seq_id")
        # Only a CAPTION index (a \c-switched index, identified by its opaque
        # seq_id) is content-matched; a bare outline TOC has no seq_id and is
        # always refreshed, never flagged.
        if seq_id and seq_id not in present_seq_ids:
            findings.append(
                Finding(
                    "index_matches_content",
                    schema.Severity.WARNING.value,
                    f"preserved caption index {idx.get('index_ref')!r} "
                    f"(seq_id {seq_id!r}) has no matching captionable content",
                )
            )
    return findings


def check_no_net_structure_loss(
    removed_refs: set[str],
    profile: dict,
    *,
    confidence: float | None = None,
) -> list[Finding]:
    """The destructive-action floor (§6): no preserved anchor/index removed unless
    the deterministic layer independently classified it placeholder/demo.

    ``removed_refs`` is the set of anchor/index refs the reconciliation actually
    deleted (supplied by the generator). For each, the comprehension must carry a
    deterministically-corroborated destructive verdict (``fill_rule='clear'`` for a
    cover anchor, ``reconcile='clear'`` for an index, or a ``verdict='demo'`` region);
    otherwise the deletion is an ERROR. Model-free (it reads the frozen verdicts).

    ``confidence`` (optional) is the model's single ``comprehension.confidence``, the
    SAME value the reconcile sites gate on. When supplied, the backstop also
    re-verifies the destructive-action confidence floor: a sanctioned removal whose
    confidence does NOT clear the floor is an ERROR, because under the uniform policy
    such a removal should have been downgraded to KEEP at the reconcile site (a wrong
    delete is unrecoverable). When ``confidence`` is ``None`` the floor re-check is
    skipped (additive, back-compatible) and only the verdict-corroboration gate runs.
    """
    comp = _present_comprehension(profile)
    findings: list[Finding] = []
    cleared_anchors = (
        {
            ref
            for ref, slot in (comp.get("cover_slots") or {}).items()
            if isinstance(slot, dict)
            and slot.get("fill_rule") == schema.FillRule.CLEAR.value
        }
        if comp
        else set()
    )
    cleared_indexes = (
        {
            idx.get("index_ref")
            for idx in ((comp.get("conventions") or {}).get("indexes") or [])
            if isinstance(idx, dict)
            and idx.get("reconcile") == schema.Reconcile.CLEAR.value
        }
        if comp
        else set()
    )
    demo_regions = (
        {
            r.get("region_ref")
            for r in ((comp.get("demo_classification") or {}).get("regions") or [])
            if isinstance(r, dict) and r.get("verdict") == schema.Verdict.DEMO.value
        }
        if comp
        else set()
    )
    sanctioned = cleared_anchors | cleared_indexes | demo_regions
    floor_blocks = confidence is not None and not confidence_clears_floor(confidence)
    for ref in sorted(removed_refs):
        if ref not in sanctioned:
            findings.append(
                Finding(
                    "no_net_structure_loss",
                    schema.Severity.ERROR.value,
                    f"reconciliation removed {ref!r} without a corroborated "
                    f"destructive verdict",
                )
            )
        elif floor_blocks:
            # The verdict is sanctioned, but the model's confidence is below the
            # destructive floor: under the uniform policy the reconcile site should
            # have downgraded this to KEEP. A removal that still happened is a floor
            # breach (a wrong delete is unrecoverable), so surface it LOUDLY.
            findings.append(
                Finding(
                    "no_net_structure_loss",
                    schema.Severity.ERROR.value,
                    f"reconciliation removed {ref!r} below the destructive "
                    f"confidence floor ({confidence:.2f})",
                )
            )
    return findings


# ---------------------------------------------------------------------------
# Shell-vs-output structural diffs (catch silent corruption the text scans miss)
# ---------------------------------------------------------------------------
def _xlsx_formula_map(path) -> dict[str, str]:
    """Map ``Sheet!Coord`` -> formula string for every formula cell in a workbook.

    Reads the live file (``data_only=False``) so the comparison reflects what is
    actually on disk, not a possibly-stale catalog snapshot. A formula is identified
    by ``data_type == "f"`` (authoritative) - NOT by a leading ``=`` in the string,
    so a cell deliberately neutralized to TEXT (an author ``=...`` value written as a
    string literal) is correctly NOT counted as a formula. Iterating ``_cells`` keeps
    this O(populated) on sparse corporate models.
    """
    out: dict[str, str] = {}
    wb = load_workbook(path, data_only=False)
    for ws in wb.worksheets:
        for cell in ws._cells.values():
            if cell.data_type == "f" and isinstance(cell.value, str):
                out[f"{ws.title}!{cell.coordinate}"] = cell.value
    return out


def check_formula_preservation(shell, output, profile: dict) -> list[Finding]:
    """ERROR on any formula the generation LOST or MUTATED vs the brand shell.

    The brand guarantee for xlsx is that the shell's own formulas are preserved
    verbatim (only the shell's inputs are refilled, never its formulas). A region
    fill that overwrites a formula cell is silent data corruption that no text
    scan can see, so this diffs the shell's ``address->formula`` set against the
    output's and emits an ERROR for every entry that disappeared or changed.

    Model-free and deterministic. No-ops (returns ``[]``) when either file is
    absent (e.g. verify time, when there is no output yet) so it is always safe to
    call from the gate. The extractor's ``artifact_catalog.formulas`` is the same
    baseline; here we read the live shell so the check stands even if the catalog
    drifts.
    """
    if shell is None or output is None:
        return []
    if profile.get("kind") != schema.Kind.XLSX.value:
        return []
    try:
        shell_formulas = _xlsx_formula_map(shell)
        output_formulas = _xlsx_formula_map(output)
    except Exception as exc:  # opening a workbook must never crash the gate
        return [
            Finding(
                "formula_preservation",
                schema.Severity.WARNING.value,
                f"could not verify formula preservation: {exc}",
            )
        ]
    findings: list[Finding] = []
    for address in sorted(shell_formulas):
        shell_formula = shell_formulas[address]
        out_formula = output_formulas.get(address)
        if out_formula is None:
            findings.append(
                Finding(
                    "formula_preservation",
                    schema.Severity.ERROR.value,
                    f"shell formula at {address} ({shell_formula!r}) was erased in the output",
                    location=address,
                )
            )
        elif out_formula != shell_formula:
            findings.append(
                Finding(
                    "formula_preservation",
                    schema.Severity.ERROR.value,
                    f"shell formula at {address} was mutated: "
                    f"{shell_formula!r} -> {out_formula!r}",
                    location=address,
                )
            )
    # The engine NEVER authors formulas: any formula in the OUTPUT that the shell did
    # not have was introduced by author content (a string starting with '='), which is
    # both an invariant break and a formula-injection risk. Fail closed even if the
    # write-path neutralization ever regresses.
    for address in sorted(set(output_formulas) - set(shell_formulas)):
        findings.append(
            Finding(
                "formula_preservation",
                schema.Severity.ERROR.value,
                f"output has a formula at {address} ({output_formulas[address]!r}) that "
                "the shell did not: the engine never authors formulas (author '='-led "
                "content must be neutralized to text)",
                location=address,
            )
        )
    return findings


def _docx_component_counts(path) -> dict[str, int]:
    # Only count components the generator preserves rather than re-authors. Tables
    # are a meaningful survival signal (a flattened table is a real defect); raw
    # list-paragraph counts are NOT compared, because docx generation rewrites the
    # body from new content, so a different list length is expected, not a loss.
    doc = Document(path)
    return {"tables": len(doc.tables)}


def _pptx_component_counts(path) -> dict[str, int]:
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    prs = Presentation(path)
    tables = charts = pictures = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            if getattr(shape, "has_table", False):
                tables += 1
            if getattr(shape, "has_chart", False):
                charts += 1
            if getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.PICTURE:
                pictures += 1
    return {"tables": tables, "charts": charts, "pictures": pictures}


def _xlsx_component_counts(path) -> dict[str, int]:
    wb = load_workbook(path, data_only=False)
    tables = charts = 0
    for ws in wb.worksheets:
        try:
            tables += len(ws.tables)
        except Exception:
            pass
        try:
            charts += len(ws._charts)
        except Exception:
            pass
    return {"tables": tables, "charts": charts}


_COMPONENT_COUNTERS = {
    schema.Kind.DOCX.value: _docx_component_counts,
    schema.Kind.PPTX.value: _pptx_component_counts,
    schema.Kind.XLSX.value: _xlsx_component_counts,
}


def check_component_survival(shell, output, profile: dict) -> list[Finding]:
    """WARN when a native component present in the shell has no counterpart in the output.

    Catches the silent down-render class (e.g. pptx flattening a native table to
    text, or a docx list losing its numbering) that no text scan detects: it
    counts native components (tables/charts/lists/pictures, per format) in the
    shell and in the output and emits a WARNING for every component family whose
    output count dropped below the shell's. A WARNING (not ERROR) because some
    drops are legitimate (the new content genuinely has fewer of a component); the
    signal is "you may have lost a native object", surfaced rather than silent.

    Model-free, deterministic, no-ops when either file is absent.
    """
    if shell is None or output is None:
        return []
    counter = _COMPONENT_COUNTERS.get(profile.get("kind"))
    if counter is None:
        return []
    try:
        shell_counts = counter(shell)
        output_counts = counter(output)
    except Exception as exc:  # opening the package must never crash the gate
        return [
            Finding(
                "component_survival",
                schema.Severity.WARNING.value,
                f"could not verify component survival: {exc}",
            )
        ]
    findings: list[Finding] = []
    for family in sorted(shell_counts):
        before = shell_counts[family]
        after = output_counts.get(family, 0)
        if before > 0 and after < before:
            findings.append(
                Finding(
                    "component_survival",
                    schema.Severity.WARNING.value,
                    f"native {family} count dropped {before} -> {after} "
                    f"between shell and output (possible down-render)",
                    # Cluster C2: the dropped FAMILY is the addressable location, so
                    # two dropped families no longer collide on (check, None) and a
                    # model-assisted triage entry can name a single family precisely.
                    location=family,
                )
            )
    return findings


def check_docx(path, profile: dict, shell=None) -> list[Finding]:
    findings = check_profile(profile)
    doc = Document(path)
    text = "\n".join(
        [p.text for p in doc.paragraphs]
        + [cell.text for t in doc.tables for row in t.rows for cell in row.cells]
    )
    for hit in textutil.find_markdown_literals(text):
        findings.append(
            Finding(
                "no_literal_markdown",
                schema.Severity.ERROR.value,
                f"literal markdown leaked: {hit['match']!r}",
            )
        )
    findings.extend(check_residual_template_text(text, profile))
    findings.extend(check_no_orphan_cover_placeholder(text, profile))
    findings.extend(check_component_survival(shell, path, profile))
    return findings


def check_pptx(path, profile: dict, shell=None) -> list[Finding]:
    findings = check_profile(profile)
    prs = Presentation(path)
    text = "\n".join(
        shape.text
        for slide in prs.slides
        for shape in slide.shapes
        if hasattr(shape, "text")
    )
    for hit in textutil.find_markdown_literals(text):
        findings.append(
            Finding(
                "no_literal_markdown",
                schema.Severity.ERROR.value,
                f"literal markdown leaked: {hit['match']!r}",
            )
        )
    # Uniform, model-free residual check: compare against THIS template's captured
    # demo/placeholder text, never a fixed phrase (de-literalized).
    findings.extend(check_residual_template_text(text, profile))
    findings.extend(check_no_orphan_cover_placeholder(text, profile))
    findings.extend(check_component_survival(shell, path, profile))
    return findings


def check_xlsx(path, profile: dict, shell=None) -> list[Finding]:
    findings = check_profile(profile)
    wb = load_workbook(path, data_only=False)
    cell_texts: list[str] = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    cell_texts.append(cell.value)
                    for hit in textutil.find_markdown_literals(cell.value):
                        findings.append(
                            Finding(
                                "no_literal_markdown",
                                schema.Severity.ERROR.value,
                                f"literal markdown leaked: {hit['match']!r}",
                            )
                        )
    # Uniform, model-free residual check across all kinds.
    text = "\n".join(cell_texts)
    findings.extend(check_residual_template_text(text, profile))
    findings.extend(check_no_orphan_cover_placeholder(text, profile))
    # Deterministic structural diffs that the text scan above is blind to: a fill
    # that erased a shell formula, or a native component lost in the output.
    findings.extend(check_formula_preservation(shell, path, profile))
    findings.extend(check_component_survival(shell, path, profile))
    return findings
