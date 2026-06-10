# SPDX-License-Identifier: MIT
"""DOCX extraction for M1."""

from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path

from docx import Document

from brandkit.common import color, profilemd
from brandkit.formats import catalog
from brandkit.formats.docx import cover, roles, structure, typography
from brandkit.ooxml import names, pack
from brandkit.profile import schema, store


def extract(
    template: str | Path,
    name: str,
    *,
    scope: str = "project",
    cwd: str | Path | None = None,
) -> Path:
    template_path = Path(template)
    shell_bytes = template_path.read_bytes()
    doc = Document(template_path)

    theme = _extract_theme(template_path)
    role_registry = roles.infer_roles(doc)
    # One materialized direct-run facts pass shared by capture_fonts and
    # capture_pseudo_headings (which scans twice): pure recomputation removal,
    # identical items in identical order, byte-identical profile.
    run_facts = typography.collect_font_run_facts(doc)
    # Capture the template's REAL visible fonts (often direct run-level overrides
    # the named styles/theme never carry) into role.appearance + theme.fonts.body.
    # Additive and deterministic: a template with no dominant direct font is a no-op.
    typography.capture_fonts(doc, role_registry, theme, facts=run_facts)
    # Detect faked-heading-in-body-style candidates (Cluster E2): a body-style run
    # whose captured size/color is a clear OUTLIER vs the just-captured dominant body
    # appearance is surfaced as a ``theme.pseudo_headings`` fact for the model to
    # adjudicate (promote onto a real heading role via comprehension). DOCX-FIRST and
    # additive: a uniform body leaves it absent (no-op, byte-identical comprehend bundle
    # + generation). MUST run after capture_fonts so the body dominant exists.
    typography.capture_pseudo_headings(doc, role_registry, theme, facts=run_facts)
    # Capture the brand PALETTE (theme.palette): the UNDERSTAND half of model-driven
    # color, built from observed facts only (theme slots, run colors incl. a low-floor
    # accent aggregation, per-role appearance colors, link colors). Additive and
    # deterministic; a template with no observed color leaves an empty palette.
    typography.capture_palette(doc, role_registry, theme)
    # Capture the template's dominant paragraph GEOMETRY (spacing/indent/borders/
    # shading from the runs' own w:pPr) into role.appearance.geometry + theme.geometry.
    # DOCX-ONLY (Cluster D1): WordprocessingML pPr has no pptx/xlsx peer. Additive and
    # deterministic - a template with no dominant geometry leaves it absent (no-op).
    typography.capture_geometry(doc, role_registry, theme)
    # Capture the template's dominant TABLE conditional-format facts (the w:tblLook
    # bitmask, the referenced table-style id, and the w:tblCellMar cell margins from the
    # template's OWN tables' w:tblPr) into role.appearance.table + theme.table.body.
    # DOCX-ONLY (Cluster D2): the bitmask only ENABLES the shell style's own
    # w:tblStylePr banding/first-last emphasis; the engine never synthesizes a fill.
    # Additive and deterministic - a template with no dominant table fact leaves it
    # absent (no-op, byte-identical generation).
    typography.capture_table_appearance(doc, role_registry, theme)
    cover_anchors, anchors = cover.discover_cover(doc)
    demo_region = structure.detect_demo_region(doc)
    toc_present = structure.is_toc_present(doc)
    skeleton = structure.detect_skeleton(doc, cover_anchors)
    # Format-uniform inventories the model reasons over and the validator binds to
    # (plan §4): every TOC/index complex field (opaque ``\c`` seq_id) and every
    # region. These back ``comprehension`` refs; an absent comprehension never
    # reads them, so the deterministic path is unaffected.
    fields = structure.inventory_fields(doc)
    regions = structure.inventory_regions(doc)
    anchors["demo_region"] = {"present": bool(demo_region.get("present"))}
    anchors["toc"] = {"present": bool(toc_present)}

    surface = {
        "docx": {
            "styles": _style_names(doc),
            "cover_anchors": cover_anchors,
            "fields": fields,
            "regions": regions,
            "demo_region": demo_region,
        }
    }
    profile = schema.build_envelope(
        "docx",
        {"name": name, "display_name": name},
        extracted_at=datetime.now(timezone.utc).isoformat(),
        source_template={
            "filename": template_path.name,
            "sha256": store.sha256_file(template_path),
        },
        theme=theme,
        roles=role_registry,
        surface=surface,
        structure=skeleton,
    )
    profile["anchors"] = anchors
    profile["provenance"]["ooxml_parts_seen"] = pack.list_parts(template_path)
    profile["artifact_catalog"] = _artifact_catalog(
        template_path, doc, profile["provenance"]["ooxml_parts_seen"]
    )
    profile["capabilities"] = _capabilities()

    target = store.target_dir_for_save(name, scope, cwd=cwd)
    extra = {"PROFILE.md": _profile_markdown(profile)}
    return store.save_profile(
        target, profile, shell_bytes, extra_files=extra, overwrite=True
    )


def _extract_theme(path: Path) -> dict:
    # A missing theme part is legitimate (KeyError from the zip); treat that as
    # "no theme colors". Any other failure (a corrupt/malicious theme1.xml, an
    # invalid package) is a real error and must NOT silently blank the palette,
    # so it propagates instead of being swallowed.
    colors = {}
    try:
        colors = color.parse_theme_colors(pack.read_part(path, "word/theme/theme1.xml"))
    except KeyError:
        colors = {}
    return {
        "colors": {slot: {"hex": hex_value} for slot, hex_value in colors.items()},
        "palette_roles": {
            "primary": {"theme": "accent1"},
            "text": {"theme": "dk1"},
            "surface": {"theme": "lt2"},
            "danger": {"theme": "accent4"},
        },
        "fonts": _extract_theme_fonts(path),
        "embedded_fonts": [],
    }


# DrawingML + WordprocessingML namespace-bound qualifiers for the font reader.
_A = names.make_qn("a")
_W = names.make_qn("w")


def _major_minor_latin(path: Path) -> tuple[str | None, str | None]:
    """Read the major/minor latin typefaces from ``word/theme/theme1.xml``.

    Returns ``(major, minor)`` where each is the
    ``a:fontScheme/a:majorFont|a:minorFont/a:latin@typeface`` value, or ``None``
    when the part, the element, or the attribute is missing - or when the
    typeface is the empty string (Office writes ``typeface=""`` to mean "none").
    A missing theme part is the only swallowed error (KeyError), preserving the
    extractor's contract; any other parse error propagates. The parser is the
    hardened OOXML one (``resolve_entities=False``), so a malicious theme cannot
    pull in external entities.
    """
    try:
        xml = pack.read_part(path, "word/theme/theme1.xml")
    except KeyError:
        return None, None
    root = pack.parse_xml_bytes(xml)
    scheme = root.find(f".//{_A('fontScheme')}")
    if scheme is None:
        return None, None

    def latin_of(font_tag: str) -> str | None:
        font = scheme.find(_A(font_tag))
        if font is None:
            return None
        latin = font.find(_A("latin"))
        if latin is None:
            return None
        # An empty ``typeface=""`` is Office's "no face here" sentinel, not a font.
        return latin.get("typeface") or None

    return latin_of("majorFont"), latin_of("minorFont")


def _doc_default_ascii(path: Path) -> str | None:
    """Read the document-default ascii body face from ``word/styles.xml``.

    Returns ``docDefaults/rPrDefault/rPr/rFonts@w:ascii`` (else ``@w:hAnsi``), or
    ``None`` when any link in that chain is absent - including the common case
    where the default ``rFonts`` carries only a THEME reference
    (``w:asciiTheme``) and no literal face. A missing styles part is swallowed
    (KeyError); any other parse error propagates. Hardened parser, as above.
    """
    try:
        xml = pack.read_part(path, "word/styles.xml")
    except KeyError:
        return None
    root = pack.parse_xml_bytes(xml)
    rfonts = root.find(
        f"{_W('docDefaults')}/{_W('rPrDefault')}/{_W('rPr')}/{_W('rFonts')}"
    )
    if rfonts is None:
        return None
    return rfonts.get(_W("ascii")) or rfonts.get(_W("hAnsi")) or None


def _extract_theme_fonts(path: Path) -> dict:
    """The template's TRUTHFUL theme fonts (major/minor), read from the package.

    The latin faces are the theme's own ``a:majorFont``/``a:minorFont`` latin
    typefaces; the fallback is the document-default ascii body face. The major
    fallback keeps ``'Arial'`` (Word's universal baseline, NOT a brand value)
    when no doc-default ascii is declared; the minor fallback is the real
    doc-default ascii (``None`` when absent). These declarations are read-only
    here: the resolver never reads major/minor - they only widen the allow-set
    that ``check_appearance_targets`` validates applied fonts against.
    """
    major_latin, minor_latin = _major_minor_latin(path)
    doc_ascii = _doc_default_ascii(path)
    return {
        "major": {"latin": major_latin, "fallback": doc_ascii or "Arial"},
        "minor": {"latin": minor_latin, "fallback": doc_ascii},
    }


def _style_names(doc) -> dict:
    paragraph, table, numbering = [], [], []
    for style in doc.styles:
        stype = str(style.type)
        if "PARAGRAPH" in stype:
            paragraph.append(style.name)
        elif "TABLE" in stype:
            table.append(style.name)
    return {"paragraph": paragraph, "table": table, "numbering": numbering}


def _artifact_catalog(path: Path, doc, parts: list[str]) -> dict:
    out = catalog.part_catalog(path)
    out["ooxml_parts"] = parts
    out["styles"] = _style_names(doc)
    out["style_details"] = _style_details(doc)
    # ``structure.section_length_emu`` is robust to malformed twips attributes
    # (some editors emit non-integer twips that make python-docx raise on access).
    out["sections"] = [
        {
            "page_width_emu": structure.section_length_emu(section, "page_width"),
            "page_height_emu": structure.section_length_emu(section, "page_height"),
            "margins_emu": {
                "top": structure.section_length_emu(section, "top_margin"),
                "right": structure.section_length_emu(section, "right_margin"),
                "bottom": structure.section_length_emu(section, "bottom_margin"),
                "left": structure.section_length_emu(section, "left_margin"),
            },
        }
        for section in doc.sections
    ]
    out["paragraph_samples"] = [
        {"style": para.style.name if para.style else None, "text": para.text[:200]}
        for para in doc.paragraphs[:25]
        if para.text
    ]
    out["table_count"] = len(doc.tables)
    return out


def _style_details(doc) -> list[dict]:
    details = []
    for style in doc.styles:
        font = getattr(style, "font", None)
        details.append(
            {
                "style_id": getattr(style, "style_id", None),
                "name": style.name,
                "type": str(style.type),
                "based_on": style.base_style.name
                if getattr(style, "base_style", None)
                else None,
                "font": {
                    "name": getattr(font, "name", None),
                    "size_pt": (font.size.pt if getattr(font, "size", None) else None),
                    "bold": getattr(font, "bold", None),
                    "italic": getattr(font, "italic", None),
                },
            }
        )
    return details


def _capabilities() -> dict:
    return {
        "extracts_all_ooxml_parts": True,
        "extracts_style_catalog": True,
        "extracts_layout_geometry": True,
        "generates_from_shell": True,
        "visual_overflow_requires_render": True,
        "native_charts": True,
        "native_smartart": True,
    }


def _profile_markdown(profile: dict) -> str:
    lines = [
        f"# Brand Profile: {profile['identity']['display_name']}",
        "",
        f"- kind: {profile['kind']}",
        f"- verification: {profile['verification']['status']}",
    ]
    lines.extend(_structure_markdown(profile))
    lines.extend(profilemd.roles_md(profile))
    lines.extend(profilemd.palette_roles_md(profile))
    lines.extend(profilemd.authoring_hints_md(profile))
    return "\n".join(lines) + "\n"


def _structure_markdown(profile: dict) -> list[str]:
    structure = profile.get("structure") or {}
    skeleton = structure.get("skeleton") or []
    if not skeleton:
        return []
    ordered = structure.get("ordered", True)
    lines = ["", "## Structure", ""]
    lines.append(
        "The template's ordered top-level skeleton. "
        + (
            "Region order **must** be respected on generation."
            if ordered
            else "Region order is informational."
        )
    )
    lines.append("")
    for region in sorted(
        skeleton, key=lambda r: r.get("order") if r.get("order") is not None else 0
    ):
        flags = []
        if region.get("required"):
            flags.append("required")
        if region.get("repeatable"):
            flags.append("repeatable")
        if region.get("freeform"):
            flags.append("freeform")
        flag_str = (" · " + " · ".join(flags)) if flags else ""
        lines.append(
            f"{region.get('order')}. **{region.get('region')}** "
            f"(`{region.get('role')}`){flag_str} - {region.get('evidence', '')}"
        )
    return lines
