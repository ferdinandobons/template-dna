# SPDX-License-Identifier: MIT
"""DOCX extraction for M1."""

from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path

from docx import Document

from brandkit.common import color
from brandkit.formats import catalog
from brandkit.formats.docx import cover, roles, structure
from brandkit.ooxml import pack
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
        "fonts": {
            "major": {"latin": None, "fallback": "Arial"},
            "minor": {"latin": None, "fallback": "Calibri"},
        },
        "embedded_fonts": [],
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
    lines.extend(["", "## Roles", ""])
    lines.append(
        "Each role lists its concrete style and its usage (scope · placement · required)."
    )
    lines.append(
        "`structural` roles belong to the ordered skeleton and must appear in their slot;"
    )
    lines.append("`freeform` roles are used on demand inside the freeform body region.")
    lines.append("")
    for rid in profile.get("roles", {}).get("_index", []):
        entry = profile["roles"][rid]
        resolver = entry.get("resolver", {})
        style = resolver.get("style_name") or resolver.get("style_id")
        usage = entry.get("usage") or {}
        usage_str = _usage_str(usage)
        lines.append(f"- `{rid}`: {style} ({entry.get('status')}) - {usage_str}")
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


def _usage_str(usage: dict) -> str:
    if not usage:
        return "usage: n/a"
    scope = usage.get("scope", "?")
    placement = usage.get("placement", "?")
    required = "required" if usage.get("required") else "optional"
    order = usage.get("order")
    order_str = f" · order={order}" if order is not None else ""
    return f"scope={scope} · {placement} · {required}{order_str}"
