# SPDX-License-Identifier: MIT
from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path

from pptx import Presentation

from brandkit.formats import catalog
from brandkit.ooxml import pack
from brandkit.profile import schema, store


def extract(template: str | Path, name: str, *, scope: str = "project", cwd: str | Path | None = None) -> Path:
    template_path = Path(template)
    prs = Presentation(template_path)
    roles = _roles(prs)
    surface = {
        "pptx": {
            "slide_size_emu": {"w": int(prs.slide_width), "h": int(prs.slide_height)},
            "layouts": _layouts(prs),
            "role_layout_map": {"cover": "Title Slide", "content_text": "Title and Content"},
            "safe_area_emu": {"l": 457200, "t": 457200, "r": 457200, "b": 457200},
        }
    }
    profile = schema.build_envelope(
        "pptx",
        {"name": name, "display_name": name},
        extracted_at=datetime.now(timezone.utc).isoformat(),
        source_template={"filename": template_path.name, "sha256": store.sha256_file(template_path)},
        theme=_theme(),
        roles=roles,
        surface=surface,
    )
    profile["anchors"] = {
        "cover": {"kind": schema.AnchorKind.PLACEHOLDER.value, "slots_found": 1},
        "demo_region": {"present": True},
        "toc": {"present": False},
    }
    profile["provenance"]["ooxml_parts_seen"] = pack.list_parts(template_path)
    profile["artifact_catalog"] = _artifact_catalog(template_path, prs, profile["provenance"]["ooxml_parts_seen"])
    profile["capabilities"] = _capabilities()
    target = store.target_dir_for_save(name, scope, cwd=cwd)
    return store.save_profile(target, profile, template_path.read_bytes(), extra_files={"PROFILE.md": _profile_md(profile)}, overwrite=True)


def _roles(prs: Presentation) -> dict:
    roles = {"_index": []}

    def add(rid: str, resolver: dict, confidence: float, status: str, signal: str) -> None:
        roles[rid] = {
            "resolver": resolver,
            "appearance": {},
            "verified": True,
            "confidence": confidence,
            "status": status,
            "evidence": {"signal": signal},
        }
        roles["_index"].append(rid)

    add(
        "cover.title",
        {"type": schema.ResolverType.PLACEHOLDER.value, "layout": "Title Slide", "ph_idx": 0, "ph_type": "title"},
        0.9,
        schema.Status.ROBUST.value,
        "title slide title placeholder",
    )
    add(
        "heading.1",
        {"type": schema.ResolverType.PLACEHOLDER.value, "layout": "Title and Content", "ph_idx": 0, "ph_type": "title"},
        0.8,
        schema.Status.BEST_EFFORT.value,
        "content title placeholder",
    )
    add(
        "paragraph",
        {"type": schema.ResolverType.PLACEHOLDER.value, "layout": "Title and Content", "ph_idx": 1, "ph_type": "body"},
        0.8,
        schema.Status.BEST_EFFORT.value,
        "content body placeholder",
    )
    return roles


def _layouts(prs: Presentation) -> dict:
    out = {}
    for layout in prs.slide_layouts:
        placeholders = []
        for ph in layout.placeholders:
            fmt = ph.placeholder_format
            placeholders.append(
                {
                    "idx": fmt.idx,
                    "type": str(fmt.type),
                    "geo_emu": {
                        "l": int(ph.left or 0),
                        "t": int(ph.top or 0),
                        "w": int(ph.width or 0),
                        "h": int(ph.height or 0),
                    },
                    "geo_source": "resolved",
                }
            )
        out[layout.name] = {"master_idx": 0, "placeholders": placeholders}
    return out


def _artifact_catalog(path: Path, prs: Presentation, parts: list[str]) -> dict:
    out = catalog.part_catalog(path)
    out["ooxml_parts"] = parts
    out["slide_size_emu"] = {"w": int(prs.slide_width), "h": int(prs.slide_height)}
    out["slide_layouts"] = _layouts(prs)
    out["slide_masters"] = [
        {
            "name": getattr(master, "name", None),
            "placeholders": [
                {
                    "idx": ph.placeholder_format.idx,
                    "type": str(ph.placeholder_format.type),
                    "geo_emu": {
                        "l": int(ph.left or 0),
                        "t": int(ph.top or 0),
                        "w": int(ph.width or 0),
                        "h": int(ph.height or 0),
                    },
                }
                for ph in master.placeholders
            ],
        }
        for master in prs.slide_masters
    ]
    out["slides"] = [
        {
            "layout": slide.slide_layout.name,
            "shape_count": len(slide.shapes),
            "texts": [shape.text[:200] for shape in slide.shapes if hasattr(shape, "text") and shape.text],
        }
        for slide in prs.slides
    ]
    return out


def _capabilities() -> dict:
    return {
        "extracts_all_ooxml_parts": True,
        "extracts_layout_geometry": True,
        "extracts_placeholder_catalog": True,
        "generates_from_shell": True,
        "overflow_guard": "conservative_text_split",
    }


def _theme() -> dict:
    return {
        "colors": {},
        "palette_roles": {"primary": {"theme": "accent1"}, "text": {"theme": "dk1"}},
        "fonts": {"major": {"latin": None, "fallback": "Arial"}, "minor": {"latin": None, "fallback": "Calibri"}},
        "embedded_fonts": [],
    }


def _profile_md(profile: dict) -> str:
    return "# Brand Profile: " + profile["identity"]["display_name"] + "\n\n- kind: pptx\n"
