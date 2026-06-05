# SPDX-License-Identifier: MIT
from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path

from pptx import Presentation

from brandkit.formats import catalog
from brandkit.formats.pptx import structure
from brandkit.ooxml import pack
from brandkit.profile import schema, store


def extract(template: str | Path, name: str, *, scope: str = "project", cwd: str | Path | None = None) -> Path:
    template_path = Path(template)
    prs = Presentation(template_path)
    layouts = _layouts(prs)
    roles = _roles(prs, layouts)
    # Format-uniform comprehension inventories (schema 1.2.0, plan §4 / M-i-7).
    # Every load-bearing ref the model writes binds to one of these ids; the
    # validator checks membership. They are the PPTX peers of the docx
    # cover_anchors / fields / regions:
    #   - cover_anchors: every placeholder on the cover layout (multi-placeholder
    #     cover), keyed by layout+ph idx + captured layout prompt as the demo value;
    #   - fields: the deck's real section list as an agenda/section-list index
    #     (empty when the deck has no p14:sectionLst);
    #   - regions: every slide classified cover / structural / demo (a demo slide is
    #     one whose body text equals a layout placeholder prompt).
    # An absent comprehension never reads them, so the deterministic path is
    # unaffected; a deck whose inventory is genuinely empty (e.g. no cover layout)
    # surfaces an empty list and a comprehension ref into it is fail-closed at QA.
    cover_anchors = structure.inventory_cover_anchors(prs)
    fields = structure.inventory_fields(prs)
    regions = structure.inventory_regions(prs)
    sections = structure.detect_sections(prs)
    skeleton = structure.detect_skeleton(prs)
    surface = {
        "pptx": {
            "slide_size_emu": {"w": int(prs.slide_width), "h": int(prs.slide_height)},
            "layouts": layouts,
            "safe_area_emu": {"l": 457200, "t": 457200, "r": 457200, "b": 457200},
            "cover_anchors": cover_anchors,
            "fields": fields,
            "regions": regions,
            "sections": sections,
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
        structure=skeleton,
    )
    # Anchors track reality: the cover slot count is the real multi-placeholder
    # count (never a hardcoded ``1``); the demo region is present only when a slide
    # actually reads as demo (NOT ``bool(prs.slides)``); the section list is present
    # only when the deck carries one (NOT always-False).
    demo_present = any(r.get("kind") == "demo" for r in regions)
    profile["anchors"] = {
        "cover": {
            "kind": schema.AnchorKind.PLACEHOLDER.value if cover_anchors else schema.AnchorKind.NONE.value,
            "slots_found": len(cover_anchors),
        },
        "demo_region": {"present": demo_present},
        "sections": {"present": bool(sections)},
    }
    profile["provenance"]["ooxml_parts_seen"] = pack.list_parts(template_path)
    profile["artifact_catalog"] = _artifact_catalog(template_path, prs, profile["provenance"]["ooxml_parts_seen"], layouts)
    profile["capabilities"] = _capabilities()
    target = store.target_dir_for_save(name, scope, cwd=cwd)
    return store.save_profile(target, profile, template_path.read_bytes(), extra_files={"PROFILE.md": _profile_md(profile)}, overwrite=True)


# Layout classification, cover/content picking, and placeholder-type families now
# live in the structure peer (one source of truth, plan §4 / M-i-8). The thin
# re-exports below keep the role-derivation code below (and any caller) reading the
# same descriptors the inventory uses, so role layouts and cover anchors can never
# disagree about which layout is the cover.
_TITLE_TYPES = structure.TITLE_TYPES
_SUBTITLE_TYPES = structure.SUBTITLE_TYPES
_BODY_TYPES = structure.BODY_TYPES
_classify_layouts = structure.classify_layouts
_pick_cover = structure.pick_cover
_pick_content = structure.pick_content


def _roles(prs: Presentation, layouts: dict) -> dict:
    """Derive pptx roles from the REAL parsed layouts (§C3).

    Every emitted resolver points at an actual ``layout.name`` and an actual
    placeholder ``ph_idx`` proven to exist in that layout. When no suitable
    layout is found, the role is emitted as a ``stub`` with honest (low)
    confidence and an empty resolver target rather than a confident fiction.
    """
    described = _classify_layouts(prs)
    roles: dict = {"_index": []}

    def add(rid: str, resolver: dict, confidence: float, status: str, signal: str) -> None:
        roles[rid] = {
            "resolver": resolver,
            "appearance": {},
            "verified": status != schema.Status.STUB.value,
            "confidence": confidence,
            "status": status,
            "evidence": {"signal": signal},
        }
        roles["_index"].append(rid)

    def stub_resolver() -> dict:
        # A stub names no layout, so the intra-profile consistency check (which
        # only fires when ``layout`` is non-null) cannot be tripped by a fiction.
        return {"type": schema.ResolverType.PLACEHOLDER.value, "layout": None, "ph_idx": None, "ph_type": None}

    cover = _pick_cover(described)
    if cover is not None:
        add(
            "cover.title",
            {
                "type": schema.ResolverType.PLACEHOLDER.value,
                "layout": cover["name"],
                "ph_idx": cover["title_idx"],
                "ph_type": "title",
            },
            0.9 if cover["subtitle_idx"] is not None else 0.7,
            schema.Status.ROBUST.value if cover["subtitle_idx"] is not None else schema.Status.BEST_EFFORT.value,
            f"layout {cover['name']!r} title placeholder (idx {cover['title_idx']})",
        )
    else:
        add(
            "cover.title",
            stub_resolver(),
            0.0,
            schema.Status.STUB.value,
            "no layout in this template exposes a title placeholder",
        )

    cover_idx = cover["idx"] if cover is not None else None
    content = _pick_content(described, exclude_idx=cover_idx)

    heading_layout = content if (content and content["title_idx"] is not None) else None
    if heading_layout is not None:
        add(
            "heading.1",
            {
                "type": schema.ResolverType.PLACEHOLDER.value,
                "layout": heading_layout["name"],
                "ph_idx": heading_layout["title_idx"],
                "ph_type": "title",
            },
            0.8,
            schema.Status.BEST_EFFORT.value,
            f"layout {heading_layout['name']!r} title placeholder (idx {heading_layout['title_idx']})",
        )
    else:
        add(
            "heading.1",
            stub_resolver(),
            0.0,
            schema.Status.STUB.value,
            "no content layout in this template exposes a title placeholder",
        )

    body_layout = content if (content and content["body_idx"] is not None) else None
    if body_layout is not None:
        add(
            "paragraph",
            {
                "type": schema.ResolverType.PLACEHOLDER.value,
                "layout": body_layout["name"],
                "ph_idx": body_layout["body_idx"],
                "ph_type": "body",
            },
            0.8,
            schema.Status.BEST_EFFORT.value,
            f"layout {body_layout['name']!r} body placeholder (idx {body_layout['body_idx']})",
        )
    else:
        add(
            "paragraph",
            stub_resolver(),
            0.0,
            schema.Status.STUB.value,
            "no content layout in this template exposes a body placeholder",
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


def _artifact_catalog(path: Path, prs: Presentation, parts: list[str], layouts: dict) -> dict:
    out = catalog.part_catalog(path)
    out["ooxml_parts"] = parts
    out["slide_size_emu"] = {"w": int(prs.slide_width), "h": int(prs.slide_height)}
    out["slide_layouts"] = layouts
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
    # Typed native-component inventory per slide (table/chart/picture counts), the
    # baseline the component-survival check diffs against the output deck so a
    # down-rendered native object is caught deterministically (plan P5).
    components = {c["index"]: c["components"] for c in structure.slide_component_inventory(prs)}
    out["slides"] = [
        {
            "layout": slide.slide_layout.name,
            "shape_count": len(slide.shapes),
            "texts": [shape.text[:200] for shape in slide.shapes if hasattr(shape, "text") and shape.text],
            "components": components.get(i, {"table": 0, "chart": 0, "picture": 0}),
        }
        for i, slide in enumerate(prs.slides)
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
