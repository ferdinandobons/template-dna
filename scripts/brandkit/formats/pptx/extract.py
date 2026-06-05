# SPDX-License-Identifier: MIT
from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER

from brandkit.formats import catalog
from brandkit.ooxml import pack
from brandkit.profile import schema, store


def extract(template: str | Path, name: str, *, scope: str = "project", cwd: str | Path | None = None) -> Path:
    template_path = Path(template)
    prs = Presentation(template_path)
    layouts = _layouts(prs)
    roles = _roles(prs, layouts)
    surface = {
        "pptx": {
            "slide_size_emu": {"w": int(prs.slide_width), "h": int(prs.slide_height)},
            "layouts": layouts,
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
    cover_role = roles.get("cover.title") if isinstance(roles, dict) else None
    cover_present = isinstance(cover_role, dict) and cover_role.get("status") != schema.Status.STUB.value
    profile["anchors"] = {
        "cover": {
            "kind": schema.AnchorKind.PLACEHOLDER.value if cover_present else schema.AnchorKind.NONE.value,
            "slots_found": 1 if cover_present else 0,
        },
        "demo_region": {"present": bool(prs.slides)},
        "toc": {"present": False},
    }
    profile["provenance"]["ooxml_parts_seen"] = pack.list_parts(template_path)
    profile["artifact_catalog"] = _artifact_catalog(template_path, prs, profile["provenance"]["ooxml_parts_seen"], layouts)
    profile["capabilities"] = _capabilities()
    target = store.target_dir_for_save(name, scope, cwd=cwd)
    return store.save_profile(target, profile, template_path.read_bytes(), extra_files={"PROFILE.md": _profile_md(profile)}, overwrite=True)


# Placeholder-type families used to classify real layouts (PP_PLACEHOLDER, §C3).
# A *title* slot is any of TITLE/CENTER_TITLE/VERTICAL_TITLE; a *body* slot is any
# text-bearing content placeholder (BODY/OBJECT/SUBTITLE/...). We never invent a
# layout: every resolver here points at a layout.name that prs actually parsed.
_TITLE_TYPES = frozenset({
    PP_PLACEHOLDER.TITLE,
    PP_PLACEHOLDER.CENTER_TITLE,
    PP_PLACEHOLDER.VERTICAL_TITLE,
})
_SUBTITLE_TYPES = frozenset({PP_PLACEHOLDER.SUBTITLE})
_BODY_TYPES = frozenset({
    PP_PLACEHOLDER.BODY,
    PP_PLACEHOLDER.OBJECT,
    PP_PLACEHOLDER.VERTICAL_BODY,
    PP_PLACEHOLDER.VERTICAL_OBJECT,
})


def _classify_layouts(prs: Presentation) -> list[dict]:
    """Describe each REAL slide layout by the slots it actually exposes.

    Returns a list (in deck order) of ``{name, idx, title_idx, subtitle_idx,
    body_idx}`` where each ``*_idx`` is the placeholder ``idx`` of the first slot
    of that family present in the layout, or ``None``. This is the only source of
    truth for role derivation — nothing here is fabricated.
    """
    described: list[dict] = []
    for pos, layout in enumerate(prs.slide_layouts):
        title_idx = subtitle_idx = body_idx = None
        for ph in layout.placeholders:
            fmt = ph.placeholder_format
            ptype = fmt.type
            if ptype in _TITLE_TYPES and title_idx is None:
                title_idx = fmt.idx
            elif ptype in _SUBTITLE_TYPES and subtitle_idx is None:
                subtitle_idx = fmt.idx
            elif ptype in _BODY_TYPES and body_idx is None:
                body_idx = fmt.idx
        described.append({
            "name": layout.name,
            "idx": pos,
            "title_idx": title_idx,
            "subtitle_idx": subtitle_idx,
            "body_idx": body_idx,
        })
    return described


def _pick_cover(described: list[dict]) -> dict | None:
    """Pick the layout that best reads as a cover/title slide.

    Strongest signal: a title slot paired with a subtitle slot (the canonical
    cover shape). Falls back to the first layout exposing any title slot.
    """
    for d in described:
        if d["title_idx"] is not None and d["subtitle_idx"] is not None:
            return d
    for d in described:
        if d["title_idx"] is not None:
            return d
    return None


def _pick_content(described: list[dict], *, exclude_idx: int | None = None) -> dict | None:
    """Pick the layout that best reads as a title+body content slide.

    Prefers a layout with BOTH a title and a body slot, skipping ``exclude_idx``
    (the cover) when an alternative exists. Falls back to any title-bearing
    layout, then any body-bearing layout.
    """
    title_body = [d for d in described if d["title_idx"] is not None and d["body_idx"] is not None]
    for d in title_body:
        if d["idx"] != exclude_idx:
            return d
    if title_body:
        return title_body[0]
    for d in described:
        if d["title_idx"] is not None:
            return d
    for d in described:
        if d["body_idx"] is not None:
            return d
    return None


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
