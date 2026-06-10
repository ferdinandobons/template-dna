# SPDX-License-Identifier: MIT
"""Shared PROFILE.md sections - the human/agent-facing authoring surface.

PROFILE.md is written at extract time next to profile.json and is the page an
authoring agent reads BEFORE building an IntermediateDocument/GridDocument. The
sections here are format-uniform so the three extract writers never drift:
the role table, the semantic palette-role table, and the authoring hints that
teach role-first (never style/hex-first) composition. Everything rendered is
derived from the profile dict alone - deterministic, no model, no IO.
"""

from __future__ import annotations


def usage_str(usage: dict) -> str:
    """One-line usage summary for a role (scope, placement, required, order)."""
    if not usage:
        return "usage: n/a"
    scope = usage.get("scope", "?")
    placement = usage.get("placement", "?")
    required = "required" if usage.get("required") else "optional"
    order = usage.get("order")
    order_str = f" · order={order}" if order is not None else ""
    return f"scope={scope} · {placement} · {required}{order_str}"


def roles_md(profile: dict) -> list[str]:
    """The role table: every role id with its concrete artifact and usage."""
    index = (profile.get("roles") or {}).get("_index") or []
    if not index:
        return []
    lines = ["", "## Roles", ""]
    lines.append(
        "Each role lists its concrete style and its usage (scope · placement · required)."
    )
    lines.append(
        "`structural` roles belong to the ordered skeleton and must appear in their slot;"
    )
    lines.append("`freeform` roles are used on demand inside the freeform body region.")
    lines.append("")
    for rid in index:
        entry = profile["roles"].get(rid) or {}
        resolver = entry.get("resolver", {})
        style = (
            resolver.get("style_name")
            or resolver.get("style_id")
            or resolver.get("type")
        )
        lines.append(
            f"- `{rid}`: {style} ({entry.get('status')}) - "
            f"{usage_str(entry.get('usage') or {})}"
        )
    return lines


def palette_roles_md(profile: dict) -> list[str]:
    """The semantic palette-role table: the COLOR TOKENS an author may name."""
    theme = profile.get("theme") or {}
    proles = theme.get("palette_roles") or {}
    if not proles:
        return []
    colors = theme.get("colors") or {}
    lines = ["", "## Brand palette roles", ""]
    lines.append(
        "Semantic color tokens captured from the template. Reference THESE names"
    )
    lines.append(
        "(or a theme slot like `accent1`) as run color tokens; never a raw hex."
    )
    lines.append("")
    for name in sorted(proles):
        val = proles[name] or {}
        slot = val.get("theme") if isinstance(val, dict) else str(val)
        hexv = (colors.get(slot) or {}).get("hex") if slot else None
        hex_str = f" (#{hexv})" if hexv else ""
        lines.append(f"- `{name}` -> `{slot}`{hex_str}")
    return lines


def authoring_hints_md(profile: dict) -> list[str]:
    """How to author content against this profile - role-first, never style-first."""
    lines = ["", "## Authoring hints", ""]
    lines.append(
        "- Pick blocks by MEANING using the role table above; the engine resolves"
    )
    lines.append(
        "  every role to the template's own artifacts. Never name a style, font,"
    )
    lines.append("  or hex anywhere in the input: the profile is the only source.")
    lines.append(
        "- Respect the ordered skeleton (Structure section, when present): cover"
    )
    lines.append(
        "  content first, derived indexes where the template keeps them, then the"
    )
    lines.append("  freeform body in the template's own order.")
    comp = profile.get("comprehension") or {}
    fragments = comp.get("fragments") or []
    if comp.get("status") == "present" and fragments:
        names = ", ".join(
            sorted(str(f.get("name") or f.get("id") or "?") for f in fragments)
        )
        lines.append(
            f"- REUSE the comprehended fragments ({names}) via `component`/`section`"
        )
        lines.append(
            "  blocks with `{{slot}}` values instead of re-deriving their layout."
        )
    else:
        lines.append(
            "- After `comprehend`, reusable fragments (components/sections) may be"
        )
        lines.append(
            "  available in `profile.json` under `comprehension.fragments`: prefer"
        )
        lines.append("  them over re-deriving recurring layouts.")
    return lines
