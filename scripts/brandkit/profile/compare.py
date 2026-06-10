# SPDX-License-Identifier: MIT
"""Read-only comparison of two Brand Profiles - the cross-template drift report.

The same business brand extracted from two templates (the Word report and the
PowerPoint deck, or two revisions of the same template) can yield profiles
that silently disagree: a theme slot resolves to different hexes, the corporate
blue lives as a theme slot in one file and as a raw off-theme hex in the other,
the captured body font differs. Nothing in a single profile can see this;
:func:`compare_profiles` can.

Strictly diagnostic and read-only: nothing here writes a profile, and nothing
here feeds generation. Structural facts (named styles, layouts, anchors,
geometry) are deliberately NOT compared as drift - they are per-shell by design
and legitimately differ across formats; only the role-id coverage is reported,
as information. Brand-level facts (theme colors, theme/captured fonts, semantic
palette roles, off-theme hex usage) are the drift surface: those are the facts
a company defines once for the whole brand.
"""

from __future__ import annotations

from typing import Any

# A palette key like ``hex:C24D2B`` records a raw color seen OUTSIDE the theme
# slots. This prefix is part of the frozen palette vocabulary (CONVENTIONS §13).
_RAW_HEX_PREFIX = "hex:"

VERDICT_ALIGNED = "aligned"
VERDICT_DRIFT = "drift_detected"


def _theme_colors(profile: dict) -> dict[str, str]:
    colors = (profile.get("theme") or {}).get("colors") or {}
    out: dict[str, str] = {}
    for slot, val in colors.items():
        hexv = (val or {}).get("hex") if isinstance(val, dict) else None
        if isinstance(hexv, str) and hexv:
            out[str(slot)] = hexv.upper()
    return out


def _fonts(profile: dict) -> dict[str, Any]:
    fonts = (profile.get("theme") or {}).get("fonts") or {}
    out: dict[str, Any] = {}
    for slot in ("major", "minor"):
        latin = (fonts.get(slot) or {}).get("latin")
        if latin:
            out[slot] = str(latin)
    body = fonts.get("body") or {}
    if body.get("latin"):
        out["body"] = str(body["latin"])
    if body.get("size_hp") is not None:
        out["body_size_hp"] = body["size_hp"]
    return out


def _palette_roles(profile: dict) -> dict[str, str]:
    roles = (profile.get("theme") or {}).get("palette_roles") or {}
    out: dict[str, str] = {}
    for name, val in roles.items():
        ref = None
        if isinstance(val, dict):
            ref = val.get("theme") or val.get("hex")
        elif isinstance(val, str):
            ref = val
        if ref:
            out[str(name)] = str(ref)
    return out


def _raw_hexes(profile: dict) -> set[str]:
    palette = (profile.get("theme") or {}).get("palette") or {}
    return {
        key[len(_RAW_HEX_PREFIX) :].upper()
        for key in palette
        if isinstance(key, str) and key.startswith(_RAW_HEX_PREFIX)
    }


def _role_ids(profile: dict) -> list[str]:
    index = (profile.get("roles") or {}).get("_index") or []
    return [str(r) for r in index]


def _identity(profile: dict) -> dict[str, Any]:
    ident = profile.get("identity") or {}
    surface = profile.get("surface") or {}
    kind = profile.get("kind") or next(iter(surface), None)
    return {
        "name": ident.get("name"),
        "kind": kind,
        "locale": ident.get("locale"),
        "schema": profile.get("$schema"),
    }


def _diff_maps(a: dict, b: dict) -> dict[str, Any]:
    """Split two flat maps into agree / differ / only_a / only_b."""
    agree: dict[str, Any] = {}
    differ: dict[str, dict[str, Any]] = {}
    for key in sorted(set(a) & set(b)):
        if a[key] == b[key]:
            agree[key] = a[key]
        else:
            differ[key] = {"a": a[key], "b": b[key]}
    only_a = {k: a[k] for k in sorted(set(a) - set(b))}
    only_b = {k: b[k] for k in sorted(set(b) - set(a))}
    return {"agree": agree, "differ": differ, "only_a": only_a, "only_b": only_b}


def _off_theme_findings(
    raw_hexes: set[str], other_colors: dict[str, str]
) -> list[dict[str, str]]:
    """Raw off-theme hexes in one profile that ARE a theme slot in the other.

    The sharpest drift signal: the same physical color is wired through the
    theme in one template and hardcoded in the other, so theme-level restyling
    diverges between the two documents.
    """
    by_hex: dict[str, list[str]] = {}
    for slot, hexv in other_colors.items():
        by_hex.setdefault(hexv, []).append(slot)
    return [
        {"hex": hexv, "other_slots": ", ".join(sorted(by_hex[hexv]))}
        for hexv in sorted(raw_hexes)
        if hexv in by_hex
    ]


def compare_profiles(profile_a: dict, profile_b: dict) -> dict[str, Any]:
    """Pure, read-only comparison; returns a structured drift report.

    Drift (the verdict flips to :data:`VERDICT_DRIFT`) is ONLY brand-level
    disagreement: a shared theme slot with different hexes, shared font slots
    with different families or captured sizes, shared semantic palette roles
    bound differently, or an off-theme raw hex in one profile that is a theme
    slot in the other. Role coverage and slots present in only one profile are
    reported as information, never as drift - per-format surfaces legitimately
    capture different subsets.
    """
    colors_a, colors_b = _theme_colors(profile_a), _theme_colors(profile_b)
    fonts = _diff_maps(_fonts(profile_a), _fonts(profile_b))
    colors = _diff_maps(colors_a, colors_b)
    proles = _diff_maps(_palette_roles(profile_a), _palette_roles(profile_b))
    roles_a, roles_b = set(_role_ids(profile_a)), set(_role_ids(profile_b))
    off_theme = {
        "a_raw_is_b_slot": _off_theme_findings(_raw_hexes(profile_a), colors_b),
        "b_raw_is_a_slot": _off_theme_findings(_raw_hexes(profile_b), colors_a),
    }
    drift = bool(
        colors["differ"]
        or fonts["differ"]
        or proles["differ"]
        or off_theme["a_raw_is_b_slot"]
        or off_theme["b_raw_is_a_slot"]
    )
    return {
        "a": _identity(profile_a),
        "b": _identity(profile_b),
        "theme_colors": colors,
        "fonts": fonts,
        "palette_roles": proles,
        "off_theme": off_theme,
        "roles": {
            "common": sorted(roles_a & roles_b),
            "only_a": sorted(roles_a - roles_b),
            "only_b": sorted(roles_b - roles_a),
        },
        "verdict": VERDICT_DRIFT if drift else VERDICT_ALIGNED,
    }


def render_report(result: dict[str, Any]) -> str:
    """Human-readable rendering of a :func:`compare_profiles` result."""
    a, b = result["a"], result["b"]
    lines = [
        "compare-profiles: %s (%s)  vs  %s (%s)"
        % (a.get("name"), a.get("kind"), b.get("name"), b.get("kind")),
    ]

    def _section(title: str, diff: dict[str, Any]) -> None:
        lines.append("")
        lines.append(
            "%s: %d agree, %d differ, %d only-A, %d only-B"
            % (
                title,
                len(diff["agree"]),
                len(diff["differ"]),
                len(diff["only_a"]),
                len(diff["only_b"]),
            )
        )
        for key, pair in diff["differ"].items():
            lines.append("  DRIFT %s: A=%s  B=%s" % (key, pair["a"], pair["b"]))
        for key, val in diff["only_a"].items():
            lines.append("  info  %s: only in A (%s)" % (key, val))
        for key, val in diff["only_b"].items():
            lines.append("  info  %s: only in B (%s)" % (key, val))

    _section("theme colors", result["theme_colors"])
    _section("fonts", result["fonts"])
    _section("palette roles", result["palette_roles"])

    off = result["off_theme"]
    if off["a_raw_is_b_slot"] or off["b_raw_is_a_slot"]:
        lines.append("")
        lines.append("off-theme usage:")
        for f in off["a_raw_is_b_slot"]:
            lines.append(
                "  DRIFT raw #%s in A is theme slot %s in B"
                % (f["hex"], f["other_slots"])
            )
        for f in off["b_raw_is_a_slot"]:
            lines.append(
                "  DRIFT raw #%s in B is theme slot %s in A"
                % (f["hex"], f["other_slots"])
            )

    roles = result["roles"]
    lines.append("")
    lines.append(
        "roles: %d common, %d only-A, %d only-B (informational; per-format "
        "surfaces differ by design)"
        % (len(roles["common"]), len(roles["only_a"]), len(roles["only_b"]))
    )
    if roles["only_a"]:
        lines.append("  only-A: " + ", ".join(roles["only_a"]))
    if roles["only_b"]:
        lines.append("  only-B: " + ", ".join(roles["only_b"]))

    lines.append("")
    lines.append("verdict: " + result["verdict"])
    return "\n".join(lines)
