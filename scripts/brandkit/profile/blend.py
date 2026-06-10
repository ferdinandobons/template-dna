# SPDX-License-Identifier: MIT
"""Multi-template VALUE-FACT blending (REFLECTIONS P3) - the ONE writer of the
``blend`` ledger and ``provenance.blended_shells``.

A single same-format template is the capture-quality ceiling: a sparse primary
trips the ``MIN_RUNS`` / ``MIN_DOMINANCE`` floors (``common/typography.py``) and
drops facts the shell physically carries. :func:`blend` merges the VALUE facts of
a second same-format template's freshly-extracted profile into an existing one,
under five hard rules:

1. SAME-FORMAT ONLY: a kind mismatch is rejected with a clear finding
   (cross-format comparison stays the read-only ``compare-profiles`` verb).
2. VALUE-FACTS ONLY: only captured appearance-axis values (font / size_hp /
   color), the captured ``theme.fonts.body`` facts, the ``hex:RRGGBB`` palette
   inventory, and confidence corroboration may cross. Artifact POINTERS (style
   ids, layout ids, anchors, numbering ids, table style ids) NEVER cross -
   generation still opens the PRIMARY shell and the resolver still
   membership-validates every pointer against it. A fact class that is not
   cleanly value-shaped is NOT blended (it is surfaced in the report's rejected
   classes instead).
3. PRIMARY WINS: a secondary only FILLS facts the primary left unset and
   CORROBORATES agreeing facts (a bounded, deterministic confidence bump). It
   never overrides a primary value (conflicts are reported, kept-primary).
4. PRIMARY-SHELL PROOF: a candidate fill must be PROVABLE on the primary shell -
   it is pre-proved against the SAME membership sets
   ``check_appearance_targets`` re-validates every captured value against
   (fonts allow-set, the shell's own run sizes, the clrScheme palette union the
   run hexes). A secondary may only LOWER the capture floor, never introduce a
   value the primary shell does not carry; an unprovable candidate lands in the
   rejected report (class ``unprovable``). As a belt against mirror drift, the
   transaction re-runs the REAL ``check_appearance_targets`` on the trial
   profile BASELINE-RELATIVE (the overrides block-scoped rationale: a
   pre-existing unrelated red must not block a clean blend, but the blend may
   introduce NOTHING new).
5. FAIL-CLOSED, ALL-OR-NOTHING, IDEMPOTENT: everything validates first, then one
   binary write + one ``profile.json`` write. Any inconsistency rejects with
   findings and leaves ``profile.json`` byte-identical (a reject writes NOTHING,
   unlike the comprehension/overrides rejected blocks). Re-blending the same
   bytes is a structural no-op (sha dedupe), so a repeat blend never moves the
   confidence math; fills across multiple donors are first-writer-wins (the
   precedence rule applied transitively - an earlier blend's fill becomes a
   primary fact), and corroboration is commutative across distinct donors.

The secondary's per-format extraction is REUSED, never reimplemented: the CLI
extracts the donor into a temporary profile (``cwd=`` temp dir, never persisted)
and hands its dict here, so this module is format-neutral (the QA collector
registry is kind-keyed; there is no per-format branch).
"""

from __future__ import annotations

import copy
import json
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional

from brandkit.profile import schema, store

# The closed, geometry-free appearance axes blending may touch (v1). The
# geometry / table / numbering axes stay frozen primary-only: geometry is a mixed
# value + nested dotted-path confidence map (one capture unit a partial import
# would split), and table/numbering carry pointers (style ids / num ids).
BLEND_AXES: tuple[str, ...] = ("font", "size_hp", "color")
# The sibling confidence key each axis value carries (the capture writer's shape).
AXIS_CONFIDENCE_KEY: dict[str, str] = {
    "font": "confidence",
    "size_hp": "size_confidence",
    "color": "color_confidence",
}
# Corroboration bump: a fixed bonus per DISTINCT donor sha, capped, rounded like
# the capture writer (round(x, 3)). The donor's own measured confidence is
# deliberately NOT an input (it would break order-independence under rounding);
# the donor's own dominance floor at its extract is the quality gate.
CORROBORATION_BONUS: float = 0.05
CONFIDENCE_CAP: float = 1.0
# The temp-dir extraction name the CLI uses for the secondary (never persisted).
TEMP_PROFILE_NAME: str = "blend-secondary"

# The closed rejected-class identifiers, in their fixed report order.
_REJECTED_CLASS_ORDER: tuple[str, ...] = (
    "kind",
    "roles",
    "appearance.geometry",
    "appearance.table",
    "appearance.numbering",
    "theme.text.body",
    "theme.palette.slot",
    "unprovable",
)
# The frozen secondary appearance axes surfaced (never crossed) per shared role.
_FROZEN_APPEARANCE_AXES: tuple[str, ...] = ("geometry", "table", "numbering")


@dataclass
class BlendReport:
    """The deterministic outcome ledger of one successful blend computation."""

    noop: bool = False
    filled: list[dict] = field(default_factory=list)
    corroborated: list[dict] = field(default_factory=list)
    conflicts: list[dict] = field(default_factory=list)
    rejected: list[dict] = field(default_factory=list)


class BlendResult:
    """Outcome of a :func:`blend` attempt (mirrors ``overrides.MergeResult``).

    Attributes:
        ok: True iff the blend was written (also True for the sha-dedupe no-op).
        problems: path-prefixed findings; non-empty iff ``ok`` is False.
        report: the :class:`BlendReport` (None on reject).
        profile: the post-blend profile dict as written (the live, unchanged
            primary dict on a no-op; None on reject).
    """

    __slots__ = ("ok", "problems", "report", "profile")

    def __init__(
        self,
        ok: bool,
        problems: list[str],
        report: Optional[BlendReport],
        profile: Optional[dict],
    ):
        self.ok = ok
        self.problems = problems
        self.report = report
        self.profile = profile


def _bump(old: Any) -> float:
    """The bounded, deterministic corroboration formula (applied once per donor)."""
    return round(min(CONFIDENCE_CAP, float(old) + CORROBORATION_BONUS), 3)


def _new_blend_block(live_shell_sha: Optional[str]) -> dict:
    """A fresh ``blend`` ledger block (written only on the first successful blend)."""
    return {
        "schema_version": schema.BLEND_SCHEMA_VERSION,
        # Only ``present`` is ever serialized: a rejected blend writes NOTHING.
        "status": schema.ComprehensionStatus.PRESENT.value,
        "source_shell_sha256": live_shell_sha,
        "ledger": {"filled": {}, "corroborated": {}},
    }


def _collect_primary_facts(primary: store.LoadedProfile):
    """The primary shell's appearance fact sets, via the QA collector registry.

    Late import (the ``overrides.check_membership`` precedent) keeps this module
    free of the qa layer at import time. A collector failure degrades to EMPTY
    fact sets - every candidate fill is then unprovable (fail-closed), mirroring
    ``check_appearance_targets``.
    """
    from brandkit.qa import checks_deterministic as cd

    collector = cd._SHELL_APPEARANCE_COLLECTORS.get(primary.kind)
    if collector is None:
        return cd.ShellAppearanceFacts(
            fonts=set(), sizes=set(), hexes=set(), palette={}
        )
    try:
        return collector(primary.shell_path, primary.profile)
    except Exception:
        return cd.ShellAppearanceFacts(
            fonts=set(), sizes=set(), hexes=set(), palette={}
        )


def _allowed_hexes(facts) -> set[str]:
    """The hex allow-set ``check_appearance_targets`` uses: clrScheme UNION runs."""
    from brandkit.common import color as colorutil

    palette_hexes: set[str] = set()
    for h in facts.palette.values():
        try:
            palette_hexes.add(colorutil.normalize_hex(h))
        except (ValueError, AttributeError):
            continue
    return palette_hexes | set(facts.hexes)


def _provable(axis: str, value: Any, facts, allowed_hexes: set[str]) -> bool:
    """Mirror of the per-axis membership predicate in ``check_appearance_targets``."""
    from brandkit.common import color as colorutil
    from brandkit.qa import checks_deterministic as cd

    if axis == "font":
        latin = (value or {}).get("latin") if isinstance(value, dict) else None
        return isinstance(latin, str) and latin in facts.fonts
    if axis == "size_hp":
        try:
            return int(value) in facts.sizes
        except (TypeError, ValueError):
            return False
    if axis == "color":
        if not isinstance(value, dict):
            return False
        kind = value.get("kind")
        if kind == "hex":
            try:
                return colorutil.normalize_hex(value.get("hex") or "") in allowed_hexes
            except (ValueError, AttributeError):
                return False
        if kind == "theme":
            token = value.get("theme")
            slot = cd._WML_THEME_TO_SLOT.get(token)
            return (slot is not None and slot in facts.palette) or (
                token in colorutil.THEME_SLOTS and token in facts.palette
            )
        return False
    return False


def _role_walk_order(roles: dict) -> list[str]:
    """Primary ``roles._index`` order, then any remainder sorted (deterministic)."""
    index = roles.get("_index")
    ordered = [r for r in index if r in roles] if isinstance(index, list) else []
    remainder = sorted(r for r in roles if r != "_index" and r not in ordered)
    return ordered + remainder


def _appearance_errors(shell_path, profile: dict) -> Counter:
    """The ERROR multiset of the REAL ``check_appearance_targets`` on ``profile``."""
    from brandkit.qa import checks_deterministic as cd

    findings = cd.check_appearance_targets(shell_path, profile)
    return Counter(
        (f.severity, f.message, f.location)
        for f in findings
        if f.severity == schema.Severity.ERROR.value
    )


def blend(
    primary: store.LoadedProfile,
    secondary_profile: dict,
    secondary_bytes: bytes,
    secondary_filename: str,
) -> BlendResult:
    """Merge the secondary profile's VALUE facts into the primary, all-or-nothing.

    Validates everything on a trial deepcopy first, then writes in one
    transaction (binary first, then ``profile.json``: a failed json write leaves
    ``profile.json`` byte-identical and the orphan content-addressed binary inert,
    pruned by the next blend). On ANY problem nothing is written.
    """
    # ----- 1. Guards (each failure rejects with nothing touched) -------------
    p_kind = primary.profile.get("kind")
    s_kind = (secondary_profile or {}).get("kind")
    if p_kind != s_kind:
        return BlendResult(
            False,
            [
                f"blend: kind mismatch: profile is {p_kind!r}, secondary template is "
                f"{s_kind!r}; cross-format comparison is the read-only "
                "compare-profiles verb"
            ],
            None,
            None,
        )
    if primary.shell_drift:
        return BlendResult(
            False,
            [
                "blend: primary shell drifted from recorded provenance; refusing to "
                "blend against an unproven shell"
            ],
            None,
            None,
        )
    if not primary.shell_exists:
        return BlendResult(
            False,
            ["blend: primary shell is missing; refusing to blend without its proof"],
            None,
            None,
        )

    # ----- 2. Dedupe (structural no-op, nothing computed or written) ---------
    sha = store.sha256_bytes(secondary_bytes)
    live_shell_sha = ((primary.profile.get("provenance") or {}).get("shell") or {}).get(
        "sha256"
    )
    recorded_shas = {e.get("sha256") for e in primary.blended_shells}
    if sha == live_shell_sha or sha in recorded_shas:
        return BlendResult(True, [], BlendReport(noop=True), primary.profile)

    # ----- 3+4. Pre-proof facts, then the deterministic plan/apply walk ------
    facts = _collect_primary_facts(primary)
    allowed_hexes = _allowed_hexes(facts)
    trial = copy.deepcopy(primary.profile)
    report = BlendReport()
    ledger_filled: dict[str, str] = {}  # path -> donor sha
    ledger_corroborated: list[str] = []  # paths bumped by THIS donor
    rejected_counts: Counter = Counter()
    unprovable: list[tuple[str, Any]] = []

    def _fill(container: dict, key: str, value: Any, path: str, sibling: Any) -> None:
        container[key] = copy.deepcopy(value)
        if sibling is not None:
            conf_key, conf_val = sibling
            container[conf_key] = conf_val
        report.filled.append({"path": path, "value": value, "source_sha256": sha})
        ledger_filled[path] = sha

    def _corroborate(container: dict, conf_key: str, path: str) -> None:
        # Tolerated defensively: a value with no sibling confidence (cannot happen
        # via the capture writer) is skipped - no bump, no report line.
        old = container.get(conf_key)
        if not isinstance(old, (int, float)):
            return
        new = _bump(old)
        container[conf_key] = new
        report.corroborated.append(
            {"path": path, "before": old, "after": new, "source_sha256": sha}
        )
        ledger_corroborated.append(path)

    # --- roles: shared role ids only, in primary walk order ---
    p_roles = trial.get("roles") or {}
    s_roles = secondary_profile.get("roles") or {}
    shared: list[str] = []
    for rid in _role_walk_order(p_roles):
        if isinstance(p_roles.get(rid), dict) and isinstance(s_roles.get(rid), dict):
            shared.append(rid)
    only_secondary = sorted(
        r
        for r in s_roles
        if r != "_index" and isinstance(s_roles.get(r), dict) and r not in p_roles
    )
    if only_secondary:
        rejected_counts["roles"] = len(only_secondary)

    for rid in shared:
        p_entry = p_roles[rid]
        s_entry = s_roles[rid]
        s_app = s_entry.get("appearance")
        s_app = s_app if isinstance(s_app, dict) else {}
        p_app = p_entry.get("appearance")
        p_app = p_app if isinstance(p_app, dict) else {}
        for axis in BLEND_AXES:
            if axis not in s_app:
                continue
            s_val = s_app[axis]
            conf_key = AXIS_CONFIDENCE_KEY[axis]
            path = f"roles.{rid}.appearance.{axis}"
            # Gap test is STRICTLY "key not in dict": appearance may be {} or
            # missing entirely - both mean every axis is unset.
            if axis not in p_app:
                if _provable(axis, s_val, facts, allowed_hexes):
                    # Materialize the appearance dict ONLY when a fill lands.
                    target = p_entry.get("appearance")
                    if not isinstance(target, dict):
                        target = {}
                        p_entry["appearance"] = target
                    sibling = (conf_key, s_app[conf_key]) if conf_key in s_app else None
                    _fill(target, axis, s_val, path, sibling)
                    p_app = target
                else:
                    unprovable.append((path, s_val))
            elif p_app[axis] == s_val:
                _corroborate(p_app, conf_key, path)
            else:
                report.conflicts.append(
                    {"path": path, "primary": p_app[axis], "secondary": s_val}
                )
        for frozen_axis in _FROZEN_APPEARANCE_AXES:
            if frozen_axis in s_app:
                rejected_counts[f"appearance.{frozen_axis}"] += 1
        # Role-detection confidence: corroborate-only, and ONLY under exact
        # resolver equality (the pointers themselves never cross).
        p_res = p_entry.get("resolver")
        s_res = s_entry.get("resolver")
        if (
            isinstance(p_res, dict)
            and isinstance(s_res, dict)
            and p_res == s_res
            and "confidence" in p_entry
        ):
            _corroborate(p_entry, "confidence", f"roles.{rid}.confidence")

    # --- theme.fonts.body: fill-if-unset / corroborate-if-agreeing ---
    s_theme = secondary_profile.get("theme") or {}
    s_body = (s_theme.get("fonts") or {}).get("body")
    s_body = s_body if isinstance(s_body, dict) else {}
    t_theme = trial.setdefault("theme", {})
    for axis_key, conf_key, provable_axis in (
        ("latin", "confidence", "font"),
        ("size_hp", "size_confidence", "size_hp"),
    ):
        if axis_key not in s_body:
            continue
        p_fonts = t_theme.get("fonts")
        p_fonts = p_fonts if isinstance(p_fonts, dict) else {}
        p_body = p_fonts.get("body")
        p_body = p_body if isinstance(p_body, dict) else {}
        path = f"theme.fonts.body.{axis_key}"
        s_val = s_body[axis_key]
        if axis_key not in p_body:
            wrapped = {"latin": s_val} if axis_key == "latin" else s_val
            if _provable(provable_axis, wrapped, facts, allowed_hexes):
                # Create fonts/body only when a sub-fact actually lands (the
                # capture writer's shape, common/typography.py).
                body = t_theme.setdefault("fonts", {}).setdefault("body", {})
                sibling = (conf_key, s_body[conf_key]) if conf_key in s_body else None
                _fill(body, axis_key, s_val, path, sibling)
            else:
                unprovable.append((path, s_val))
        elif p_body[axis_key] == s_val:
            _corroborate(p_body, conf_key, path)
        else:
            report.conflicts.append(
                {"path": path, "primary": p_body[axis_key], "secondary": s_val}
            )

    # --- theme.text.body: value-class but OUTSIDE the enumerated surface (v1) ---
    s_text_body = (s_theme.get("text") or {}).get("body") or {}
    if isinstance(s_text_body, dict) and "color" in s_text_body:
        rejected_counts["theme.text.body"] += 1

    # --- theme.palette: hex: keys union (fill-only); slot tokens frozen ---
    s_palette = s_theme.get("palette")
    s_palette = s_palette if isinstance(s_palette, dict) else {}
    p_palette = t_theme.get("palette")
    p_palette = p_palette if isinstance(p_palette, dict) else {}
    from brandkit.common import color as colorutil

    for key in sorted(s_palette):
        path = f"theme.palette.{key}"
        if not isinstance(key, str) or not key.startswith("hex:"):
            if key not in p_palette:
                rejected_counts["theme.palette.slot"] += 1
            continue
        if key in p_palette:
            # Agreement on an inventory entry: a report line only (the entry
            # carries no confidence to bump, and the primary entry stays
            # byte-identical). Never persisted to the ledger.
            report.corroborated.append(
                {"path": path, "before": None, "after": None, "source_sha256": sha}
            )
            continue
        try:
            provable = colorutil.normalize_hex(key[len("hex:") :]) in allowed_hexes
        except (ValueError, AttributeError):
            provable = False
        if not provable:
            unprovable.append((path, key[len("hex:") :]))
            continue
        # Verbatim deepcopy: ref byte-copy, provenance facts verbatim (they
        # satisfy the closed PALETTE_WHERE), frequency verbatim, advisory
        # name/purpose/use_when verbatim. Deeper shape problems are caught by
        # the trial-wide schema validation below (fail-closed).
        target_palette = t_theme.setdefault("palette", {})
        target_palette[key] = copy.deepcopy(s_palette[key])
        report.filled.append(
            {"path": path, "value": s_palette[key], "source_sha256": sha}
        )
        ledger_filled[path] = sha

    # --- assemble the closed rejected-class list (fixed order) ---
    for cls in _REJECTED_CLASS_ORDER:
        if cls == "unprovable":
            for path, value in unprovable:
                report.rejected.append(
                    {
                        "class": "unprovable",
                        "detail": f"{path}: "
                        + json.dumps(value, sort_keys=True, ensure_ascii=False),
                    }
                )
        elif rejected_counts.get(cls):
            report.rejected.append(
                {
                    "class": cls,
                    "detail": f"{rejected_counts[cls]} fact(s) kept primary-only",
                }
            )

    # ----- 5. Provenance entry + blend ledger on the trial -------------------
    prov = trial.setdefault("provenance", {})
    entries = prov.get("blended_shells")
    entries = list(entries) if isinstance(entries, list) else []
    entries.append(
        {
            "filename": Path(secondary_filename).name,
            "path": store.blend_shell_relpath(p_kind, sha),
            "sha256": sha,
        }
    )
    prov["blended_shells"] = sorted(entries, key=lambda e: str(e.get("sha256")))

    block = trial.get("blend")
    if not isinstance(block, dict):
        block = _new_blend_block(live_shell_sha)
        trial["blend"] = block
    block["status"] = schema.ComprehensionStatus.PRESENT.value
    # Re-stamped from the LIVE primary sha (which blend never changes).
    block["source_shell_sha256"] = live_shell_sha
    ledger = block.setdefault("ledger", {"filled": {}, "corroborated": {}})
    filled_ledger = ledger.setdefault("filled", {})
    for path, donor in ledger_filled.items():
        # First-writer-wins: an earlier blend's fill is never overwritten.
        filled_ledger.setdefault(path, {"from": donor})
    corroborated_ledger = ledger.setdefault("corroborated", {})
    for path in ledger_corroborated:
        mark = corroborated_ledger.setdefault(path, {"by": []})
        if sha not in mark["by"]:
            mark["by"] = sorted([*mark["by"], sha])

    # ----- 6. Validate-all (baseline-relative; the mirror-drift belt) --------
    new_schema = sorted(
        set(schema.validate(trial)) - set(schema.validate(primary.profile))
    )
    baseline = _appearance_errors(primary.shell_path, primary.profile)
    after = _appearance_errors(primary.shell_path, trial)
    new_qa = after - baseline
    problems = list(new_schema)
    for severity, message, location in sorted(new_qa, key=str):
        where = f" [{location}]" if location else ""
        problems.append(f"appearance_targets_exist: {severity} {message}{where}")
    if problems:
        return BlendResult(False, problems, None, None)

    # ----- 7. Write-once (binary first, then json) ----------------------------
    store.save_blend_shell(primary.directory, p_kind, secondary_bytes)
    store.prune_blend_shells(
        primary.directory,
        p_kind,
        {
            e.get("sha256")
            for e in prov["blended_shells"]
            if isinstance(e.get("sha256"), str)
        },
    )
    store.write_profile_json(primary.directory, trial)
    return BlendResult(True, [], report, trial)


def _render_value(value: Any) -> str:
    return json.dumps(value, sort_keys=True, ensure_ascii=False)


def render_report(name: str, filename: str, sha256: str, report: BlendReport) -> str:
    """The deterministic human report for one blend (plan order, empty sections
    omitted)."""
    sha12 = sha256[:12]
    lines = [f"blended {name} <- {filename} (sha256 {sha12})"]
    if report.filled:
        lines.append(f"  filled ({len(report.filled)}):")
        for entry in report.filled:
            lines.append(
                f"    {entry['path']} = {_render_value(entry['value'])} "
                f"[from {entry['source_sha256'][:12]}]"
            )
    if report.corroborated:
        lines.append(f"  corroborated ({len(report.corroborated)}):")
        for entry in report.corroborated:
            if entry["before"] is None:
                lines.append(f"    {entry['path']} (agreement)")
            else:
                lines.append(
                    f"    {entry['path']} (confidence {entry['before']} -> "
                    f"{entry['after']})"
                )
    if report.conflicts:
        lines.append(f"  conflicts kept-primary ({len(report.conflicts)}):")
        for entry in report.conflicts:
            lines.append(
                f"    {entry['path']} (primary {_render_value(entry['primary'])} | "
                f"secondary {_render_value(entry['secondary'])})"
            )
    if report.rejected:
        lines.append(f"  rejected ({len(report.rejected)}):")
        for entry in report.rejected:
            lines.append(f"    {entry['class']}: {entry['detail']}")
    lines.append(
        f"blend {name}: {len(report.filled)} filled, "
        f"{len(report.corroborated)} corroborated, "
        f"{len(report.conflicts)} conflicts kept-primary, "
        f"{len(report.rejected)} rejected"
    )
    return "\n".join(lines)
