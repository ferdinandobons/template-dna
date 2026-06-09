# SPDX-License-Identifier: MIT
"""The ONE writer of the learned-overrides block (Cluster B, ``rules.overrides``).

Two producers feed a SINGLE canonical sink:

1. :func:`learn` (B3, deterministic, model-free) distills a brand-safe lesson from
   the cross-run ``generation_report.json`` history (B1/B2): it keeps the findings
   whose ``check`` is in :data:`schema.LEARNABLE_CHECKS` that recurred across
   ``>= min_runs`` SAME-shell runs, maps each to exactly ONE closed-kind override
   entry whose target it re-validates against the profile's surfaced inventories,
   drops any entry that fails membership, proves the reroute graph is ACYCLIC, and
   routes the survivors through :func:`merge_overrides`.

2. ``merge_overrides`` (also reused by the B4 model proposer, a later increment) is
   the canonical writer: shape-validate -> fail-closed :func:`check_membership`
   binding -> ALL-OR-NOTHING. A clean proposal becomes a canonical, sorted
   ``present`` block with ``source_shell_sha256`` stamped from the live
   ``provenance.shell.sha256`` and a flat ``provenance`` record of which
   ``(check, location, recurred_runs)`` produced each entry; any unbound pointer
   rejects the WHOLE proposal (``status='rejected'``, nothing load-bearing written).

Because both producers write through ``merge_overrides`` there is exactly one sink,
one sort order, and one ``source_shell_sha256`` stamp per write, so a deterministic
lesson and a model proposal can never clobber each other or disagree on shape. This
module is structurally incapable of authoring a style/font/hex: a reroute reuses an
EXISTING role's resolver verbatim (still through the resolver's legal-type gate), a
number_format swap only names a mask the shell already uses, and a demo-clear only
names a captured demo string - the resolver stays the single brand chokepoint.
"""

from __future__ import annotations

from typing import Any, Optional

from brandkit.profile import comprehension as comprehension_mod
from brandkit.profile import schema

# The recurrence threshold the deterministic ``learn`` step gates on: a
# ``(check, location)`` pair seen in a SINGLE run is never distilled into a lesson (a
# one-off is noise); it must recur across at least this many same-shell runs (the
# ``recurred_runs`` count B2 records, i.e. the number of persisted reports that carry
# the pair). Advisory on top of this: the CLI keeps the written block out of the LIVE
# resolver until an explicit ``--accept`` (see ``cli.py``), so even a recurrence
# cannot mint a permanent live lesson on its own.
LEARN_MIN_RUNS: int = 2


# ---------------------------------------------------------------------------
# Outcome objects (mirror comprehension.MergeResult)
# ---------------------------------------------------------------------------
class MergeResult:
    """Outcome of a :func:`merge_overrides` attempt.

    Attributes:
        ok: True iff the overrides were written ``present``.
        status: the resulting ``rules.overrides.status`` (``present`` | ``rejected``).
        problems: the validation/membership findings (empty iff ``ok``).
    """

    __slots__ = ("ok", "status", "problems")

    def __init__(self, ok: bool, status: str, problems: list[str]):
        self.ok = ok
        self.status = status
        self.problems = problems


class LearnResult:
    """Outcome of a :func:`learn` attempt.

    Attributes:
        ok: True iff a lesson was written ``present`` (False when nothing crossed the
            threshold/bound, OR the proposal was rejected by the sink).
        status: the resulting ``rules.overrides.status`` (``present`` | ``rejected``
            | ``absent`` when there was nothing to distil).
        problems: the sink's findings (empty unless the sink rejected).
        distilled: the count of closed-kind entries the threshold + membership +
            acyclicity gates let through to the sink.
    """

    __slots__ = ("ok", "status", "problems", "distilled")

    def __init__(self, ok: bool, status: str, problems: list[str], distilled: int):
        self.ok = ok
        self.status = status
        self.problems = problems
        self.distilled = distilled


# ---------------------------------------------------------------------------
# Membership (fail-closed, reject-never-skip) -- the binding gate
# ---------------------------------------------------------------------------
def _surface_number_format_masks(profile: dict) -> set[str]:
    """The masks the shell uses, read from ``surface.<kind>.number_formats``.

    Profile-internal peer of ``checks_deterministic._xlsx_number_format_masks`` (the
    same set the schema's intra-profile number_format check binds to, schema.py:716),
    so ``merge_overrides`` can bind a swap mask without opening the package. The
    fail-closed ``check_override_targets`` QA check re-proves the SAME mask against
    the live shell bytes at verify time, so an empty surface inventory here still
    can't smuggle a fabricated mask past the gate.
    """
    kind = profile.get("kind")
    sub = ((profile.get("surface") or {}).get(kind) or {}) if kind else {}
    if not isinstance(sub, dict):
        return set()
    return {
        f.get("format")
        for f in (sub.get("number_formats") or [])
        if isinstance(f, dict) and f.get("format")
    }


def check_membership(profile: dict, overrides: dict) -> list[str]:
    """Return fail-closed membership problems for an overrides block.

    Every re-point pointer must be a verbatim id from the profile's OWN surfaced
    inventories, mirroring :func:`comprehension.check_membership` (reject-never-skip,
    NOT the namespace-guarded skip-on-empty of ``_validate_resolver_consistency``):
    a pointer into an EMPTY or ABSENT inventory is itself a problem, so an override
    can never bind into nothing.

      - every ``reroute_roles`` TARGET is a declared role id
        (:func:`schema.list_role_ids`); the requested KEY need not be a declared role
        (rerouting an ABSENT/stub role is the whole point), only the target must be
        real and shell-backed (the resolver re-proof + ``check_override_targets``);
      - every ``number_format_swaps`` MASK is in ``surface.<kind>.number_formats``;
      - every ``demo_clears`` VALUE is in the captured demo set
        (``checks_deterministic.captured_template_texts``).

    Returns ``[]`` when ``overrides`` is absent / not ``present`` (nothing to bind),
    matching the comprehension contract.
    """
    if not isinstance(overrides, dict):
        return []
    status = overrides.get("status")
    # Only a PRESENT (or status-less, mid-merge trial) block carries load-bearing
    # pointers to enforce. ABSENT/REJECTED carry nothing enforceable.
    if status not in (None, schema.ComprehensionStatus.PRESENT.value):
        return []

    # Late import keeps this module free of the qa layer at import time; the captured
    # demo set is the same one ``check_override_targets`` rejects against.
    from brandkit.qa import checks_deterministic as cd

    role_ids = set(schema.list_role_ids(profile))
    masks = _surface_number_format_masks(profile)
    captured = set(cd.captured_template_texts(profile))

    problems: list[str] = []

    # (a) reroute targets ∈ declared roles (FAIL-CLOSED on empty).
    for requested, target in (overrides.get("reroute_roles") or {}).items():
        if target not in role_ids:
            problems.append(
                f"rules.overrides.reroute_roles.{requested}: target {target!r} not in "
                f"declared roles {sorted(role_ids)}"
            )

    # (b) swap masks ∈ surfaced number formats (FAIL-CLOSED on empty).
    for rid, mask in (overrides.get("number_format_swaps") or {}).items():
        if mask not in masks:
            problems.append(
                f"rules.overrides.number_format_swaps.{rid}: mask {mask!r} not in "
                f"surface.{profile.get('kind')}.number_formats {sorted(masks)}"
            )

    # (c) demo-clear values ∈ captured demo set (FAIL-CLOSED on empty).
    for i, value in enumerate(overrides.get("demo_clears") or []):
        if value not in captured:
            problems.append(
                f"rules.overrides.demo_clears[{i}]: value {value!r} was not captured "
                "for this template"
            )

    return problems


# ---------------------------------------------------------------------------
# Acyclicity (reuse the comprehension colored-DFS)
# ---------------------------------------------------------------------------
def _reroute_cycle(reroutes: dict) -> set:
    """Return the role ids on a cycle in the reroute graph (or empty if acyclic).

    Reuses :func:`comprehension._detect_fragment_cycle` (the iterative colored DFS
    that can never overflow the recursion limit on an adversarial cycle). Nodes are
    role ids; an edge ``A -> B`` exists for each ``reroute_roles[A] == B``. Only
    edges whose target is itself a reroute KEY can extend a path, so the resolver's
    SINGLE-HOP runtime guard and this build-time acyclicity proof agree: a
    ``A -> B -> A`` proposal is rejected here before it can ever be written.
    """
    graph: dict[str, list[str]] = {}
    for requested, target in reroutes.items():
        # Record the edge only when the target is itself a reroute source, so the
        # graph keys/edges are exactly the nodes that can form a cycle.
        graph.setdefault(requested, [])
        if target in reroutes:
            graph[requested].append(target)
        graph.setdefault(target, graph.get(target, []))
    return comprehension_mod._detect_fragment_cycle(graph)


# ---------------------------------------------------------------------------
# merge_overrides -- the single canonical writer (shared by learn + B4 model)
# ---------------------------------------------------------------------------
def _canonicalize_overrides(
    proposal: dict,
    shell_sha: Optional[str],
    generated_by: Optional[dict],
) -> dict:
    """Return the overrides block with stable order and stamped provenance.

    Mirrors :func:`comprehension._canonicalize`: starts from the shaped
    :func:`schema.empty_overrides` default, forces ``status='present'``, stamps the
    live shell sha + ``generated_by`` + ``confidence``, and rewrites the three
    closed-kind containers with sorted keys / sorted list order so a re-write of the
    same proposal serializes byte-identically (idempotency).
    """
    out = schema.empty_overrides()
    out["status"] = schema.ComprehensionStatus.PRESENT.value
    out["source_shell_sha256"] = shell_sha
    if generated_by is not None:
        out["generated_by"] = dict(generated_by)
    elif proposal.get("generated_by") is not None:
        out["generated_by"] = dict(proposal["generated_by"])
    conf = proposal.get("confidence")
    out["confidence"] = float(conf) if isinstance(conf, (int, float)) else 0.0

    reroutes = proposal.get("reroute_roles") or {}
    out["reroute_roles"] = {
        k: reroutes[k]
        for k in sorted(reroutes)
        if isinstance(k, str) and isinstance(reroutes.get(k), str)
    }
    swaps = proposal.get("number_format_swaps") or {}
    out["number_format_swaps"] = {
        k: swaps[k]
        for k in sorted(swaps)
        if isinstance(k, str) and isinstance(swaps.get(k), str)
    }
    clears = proposal.get("demo_clears") or []
    out["demo_clears"] = sorted({c for c in clears if isinstance(c, str) and c})

    prov = proposal.get("provenance") or {}
    out["provenance"] = {
        k: dict(prov[k]) for k in sorted(prov) if isinstance(prov.get(k), dict)
    }
    return out


def merge_overrides(
    profile: dict,
    proposal: dict,
    *,
    generated_by: Optional[dict] = None,
) -> MergeResult:
    """Validate ``proposal`` against ``profile`` and, if clean, write it in.

    The SINGLE canonical writer of ``rules.overrides`` (Cluster B). Mutates
    ``profile`` in place: on success ``profile['rules']['overrides']`` is the
    canonical ``present`` block (sorted keys / sorted lists, ``source_shell_sha256``
    = live ``provenance.shell.sha256``); on ANY problem NOTHING load-bearing is
    written - the block becomes ``status='rejected'`` carrying the findings
    (all-or-nothing, the comprehension transaction shape).

    Order of gates (each fail-closed):
      1. shape (:func:`schema._validate_overrides`, the SAME shape-only validator the
         load-time ``schema.validate`` runs on this block - it binds only the
         overrides block, never the surrounding profile, so a pre-existing unrelated
         profile shape issue cannot reject a clean lesson);
      2. fail-closed :func:`check_membership` binding of every pointer;
      3. reroute acyclicity (a cycle is rejected before it can be written).

    Args:
        profile: the loaded profile dict (mutated in place).
        proposal: the override proposal (``reroute_roles`` / ``number_format_swaps``
            / ``demo_clears`` / ``confidence`` / ``provenance``); any supplied
            ``status`` is DISPOSED (never trusted, forced to ``present`` for the
            trial so every binding gate always runs).
        generated_by: optional ``{"model","prompt_version","generated_at"}``
            provenance to stamp.

    Returns:
        A :class:`MergeResult`.
    """
    if not isinstance(proposal, dict):
        return _reject_overrides(
            profile, ["rules.overrides: proposal must be an object"], generated_by
        )

    trial = dict(proposal)
    # Dispose any supplied status: a model/learn-supplied status is never trusted (it
    # would otherwise let status='absent'/'rejected' short-circuit the binding gates).
    trial["status"] = schema.ComprehensionStatus.PRESENT.value

    # 1) Shape: the SAME shape-only validator load-time ``schema.validate`` runs on
    # the block. It binds only the overrides block (never the surrounding profile),
    # so a clean lesson is never rejected by an unrelated pre-existing profile shape
    # issue - and a model/learn proposal that is itself malformed is still rejected.
    problems = list(schema._validate_overrides(trial))

    # 2) Fail-closed membership of every pointer.
    problems.extend(check_membership(profile, trial))

    # 3) Reroute acyclicity (reuse the comprehension colored-DFS).
    reroutes = trial.get("reroute_roles") or {}
    if isinstance(reroutes, dict) and reroutes:
        cycle = _reroute_cycle(reroutes)
        if cycle:
            problems.append(
                "rules.overrides.reroute_roles: reroute graph has a cycle through "
                f"{sorted(cycle)}"
            )

    if problems:
        return _reject_overrides(profile, sorted(problems), generated_by)

    # Clean: write the canonical block with stable, sorted serialization.
    shell_sha = ((profile.get("provenance") or {}).get("shell") or {}).get("sha256")
    canonical = _canonicalize_overrides(trial, shell_sha, generated_by)
    rules = profile.setdefault("rules", {})
    rules["overrides"] = canonical
    return MergeResult(True, schema.ComprehensionStatus.PRESENT.value, [])


def _reject_overrides(
    profile: dict, problems: list[str], generated_by: Optional[dict]
) -> MergeResult:
    """Write a ``rejected`` overrides block carrying the findings; nothing else."""
    rejected = schema.empty_overrides()
    rejected["status"] = schema.ComprehensionStatus.REJECTED.value
    rejected["findings"] = sorted(problems)
    if generated_by is not None:
        rejected["generated_by"] = dict(generated_by)
    rules = profile.setdefault("rules", {})
    rules["overrides"] = rejected
    return MergeResult(False, schema.ComprehensionStatus.REJECTED.value, problems)


# ---------------------------------------------------------------------------
# learn -- deterministic distillation of recurring findings into a proposal
# ---------------------------------------------------------------------------
def _recurrence_counts(
    reports: list[dict], shell_sha: Optional[str]
) -> dict[tuple[str, Any], int]:
    """Count how many SAME-shell reports carry each ``(check, location)`` pair.

    The recurrence key is STRICTLY ``(check, location)`` - never the ``message``
    body (which can carry brand/template text) - matching the B2 contract
    (``report._location_key``). Only reports whose ``shell_sha256`` matches the live
    shell contribute, so a re-extract (new sha) starts a fresh history
    (SHELL-FROZEN). The count is the number of distinct RUNS that carried the pair (a
    finding seen twice IN one report is one run); each persisted report is one run, so
    a pair seen in N reports has run across N runs - the ``recurred_runs`` count
    ``learn`` thresholds on (a pair in a SINGLE report is "seen once").
    """
    counts: dict[tuple[str, Any], int] = {}
    for report in reports:
        if not isinstance(report, dict):
            continue
        if shell_sha is not None and report.get("shell_sha256") != shell_sha:
            continue
        seen: set[tuple[str, Any]] = set()
        for finding in report.get("findings") or []:
            if not isinstance(finding, dict):
                continue
            key = (finding.get("check"), finding.get("location"))
            if key in seen:
                continue
            seen.add(key)
            counts[key] = counts.get(key, 0) + 1
    return counts


def learn(
    profile: dict,
    reports: list[dict],
    *,
    min_runs: int = LEARN_MIN_RUNS,
    generated_by: Optional[dict] = None,
) -> LearnResult:
    """Distil recurring QA findings into a brand-safe overrides lesson (B3).

    The deterministic, model-free counterpart of :func:`comprehension.merge`. Reads
    the SAME-shell ``generation_report.json`` digests (B1/B2 history), keeps only the
    findings whose ``check`` is in :data:`schema.LEARNABLE_CHECKS` that recurred
    across ``>= min_runs`` same-shell runs, maps each surviving ``(check, location)``
    to exactly ONE closed-kind override entry, then routes the proposal through
    :func:`merge_overrides` - the SAME sink the B4 model proposer uses - so there is
    one writer, one sort order, one membership/acyclicity gate.

    Mapping (``location`` is the binding pointer the report carries, never the
    brand-bearing ``message``):
      - ``resolver_targets_exist`` / ``style_fallback`` whose ``location`` is a
        declared role id whose own resolver is a STUB ({} or no shell-backed target)
        AND a sibling role of the SAME family resolves healthily -> a
        ``reroute_role`` to that sibling. (A finding whose ``location`` cannot be
        bound to a stub-role-with-a-healthy-sibling is DROPPED - never guessed.)
      - ``no_residual_template_text`` whose ``location`` is a captured demo string ->
        a ``register_demo_clear`` of that string.

    Every mapped entry is independently re-validated by ``merge_overrides``'
    fail-closed ``check_membership`` (and, for reroutes, the acyclicity proof), so an
    entry that does not bind is dropped by the sink rather than written. The reroute
    here is single-source -> single-target, so it is structurally acyclic; the
    acyclicity gate is the backstop for a future multi-entry proposal.

    Args:
        profile: the loaded profile (mutated in place by ``merge_overrides``).
        reports: prior ``generation_report.json`` digests (any shell; same-shell is
            enforced internally against the live ``provenance.shell.sha256``).
        min_runs: the recurrence threshold (default :data:`LEARN_MIN_RUNS`).
        generated_by: optional provenance stamp passed through to the sink.

    Returns:
        A :class:`LearnResult`. When nothing crosses the threshold/binds, the block
        is left/written ``absent`` and ``ok`` is False with ``distilled == 0``.
    """
    shell_sha = ((profile.get("provenance") or {}).get("shell") or {}).get("sha256")
    counts = _recurrence_counts(reports, shell_sha)

    role_ids = set(schema.list_role_ids(profile))
    roles = profile.get("roles") or {}

    from brandkit.qa import checks_deterministic as cd

    captured = set(cd.captured_template_texts(profile))

    reroute_roles: dict[str, str] = {}
    demo_clears: set[str] = set()
    provenance: dict[str, dict] = {}

    # Iterate the recurring keys in a deterministic order so the proposal (and so the
    # written block) is stable run to run.
    for (check, location), recurred_runs in sorted(
        counts.items(), key=lambda kv: (str(kv[0][0]), str(kv[0][1]))
    ):
        if check not in schema.LEARNABLE_CHECKS:
            continue
        # ``recurred_runs`` is the number of distinct same-shell runs that carried the
        # pair (one per report). A pair seen in a SINGLE run never crosses the default
        # ``min_runs`` of 2: a one-off is noise, not a lesson.
        if recurred_runs < min_runs:
            continue
        if not isinstance(location, str) or not location:
            # Unbindable (no structured pointer) -> drop, never guess a target.
            continue

        if check in ("resolver_targets_exist", "style_fallback"):
            target = _healthy_sibling(location, role_ids, roles)
            if target is None:
                continue  # no brand-safe re-point -> drop
            reroute_roles[location] = target
            provenance[f"reroute_roles.{location}"] = {
                "check": check,
                "location": location,
                "recurred_runs": recurred_runs,
            }
        elif check == "no_residual_template_text":
            if location not in captured:
                continue  # not a captured demo string -> drop
            demo_clears.add(location)
            provenance[f"demo_clears.{location}"] = {
                "check": check,
                "location": location,
                "recurred_runs": recurred_runs,
            }

    if not reroute_roles and not demo_clears:
        # Nothing crossed the threshold and bound: leave the block ``absent`` (the
        # byte-identical default), report nothing distilled.
        return LearnResult(False, schema.ComprehensionStatus.ABSENT.value, [], 0)

    proposal = {
        "reroute_roles": reroute_roles,
        "demo_clears": sorted(demo_clears),
        "confidence": _confidence(provenance, min_runs),
        "provenance": provenance,
    }
    result = merge_overrides(profile, proposal, generated_by=generated_by)
    distilled = len(reroute_roles) + len(demo_clears) if result.ok else 0
    return LearnResult(result.ok, result.status, result.problems, distilled)


def _healthy_sibling(requested: str, role_ids: set[str], roles: dict) -> Optional[str]:
    """The declared SAME-family role to reroute a stub ``requested`` role at.

    A brand-safe re-point only exists when ``requested`` is genuinely a stub (it has
    no shell-backed resolver) and a SIBLING role in the same dotted family (e.g.
    ``heading.3`` for ``heading.9``, ``paragraph`` for ``paragraph.lead``) DOES carry
    a concrete resolver. Returns that sibling, or ``None`` when no brand-safe target
    exists (so the finding is dropped, never guessed). ``merge_overrides`` re-proves
    the chosen target against the surfaced inventory; this only narrows the search to
    a same-family, concretely-resolved role.
    """
    # A requested role that already resolves healthily is NOT a stub -> nothing to
    # learn (rerouting it would silently change branded output).
    req_resolver = (roles.get(requested) or {}).get("resolver") or {}
    if req_resolver:
        return None

    family = requested.split(".", 1)[0]
    candidates = []
    for rid in role_ids:
        if rid == requested:
            continue
        resolver = (roles.get(rid) or {}).get("resolver") or {}
        if not resolver:
            continue  # the sibling must itself be concretely resolved
        rid_family = rid.split(".", 1)[0]
        if rid_family != family:
            continue
        candidates.append(rid)
    if not candidates:
        return None
    # Deterministic pick: the same-family role with the shortest id, ties broken
    # lexicographically (prefers the family root, e.g. ``paragraph`` over
    # ``paragraph.lead``, ``heading.1`` over ``heading.10``).
    return sorted(candidates, key=lambda r: (len(r), r))[0]


def _confidence(provenance: dict, min_runs: int) -> float:
    """A bounded [0,1] confidence from how strongly the entries recurred.

    Higher recurrence -> higher confidence, capped at 1.0. Purely advisory (the
    schema only requires [0,1]); the CLI ``--accept`` gate, not this number, decides
    whether the lesson goes live, so it never on its own entrenches a fix.
    """
    if not provenance:
        return 0.0
    runs = [
        e.get("recurred_runs", min_runs)
        for e in provenance.values()
        if isinstance(e, dict)
    ]
    if not runs:
        return 0.0
    # Map the minimum recurrence across entries onto [0,1]: exactly min_runs -> 0.5,
    # growing toward 1.0 as recurrence climbs.
    weakest = min(runs)
    return min(1.0, 0.5 + 0.1 * max(0, weakest - min_runs))
