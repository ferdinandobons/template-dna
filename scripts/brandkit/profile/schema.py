# SPDX-License-Identifier: MIT
"""The FROZEN Brand-Profile vocabulary - the single source of truth.

Every other module imports its enum values, key names, and the ``profile.json``
envelope shape from here. The plan (§3) froze the vocabulary to resolve the
cohesion conflicts between dimensions; this module *is* that freeze:

- The discriminator key is ``kind`` (``docx`` | ``pptx`` | ``xlsx``), never
  ``doc_type``. See :class:`Kind`.
- The role registry key is ``roles``, never ``bindings``.
- Resolver types are ``named_style`` | ``placeholder`` | ``cell_style`` |
  ``named_range`` | ``number_format`` (all first-class), plus the still
  maturity-staged ``chart_template``. The redundant ``layout_placeholder`` is
  dropped. See :class:`ResolverType`.
- ``schema_version`` is semver; this module pins :data:`SCHEMA_VERSION`.
- The shell always lives at ``template/shell.<ext>``.
- Role status is ``robust`` | ``best_effort`` | ``stub`` (:class:`Status`);
  finding severity is ``INFO`` | ``WARNING`` | ``ERROR`` (:class:`Severity`).

Two helpers complete the contract:
  - :func:`build_envelope` - construct a minimal, valid envelope skeleton.
  - :func:`validate` - return a list of human-readable problems (``[]`` == ok).

Plus role-id helpers (:func:`role_id`, :func:`parse_role_id`) so callers never
hand-concatenate role ids like ``"heading.1"``.
"""

from __future__ import annotations

import re
from enum import Enum
from typing import Any, Optional

# ---------------------------------------------------------------------------
# Pinned schema version (semver). Bump MINOR for additive, MAJOR for breaking.
# ---------------------------------------------------------------------------
SCHEMA_VERSION: str = "1.2.0"
SCHEMA_ID: str = "https://brand-docs/schema/profile-1.json"

# The newest MAJOR version this reader understands. The schema's contract is
# "additive forever within a major": any 1.x profile - older or newer MINOR -
# is readable by a 1.x reader because every cross-minor change only adds
# optional keys (an absent key is its documented default). A bump to MAJOR 2
# would signal a breaking reshape, which a 1.x reader must refuse cleanly
# rather than mis-parse. :func:`validate` enforces this floor; :func:`migrate`
# is the forward hook a future reader would grow into.
SUPPORTED_MAJOR: int = 1

# The comprehension sub-block carries its own independent schema tag so the
# model-facing contract can evolve without re-versioning the whole envelope.
COMPREHENSION_SCHEMA_VERSION: str = "comprehension-1"

# The learned-overrides sub-block (``rules.overrides``, Cluster B) carries its own
# independent schema tag for the same reason. Additive: the slot was reserved at
# SCHEMA_VERSION 1.2.0; stamping a shaped ``absent`` block does NOT bump the major.
OVERRIDES_SCHEMA_VERSION: str = "overrides-1"

# The multi-template BLEND ledger (top-level ``blend`` + ``provenance.blended_shells``,
# REFLECTIONS P3) carries its own independent schema tag for the same reason. Strictly
# additive ONLY-ON-USE: unlike comprehension/overrides there is NO stamped default -
# a profile that never blended serializes without either key (not one new byte), and
# only ``profile/blend.py`` ever writes them on the first successful blend.
BLEND_SCHEMA_VERSION: str = "blend-1"


# ---------------------------------------------------------------------------
# Frozen enums
# ---------------------------------------------------------------------------
class Kind(str, Enum):
    """The format discriminator. Selects ``surface.*`` and resolver dispatch."""

    DOCX = "docx"
    PPTX = "pptx"
    XLSX = "xlsx"


# File extension for each kind's shell (``template/shell.<ext>``).
KIND_EXTENSION: dict[str, str] = {
    Kind.DOCX.value: "docx",
    Kind.PPTX.value: "pptx",
    Kind.XLSX.value: "xlsx",
}


class ResolverType(str, Enum):
    """The concrete kinds a role can resolve to.

    ``NAMED_STYLE`` (docx), ``PLACEHOLDER`` (pptx), ``CELL_STYLE`` /
    ``NAMED_RANGE`` / ``NUMBER_FORMAT`` (xlsx) are all first-class and applied by
    the generators. ``CHART_TEMPLATE`` remains a staged maturity type carried by
    the schema (status ``stub``/``best_effort``); native charts are authored
    directly today rather than via a resolver.
    """

    NAMED_STYLE = "named_style"
    PLACEHOLDER = "placeholder"
    CELL_STYLE = "cell_style"
    NUMBER_FORMAT = "number_format"
    NAMED_RANGE = "named_range"
    CHART_TEMPLATE = "chart_template"


# Which resolver types are legal for which kind (the only dispatch gate).
LEGAL_RESOLVER_TYPES: dict[str, frozenset[str]] = {
    Kind.DOCX.value: frozenset(
        {
            ResolverType.NAMED_STYLE.value,
            ResolverType.CHART_TEMPLATE.value,
        }
    ),
    Kind.PPTX.value: frozenset(
        {
            ResolverType.PLACEHOLDER.value,
            ResolverType.NAMED_STYLE.value,
            ResolverType.CHART_TEMPLATE.value,
        }
    ),
    Kind.XLSX.value: frozenset(
        {
            ResolverType.CELL_STYLE.value,
            ResolverType.NUMBER_FORMAT.value,
            ResolverType.NAMED_RANGE.value,
            ResolverType.CHART_TEMPLATE.value,
        }
    ),
}


class Status(str, Enum):
    """Capability maturity of a role / component / specimen (§8).

    Maturity is *data*, not structure: the schema holds charts/SmartArt from day
    one with ``STUB``; generation degrades gracefully on anything below
    ``ROBUST``.
    """

    ROBUST = "robust"
    BEST_EFFORT = "best_effort"
    STUB = "stub"


class Severity(str, Enum):
    """Finding severity, shared by L0/L1/L2 QA."""

    INFO = "INFO"
    WARNING = "WARNING"
    ERROR = "ERROR"


class VerificationStatus(str, Enum):
    """Overall verification verdict stamped into ``verification.status``."""

    PASSED = "passed"
    PASSED_WITH_WARNINGS = "passed_with_warnings"
    FAILED = "failed"
    UNVERIFIED = "unverified"


class AnchorKind(str, Enum):
    """How a cover/anchor region is realized in the shell (§3.2 ``anchors``)."""

    SDT_ANCHORED = "sdt_anchored"
    PLACEHOLDER = "placeholder"
    NAMED_RANGE = "named_range"
    NONE = "NONE"


# ---------------------------------------------------------------------------
# Comprehension executor enums (Ruling A: the ONLY closed value sets the model
# may write). Every value maps to a real generator code branch. Semantic labels
# (semantic_role, region names, purpose, kind, evidence) are OPEN advisory tokens
# the generator never pattern-matches on - they are NOT enums.
# ---------------------------------------------------------------------------
class ComprehensionStatus(str, Enum):
    """State of the comprehension block.

    ``present`` - merged + validated, frozen into the profile.
    ``absent``  - no comprehension yet (today's deterministic path; the default
                  every extractor stamps, including pptx/xlsx until enriched).
    ``rejected``- a merge was attempted but failed validation (fail-closed); the
                  findings are recorded and the model must retry.
    """

    PRESENT = "present"
    ABSENT = "absent"
    REJECTED = "rejected"


class FillRule(str, Enum):
    """What to do with a discovered cover slot at generate time.

    ``in_place`` - FILL: write the bound content into the preserved anchor.
    ``clear``    - CLEAR: empty the anchor / re-arm its placeholder prompt.
    ``leave``    - KEEP: do not touch the anchor.
    """

    IN_PLACE = "in_place"
    CLEAR = "clear"
    LEAVE = "leave"


class Reconcile(str, Enum):
    """How to reconcile a preserved derived index with the new content.

    ``regenerate`` - rebuild/refresh the index from the new content's items.
    ``preserve``   - keep the index untouched.
    ``clear``      - REMOVE the orphan index block (no items feed it).
    """

    REGENERATE = "regenerate"
    PRESERVE = "preserve"
    CLEAR = "clear"


class Verdict(str, Enum):
    """The demo-vs-real classification of a region."""

    DEMO = "demo"
    REAL = "real"
    MIXED = "mixed"


class AuditVerdict(str, Enum):
    """The closed L2 visual-audit disposition the model may write (Cluster C1).

    Keyed by a ``visual_manifest.json`` ``checklist[*].id``, this is the ONLY value
    the model writes into the ``comprehension.audit`` sink. Each maps to exactly one
    short-circuit branch:

    ``PASS`` - the rendered artifact satisfies the checklist item; combined with a
               matching ``(shell_sha256, content_sha256)`` it lets a same-shell /
               same-content ``generate`` skip the L2 render+manifest round.
    ``FAIL`` - the item failed; ALWAYS forces a full L2 re-render (never
               short-circuited).
    ``NA``   - not applicable for this artifact; also forces a full L2 round (the
               short-circuit requires every current item at ``PASS``).

    Closed because each value maps to a real engine branch, mirroring
    :class:`Verdict` / :class:`ComprehensionStatus`.
    """

    PASS = "PASS"
    FAIL = "FAIL"
    NA = "NA"


class TriageDisposition(str, Enum):
    """The closed model-assisted QA-triage disposition (Cluster C2).

    The ONLY value the model writes into a ``comprehension.triage`` entry. Each maps
    to exactly one branch in ``qa.gate._apply_triage`` and, by deliberate design,
    NEITHER value can raise severity or touch anything but a WARNING:

    ``expected`` - the model confirms the matched AMBIGUOUS WARNING is an intended
                   property of this template (e.g. a full-bleed cover that the
                   edge-bleed proxy flags); demote that one WARNING to INFO.
    ``defect``   - the model agrees the WARNING is a real defect; keep it verbatim.

    There is intentionally no value that lowers an ERROR or raises severity, so a
    triage entry can NEVER mask a real ERROR. Closed because each value maps to a
    real engine branch, mirroring :class:`AuditVerdict` / :class:`Verdict`.
    """

    EXPECTED = "expected"
    DEFECT = "defect"


class FragmentKind(str, Enum):
    """Which reusable-fragment registry a comprehension proposal feeds.

    ``component`` - a single inline fragment; lands in ``profile['components']``
                    and is referenced by a ``component`` IID block.
    ``section``   - a multi-block reusable unit; lands in ``profile['sections']``
                    and is referenced by a ``section`` IID block.

    This is a closed enum because each value maps to a real engine branch (which
    registry the validated proposal is written into).
    """

    COMPONENT = "component"
    SECTION = "section"


class OverrideKind(str, Enum):
    """The closed re-point vocabulary a learned override entry may use (Cluster B).

    Each value names exactly ONE brand-safe rewrite the resolver/generate path
    already supports, so a lesson can NEVER invent a style/font/hex:

    ``reroute_role``        - re-point a role that resolves to a stub at a DIFFERENT
                              role the profile already declares; the rerouted role
                              inherits that role's existing (shell-proven) resolver
                              verbatim, still through the legal-type gate.
    ``number_format``       - swap a role's number-format mask to another mask the
                              shell's ``surface.xlsx.number_formats`` already uses.
    ``register_demo_clear`` - register a captured demo string so the residual /
                              demo-clear path removes it; only removes text.

    This is a closed enum because each value maps to a real engine branch (which
    consumer the validated override entry is routed through), mirroring
    :class:`ResolverType` and :class:`ComprehensionStatus`.
    """

    REROUTE_ROLE = "reroute_role"
    NUMBER_FORMAT = "number_format"
    REGISTER_DEMO_CLEAR = "register_demo_clear"


class OverflowCapability(str, Enum):
    """Per-format overflow detection mechanism (§6.5). docx never estimates."""

    ESTIMATOR = "estimator"  # pptx
    CELLFIT = "cellfit"  # xlsx
    RENDER = "render"  # docx (LibreOffice render is the only detector)
    NONE = "none"


# Default overflow capability per kind (the frozen §6.5 mapping).
DEFAULT_OVERFLOW_CAPABILITY: dict[str, str] = {
    Kind.DOCX.value: OverflowCapability.RENDER.value,
    Kind.PPTX.value: OverflowCapability.ESTIMATOR.value,
    Kind.XLSX.value: OverflowCapability.CELLFIT.value,
}

# The L0 invariant ids every profile declares it wants enforced (§3.2 qa).
# NOTE: ``lists_use_named_numbering`` was removed (staged) - it was advertised in
# every profile but enforced nowhere (no checker, not in ``registry.CHECKS``), so a
# profile must not claim it. Re-add it here ONLY together with a real checker that
# fails on a direct ``w:numPr`` not backed by a named numbering definition.
DEFAULT_L0_INVARIANTS: tuple[str, ...] = (
    "every_role_resolves",
    "resolver_targets_exist",
    "no_literal_markdown",
    "no_residual_template_text",
    # Fail-closed membership of every load-bearing comprehension ref against the
    # surfaced deterministic inventories. Wired into ``run_qa`` next to
    # ``resolver_targets_exist`` in the SAME change that added this id (the
    # invariant list is documentation; ``run_qa`` is the enforcement). No-ops on
    # a profile with no (or an absent) comprehension block, so it is safe for the
    # model-free CI path and for pptx/xlsx.
    "comprehension_targets_exist",
    # Fail-closed membership of every persisted L2 visual-AUDIT verdict key against
    # the profile-derived visual checklist (Cluster C1). Wired into ``run_qa`` next
    # to ``comprehension_targets_exist`` in the SAME change that added the audit sink
    # (the invariant list is documentation; ``run_qa`` / ``check_audit_targets`` is
    # the enforcement). ERRORs on an audit key that is not a current checklist id and
    # rejects-never-skips on an empty derived checklist. No-ops on a profile with no
    # (or an absent) comprehension block, so it is safe for the model-free CI path
    # and for all three formats.
    "audit_targets_exist",
    # Fail-closed membership of every model-assisted QA-TRIAGE entry against the
    # closed eligible-check set + uniqueness of its (check, location) pair (Cluster
    # C2). Wired into ``run_qa`` next to ``audit_targets_exist`` in the SAME change
    # that added the triage sink (the invariant list is documentation; ``run_qa`` /
    # ``check_triage_targets`` is the enforcement). ERRORs on a triage entry naming a
    # non-eligible check or a duplicate (check, location). No-ops on a profile with no
    # (or an absent) comprehension block, so it is safe for the model-free CI path and
    # for all three formats. A triage entry can NEVER demote an ERROR: the eligible
    # set is WARNING-only and ``_apply_triage`` guards on ``severity == WARNING``.
    "triage_targets_exist",
    # Fail-closed membership of every LEARNED override target against the surfaced
    # deterministic inventories (Cluster B). Wired into ``run_qa`` next to
    # ``comprehension_targets_exist`` in the SAME change that added the override
    # consumer (the invariant list is documentation; ``run_qa`` /
    # ``check_override_targets`` is the enforcement). ERRORs on a now-missing reroute
    # target / mask / demo string and rejects-never-skips on an empty inventory.
    # No-ops on a profile with no (or an ``absent``) overrides block, so it is safe
    # for the model-free CI path and for all three formats.
    "override_targets_exist",
    # Fail-closed integrity of every PALETTE ALIAS token (Cluster E1): a model NAMES
    # an alias for a captured palette entry and the engine mints a dotted token whose
    # ref is a BYTE-COPY of that entry's ref. The check rejects an alias whose token is
    # syntactically illegal (not a dotted role-id), collides with an existing
    # theme.palette key, or whose minted ref is not byte-identical to its declared
    # source entry's ref (the engine never authors a color). Wired into ``run_qa`` next
    # to ``check_color_token_targets`` in the SAME change that added the alias mint (the
    # invariant list is documentation; ``run_qa`` / ``check_palette_alias_targets`` is
    # the enforcement). No-ops on a profile with no (or an absent) comprehension block
    # and on one with no alias proposals, so the model-free CI path, pptx/xlsx, and the
    # frozen no-alias byte-identity all stay green.
    "palette_alias_targets_exist",
    # The destructive-action floor (§6): reconciliation must never remove a
    # preserved cover anchor / index block the deterministic layer did not also
    # classify as placeholder/demo. Enforced at generate time by the generators
    # surfacing a finding with this id; listed here so the invariant is declared.
    "no_net_structure_loss",
    # Shell-vs-output structural diffs the text scans cannot see. Wired into the
    # per-format checks in ``run_qa`` (which now receives the shell at generate
    # time): ``formula_preservation`` ERRORs when an xlsx fill erased/mutated a
    # shell formula; ``component_survival`` WARNs when a native component
    # (table/chart/list/picture) present in the shell is missing from the output.
    # Both no-op when there is no output to diff (verify time), so the model-free
    # CI path is unaffected.
    "formula_preservation",
    "component_survival",
)

# Advisory ONLY (staged, like the removed ``lists_use_named_numbering``): this
# value is stamped into every profile's ``qa`` block but is NOT enforced by any L0
# check today (no checker feeds the ``color.contrast_ratio`` helpers a fg/bg pair).
# It is retained as the declared target a future contrast check would read; do NOT
# assume the gate enforces contrast until such a check is added to ``run_qa`` and
# listed in :data:`DEFAULT_L0_INVARIANTS`.
DEFAULT_CONTRAST_MIN: float = 4.5


# ---------------------------------------------------------------------------
# Role-id helpers
# ---------------------------------------------------------------------------
# A role id is a dotted lowercase path: family[.qualifier[.qualifier...]].
# Examples: "heading.1", "paragraph", "list.bullet.1", "callout.info",
# "table.default", "cover.title", "chart.bar".
_ROLE_ID_RE = re.compile(r"^[a-z][a-z0-9]*(\.[a-z0-9]+)*$")


def role_id(family: str, *qualifiers: Any) -> str:
    """Compose a canonical role id from a family and optional qualifiers.

    ``role_id("heading", 1)`` -> ``"heading.1"``;
    ``role_id("list", "bullet", 2)`` -> ``"list.bullet.2"``;
    ``role_id("paragraph")`` -> ``"paragraph"``.

    The composed id is validated against :data:`_ROLE_ID_RE` before returning, so
    this helper can never emit an id :func:`is_valid_role_id` would reject (an
    empty family, uppercase characters, or special characters in a qualifier
    raise ``ValueError`` at the composition point instead of silently failing a
    later role lookup).
    """
    parts = [str(family)] + [str(q) for q in qualifiers if q is not None and q != ""]
    rid = ".".join(parts)
    if not is_valid_role_id(rid):
        raise ValueError(f"composed role_id {rid!r} is not valid")
    return rid


def parse_role_id(rid: str) -> tuple[str, list[str]]:
    """Split a role id into ``(family, [qualifiers])``.

    ``"list.bullet.1"`` -> ``("list", ["bullet", "1"])``.
    """
    parts = rid.split(".")
    return parts[0], parts[1:]


def is_valid_role_id(rid: str) -> bool:
    """Return True if ``rid`` is a syntactically valid role id."""
    return bool(_ROLE_ID_RE.match(rid or ""))


# ---------------------------------------------------------------------------
# Envelope construction
# ---------------------------------------------------------------------------
def build_envelope(
    kind: str,
    identity: dict,
    *,
    extractor_version: str = "0.1.0",
    extracted_at: Optional[str] = None,
    source_template: Optional[dict] = None,
    theme: Optional[dict] = None,
    roles: Optional[dict] = None,
    surface: Optional[dict] = None,
    structure: Optional[dict] = None,
) -> dict:
    """Build a minimal but schema-valid ``profile.json`` envelope dict.

    This is the canonical constructor every extractor calls; it guarantees the
    frozen top-level keys exist with the right defaults so downstream code never
    has to guard for a missing section.

    Args:
        kind: one of :class:`Kind` values (``"docx"``/``"pptx"``/``"xlsx"``).
        identity: at minimum ``{"name": ...}``; ``display_name``/``locale``/
            ``tags`` are filled with defaults if absent.
        extractor_version: stamped into ``provenance.extractor_version``.
        extracted_at: ISO-8601 timestamp; caller supplies (kept out of this pure
            helper so it stays deterministic when given one).
        source_template: ``{"filename", "sha256"}`` of the source, or None.
        theme: a pre-built ``theme`` block, or None for an empty skeleton.
        roles: a pre-built ``roles`` registry, or None for an empty one.
        surface: the type-specific ``surface`` block, or None.
        structure: the optional ordered-skeleton block
            (``{"ordered": bool, "skeleton": [...]}``), or None to omit it. Additive
            since schema 1.1.0; absent on older profiles and never required by
            :func:`validate`.

    Returns:
        A dict ready to be augmented and saved. It already passes the
        structural part of :func:`validate` (semantic gaps like empty roles are
        reported as warnings by the caller's verify step, not here).

    Raises:
        ValueError: if ``kind`` is not a recognized :class:`Kind`.
    """
    if kind not in KIND_EXTENSION:
        raise ValueError(
            f"unknown kind {kind!r}; expected one of {list(KIND_EXTENSION)}"
        )
    ext = KIND_EXTENSION[kind]

    ident = {
        "name": identity["name"],
        "display_name": identity.get("display_name", identity["name"]),
        "locale": identity.get("locale", "en-US"),
        "tags": list(identity.get("tags", [])),
    }

    envelope: dict = {
        "$schema": SCHEMA_ID,
        "schema_version": SCHEMA_VERSION,
        "kind": kind,
        "identity": ident,
        "provenance": {
            "extractor_version": extractor_version,
            "extracted_at": extracted_at,
            "source_template": source_template or {"filename": None, "sha256": None},
            "shell": {"path": f"template/shell.{ext}", "sha256": None},
            "ooxml_parts_seen": [],
        },
        "theme": theme or _empty_theme(),
        "roles": roles or {"_index": []},
        "assets": {},
        "components": {},
        "sections": {},
        "rules": {
            "auto_derived": {},
            # Additive since 1.2.0 (the slot was reserved here). Always present and
            # ``absent`` by default so the resolver takes ZERO new branches and bytes
            # stay byte-identical until ``learn`` (B3) populates a sha-frozen lesson.
            "overrides": empty_overrides(),
        },
        "anchors": {},
        "structure": structure if structure is not None else _empty_structure(),
        # Additive since 1.2.0. Always present and ``absent`` by default so the
        # deterministic path is the ground truth until ``comprehend`` runs.
        "comprehension": empty_comprehension(),
        "qa": {
            "l0_invariants": list(DEFAULT_L0_INVARIANTS),
            "overflow_capability": DEFAULT_OVERFLOW_CAPABILITY[kind],
            "color_contrast_min": DEFAULT_CONTRAST_MIN,
        },
        "verification": {
            "status": VerificationStatus.UNVERIFIED.value,
            "checked_at": None,
            "roles_verified": 0,
            "roles_total": 0,
            "warnings": [],
        },
        "surface": surface or {kind: {}},
    }
    return envelope


def _empty_theme() -> dict:
    """Return an empty-but-shaped ``theme`` block."""
    return {
        "colors": {},
        "palette_roles": {},
        # Additive, optional. The brand PALETTE: a template-derived map keyed by a
        # theme slot token or ``hex:RRGGBB``. Empty until ``capture_palette`` runs;
        # never load-bearing (so no SCHEMA_VERSION bump - it is an additive optional
        # key, readable by every 1.x reader as its documented default ``{}``).
        "palette": {},
        "fonts": {
            "major": {"latin": None, "fallback": None},
            "minor": {"latin": None, "fallback": None},
        },
        "embedded_fonts": [],
    }


def _empty_structure() -> dict:
    """Return an empty-but-shaped ``structure`` block (no detected regions).

    Additive since schema 1.1.0. ``ordered`` defaults to True (the top-level region
    order is always meaningful); ``skeleton`` is empty until the extractor populates
    it with the regions actually present in the template.
    """
    return {"ordered": True, "skeleton": []}


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------
# Top-level keys every envelope must carry.
REQUIRED_TOP_KEYS: tuple[str, ...] = (
    "schema_version",
    "kind",
    "identity",
    "provenance",
    "theme",
    "roles",
    "rules",
    "anchors",
    "qa",
    "verification",
    "surface",
)


def validate(profile: dict) -> list[str]:
    """Return a list of human-readable problems with ``profile`` (``[]`` == ok).

    This is a *structural* checker - it verifies the frozen vocabulary is
    respected (keys present, enum values legal, resolver types legal for the
    kind, role ids well-formed, the shell path matches the kind, the surface
    block matches the kind). It does NOT judge brand quality (that is the QA
    gate's job) and never raises - it returns the problems so callers can decide
    severity.

    Each returned string is prefixed with a dotted path so the source of the
    problem is obvious (e.g. ``"roles.heading.1.resolver.type: ..."``).
    """
    problems: list[str] = []

    if not isinstance(profile, dict):
        return ["<root>: profile must be a JSON object"]

    for key in REQUIRED_TOP_KEYS:
        if key not in profile:
            problems.append(f"<root>: missing required key {key!r}")

    # schema_version (semver-ish, only the X.Y.Z shape is enforced here).
    sv = profile.get("schema_version")
    if sv is not None:
        m = re.match(r"^(\d+)\.\d+\.\d+", str(sv))
        if not m:
            problems.append(f"schema_version: not semver: {sv!r}")
        elif int(m.group(1)) > SUPPORTED_MAJOR:
            # A newer MAJOR is a breaking reshape this reader cannot trust. Return
            # the single actionable message NOW instead of letting the foreign
            # shape scatter a pile of confusing per-field enum errors below.
            return [
                f"schema_version {sv!r} has major {int(m.group(1))} newer than this "
                f"reader supports (max {SUPPORTED_MAJOR}); upgrade brand-docs or "
                f"re-extract the template"
            ]

    # kind discriminator
    kind = profile.get("kind")
    if kind not in KIND_EXTENSION:
        problems.append(f"kind: must be one of {list(KIND_EXTENSION)}, got {kind!r}")
        # Without a valid kind, kind-dependent checks below are meaningless.
        kind = None

    # identity
    ident = profile.get("identity")
    if not isinstance(ident, dict) or not ident.get("name"):
        problems.append("identity.name: required non-empty string")

    # provenance shell path must match the kind extension.
    prov = profile.get("provenance") or {}
    shell = (prov.get("shell") or {}) if isinstance(prov, dict) else {}
    if kind is not None:
        want = f"template/shell.{KIND_EXTENSION[kind]}"
        if shell.get("path") not in (None, want):
            problems.append(
                f"provenance.shell.path: must be {want!r} for kind {kind!r}, "
                f"got {shell.get('path')!r}"
            )

    # theme block shape
    theme = profile.get("theme")
    if not isinstance(theme, dict) or "colors" not in theme:
        problems.append("theme: must be an object with a 'colors' map")

    # theme.palette (optional, additive). Absent is fine; present must be
    # well-shaped (shape-only - membership of a color token against the palette is
    # the fail-closed QA check's job, not this structural validator's).
    if isinstance(theme, dict):
        problems.extend(_validate_palette(theme.get("palette")))

    # roles registry
    problems.extend(_validate_roles(profile.get("roles"), kind))

    # structure block (optional, additive since 1.1.0). Absent is fine; present
    # must be well-shaped.
    problems.extend(_validate_structure(profile.get("structure")))

    # comprehension block (optional, additive since 1.2.0). Absent is fine and is
    # the default; present must be well-shaped (shape-only - membership of refs
    # against the surfaced inventories is the fail-closed QA check's job, not
    # this structural validator's). NEVER required.
    problems.extend(_validate_comprehension(profile.get("comprehension")))

    # learned overrides block (optional, additive since 1.2.0 - the reserved
    # ``rules.overrides`` slot). Absent / ``{}`` / ``status='absent'`` is fine and is
    # the default; present must be well-shaped (shape-only - membership of a reroute
    # target / mask / demo string against the surfaced shell inventories is the
    # fail-closed ``check_override_targets`` job, not this structural validator's).
    # NEVER required.
    problems.extend(_validate_overrides((profile.get("rules") or {}).get("overrides")))

    # blended-shell provenance + blend ledger (optional, additive ONLY-ON-USE -
    # REFLECTIONS P3). Absent is fine and is the default for every never-blended
    # profile; present must be well-shaped and internally consistent (every ledger
    # sha a member of the recorded blended shells). Shape-only here - that the
    # recorded secondary binaries still hash to their recorded shas is the
    # fail-closed ``check_blend_shell_provenance`` QA check's job. NEVER required.
    problems.extend(
        _validate_blended_shells(
            (prov.get("blended_shells") if isinstance(prov, dict) else None), kind
        )
    )
    problems.extend(_validate_blend(profile.get("blend"), profile))

    # components / sections reusable-fragment registries. Each entry must be a dict
    # carrying a ``blocks`` list (the primitive template ``expand_components``
    # inlines). Shape-only - the block CONTENTS are enforced by ``block_from_dict``
    # at expansion time, not duplicated here. A malformed entry is surfaced now
    # (fail-closed) rather than blowing up later inside the expander.
    problems.extend(
        _validate_fragment_registry(profile.get("components"), "components")
    )
    problems.extend(_validate_fragment_registry(profile.get("sections"), "sections"))

    # surface must contain exactly the one kind sub-block.
    surface = profile.get("surface")
    if isinstance(surface, dict) and kind is not None:
        if kind not in surface:
            problems.append(f"surface: missing the {kind!r} sub-block")
        extra = [k for k in surface if k in KIND_EXTENSION and k != kind]
        if extra:
            problems.append(
                f"surface: must hold only {kind!r}; found foreign block(s) {extra}"
            )
    elif not isinstance(surface, dict):
        problems.append("surface: must be an object")

    # qa.overflow_capability must be a legal enum value.
    qa = profile.get("qa") or {}
    oc = qa.get("overflow_capability") if isinstance(qa, dict) else None
    legal_oc = {e.value for e in OverflowCapability}
    if oc is not None and oc not in legal_oc:
        problems.append(f"qa.overflow_capability: illegal value {oc!r}")

    # verification.status enum
    ver = profile.get("verification") or {}
    vs = ver.get("status") if isinstance(ver, dict) else None
    legal_vs = {e.value for e in VerificationStatus}
    if vs is not None and vs not in legal_vs:
        problems.append(f"verification.status: illegal value {vs!r}")

    # Intra-profile consistency (no shell I/O): a resolver target that the profile
    # *itself* contradicts is caught here. A placeholder resolver's ``layout`` must
    # exist in ``surface.pptx.layouts``; a named_range resolver's ``name`` must
    # exist in ``surface.xlsx.named_regions``. This catches a fabricated profile
    # (e.g. a pptx extractor that invents "Title Slide" while surface lists other
    # layouts) without opening the template.
    problems.extend(_validate_resolver_consistency(profile, kind))

    return problems


def migrate(profile: dict) -> dict:
    """Bring a loaded ``profile`` dict up to the reader's current schema.

    The schema's contract is **additive forever within a major**: across MINOR
    versions a profile only ever *gains* optional keys, and an absent key is
    always its documented default. A reader of MAJOR ``SUPPORTED_MAJOR`` can
    therefore consume any ``SUPPORTED_MAJOR``.x profile - older or newer MINOR -
    without rewriting it, which is why this is the **identity** today: it returns
    the profile unchanged.

    This function exists as the documented forward seam. Real cross-major
    migration (a MAJOR 1 -> 2 reshape, only ever reached for profiles whose major
    is <= :data:`SUPPORTED_MAJOR`) would live here, branching on
    ``profile["schema_version"]``. Profiles whose major is *newer* than this
    reader never reach migration - :func:`validate` refuses them with a single
    clear message first. Wiring this into the load path (see
    ``store.load_profile``) is behaviour-identical while the body is identity, so
    callers can adopt the seam now and gain real migration for free later.
    """
    return profile


def _validate_resolver_consistency(profile: dict, kind: Optional[str]) -> list[str]:
    """Cross-check each resolver target against the profile's own surface map."""
    problems: list[str] = []
    roles = profile.get("roles")
    if not isinstance(roles, dict):
        return problems
    surface = profile.get("surface")
    sub = surface.get(kind) if isinstance(surface, dict) and kind else None
    if not isinstance(sub, dict):
        return problems

    layouts = sub.get("layouts") if isinstance(sub.get("layouts"), dict) else {}
    named_regions = (
        sub.get("named_regions") if isinstance(sub.get("named_regions"), dict) else {}
    )

    for rid, entry in roles.items():
        if rid == "_index" or not isinstance(entry, dict):
            continue
        resolver = entry.get("resolver")
        if not isinstance(resolver, dict):
            continue
        rtype = resolver.get("type")
        if rtype == ResolverType.PLACEHOLDER.value:
            layout = resolver.get("layout")
            # Only enforce when the surface actually declares layouts (else we
            # cannot prove absence - that is the shell-backed check's job).
            if layout is not None and layouts and layout not in layouts:
                problems.append(
                    f"roles.{rid}.resolver.layout: {layout!r} not in surface.{kind}.layouts "
                    f"(have {sorted(layouts)})"
                )
        elif rtype == ResolverType.NAMED_RANGE.value:
            name = resolver.get("name")
            if name is not None and named_regions and name not in named_regions:
                problems.append(
                    f"roles.{rid}.resolver.name: named range {name!r} not in "
                    f"surface.{kind}.named_regions (have {sorted(named_regions)})"
                )
        elif rtype == ResolverType.NUMBER_FORMAT.value:
            # A number_format resolver may only carry a mask the template actually
            # uses (surface.xlsx.number_formats), so the resolver can never fabricate
            # a format - the same brand guarantee as the named_range check above.
            mask = resolver.get("number_format")
            known = {
                f.get("format")
                for f in (sub.get("number_formats") or [])
                if isinstance(f, dict)
            }
            if mask is not None and known and mask not in known:
                problems.append(
                    f"roles.{rid}.resolver.number_format: mask {mask!r} not in "
                    f"surface.{kind}.number_formats"
                )
    return problems


def _validate_roles(roles: Any, kind: Optional[str]) -> list[str]:
    """Validate the ``roles`` registry against the frozen vocabulary."""
    problems: list[str] = []
    if not isinstance(roles, dict):
        return ["roles: must be an object"]

    index = roles.get("_index")
    if index is not None and not isinstance(index, list):
        problems.append("roles._index: must be a list of role ids")
        index = None

    legal_types = LEGAL_RESOLVER_TYPES.get(kind, set()) if kind else set()
    legal_status = {e.value for e in Status}

    for rid, entry in roles.items():
        if rid == "_index":
            continue
        if not is_valid_role_id(rid):
            problems.append(f"roles.{rid}: invalid role id syntax")
        if not isinstance(entry, dict):
            problems.append(f"roles.{rid}: must be an object")
            continue
        resolver = entry.get("resolver")
        if not isinstance(resolver, dict):
            problems.append(f"roles.{rid}.resolver: required object")
        else:
            rtype = resolver.get("type")
            if rtype not in {e.value for e in ResolverType}:
                problems.append(f"roles.{rid}.resolver.type: unknown {rtype!r}")
            elif kind is not None and rtype not in legal_types:
                problems.append(
                    f"roles.{rid}.resolver.type: {rtype!r} is not legal for "
                    f"kind {kind!r} (legal: {sorted(legal_types)})"
                )
        status = entry.get("status")
        if status is not None and status not in legal_status:
            problems.append(f"roles.{rid}.status: illegal value {status!r}")

        # usage is optional (additive since 1.1.0). Profiles lacking it are NOT
        # rejected; if present it must be a well-shaped object.
        usage = entry.get("usage")
        if usage is not None:
            problems.extend(_validate_usage(f"roles.{rid}.usage", usage))

        # _index, if present, should list every concrete role.
        if index is not None and rid not in index:
            problems.append(f"roles._index: missing role {rid!r}")

    return problems


# Legal vocabulary for the per-artifact ``usage`` annotation (additive, 1.1.0).
USAGE_SCOPES: frozenset[str] = frozenset({"cover", "toc", "body", "anywhere"})
USAGE_PLACEMENTS: frozenset[str] = frozenset({"structural", "freeform"})

# Region NAMES are OPEN tokens (schema 1.2.0). The docx derivation still uses the
# canonical {cover,toc,body} trio and pptx/xlsx may use their own honest names
# (agenda, appendix, sheet, ...); none of these is a frozen MATCHING rule. A
# region name is validated for *syntax only* (same shape as a role id) so a typo
# is caught without committing a per-format word-list (which would re-commit the
# lexicon sin). The legacy trio is kept as documentation of the docx convention,
# never as the validation gate.
STRUCTURE_REGIONS: frozenset[str] = frozenset({"cover", "toc", "body"})
# The four boolean/closed attributes the generator is allowed to branch on. Same
# four for every format; this REPLACES per-format region word-lists.
STRUCTURE_REGION_ATTRS: tuple[str, ...] = ("freeform", "demo", "ordered", "required")


def is_valid_region_token(name: Any) -> bool:
    """Return True if ``name`` is a syntactically valid OPEN region token.

    A region token has the same shape as a role id (dotted lowercase path), so a
    typo is caught structurally without freezing a per-format vocabulary.
    """
    return isinstance(name, str) and bool(_ROLE_ID_RE.match(name))


def _validate_usage(path: str, usage: Any) -> list[str]:
    """Validate a per-artifact ``usage`` object (never required; structural only)."""
    problems: list[str] = []
    if not isinstance(usage, dict):
        return [f"{path}: must be an object"]
    scope = usage.get("scope")
    if scope is not None and scope not in USAGE_SCOPES:
        problems.append(
            f"{path}.scope: illegal value {scope!r} (legal: {sorted(USAGE_SCOPES)})"
        )
    placement = usage.get("placement")
    if placement is not None and placement not in USAGE_PLACEMENTS:
        problems.append(
            f"{path}.placement: illegal value {placement!r} (legal: {sorted(USAGE_PLACEMENTS)})"
        )
    order = usage.get("order")
    if order is not None and not isinstance(order, int):
        problems.append(f"{path}.order: must be an int or null, got {order!r}")
    return problems


def _validate_structure(structure: Any) -> list[str]:
    """Validate the optional ``structure`` block (absent is fine).

    Checks shape only: ``ordered`` is a bool, ``skeleton`` is a list of region
    descriptors with a legal ``region`` and an int ``order``. Never required; a
    1.0.0 profile that omits the key entirely produces no problems.
    """
    if structure is None:
        return []
    problems: list[str] = []
    if not isinstance(structure, dict):
        return ["structure: must be an object"]
    ordered = structure.get("ordered")
    if ordered is not None and not isinstance(ordered, bool):
        problems.append(f"structure.ordered: must be a bool, got {ordered!r}")
    skeleton = structure.get("skeleton")
    if skeleton is None:
        return problems
    if not isinstance(skeleton, list):
        return problems + ["structure.skeleton: must be a list"]
    for i, region in enumerate(skeleton):
        if not isinstance(region, dict):
            problems.append(f"structure.skeleton[{i}]: must be an object")
            continue
        rname = region.get("region")
        # OPEN token (schema 1.2.0): validated for SYNTAX only, never against a
        # frozen per-format word-list. The generator branches on the boolean
        # attributes below, not on the name.
        if not is_valid_region_token(rname):
            problems.append(
                f"structure.skeleton[{i}].region: not a valid region token {rname!r} "
                f"(dotted lowercase, e.g. {sorted(STRUCTURE_REGIONS)})"
            )
        order = region.get("order")
        if order is not None and not isinstance(order, int):
            problems.append(
                f"structure.skeleton[{i}].order: must be an int, got {order!r}"
            )
        for attr in STRUCTURE_REGION_ATTRS:
            val = region.get(attr)
            if val is not None and not isinstance(val, bool):
                problems.append(
                    f"structure.skeleton[{i}].{attr}: must be a bool or null, got {val!r}"
                )
    return problems


# The closed ``where`` vocabulary for a ``theme.palette`` entry's provenance. A
# palette provenance fact records WHICH observed source named the color; nothing
# outside this set may appear. The frozen mirror of
# ``docx.typography.PALETTE_WHERE`` (kept here so the validator has no format
# import). ``palette_role`` is the only NON-authoritative source. ``palette.alias``
# is the ENGINE-minted source for an alias token (Cluster E1): the model NAMES an
# alias for a captured entry, the engine mints a dotted token whose ref is a
# byte-copy of that entry's ref, stamping this provenance ``where`` with the source
# palette key as ``detail`` (so an alias is never confused for an observed capture).
PALETTE_WHERE: frozenset[str] = frozenset(
    {"palette_role", "role.appearance", "run.color", "link.color", "palette.alias"}
)
# The COARSE frequency buckets a palette entry may carry (never raw counts).
PALETTE_FREQUENCIES: frozenset[str] = frozenset({"dominant", "accent", "rare"})


def _validate_palette(palette: Any) -> list[str]:
    """Validate the optional ``theme.palette`` block (absent is fine).

    SHAPE-ONLY, consistent with the rest of this structural validator. Each entry
    (keyed by a template-derived id - a theme slot token or ``hex:RRGGBB``) must
    carry:
      - ``ref``: an object with a closed ``kind`` (``theme`` -> ``theme`` slot
        string; ``hex`` -> ``hex`` string);
      - ``provenance`` (if present): a list of ``{where, detail}`` facts whose
        ``where`` is in the closed :data:`PALETTE_WHERE` vocabulary;
      - ``frequency`` (if present): one of :data:`PALETTE_FREQUENCIES`;
      - ``name`` / ``purpose`` / ``use_when`` (if present): advisory strings or
        null (the model fills these; never gated for VALUE).

    It deliberately does NOT check that a color TOKEN referenced elsewhere is a
    key of the palette - that is the fail-closed ``check_color_token_targets`` QA
    check. NEVER required: an absent / empty palette yields no problems.
    """
    if palette is None:
        return []
    if not isinstance(palette, dict):
        return ["theme.palette: must be an object mapping color id -> entry"]
    problems: list[str] = []
    for key, entry in palette.items():
        path = f"theme.palette.{key}"
        if not isinstance(key, str) or not key:
            problems.append(f"{path}: palette key must be a non-empty string")
        if not isinstance(entry, dict):
            problems.append(f"{path}: must be an object")
            continue
        ref = entry.get("ref")
        if not isinstance(ref, dict):
            problems.append(f"{path}.ref: required object")
        else:
            rkind = ref.get("kind")
            if rkind == "theme":
                if not isinstance(ref.get("theme"), str) or not ref.get("theme"):
                    problems.append(f"{path}.ref.theme: required non-empty string")
            elif rkind == "hex":
                if not isinstance(ref.get("hex"), str) or not ref.get("hex"):
                    problems.append(f"{path}.ref.hex: required non-empty string")
            else:
                problems.append(
                    f"{path}.ref.kind: illegal value {rkind!r} (legal: 'theme' | 'hex')"
                )
        provenance = entry.get("provenance")
        if provenance is not None:
            if not isinstance(provenance, list):
                problems.append(f"{path}.provenance: must be a list")
            else:
                for i, fact in enumerate(provenance):
                    fpath = f"{path}.provenance[{i}]"
                    if not isinstance(fact, dict):
                        problems.append(f"{fpath}: must be an object")
                        continue
                    where = fact.get("where")
                    if where not in PALETTE_WHERE:
                        problems.append(
                            f"{fpath}.where: illegal value {where!r} "
                            f"(legal: {sorted(PALETTE_WHERE)})"
                        )
                    detail = fact.get("detail")
                    if detail is not None and not isinstance(detail, str):
                        problems.append(
                            f"{fpath}.detail: must be a string or null, got {detail!r}"
                        )
        freq = entry.get("frequency")
        if freq is not None and freq not in PALETTE_FREQUENCIES:
            problems.append(
                f"{path}.frequency: illegal value {freq!r} "
                f"(legal: {sorted(PALETTE_FREQUENCIES)})"
            )
        for advisory in ("name", "purpose", "use_when"):
            val = entry.get(advisory)
            if val is not None and not isinstance(val, str):
                problems.append(
                    f"{path}.{advisory}: must be a string or null, got {val!r}"
                )
    return problems


# ---------------------------------------------------------------------------
# Comprehension block (additive, schema 1.2.0)
# ---------------------------------------------------------------------------
# The single canonical, additive top-level sink for model output (Ruling B). One
# writer (``profile/comprehension.py::merge``), one home. ``roles[*].usage`` /
# ``structure.skeleton`` / ``anchors.*`` are DERIVED from this block at merge
# time, never written independently.
COMPREHENSION_STATUSES: frozenset[str] = frozenset(e.value for e in ComprehensionStatus)
FILL_RULES: frozenset[str] = frozenset(e.value for e in FillRule)
RECONCILE_RULES: frozenset[str] = frozenset(e.value for e in Reconcile)
VERDICTS: frozenset[str] = frozenset(e.value for e in Verdict)
# Closed disposition the L2 model writes into ``comprehension.audit`` (Cluster C1).
AUDIT_VERDICTS: frozenset[str] = frozenset(e.value for e in AuditVerdict)
# Closed disposition the model writes into a ``comprehension.triage`` entry (C2).
TRIAGE_DISPOSITIONS: frozenset[str] = frozenset(e.value for e in TriageDisposition)
# The ONLY check ids a triage entry may name (Cluster C2). DELIBERATELY restricted
# to WARNING-only, genuinely-ambiguous proxies: a full-bleed cover the edge-bleed
# proxy flags, a deliberately blank section page, a legitimately-dropped native
# component family. NONE of these is an ERROR-emitting check, so a triage entry can
# never even be aimed at an ERROR (it is rejected at merge as a non-member) - the
# belt to ``_apply_triage``'s ``severity == WARNING`` suspenders. Closed: a check id
# joins this set ONLY when its findings are WARNING-only and genuinely ambiguous.
AMBIGUOUS_TRIAGE_CHECKS: frozenset[str] = frozenset(
    {"visual.blank_page", "visual.edge_bleed", "component_survival"}
)
FRAGMENT_KINDS: frozenset[str] = frozenset(e.value for e in FragmentKind)
# Closed enum for a caption index's content TARGET: which captionable kind feeds it
# (a list-of-tables is fed by table captions, a list-of-figures by figure captions).
# Optional and additive; lets the generator map a caption's ``target`` to the index's
# opaque ``seq_id`` brand-agnostically, without a language heuristic on the seq name.
CAPTION_TARGETS: frozenset[str] = frozenset({"table", "figure"})

# Closed re-point vocabulary for a learned override entry (Cluster B).
OVERRIDE_KINDS: frozenset[str] = frozenset(e.value for e in OverrideKind)

# The UNAMBIGUOUS recurring check-ids the deterministic ``learn`` step distills a
# lesson from. Each maps cleanly to exactly one closed override kind, so a recurrence
# of one of these can be turned into a brand-safe re-point without a model. (The
# number-format rejection surfaces as a ``resolver_targets_exist`` finding -
# checks_deterministic.py:224 - so it is covered by that id, not a fourth one.)
LEARNABLE_CHECKS: frozenset[str] = frozenset(
    {
        "resolver_targets_exist",
        "style_fallback",
        "no_residual_template_text",
    }
)


def empty_comprehension() -> dict:
    """Return an empty-but-shaped, ``absent`` comprehension block.

    This is what every extractor stamps by default (docx/pptx/xlsx) so the block
    always exists and generation can ask ``status`` without guarding for a missing
    key. ``absent`` ⇒ today's deterministic path.
    """
    return {
        "schema_version": COMPREHENSION_SCHEMA_VERSION,
        "status": ComprehensionStatus.ABSENT.value,
        "generated_by": None,
        "source_shell_sha256": None,
        "confidence": 0.0,
        "cover_slots": {},
        "conventions": {"indexes": [], "sections": []},
        "role_annotations": {},
        "demo_classification": {"regions": []},
        # Additive: model-proposed reusable fragments. On a clean merge each entry
        # is DERIVED into profile['components'] / profile['sections'] (the existing
        # registries expand_components inlines). Empty is the norm.
        "fragments": [],
        # Additive (model-driven color): the model NAMES each captured palette color
        # (a key of ``theme.palette``) - name / purpose / use_when / semantic_role -
        # without ever authoring a real color. On a clean merge the advisory fields
        # are mirrored onto ``theme.palette[key]`` (the derived sink). Empty is the
        # norm; the model never writes ``ref`` / a hex.
        "palette_annotations": {},
        # Additive (Cluster C1): the persisted L2 visual-audit verdict map, keyed by a
        # ``visual_manifest.json`` checklist id. Each value is { verdict: PASS|FAIL|NA,
        # evidence?, shell_sha256?, content_sha256? }. The model never writes a brand
        # value here - only a closed disposition against a structural checklist id it
        # did not author. Empty is the norm; a same-shell/same-content generate
        # short-circuits L2 only when EVERY current checklist id PASSes at a matching
        # sha pair. (Omitting this key would silently drop the verdict in
        # ``_canonicalize``, which rebuilds from this empty block.)
        "audit": {},
        # Additive (Cluster C2): the model-assisted QA-triage list. Each entry is
        # { check, location, disposition: expected|defect, evidence? } naming exactly
        # one AMBIGUOUS WARNING the model judged. On a clean merge ``_canonicalize``
        # copies it sorted; ``qa.gate._apply_triage`` is the SOLE consumer, demoting a
        # matched WARNING to INFO iff disposition==expected. The model never writes a
        # brand value here, only a closed disposition + advisory evidence against a
        # closed (check, location) pair; an ERROR can NEVER be demoted (no enum value
        # lowers an ERROR, and the eligible-check set is WARNING-only). Empty is the
        # norm. (Omitting this key would silently drop triage in ``_canonicalize``,
        # which rebuilds from this empty block.)
        "triage": [],
        # Additive (Cluster E2): the model-adjudicated faked-heading promotions. Each
        # entry is { pseudo_heading_ref, target_role_id } naming a SURFACED
        # ``pseudo_heading`` fact (a body-style size/color outlier the detector found)
        # and a DECLARED heading role. The model authors no size/color: on a clean
        # merge ``_derive_promote_appearance`` COPIES the captured outlier size/color
        # from the detector fact onto ``roles[target_role_id].appearance`` (re-validated
        # shell-backed by ``check_appearance_targets``). Empty is the norm. (Omitting
        # this key would silently drop the promotions in ``_canonicalize``, which
        # rebuilds from this empty block.)
        "promote_appearance": [],
    }


def empty_overrides() -> dict:
    """Return an empty-but-shaped, ``absent`` learned-overrides block (Cluster B).

    This is what every extractor stamps into ``rules.overrides`` by default
    (docx/pptx/xlsx), replacing the reserved ``{}`` so the block always exists and
    the resolver/store can ask ``status`` / ``source_shell_sha256`` without guarding
    for a missing key. ``absent`` (or a ``source_shell_sha256`` that no longer
    matches the live shell) ⇒ the resolver takes ZERO new branches and generated
    bytes are byte-identical to the no-overrides path.

    Mirrors :func:`empty_comprehension` one-for-one (same freeze fields:
    ``status`` / ``source_shell_sha256`` / ``confidence``) so the same
    sha-bound presence test applies (``store.overrides_are_present``). The three
    containers are the closed-kind sinks the deterministic ``learn`` writer fills;
    ``provenance`` is a flat record of which ``(check, location, recurred_runs)``
    produced each entry.
    """
    return {
        "schema_version": OVERRIDES_SCHEMA_VERSION,
        "status": ComprehensionStatus.ABSENT.value,
        "source_shell_sha256": None,
        "generated_by": None,
        "confidence": 0.0,
        # reroute_roles: { <requested_role_id>: <target_role_id> }
        "reroute_roles": {},
        # number_format_swaps: { <role_id>: <mask> }
        "number_format_swaps": {},
        # demo_clears: [ <captured demo string> ]
        "demo_clears": [],
        # provenance: { <opaque entry key>: { check, location, recurred_runs } }
        "provenance": {},
    }


def _validate_comprehension(comp: Any) -> list[str]:
    """Validate the optional ``comprehension`` block (absent is fine).

    SHAPE-ONLY (Ruling A/B). This checks that:
      - ``status`` (if present) is in the closed :class:`ComprehensionStatus` enum;
      - every EXECUTOR field is in its closed enum (``fill_rule`` / ``reconcile``
        / ``verdict``);
      - the sub-collections are the right container types and their entries carry
        the load-bearing ref keys as strings;
      - advisory free-text fields (``purpose`` / ``semantic_role`` / ``kind`` /
        ``evidence`` / ``generation_rules``) are accepted as-is and NEVER gated.

    It deliberately does NOT check that a ref resolves to an inventory id - that
    is the fail-closed ``comprehension_targets_exist`` QA check (it needs the
    surfaced inventories, and it must be able to ERROR on an empty inventory,
    which a never-required structural validator must not do). NEVER required: a
    profile without the key, or with ``status='absent'``, yields no problems.
    """
    if comp is None:
        return []
    problems: list[str] = []
    if not isinstance(comp, dict):
        return ["comprehension: must be an object"]

    status = comp.get("status")
    if status is not None and status not in COMPREHENSION_STATUSES:
        problems.append(
            f"comprehension.status: illegal value {status!r} "
            f"(legal: {sorted(COMPREHENSION_STATUSES)})"
        )

    conf = comp.get("confidence")
    if conf is not None and not isinstance(conf, (int, float)):
        problems.append(
            f"comprehension.confidence: must be a number or null, got {conf!r}"
        )
    elif isinstance(conf, (int, float)) and not (0.0 <= float(conf) <= 1.0):
        problems.append(f"comprehension.confidence: must be in [0,1], got {conf!r}")

    sha = comp.get("source_shell_sha256")
    if sha is not None and not isinstance(sha, str):
        problems.append(
            f"comprehension.source_shell_sha256: must be a hex string or null, got {sha!r}"
        )

    # cover_slots: { <anchor_ref>: { fill_rule, binds_to?, semantic_role?, ... } }
    slots = comp.get("cover_slots")
    if slots is not None:
        if not isinstance(slots, dict):
            problems.append("comprehension.cover_slots: must be an object")
        else:
            for anchor_ref, slot in slots.items():
                path = f"comprehension.cover_slots.{anchor_ref}"
                if not isinstance(anchor_ref, str) or not anchor_ref:
                    problems.append(
                        f"{path}: anchor_ref key must be a non-empty string"
                    )
                if not isinstance(slot, dict):
                    problems.append(f"{path}: must be an object")
                    continue
                fr = slot.get("fill_rule")
                if fr is not None and fr not in FILL_RULES:
                    problems.append(
                        f"{path}.fill_rule: illegal value {fr!r} (legal: {sorted(FILL_RULES)})"
                    )
                bt = slot.get("binds_to")
                if bt is not None and not isinstance(bt, str):
                    problems.append(
                        f"{path}.binds_to: must be a string or null, got {bt!r}"
                    )

    # conventions.indexes / conventions.sections
    conventions = comp.get("conventions")
    if conventions is not None:
        if not isinstance(conventions, dict):
            problems.append("comprehension.conventions: must be an object")
        else:
            problems.extend(_validate_comp_indexes(conventions.get("indexes")))
            problems.extend(_validate_comp_sections(conventions.get("sections")))

    # role_annotations: { <role_id>: { purpose?, generation_rules? } } - advisory.
    annotations = comp.get("role_annotations")
    if annotations is not None:
        if not isinstance(annotations, dict):
            problems.append("comprehension.role_annotations: must be an object")
        else:
            for rid, ann in annotations.items():
                if not is_valid_role_id(rid):
                    problems.append(
                        f"comprehension.role_annotations.{rid}: not a valid role id"
                    )
                if ann is not None and not isinstance(ann, dict):
                    problems.append(
                        f"comprehension.role_annotations.{rid}: must be an object"
                    )

    # demo_classification.regions: [ { region_ref, verdict, evidence? } ]
    demo = comp.get("demo_classification")
    if demo is not None:
        if not isinstance(demo, dict):
            problems.append("comprehension.demo_classification: must be an object")
        else:
            regions = demo.get("regions")
            if regions is not None:
                if not isinstance(regions, list):
                    problems.append(
                        "comprehension.demo_classification.regions: must be a list"
                    )
                else:
                    for i, reg in enumerate(regions):
                        path = f"comprehension.demo_classification.regions[{i}]"
                        if not isinstance(reg, dict):
                            problems.append(f"{path}: must be an object")
                            continue
                        ref = reg.get("region_ref")
                        if not isinstance(ref, str) or not ref:
                            problems.append(
                                f"{path}.region_ref: required non-empty string"
                            )
                        verdict = reg.get("verdict")
                        if verdict is not None and verdict not in VERDICTS:
                            problems.append(
                                f"{path}.verdict: illegal value {verdict!r} "
                                f"(legal: {sorted(VERDICTS)})"
                            )

    # palette_annotations: { <palette_key>: { name?, purpose?, use_when?,
    # semantic_role? } } - the model NAMES a captured color. SHAPE-ONLY (advisory
    # strings); membership of the KEY against the surfaced palette inventory is the
    # fail-closed check_membership / check_color_token_targets job, not this
    # never-required structural validator's.
    problems.extend(_validate_palette_annotations(comp.get("palette_annotations")))

    # fragments: [ { ref, kind, blocks, purpose? } ] - reusable-fragment proposals.
    problems.extend(_validate_comp_fragments(comp.get("fragments")))

    # audit: { <checklist_id>: { verdict, evidence?, shell_sha256?, content_sha256? } }
    # - the persisted L2 visual-audit verdict map (Cluster C1). SHAPE-ONLY; that the
    # KEY is a real checklist id is the fail-closed ``check_audit_targets`` job.
    problems.extend(_validate_comp_audit(comp.get("audit")))

    # triage: [ { check, location, disposition, evidence? } ] - the model-assisted
    # QA-triage list (Cluster C2). SHAPE-ONLY; that the (check, location) was actually
    # emitted is the fail-closed ``check_triage_targets`` / ``check_triage`` job.
    problems.extend(_validate_comp_triage(comp.get("triage")))

    # promote_appearance: [ { pseudo_heading_ref, target_role_id } ] - the faked-heading
    # promotion list (Cluster E2). SHAPE-ONLY; that the ref was SURFACED by the detector
    # and the target is a DECLARED heading role is the fail-closed ``check_promote_appearance``
    # job (same split as triage refs vs ``check_triage``).
    problems.extend(_validate_comp_promote_appearance(comp.get("promote_appearance")))
    return problems


def _validate_overrides(overrides: Any) -> list[str]:
    """Validate the optional ``rules.overrides`` block (absent is fine).

    SHAPE-ONLY, mirroring :func:`_validate_comprehension` (Ruling A/B). This checks
    that:
      - ``status`` (if present) is in the closed :class:`ComprehensionStatus` enum
        (overrides reuse the present/absent/rejected freeze contract);
      - ``confidence`` is a number in ``[0,1]`` (or null) and
        ``source_shell_sha256`` is a hex string (or null);
      - ``reroute_roles`` / ``number_format_swaps`` are objects of string→string;
      - ``demo_clears`` is a list of strings;
      - ``provenance`` is an object.

    It deliberately does NOT check that a reroute target is a declared role, that a
    mask is one the shell uses, or that a demo string was captured - that is the
    fail-closed ``check_override_targets`` QA check (it needs the surfaced shell
    inventories and must be able to ERROR on an empty inventory, which a
    never-required structural validator must not do). NEVER required: a profile
    without the key, with ``rules.overrides == {}``, or with ``status='absent'``,
    yields no problems.
    """
    if overrides is None:
        return []
    if not isinstance(overrides, dict):
        return ["rules.overrides: must be an object"]
    # The reserved-empty slot (and the additive default before B3) is ``{}``: a
    # never-required validator must accept it as the documented "absent" default.
    if not overrides:
        return []
    problems: list[str] = []

    status = overrides.get("status")
    if status is not None and status not in COMPREHENSION_STATUSES:
        problems.append(
            f"rules.overrides.status: illegal value {status!r} "
            f"(legal: {sorted(COMPREHENSION_STATUSES)})"
        )

    conf = overrides.get("confidence")
    if conf is not None and not isinstance(conf, (int, float)):
        problems.append(
            f"rules.overrides.confidence: must be a number or null, got {conf!r}"
        )
    elif isinstance(conf, (int, float)) and not (0.0 <= float(conf) <= 1.0):
        problems.append(f"rules.overrides.confidence: must be in [0,1], got {conf!r}")

    sha = overrides.get("source_shell_sha256")
    if sha is not None and not isinstance(sha, str):
        problems.append(
            f"rules.overrides.source_shell_sha256: must be a hex string or null, "
            f"got {sha!r}"
        )

    # reroute_roles: { <requested_role_id>: <target_role_id> } - both strings.
    reroutes = overrides.get("reroute_roles")
    if reroutes is not None:
        if not isinstance(reroutes, dict):
            problems.append("rules.overrides.reroute_roles: must be an object")
        else:
            for requested, target in reroutes.items():
                path = f"rules.overrides.reroute_roles.{requested}"
                if not isinstance(requested, str) or not requested:
                    problems.append(f"{path}: role id key must be a non-empty string")
                if not isinstance(target, str) or not target:
                    problems.append(
                        f"{path}: target must be a non-empty string, got {target!r}"
                    )

    # number_format_swaps: { <role_id>: <mask> } - both strings.
    swaps = overrides.get("number_format_swaps")
    if swaps is not None:
        if not isinstance(swaps, dict):
            problems.append("rules.overrides.number_format_swaps: must be an object")
        else:
            for rid, mask in swaps.items():
                path = f"rules.overrides.number_format_swaps.{rid}"
                if not isinstance(rid, str) or not rid:
                    problems.append(f"{path}: role id key must be a non-empty string")
                if not isinstance(mask, str) or not mask:
                    problems.append(
                        f"{path}: mask must be a non-empty string, got {mask!r}"
                    )

    # demo_clears: [ <captured demo string> ].
    clears = overrides.get("demo_clears")
    if clears is not None:
        if not isinstance(clears, list):
            problems.append("rules.overrides.demo_clears: must be a list")
        else:
            for i, val in enumerate(clears):
                if not isinstance(val, str) or not val:
                    problems.append(
                        f"rules.overrides.demo_clears[{i}]: must be a non-empty "
                        f"string, got {val!r}"
                    )

    prov = overrides.get("provenance")
    if prov is not None and not isinstance(prov, dict):
        problems.append("rules.overrides.provenance: must be an object")

    return problems


# A recorded shell hash is always the full lowercase hex SHA-256.
_SHA256_HEX_RE = re.compile(r"^[0-9a-f]{64}$")


def _validate_blended_shells(entries: Any, kind: Optional[str]) -> list[str]:
    """Validate the optional ``provenance.blended_shells`` list (absent is fine).

    SHAPE-ONLY, mirroring :func:`_validate_overrides` (absent-tolerant, never
    required): each entry must carry a single-segment ``filename``, a full
    lowercase-hex ``sha256``, and the content-addressed ``path`` exactly
    ``template/blend-<sha256[:12]>.<ext>`` for the profile's kind. The list must be
    sorted by ``sha256`` with unique shas so the on-disk form is structurally
    deterministic. It deliberately does NOT re-hash the recorded binaries - that is
    the fail-closed ``check_blend_shell_provenance`` QA check's job.
    """
    if entries is None:
        return []
    if not isinstance(entries, list):
        return ["provenance.blended_shells: must be a list"]
    problems: list[str] = []
    shas: list[str] = []
    for i, entry in enumerate(entries):
        path = f"provenance.blended_shells[{i}]"
        if not isinstance(entry, dict):
            problems.append(f"{path}: must be an object")
            continue
        filename = entry.get("filename")
        if not isinstance(filename, str) or not filename:
            problems.append(f"{path}.filename: required non-empty string")
        elif "/" in filename or "\\" in filename:
            problems.append(
                f"{path}.filename: must be a single path segment, got {filename!r}"
            )
        sha = entry.get("sha256")
        if not isinstance(sha, str) or not _SHA256_HEX_RE.match(sha):
            problems.append(
                f"{path}.sha256: required 64-char lowercase hex sha256, got {sha!r}"
            )
            sha = None
        else:
            shas.append(sha)
        rel = entry.get("path")
        if sha is not None and kind in KIND_EXTENSION:
            want = f"template/blend-{sha[:12]}.{KIND_EXTENSION[kind]}"
            if rel != want:
                problems.append(f"{path}.path: must be {want!r}, got {rel!r}")
        elif not isinstance(rel, str) or not rel:
            problems.append(f"{path}.path: required non-empty string")
    if shas != sorted(shas):
        problems.append("provenance.blended_shells: entries must be sorted by sha256")
    if len(shas) != len(set(shas)):
        problems.append("provenance.blended_shells: sha256 values must be unique")
    return problems


def _validate_blend(block: Any, profile: dict) -> list[str]:
    """Validate the optional top-level ``blend`` ledger (absent is fine).

    SHAPE-ONLY plus ONE intra-profile cross-check (the peer of
    :func:`_validate_resolver_consistency`, no IO): every ``from``/``by`` sha the
    ledger names must be a member of ``provenance.blended_shells`` - a ledger can
    never cite a donor the provenance does not record. The ledger maps a dotted
    VALUE-fact path to the donor sha(s) that filled/corroborated it; the ``filled``
    value is ``{"from": <sha>}`` and the ``corroborated`` value is ``{"by":
    [<sha>, ...]}`` (non-empty, sorted, unique). ``status`` reuses the closed
    :class:`ComprehensionStatus` vocabulary (only ``present`` is ever serialized -
    a rejected blend writes NOTHING, unlike comprehension/overrides). NEVER
    required: a never-blended profile has no ``blend`` key and yields no problems.
    """
    if block is None:
        return []
    if not isinstance(block, dict):
        return ["blend: must be an object"]
    problems: list[str] = []

    sv = block.get("schema_version")
    if sv != BLEND_SCHEMA_VERSION:
        problems.append(
            f"blend.schema_version: must be {BLEND_SCHEMA_VERSION!r}, got {sv!r}"
        )
    status = block.get("status")
    if status not in COMPREHENSION_STATUSES:
        problems.append(
            f"blend.status: illegal value {status!r} "
            f"(legal: {sorted(COMPREHENSION_STATUSES)})"
        )
    sha = block.get("source_shell_sha256")
    if sha is not None and not isinstance(sha, str):
        problems.append(
            f"blend.source_shell_sha256: must be a hex string or null, got {sha!r}"
        )

    recorded = {
        e.get("sha256")
        for e in (
            ((profile.get("provenance") or {}).get("blended_shells") or [])
            if isinstance(profile.get("provenance"), dict)
            else []
        )
        if isinstance(e, dict)
    }

    ledger = block.get("ledger")
    if not isinstance(ledger, dict):
        problems.append("blend.ledger: required object")
        return problems
    extra = sorted(set(ledger) - {"filled", "corroborated"})
    if extra:
        problems.append(
            f"blend.ledger: only 'filled'/'corroborated' are legal keys, found {extra}"
        )

    filled = ledger.get("filled")
    if filled is not None:
        if not isinstance(filled, dict):
            problems.append("blend.ledger.filled: must be an object")
        else:
            for fact_path, mark in filled.items():
                fpath = f"blend.ledger.filled.{fact_path}"
                if not isinstance(fact_path, str) or not fact_path:
                    problems.append(f"{fpath}: fact path must be a non-empty string")
                if not isinstance(mark, dict) or set(mark) != {"from"}:
                    problems.append(f"{fpath}: must be an object with only 'from'")
                    continue
                src = mark.get("from")
                if not isinstance(src, str) or not _SHA256_HEX_RE.match(src):
                    problems.append(f"{fpath}.from: must be a 64-char hex sha256")
                elif src not in recorded:
                    problems.append(
                        f"{fpath}.from: sha {src!r} is not a recorded blended shell"
                    )

    corroborated = ledger.get("corroborated")
    if corroborated is not None:
        if not isinstance(corroborated, dict):
            problems.append("blend.ledger.corroborated: must be an object")
        else:
            for fact_path, mark in corroborated.items():
                fpath = f"blend.ledger.corroborated.{fact_path}"
                if not isinstance(fact_path, str) or not fact_path:
                    problems.append(f"{fpath}: fact path must be a non-empty string")
                if not isinstance(mark, dict) or set(mark) != {"by"}:
                    problems.append(f"{fpath}: must be an object with only 'by'")
                    continue
                by = mark.get("by")
                if not isinstance(by, list) or not by:
                    problems.append(f"{fpath}.by: must be a non-empty list")
                    continue
                # Item types FIRST: sorted()/set() on a mixed or unhashable
                # list raise, and validate() must RETURN problems, never throw.
                typed = all(isinstance(src, str) for src in by)
                for src in by:
                    if not isinstance(src, str) or not _SHA256_HEX_RE.match(src):
                        problems.append(
                            f"{fpath}.by: every entry must be a 64-char hex sha256"
                        )
                    elif src not in recorded:
                        problems.append(
                            f"{fpath}.by: sha {src!r} is not a recorded blended shell"
                        )
                if typed and (by != sorted(by) or len(by) != len(set(by))):
                    problems.append(f"{fpath}.by: must be sorted and unique")
    return problems


# The advisory free-text fields a palette annotation may carry (the model NAMES a
# captured color; it never authors a real color, so ``ref``/``hex`` are NOT here).
PALETTE_ANNOTATION_FIELDS: tuple[str, ...] = (
    "name",
    "purpose",
    "use_when",
    "semantic_role",
)


def _validate_palette_annotations(annotations: Any) -> list[str]:
    """Validate ``comprehension.palette_annotations`` (absent is fine).

    SHAPE-ONLY: a map ``{ <palette_key>: { name?, purpose?, use_when?,
    semantic_role?, alias? } }`` whose every value is an object of advisory strings
    (or null). The KEY is a template-derived palette id (``accent1`` / ``hex:RRGGBB``)
    and is NOT syntax-checked here (a ``hex:...`` key is legal and would fail a
    role-id regex); its MEMBERSHIP against the surfaced palette inventory is the
    fail-closed QA check. NEVER required.

    The optional ``alias`` (Cluster E1) is a DIRECTIVE the model writes to mint a
    syntactically-legal dotted token aliasing this captured entry (an off-theme
    ``hex:RRGGBB`` accent becomes addressable as a clean run-color token). It is a
    string here (shape only); its DOTTED SYNTAX and non-collision are validated
    fail-closed at merge time (``check_membership``) and QA time
    (``check_palette_alias_targets``), and the engine - never the model - copies the
    captured ``ref`` byte-identical into the minted token. It is NOT in
    :data:`PALETTE_ANNOTATION_FIELDS`, so it is never mirrored back onto the source
    entry as advisory text.
    """
    if annotations is None:
        return []
    if not isinstance(annotations, dict):
        return ["comprehension.palette_annotations: must be an object"]
    problems: list[str] = []
    for key, ann in annotations.items():
        path = f"comprehension.palette_annotations.{key}"
        if not isinstance(key, str) or not key:
            problems.append(f"{path}: palette key must be a non-empty string")
        if ann is None:
            continue
        if not isinstance(ann, dict):
            problems.append(f"{path}: must be an object")
            continue
        for field in (*PALETTE_ANNOTATION_FIELDS, "alias"):
            val = ann.get(field)
            if val is not None and not isinstance(val, str):
                problems.append(
                    f"{path}.{field}: must be a string or null, got {val!r}"
                )
    return problems


def _validate_comp_fragments(fragments: Any) -> list[str]:
    """Validate ``comprehension.fragments`` (reusable-fragment proposals).

    SHAPE-ONLY, consistent with the rest of ``_validate_comprehension``: each
    entry must carry a non-empty string ``ref``, a closed-enum ``kind``
    (``component`` | ``section``), and a non-empty ``blocks`` list. ``purpose`` is
    an optional advisory string. The block CONTENTS are NOT parsed here - that is
    the fail-closed ``check_fragments`` membership check (it needs
    ``ir.model.block_from_dict`` and must reject, not just shape-flag). Absent /
    empty is fine (the default).
    """
    if fragments is None:
        return []
    if not isinstance(fragments, list):
        return ["comprehension.fragments: must be a list"]
    problems: list[str] = []
    for i, frag in enumerate(fragments):
        path = f"comprehension.fragments[{i}]"
        if not isinstance(frag, dict):
            problems.append(f"{path}: must be an object")
            continue
        ref = frag.get("ref")
        if not isinstance(ref, str) or not ref:
            problems.append(f"{path}.ref: required non-empty string")
        kind = frag.get("kind")
        if kind not in FRAGMENT_KINDS:
            problems.append(
                f"{path}.kind: illegal value {kind!r} (legal: {sorted(FRAGMENT_KINDS)})"
            )
        blocks = frag.get("blocks")
        if not isinstance(blocks, list) or not blocks:
            problems.append(f"{path}.blocks: required non-empty list of block dicts")
        purpose = frag.get("purpose")
        if purpose is not None and not isinstance(purpose, str):
            problems.append(
                f"{path}.purpose: must be a string or null, got {purpose!r}"
            )
    return problems


def _validate_comp_audit(audit: Any) -> list[str]:
    """Validate ``comprehension.audit`` (the persisted L2 verdict map, Cluster C1).

    SHAPE-ONLY, mirroring :func:`_validate_comp_fragments` and never-required: a
    profile without the key, or with ``audit == {}``, yields no problems. Each value
    must be an object carrying a closed ``verdict`` (``PASS|FAIL|NA``, illegal value
    rejected like ``fill_rule``/``verdict``); ``evidence`` is advisory (a string or
    null), accepted as-is and NEVER gated; ``shell_sha256`` / ``content_sha256`` are
    optional hex strings or null (they scope the short-circuit, gated by value at
    generate, not by this structural validator).

    It deliberately does NOT check that the KEY is a real ``visual_manifest.json``
    checklist id - that is the fail-closed ``check_audit_targets`` /
    ``check_membership`` audit-arm job (it needs the profile-derived checklist and
    must be able to ERROR on an empty checklist, which a never-required structural
    validator must not do; same split as comprehension refs vs membership).
    """
    if audit is None:
        return []
    if not isinstance(audit, dict):
        return [
            "comprehension.audit: must be an object mapping checklist id -> verdict"
        ]
    problems: list[str] = []
    for key, row in audit.items():
        path = f"comprehension.audit.{key}"
        if not isinstance(key, str) or not key:
            problems.append(f"{path}: audit key must be a non-empty string")
        if not isinstance(row, dict):
            problems.append(f"{path}: must be an object")
            continue
        verdict = row.get("verdict")
        if verdict not in AUDIT_VERDICTS:
            problems.append(
                f"{path}.verdict: illegal value {verdict!r} (legal: {sorted(AUDIT_VERDICTS)})"
            )
        evidence = row.get("evidence")
        if evidence is not None and not isinstance(evidence, str):
            problems.append(
                f"{path}.evidence: must be a string or null, got {evidence!r}"
            )
        for sha_field in ("shell_sha256", "content_sha256"):
            val = row.get(sha_field)
            if val is not None and not isinstance(val, str):
                problems.append(
                    f"{path}.{sha_field}: must be a hex string or null, got {val!r}"
                )
    return problems


def _validate_comp_triage(triage: Any) -> list[str]:
    """Validate ``comprehension.triage`` (the model-assisted QA-triage list, C2).

    SHAPE-ONLY, mirroring :func:`_validate_comp_audit` and never-required: a profile
    without the key, or with ``triage == []``, yields no problems. Each entry must be
    an object carrying a ``check`` that is one of the closed
    :data:`AMBIGUOUS_TRIAGE_CHECKS` ids (illegal value rejected like ``fill_rule`` /
    ``verdict``), a ``disposition`` in the closed :class:`TriageDisposition` enum
    (``expected`` | ``defect``), an optional ``location`` (a string or null), and an
    optional advisory ``evidence`` (a string or null, accepted as-is and NEVER gated).

    It deliberately does NOT check that the ``(check, location)`` pair was actually
    EMITTED by the gate, nor that the pair is unique across the proposal - those are
    the fail-closed ``check_triage`` / ``check_triage_targets`` job (same split as
    audit refs vs ``check_audit_targets``). Restricting ``check`` to the WARNING-only
    eligible set HERE is the belt that makes an ERROR-aimed triage entry rejected at
    merge before ``_apply_triage`` ever runs - it can never demote an ERROR.
    """
    if triage is None:
        return []
    if not isinstance(triage, list):
        return ["comprehension.triage: must be a list"]
    problems: list[str] = []
    for i, entry in enumerate(triage):
        path = f"comprehension.triage[{i}]"
        if not isinstance(entry, dict):
            problems.append(f"{path}: must be an object")
            continue
        check = entry.get("check")
        if check not in AMBIGUOUS_TRIAGE_CHECKS:
            problems.append(
                f"{path}.check: illegal value {check!r} "
                f"(legal: {sorted(AMBIGUOUS_TRIAGE_CHECKS)})"
            )
        disposition = entry.get("disposition")
        if disposition not in TRIAGE_DISPOSITIONS:
            problems.append(
                f"{path}.disposition: illegal value {disposition!r} "
                f"(legal: {sorted(TRIAGE_DISPOSITIONS)})"
            )
        location = entry.get("location")
        if location is not None and not isinstance(location, str):
            problems.append(
                f"{path}.location: must be a string or null, got {location!r}"
            )
        evidence = entry.get("evidence")
        if evidence is not None and not isinstance(evidence, str):
            problems.append(
                f"{path}.evidence: must be a string or null, got {evidence!r}"
            )
    return problems


def _validate_comp_promote_appearance(promote: Any) -> list[str]:
    """Validate ``comprehension.promote_appearance`` (the faked-heading promotion list,
    Cluster E2).

    SHAPE-ONLY, mirroring :func:`_validate_comp_triage` and never-required: a profile
    without the key, or with ``promote_appearance == []``, yields no problems. Each
    entry must be an object carrying a non-empty string ``pseudo_heading_ref`` (a ref
    the detector surfaced) and a non-empty string ``target_role_id`` (a declared
    heading role). The model authors NO size/color here - only NAMES a surfaced ref +
    a target role - so an entry carrying a ``size_hp`` / ``color`` is rejected (the
    engine is the sole author of the promoted value, copied from the captured fact).

    It deliberately does NOT check that the ``pseudo_heading_ref`` was actually
    SURFACED by the detector, nor that ``target_role_id`` is a declared heading role,
    nor that the ``(ref, target)`` pair is unique - those are the fail-closed
    ``check_promote_appearance`` job (same split as triage refs vs ``check_triage``).
    """
    if promote is None:
        return []
    if not isinstance(promote, list):
        return ["comprehension.promote_appearance: must be a list"]
    problems: list[str] = []
    for i, entry in enumerate(promote):
        path = f"comprehension.promote_appearance[{i}]"
        if not isinstance(entry, dict):
            problems.append(f"{path}: must be an object")
            continue
        ref = entry.get("pseudo_heading_ref")
        if not isinstance(ref, str) or not ref:
            problems.append(f"{path}.pseudo_heading_ref: required non-empty string")
        target = entry.get("target_role_id")
        if not isinstance(target, str) or not target:
            problems.append(f"{path}.target_role_id: required non-empty string")
        # The model NAMES only; it never authors the promoted appearance. A size/color
        # in the entry would let the model inject a value the template may not carry,
        # so reject it here (the engine copies the captured fact, the sole author).
        for authored in ("size_hp", "color"):
            if entry.get(authored) is not None:
                problems.append(
                    f"{path}.{authored}: must not be authored by the model "
                    "(the engine copies the captured pseudo_heading value)"
                )
    return problems


def _validate_comp_indexes(indexes: Any) -> list[str]:
    """Validate ``comprehension.conventions.indexes`` (list of derived-index descriptors)."""
    if indexes is None:
        return []
    if not isinstance(indexes, list):
        return ["comprehension.conventions.indexes: must be a list"]
    problems: list[str] = []
    for i, idx in enumerate(indexes):
        path = f"comprehension.conventions.indexes[{i}]"
        if not isinstance(idx, dict):
            problems.append(f"{path}: must be an object")
            continue
        ref = idx.get("index_ref")
        if not isinstance(ref, str) or not ref:
            problems.append(f"{path}.index_ref: required non-empty string")
        rec = idx.get("reconcile")
        if rec is not None and rec not in RECONCILE_RULES:
            problems.append(
                f"{path}.reconcile: illegal value {rec!r} (legal: {sorted(RECONCILE_RULES)})"
            )
        feeds = idx.get("feeds_from_role_id")
        if feeds is not None and (
            not isinstance(feeds, str) or not is_valid_role_id(feeds)
        ):
            problems.append(
                f"{path}.feeds_from_role_id: must be a role id or null, got {feeds!r}"
            )
        seq = idx.get("seq_id")
        if seq is not None and not isinstance(seq, str):
            problems.append(f"{path}.seq_id: must be a string or null, got {seq!r}")
        target = idx.get("caption_target")
        if target is not None and target not in CAPTION_TARGETS:
            problems.append(
                f"{path}.caption_target: illegal value {target!r} "
                f"(legal: {sorted(CAPTION_TARGETS)} or null)"
            )
    return problems


def _validate_comp_sections(sections: Any) -> list[str]:
    """Validate ``comprehension.conventions.sections`` (list of region descriptors)."""
    if sections is None:
        return []
    if not isinstance(sections, list):
        return ["comprehension.conventions.sections: must be a list"]
    problems: list[str] = []
    for i, sec in enumerate(sections):
        path = f"comprehension.conventions.sections[{i}]"
        if not isinstance(sec, dict):
            problems.append(f"{path}: must be an object")
            continue
        ref = sec.get("region_ref")
        if not isinstance(ref, str) or not ref:
            problems.append(f"{path}.region_ref: required non-empty string")
        for attr in ("required", "repeatable"):
            val = sec.get(attr)
            if val is not None and not isinstance(val, bool):
                problems.append(f"{path}.{attr}: must be a bool or null, got {val!r}")
    return problems


def _validate_fragment_registry(registry: Any, key: str) -> list[str]:
    """Validate a reusable-fragment registry (``components`` / ``sections``).

    Well-formedness only: the registry must be a map whose every entry is a dict
    carrying a ``blocks`` list (the primitive template ``expand_components`` inlines
    when an idoc references the entry by id). The block CONTENTS are not re-checked
    here - ``block_from_dict`` enforces those at expansion time; over-constraining
    them would duplicate that contract. Absent / empty is fine (the default).
    ``slots`` PARAMETERIZATION and auto-POPULATION of these registries remain
    deferred, so nothing beyond ``blocks`` is required.
    """
    if registry is None:
        return []
    if not isinstance(registry, dict):
        return [f"{key}: must be an object mapping ref -> fragment definition"]
    problems: list[str] = []
    for ref, definition in registry.items():
        path = f"{key}.{ref}"
        if not isinstance(definition, dict):
            problems.append(f"{path}: must be an object")
            continue
        blocks = definition.get("blocks")
        if not isinstance(blocks, list):
            problems.append(f"{path}.blocks: required list of block definitions")
    return problems


def supports_role(profile: dict, rid: str) -> bool:
    """Return True if ``profile.roles`` contains a concrete entry for ``rid``."""
    roles = profile.get("roles") or {}
    return rid in roles and rid != "_index"


def list_role_ids(profile: dict) -> list[str]:
    """Return the concrete role ids in a profile (``_index`` order if present)."""
    roles = profile.get("roles") or {}
    index = roles.get("_index")
    if isinstance(index, list) and index:
        return [r for r in index if r in roles]
    return [r for r in roles if r != "_index"]


def overrides_block(profile: dict) -> dict:
    """Return the ``rules.overrides`` block (``{}`` if absent / malformed).

    A read helper peer of :func:`list_role_ids`: the resolver, store, and QA
    consumers all read overrides through this single accessor so a missing
    ``rules`` / ``overrides`` key (an old profile) degrades to the empty block
    rather than raising. This is shape-tolerant only; presence/freeze is decided by
    :func:`store.overrides_are_present`.
    """
    rules = profile.get("rules")
    if not isinstance(rules, dict):
        return {}
    overrides = rules.get("overrides")
    return overrides if isinstance(overrides, dict) else {}


def list_overrides(profile: dict) -> list[tuple[str, str, str]]:
    """Return the learned override entries as ``(kind, requested, target)`` triples.

    Flattens the three closed-kind containers of ``rules.overrides`` into one list,
    using the :class:`OverrideKind` value as the ``kind`` discriminator
    (``register_demo_clear`` entries carry the cleared string as both the
    ``requested`` and ``target`` slot - they map text, not a role pair). A read
    helper only: it does NOT consult the freeze sha, so callers that need the LIVE
    overrides must first gate on :func:`store.overrides_are_present`.

    PUBLIC INSPECTION API with no internal engine caller yet: it exists so an
    operator/agent surface (e.g. the ``list`` verb, or the B4 model bundle) can show
    a profile's learned lessons without re-implementing the flattening. Do not
    mistake it for dead code wired into the resolve path - the resolver reads the
    raw containers directly.
    """
    overrides = overrides_block(profile)
    entries: list[tuple[str, str, str]] = []
    reroutes = overrides.get("reroute_roles")
    if isinstance(reroutes, dict):
        for requested, target in reroutes.items():
            entries.append((OverrideKind.REROUTE_ROLE.value, requested, target))
    swaps = overrides.get("number_format_swaps")
    if isinstance(swaps, dict):
        for rid, mask in swaps.items():
            entries.append((OverrideKind.NUMBER_FORMAT.value, rid, mask))
    clears = overrides.get("demo_clears")
    if isinstance(clears, list):
        for val in clears:
            entries.append((OverrideKind.REGISTER_DEMO_CLEAR.value, val, val))
    return entries
