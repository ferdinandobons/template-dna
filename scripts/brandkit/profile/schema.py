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
    """
    parts = [str(family)] + [str(q) for q in qualifiers if q is not None and q != ""]
    rid = ".".join(parts)
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
            "overrides": {},
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
FRAGMENT_KINDS: frozenset[str] = frozenset(e.value for e in FragmentKind)


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

    # fragments: [ { ref, kind, blocks, purpose? } ] - reusable-fragment proposals.
    problems.extend(_validate_comp_fragments(comp.get("fragments")))
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
