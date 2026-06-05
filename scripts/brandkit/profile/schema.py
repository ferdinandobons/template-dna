# SPDX-License-Identifier: MIT
"""The FROZEN Brand-Profile vocabulary — the single source of truth.

Every other module imports its enum values, key names, and the ``profile.json``
envelope shape from here. The plan (§3) froze the vocabulary to resolve the
cohesion conflicts between dimensions; this module *is* that freeze:

- The discriminator key is ``kind`` (``docx`` | ``pptx`` | ``xlsx``), never
  ``doc_type``. See :class:`Kind`.
- The role registry key is ``roles``, never ``bindings``.
- Resolver types are ``named_style`` | ``placeholder`` | ``cell_style``
  (plus the maturity-staged ``number_format`` | ``named_range`` |
  ``chart_template``). The redundant ``layout_placeholder`` is dropped. See
  :class:`ResolverType`.
- ``schema_version`` is semver; this module pins :data:`SCHEMA_VERSION`.
- The shell always lives at ``template/shell.<ext>``.
- Role status is ``robust`` | ``best_effort`` | ``stub`` (:class:`Status`);
  finding severity is ``INFO`` | ``WARNING`` | ``ERROR`` (:class:`Severity`).

Two helpers complete the contract:
  - :func:`build_envelope` — construct a minimal, valid envelope skeleton.
  - :func:`validate` — return a list of human-readable problems (``[]`` == ok).

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
SCHEMA_VERSION: str = "1.1.0"
SCHEMA_ID: str = "https://template-dna/schema/profile-1.json"


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

    Only ``NAMED_STYLE`` (docx), ``PLACEHOLDER`` (pptx) and ``CELL_STYLE``
    (xlsx) are first-class in M1. ``NUMBER_FORMAT`` / ``NAMED_RANGE`` (xlsx) and
    ``CHART_TEMPLATE`` are staged maturity types carried by the schema from day
    one (status ``stub``/``best_effort``) per §8.
    """

    NAMED_STYLE = "named_style"
    PLACEHOLDER = "placeholder"
    CELL_STYLE = "cell_style"
    NUMBER_FORMAT = "number_format"
    NAMED_RANGE = "named_range"
    CHART_TEMPLATE = "chart_template"


# Which resolver types are legal for which kind (the only dispatch gate).
LEGAL_RESOLVER_TYPES: dict[str, frozenset[str]] = {
    Kind.DOCX.value: frozenset({
        ResolverType.NAMED_STYLE.value,
        ResolverType.CHART_TEMPLATE.value,
    }),
    Kind.PPTX.value: frozenset({
        ResolverType.PLACEHOLDER.value,
        ResolverType.NAMED_STYLE.value,
        ResolverType.CHART_TEMPLATE.value,
    }),
    Kind.XLSX.value: frozenset({
        ResolverType.CELL_STYLE.value,
        ResolverType.NUMBER_FORMAT.value,
        ResolverType.NAMED_RANGE.value,
        ResolverType.CHART_TEMPLATE.value,
    }),
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


class OverflowCapability(str, Enum):
    """Per-format overflow detection mechanism (§6.5). docx never estimates."""

    ESTIMATOR = "estimator"   # pptx
    CELLFIT = "cellfit"       # xlsx
    RENDER = "render"         # docx (LibreOffice render is the only detector)
    NONE = "none"


# Default overflow capability per kind (the frozen §6.5 mapping).
DEFAULT_OVERFLOW_CAPABILITY: dict[str, str] = {
    Kind.DOCX.value: OverflowCapability.RENDER.value,
    Kind.PPTX.value: OverflowCapability.ESTIMATOR.value,
    Kind.XLSX.value: OverflowCapability.CELLFIT.value,
}

# The L0 invariant ids every profile declares it wants enforced (§3.2 qa).
# NOTE: ``lists_use_named_numbering`` was removed (staged) — it was advertised in
# every profile but enforced nowhere (no checker, not in ``registry.CHECKS``), so a
# profile must not claim it. Re-add it here ONLY together with a real checker that
# fails on a direct ``w:numPr`` not backed by a named numbering definition.
DEFAULT_L0_INVARIANTS: tuple[str, ...] = (
    "every_role_resolves",
    "resolver_targets_exist",
    "no_literal_markdown",
    "no_residual_template_text",
)

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
        raise ValueError(f"unknown kind {kind!r}; expected one of {list(KIND_EXTENSION)}")
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
    "schema_version", "kind", "identity", "provenance",
    "theme", "roles", "rules", "anchors", "qa", "verification", "surface",
)


def validate(profile: dict) -> list[str]:
    """Return a list of human-readable problems with ``profile`` (``[]`` == ok).

    This is a *structural* checker — it verifies the frozen vocabulary is
    respected (keys present, enum values legal, resolver types legal for the
    kind, role ids well-formed, the shell path matches the kind, the surface
    block matches the kind). It does NOT judge brand quality (that is the QA
    gate's job) and never raises — it returns the problems so callers can decide
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
    if sv is not None and not re.match(r"^\d+\.\d+\.\d+", str(sv)):
        problems.append(f"schema_version: not semver: {sv!r}")

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
    named_regions = sub.get("named_regions") if isinstance(sub.get("named_regions"), dict) else {}

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
            # cannot prove absence — that is the shell-backed check's job).
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
# Legal vocabulary for a ``structure.skeleton`` region.
STRUCTURE_REGIONS: frozenset[str] = frozenset({"cover", "toc", "body"})


def _validate_usage(path: str, usage: Any) -> list[str]:
    """Validate a per-artifact ``usage`` object (never required; structural only)."""
    problems: list[str] = []
    if not isinstance(usage, dict):
        return [f"{path}: must be an object"]
    scope = usage.get("scope")
    if scope is not None and scope not in USAGE_SCOPES:
        problems.append(f"{path}.scope: illegal value {scope!r} (legal: {sorted(USAGE_SCOPES)})")
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
        if rname not in STRUCTURE_REGIONS:
            problems.append(
                f"structure.skeleton[{i}].region: illegal value {rname!r} "
                f"(legal: {sorted(STRUCTURE_REGIONS)})"
            )
        order = region.get("order")
        if order is not None and not isinstance(order, int):
            problems.append(f"structure.skeleton[{i}].order: must be an int, got {order!r}")
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
