# SPDX-License-Identifier: MIT
"""Format-neutral brand APPEARANCE apply orchestration (font / size / color /
geometry / table / numbering).

This is the shared control flow the per-format generators (docx and pptx today;
xlsx in a later PR) delegate to so the "read the appearance axes off the resolver
op, then brand each run/paragraph only when its axis is unset" logic has exactly
ONE writer across kinds.

It is lxml / python-docx / pptx / openpyxl-FREE at import (like
:mod:`brandkit.common.text`): the per-axis run/paragraph mutations and the
set-only-when-unset probes live behind a small BACKEND object the format adapter
supplies (e.g. docx's ``DOCX_BACKEND`` wrapping ``run.font.name``/``.size``/
``.color`` and the paragraph's ``w:pPr`` geometry; pptx's ``PPTX_BACKEND``). This
module only:

  1. reads the captured brand axes off the resolver op (:func:`op_latin` /
     :func:`op_size_hp` / :func:`op_color` / :func:`op_geometry` / :func:`op_table` /
     :func:`op_numbering`) - STRICTLY from ``op.appearance``, never a literal in the
     engine, so off-brand output stays impossible by construction. The axes ride
     different pathways: font/size/color are run axes applied here through the
     backend; geometry is a paragraph axis applied here via the backend's
     ``set_geometry`` hook (docx-only today); table and numbering are docx-only and
     realized by dedicated writers OUTSIDE this orchestration, but declared in
     :data:`APPEARANCE_AXES` so the parity ledger measures them;
  2. resolves a run's ``color`` palette TOKEN to its captured ref
     (:func:`resolve_run_color`), recording a graceful INFO finding for an unknown
     token (the writer never fabricates a color);
  3. drives the backend to apply the run/paragraph axes
     (:func:`apply_role_appearance` over a paragraph's runs and geometry;
     :func:`apply_run_color` for a single run), gating each write on the backend's
     ``*_unset`` probe so an inherited-but-correct value is never clobbered and
     re-runs stay byte-identical;
  4. keeps the parity ledger (Cluster E3): :func:`_record_degraded_axes` emits one
     INFO ``appearance_apply_degraded`` finding per captured axis the format backend
     does not declare it realizes, so an unmaterialized axis surfaces gracefully
     instead of silently dropping.

The brand guarantee is preserved end to end: every applied value comes only from
``op.appearance`` / the resolved palette ref, the set-only-when-unset guard is
enforced HERE via the backend's ``*_unset`` probe (so the guard reads identically
across kinds), and an empty appearance (a pre-capture profile) is a byte-identical
no-op.
"""

from __future__ import annotations

from typing import Iterable, Optional, Protocol, runtime_checkable

from brandkit.profile import schema
from brandkit.qa.model import Finding


@runtime_checkable
class AppearanceBackend(Protocol):
    """The per-format hook set the shared orchestration drives.

    The format adapter supplies the concrete run type, the per-axis ``*_unset``
    probes (which decide set-only-when-unset - e.g. docx's ``run.font.<axis> is
    None``), and the per-axis ``set_*`` writes. The shared orchestration owns the
    control flow: it reads the captured axis off the op, asks the backend whether the
    run carries that axis already, and only then writes. This module never touches a
    run attribute directly, so the same guard governs every kind.
    """

    def runs_of(self, target) -> Iterable:
        """Every run of ``target`` to brand (hyperlink runs included for docx)."""
        ...

    def font_unset(self, run) -> bool:
        """``True`` when ``run`` carries no explicit font (so it may be branded)."""
        ...

    def set_font(self, run, latin: str) -> None:
        """Set ``run``'s font to ``latin`` (called only when :meth:`font_unset`)."""
        ...

    def size_unset(self, run) -> bool:
        """``True`` when ``run`` carries no explicit size (so it may be branded)."""
        ...

    def set_size(self, run, half_pts: int) -> None:
        """Set ``run``'s size from ``half_pts`` (backend converts to its own unit;
        called only when :meth:`size_unset`)."""
        ...

    def color_unset(self, run) -> bool:
        """``True`` when ``run`` carries no explicit color (so it may be branded)."""
        ...

    def set_color(self, run, ref: dict, findings: list[Finding]) -> None:
        """Set ``run``'s color from the resolved palette ``ref``
        (``{'kind': 'hex'|'theme', 'hex'?, 'theme'?}``); called only when
        :meth:`color_unset`. A token the backend cannot realize is left inherited
        with an INFO finding rather than raising."""
        ...

    def paragraphs_of(self, target) -> Iterable:
        """Every PARAGRAPH of ``target`` that may carry brand geometry (Cluster D1).

        Geometry (``w:pPr``) lives on the PARAGRAPH, not the run, so this is a separate
        iterator from :meth:`runs_of`. A backend with no paragraph-geometry support
        (pptx/xlsx today) yields nothing, so :func:`apply_role_appearance` no-ops the
        geometry pass for it (the axis is docx-only by construction)."""
        ...

    def set_geometry(self, para, geometry: dict) -> None:
        """Apply the captured ``geometry`` dict to ``para``'s ``w:pPr``, SET-ONLY-WHEN-
        UNSET per property (Cluster D1, docx-only).

        Each property (spacing before/after/line, the four indents, the four border
        sides, shading fill) is written ONLY when the paragraph does not already carry
        it directly, so an authored/inherited value is never clobbered and re-runs stay
        byte-identical. A backend without geometry support never has this called
        (:meth:`paragraphs_of` yields nothing)."""
        ...


def op_latin(op) -> Optional[str]:
    """The captured brand latin font this resolved op applies, or ``None``.

    The brand value comes ONLY from ``op.appearance`` (which the resolver populated
    from the profile, role-specific font winning over the document body font), never
    from a literal in this engine, so off-brand output stays impossible by
    construction. ``None`` for every pre-capture profile (empty appearance)."""
    return (getattr(op, "appearance", None) or {}).get("font", {}).get("latin")


def op_size_hp(op) -> Optional[int]:
    """The captured brand run SIZE (half-points) this resolved op applies, or ``None``.

    Read STRICTLY from ``op.appearance`` (resolver-populated from the profile,
    role-specific size winning over the document body size). The body size only ever
    reaches the body/paragraph family - the resolver's family gate keeps it off
    headings - so a heading's intrinsic style size is never overridden here."""
    return (getattr(op, "appearance", None) or {}).get("size_hp")


def op_color(op) -> Optional[dict]:
    """The captured brand run COLOR this resolved op applies (a ``{'kind': ...}``
    object), or ``None``.

    Read STRICTLY from ``op.appearance``. Like size, the body color reaches only the
    body/paragraph family via the resolver's family gate, so a heading's intrinsic
    style color is never overridden here."""
    return (getattr(op, "appearance", None) or {}).get("color")


def op_geometry(op) -> Optional[dict]:
    """The captured brand paragraph GEOMETRY this resolved op applies (Cluster D1,
    docx-only), or ``None``.

    Read STRICTLY from ``op.appearance.geometry`` (resolver-populated from the profile,
    role-specific geometry winning over the body geometry; NO family gate). ``None``
    for every profile that carries no captured geometry, so the no-geometry path is a
    byte-identical no-op."""
    return (getattr(op, "appearance", None) or {}).get("geometry")


def op_table(op) -> Optional[dict]:
    """The captured brand TABLE conditional-format facts this resolved op applies
    (Cluster D2, docx-only), or ``None``.

    Read STRICTLY from ``op.appearance.table`` (resolver-populated from the profile,
    role-specific table appearance winning over the body table default; NO family gate).
    The dict carries the ``tblLook`` bitmask, the referenced table ``style_id``, and the
    ``cell_margins`` twips - facts the docx table writer re-emits set-only-when-unset
    (the band fills/borders stay in the shell's style part). ``None`` for every profile
    that carries no captured table appearance, so the no-table path is a byte-identical
    no-op. The docx table writer (``_apply_table_style``) is the ONLY consumer; there is
    deliberately NO backend protocol method, so the shared run/geometry orchestration is
    untouched and pptx/xlsx never see this axis."""
    return (getattr(op, "appearance", None) or {}).get("table")


def op_numbering(op) -> Optional[dict]:
    """The captured brand list NUMBERING facts this resolved op applies (Cluster D3,
    docx-only), or ``None``.

    Read STRICTLY from ``op.appearance.numbering`` (resolver-populated from the profile,
    role-specific numbering winning over the body numbering default; NO family gate). The
    dict carries the referenced ``num_id`` and ``abstract_num_id`` (SYMBOLIC refs into the
    shell's numbering part) and the shell's OWN ``per_level_facts`` (``{ilvl -> {numFmt,
    lvlText, indent}}``) - facts the docx list writer re-asserts onto the cloned
    ``w:abstractNum`` set-only-when-unset (the definition itself stays the shell's). The
    docx list writer (``_apply_list_numbering_appearance``) is the ONLY consumer; there is
    deliberately NO backend protocol method, so the shared run/geometry orchestration is
    untouched and pptx/xlsx never see this axis. ``None`` for every profile that carries no
    captured numbering, so the no-numbering path is a byte-identical no-op."""
    return (getattr(op, "appearance", None) or {}).get("numbering")


# The closed set of appearance AXES a resolved op can carry (the op.appearance
# keys the resolver merges). The parity ledger (Cluster E3) compares the axes an op
# CARRIES against the axes the format backend declares it REALIZES, so a captured
# brand axis that a format cannot yet honor is surfaced instead of silently dropped.
APPEARANCE_AXES: tuple[str, ...] = (
    "font",
    "size_hp",
    "color",
    "geometry",
    "table",
    "numbering",
)

# The axes every backend realizes today without declaring anything: the original
# run-typography trio. Geometry is inferred from the ``paragraphs_of`` hook for a
# backend that predates the explicit declaration.
_BASE_REALIZED_AXES: frozenset[str] = frozenset({"font", "size_hp", "color"})


def _realized_axes(backend) -> frozenset[str]:
    """The appearance axes ``backend`` can realize (Cluster E3).

    A backend DECLARES its capability via a ``realized_axes`` attribute (the docx
    backend declares all six: the table/numbering axes are realized by the docx
    writers outside this orchestration, but the declaration lives here so the parity
    ledger has ONE source of truth). A backend without the attribute falls back to
    the run-typography trio plus geometry-if-it-has-the-hook, so a third-party /
    pre-E3 backend keeps its exact prior behavior."""
    declared = getattr(backend, "realized_axes", None)
    if declared is not None:
        return frozenset(declared)
    base = set(_BASE_REALIZED_AXES)
    if hasattr(backend, "paragraphs_of"):
        base.add("geometry")
    return frozenset(base)


def _record_degraded_axes(backend, op, findings: list[Finding]) -> None:
    """Emit one INFO ``appearance_apply_degraded`` finding per captured-but-
    unrealizable axis (Cluster E3): the uniform parity ledger.

    Fires on the STRUCTURAL fact that ``op.appearance`` carries an axis the backend
    does not declare it realizes - the captured brand intent degrades gracefully
    instead of silently. The finding names ONLY the role id and the axis (never a
    brand value), keyed ``location='<role_id>:<axis>'`` so cross-run recurrence
    (B2/B4 ``generation_history``) and the L2 model can measure parity gaps.
    INFO-only: NOT in ``DEFAULT_L0_INVARIANTS`` / ``LEARNABLE_CHECKS``, so it can
    never flip a verdict or feed a deterministic lesson. Deduplicated per
    ``(role, axis)`` within a run (the apply is called per paragraph/cell). A profile
    whose axes are all realized (every existing real profile) emits NOTHING, so the
    existing paths stay byte-identical."""
    appearance = getattr(op, "appearance", None) or {}
    realized = _realized_axes(backend)
    role_id = getattr(op, "role_id", None) or "?"
    for axis in APPEARANCE_AXES:
        if axis in realized or not appearance.get(axis):
            continue
        location = f"{role_id}:{axis}"
        if any(
            f.check == "appearance_apply_degraded" and f.location == location
            for f in findings
        ):
            continue  # one ledger entry per (role, axis) per run
        findings.append(
            Finding(
                "appearance_apply_degraded",
                schema.Severity.INFO.value,
                f"captured appearance axis {axis!r} of role {role_id!r} is not "
                "realizable by this format's writer; output degrades gracefully",
                location=location,
            )
        )


def resolve_run_color(
    resolver,
    token: Optional[str],
    findings: Optional[list[Finding]],
) -> Optional[dict]:
    """Resolve a run's ``color`` palette TOKEN to its captured color ``ref`` object.

    Returns the ``{'kind': 'theme'|'hex', ...}`` object the writers apply, or
    ``None`` when there is no token, no resolver, or the token is unknown. An
    UNRESOLVED token (present, but not a key of ``theme.palette``) is recorded as a
    graceful INFO ``color_token_unresolved`` finding (the run is then left to inherit
    its color, mirroring ``appearance_color_skipped``) - the writer NEVER fabricates
    a color for an unknown token. The token itself never carries a literal color
    (``normalize_runs`` already rejected any hex-shaped value structurally)."""
    if not token or resolver is None:
        return None
    color = resolver.resolve_color(token)
    if color is None and findings is not None:
        findings.append(
            Finding(
                "color_token_unresolved",
                schema.Severity.INFO.value,
                f"run color token {token!r} is not a key of theme.palette; "
                "left inherited",
            )
        )
    return color


def apply_role_appearance(
    backend: AppearanceBackend, target, op, findings: list[Finding]
) -> None:
    """Apply captured brand typography (font, size, color) and geometry from the
    resolved op as direct run/paragraph formatting on ``target`` (hyperlink runs
    included for docx).

    The run axes are INDEPENDENT: each is applied only when the run's corresponding
    ``*_unset`` probe is true, so a role carrying a size but no font (or a color but
    no font) still applies the axes it has. Geometry is a separate PARAGRAPH-level
    axis, applied per paragraph via the backend's ``set_geometry`` hook (its
    set-only-when-unset guard lives per PROPERTY inside the backend). Table and
    numbering are MEASURED here by the parity ledger but realized by dedicated
    format-specific writers elsewhere. A target that exposes no runs (a docx table
    here) yields nothing and is skipped. An empty appearance (a pre-capture profile)
    returns before touching any run, so output stays byte-identical to today."""
    # Parity ledger (Cluster E3): surface any captured axis this backend cannot
    # realize BEFORE the early return, so a table/numbering-only op on a format
    # without those writers is still measured. Appends findings only; it never
    # touches the document, so byte-identity is unaffected.
    _record_degraded_axes(backend, op, findings)
    latin = op_latin(op)
    size_hp = op_size_hp(op)
    color = op_color(op)
    geometry = op_geometry(op)
    if not (latin or size_hp or color or geometry):
        return
    if latin or size_hp or color:
        for run in backend.runs_of(target):
            if latin and backend.font_unset(run):
                backend.set_font(run, latin)
            if size_hp and backend.size_unset(run):
                backend.set_size(run, size_hp)
            if color and backend.color_unset(run):
                backend.set_color(run, color, findings)
    # GEOMETRY (Cluster D1, docx-only) is a PARAGRAPH axis, applied separately from the
    # run axes. A backend without geometry support (pptx/xlsx) exposes no
    # ``paragraphs_of`` and is skipped, so the axis is docx-only by construction. Each
    # property is written set-only-when-unset by the backend's ``set_geometry``.
    if geometry and hasattr(backend, "paragraphs_of"):
        for para in backend.paragraphs_of(target):
            backend.set_geometry(para, geometry)


def apply_run_color(
    backend: AppearanceBackend, run, color: Optional[dict], findings: list[Finding]
) -> None:
    """Apply a single run's resolved palette ``color`` ref via the backend.

    Gated on the backend's ``color_unset`` probe, this gives an explicit per-run
    token first-writer-wins precedence over the later :func:`apply_role_appearance`
    body/role default and keeps re-runs byte-identical. A falsy ``color`` is a no-op.
    """
    if color and backend.color_unset(run):
        backend.set_color(run, color, findings)
