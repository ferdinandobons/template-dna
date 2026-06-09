# SPDX-License-Identifier: MIT
"""Format-neutral brand APPEARANCE apply orchestration (font / size / color).

This is the shared control flow the per-format generators (docx and pptx today;
xlsx in a later PR) delegate to so the "read the three axes off the resolver op,
then brand each run only when its axis is unset" logic has exactly ONE writer
across kinds.

It is lxml / python-docx / pptx / openpyxl-FREE at import (like
:mod:`brandkit.common.text`): the per-axis run mutations and the set-only-when-unset
probes live behind a small BACKEND object the format adapter supplies (e.g. docx's
``DOCX_BACKEND`` wrapping ``run.font.name``/``.size``/``.color``; pptx's
``PPTX_BACKEND``). This module only:

  1. reads the captured brand axes off the resolver op (:func:`op_latin` /
     :func:`op_size_hp` / :func:`op_color`) - STRICTLY from ``op.appearance``, never
     a literal in the engine, so off-brand output stays impossible by construction;
  2. resolves a run's ``color`` palette TOKEN to its captured ref
     (:func:`resolve_run_color`), recording a graceful INFO finding for an unknown
     token (the writer never fabricates a color);
  3. drives the backend to apply those axes (:func:`apply_role_appearance` over a
     paragraph's runs; :func:`apply_run_color` for a single run), gating each write
     on the backend's ``*_unset`` probe so an inherited-but-correct value is never
     clobbered and re-runs stay byte-identical.

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
    """Apply captured brand typography (font, size, color) from the resolved op as
    direct run formatting on ``target``'s runs (hyperlink runs included for docx).

    The three axes are INDEPENDENT: each is applied only when the run's corresponding
    ``*_unset`` probe is true, so a role carrying a size but no font (or a color but
    no font) still applies the axes it has. A target that exposes no runs (a docx
    table here) yields nothing and is skipped. An empty appearance (a pre-capture
    profile) returns before touching any run, so output stays byte-identical to
    today."""
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
