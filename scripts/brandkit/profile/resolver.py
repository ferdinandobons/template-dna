# SPDX-License-Identifier: MIT
"""Pure Brand Profile resolver - the single, KIND-AWARE resolution spine.

The resolver maps semantic roles to the concrete resolver op the *shell*
exposes, dispatching on the profile ``kind``:

  - ``docx`` -> ``named_style``   (``{"type", "style_id", "style_name"}``)
  - ``pptx`` -> ``placeholder`` (``{"type", "layout", "ph_idx", "ph_type"}``) or
                 ``named_style`` (both are legal for pptx per the schema)
  - ``xlsx`` -> ``named_range`` | ``cell_style`` | ``number_format``
                 (``{"type", "name"}`` / ``{"type", "style_name"}`` /
                  ``{"type", "number_format"}``)

All three concrete generators (docx/pptx/xlsx) route role resolution through
this one spine so the brand guarantee is enforced in **one** place: the resolver
never invents a target. Every op it emits carries a resolver dict that came
verbatim from ``profile['roles']`` (extracted from the real template), and its
``type`` is checked against :data:`schema.LEGAL_RESOLVER_TYPES` for the profile
kind - a foreign/fabricated resolver type (e.g. a ``placeholder`` smuggled into a
docx profile) is *never* emitted. The model proposes; the validator (here +
``run_qa``) disposes.

It performs no I/O and contains no brand literals; all concrete names and ids
come from the profile.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Optional

from brandkit.common import color as colorutil
from brandkit.ir import model as ir
from brandkit.profile import schema, store


@dataclass(frozen=True)
class ResolvedOp:
    """The outcome of resolving one role against the profile.

    ``resolver`` is the kind-appropriate op the generator applies; its concrete
    shape is selected by ``kind`` (and self-describes via ``resolver['type']``):

      - docx ``named_style``:   ``{"type", "style_id"?, "style_name"?}``
      - pptx ``placeholder``:   ``{"type", "layout", "ph_idx", "ph_type"}``
      - xlsx ``named_range``:   ``{"type", "name"}``
      - xlsx ``cell_style``:    ``{"type", "style_name"}``
      - xlsx ``number_format``: ``{"type", "number_format"}``

    An empty ``resolver`` ({}) means "no concrete target" - a stub/foreign role
    that callers must treat as a no-op (never a fabricated target). ``kind`` is
    the profile kind so a consumer can dispatch on it without re-reading the
    profile; ``resolver_type`` is the ``resolver['type']`` (or ``None`` when the
    op is a stub) for the same convenience.
    """

    role_id: str
    resolver: dict[str, Any]
    status: str
    confidence: float
    kind: Optional[str] = None
    appearance: dict[str, Any] = field(default_factory=dict)

    @property
    def resolver_type(self) -> Optional[str]:
        """The concrete resolver ``type`` for this op, or ``None`` for a stub."""
        return self.resolver.get("type") if self.resolver else None


class ResolveError(LookupError):
    pass


class ProfileResolver:
    """Single resolution spine: maps semantic roles to profile resolver ops.

    KIND-AWARE (plan §4, M-i-8): the resolver records the profile ``kind`` and
    dispatches the kind-appropriate op. The docx vertical drives generation from
    the semantic IR via :meth:`resolve_block`; pptx and xlsx drive generation
    from their own role inventories and route through :meth:`resolve_role`
    directly with kind-appropriate role ids (e.g. ``cover.title`` for pptx, a
    named-region role for xlsx). All three share the *same* brand guarantee,
    enforced here once: :meth:`resolve_role` only ever emits a resolver op whose
    ``type`` is legal for ``kind`` (:data:`schema.LEGAL_RESOLVER_TYPES`) and
    whose target ids the profile itself carries.
    """

    def __init__(self, profile: dict, *, strict: bool = False) -> None:
        self.profile = profile
        self.strict = strict
        self.kind = profile.get("kind")
        self.roles = profile.get("roles") or {}
        # Document-level captured typography (the dominant body font/size/color),
        # applied as a fallback to any paragraph whose role carries no role-specific
        # value on that axis. Empty for every profile that predates typography
        # capture, so behavior is unchanged there. Read once; the brand value lives
        # only in the profile.
        self._default_appearance = self._read_default_appearance(profile)
        # The model-driven color palette: ``theme.palette`` maps a template-derived
        # token (a theme slot like ``accent1``, or a comprehension-named key) to its
        # captured ``ref`` color object. This is the ONLY namespace a run color token
        # resolves off (never ``palette_roles``). Empty for every pre-palette profile,
        # so an unknown / absent token resolves to None (lenient) there.
        self._palette = (
            ((profile.get("theme") or {}).get("palette") or {}) if profile else {}
        )
        # The clrScheme slot -> {"hex": ...} map, used ONLY to resolve a theme token's
        # concrete hex for the w:val fallback (see resolve_color). Read once.
        self._theme_colors = (
            ((profile.get("theme") or {}).get("colors") or {}) if profile else {}
        )
        # The resolver types this profile's kind is allowed to emit. An unknown
        # kind yields an empty set, so every concrete resolver is treated as
        # foreign (fail-closed) rather than blindly trusted.
        self._legal_types: frozenset[str] = (
            schema.LEGAL_RESOLVER_TYPES.get(self.kind, frozenset())
            if self.kind
            else frozenset()
        )
        # Learned, sha-frozen overrides (Cluster B). Built ONLY when the lesson is
        # genuinely present AND bound to the live shell sha (the same presence test
        # the store uses); otherwise the two indexes stay EMPTY so resolve_role
        # takes ZERO new branches and generated bytes are byte-identical to the
        # no-overrides path. A re-extract re-stamps provenance.shell.sha256, which
        # drifts the recorded sha and empties these again (SHELL-FROZEN). Each
        # index re-points ONLY to artifacts the profile already declares; the actual
        # brand value is still authored by the rerouted role's own resolver / the
        # shell's own mask (the resolver remains the single brand chokepoint).
        #
        #   _overrides           : { requested_role_id -> target_role_id }  (reroute)
        #   _number_format_swaps : { role_id -> mask }                      (mask swap)
        #
        # ``register_demo_clear`` is intentionally NOT a resolver concern: the
        # resolver maps roles to ops, never text. Registered demo strings are
        # consumed by the generate/QA residual-clear path (and membership-checked by
        # ``check_override_targets``), so they are deliberately absent here.
        self._overrides: dict[str, str] = {}
        self._number_format_swaps: dict[str, str] = {}
        # Single-hop depth guard: True while inside a reroute's target resolve, so the
        # rerouted target is never itself rerouted (a chain/cycle can't recurse).
        self._in_reroute: bool = False
        if store.overrides_are_present(profile):
            block = schema.overrides_block(profile)
            reroutes = block.get("reroute_roles")
            if isinstance(reroutes, dict):
                self._overrides = {
                    str(k): str(v)
                    for k, v in reroutes.items()
                    if isinstance(k, str) and isinstance(v, str)
                }
            swaps = block.get("number_format_swaps")
            if isinstance(swaps, dict):
                self._number_format_swaps = {
                    str(k): str(v)
                    for k, v in swaps.items()
                    if isinstance(k, str) and isinstance(v, str)
                }

    # -- the single kind-aware resolution entry point -----------------------
    def resolve_role(
        self, role_id: str, *, fallback: Optional[str] = "paragraph"
    ) -> ResolvedOp:
        """Resolve ``role_id`` to its kind-appropriate :class:`ResolvedOp`.

        This is the spine both pptx and xlsx generators call. Resolution is the
        same for every kind - look the role up in ``profile['roles']`` (falling
        back to ``fallback`` when absent) and return its resolver op verbatim -
        but the brand guarantee differs only in *which* resolver types are
        legal, which is data (:data:`schema.LEGAL_RESOLVER_TYPES`), not code.

        A resolver whose ``type`` is not legal for ``kind`` is refused: in strict
        mode it raises; otherwise it degrades to an empty-resolver stub op so a
        fabricated cross-kind target can never reach a generator.
        """
        # Gate the body size/color default on the ORIGINALLY requested role, before
        # any fallback reassignment: a heading.* that falls back to the paragraph
        # style must still be treated as a heading for the body-default gate, so the
        # body size/color never leaks onto it (the fallback only swaps the resolver).
        requested_role_id = role_id
        entry = self.roles.get(role_id)
        if entry is None and fallback:
            entry = self.roles.get(fallback)
            role_id = fallback
        if entry is None:
            if self.strict:
                raise ResolveError(f"role {role_id!r} is not resolved in profile")
            # LAST-RESORT reroute: the base resolution is a genuine stub (the role,
            # post-fallback, is not in the profile). When a learned lesson re-points
            # ``requested_role_id`` at a declared role, return THAT role's op instead
            # of the empty stub. No lesson ⇒ the stub is returned verbatim (byte-id).
            return self._apply_reroute(
                requested_role_id,
                self._stub(role_id, appearance_role_id=requested_role_id),
                fallback=fallback,
            )

        resolver = dict(entry.get("resolver") or {})
        status = entry.get("status", schema.Status.STUB.value)
        confidence = float(entry.get("confidence", 0.0))
        appearance = self._merge_appearance(
            entry.get("appearance"), role_id=requested_role_id
        )

        # number_format mask swap (Cluster B): a learned lesson may re-point a role's
        # number-format MASK to another mask the shell already uses. Applied to the
        # role's OWN resolver in place, BEFORE the legal-type gate, so the swapped op
        # is still membership-gated below (``number_format`` stays legal for xlsx).
        # This is NOT a last-resort step: a number_format role with a registered swap
        # is healthy, and the swap only changes which shell-backed mask it carries -
        # it can never invent a mask (that is ``check_override_targets``'s job to
        # prove). No swap for this role ⇒ ``resolver`` is untouched (byte-identical).
        if (
            resolver.get("type") == schema.ResolverType.NUMBER_FORMAT.value
            and requested_role_id in self._number_format_swaps
        ):
            resolver["number_format"] = self._number_format_swaps[requested_role_id]

        # Brand guarantee, enforced in ONE place for all three kinds: never emit
        # a concrete resolver whose type the kind cannot legally carry. An empty
        # resolver ({}) is always fine (an honest stub); a typed resolver must be
        # legal for this kind.
        rtype = resolver.get("type")
        if resolver and rtype is not None and rtype not in self._legal_types:
            if self.strict:
                raise ResolveError(
                    f"role {role_id!r} resolver type {rtype!r} is not legal for "
                    f"kind {self.kind!r} (legal: {sorted(self._legal_types)})"
                )
            # LAST-RESORT reroute: the legal-type gate JUST produced a stub (the
            # role's own resolver type is foreign for this kind). Re-point if a lesson
            # exists; otherwise the stub is returned verbatim.
            return self._apply_reroute(
                requested_role_id,
                self._stub(role_id, appearance_role_id=requested_role_id),
                fallback=fallback,
            )

        op = ResolvedOp(
            role_id=role_id,
            resolver=resolver,
            status=status,
            confidence=confidence,
            kind=self.kind,
            appearance=appearance,
        )
        # LAST-RESORT reroute on the final return ONLY when the resolved op is a
        # genuine stub - i.e. the role exists but its resolver is empty ({}). A
        # HEALTHY typed resolve is NEVER rerouted (firing there would silently change
        # branded output), so the trigger is exactly the QA-flagged empty-resolver
        # condition. No lesson / non-empty resolver ⇒ ``op`` is returned unchanged.
        if not resolver:
            return self._apply_reroute(requested_role_id, op, fallback=fallback)
        return op

    def _apply_reroute(
        self,
        requested_role_id: str,
        stub_op: ResolvedOp,
        *,
        fallback: Optional[str],
    ) -> ResolvedOp:
        """LAST-RESORT re-point of a STUB resolution to a learned target role.

        Fires ONLY when ``requested_role_id`` carries a learned ``reroute_role``
        lesson AND the base resolution was a genuine stub (the caller only invokes
        this on a stub op). It does a SINGLE-HOP resolve of the target - overrides
        are NOT re-applied to the target (``_in_reroute`` depth guard), mirroring the
        bounded ``comprehension._apply_slots`` guard - so a chain/cycle can never
        recurse here (``learn`` proves the reroute graph acyclic before writing; this
        is the runtime backstop).

        The rerouted op inherits the TARGET role's existing, shell-proven resolver
        verbatim (still through the legal-type gate, since the single-hop resolve runs
        the same gate), so an override can never invent a style/font/hex. But it PINS
        ``role_id=requested_role_id`` and keeps the REQUESTED role's appearance, so
        the ``_allows_body_default`` heading exclusion still keys on the ORIGINAL
        role - a rerouted ``heading.*`` never has the body default leaked onto it.

        When there is no lesson for this role (the common case, and ALWAYS when
        ``overrides_are_present`` is False ⇒ ``self._overrides == {}``) the original
        stub op is returned verbatim, so the no-overrides path is byte-identical.
        """
        target = self._overrides.get(requested_role_id)
        if not target or self._in_reroute:
            return stub_op
        # Single-hop: resolve the TARGET without re-applying overrides to it, and with
        # NO fallback - a missing reroute target must stub out HONESTLY rather than
        # silently inheriting the requested role's fallback (check_override_targets
        # already rejects an unbacked target on the live path; this keeps the runtime
        # last-resort semantics honest if a target ever slips through).
        self._in_reroute = True
        try:
            target_op = self.resolve_role(target, fallback=None)
        finally:
            self._in_reroute = False
        # The target's own (shell-proven) resolver, but pinned to the requested role
        # id and carrying the requested role's appearance (not the target's), so the
        # body-default gate still excludes a rerouted heading.
        return ResolvedOp(
            role_id=requested_role_id,
            resolver=target_op.resolver,
            status=target_op.status,
            confidence=target_op.confidence,
            kind=self.kind,
            appearance=stub_op.appearance,
        )

    def resolve_color(self, token: Optional[str]) -> Optional[dict]:
        """Resolve a run COLOR palette token to its captured color ``ref`` object.

        Returns ``theme.palette[token]['ref']`` (the ``{'kind': 'theme'|'hex', ...}``
        object the writer applies via ``_brand_run_color``) when ``token`` is a key
        of the palette, else ``None`` (LENIENT - an unknown token, an empty/absent
        palette, or a falsy token). Resolution is off ``theme.palette`` ONLY (one
        namespace); ``palette_roles`` is never consulted. The model NAMES a token;
        the deterministic palette AUTHORS the actual color - the resolver never
        invents one, so off-brand color stays impossible by construction. An
        unresolved token is the generator's cue to leave the run inherited and record
        a graceful INFO finding (it never fabricates a color)."""
        if not token:
            return None
        entry = self._palette.get(token)
        if not isinstance(entry, dict):
            return None
        ref = entry.get("ref")
        if not isinstance(ref, dict):
            return None
        return self._enrich_theme_hex(ref)

    def _enrich_theme_hex(self, color: Optional[dict]) -> Optional[dict]:
        """Add the concrete ``hex`` to a THEME color from the profile's own
        ``theme.colors`` so the writer can carry it in ``w:color@w:val`` ALONGSIDE the
        ``themeColor``: Word honors the live themeColor, while renderers that ignore a
        run-level themeColor (e.g. headless LibreOffice) still show the real brand
        color instead of the black fallback. It also lets the writer realize a token
        whose clrScheme-slot name has no WordprocessingML themeColor member (dk1/lt1/
        hlink ...) via the hex. The token is mapped to its clrScheme slot through the
        shared :data:`color.WML_THEME_TO_SLOT` (a WML name) or used verbatim (already a
        slot). The hex is the profile's captured value (no literal in the writer);
        left unchanged when not a theme color or when theme.colors lacks the slot."""
        if not isinstance(color, dict) or color.get("kind") != "theme":
            return color
        token = color.get("theme")
        if not token or "hex" in color:
            return color
        slot = colorutil.WML_THEME_TO_SLOT.get(token, token)
        hexval = (self._theme_colors.get(slot) or {}).get("hex")
        return {**color, "hex": hexval} if hexval else color

    @staticmethod
    def _read_default_appearance(profile: dict) -> dict:
        """The document-level captured body typography, as an ``appearance`` dict.

        Three INDEPENDENT axes: the body font/size live under ``theme.fonts.body``
        and the body color under the additive ``theme.text.body`` key. Each axis is
        included only when the template actually captured it, so a pre-feature
        profile yields an empty dict (behavior unchanged)."""
        if not profile:
            return {}
        theme = profile.get("theme") or {}
        body = (theme.get("fonts") or {}).get("body") or {}
        text_body = (theme.get("text") or {}).get("body") or {}
        out: dict[str, Any] = {}
        latin = body.get("latin")
        if latin:
            out["font"] = {"latin": latin}
        size_hp = body.get("size_hp")
        if size_hp:
            out["size_hp"] = size_hp
        color = text_body.get("color")
        if color:
            out["color"] = color
        return out

    def _merge_appearance(
        self, role_appearance: Optional[dict], *, role_id: Optional[str] = None
    ) -> dict:
        """Effective appearance for a role across three INDEPENDENT axes (font, size,
        color): the role's own captured value wins on each axis; otherwise the
        document-level body value fills in.

        CRITICAL family gate: the FONT body-default flows to EVERY role (v1
        behavior) - a missing font means the document baseline, which is right for a
        heading too. But the body SIZE and body COLOR defaults apply ONLY to the
        body/paragraph family (and the body stub), NEVER to ``heading.*`` or any
        other role that resolves to a real named style with its own intrinsic
        larger/colored typography: forcing the body size/color as DIRECT run
        formatting would override that style's intrinsic value (shrinking/recoloring
        headings). A role-SPECIFIC captured size/color still applies to any role."""
        role_appearance = role_appearance or {}
        default = self._default_appearance

        out: dict[str, Any] = {}
        latin = (role_appearance.get("font") or {}).get("latin") or (
            default.get("font") or {}
        ).get("latin")
        if latin:
            out["font"] = {"latin": latin}

        # The body size/color defaults are gated on the role FAMILY; a role-specific
        # captured value always applies, the body default only on paragraph/body.
        allow_body_default = self._allows_body_default(role_id)
        size_hp = role_appearance.get("size_hp") or (
            default.get("size_hp") if allow_body_default else None
        )
        if size_hp:
            out["size_hp"] = size_hp
        color = role_appearance.get("color") or (
            default.get("color") if allow_body_default else None
        )
        if color:
            # Enrich a THEME color with its concrete hex (same as resolve_color), so
            # an AUTO body/role theme color renders in headless LibreOffice too and a
            # clrScheme-slot color is realizable via the hex at apply time.
            out["color"] = self._enrich_theme_hex(color)
        return out

    @staticmethod
    def _allows_body_default(role_id: Optional[str]) -> bool:
        """True when the document body SIZE/COLOR default may flow to this role.

        Only the body itself: a ``paragraph`` family role (``paragraph`` and its
        ``paragraph.*`` variants) or the body stub (``role_id`` is ``None`` /
        ``paragraph``). A ``heading.*`` (or any other named-style role with its own
        intrinsic size/color) is excluded so the body size/color never overrides the
        style's intrinsic value."""
        if role_id is None:
            return True
        family, _ = schema.parse_role_id(role_id)
        return family == "paragraph"

    def _stub(
        self, role_id: str, *, appearance_role_id: Optional[str] = None
    ) -> ResolvedOp:
        """An honest no-target op: empty resolver, zero confidence. It still carries
        the document-level body typography (the body stub is the paragraph fallback)
        so a fallback paragraph is branded.

        ``appearance_role_id`` (when given) gates the body size/color default on the
        ORIGINALLY requested role rather than the post-fallback ``role_id``, so a
        heading that fell through to a stub is still excluded from the body default.
        """
        gate = appearance_role_id if appearance_role_id is not None else role_id
        return ResolvedOp(
            role_id,
            {},
            schema.Status.STUB.value,
            0.0,
            self.kind,
            appearance=self._merge_appearance(None, role_id=gate),
        )

    # -- docx IR-stream dispatch -------------------------------------------
    def resolve_block(self, block: ir.Block) -> ResolvedOp:
        """Resolve a semantic IR block to a role op (the docx dispatch).

        Dispatches on the block's semantic family, returning the
        kind-appropriate resolver from the profile via :meth:`resolve_role`. For
        docx every role resolves to a ``named_style`` op. pptx/xlsx do not drive
        generation from this IR-block stream; they call :meth:`resolve_role`
        directly with their own role ids, sharing the same spine and the same
        brand guarantee.
        """
        if isinstance(block, ir.Heading):
            return self.resolve_role(
                schema.role_id("heading", block.level), fallback="paragraph"
            )
        if isinstance(block, ir.Paragraph):
            fallback = (
                schema.role_id("paragraph", block.variant)
                if block.variant
                else "paragraph"
            )
            return self.resolve_role(fallback, fallback="paragraph")
        if isinstance(block, ir.Callout):
            return self.resolve_role(
                schema.role_id("callout", block.intent), fallback="paragraph"
            )
        if isinstance(block, ir.ListBlock):
            # API-completeness only: the docx generator resolves lists PER ITEM via
            # resolve_list_item (it needs each item's level), so this whole-block
            # branch is not exercised by the generators - it returns the level-1
            # default so a direct resolve_block(ListBlock) caller still gets a list op.
            return self.resolve_list_item(block, None)
        if isinstance(block, ir.Table):
            return self.resolve_role(
                schema.role_id("table", block.role or "default"), fallback=None
            )
        if isinstance(block, ir.Caption):
            return self.resolve_role("caption", fallback="paragraph")
        if isinstance(block, ir.Quote):
            return self.resolve_role("quote", fallback="paragraph")
        return self.resolve_role("paragraph", fallback="paragraph")

    def resolve_list_item(
        self, block: ir.ListBlock, item: Optional[ir.ListItem]
    ) -> ResolvedOp:
        """Resolve the list role for one item, honoring its nesting ``level``.

        The list family (bullet/number) comes from the block; the *level* comes
        from the item (``item.level`` is 0-based, role ids are 1-based, so
        ``level+1``). When no per-level role exists in the profile the resolver
        falls back to the level-1 list role, then to ``paragraph`` - so a deeply
        nested item still gets a list style rather than being dropped.
        """
        family = "number" if block.ordered else "bullet"
        level = (item.level + 1) if item is not None else 1
        rid = schema.role_id("list", family, level)
        if level > 1 and not schema.supports_role(self.profile, rid):
            rid = schema.role_id("list", family, 1)
        return self.resolve_role(rid, fallback="paragraph")
