# SPDX-License-Identifier: MIT
"""Pure Brand Profile resolver - the single, KIND-AWARE resolution spine.

The resolver maps semantic roles to the concrete resolver op the *shell*
exposes, dispatching on the profile ``kind``:

  - ``docx`` -> ``named_style``   (``{"type", "style_id", "style_name"}``)
  - ``pptx`` -> ``placeholder``   (``{"type", "layout", "ph_idx", "ph_type"}``)
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

from dataclasses import dataclass
from typing import Any, Optional

from brandkit.ir import model as ir
from brandkit.profile import schema


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
        # The resolver types this profile's kind is allowed to emit. An unknown
        # kind yields an empty set, so every concrete resolver is treated as
        # foreign (fail-closed) rather than blindly trusted.
        self._legal_types: frozenset[str] = schema.LEGAL_RESOLVER_TYPES.get(
            self.kind, frozenset()
        ) if self.kind else frozenset()

    # -- the single kind-aware resolution entry point -----------------------
    def resolve_role(self, role_id: str, *, fallback: Optional[str] = "paragraph") -> ResolvedOp:
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
        entry = self.roles.get(role_id)
        if entry is None and fallback:
            entry = self.roles.get(fallback)
            role_id = fallback
        if entry is None:
            if self.strict:
                raise ResolveError(f"role {role_id!r} is not resolved in profile")
            return self._stub(role_id)

        resolver = dict(entry.get("resolver") or {})
        status = entry.get("status", schema.Status.STUB.value)
        confidence = float(entry.get("confidence", 0.0))

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
            return self._stub(role_id)

        return ResolvedOp(
            role_id=role_id,
            resolver=resolver,
            status=status,
            confidence=confidence,
            kind=self.kind,
        )

    def _stub(self, role_id: str) -> ResolvedOp:
        """An honest no-target op: empty resolver, zero confidence."""
        return ResolvedOp(role_id, {}, schema.Status.STUB.value, 0.0, self.kind)

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
            return self.resolve_role(schema.role_id("heading", block.level), fallback="paragraph")
        if isinstance(block, ir.Paragraph):
            fallback = schema.role_id("paragraph", block.variant) if block.variant else "paragraph"
            return self.resolve_role(fallback, fallback="paragraph")
        if isinstance(block, ir.Callout):
            return self.resolve_role(schema.role_id("callout", block.intent), fallback="paragraph")
        if isinstance(block, ir.ListBlock):
            return self.resolve_list_item(block, None)
        if isinstance(block, ir.Table):
            return self.resolve_role(schema.role_id("table", block.role or "default"), fallback=None)
        if isinstance(block, ir.Caption):
            return self.resolve_role("caption", fallback="paragraph")
        if isinstance(block, ir.Quote):
            return self.resolve_role("quote", fallback="paragraph")
        return self.resolve_role("paragraph", fallback="paragraph")

    def resolve_list_item(self, block: ir.ListBlock, item: Optional[ir.ListItem]) -> ResolvedOp:
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


def resolve_block(profile: dict, block: ir.Block, *, strict: bool = False) -> ResolvedOp:
    return ProfileResolver(profile, strict=strict).resolve_block(block)


def resolve_role(profile: dict, role_id: str, *, fallback: Optional[str] = "paragraph", strict: bool = False) -> ResolvedOp:
    """Module-level convenience: resolve one role against ``profile``.

    Mirrors :func:`resolve_block`; the kind-aware spine both pptx and xlsx route
    through when they only need a single role (not the docx IR stream).
    """
    return ProfileResolver(profile, strict=strict).resolve_role(role_id, fallback=fallback)
