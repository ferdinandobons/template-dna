# SPDX-License-Identifier: MIT
"""Pure Brand Profile resolver.

The resolver is deliberately small in M1: it maps semantic IID blocks to role
entries already extracted into ``profile.json``. It performs no I/O and never
contains brand literals; all concrete names and ids come from the profile.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Optional

from brandkit.ir import model as ir
from brandkit.profile import schema


@dataclass(frozen=True)
class ResolvedOp:
    role_id: str
    resolver: dict[str, Any]
    status: str
    confidence: float


class ResolveError(LookupError):
    pass


class ProfileResolver:
    def __init__(self, profile: dict, *, strict: bool = False) -> None:
        self.profile = profile
        self.strict = strict
        self.roles = profile.get("roles") or {}

    def resolve_role(self, role_id: str, *, fallback: Optional[str] = "paragraph") -> ResolvedOp:
        entry = self.roles.get(role_id)
        if entry is None and fallback:
            entry = self.roles.get(fallback)
            role_id = fallback
        if entry is None:
            if self.strict:
                raise ResolveError(f"role {role_id!r} is not resolved in profile")
            return ResolvedOp(role_id, {}, schema.Status.STUB.value, 0.0)
        return ResolvedOp(
            role_id=role_id,
            resolver=dict(entry.get("resolver") or {}),
            status=entry.get("status", schema.Status.STUB.value),
            confidence=float(entry.get("confidence", 0.0)),
        )

    def resolve_block(self, block: ir.Block) -> ResolvedOp:
        if isinstance(block, ir.Heading):
            return self.resolve_role(schema.role_id("heading", block.level), fallback="paragraph")
        if isinstance(block, ir.Paragraph):
            fallback = schema.role_id("paragraph", block.variant) if block.variant else "paragraph"
            return self.resolve_role(fallback, fallback="paragraph")
        if isinstance(block, ir.Callout):
            return self.resolve_role(schema.role_id("callout", block.intent), fallback="paragraph")
        if isinstance(block, ir.ListBlock):
            family = "number" if block.ordered else "bullet"
            return self.resolve_role(schema.role_id("list", family, 1), fallback="paragraph")
        if isinstance(block, ir.Table):
            return self.resolve_role(schema.role_id("table", block.role or "default"), fallback=None)
        if isinstance(block, ir.Caption):
            return self.resolve_role("caption", fallback="paragraph")
        if isinstance(block, ir.Quote):
            return self.resolve_role("quote", fallback="paragraph")
        return self.resolve_role("paragraph", fallback="paragraph")


def resolve_block(profile: dict, block: ir.Block, *, strict: bool = False) -> ResolvedOp:
    return ProfileResolver(profile, strict=strict).resolve_block(block)

