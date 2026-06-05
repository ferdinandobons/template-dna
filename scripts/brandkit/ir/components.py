# SPDX-License-Identifier: MIT
"""Component / section expansion.

``component`` and ``section`` are first-class IID blocks: a ``component`` names a
``profile["components"][ref]`` entry and a ``section`` names a
``profile["sections"][ref]`` entry, each carrying a ``blocks`` template of
primitive sub-blocks. The writers cannot render a ``component``/``section`` block
directly, so this hook expands every such block into its primitive sub-blocks
*before* resolution. A reference the profile does not define is rejected (raised)
rather than silently dropped — off-brand / missing content must never pass
silently.
"""
from __future__ import annotations

from brandkit.ir.model import IntermediateDocument, block_from_dict
from brandkit.ir.model import Component, Section


class ComponentExpansionError(ValueError):
    """Raised when a component/section ref is not defined in the profile."""


def expand_components(document: IntermediateDocument, profile: dict) -> IntermediateDocument:
    """Return a document whose component/section blocks are expanded to primitives.

    A ``component``/``section`` block is replaced, in place, by the primitive
    sub-blocks of its profile definition (``profile["components"][ref]["blocks"]``
    / ``profile["sections"][ref]["blocks"]``). Expansion is recursive (a section
    may contain components). When a ref is undefined the call raises
    :class:`ComponentExpansionError` so the missing fragment is loud, not dropped.
    """
    components = profile.get("components") or {}
    sections = profile.get("sections") or {}
    document.blocks = _expand_blocks(document.blocks, components, sections, _depth=0)
    return document


def _expand_blocks(blocks, components, sections, *, _depth: int):
    if _depth > 16:
        raise ComponentExpansionError("component/section expansion exceeded max depth (cycle?)")
    out = []
    for block in blocks:
        if isinstance(block, Component):
            out.extend(_expand_ref(block.ref, components, "component", components, sections, _depth))
        elif isinstance(block, Section):
            out.extend(_expand_ref(block.ref, sections, "section", components, sections, _depth))
        else:
            out.append(block)
    return out


def _expand_ref(ref, registry, kind, components, sections, depth):
    definition = registry.get(ref)
    if definition is None:
        raise ComponentExpansionError(f"{kind} ref {ref!r} is not defined in the profile")
    raw_blocks = definition.get("blocks")
    if not isinstance(raw_blocks, list):
        raise ComponentExpansionError(f"{kind} {ref!r} has no 'blocks' template")
    sub = [block_from_dict(b) for b in raw_blocks]
    return _expand_blocks(sub, components, sections, _depth=depth + 1)
