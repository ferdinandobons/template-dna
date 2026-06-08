# SPDX-License-Identifier: MIT
"""Component / section expansion.

``component`` and ``section`` are first-class IID blocks: a ``component`` names a
``profile["components"][ref]`` entry and a ``section`` names a
``profile["sections"][ref]`` entry, each carrying a ``blocks`` template of
primitive sub-blocks. The writers cannot render a ``component``/``section`` block
directly, so this hook expands every such block into its primitive sub-blocks
*before* resolution. A reference the profile does not define is rejected (raised)
rather than silently dropped - off-brand / missing content must never pass
silently.
"""

from __future__ import annotations

import re

from brandkit.ir.model import IntermediateDocument, block_from_dict
from brandkit.ir.model import Component, Section

#: A ``{{name}}`` placeholder inside a fragment template's text, filled from the
#: referencing block's ``slots``. The name is a conservative identifier so it
#: cannot accidentally swallow real text containing braces.
_SLOT_TOKEN = re.compile(r"\{\{\s*([A-Za-z0-9_.\-]+)\s*\}\}")


class ComponentExpansionError(ValueError):
    """Raised when a component/section ref is not defined in the profile."""


def expand_components(
    document: IntermediateDocument, profile: dict
) -> IntermediateDocument:
    """Return a document whose component/section blocks are expanded to primitives.

    A ``component``/``section`` block is replaced, in place, by the primitive
    sub-blocks of its profile definition (``profile["components"][ref]["blocks"]``
    / ``profile["sections"][ref]["blocks"]``). Expansion is recursive (a section
    may contain components). When a ref is undefined the call raises
    :class:`ComponentExpansionError` so the missing fragment is loud, not dropped.

    Both IID legs route through here (docx and pptx ``generate`` call this at the
    top), so a profile-defined fragment expands identically across formats. The
    referencing block's ``slots`` parameterize the template: a ``{{name}}`` token
    in the template's text is replaced with ``slots[name]`` (an unfilled token
    resolves to the empty string, never leaked verbatim). The registries are
    populated at comprehend time (the model proposes fragments, the merge boundary
    validates them fail-closed and derives them in); this hook is the consumer.
    """
    components = profile.get("components") or {}
    sections = profile.get("sections") or {}
    document.blocks = _expand_blocks(
        document.blocks, components, sections, _depth=0, _emitted=[0]
    )
    return document


# Depth bounds a self-/mutually-referential registry (a cycle); the node budget
# bounds runaway FAN-OUT (a tree of refs that stays within the depth cap can still
# expand to astronomically many primitives) - both fail LOUD, never hang/OOM.
_MAX_DEPTH = 16
_MAX_EXPANDED_BLOCKS = 50_000


def _expand_blocks(blocks, components, sections, *, _depth: int, _emitted: list[int]):
    if _depth > _MAX_DEPTH:
        raise ComponentExpansionError(
            "component/section expansion exceeded max depth (cycle?)"
        )
    out = []
    for block in blocks:
        if isinstance(block, Component):
            out.extend(
                _expand_ref(
                    block.ref,
                    components,
                    "component",
                    components,
                    sections,
                    _depth,
                    block.slots,
                    _emitted,
                )
            )
        elif isinstance(block, Section):
            out.extend(
                _expand_ref(
                    block.ref,
                    sections,
                    "section",
                    components,
                    sections,
                    _depth,
                    block.slots,
                    _emitted,
                )
            )
        else:
            _emitted[0] += 1
            if _emitted[0] > _MAX_EXPANDED_BLOCKS:
                raise ComponentExpansionError(
                    f"component/section expansion exceeded {_MAX_EXPANDED_BLOCKS} "
                    "primitive blocks (runaway fan-out?)"
                )
            out.append(block)
    return out


def _expand_ref(ref, registry, kind, components, sections, depth, slots, _emitted):
    definition = registry.get(ref)
    if definition is None:
        raise ComponentExpansionError(
            f"{kind} ref {ref!r} is not defined in the profile"
        )
    raw_blocks = definition.get("blocks")
    if not isinstance(raw_blocks, list):
        raise ComponentExpansionError(f"{kind} {ref!r} has no 'blocks' template")
    # Slot parameterization: substitute ``{{name}}`` tokens in the template's text
    # from the referencing block's ``slots``. Done at the dict level (a deep copy),
    # BEFORE block_from_dict, so it works uniformly for every text-bearing field and
    # never mutates the shared profile registry. Run ALWAYS (even with no slots) so
    # an unfilled token resolves to the empty string and is never leaked into the
    # output as a literal ``{{...}}``.
    raw_blocks = [_apply_slots(b, slots or {}) for b in raw_blocks]
    sub = [block_from_dict(b) for b in raw_blocks]
    return _expand_blocks(
        sub, components, sections, _depth=depth + 1, _emitted=_emitted
    )


def _apply_slots(value, slots, _depth: int = 0):
    """Deep-copy ``value`` substituting ``{{name}}`` tokens from ``slots``.

    Strings get token replacement (an unfilled or ``None`` slot -> ``""``, never the
    literal token or the string ``"None"``); lists/dicts recurse; other scalars pass
    through. Returns NEW containers so the profile registry the template lives in is
    never mutated. A depth bound turns a pathologically nested template into a
    fail-closed error instead of an unhandled ``RecursionError``.
    """
    if _depth > 64:
        raise ComponentExpansionError("slot substitution exceeded max nesting depth")
    if isinstance(value, str):

        def _sub(match):
            filled = slots.get(match.group(1))
            return "" if filled is None else str(filled)

        return _SLOT_TOKEN.sub(_sub, value)
    if isinstance(value, list):
        return [_apply_slots(v, slots, _depth + 1) for v in value]
    if isinstance(value, dict):
        return {k: _apply_slots(v, slots, _depth + 1) for k, v in value.items()}
    return value
