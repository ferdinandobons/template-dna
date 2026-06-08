# SPDX-License-Identifier: MIT
"""Shared OOXML name helpers - one place for Clark-notation qualified names.

Every format walker (docx structure, pptx structure, the color resolver) used to
carry its own copy of two tiny but error-prone helpers: a ``{NS}tag`` qualifier
bound to that format's namespace, and a ``local_name(tag)`` that strips the Clark
prefix back off. This module is the single home for both, so the engine never
re-derives raw lxml name glue per format (the same intent :mod:`brandkit.ooxml.pack`
states for (un)packing).

Two shapes are offered, deliberately:

- :func:`local_name` - the namespace-AGNOSTIC direction (``{NS}tag`` -> ``tag``),
  robust to a non-string lxml tag (comments / PIs carry a callable ``.tag``).
- :func:`qname` / :func:`make_qn` - the namespace-SPECIFIC direction. A walker
  binds ``make_qn("w")`` once and then calls it with a BARE local name
  (``w("pPr")`` -> ``{...wordprocessingml...}pPr``). This is the shape the docx and
  color helpers already used, and is distinct from :func:`brandkit.ooxml.pack.qn`,
  which takes a ``"prefix:tag"`` string. Both coexist: ``pack.qn`` is the
  prefix-string flavor, ``make_qn`` is the per-namespace bare-tag flavor.

Namespace URIs are sourced from :data:`brandkit.ooxml.pack.NAMESPACES`, so there
is one registry of OOXML namespaces for the whole engine.
"""

from __future__ import annotations

from typing import Callable


def local_name(tag) -> str:
    """Return the local (un-namespaced) name of an lxml tag.

    Robust to a non-string ``tag`` (lxml comment / processing-instruction nodes
    carry a callable ``.tag``): those yield ``""``. A Clark-notation
    ``{uri}local`` collapses to ``local``; a bare name is returned unchanged.
    """
    if not isinstance(tag, str):
        return ""
    return tag.rsplit("}", 1)[-1]


def qname(namespace: str, tag: str) -> str:
    """Return the Clark-notation qualified name ``{namespace}tag``.

    ``namespace`` is a raw URI (not a prefix). This is the primitive both
    :func:`make_qn` and the per-format ``w``/``a`` helpers build on.
    """
    return f"{{{namespace}}}{tag}"


def make_qn(prefix: str) -> Callable[[str], str]:
    """Return a namespace-bound qualifier for ``prefix`` (e.g. ``"w"``/``"a"``).

    The returned callable maps a BARE local name to ``{uri}tag`` for the
    namespace registered under ``prefix`` in
    :data:`brandkit.ooxml.pack.NAMESPACES`::

        w = make_qn("w")
        w("pPr")  # -> "{http://...wordprocessingml...}pPr"

    Raises:
        KeyError: if ``prefix`` is not a registered OOXML namespace prefix.
    """
    # Imported lazily so importing this module never pulls in lxml (via
    # :mod:`brandkit.ooxml.pack`); :func:`local_name` / :func:`qname` stay usable
    # from import-safe (no-lxml) callers such as :mod:`brandkit.common.color`.
    from brandkit.ooxml.pack import NAMESPACES

    namespace = NAMESPACES[prefix]

    def qn(tag: str) -> str:
        return f"{{{namespace}}}{tag}"

    return qn
