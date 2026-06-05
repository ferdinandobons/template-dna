# SPDX-License-Identifier: MIT
"""DOCX style lookup — the single place that maps a resolver dict to a live style.

This is the one bridge between a profile resolver (``{"style_id", "style_name",
...}``) and a python-docx ``Document``. It is deliberately tiny and shared by
``generate.py`` and ``cover.py`` so the brand guarantee has exactly one
enforcement point: a writer can only ever apply a style the resolver named AND
the shell actually carries. There is no literal style name anywhere in the
writers — every style comes from the profile via :func:`lookup_style`.
"""
from __future__ import annotations


def lookup_style(doc, resolver: dict):
    """Return the live ``doc`` style matching ``resolver`` (id or name), or None.

    A resolver matches a style when its ``style_id`` equals the style's
    ``style_id`` OR its ``style_name`` equals the style's ``name``. An empty
    resolver (neither key set) never matches — returning None here is what stops
    the empty-resolver false match (every style's missing attributes compared
    against ``None``).

    Returns None (not ``Normal``) when nothing matches; callers decide whether a
    miss is a degradation (skip styling) or an error (emit a finding). The lookup
    itself stays silent and pure.
    """
    style_id = resolver.get("style_id")
    style_name = resolver.get("style_name")
    # Guard the empty-resolver false match: with neither key set there is nothing
    # to resolve, so do not let ``getattr(style, "style_id", None) == None`` (true
    # for styles without an id) masquerade as a match.
    if not style_id and not style_name:
        return None
    for style in doc.styles:
        if style_id and getattr(style, "style_id", None) == style_id:
            return style
        if style_name and getattr(style, "name", None) == style_name:
            return style
    return None
