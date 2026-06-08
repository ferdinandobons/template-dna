# SPDX-License-Identifier: MIT
"""Shared OOXML complex-field walk - the begin/separate/end + instrText skeleton.

A WordprocessingML *complex field* is delimited by ``w:fldChar`` characters
(``fldCharType`` ``begin`` / ``separate`` / ``end``) and carries its field code in
one or more ``w:instrText`` runs between ``begin`` and ``separate``. Fields nest
(a ``PAGEREF`` inside a rendered ``TOC`` entry), so any code that attributes an
instruction to *its own* enclosing field must walk ``fldChar``/``instrText`` in
document order while tracking a nesting stack.

That iteration-and-discrimination boilerplate appeared three times verbatim in the
docx structure walker (TOC inventory, dirty-mark, outline-cache rewrite), each with
the SAME skeleton but a DIFFERENT per-field payload (one builds index spans, one
flips a ``w:dirty`` attribute, one accumulates the instruction string). This module
factors out only the shared part: :func:`iter_complex_field_events` yields a
normalized ``(kind, element)`` event per ``fldChar``/``instrText`` so each caller
keeps its own stack and payload handling, without re-deriving the
local-name + ``fldCharType`` discrimination every time.
"""

from __future__ import annotations

from typing import Iterator, Tuple

from brandkit.ooxml.names import local_name, make_qn

_w = make_qn("w")
_FLDCHAR = _w("fldChar")
_INSTRTEXT = _w("instrText")
_FLDCHARTYPE = _w("fldCharType")

# Event kinds yielded by :func:`iter_complex_field_events`.
BEGIN = "begin"
SEPARATE = "separate"
END = "end"
INSTR = "instr"


def iter_complex_field_events(root) -> Iterator[Tuple[str, object]]:
    """Yield ``(kind, element)`` for every complex-field marker under ``root``.

    Walks ``root.iter(w:fldChar, w:instrText)`` in document order and classifies
    each hit:

    - a ``w:fldChar`` yields its ``fldCharType`` as ``kind`` - one of
      :data:`BEGIN` / :data:`SEPARATE` / :data:`END` (an unknown/absent type
      yields that raw string, so callers see exactly what the document carried);
    - a ``w:instrText`` yields :data:`INSTR`.

    ``element`` is always the live lxml node, so a caller can read ``.text`` off an
    ``INSTR`` event or set an attribute (e.g. ``w:dirty``) on a ``BEGIN`` fldChar.
    The caller owns the nesting stack and whatever per-field payload it builds; this
    generator owns only the iteration order and the begin/separate/end + instrText
    discrimination that all callers share.
    """
    for el in root.iter(_FLDCHAR, _INSTRTEXT):
        if local_name(el.tag) == "fldChar":
            # An absent fldCharType yields "" (not None), matching the docstring's
            # promise that callers see exactly the raw string the document carried.
            yield el.get(_FLDCHARTYPE, ""), el
        else:
            yield INSTR, el
