# SPDX-License-Identifier: MIT
"""M1 component expansion.

Components and sections are part of the public IID vocabulary from day one, but
M1 does not implement reusable fragment rendering yet. The expansion hook keeps
the generate pipeline shape stable and passes primitive blocks through.
"""
from __future__ import annotations

from brandkit.ir.model import IntermediateDocument


def expand_components(document: IntermediateDocument, profile: dict) -> IntermediateDocument:
    return document

