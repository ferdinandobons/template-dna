# SPDX-License-Identifier: MIT
"""Shared link-target safety check for the docx/pptx hyperlink writers.

Author content is untrusted relative to whoever opens the generated file, so a
hyperlink target must be restricted to a safe scheme: a crafted ``file:``,
``smb:``, ``javascript:`` or ``data:`` link must never be wired into an otherwise
on-brand document. A relative / fragment / same-document target (no scheme) is
allowed; everything outside the allowlist is refused at the single chokepoint.
"""

from __future__ import annotations

import urllib.parse

#: Schemes a generated document may safely link to.
SAFE_LINK_SCHEMES = frozenset({"http", "https", "mailto", "tel"})


def is_safe_link_url(url: str) -> bool:
    """Return True iff ``url`` is a safe hyperlink target (allowlisted scheme).

    Empty/relative/fragment targets (no scheme) are safe; ``file``/``smb``/
    ``javascript``/``data``/etc. are not. Never raises.
    """
    if not isinstance(url, str) or not url.strip():
        return False
    try:
        scheme = urllib.parse.urlsplit(url.strip()).scheme.lower()
    except ValueError:
        return False
    return scheme == "" or scheme in SAFE_LINK_SCHEMES
