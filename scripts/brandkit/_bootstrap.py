# SPDX-License-Identifier: MIT
"""Import-path bootstrap for skill-local shims."""
from __future__ import annotations

import os
import sys
from pathlib import Path


def plugin_root(start: str | Path | None = None) -> Path:
    env = os.environ.get("CLAUDE_PLUGIN_ROOT")
    if env:
        return Path(env)
    cur = Path(start or __file__).resolve()
    for parent in [cur] + list(cur.parents):
        if (parent / ".claude-plugin").is_dir():
            return parent
    return Path(__file__).resolve().parents[2]


def ensure_scripts_path(start: str | Path | None = None) -> Path:
    root = plugin_root(start)
    scripts = root / "scripts"
    if str(scripts) not in sys.path:
        sys.path.insert(0, str(scripts))
    return root

