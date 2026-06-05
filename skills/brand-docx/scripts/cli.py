# SPDX-License-Identifier: MIT
from __future__ import annotations

import os
import sys
from pathlib import Path


def _root(start: Path) -> Path:
    env_root = os.environ.get("TEMPLATE_DNA_ROOT")
    if env_root:
        root = Path(env_root).expanduser().resolve()
        if (root / "scripts" / "brandkit").is_dir():
            return root

    for parent in [start] + list(start.parents):
        if (parent / ".claude-plugin").is_dir():
            return parent
    return start.parents[3]


ROOT = _root(Path(__file__).resolve())
sys.path.insert(0, str(ROOT / "scripts"))

from brandkit.cli import main  # noqa: E402


if __name__ == "__main__":
    raise SystemExit(main())
