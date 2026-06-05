# SPDX-License-Identifier: MIT
"""Guard against proprietary / vendor leaks in the tracked repository.

Two invariants, both scoped to *git-tracked* files (so gitignored scratch such as
``brand-kit/`` and ``generated/`` is ignored):

1. No Office binary (``.docx``/``.pptx``/``.xlsx``/legacy) is tracked anywhere
   except ``tests/fixtures/`` — company templates and generated samples must never
   be committed, regardless of filename.
2. No tracked source imports Bedrock/boto3 or a vendored proprietary Office helper
   package (``office.*``) — the engine is self-contained.
"""
from __future__ import annotations

import subprocess
import unittest
from pathlib import Path

FORBIDDEN_RUNTIME_PATTERNS = (
    "import boto3",
    "from boto3",
    "bedrock-runtime",
    "from office.",
    "import office.",
)
OFFICE_SUFFIXES = {".docx", ".pptx", ".xlsx", ".doc", ".ppt", ".xls"}
TEXT_SUFFIXES = {".py", ".md", ".json", ".txt", ".svg", ".yml", ".yaml", ""}
FIXTURES = ("tests", "fixtures")


def _tracked_files(root: Path) -> list[Path]:
    out = subprocess.check_output(["git", "ls-files"], cwd=str(root), text=True)
    return [root / line for line in out.splitlines() if line]


class NoProprietaryTest(unittest.TestCase):
    def test_no_proprietary_or_bedrock_leaks(self) -> None:
        root = Path(__file__).resolve().parents[1]
        self_path = Path(__file__).resolve()
        offenders: list[str] = []
        for path in _tracked_files(root):
            if not path.is_file():
                continue
            rel = path.relative_to(root)
            suffix = path.suffix.lower()
            if suffix in OFFICE_SUFFIXES and tuple(rel.parts[:2]) != FIXTURES:
                offenders.append(f"{rel}: tracked Office asset outside tests/fixtures")
                continue
            if path == self_path or suffix not in TEXT_SUFFIXES:
                continue
            lowered = path.read_text(encoding="utf-8", errors="ignore").lower()
            for token in FORBIDDEN_RUNTIME_PATTERNS:
                if token in lowered:
                    offenders.append(f"{rel}: contains {token!r}")
        self.assertEqual(offenders, [])


if __name__ == "__main__":
    unittest.main()
