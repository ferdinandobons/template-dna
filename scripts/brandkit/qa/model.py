# SPDX-License-Identifier: MIT
from __future__ import annotations

from dataclasses import dataclass, field

from brandkit.profile import schema


@dataclass
class Finding:
    check: str
    severity: str
    message: str
    location: str | None = None


@dataclass
class QAReport:
    verdict: str
    findings: list[Finding] = field(default_factory=list)

    @property
    def passed(self) -> bool:
        return not any(f.severity == schema.Severity.ERROR.value for f in self.findings)

