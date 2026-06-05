# SPDX-License-Identifier: MIT
"""Adaptive QA gate entrypoint.

M1 implements L0 deterministic checks and treats visual QA as gracefully
unavailable unless external renderers are installed.
"""
from __future__ import annotations

from pathlib import Path

from brandkit.profile import schema
from brandkit.qa import checks_deterministic
from brandkit.qa.model import QAReport


def run_qa(target: str | Path | None, profile: dict, *, mode: str = "generate", qa: str = "auto") -> QAReport:
    findings = checks_deterministic.check_profile(profile)
    if target is not None and profile.get("kind") == "docx":
        findings = checks_deterministic.check_docx(target, profile)
    elif target is not None and profile.get("kind") == "pptx":
        findings = checks_deterministic.check_pptx(target, profile)
    elif target is not None and profile.get("kind") == "xlsx":
        findings = checks_deterministic.check_xlsx(target, profile)
    verdict = schema.VerificationStatus.PASSED.value
    if any(f.severity == schema.Severity.ERROR.value for f in findings):
        verdict = schema.VerificationStatus.FAILED.value
    elif findings:
        verdict = schema.VerificationStatus.PASSED_WITH_WARNINGS.value
    return QAReport(verdict=verdict, findings=findings)
