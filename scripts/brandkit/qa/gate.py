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


def run_qa(
    target: str | Path | None,
    profile: dict,
    *,
    mode: str = "generate",
    qa: str = "auto",
    shell: str | Path | None = None,
    extra_findings: list | None = None,
) -> QAReport:
    """Run the L0 deterministic gate.

    Args:
        target: the produced artifact to scan for leaked markdown / residual text
            (None at verify time, when there is no output yet).
        profile: the brand profile under test.
        shell: the brand shell (``template/shell.<ext>``). When given, the
            ``resolver_targets_exist`` invariant opens it and verifies every role's
            target actually exists - the deterministic backstop for the core
            promise. Passed at verify time so the shell is checked even without a
            generated output.
        extra_findings: findings recorded elsewhere (e.g. by the generator's loud
            style lookup) to fold into this report.
    """
    findings = checks_deterministic.check_profile(profile)
    # The per-format checks also receive the brand ``shell`` so they can run the
    # shell-vs-output structural diffs (formula preservation, component survival)
    # that no text scan can detect. ``shell`` is None at verify time (no output to
    # diff), and those diffs no-op on a missing file, so this is always safe.
    if target is not None and profile.get("kind") == "docx":
        findings = checks_deterministic.check_docx(target, profile, shell=shell)
    elif target is not None and profile.get("kind") == "pptx":
        findings = checks_deterministic.check_pptx(target, profile, shell=shell)
    elif target is not None and profile.get("kind") == "xlsx":
        findings = checks_deterministic.check_xlsx(target, profile, shell=shell)

    # Deterministic resolver-target existence check (opens the shell once).
    findings = findings + checks_deterministic.check_resolver_targets(shell, profile)
    # Fail-closed comprehension-target membership (sibling of resolver targets):
    # every load-bearing comprehension ref must be a verbatim id from the surfaced
    # inventories. No-ops when comprehension is absent (model-free CI path,
    # pptx/xlsx), so it is always safe to run unconditionally here.
    findings = findings + checks_deterministic.check_comprehension_targets(profile)
    if extra_findings:
        findings = findings + list(extra_findings)

    verdict = schema.VerificationStatus.PASSED.value
    if any(f.severity == schema.Severity.ERROR.value for f in findings):
        verdict = schema.VerificationStatus.FAILED.value
    elif findings:
        verdict = schema.VerificationStatus.PASSED_WITH_WARNINGS.value
    return QAReport(verdict=verdict, findings=findings)
