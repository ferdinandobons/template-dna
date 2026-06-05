# SPDX-License-Identifier: MIT
"""Adaptive QA gate entrypoint.

L0 implements deterministic byte/structure checks. On top of it a two-stage
visual audit runs *only* when the external renderers are installed and the QA
mode asks for it (``auto``/``deep``):

  * **L1** = deterministic pixel proxies on the rendered PNGs (blank pages, edge
    bleed); each defect is a WARNING ``Finding`` (never ERROR -> never changes a
    passing verdict).
  * **L2** = a ``visual_manifest.json`` (PNG paths + a profile-derived checklist)
    the orchestrator reads to make the qualitative judgement and drive a repair
    loop. The engine never calls a model.

Backward-compat: ``fast`` and ``target is None`` never touch the visual path;
when renderers are absent (CI) ``auto``/``deep`` degrade to L0 plus a single INFO
``visual.unavailable`` finding. The PNGs and manifest are SIDE artifacts in an
out dir next to the output; the generated document's bytes never change.
"""
from __future__ import annotations

from pathlib import Path

from brandkit.profile import schema
from brandkit.qa import checks_deterministic
from brandkit.qa.model import Finding, QAReport


def run_qa(
    target: str | Path | None,
    profile: dict,
    *,
    mode: str = "generate",
    qa: str = "auto",
    shell: str | Path | None = None,
    extra_findings: list | None = None,
    out_dir: str | Path | None = None,
    visual=None,
) -> QAReport:
    """Run the L0 deterministic gate, then the visual audit when asked/available.

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
        out_dir: where to write the visual side artifacts (PNGs + manifest). When
            None a conventional ``<output>.visual`` dir next to ``target`` is used.
            Only consulted on the ``auto``/``deep`` path with renderers present.
        visual: a pre-rendered ``(renderers_ok, png_paths)`` tuple injected by
            tests to drive the ``auto``/``deep`` path without ``soffice``. When
            None the real env-aware renderer is used. Default None keeps every
            existing call identical.
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

    # The saved shell is part of the verified brand contract. If it drifted from
    # the profile's recorded sha, verification/generation must fail closed.
    findings = findings + checks_deterministic.check_shell_provenance(shell, profile)

    findings = findings + _run_visual_audit(
        target, profile, qa=qa, out_dir=out_dir, visual=visual,
    )

    verdict = schema.VerificationStatus.PASSED.value
    if any(f.severity == schema.Severity.ERROR.value for f in findings):
        verdict = schema.VerificationStatus.FAILED.value
    elif findings:
        verdict = schema.VerificationStatus.PASSED_WITH_WARNINGS.value
    return QAReport(verdict=verdict, findings=findings)


def _run_visual_audit(
    target,
    profile: dict,
    *,
    qa: str,
    out_dir,
    visual,
) -> list[Finding]:
    """Run the optional visual audit and return its findings (never raises).

    Semantics:
      * ``fast`` or ``target is None`` -> no visual path at all (``[]``). This is
        the pillar of backward-compat: every smoke test uses ``--qa fast`` and so
        never imports or touches the visual module.
      * ``auto``/``deep`` with renderers ABSENT -> a single INFO
        ``visual.unavailable`` (clean degrade to L0; no ERROR, no verdict change).
      * ``auto`` with renderers present -> L0 + L1 pixel proxies.
      * ``deep`` with renderers present -> L0 + L1 + a written manifest, signalled
        back via an INFO ``visual.manifest`` carrying the manifest path so the
        orchestrator can run the L2 step.

    The visual module is imported lazily so the fast/CI path never imports it and
    an import-time environment problem cannot touch the L0 path.
    """
    qa = (qa or "auto").lower()
    if qa == "fast" or target is None:
        return []

    # Lazy import: keeps the fast/CI path free of the visual module entirely.
    from brandkit.qa import visual as vqa

    if visual is not None:
        renderers_ok, png_paths = visual
    else:
        renderers_ok = vqa.renderers_available()
        png_paths = []

    if not renderers_ok:
        return [Finding(
            "visual.unavailable",
            schema.Severity.INFO.value,
            "visual QA unavailable (soffice/pdftoppm absent); L0 only",
        )]

    resolved_out = Path(out_dir) if out_dir is not None else vqa.default_out_dir(target)
    if visual is None:
        png_paths = vqa.render_to_pngs(target, resolved_out)

    findings: list[Finding] = list(vqa.run_visual_l1(png_paths))
    findings.extend(vqa.check_page_count_sane(png_paths))

    if qa == "deep":
        manifest = vqa.build_visual_manifest(
            profile=profile,
            document=target,
            png_paths=png_paths,
            l1_findings=findings,
            renderers_ok=True,
            out_dir=resolved_out,
        )
        findings.append(Finding(
            "visual.manifest",
            schema.Severity.INFO.value,
            f"visual audit manifest written: {manifest}",
            location=str(manifest),
        ))
    return findings
