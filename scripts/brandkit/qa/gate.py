# SPDX-License-Identifier: MIT
"""Adaptive QA gate entrypoint.

L0 implements deterministic byte/structure checks. On top of it a two-stage
visual audit runs *only* when the external renderers are installed and the QA
mode asks for it (``auto``/``deep``/``strict``):

  * **L1** = deterministic pixel proxies on the rendered PNGs (blank pages, edge
    bleed); each defect is a WARNING ``Finding`` (never ERROR -> never changes a
    passing verdict).
  * **L2** = a ``visual_manifest.json`` (PNG paths + a profile-derived checklist)
    the orchestrator reads to make the qualitative judgement and drive a repair
    loop. The engine never calls a model.

Backward-compat: ``fast`` and ``target is None`` never touch the visual path;
when renderers are absent (CI) ``auto`` degrades to L0 plus a single INFO
``visual.unavailable`` finding; ``deep`` also writes a degraded manifest so the
orchestrator still has the L2 checklist. ``strict`` writes the same manifest but
fails when full visual proof is unavailable or L1/OCR findings are present. The
PNGs and manifest are SIDE artifacts in an out dir next to the output; the
generated document's bytes never change.
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
            None a conventional ``<output-filename>.visual`` dir next to
            ``target`` is used (for example ``out.docx.visual``).
            Only consulted on the ``auto``/``deep``/``strict`` path with renderers
            present.
        visual: a pre-rendered ``(renderers_ok, png_paths)`` tuple injected by
            tests to drive the visual path without ``soffice``. When None the
            real env-aware renderer is used. Default None keeps every existing
            call identical.
    """
    # ``check_profile`` runs EXACTLY ONCE: the per-format checks call it internally,
    # so when one of them runs (target present) calling it here too would
    # double-validate. Run it standalone only when no format-specific check applies
    # (verify time / target None / unknown kind). The per-format checks also receive
    # the brand ``shell`` for the shell-vs-output structural diffs (formula
    # preservation, component survival) that no text scan can detect; ``shell`` is
    # None at verify time and those diffs no-op on a missing file, so this is safe.
    if target is not None and profile.get("kind") == "docx":
        findings = checks_deterministic.check_docx(target, profile, shell=shell)
    elif target is not None and profile.get("kind") == "pptx":
        findings = checks_deterministic.check_pptx(target, profile, shell=shell)
    elif target is not None and profile.get("kind") == "xlsx":
        findings = checks_deterministic.check_xlsx(target, profile, shell=shell)
    else:
        findings = checks_deterministic.check_profile(profile)

    # Format-agnostic OPC integrity backstop: a generated package with duplicate
    # ZIP part names is corrupt (Office repair dialog). No-ops at verify time / on a
    # missing file.
    if target is not None:
        findings = findings + checks_deterministic.check_no_duplicate_parts(target)

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
        target,
        profile,
        qa=qa,
        out_dir=out_dir,
        visual=visual,
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
      * ``auto``/``deep``/``strict`` with renderers ABSENT -> INFO
        ``visual.unavailable``; ``deep`` additionally writes a degraded manifest
        with the checklist; ``strict`` writes the manifest and adds an ERROR.
      * ``auto`` with renderers present -> L0 + L1 pixel proxies.
      * ``deep`` with renderers present -> L0 + L1 + a written manifest, signalled
        back via an INFO ``visual.manifest`` carrying the manifest path so the
        orchestrator can run the L2 step.
      * ``strict`` with renderers present -> ``deep`` plus ERROR findings when
        render/L1/OCR evidence is not clean.

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

    resolved_out = Path(out_dir) if out_dir is not None else vqa.default_out_dir(target)

    if not renderers_ok:
        findings = [
            Finding(
                "visual.unavailable",
                schema.Severity.INFO.value,
                "visual QA unavailable (soffice/pdftoppm absent); L0 only",
            )
        ]
        if qa in ("deep", "strict"):
            render_errors: list[str] = []
            render_warnings: list[str] = []
            if qa == "deep":
                png_paths = vqa.render_to_pngs(
                    target,
                    resolved_out,
                    check_available=False,
                    quicklook_only=True,
                    render_errors=render_errors,
                    render_warnings=render_warnings,
                )
                findings.extend(vqa.run_visual_l1(png_paths))
                for warning in render_warnings:
                    findings.append(
                        Finding(
                            "visual.render_degraded",
                            schema.Severity.WARNING.value,
                            warning,
                        )
                    )
                if not png_paths:
                    detail = (
                        render_errors[-1]
                        if render_errors
                        else "renderer produced no pages"
                    )
                    findings.append(
                        Finding(
                            "visual.render_failed",
                            schema.Severity.WARNING.value,
                            f"visual fallback render failed: {detail}",
                        )
                    )
                    # ``visual.render_failed`` already signals zero pages here; do
                    # NOT also run check_page_count_sane (it would add a redundant
                    # ``visual.no_pages`` WARNING for the same root cause).
            else:
                findings.append(
                    Finding(
                        "visual.strict_unavailable",
                        schema.Severity.ERROR.value,
                        "strict visual QA requires full render proof, but renderers are unavailable",
                    )
                )
            ocr_report = vqa.run_visual_ocr(png_paths, profile)
            findings.extend(vqa.ocr_findings(ocr_report))
            manifest = vqa.build_visual_manifest(
                profile=profile,
                document=target,
                png_paths=png_paths,
                l1_findings=findings,
                renderers_ok=False,
                out_dir=resolved_out,
                degraded=True,
                environment_status=vqa.last_renderer_status(),
                ocr_report=ocr_report,
                qa_mode=qa,
            )
            findings.append(
                Finding(
                    "visual.manifest",
                    schema.Severity.INFO.value,
                    f"degraded visual audit manifest written: {manifest}",
                    location=str(manifest),
                )
            )
        return findings

    render_errors: list[str] = []
    render_warnings: list[str] = []
    if visual is None:
        png_paths = vqa.render_to_pngs(
            target,
            resolved_out,
            check_available=False,
            render_errors=render_errors,
            render_warnings=render_warnings,
        )

    findings: list[Finding] = list(vqa.run_visual_l1(png_paths))
    if visual is None and render_warnings:
        for warning in render_warnings:
            findings.append(
                Finding(
                    "visual.render_degraded",
                    schema.Severity.WARNING.value,
                    warning,
                )
            )
    if visual is None and not png_paths:
        detail = render_errors[-1] if render_errors else "renderer produced no pages"
        findings.append(
            Finding(
                "visual.render_failed",
                schema.Severity.WARNING.value,
                f"visual render failed after renderer probe: {detail}",
            )
        )
    findings.extend(vqa.check_page_count_sane(png_paths))

    if qa in ("deep", "strict"):
        manifest_renderers_ok = (
            renderers_ok
            if visual is not None
            else (bool(png_paths) and not _degraded_render_warnings(render_warnings))
        )
        ocr_report = vqa.run_visual_ocr(png_paths, profile)
        findings.extend(vqa.ocr_findings(ocr_report))
        if qa == "strict":
            findings.extend(_strict_visual_errors(findings))
        manifest = vqa.build_visual_manifest(
            profile=profile,
            document=target,
            png_paths=png_paths,
            l1_findings=findings,
            renderers_ok=manifest_renderers_ok,
            out_dir=resolved_out,
            degraded=bool(render_warnings) or not manifest_renderers_ok,
            environment_status=vqa.last_renderer_status(),
            ocr_report=ocr_report,
            qa_mode=qa,
        )
        findings.append(
            Finding(
                "visual.manifest",
                schema.Severity.INFO.value,
                f"visual audit manifest written: {manifest}",
                location=str(manifest),
            )
        )
    return findings


def _degraded_render_warnings(warnings: list[str]) -> bool:
    """Return True when PNGs came from a partial, non-PDF visual fallback."""
    return any("Quick Look" in warning for warning in warnings)


_STRICT_BLOCKING_CHECKS = {
    "visual.blank_page",
    "visual.edge_bleed",
    "visual.no_pages",
    "visual.render_degraded",
    "visual.render_failed",
    "visual.ocr_residual_text",
    "visual.ocr_degraded",
}


def _strict_visual_errors(findings: list[Finding]) -> list[Finding]:
    """Promote concrete visual-audit findings into strict-mode gate errors."""
    errors: list[Finding] = []
    for finding in findings:
        if finding.check not in _STRICT_BLOCKING_CHECKS:
            continue
        errors.append(
            Finding(
                "visual.strict",
                schema.Severity.ERROR.value,
                f"strict visual QA blocks on {finding.check}: {finding.message}",
                location=finding.location,
            )
        )
    return errors
