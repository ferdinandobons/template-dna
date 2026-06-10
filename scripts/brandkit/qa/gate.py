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
    content_hash: str | None = None,
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
        content_hash: the sha of the canonical generate input (``to_dict()`` JSON),
            passed ONLY by ``cli.generate``. It is recorded into the visual manifest
            (so the L2 model can scope a verdict to the exact artifact) and gates the
            generate-time L2 short-circuit. ``cli.verify`` passes None, so the
            short-circuit can NEVER fire at verify - verify behaviour is unchanged.
    """
    # ``check_profile`` runs EXACTLY ONCE: the per-format checks call it internally,
    # so when one of them runs (target present) calling it here too would
    # double-validate. Run it standalone only when no format-specific check applies
    # (verify time / target None / unknown kind). The per-format checks also receive
    # the brand ``shell`` for the shell-vs-output structural diffs (formula
    # preservation, component survival) that no text scan can detect; ``shell`` is
    # None at verify time and those diffs no-op on a missing file, so this is safe.
    # Per-pass artifact-load memo (see checks_deterministic.load_memo): within
    # this block the shell/output are each opened at most once across the
    # independent checks below; the memo dies with the block, so no loaded
    # object or fact is ever cached across run_qa invocations (shell-frozen
    # sha semantics still read file bytes straight from disk).
    with checks_deterministic.load_memo():
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
        findings = findings + checks_deterministic.check_resolver_targets(
            shell, profile
        )
        # Shell-backed peer of resolver targets: every captured/applied font must be one
        # the shell actually carries (fail-closed). No-op when no appearance is present.
        findings = findings + checks_deterministic.check_appearance_targets(
            shell, profile
        )
        # Honest fail-closed peer for the paragraph-GEOMETRY axis (Cluster D1, docx-only):
        # every applied spacing/indent/border/shading value must be WELL-FORMED and a value
        # the template's OWN paragraphs carried (the captured floor), never synthesized.
        # No-op for non-docx kinds and when no geometry is captured (pre-D1 profiles).
        findings = findings + checks_deterministic.check_geometry_targets(
            shell, profile
        )
        # Honest fail-closed peer for the TABLE conditional-format axis (Cluster D2,
        # docx-only): every applied tblLook bitmask must be WELL-FORMED (shape/sanity), every
        # referenced table style must be one the shell's styles part DEFINES (symbolic
        # name-membership, like fonts), and every cell margin must be a value the template's
        # OWN tables carried (observed-floor, like geometry). The band fills stay in the shell
        # style part - the engine only enables them. No-op for non-docx kinds and when no
        # table appearance is captured (pre-D2 profiles).
        findings = findings + checks_deterministic.check_table_targets(shell, profile)
        # Honest fail-closed peer for the LIST / NUMBERING-definition axis (Cluster D3,
        # docx-only): every referenced num_id / abstract_num_id must be one the shell's
        # numbering part DEFINES (symbolic membership), every per-level numFmt must be a valid
        # OOXML field code (shape), and every per-level lvlText / indent must be byte-identical
        # to the shell's OWN abstractNum for that level (observed-floor, never synthesized). The
        # numbering definition stays the shell's - the engine only references/clones it by id.
        # No-op for non-docx kinds and when no numbering is captured (pre-D3 profiles).
        findings = findings + checks_deterministic.check_numbering_targets(
            shell, profile
        )
        # Fail-closed comprehension-target membership (sibling of resolver targets):
        # every load-bearing comprehension ref must be a verbatim id from the surfaced
        # inventories. No-ops when comprehension is absent (model-free CI path,
        # pptx/xlsx), so it is always safe to run unconditionally here.
        findings = findings + checks_deterministic.check_comprehension_targets(profile)
        # Fail-closed L2 visual-AUDIT-target membership (sibling of comprehension targets,
        # Cluster C1): every persisted ``comprehension.audit`` key must be a verbatim id
        # from the profile's derived visual checklist; rejects-never-skips on an empty
        # checklist. No-ops when comprehension is absent, so the model-free CI path and
        # pptx/xlsx are unaffected.
        findings = findings + checks_deterministic.check_audit_targets(profile)
        # Fail-closed model-assisted QA-TRIAGE-target membership (sibling of audit
        # targets, Cluster C2): every ``comprehension.triage`` entry must name a check in
        # the closed eligible set and a unique (check, location) pair; rejects-never-skips
        # a non-eligible check / duplicate. No-ops when comprehension is absent, so the
        # model-free CI path and pptx/xlsx are unaffected. (The triage entries are CONSUMED
        # - demoting a matched WARNING to INFO - inside ``_run_visual_audit`` via the single
        # ``_apply_triage``, BEFORE the strict promoter and this final fold; this line only
        # ENFORCES that every triage entry is well-formed.)
        findings = findings + checks_deterministic.check_triage_targets(profile)
        # Fail-closed LEARNED-override-target membership (sibling of comprehension
        # targets, Cluster B): every reroute target / number_format mask / demo-clear
        # value a learned lesson re-points to must be proven by this shell / captured for
        # this template. Rejects-never-skips on an empty inventory. No-ops when overrides
        # are absent (status != present), so the model-free CI path and pre-B3 profiles
        # are unaffected.
        findings = findings + checks_deterministic.check_override_targets(
            shell, profile
        )
        # Audit visibility (Cluster B4): surface every LIVE learned override as an INFO
        # ``override_applied`` finding so a learned re-point is never silent and verify
        # re-surfaces it. Gated on the SAME presence+freeze predicate the resolver consumes
        # on (``store.overrides_are_present``), so it emits iff an override branch is
        # actually live; INFO-only, not in ``DEFAULT_L0_INVARIANTS``, so it can never flip a
        # verdict. Reads only the profile, so it runs in both generate and verify.
        findings = findings + checks_deterministic.check_overrides_applied(profile)
        # Fail-closed COLOR-token membership (sibling of comprehension targets): every
        # palette color token the comprehension/IDoc references must be a verbatim key
        # of theme.palette. No-ops when comprehension is absent, so the model-free CI
        # path and pre-palette profiles are unaffected.
        findings = findings + checks_deterministic.check_color_token_targets(profile)
        # Fail-closed PALETTE-ALIAS integrity (sibling of color-token targets, Cluster E1):
        # every model-NAMED alias minted into theme.palette must be a legal dotted token
        # whose ref is a byte-copy of its captured source entry's ref (the engine never
        # authors a color). No-ops when comprehension is absent and on annotations with no
        # alias, so the model-free CI path, pre-palette profiles, pptx/xlsx, and the
        # no-alias byte-identity path are all unaffected.
        findings = findings + checks_deterministic.check_palette_alias_targets(profile)
        if extra_findings:
            findings = findings + list(extra_findings)

        # The saved shell is part of the verified brand contract. If it drifted from
        # the profile's recorded sha, verification/generation must fail closed.
        findings = findings + checks_deterministic.check_shell_provenance(
            shell, profile
        )

    findings = findings + _run_visual_audit(
        target,
        profile,
        qa=qa,
        out_dir=out_dir,
        visual=visual,
        shell=shell,
        content_hash=content_hash,
    )

    # Cluster C2: the SINGLE model-assisted triage demotion over the FULL assembled
    # list. This covers the format-level ``component_survival`` WARNINGs (built above,
    # not inside ``_run_visual_audit``) and is idempotent over the visual L1 findings
    # already triaged inside ``_run_visual_audit`` before the strict promoter (an
    # already-demoted INFO fails the WARNING guard, so no double-suffix). The fold
    # below then keys on this already-triaged list. An ERROR can NEVER be demoted -
    # ``_apply_triage`` guards on ``severity == WARNING``.
    findings = _apply_triage(findings, profile)

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
    shell=None,
    content_hash: str | None = None,
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

    # L2 SHORT-CIRCUIT (Cluster C1): skip the model-driven L2 render+manifest round
    # ONLY when a previously persisted audit verdict proves the EXACT same artifact
    # already passed. It can never mask a regression: it requires coverage of EVERY
    # current checklist id at verdict==PASS AND an exact (shell_sha, content_sha)
    # match, it is disabled under strict (which always re-renders for full proof) and
    # at verify (content_hash is None there). It skips the WHOLE visual round - the
    # render, the side artifacts, AND the L1 pixel proxies that live inside it (which
    # are WARNING-only and redundant on a byte-identical render) - but NEVER the
    # document bytes and NEVER the ERROR-bearing L0 checks above (those already ran).
    short = _l2_short_circuit(profile, qa=qa, shell=shell, content_hash=content_hash)
    if short is not None:
        return [short]

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
                shell_sha256=_manifest_shell_sha256(profile, shell),
                content_sha256=content_hash,
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
        # Cluster C2: triage demotion MUST precede the strict promotion, or a
        # model-confirmed-EXPECTED ambiguous WARNING (a full-bleed cover the
        # edge-bleed proxy flags, a deliberately blank page) would already be
        # promoted to a ``visual.strict`` ERROR and stay failing under strict. The
        # single ``_apply_triage`` runs here on the assembled L1/OCR findings; the
        # downstream strict promoter and ``run_qa``'s final fold then read the
        # already-triaged list. ``visual.strict`` is an ERROR id NOT in the eligible
        # set, so once raised it is itself undemotable.
        findings = _apply_triage(findings, profile)
        if qa == "strict":
            findings.extend(
                _strict_visual_errors(findings, _expected_triage_keys(profile))
            )
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
            shell_sha256=_manifest_shell_sha256(profile, shell),
            content_sha256=content_hash,
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


def _apply_triage(findings: list[Finding], profile: dict) -> list[Finding]:
    """Demote model-confirmed-EXPECTED ambiguous WARNINGs to INFO (Cluster C2).

    The SOLE consumer of ``comprehension.triage``. Returns a new list where a finding
    is demoted to INFO ONLY when ALL hold:

      * ``finding.severity == Severity.WARNING`` (an ERROR or INFO is left verbatim -
        this single guard is the load-bearing proof that an ERROR can NEVER be
        demoted: the enum has no value that lowers an ERROR);
      * ``finding.check`` is in the closed :data:`schema.AMBIGUOUS_TRIAGE_CHECKS`
        (defense-in-depth: the lookup is already restricted to that set);
      * ``(finding.check, finding.location)`` has a matching triage entry whose
        ``disposition == expected``.

    A ``defect`` disposition, an entry for an ERROR-emitting check (rejected at merge,
    never reaching the lookup), or any miss leaves the finding verbatim. No-op when
    the comprehension is absent / not present / carries no triage, so an empty triage
    is byte-identical to today's gate. Idempotent: a finding already demoted to INFO
    fails the WARNING guard, so a second pass is a no-op (this lets the single
    implementation be called both inside ``_run_visual_audit`` - before the strict
    promoter - and again over the full assembled list, with no double-suffix).
    """
    comp = profile.get("comprehension")
    if (
        not isinstance(comp, dict)
        or comp.get("status") != schema.ComprehensionStatus.PRESENT.value
    ):
        return findings
    triage = comp.get("triage")
    if not isinstance(triage, list) or not triage:
        return findings
    # Build the (check, location) -> entry lookup, restricted to the eligible set
    # (defense-in-depth; the merge already rejected a non-member check).
    lookup: dict[tuple, dict] = {}
    for entry in triage:
        if not isinstance(entry, dict):
            continue
        check = entry.get("check")
        if check not in schema.AMBIGUOUS_TRIAGE_CHECKS:
            continue
        lookup[(check, entry.get("location"))] = entry

    out: list[Finding] = []
    for finding in findings:
        if (
            finding.severity == schema.Severity.WARNING.value
            and finding.check in schema.AMBIGUOUS_TRIAGE_CHECKS
        ):
            entry = lookup.get((finding.check, finding.location))
            if (
                entry is not None
                and entry.get("disposition") == schema.TriageDisposition.EXPECTED.value
            ):
                evidence = entry.get("evidence")
                suffix = (
                    f" (triaged EXPECTED: {evidence})"
                    if evidence
                    else " (triaged EXPECTED)"
                )
                out.append(
                    Finding(
                        finding.check,
                        schema.Severity.INFO.value,
                        finding.message + suffix,
                        location=finding.location,
                    )
                )
                continue
        out.append(finding)
    return out


def _manifest_shell_sha256(profile: dict, shell) -> str | None:
    """The shell sha stamped into the manifest / matched by the short-circuit.

    Single derivation shared by both manifest writes AND :func:`_l2_short_circuit`
    (via ``qa.report.report_shell_sha256``), so the per-row ``shell_sha256`` the
    model wrote against a manifest can be compared bit-for-bit to the sha a later
    generate derives. Prefers the live shell bytes, else ``provenance.shell.sha256``.
    """
    from brandkit.qa import report as vreport

    return vreport.report_shell_sha256(profile, shell)


def _l2_short_circuit(
    profile: dict, *, qa: str, shell, content_hash: str | None
) -> Finding | None:
    """Return a single INFO finding when the L2 round may be SKIPPED, else None (C1).

    Fires ONLY when ALL hold (so it can never mask a regression):

      1. ``qa != "strict"`` - strict ALWAYS re-renders for full proof.
      2. ``content_hash is not None`` - generate-time only; verify passes None.
      3. the present comprehension ``audit`` map covers EVERY current
         ``visual_checklist_ids(profile)`` id (coverage of the CURRENTLY derived set,
         not just the keys present - a newly-derived id with no audit row forces a
         full L2).
      4. for every such id: ``verdict == "PASS"`` AND ``shell_sha256`` matches the
         live shell sha AND ``content_sha256 == content_hash``.

    Any missing id / any FAIL / any NA / any sha mismatch -> return None (fall
    through to the full render+manifest path). When it fires it skips the WHOLE
    visual round - the render, the side artifacts (PNGs+manifest), AND the L1 pixel
    proxies (blank-page/edge-bleed/OCR), which run inside that round. That is sound:
    L1 reads the rendered pixels, and an identical (shell_sha, content_sha) yields a
    byte-identical render, so re-running L1 is redundant; ANY content/shell change
    defeats the sha gate and forces the full round. Only L0 (run BEFORE the visual
    round) is guaranteed to have already executed.
    """
    if qa == "strict" or content_hash is None:
        return None
    comp = profile.get("comprehension")
    if (
        not isinstance(comp, dict)
        or comp.get("status") != schema.ComprehensionStatus.PRESENT.value
    ):
        return None
    audit = comp.get("audit")
    if not isinstance(audit, dict) or not audit:
        return None

    from brandkit.qa.visual import visual_checklist_ids

    checklist_ids = visual_checklist_ids(profile)
    if not checklist_ids:
        # No derived checklist => nothing the verdict could cover => never skip.
        return None

    shell_sha = _manifest_shell_sha256(profile, shell)
    for cid in checklist_ids:
        row = audit.get(cid)
        if not isinstance(row, dict):
            return None  # a current checklist id with no verdict -> full L2
        if row.get("verdict") != schema.AuditVerdict.PASS.value:
            return None  # any FAIL/NA/missing -> full L2
        if row.get("shell_sha256") != shell_sha:
            return None  # stale shell -> full L2
        if row.get("content_sha256") != content_hash:
            return None  # different content -> full L2

    return Finding(
        "visual.l2_short_circuit",
        schema.Severity.INFO.value,
        (
            f"L2 visual audit short-circuited: all {len(checklist_ids)} checklist "
            "item(s) PASSed for this exact shell+content in a prior audit; "
            "render+manifest skipped"
        ),
        location=None,
    )


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


def _expected_triage_keys(profile: dict) -> set[tuple]:
    """The set of (check, location) pairs a present comprehension triaged EXPECTED.

    Restricted to the closed eligible set (defense-in-depth; merge already rejected a
    non-member). Empty when comprehension is absent / not present / carries no triage,
    so it is byte-identical to today's gate. Shared by :func:`_strict_visual_errors`
    so the strict promoter skips EXACTLY the WARNINGs ``_apply_triage`` demoted - and
    nothing else (a genuinely-INFO blocking finding like ``visual.ocr_degraded`` is
    still promoted, since it is not an EXPECTED-triaged key).
    """
    comp = profile.get("comprehension")
    if (
        not isinstance(comp, dict)
        or comp.get("status") != schema.ComprehensionStatus.PRESENT.value
    ):
        return set()
    triage = comp.get("triage")
    if not isinstance(triage, list):
        return set()
    keys: set[tuple] = set()
    for entry in triage:
        if (
            isinstance(entry, dict)
            and entry.get("check") in schema.AMBIGUOUS_TRIAGE_CHECKS
            and entry.get("disposition") == schema.TriageDisposition.EXPECTED.value
        ):
            keys.add((entry.get("check"), entry.get("location")))
    return keys


def _strict_visual_errors(
    findings: list[Finding], expected_triage: set[tuple] | None = None
) -> list[Finding]:
    """Promote concrete visual-audit findings into strict-mode gate errors.

    ``expected_triage`` is the set of ``(check, location)`` pairs a model confirmed
    EXPECTED (Cluster C2). A blocking finding matching one of those pairs is NOT
    promoted: triage already demoted that WARNING to INFO, and strict must honour the
    model's judgement rather than re-promote the same defect. This is the keyed
    counterpart to the ``_apply_triage`` demotion (same closed lookup), so it skips
    EXACTLY the demoted findings and never a genuinely-INFO blocking finding (e.g.
    ``visual.ocr_degraded``, which is INFO-by-design and still blocks strict). With no
    triage the set is empty and every existing strict promotion is unchanged.
    """
    expected_triage = expected_triage or set()
    errors: list[Finding] = []
    for finding in findings:
        if finding.check not in _STRICT_BLOCKING_CHECKS:
            continue
        if (finding.check, finding.location) in expected_triage:
            continue  # model-confirmed EXPECTED -> undemotable AND unpromotable
        errors.append(
            Finding(
                "visual.strict",
                schema.Severity.ERROR.value,
                f"strict visual QA blocks on {finding.check}: {finding.message}",
                location=finding.location,
            )
        )
    return errors
