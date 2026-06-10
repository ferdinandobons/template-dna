# SPDX-License-Identifier: MIT
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

# The six per-format extract/generate modules are imported INSIDE the command
# branches that dispatch on verb+format: only the active format's Office lib is
# paid for per invocation (pure import deferral - the import still happens
# before any use, so behavior and findings are unchanged).
from brandkit import doctor
from brandkit.grid.model import parse_grid
from brandkit.ir.model import parse_idoc
from brandkit.profile import comprehension as comprehension_mod
from brandkit.profile import schema, store
from brandkit.qa.gate import run_qa


def _content_hash(canonical: dict) -> str:
    """Hash a canonical input dict deterministically.

    ``sort_keys`` + ``ensure_ascii=False`` make logically-equal inputs (differing
    only in author-side key order / whitespace) hash identically -- the stable
    "same input" key the cross-run report (Cluster B) reads.
    """
    return store.sha256_bytes(
        json.dumps(canonical, sort_keys=True, ensure_ascii=False).encode("utf-8")
    )


def _grid_to_dict(grid) -> dict:
    """Canonical dict form of a parsed ``GridDocument`` (peer of ``idoc.to_dict()``).

    The grid model carries no ``to_dict``; this projects its load-bearing fields
    into a stable shape for the content hash. Empty containers are omitted so the
    hash matches what an equivalent minimal input would produce.
    """
    out: dict = {"cells": dict(grid.cells), "regions": dict(grid.regions)}
    if grid.formats:
        out["formats"] = dict(grid.formats)
    if grid.charts:
        out["charts"] = [dict(c) for c in grid.charts]
    return out


def _discover_learn_reports(profile: dict, shell_path) -> list[dict]:
    """Discover this profile's SAME-shell generation_report.json history for ``learn``.

    The deterministic ``learn`` verb distils a lesson from the cross-run report
    history B1/B2 persist. Each ``generate`` writes its report into the
    ``<output>.visual`` dir NEXT TO its output (an operator-chosen path), so there is
    no single registry of past outputs; we walk the current working directory tree
    for every ``*.visual/generation_report.json`` and keep ONLY those whose
    ``shell_sha256`` matches THIS profile's live shell. The sha partition both scopes
    the history to this profile (a different profile has a different shell sha) and
    enforces SHELL-FROZEN (a re-extract re-stamps the sha, starting a fresh history).

    A pure side artifact: any IO error / malformed report degrades to fewer (or zero)
    priors, never an exception - mirroring ``report.discover_prior_reports``.
    """
    from brandkit.qa import report as vreport

    shell_sha256 = vreport.report_shell_sha256(profile, shell_path)
    if not shell_sha256:
        return []
    try:
        # Only the candidate set is ours (a cwd-wide walk); the per-candidate
        # read/filter/order pipeline is the SAME one discover_prior_reports uses
        # (report.load_same_shell_reports), so the two read paths can never drift.
        candidates = sorted(Path.cwd().rglob(f"*.visual/{vreport.REPORT_FILENAME}"))
        return vreport.load_same_shell_reports(candidates, shell_sha256=shell_sha256)
    except Exception:
        return []


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="brand-docs")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("extract")
    p.add_argument("--name", required=True)
    p.add_argument("--template", required=True)
    p.add_argument("--scope", default="auto", choices=("auto", "project", "global"))

    p = sub.add_parser("verify")
    p.add_argument("--name", required=True)
    p.add_argument("--scope", default="auto", choices=("auto", "project", "global"))
    p.add_argument("--qa", default="auto", choices=("auto", "fast", "deep", "strict"))
    p.add_argument("--accept", action="store_true")

    p = sub.add_parser("generate")
    p.add_argument("--name", required=True)
    p.add_argument("--input", required=True)
    p.add_argument("--output", required=True)
    p.add_argument("--scope", default="auto", choices=("auto", "project", "global"))
    p.add_argument("--qa", default="auto", choices=("auto", "fast", "deep", "strict"))

    # comprehend-input: print the bounded {facts, excerpt} bundle the MODEL reads.
    p = sub.add_parser("comprehend-input")
    p.add_argument("--name", required=True)
    p.add_argument("--scope", default="auto", choices=("auto", "project", "global"))

    # comprehend: merge+validate+cache a model-authored comprehension.json. This is
    # the ONLY writer of the comprehension block.
    p = sub.add_parser("comprehend")
    p.add_argument("--name", required=True)
    p.add_argument(
        "--input", required=True, help="path to the model-authored comprehension.json"
    )
    p.add_argument("--scope", default="auto", choices=("auto", "project", "global"))

    # refine: overlay a model-authored qualitative-feedback delta onto the EXISTING
    # present comprehension, then route the WHOLE block back through merge (the single
    # fail-closed writer). ADVISORY by default (mirrors learn --accept): without
    # --accept the post-overlay diff is printed and the prior block stays authoritative
    # on disk; --accept persists the refined block. No schema change (1.2.0).
    p = sub.add_parser("refine")
    p.add_argument("--name", required=True)
    p.add_argument(
        "--input", required=True, help="path to the model-authored refinement.json"
    )
    p.add_argument("--scope", default="auto", choices=("auto", "project", "global"))
    p.add_argument(
        "--accept",
        action="store_true",
        help="persist the refined comprehension (else print the diff and keep the prior block)",
    )

    # learn: deterministically distil recurring QA findings (the cross-run
    # generation_report.json history, B1/B2) into a brand-safe overrides lesson and
    # cache it via the single merge_overrides sink. ADVISORY by default: the lesson is
    # written but kept OUT of the live resolver (status forced 'absent', zero new
    # resolver branches, byte-identical) until an explicit --accept promotes it to
    # 'present', so a single noisy run can never mint a permanent live lesson.
    p = sub.add_parser("learn")
    p.add_argument("--name", required=True)
    p.add_argument("--scope", default="auto", choices=("auto", "project", "global"))
    p.add_argument(
        "--accept",
        action="store_true",
        help="promote the distilled lesson to a LIVE override the resolver consumes",
    )

    # propose-overrides: the MODEL-assisted sibling of ``learn`` (B4). The model reads
    # the comprehend-input bundle's ``generation_history`` (the AMBIGUOUS recurring
    # remainder deterministic ``learn`` could not bind) and authors an overrides
    # proposal; this OVERLAYS it onto any existing lesson (overlay_overrides) and routes
    # the WHOLE block through the SINGLE merge_overrides sink (shape + fail-closed
    # membership + acyclicity). ADVISORY by default like ``learn``: written but kept OUT
    # of the live resolver (status forced 'absent', byte-identical) until --accept
    # promotes it to 'present'. No schema change (1.2.0).
    p = sub.add_parser("propose-overrides")
    p.add_argument("--name", required=True)
    p.add_argument(
        "--input",
        required=True,
        help="path to the model-authored overrides proposal JSON",
    )
    p.add_argument("--scope", default="auto", choices=("auto", "project", "global"))
    p.add_argument(
        "--accept",
        action="store_true",
        help="promote the proposed correction to a LIVE override the resolver consumes",
    )

    p = sub.add_parser("list")
    p.add_argument("--scope", default="auto", choices=("auto", "project", "global"))

    # Read-only cross-template drift report (REFLECTIONS P3): compares the
    # BRAND-level facts of two saved profiles (theme colors, fonts, semantic
    # palette roles, off-theme hex usage). Writes nothing; exit 1 on drift so
    # the verb is scriptable as a brand-coherence gate.
    p = sub.add_parser("compare-profiles")
    p.add_argument("--name-a", required=True, help="first saved profile name")
    p.add_argument("--name-b", required=True, help="second saved profile name")
    p.add_argument("--scope-a", default="auto", choices=("auto", "project", "global"))
    p.add_argument("--scope-b", default="auto", choices=("auto", "project", "global"))
    p.add_argument(
        "--json",
        action="store_true",
        help="print the structured comparison as JSON instead of the report",
    )

    p = sub.add_parser("doctor")
    p.add_argument(
        "--json",
        action="store_true",
        help="print the verbatim probe() dict as JSON and skip the human report",
    )
    p.add_argument(
        "--fast",
        action="store_true",
        help="skip the slow soffice render probes; mark visual QA as not probed",
    )
    args = parser.parse_args(argv)

    if args.cmd == "doctor":
        status = doctor.probe(skip_visual_pipeline=args.fast)
        if args.json:
            print(json.dumps(status, indent=2, sort_keys=True))
        else:
            doctor.print_report(status)
        return 0 if doctor.required_ok(status) else 1
    if args.cmd == "extract":
        path = Path(args.template)
        suffix = path.suffix.lower()
        # A malformed/unreadable template must fail with a clean "ERROR extract: ..."
        # and exit 1 (mirroring the generate command), not an unhandled traceback.
        # The bad-suffix SystemExit is a BaseException, so it is NOT swallowed here.
        try:
            if suffix == ".docx":
                from brandkit.formats.docx import extract as docx_extract

                profile_json = docx_extract.extract(path, args.name, scope=args.scope)
            elif suffix == ".pptx":
                from brandkit.formats.pptx import extract as pptx_extract

                profile_json = pptx_extract.extract(path, args.name, scope=args.scope)
            elif suffix == ".xlsx":
                from brandkit.formats.xlsx import extract as xlsx_extract

                profile_json = xlsx_extract.extract(path, args.name, scope=args.scope)
            else:
                raise SystemExit("supported templates: .docx, .pptx, .xlsx")
        except Exception as exc:
            print(f"ERROR extract: {exc}")
            return 1
        loaded = store.load_profile(args.name, args.scope)
        report = run_qa(None, loaded.profile, qa="fast", shell=loaded.shell_path)
        loaded.profile["verification"]["status"] = report.verdict
        loaded.profile["verification"]["roles_total"] = len(
            schema.list_role_ids(loaded.profile)
        )
        loaded.profile["verification"]["roles_verified"] = loaded.profile[
            "verification"
        ]["roles_total"]
        store.write_profile_json(loaded.directory, loaded.profile)
        print(f"extracted {args.name} -> {profile_json}")
        return 0 if report.passed else 1
    if args.cmd == "verify":
        loaded = store.load_profile(args.name, args.scope)
        report = run_qa(None, loaded.profile, qa=args.qa, shell=loaded.shell_path)
        loaded.profile["verification"]["status"] = report.verdict
        if args.accept and report.passed:
            loaded.profile.setdefault("verification", {})["accepted"] = True
        store.write_profile_json(loaded.directory, loaded.profile)
        for finding in report.findings:
            print(f"{finding.severity} {finding.check}: {finding.message}")
        print(f"verification: {report.verdict}")
        return 0 if report.passed else 1
    if args.cmd == "generate":
        loaded = store.load_profile(args.name, args.scope)
        data = json.loads(Path(args.input).read_text(encoding="utf-8"))
        gen_findings: list = []
        # The content hash is taken from the CANONICAL parsed input (``to_dict()``),
        # never the raw author bytes: cosmetic key-order/whitespace changes hash
        # equal, which is what B2 keys "same input" on. ``None`` when the parse
        # produced no canonical form (never blocks generation).
        content_hash: str | None = None
        try:
            if loaded.kind == "docx":
                from brandkit.formats.docx import generate as docx_generate

                idoc = parse_idoc(data)
                content_hash = _content_hash(idoc.to_dict())
                out = docx_generate.generate(
                    loaded.profile,
                    loaded.shell_path,
                    idoc,
                    args.output,
                    findings=gen_findings,
                )
            elif loaded.kind == "pptx":
                from brandkit.formats.pptx import generate as pptx_generate

                idoc = parse_idoc(data)
                content_hash = _content_hash(idoc.to_dict())
                out = pptx_generate.generate(
                    loaded.profile,
                    loaded.shell_path,
                    idoc,
                    args.output,
                    findings=gen_findings,
                )
            elif loaded.kind == "xlsx":
                from brandkit.formats.xlsx import generate as xlsx_generate

                grid = parse_grid(data)
                content_hash = _content_hash(_grid_to_dict(grid))
                out = xlsx_generate.generate(
                    loaded.profile,
                    loaded.shell_path,
                    grid,
                    args.output,
                    findings=gen_findings,
                )
            else:
                raise ValueError(f"unsupported profile kind: {loaded.kind}")
        except Exception as exc:
            print(f"ERROR generate: {exc}")
            return 1
        from brandkit.qa import report as vreport
        from brandkit.qa import visual as vqa

        visual_dir = vqa.default_out_dir(args.output)
        # B2: discover prior SAME-SHELL reports BEFORE writing this run's report,
        # so the cross-run regression findings can be computed and folded in. This
        # discovery (sha derivation included) is a pure SIDE artifact: it degrades
        # to no priors on any error so it can NEVER raise into the gate.
        try:
            prior_reports = vreport.discover_prior_reports(
                visual_dir,
                shell_sha256=vreport.report_shell_sha256(
                    loaded.profile, loaded.shell_path
                ),
                exclude=visual_dir / vreport.REPORT_FILENAME,
            )
        except Exception:
            prior_reports = []
        report = run_qa(
            out,
            loaded.profile,
            qa=args.qa,
            shell=loaded.shell_path,
            extra_findings=gen_findings,
            out_dir=visual_dir,
            # Generate-time only: records the artifact's content sha into the visual
            # manifest (so the L2 model can scope a verdict to it) and gates the L2
            # short-circuit. ``cli.verify`` passes no content_hash => stays None =>
            # the short-circuit never fires there (verify behaviour unchanged).
            content_hash=content_hash,
        )
        # B2: fold cross-run regression findings into THIS run's report so they
        # self-record into the persisted generation_report.json (making recurrence
        # visible to the next run). They are advisory INFO/WARNING -- never ERROR,
        # never in DEFAULT_L0_INVARIANTS -- so they can lift a clean verdict to
        # passed_with_warnings but can NEVER flip it to failed (report.passed, the
        # CLI return code, keys only on the absence of ERRORs and is unaffected).
        try:
            regression_findings = vreport.compute_regression_findings(
                loaded.profile, report, prior_reports
            )
        except Exception:
            regression_findings = []
        if regression_findings:
            report.findings.extend(regression_findings)
            if report.verdict != schema.VerificationStatus.FAILED.value:
                report.verdict = schema.VerificationStatus.PASSED_WITH_WARNINGS.value
        # Persist the run as a durable side artifact (degrade-to-no-op on any
        # error; the timestamp lives only in this JSON, never in the doc bytes).
        report_path = vreport.build_generation_report(
            profile=loaded.profile,
            document=out,
            report=report,
            shell_path=loaded.shell_path,
            out_dir=visual_dir,
            content_hash=content_hash,
        )
        for finding in report.findings:
            print(f"{finding.severity} {finding.check}: {finding.message}")
        # Surface the manifest path on stdout so the orchestrator can read it
        # deterministically and run the L2 visual-audit step.
        manifest_findings = [f for f in report.findings if f.check == "visual.manifest"]
        if manifest_findings:
            print(f"visual manifest: {manifest_findings[0].location}")
        if report_path is not None:
            print(f"generation report: {report_path}")
        print(f"generated {out}")
        return 0 if report.passed else 1
    if args.cmd == "comprehend-input":
        loaded = store.load_profile(args.name, args.scope)
        # B4: surface the SAME-shell generation_report.json history (the AMBIGUOUS
        # recurring-finding remainder) in the bundle so the model can propose overrides
        # corrections via ``propose-overrides``. Reuses the ``learn`` discovery (a pure
        # side artifact: degrades to no history on any error). An empty history adds NO
        # bundle key, so the no-history bundle stays byte-identical to pre-B4.
        prior_reports = _discover_learn_reports(loaded.profile, loaded.shell_path)
        bundle = comprehension_mod.comprehend_input_bundle(
            loaded.profile, prior_reports=prior_reports
        )
        print(json.dumps(bundle, indent=2, ensure_ascii=False, sort_keys=True))
        return 0
    if args.cmd == "comprehend":
        loaded = store.load_profile(args.name, args.scope)
        try:
            comp = json.loads(Path(args.input).read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            print(f"ERROR comprehend: cannot read {args.input}: {exc}")
            return 1
        generated_by = (
            comp.pop("generated_by", None) if isinstance(comp, dict) else None
        )
        result = comprehension_mod.merge(
            loaded.profile, comp, generated_by=generated_by
        )
        store.write_profile_json(loaded.directory, loaded.profile)
        if not result.ok:
            print(f"comprehension REJECTED ({len(result.problems)} problem(s)):")
            for problem in result.problems:
                print(f"  {problem}")
            return 1
        n_slots = len(loaded.profile["comprehension"].get("cover_slots") or {})
        n_idx = len(
            (loaded.profile["comprehension"].get("conventions") or {}).get("indexes")
            or []
        )
        n_frag = len(loaded.profile["comprehension"].get("fragments") or [])
        print(
            f"comprehended {args.name}: {n_slots} cover slot(s), "
            f"{n_idx} index convention(s), {n_frag} fragment(s) [present]"
        )
        return 0
    if args.cmd == "refine":
        import copy

        loaded = store.load_profile(args.name, args.scope)
        try:
            delta = json.loads(Path(args.input).read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            print(f"ERROR refine: cannot read {args.input}: {exc}")
            return 1
        # The provenance stamp travels with the delta verbatim (same as comprehend);
        # it is never one of the refinable sinks the overlay reads.
        generated_by = (
            delta.pop("generated_by", None) if isinstance(delta, dict) else None
        )
        # OVERLAY TRAP: merge is REPLACE-from-single-source, so a raw delta would WIPE
        # every existing sink. Overlay the delta onto the EXISTING present block first,
        # then route the WHOLE combined block through merge so the full fail-closed
        # validation (schema.validate + check_membership + check_fragments + check_triage)
        # re-runs and every ref re-binds to the live surface_inventories.
        existing = copy.deepcopy(loaded.profile.get("comprehension") or {})
        combined = comprehension_mod.overlay_refinement(existing, delta)
        # ADVISORY-until-accept (mirrors learn --accept): without --accept, merge into a
        # COPY of the profile so the diff can be computed without persisting a live block;
        # the prior block on disk stays authoritative. With --accept, merge into the real
        # profile and persist. A bad refinement is all-or-nothing: merge writes
        # status='rejected' and (advisory) the prior present block is left untouched.
        target = loaded.profile if args.accept else copy.deepcopy(loaded.profile)
        result = comprehension_mod.merge(target, combined, generated_by=generated_by)
        if not result.ok:
            print(f"refinement REJECTED ({len(result.problems)} problem(s)):")
            for problem in result.problems:
                print(f"  {problem}")
            return 1
        if args.accept:
            store.write_profile_json(loaded.directory, loaded.profile)
        # Confirm-as-diff: existing canonical vs post-overlay canonical, per sink. Both
        # are canonical (sorted/stable) so the diff is deterministic.
        before = json.dumps(existing, indent=2, sort_keys=True, ensure_ascii=False)
        after = json.dumps(
            target["comprehension"], indent=2, sort_keys=True, ensure_ascii=False
        )
        import difflib

        diff = "\n".join(
            difflib.unified_diff(
                before.splitlines(),
                after.splitlines(),
                fromfile="comprehension (current)",
                tofile="comprehension (refined)",
                lineterm="",
            )
        )
        if diff:
            print(diff)
        state = "persisted (LIVE)" if args.accept else "preview (--accept to persist)"
        print(f"refined {args.name}: comprehension overlay [{state}]")
        return 0
    if args.cmd == "learn":
        from brandkit.profile import overrides as overrides_mod
        from brandkit.qa import report as vreport

        loaded = store.load_profile(args.name, args.scope)
        # Discover the SAME-shell generation_report.json history this profile's prior
        # generate runs persisted (B1/B2). Mirrors the generate verb's discovery: it
        # walks the side-artifact dirs next to the profile's outputs. The discovery
        # is a pure side artifact (degrade-to-no-priors on any error), so a missing
        # history simply distils nothing rather than failing.
        reports = _discover_learn_reports(loaded.profile, loaded.shell_path)
        result = overrides_mod.learn(loaded.profile, reports)
        # ADVISORY accept gate (mirrors verify --accept): learn writes the lesson via
        # the single merge_overrides sink, but it stays OUT of the live resolver
        # (status forced 'absent' so resolve_role takes zero new branches and bytes
        # stay byte-identical) until --accept promotes it to 'present'. A single noisy
        # run therefore cannot mint a permanent LIVE lesson; the operator must opt in.
        block = (loaded.profile.get("rules") or {}).get("overrides") or {}
        if result.ok and not args.accept:
            block["status"] = schema.ComprehensionStatus.ABSENT.value
        store.write_profile_json(loaded.directory, loaded.profile)
        if not result.ok:
            if result.problems:
                print(f"learn REJECTED ({len(result.problems)} problem(s)):")
                for problem in result.problems:
                    print(f"  {problem}")
                return 1
            # Nothing crossed the recurrence threshold / bound to a brand-safe target.
            print(f"learn {args.name}: no recurring finding distilled (0 override(s))")
            return 0
        state = (
            "present (LIVE)"
            if args.accept
            else "absent (advisory; --accept to go live)"
        )
        print(
            f"learned {args.name}: {result.distilled} override(s) distilled [{state}]"
        )
        return 0
    if args.cmd == "propose-overrides":
        import copy

        from brandkit.profile import overrides as overrides_mod

        loaded = store.load_profile(args.name, args.scope)
        try:
            delta = json.loads(Path(args.input).read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            print(f"ERROR propose-overrides: cannot read {args.input}: {exc}")
            return 1
        # The provenance stamp travels with the delta verbatim (same as comprehend); it
        # is never one of the overlaid override containers.
        generated_by = (
            delta.pop("generated_by", None) if isinstance(delta, dict) else None
        )
        # OVERLAY TRAP: merge_overrides is REPLACE-from-single-source, so a raw delta
        # would WIPE a deterministic ``learn`` lesson. Overlay onto the EXISTING block
        # first, then route the WHOLE combined proposal through the single sink so the
        # full fail-closed validation (shape + membership + acyclicity) re-runs and
        # every pointer - old or new - re-binds to the live surface inventories.
        existing = copy.deepcopy(
            (loaded.profile.get("rules") or {}).get("overrides") or {}
        )
        combined = overrides_mod.overlay_overrides(
            existing, delta if isinstance(delta, dict) else {}
        )
        result = overrides_mod.merge_overrides(
            loaded.profile, combined, generated_by=generated_by
        )
        # ADVISORY accept gate (mirrors learn --accept): the proposal is written via the
        # single sink, but it stays OUT of the live resolver (status forced 'absent', so
        # resolve_role takes zero new branches and bytes stay byte-identical) until
        # --accept promotes it to 'present'. A single noisy proposal therefore cannot
        # mint a permanent LIVE correction; the operator must opt in.
        block = (loaded.profile.get("rules") or {}).get("overrides") or {}
        if result.ok and not args.accept:
            block["status"] = schema.ComprehensionStatus.ABSENT.value
        store.write_profile_json(loaded.directory, loaded.profile)
        if not result.ok:
            print(f"propose-overrides REJECTED ({len(result.problems)} problem(s)):")
            for problem in result.problems:
                print(f"  {problem}")
            return 1
        n_reroute = len(block.get("reroute_roles") or {})
        n_swap = len(block.get("number_format_swaps") or {})
        n_demo = len(block.get("demo_clears") or [])
        state = (
            "present (LIVE)"
            if args.accept
            else "absent (advisory; --accept to go live)"
        )
        print(
            f"proposed-overrides {args.name}: {n_reroute} reroute(s), {n_swap} swap(s), "
            f"{n_demo} demo-clear(s) [{state}]"
        )
        return 0
    if args.cmd == "compare-profiles":
        from brandkit.profile import compare

        a = store.load_profile(args.name_a, args.scope_a)
        b = store.load_profile(args.name_b, args.scope_b)
        result = compare.compare_profiles(a.profile, b.profile)
        if args.json:
            print(json.dumps(result, indent=2, sort_keys=True))
        else:
            print(compare.render_report(result))
        return 0 if result["verdict"] == compare.VERDICT_ALIGNED else 1
    if args.cmd == "list":
        for summary in store.list_profiles():
            if args.scope != "auto" and summary.scope != args.scope:
                continue
            shadow = " shadowed" if summary.shadowed else ""
            print(
                f"{summary.name}\t{summary.scope}\t{summary.kind}\t{summary.verification_status}{shadow}"
            )
        return 0
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
