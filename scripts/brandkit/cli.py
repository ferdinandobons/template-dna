# SPDX-License-Identifier: MIT
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from brandkit import doctor
from brandkit.formats.docx import extract as docx_extract
from brandkit.formats.docx import generate as docx_generate
from brandkit.formats.pptx import extract as pptx_extract
from brandkit.formats.pptx import generate as pptx_generate
from brandkit.formats.xlsx import extract as xlsx_extract
from brandkit.formats.xlsx import generate as xlsx_generate
from brandkit.grid.model import parse_grid
from brandkit.ir.model import parse_idoc
from brandkit.profile import schema, store
from brandkit.qa.gate import run_qa


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="template-dna")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("extract")
    p.add_argument("--name", required=True)
    p.add_argument("--template", required=True)
    p.add_argument("--scope", default="project", choices=("auto", "project", "global"))

    p = sub.add_parser("verify")
    p.add_argument("--name", required=True)
    p.add_argument("--scope", default="auto", choices=("auto", "project", "global"))
    p.add_argument("--qa", default="auto", choices=("auto", "fast", "deep"))
    p.add_argument("--accept", action="store_true")

    p = sub.add_parser("generate")
    p.add_argument("--name", required=True)
    p.add_argument("--input", required=True)
    p.add_argument("--output", required=True)
    p.add_argument("--scope", default="auto", choices=("auto", "project", "global"))
    p.add_argument("--qa", default="auto", choices=("auto", "fast", "deep"))

    p = sub.add_parser("list")
    p.add_argument("--scope", default="auto", choices=("auto", "project", "global"))

    sub.add_parser("doctor")
    args = parser.parse_args(argv)

    if args.cmd == "doctor":
        doctor.print_report()
        return 0
    if args.cmd == "extract":
        path = Path(args.template)
        suffix = path.suffix.lower()
        if suffix == ".docx":
            profile_json = docx_extract.extract(path, args.name, scope=args.scope)
        elif suffix == ".pptx":
            profile_json = pptx_extract.extract(path, args.name, scope=args.scope)
        elif suffix == ".xlsx":
            profile_json = xlsx_extract.extract(path, args.name, scope=args.scope)
        else:
            raise SystemExit("supported templates: .docx, .pptx, .xlsx")
        loaded = store.load_profile(args.name, args.scope)
        report = run_qa(None, loaded.profile, mode="verify", qa="fast", shell=loaded.shell_path)
        loaded.profile["verification"]["status"] = report.verdict
        loaded.profile["verification"]["roles_total"] = len(schema.list_role_ids(loaded.profile))
        loaded.profile["verification"]["roles_verified"] = loaded.profile["verification"]["roles_total"]
        profile_json.write_text(json.dumps(loaded.profile, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
        print(f"extracted {args.name} -> {profile_json}")
        return 0 if report.passed else 1
    if args.cmd == "verify":
        loaded = store.load_profile(args.name, args.scope)
        report = run_qa(None, loaded.profile, mode="verify", qa=args.qa, shell=loaded.shell_path)
        loaded.profile["verification"]["status"] = report.verdict
        if args.accept and report.passed:
            loaded.profile.setdefault("verification", {})["accepted"] = True
        (loaded.directory / "profile.json").write_text(
            json.dumps(loaded.profile, indent=2, ensure_ascii=False) + "\n",
            encoding="utf-8",
        )
        for finding in report.findings:
            print(f"{finding.severity} {finding.check}: {finding.message}")
        print(f"verification: {report.verdict}")
        return 0 if report.passed else 1
    if args.cmd == "generate":
        loaded = store.load_profile(args.name, args.scope)
        data = json.loads(Path(args.input).read_text(encoding="utf-8"))
        gen_findings: list = []
        try:
            if loaded.kind == "docx":
                out = docx_generate.generate(
                    loaded.profile, loaded.shell_path, parse_idoc(data), args.output, findings=gen_findings
                )
            elif loaded.kind == "pptx":
                out = pptx_generate.generate(loaded.profile, loaded.shell_path, parse_idoc(data), args.output)
            elif loaded.kind == "xlsx":
                out = xlsx_generate.generate(loaded.profile, loaded.shell_path, parse_grid(data), args.output)
            else:
                raise ValueError(f"unsupported profile kind: {loaded.kind}")
        except Exception as exc:
            print(f"ERROR generate: {exc}")
            return 1
        report = run_qa(
            out, loaded.profile, mode="generate", qa=args.qa,
            shell=loaded.shell_path, extra_findings=gen_findings,
        )
        for finding in report.findings:
            print(f"{finding.severity} {finding.check}: {finding.message}")
        print(f"generated {out}")
        return 0 if report.passed else 1
    if args.cmd == "list":
        for summary in store.list_profiles():
            if args.scope != "auto" and summary.scope != args.scope:
                continue
            shadow = " shadowed" if summary.shadowed else ""
            print(f"{summary.name}\t{summary.scope}\t{summary.kind}\t{summary.verification_status}{shadow}")
        return 0
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
