#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
"""Fidelity benchmark over a LOCAL-ONLY corpus of real templates.

Real company templates never enter the repo (tests/test_no_proprietary.py
enforces it), but fidelity must still be measured against them. This runner
walks a corpus directory that lives OUTSIDE the repository, runs the full
pipeline on every template, and writes a dated fidelity report NEXT TO the
corpus (never inside the repo):

    <corpus>/
    |- templates/   real .docx / .pptx / .xlsx files (yours, local-only)
    `- reports/     written by this runner: <UTC date>/report.{md,json}

Per template: extract -> verify -> (docx/pptx) generate a canonical
brand-agnostic probe document -> QA. The probe input contains no styles,
colors, fonts or template-specific words, so the corpus never tunes the
engine - it only measures it. xlsx templates run extract + verify only in v1
(a meaningful workbook fill needs template-specific region intent, which a
generic probe cannot author honestly).

Usage (from the repo root):

    python scripts/corpus_benchmark.py --corpus /path/to/corpus [--qa auto]

Profiles are extracted into a throwaway temp directory; nothing is written
into the repo or the user's brand-kit stores. Renderer availability degrades
QA exactly like the engine does (L0-only), and the report records the level
actually achieved, with Word-vs-LibreOffice caveats stated in the header.
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO / "scripts"))

SUPPORTED = (".docx", ".pptx", ".xlsx")

# The canonical brand-agnostic probe: typed blocks only, no brand facts, no
# template-specific words. Exercises cover, toc, heading hierarchy, list,
# table + caption - the surfaces the fidelity guarantee covers.
PROBE_IDOC = {
    "cover": {"title": "Corpus probe", "subtitle": "Fidelity benchmark run"},
    "blocks": [
        {"type": "toc"},
        {"type": "heading", "level": 1, "text": "First probe section"},
        {"type": "paragraph", "text": "A short body paragraph for the probe."},
        {"type": "heading", "level": 2, "text": "Nested probe subsection"},
        {
            "type": "list",
            "style": "bullet",
            "items": ["First probe item", "Second probe item"],
        },
        {"type": "heading", "level": 1, "text": "Second probe section"},
        {
            "type": "table",
            "columns": ["Alpha", "Beta"],
            "rows": [["1", "2"], ["3", "4"]],
            "caption": "Probe data table",
        },
    ],
}


def _utc_date() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%d")


def bench_template(template: Path, qa: str, scratch: Path) -> dict:
    """Run the pipeline on one template; never raises (records the failure)."""
    from brandkit import cli

    entry: dict = {"template": template.name, "kind": template.suffix.lstrip(".")}
    name = "corpus-" + template.stem.lower().replace(" ", "-")[:40]
    t0 = time.monotonic()
    step = "extract"
    try:
        rc = cli.main(
            [
                "extract",
                "--name",
                name,
                "--template",
                str(template),
                "--scope",
                "project",
            ]
        )
        entry["extract"] = "ok" if rc == 0 else f"rc={rc}"
        profile = json.loads(
            (scratch / "brand-kit" / name / "profile.json").read_text()
        )
        entry["roles"] = len((profile.get("roles") or {}).get("_index") or [])
        entry["cover_kind"] = ((profile.get("anchors") or {}).get("cover") or {}).get(
            "kind"
        )

        step = "verify"
        rc = cli.main(["verify", "--name", name, "--scope", "project", "--qa", qa])
        entry["verify"] = "passed" if rc == 0 else "failed"

        if template.suffix.lower() in (".docx", ".pptx"):
            step = "generate"
            idoc = scratch / f"{name}.idoc.json"
            idoc.write_text(json.dumps(PROBE_IDOC), encoding="utf-8")
            out = scratch / f"{name}-out{template.suffix.lower()}"
            rc = cli.main(
                [
                    "generate",
                    "--name",
                    name,
                    "--input",
                    str(idoc),
                    "--output",
                    str(out),
                    "--scope",
                    "project",
                    "--qa",
                    qa,
                ]
            )
            entry["generate"] = "ok" if rc == 0 else f"rc={rc}"
            entry["output_bytes"] = out.stat().st_size if out.is_file() else 0
        else:
            entry["generate"] = "skipped (xlsx v1: extract+verify only)"
    except Exception as exc:  # the report IS the error channel here
        entry[step] = f"EXCEPTION: {exc}"
    entry["seconds"] = round(time.monotonic() - t0, 1)
    return entry


def render_markdown(results: list[dict], corpus: Path, qa: str) -> str:
    lines = [
        "# BrandDocs corpus fidelity report",
        "",
        f"- corpus: {corpus}",
        f"- date (UTC): {_utc_date()}",
        f"- qa level requested: {qa}",
        "- renderer caveat: visual QA certifies faithfulness ACCORDING TO"
        " LibreOffice; Word may differ on field caches (TOC page numbers"
        " populate on open) and font fallback when brand fonts are not"
        " installed.",
        "",
        "| template | kind | roles | cover | extract | verify | generate | s |",
        "|---|---|---|---|---|---|---|---|",
    ]
    for r in results:
        lines.append(
            "| {template} | {kind} | {roles} | {cover_kind} | {extract} "
            "| {verify} | {generate} | {seconds} |".format(
                **{
                    "roles": r.get("roles", "-"),
                    "cover_kind": r.get("cover_kind", "-"),
                    "extract": r.get("extract", "-"),
                    "verify": r.get("verify", "-"),
                    "generate": r.get("generate", "-"),
                    "seconds": r.get("seconds", "-"),
                    "template": r.get("template", "?"),
                    "kind": r.get("kind", "?"),
                }
            )
        )
    lines.append("")
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="corpus_benchmark")
    parser.add_argument(
        "--corpus",
        required=True,
        help="corpus directory (holds templates/; reports/ is written next to it)",
    )
    parser.add_argument(
        "--qa", default="auto", choices=("auto", "fast", "deep", "strict")
    )
    args = parser.parse_args(argv)

    corpus = Path(args.corpus).resolve()
    templates_dir = corpus / "templates"
    if not templates_dir.is_dir():
        print(f"ERROR: {templates_dir} does not exist", file=sys.stderr)
        return 2
    try:
        corpus.relative_to(REPO)
        print(
            "ERROR: the corpus must live OUTSIDE the repository "
            "(real templates are never committed)",
            file=sys.stderr,
        )
        return 2
    except ValueError:
        pass  # outside the repo: correct

    templates = sorted(
        p
        for p in templates_dir.iterdir()
        if p.suffix.lower() in SUPPORTED and not p.name.startswith("~")
    )
    if not templates:
        print(f"no templates found in {templates_dir}", file=sys.stderr)
        return 2

    import os
    import tempfile

    results = []
    with tempfile.TemporaryDirectory(prefix="branddocs-corpus-") as td:
        scratch = Path(td)
        old = Path.cwd()
        os.chdir(scratch)  # project-scope brand-kit lands in the scratch dir
        try:
            for template in templates:
                print(f"== {template.name}")
                results.append(bench_template(template, args.qa, scratch))
        finally:
            os.chdir(old)

    report_dir = corpus / "reports" / _utc_date()
    report_dir.mkdir(parents=True, exist_ok=True)
    (report_dir / "report.json").write_text(
        json.dumps(results, indent=2, sort_keys=True), encoding="utf-8"
    )
    md = render_markdown(results, corpus, args.qa)
    (report_dir / "report.md").write_text(md, encoding="utf-8")
    print(md)
    print(f"report written to {report_dir}")

    # A row is healthy only when every step landed in its SUCCESS vocabulary;
    # anything else (rc=, EXCEPTION, failed, a step never reached) fails the
    # run, so the exit code is a real fidelity gate.
    def _row_ok(r: dict) -> bool:
        gen = str(r.get("generate", ""))
        return (
            r.get("extract") == "ok"
            and r.get("verify") == "passed"
            and (gen == "ok" or gen.startswith("skipped"))
        )

    failed = [r for r in results if not _row_ok(r)]
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
