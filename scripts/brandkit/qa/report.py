# SPDX-License-Identifier: MIT
"""Persisted generation report (model-free engine side artifact).

This module writes ``generation_report.json`` next to a generated output, the
durable record of one ``generate`` run's QA verdict + findings + the hashes that
let a LATER run recognize the SAME input against the SAME shell. It is the
substrate the "learn-from-errors" overrides (Cluster B) read across runs: B2
diffs successive reports for recurring findings, B3 distills the recurrences
into deterministic overrides.

It is a pure SIDE ARTIFACT, mirroring ``qa/visual.py`` exactly:

  * written into the same ``<output>.visual`` dir as the PNGs + visual manifest;
  * deterministic JSON (``indent=2``, trailing newline) -- the ONLY volatile
    field is ``generated_at`` (an ISO-8601 UTC timestamp), which lives ONLY in
    this JSON and is read by NO generator path, so the generated ``.docx`` /
    ``.pptx`` / ``.xlsx`` bytes are identical across runs;
  * the ``findings`` list is recorded VERBATIM (never reordered or deduped) --
    that ordering/multiplicity IS the recurrence signal B2/B3 key on;
  * the ``content_sha256`` is the hash of the CANONICAL ``to_dict()`` JSON of the
    parsed input (not the raw author file bytes), so two inputs that differ only
    in key order / whitespace hash equal -- this is what B2 keys "same input" on;
  * nothing here raises into the gate: any IO/serialization failure degrades to a
    no-op (returns ``None``). A failed report write can NEVER flip a verdict.

The writer is GENERATE-ONLY. ``verify`` also produces a ``QAReport`` but has no
output / content hash; a hash-less row would pollute the same-shell history B2
partitions on, so ``verify`` writes no report.
"""

from __future__ import annotations

import datetime
import json
from pathlib import Path

from brandkit.profile import schema
from brandkit.profile.store import sha256_file
from brandkit.qa.model import Finding, QAReport

# ---------------------------------------------------------------------------
# Report constants (mirror qa/visual.py's MANIFEST_FILENAME / SCHEMA_VERSION)
# ---------------------------------------------------------------------------
REPORT_FILENAME: str = "generation_report.json"
REPORT_SCHEMA_VERSION: str = "generation-report-1"


def report_shell_sha256(profile: dict, shell_path: str | Path | None) -> str | None:
    """Derive the shell sha a report is keyed on: live bytes, else provenance.

    Single derivation shared by ``build_generation_report`` (what it records) and
    ``discover_prior_reports``' caller (what it partitions on), so the new run's
    recorded ``shell_sha256`` and the sha used to find its priors can NEVER drift.
    Prefers the live shell bytes when the file exists; falls back to the
    ``provenance.shell.sha256`` recorded in the profile (mirror
    ``store.comprehension_is_present``'s provenance access path).
    """
    if shell_path is not None and Path(shell_path).is_file():
        return sha256_file(shell_path)
    return ((profile.get("provenance") or {}).get("shell") or {}).get("sha256")


def _now_iso_utc() -> str:
    """Return the current instant as an ISO-8601 UTC string (the volatile field)."""
    return (
        datetime.datetime.now(datetime.timezone.utc)
        .replace(microsecond=0)
        .isoformat()
        .replace("+00:00", "Z")
    )


def build_generation_report(
    *,
    profile: dict,
    document: str | Path,
    report: QAReport,
    shell_path: str | Path | None,
    out_dir: str | Path,
    content_hash: str | None,
    generated_at: str | None = None,
) -> Path | None:
    """Build and write ``<out_dir>/generation_report.json`` (a SIDE artifact).

    Returns the report path, or ``None`` if anything failed (degrade-to-no-op:
    this writer must NEVER raise into the gate, and a failed write must never
    flip the run's verdict).

    Args:
        profile: the brand profile under test (read-only).
        document: the generated output path (only its basename is recorded).
        report: the ``QAReport`` from ``run_qa`` (verdict + findings, verbatim).
        shell_path: the brand shell path; hashed when it exists, else the report
            falls back to ``provenance.shell.sha256`` recorded in the profile.
        out_dir: the side-artifact dir (the same ``<output>.visual`` dir the PNGs
            + visual manifest use). Created if absent (mirrors ``visual.py``), so
            the report is written even when the gate skipped the visual render.
        content_hash: the precomputed hash of the canonical input ``to_dict()``
            JSON (the caller derives it; ``None`` is recorded as-is when unknown).
        generated_at: override for the volatile timestamp (tests inject a fixed
            value); defaults to the current UTC instant.
    """
    try:
        out_dir = Path(out_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        document = Path(document)

        # Shell hash: prefer the live shell bytes, else the recorded provenance sha
        # (mirror store.comprehension_is_present's provenance access path). Shared
        # with discover_prior_reports so what we record and what we partition on
        # can never drift.
        shell_sha256: str | None = report_shell_sha256(profile, shell_path)

        report_doc: dict = {
            "schema_version": REPORT_SCHEMA_VERSION,
            "kind": profile.get("kind"),
            "profile_name": (profile.get("identity") or {}).get("name"),
            "document": document.name,
            "verdict": report.verdict,
            "shell_sha256": shell_sha256,
            "content_sha256": content_hash,
            "output_sha256": sha256_file(document) if document.is_file() else None,
            # Verbatim: NOT reordered, NOT deduped -- the ordering/multiplicity is
            # the recurrence signal B2/B3 read.
            "findings": [
                {
                    "check": f.check,
                    "severity": f.severity,
                    "message": f.message,
                    "location": f.location,
                }
                for f in report.findings
            ],
            # The ONLY volatile field; lives only here, read by no generator path.
            "generated_at": generated_at or _now_iso_utc(),
        }

        path = out_dir / REPORT_FILENAME
        path.write_text(
            json.dumps(report_doc, indent=2, ensure_ascii=False) + "\n",
            encoding="utf-8",
        )
        return path
    except Exception:
        # Side-artifact failure contract: never raise into the gate, never flip a
        # verdict. The CLI keeps returning ``0 if report.passed else 1`` regardless.
        return None


# ---------------------------------------------------------------------------
# B2 -- cross-run regression findings (still a pure side artifact)
# ---------------------------------------------------------------------------
# These regression findings are ADVISORY (INFO/WARNING). They are NOT listed in
# ``schema.DEFAULT_L0_INVARIANTS`` and carry no ERROR severity, so they can never
# flip a verdict. They are folded into the QAReport (and so into the persisted
# ``generation_report.json``) of the run that detects them, which makes
# recurrence self-recording for the NEXT run.
REGRESSION_RECURRED: str = "regression.recurred"
REGRESSION_REINTRODUCED: str = "regression.reintroduced"


def discover_prior_reports(
    side_artifact_dir: str | Path,
    *,
    shell_sha256: str | None,
    exclude: str | Path | None = None,
) -> list[dict]:
    """Discover prior ``generation_report.json`` digests for the SAME shell.

    The CLI calls this BEFORE writing the new run's report. It globs the
    side-artifact dir convention fixed in B1 -- each ``generate`` writes its
    report into ``<output>.visual/generation_report.json`` next to the output,
    so sibling runs live in sibling ``*.visual`` dirs under the same parent. We
    walk that parent for every ``generation_report.json`` and keep ONLY those
    whose ``shell_sha256`` matches the live shell: history is partitioned by
    shell so a re-extract (new sha) starts a fresh history (SHELL-FROZEN).

    Args:
        side_artifact_dir: the ``<output>.visual`` dir B1 writes the new report
            into (its PARENT is scanned for sibling reports).
        shell_sha256: the live shell sha to partition on; when ``None`` no prior
            can match (an unknown shell shares history with nothing).
        exclude: a report path to skip (the new run's own ``visual_dir``), so a
            re-generate to the same output never counts its own stale report as a
            prior.

    Never raises: a malformed/partial report on disk is skipped, not fatal, and
    any unexpected IO error degrades the whole discovery to an empty list (the
    same side-artifact contract as ``build_generation_report``). The result is
    sorted by ``generated_at`` ascending (oldest first) so callers can treat the
    LAST element as the immediately-prior run deterministically.
    """
    if not shell_sha256:
        return []
    try:
        side_artifact_dir = Path(side_artifact_dir)
        candidates = sorted(
            side_artifact_dir.parent.glob(f"*.visual/{REPORT_FILENAME}")
        )
        return load_same_shell_reports(
            candidates, shell_sha256=shell_sha256, exclude=exclude
        )
    except Exception:
        return []


def load_same_shell_reports(
    candidates,
    *,
    shell_sha256: str,
    exclude: str | Path | None = None,
) -> list[dict]:
    """Read, filter, and order report candidates - the ONE per-candidate pipeline.

    Shared by :func:`discover_prior_reports` (sibling ``*.visual`` dirs) and the CLI
    ``learn`` discovery (a cwd-wide walk): each caller supplies its own candidate
    paths; the dedup (by resolved path), the JSON read, the dict guard, the
    same-shell ``shell_sha256`` partition, and the deterministic oldest->newest
    ``generated_at`` ordering live HERE so the two read paths can never drift. A
    malformed/unreadable candidate is skipped, never fatal (the side-artifact
    contract).
    """
    exclude_resolved: Path | None = None
    if exclude is not None:
        try:
            exclude_resolved = Path(exclude).resolve()
        except OSError:
            exclude_resolved = None
    reports: list[dict] = []
    seen: set[Path] = set()
    for candidate in candidates:
        try:
            resolved = Path(candidate).resolve()
            if resolved in seen or resolved == exclude_resolved:
                continue
            seen.add(resolved)
            doc = json.loads(Path(candidate).read_text(encoding="utf-8"))
        except (OSError, ValueError):
            continue
        if not isinstance(doc, dict):
            continue
        if doc.get("shell_sha256") != shell_sha256:
            continue
        reports.append(doc)
    # Deterministic order: oldest -> newest by the ONLY volatile field. Missing
    # timestamps sort first (treated as oldest); ties keep the caller's path order.
    reports.sort(key=lambda d: d.get("generated_at") or "")
    return reports


def _location_key(finding: dict) -> tuple[str, str | None]:
    """The stable recurrence key for a serialized finding: ``(check, location)``.

    NEVER the ``message`` body -- the message can carry brand/template text (which
    would leak brand words into the key) and varies run to run (which would
    over-fragment recurrence). ``(check, location)`` is the universal,
    template-agnostic identity B2/B3/B4 key on.
    """
    return (finding.get("check"), finding.get("location"))


def _finding_keys(report_findings: list) -> list[tuple[str, str | None]]:
    """Project a run's findings (objects OR dicts) to ``(check, location)`` keys.

    Accepts either ``Finding`` dataclasses (this run's live ``report.findings``)
    or the serialized dicts persisted in a prior ``generation_report.json``.
    """
    keys: list[tuple[str, str | None]] = []
    for f in report_findings:
        if isinstance(f, dict):
            keys.append(_location_key(f))
        else:
            keys.append((f.check, f.location))
    return keys


def compute_regression_findings(
    profile: dict,
    report: QAReport,
    prior_reports: list[dict],
) -> list[Finding]:
    """Diff this run's findings against prior SAME-SHELL runs for recurrences.

    The recurrence key is STRICTLY the ``(check, location)`` multiset, never the
    ``message`` body (which can carry brand/template text). ``generated_at`` is
    ignored entirely (it is the only volatile field). The caller is responsible
    for passing ONLY same-shell priors (see ``discover_prior_reports``); this
    function trusts that partition and does not re-filter.

    Emits, for each ``(check, location)`` present in THIS run:
      * ``regression.recurred`` -- the pair also appears in >= 1 prior same-shell
        run. The message carries a ``recurred_runs`` count (this run + the priors
        that show it) -- the threshold B3/B4 distillation gates on.
      * ``regression.reintroduced`` -- the pair is ABSENT in the immediately-prior
        run but present in an EARLIER one (it came back after a clean run).

    Both are ADVISORY: INFO for recurred, WARNING for reintroduced (a clean run
    going dirty again is the louder signal). Neither is an ERROR, neither is in
    ``DEFAULT_L0_INVARIANTS``, so neither can flip a verdict.

    ``profile`` is accepted for signature parity with the other report builders
    (and so a future refinement could scope by kind); it is read-only here.
    """
    _ = profile  # signature parity; not consulted today (key is shell-agnostic)
    if not prior_reports:
        return []

    this_keys = _finding_keys(report.findings)
    if not this_keys:
        return []

    # Per-prior key sets (for the recurred count + the reintroduced gap test).
    prior_key_sets: list[set[tuple[str, str | None]]] = []
    for prior in prior_reports:
        findings = prior.get("findings")
        if not isinstance(findings, list):
            findings = []
        prior_key_sets.append(set(_finding_keys(findings)))

    # ``prior_reports`` is sorted oldest -> newest; the LAST is immediately-prior.
    immediately_prior = prior_key_sets[-1]
    earlier = prior_key_sets[:-1]

    findings: list[Finding] = []
    # Iterate the DISTINCT keys of this run, in first-seen order, so the emitted
    # findings are deterministic and stable across runs.
    seen: set[tuple[str, str | None]] = set()
    for key in this_keys:
        if key in seen:
            continue
        seen.add(key)
        check, location = key
        n_prior_with = sum(1 for s in prior_key_sets if key in s)
        if n_prior_with == 0:
            continue  # brand-new this run; not a regression
        recurred_runs = n_prior_with + 1  # the priors that show it + this run
        findings.append(
            Finding(
                REGRESSION_RECURRED,
                schema.Severity.INFO.value,
                f"finding '{check}' recurred across {recurred_runs} same-shell "
                f"run(s) (recurred_runs={recurred_runs})",
                location=location,
            )
        )
        # Reintroduced: gone in the immediately-prior run, present in an earlier
        # one -> it came back after a clean run. (Strict subset of recurred.)
        if key not in immediately_prior and any(key in s for s in earlier):
            findings.append(
                Finding(
                    REGRESSION_REINTRODUCED,
                    schema.Severity.WARNING.value,
                    f"finding '{check}' reintroduced: absent in the prior run, "
                    f"present in an earlier same-shell run",
                    location=location,
                )
            )
    return findings
