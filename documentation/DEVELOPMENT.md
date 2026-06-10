# Development

```bash
python3 -m venv .venv && . .venv/bin/activate
pip install -r requirements.txt pytest ruff
PYTHONPATH=scripts pytest -q        # docx / pptx / xlsx / security / integration / smoke suites
```

The full local gate (what CI runs, mirror of
[`.github/workflows/ci.yml`](../.github/workflows/ci.yml)):

```bash
ruff check . && ruff format --check .                  # lint + format, ruff.toml is the source of truth
PYTHONPATH=scripts python -m pytest -q                 # full model-free suite
BRANDDOCS_RUN_REAL_RENDER=1 PYTHONPATH=scripts \
  python -m pytest tests/test_visual_qa.py -q          # real-render lane (needs LibreOffice + Poppler)
```

> **Never commit real templates or company assets.** `brand-kit/` and `generated/`
> are intentionally git-ignored, and `tests/test_no_proprietary.py` fails the build
> if any Office binary is tracked outside `tests/fixtures` or `examples/` (or a
> vendored proprietary import sneaks in).

See [`CONTRIBUTING.md`](../CONTRIBUTING.md) for the PR rules, the frozen
vocabulary in [`CONVENTIONS.md`](../CONVENTIONS.md), and the engine internals
map in [`scripts/brandkit/README.md`](../scripts/brandkit/README.md) before
opening a PR.

## Fidelity corpus (local-only)

Real company templates never enter the repo, but fidelity should be measured
against them. Keep a corpus directory OUTSIDE the repository:

```text
~/my-branddocs-corpus/
|- templates/   your real .docx / .pptx / .xlsx files (never committed)
`- reports/     written by the runner, dated
```

Run the benchmark from the repo root; it extracts each template into a
throwaway directory, verifies, generates a brand-agnostic probe document
(docx/pptx), runs QA, and writes `reports/<date>/report.{md,json}` next to the
corpus. Exit 1 when any template fails a step.

```bash
python scripts/corpus_benchmark.py --corpus ~/my-branddocs-corpus [--qa auto]
```

The probe input is brand-agnostic by construction (typed blocks, no styles or
colors), so the corpus measures the engine without ever tuning it. The report
header states the renderer caveat: visual QA certifies faithfulness according
to LibreOffice; Word may differ on field caches and font fallback.

## Release checklist

Releases are maintainer-driven and cumulative: push to `main` freely, tag only
at a stable point that deserves a release on its own. The practiced sequence:

1. **Local gate green first.** Run the full local gate above (lint, suite,
   real-render lane). Never bump a version on a red gate.
2. **README sync check.** Grep the README for every term the release touched
   (command names, verbs, paths, behaviors) and update it in the same release.
3. Push the pending commits to `origin/main`.
4. **Bump the version everywhere it lives:** `CITATION.cff`, `pyproject.toml`,
   `docs/index.html` (JSON-LD `softwareVersion`), `.claude-plugin/plugin.json`,
   `.claude-plugin/marketplace.json`.
5. Convert the `[Unreleased]` section of [`CHANGELOG.md`](../CHANGELOG.md) into
   the new version entry (with date) and add the link references.
6. Update the "Latest release" links in the README to the new tag.
7. Commit the bump: `chore(release): vX.Y.Z`, push it.
8. **Wait for remote CI to be green on the bump commit.** Tag only after.
9. Annotated tag: `git tag -a vX.Y.Z -m "..."`, then `git push origin vX.Y.Z`.
10. `gh release create vX.Y.Z` with the CHANGELOG section as the body.

Two hard rules: never force-push a tag (a bad release gets a follow-up patch
vX.Y.Z+1, never a re-pointed tag), and never tag with a red or pending CI.

Release notes live in [`CHANGELOG.md`](../CHANGELOG.md). The shared engine,
profile schema, and QA gate are exercised by the CI suite
([`.github/workflows/ci.yml`](../.github/workflows/ci.yml)) on every push.
