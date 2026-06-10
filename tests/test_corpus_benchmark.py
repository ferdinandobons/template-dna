# SPDX-License-Identifier: MIT
"""Smoke guard for the local-corpus fidelity runner (scripts/corpus_benchmark.py).

The real corpus is local-only and never committed; CI smokes the runner with
the SYNTHETIC example template as a stand-in corpus, proving the pipeline
walk, the outside-the-repo guard, and the report artifacts.
"""

from __future__ import annotations

import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

import corpus_benchmark  # noqa: E402


class CorpusBenchmarkTest(unittest.TestCase):
    def test_refuses_a_corpus_inside_the_repo(self):
        rc = corpus_benchmark.main(["--corpus", str(ROOT / "examples")])
        self.assertEqual(rc, 2)

    def test_failing_template_yields_exit_1(self):
        # A garbage "template" fails extract; the runner must aggregate that
        # row as a failure and exit 1 (the scriptable-gate contract). Guards
        # the predicate that also catches EXCEPTION rows on any step.
        with tempfile.TemporaryDirectory() as td:
            corpus = Path(td) / "corpus"
            (corpus / "templates").mkdir(parents=True)
            (corpus / "templates" / "broken.docx").write_bytes(b"not a zip")
            rc = corpus_benchmark.main(["--corpus", str(corpus), "--qa", "fast"])
            self.assertEqual(rc, 1)
            reports = list((corpus / "reports").glob("*/report.json"))
            self.assertEqual(len(reports), 1)
            row = json.loads(reports[0].read_text())[0]
            self.assertNotEqual(row.get("extract"), "ok")

    def test_smoke_run_on_synthetic_corpus(self):
        with tempfile.TemporaryDirectory() as td:
            corpus = Path(td) / "corpus"
            (corpus / "templates").mkdir(parents=True)
            shutil.copy(
                ROOT / "examples" / "templates" / "branddocs_template.docx",
                corpus / "templates" / "synthetic.docx",
            )
            rc = corpus_benchmark.main(["--corpus", str(corpus), "--qa", "fast"])
            self.assertEqual(rc, 0)
            reports = list((corpus / "reports").glob("*/report.json"))
            self.assertEqual(len(reports), 1)
            results = json.loads(reports[0].read_text())
            self.assertEqual(len(results), 1)
            entry = results[0]
            self.assertEqual(entry["extract"], "ok")
            self.assertEqual(entry["verify"], "passed")
            self.assertEqual(entry["generate"], "ok")
            self.assertGreater(entry["output_bytes"], 0)
            md = reports[0].with_name("report.md").read_text()
            self.assertIn("synthetic.docx", md)
            self.assertIn("LibreOffice", md)


if __name__ == "__main__":
    unittest.main()
