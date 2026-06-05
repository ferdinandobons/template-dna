# SPDX-License-Identifier: MIT
"""Two-stage visual audit tests (no soffice required).

The L1 proxy tests feed SYNTHETIC PIL images (blank page, edge bleed, centered
content, landscape) so they assert the deterministic pixel proxies without ever
invoking ``soffice``. The wiring/degrade tests monkeypatch the renderer so the
gate's ``--qa fast|auto|deep`` semantics and the clean CI degrade are proven
without external tools. One gated end-to-end test runs the real render when the
binaries are present (skipped in CI).
"""
from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from docx import Document
from PIL import Image, ImageDraw

from brandkit import doctor
from brandkit.profile import schema
from brandkit.qa import gate
from brandkit.qa import visual as vqa


def _real_docx(path: Path) -> Path:
    """Write a minimal but VALID .docx so the L0 docx check can open it as a zip.

    The wiring tests only care about the visual path; the file just has to be a
    real OOXML package the L0 layer can read without errors.
    """
    doc = Document()
    doc.add_paragraph("placeholder body")
    doc.save(path)
    return path


# ---------------------------------------------------------------------------
# Synthetic-image helpers (US-Letter-ish portrait at 100 DPI = 850x1100)
# ---------------------------------------------------------------------------
def _blank(width: int = 850, height: int = 1100) -> Image.Image:
    return Image.new("L", (width, height), 255)


def _centered_content(width: int = 850, height: int = 1100) -> Image.Image:
    """A page with a dark block well inside the margins (clean edges)."""
    img = _blank(width, height)
    draw = ImageDraw.Draw(img)
    # A large central block (luma 0 = ink) covering > the blank ink threshold.
    draw.rectangle([width // 4, height // 4, 3 * width // 4, 3 * height // 4], fill=0)
    return img


def _bottom_bleed(width: int = 850, height: int = 1100) -> Image.Image:
    """A near-blank page with a solid ink bar along the very bottom edge."""
    img = _blank(width, height)
    draw = ImageDraw.Draw(img)
    # A bar inside the bottom 1% of the page -> falls in the bottom margin band.
    draw.rectangle([0, height - max(1, height // 100), width, height], fill=0)
    return img


# ---------------------------------------------------------------------------
# §7.1 L1 proxies on synthetic PNGs
# ---------------------------------------------------------------------------
class L1ProxyTest(unittest.TestCase):
    def test_blank_page_flagged(self) -> None:
        findings = vqa.check_blank_page(_blank(), page_index=0)
        self.assertEqual(len(findings), 1)
        self.assertEqual(findings[0].check, "visual.blank_page")
        self.assertEqual(findings[0].severity, schema.Severity.WARNING.value)
        self.assertEqual(findings[0].location, "page:1")

    def test_content_page_not_flagged(self) -> None:
        self.assertEqual(vqa.check_blank_page(_centered_content(), page_index=0), [])

    def test_edge_bleed_flagged(self) -> None:
        findings = vqa.check_edge_bleed(_bottom_bleed(), page_index=1)
        self.assertTrue(findings)
        self.assertTrue(all(f.check == "visual.edge_bleed" for f in findings))
        self.assertTrue(all(f.severity == schema.Severity.WARNING.value for f in findings))
        self.assertTrue(any("bottom" in (f.location or "") for f in findings))
        self.assertTrue(any(f.location == "page:2:bottom" for f in findings))

    def test_centered_content_no_bleed(self) -> None:
        self.assertEqual(vqa.check_edge_bleed(_centered_content(), page_index=0), [])

    def test_landscape_dimensions_handled(self) -> None:
        # 1100x850 (landscape): a centered block must still read as clean edges,
        # proving the band is computed off the correct (per-side) dimension.
        landscape = _centered_content(width=1100, height=850)
        self.assertEqual(vqa.check_edge_bleed(landscape, page_index=0), [])
        self.assertEqual(vqa.check_blank_page(landscape, page_index=0), [])
        # And a landscape bottom bleed is still caught.
        bleed = _bottom_bleed(width=1100, height=850)
        findings = vqa.check_edge_bleed(bleed, page_index=0)
        self.assertTrue(any("bottom" in (f.location or "") for f in findings))

    def test_proxies_accept_path_and_image(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            png = Path(td) / "page-1.png"
            _blank().save(png)
            from_image = vqa.check_blank_page(_blank(), page_index=0)
            from_path = vqa.check_blank_page(png, page_index=0)
            self.assertEqual(len(from_image), len(from_path))
            self.assertEqual(from_image[0].check, from_path[0].check)
            self.assertEqual(from_image[0].location, from_path[0].location)

    def test_proxies_never_raise_on_bad_image(self) -> None:
        missing = Path("/nonexistent/page-1.png")
        self.assertEqual(vqa.check_blank_page(missing, page_index=0), [])
        self.assertEqual(vqa.check_edge_bleed(missing, page_index=0), [])
        with tempfile.TemporaryDirectory() as td:
            bogus = Path(td) / "page-1.png"
            bogus.write_text("not a png", encoding="utf-8")
            self.assertEqual(vqa.check_blank_page(bogus, page_index=0), [])
            self.assertEqual(vqa.check_edge_bleed(bogus, page_index=0), [])

    def test_run_visual_l1_concatenates(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            p1 = Path(td) / "page-1.png"
            p2 = Path(td) / "page-2.png"
            _blank().save(p1)           # -> blank_page on page 1
            _bottom_bleed().save(p2)    # -> edge_bleed on page 2
            findings = vqa.run_visual_l1([p1, p2])
            checks = {f.check for f in findings}
            self.assertIn("visual.blank_page", checks)
            self.assertIn("visual.edge_bleed", checks)

    def test_page_count_sane_flags_empty(self) -> None:
        findings = vqa.check_page_count_sane([])
        self.assertEqual(len(findings), 1)
        self.assertEqual(findings[0].check, "visual.no_pages")
        self.assertEqual(
            vqa.check_page_count_sane([_blank()]), []
        )


class RendererAvailabilityTest(unittest.TestCase):
    def test_probe_marks_visual_unavailable_when_binary_probe_fails(self) -> None:
        """PATH presence alone is not enough; broken renderers must degrade."""
        orig_which = doctor.shutil.which
        orig_run = subprocess.run
        try:
            doctor.shutil.which = lambda name: f"/fake/{name}"

            def fake_run(*args, **kwargs):
                return SimpleNamespace(returncode=134, stdout=b"", stderr=b"abort")

            subprocess.run = fake_run
            status = doctor.probe()
        finally:
            doctor.shutil.which = orig_which
            subprocess.run = orig_run

        self.assertFalse(status["binaries"]["soffice"])
        self.assertFalse(status["visual_qa"])

    def test_probe_marks_visual_unavailable_when_conversion_probe_fails(self) -> None:
        """Version commands can pass while headless conversion is unusable."""
        orig_which = doctor.shutil.which
        orig_run = subprocess.run
        calls = []
        try:
            doctor.shutil.which = lambda name: f"/fake/{name}"

            def fake_run(args, *unused_args, **unused_kwargs):
                calls.append(list(args))
                if "--convert-to" in args:
                    return SimpleNamespace(returncode=134, stdout=b"", stderr=b"abort")
                return SimpleNamespace(returncode=0, stdout=b"version ok", stderr=b"")

            subprocess.run = fake_run
            status = doctor.probe()
        finally:
            doctor.shutil.which = orig_which
            subprocess.run = orig_run

        self.assertTrue(status["binaries"]["soffice"])
        self.assertTrue(status["binaries"]["pdftoppm"])
        self.assertFalse(status["visual_qa"])
        self.assertIn("soffice convert failed", status["binary_errors"]["visual_qa"])
        self.assertTrue(any("--convert-to" in call for call in calls))


# ---------------------------------------------------------------------------
# §7.2 Manifest + checklist (model-free)
# ---------------------------------------------------------------------------
def _extract_real_profile(td: Path) -> dict:
    """Extract the repo template into a temp store and return its profile dict."""
    from brandkit.formats.docx import extract as docx_extract
    from brandkit.profile import store

    template = Path(__file__).resolve().parents[1] / "examples" / "templates" / "branddocs_template.docx"
    docx_extract.extract(template, "vqa", scope="project", cwd=td)
    return store.load_profile("vqa", "project", cwd=td).profile


class ManifestTest(unittest.TestCase):
    def test_build_manifest_shape(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            profile = _extract_real_profile(td)
            out_dir = td / "out.visual"
            out_dir.mkdir(parents=True, exist_ok=True)
            png_paths = []
            for i in (1, 2):
                p = out_dir / f"page-{i}.png"
                _centered_content().save(p)
                png_paths.append(p)
            l1 = vqa.run_visual_l1(png_paths)
            manifest_path = vqa.build_visual_manifest(
                profile=profile, document=td / "out.docx", png_paths=png_paths,
                l1_findings=l1, renderers_ok=True, out_dir=out_dir,
            )
            self.assertTrue(manifest_path.is_file())
            data = json.loads(manifest_path.read_text(encoding="utf-8"))
            self.assertEqual(data["schema_version"], "visual-manifest-1")
            self.assertEqual(data["kind"], "docx")
            self.assertEqual(len(data["pages"]), 2)
            self.assertIn("checklist", data)
            self.assertIn("l1_findings", data)
            self.assertIn("instructions", data)
            self.assertTrue(data["renderers_available"])
            # Page records carry dimensions + orientation.
            self.assertEqual(data["pages"][0]["width"], 850)
            self.assertEqual(data["pages"][0]["orientation"], "portrait")

    def test_checklist_derives_from_profile(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            profile = _extract_real_profile(Path(td))
            checklist = vqa.derive_visual_checklist(profile)
            ids = {item["id"] for item in checklist}
            self.assertIn("regions_present", ids)   # from structure.skeleton
            self.assertIn("cover_correct", ids)     # from anchors.cover
            self.assertIn("palette_on_brand", ids)  # from theme.colors
            self.assertIn("roles_styled", ids)      # from roles._index
            self.assertIn("overflow_clean", ids)    # from qa.overflow_capability=render
            # No chart in this template -> charts_rendered must be ABSENT.
            self.assertNotIn("charts_rendered", ids)
            # Every item is traceable.
            for item in checklist:
                self.assertIn("derived_from", item)
                self.assertIn("what", item)

    def test_checklist_includes_charts_when_present(self) -> None:
        profile = {"kind": "docx", "roles": {"_index": ["chart.bar"]}}
        ids = {item["id"] for item in vqa.derive_visual_checklist(profile)}
        self.assertIn("charts_rendered", ids)

    def test_manifest_paths_relative(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            profile = _extract_real_profile(td)
            out_dir = td / "out.visual"
            out_dir.mkdir(parents=True, exist_ok=True)
            png = out_dir / "page-1.png"
            _centered_content().save(png)
            manifest_path = vqa.build_visual_manifest(
                profile=profile, document=td / "out.docx", png_paths=[png],
                l1_findings=[], renderers_ok=True, out_dir=out_dir,
            )
            data = json.loads(manifest_path.read_text(encoding="utf-8"))
            self.assertEqual(data["pages"][0]["png"], "page-1.png")  # relative, no dir

    def test_manifest_degraded_keeps_checklist(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            profile = _extract_real_profile(td)
            out_dir = td / "out.visual"
            manifest_path = vqa.build_visual_manifest(
                profile=profile, document=td / "out.docx", png_paths=[],
                l1_findings=[], renderers_ok=False, out_dir=out_dir,
            )
            data = json.loads(manifest_path.read_text(encoding="utf-8"))
            self.assertTrue(data["degraded"])
            self.assertEqual(data["pages"], [])
            self.assertEqual(data["l1_findings"], [])
            self.assertTrue(data["checklist"])  # still populated


# ---------------------------------------------------------------------------
# §7.3 Degrade and gate wiring (no renderer)
# ---------------------------------------------------------------------------
def _minimal_profile() -> dict:
    prof = schema.build_envelope("docx", {"name": "wire"})
    prof["surface"] = {"docx": {}}
    prof["roles"] = {"_index": []}
    return prof


class GateWiringTest(unittest.TestCase):
    def test_run_qa_fast_never_renders(self) -> None:
        called = {"n": 0}

        def boom(*a, **k):  # pragma: no cover - must never run
            called["n"] += 1
            raise AssertionError("render_to_pngs must not be called on --qa fast")

        orig = vqa.render_to_pngs
        vqa.render_to_pngs = boom
        try:
            with tempfile.TemporaryDirectory() as td:
                target = Path(td) / "out.docx"
                _real_docx(target)  # presence only; never rendered
                report = gate.run_qa(target, _minimal_profile(), qa="fast")
            self.assertEqual(called["n"], 0)
            self.assertFalse(
                any(f.check.startswith("visual.") for f in report.findings)
            )
        finally:
            vqa.render_to_pngs = orig

    def test_run_qa_auto_degrades_when_no_renderers(self) -> None:
        orig = vqa.renderers_available
        vqa.renderers_available = lambda: False
        try:
            with tempfile.TemporaryDirectory() as td:
                target = Path(td) / "out.docx"
                _real_docx(target)
                report = gate.run_qa(target, _minimal_profile(), qa="auto")
            unavailable = [f for f in report.findings if f.check == "visual.unavailable"]
            self.assertEqual(len(unavailable), 1)
            self.assertEqual(unavailable[0].severity, schema.Severity.INFO.value)
            self.assertTrue(report.passed)
            self.assertFalse(
                any(f.severity == schema.Severity.ERROR.value for f in report.findings)
            )
        finally:
            vqa.renderers_available = orig

    def test_run_qa_deep_injected_visual(self) -> None:
        """Drive ``deep`` without soffice via the ``visual=`` injection hook."""
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            profile = _extract_real_profile(td)
            out_dir = td / "out.visual"
            out_dir.mkdir(parents=True, exist_ok=True)
            p1 = out_dir / "page-1.png"
            p2 = out_dir / "page-2.png"
            _blank().save(p1)            # blank page -> L1 finding
            _centered_content().save(p2)
            target = td / "out.docx"
            _real_docx(target)

            report = gate.run_qa(
                target, profile, qa="deep", out_dir=out_dir,
                visual=(True, [p1, p2]),
            )
            # L1 blank-page finding present.
            self.assertTrue(any(f.check == "visual.blank_page" for f in report.findings))
            # Manifest finding present, pointing at a written file.
            manifest = [f for f in report.findings if f.check == "visual.manifest"]
            self.assertEqual(len(manifest), 1)
            self.assertEqual(manifest[0].severity, schema.Severity.INFO.value)
            self.assertTrue(Path(manifest[0].location).is_file())
            # The manifest records both pages and the L1 finding.
            data = json.loads(Path(manifest[0].location).read_text(encoding="utf-8"))
            self.assertEqual(len(data["pages"]), 2)
            self.assertTrue(
                any(lf["check"] == "visual.blank_page" for lf in data["l1_findings"])
            )
            # Still passes (only INFO/WARNING).
            self.assertTrue(report.passed)

    def test_run_qa_auto_injected_visual_runs_l1_no_manifest(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            out_dir = td / "out.visual"
            out_dir.mkdir(parents=True, exist_ok=True)
            p1 = out_dir / "page-1.png"
            _blank().save(p1)
            target = td / "out.docx"
            _real_docx(target)
            report = gate.run_qa(
                target, _minimal_profile(), qa="auto", out_dir=out_dir,
                visual=(True, [p1]),
            )
            self.assertTrue(any(f.check == "visual.blank_page" for f in report.findings))
            # auto does NOT emit a manifest.
            self.assertFalse(any(f.check == "visual.manifest" for f in report.findings))

    def test_visual_findings_never_change_verdict(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            out_dir = td / "out.visual"
            out_dir.mkdir(parents=True, exist_ok=True)
            p1 = out_dir / "page-1.png"
            _blank().save(p1)  # WARNING-producing
            target = td / "out.docx"
            _real_docx(target)
            report = gate.run_qa(
                target, _minimal_profile(), qa="auto", out_dir=out_dir,
                visual=(True, [p1]),
            )
            self.assertTrue(any(f.severity == schema.Severity.WARNING.value
                                for f in report.findings))
            self.assertEqual(
                report.verdict, schema.VerificationStatus.PASSED_WITH_WARNINGS.value
            )
            self.assertNotEqual(report.verdict, schema.VerificationStatus.FAILED.value)
            self.assertTrue(report.passed)

    def test_verify_time_target_none_never_renders(self) -> None:
        called = {"n": 0}
        orig = vqa.render_to_pngs

        def boom(*a, **k):  # pragma: no cover
            called["n"] += 1
            raise AssertionError("no render at verify time (target is None)")

        vqa.render_to_pngs = boom
        try:
            report = gate.run_qa(None, _minimal_profile(), mode="verify", qa="deep")
            self.assertEqual(called["n"], 0)
            self.assertFalse(any(f.check.startswith("visual.") for f in report.findings))
        finally:
            vqa.render_to_pngs = orig


# ---------------------------------------------------------------------------
# §7.4 Backward-compat guard: CI (no renderer) == fast modulo one INFO
# ---------------------------------------------------------------------------
class BackwardCompatTest(unittest.TestCase):
    def test_ci_degrade_matches_fast_modulo_info(self) -> None:
        """CI (no renderer) on --qa auto == --qa fast, modulo one INFO finding.

        The load-bearing invariant is the EXIT CODE: ``report.passed`` (no ERROR)
        is identical, so a green gate stays green and a red one stays red. The
        only delta is a single INFO ``visual.unavailable`` finding (which flips
        the human-readable verdict label from ``passed`` to
        ``passed_with_warnings`` but never the exit code). Crucially, no existing
        test exercises this path: every smoke test passes ``--qa fast``.
        """
        orig = vqa.renderers_available
        vqa.renderers_available = lambda: False
        try:
            with tempfile.TemporaryDirectory() as td:
                target = Path(td) / "out.docx"
                _real_docx(target)
                profile = _minimal_profile()
                fast = gate.run_qa(target, profile, mode="generate", qa="fast")
                auto = gate.run_qa(target, profile, mode="generate", qa="auto")
            # Exit code (passed) is invariant -- the real backward-compat promise.
            self.assertEqual(fast.passed, auto.passed)
            self.assertTrue(auto.passed)
            # No new ERROR is ever introduced by the degrade path.
            self.assertFalse(
                any(f.severity == schema.Severity.ERROR.value for f in auto.findings)
            )
            # The only delta is exactly one INFO visual.unavailable finding.
            extra = [f for f in auto.findings if f not in fast.findings]
            self.assertEqual(len(extra), 1)
            self.assertEqual(extra[0].check, "visual.unavailable")
            self.assertEqual(extra[0].severity, schema.Severity.INFO.value)
        finally:
            vqa.renderers_available = orig


# ---------------------------------------------------------------------------
# §7.5 Gated real end-to-end render (skipped in CI)
# ---------------------------------------------------------------------------
@unittest.skipUnless(vqa.renderers_available(), "no soffice/pdftoppm renderers")
class RealRenderE2ETest(unittest.TestCase):
    def test_render_produces_ordered_pngs(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            template = (
                Path(__file__).resolve().parents[1]
                / "examples" / "templates" / "branddocs_template.docx"
            )
            out_dir = td / "render"
            pngs = vqa.render_to_pngs(template, out_dir)
            self.assertTrue(pngs, "real render produced no PNGs")
            self.assertTrue(all(p.is_file() for p in pngs))
            # Ordered numerically (page-1 before page-2 ... page-10).
            indices = [vqa._page_sort_key(p) for p in pngs]
            self.assertEqual(indices, sorted(indices))

    def test_deep_generate_writes_manifest_and_pngs(self) -> None:
        from brandkit.cli import main

        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            old = Path.cwd()
            import os
            os.chdir(td)
            try:
                template = (
                    Path(__file__).resolve().parents[1]
                    / "examples" / "templates" / "branddocs_template.docx"
                )
                self.assertEqual(
                    main(["extract", "--name", "e2e", "--template", str(template),
                          "--scope", "project"]),
                    0,
                )
                idoc = td / "idoc.json"
                idoc.write_text(json.dumps({
                    "cover": {"title": "Visual Audit Demo"},
                    "blocks": [
                        {"type": "heading", "level": 1, "text": "Section"},
                        {"type": "paragraph", "text": "Rendered for the visual audit."},
                    ],
                }), encoding="utf-8")
                out = td / "out.docx"
                rc = main(["generate", "--name", "e2e", "--input", str(idoc),
                           "--output", str(out), "--scope", "project", "--qa", "deep"])
                self.assertEqual(rc, 0)
                visual_dir = out.parent / "out.visual"
                self.assertTrue((visual_dir / "visual_manifest.json").is_file())
                self.assertTrue((visual_dir / "page-1.png").is_file())
            finally:
                os.chdir(old)


if __name__ == "__main__":
    unittest.main()
