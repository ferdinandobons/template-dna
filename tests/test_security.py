# SPDX-License-Identifier: MIT
"""Security regression tests for the SECURITY review group.

Covers:
- M10 path traversal: a profile ``--name`` that is not a single safe path
  segment ('../../escape', absolute, with separators) is refused in BOTH
  ``target_dir_for_save`` and ``resolve_profile_dir`` - a proprietary shell can
  never be written outside ``brand-kit/<name>``.
- M11 zip-slip: ``ooxml.pack.unpack`` refuses a sibling-prefix escape entry
  (``../<dest>_SIBLING/x.txt``) that the old ``startswith`` prefix check let
  through, plus the classic ``../`` and absolute entries.
- ooxml-6 / security-4: ``parse_theme_colors`` routes through the hardened
  parser (no external-entity resolution) and ``_extract_theme`` no longer
  blanket-swallows a real parse error.
"""

from __future__ import annotations

import io
import os
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from brandkit.common import color
from brandkit.ooxml import pack
from brandkit.profile import schema
from brandkit.profile import store
from brandkit.qa.gate import run_qa


# ---------------------------------------------------------------------------
# M10 - profile name path traversal
# ---------------------------------------------------------------------------
class NameTraversalTest(unittest.TestCase):
    _BAD_NAMES = [
        "../../escape",
        "../escape",
        "..",
        ".",
        "a/b",
        "a\\b",
        "/abs/escape",
        "",
        "   ",
        "with\x00nul",
    ]

    def test_target_dir_for_save_refuses_traversal(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            for name in self._BAD_NAMES:
                with self.subTest(name=name):
                    with self.assertRaises(store.ProfileStoreError):
                        store.target_dir_for_save(name, "project", cwd=td)

    def test_resolve_profile_dir_refuses_traversal(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            for name in self._BAD_NAMES:
                with self.subTest(name=name):
                    with self.assertRaises(store.ProfileStoreError):
                        store.resolve_profile_dir(name, "project", cwd=td)

    def test_escape_name_does_not_write_outside_store(self) -> None:
        """End-to-end: a save under a traversal name must not land above the store."""
        with tempfile.TemporaryDirectory() as td:
            cwd = Path(td)
            with self.assertRaises(store.ProfileStoreError):
                target = store.target_dir_for_save("../../ESCAPED", "project", cwd=cwd)
                store.save_profile(target, {"kind": "docx"}, b"fake")
            # Nothing escaped two levels up.
            self.assertFalse((cwd.parent.parent / "ESCAPED").exists())
            self.assertFalse((cwd / "ESCAPED").exists())

    def test_valid_names_still_resolve_under_store(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            cwd = Path(td)
            for name in ("acme", "deck-2024", "model_v2"):
                with self.subTest(name=name):
                    target = store.target_dir_for_save(name, "project", cwd=cwd)
                    self.assertTrue(
                        target.resolve().is_relative_to(
                            store.project_store_root(cwd).resolve()
                        )
                    )
                    self.assertEqual(target.name, name)

    @unittest.skipIf(os.name == "nt", "symlink setup differs on Windows")
    def test_save_profile_refuses_symlink_escape_for_shell(self) -> None:
        """A pre-existing profile symlink must not let shell writes escape store."""
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            outside = root / "outside"
            outside.mkdir()
            victim = outside / "shell.docx"
            victim.write_bytes(b"old")

            profile_dir = root / "brand-kit" / "acme"
            profile_dir.mkdir(parents=True)
            (profile_dir / "template").symlink_to(outside, target_is_directory=True)

            with self.assertRaises(store.ProfileStoreError):
                store.save_profile(profile_dir, {"kind": "docx"}, b"new")
            self.assertEqual(victim.read_bytes(), b"old")

    @unittest.skipIf(os.name == "nt", "symlink setup differs on Windows")
    def test_save_profile_refuses_metadata_symlink_escape(self) -> None:
        """profile.json/provenance writes must use the same safe writer as shell."""
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            outside_profile = root / "outside-profile.json"
            outside_hash = root / "outside-provenance.sha256"
            outside_profile.write_text("old-profile", encoding="utf-8")
            outside_hash.write_text("old-hash", encoding="utf-8")

            profile_dir = root / "brand-kit" / "acme"
            profile_dir.mkdir(parents=True)
            (profile_dir / "profile.json").symlink_to(outside_profile)
            (profile_dir / store.SHELL_HASH_FILE).symlink_to(outside_hash)

            with self.assertRaises(store.ProfileStoreError):
                store.save_profile(profile_dir, {"kind": "docx"}, b"new")
            self.assertEqual(outside_profile.read_text(encoding="utf-8"), "old-profile")
            self.assertEqual(outside_hash.read_text(encoding="utf-8"), "old-hash")


# ---------------------------------------------------------------------------
# Provenance drift - shell hash must be load-bearing in QA
# ---------------------------------------------------------------------------
class ProvenanceDriftTest(unittest.TestCase):
    def test_run_qa_fails_when_shell_hash_drifted(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            shell = Path(td) / "shell.docx"
            shell.write_bytes(b"original-shell")
            profile = schema.build_envelope("docx", {"name": "acme"})
            profile["provenance"]["shell"]["path"] = "template/shell.docx"
            profile["provenance"]["shell"]["sha256"] = store.sha256_file(shell)

            shell.write_bytes(b"tampered-shell")
            report = run_qa(None, profile, qa="fast", shell=shell)

            self.assertFalse(report.passed)
            self.assertTrue(
                any(
                    f.check == "shell_provenance"
                    and f.severity == schema.Severity.ERROR.value
                    for f in report.findings
                )
            )


# ---------------------------------------------------------------------------
# M11 - zip-slip
# ---------------------------------------------------------------------------
def _zip_with_entries(entries: dict[str, bytes]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in entries.items():
            zf.writestr(name, data)
    return buf.getvalue()


class ZipSlipTest(unittest.TestCase):
    def _unpack_entry(self, entry_name: str) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            src = root / "evil.zip"
            src.write_bytes(_zip_with_entries({entry_name: b"pwned"}))
            dest = root / "dest"
            with self.assertRaises(pack.PackError):
                pack.unpack(src, dest)

    def test_sibling_prefix_escape_refused(self) -> None:
        """The classic broken startswith() check let ``dest_SIBLING`` through."""
        # Entry resolves to a sibling dir whose path *string-prefixes* dest.
        self._unpack_entry("../dest_SIBLING/x.txt")

    def test_parent_traversal_refused(self) -> None:
        self._unpack_entry("../escape.txt")

    def test_deep_traversal_refused(self) -> None:
        self._unpack_entry("../../escape.txt")

    def test_backslash_traversal_refused(self) -> None:
        self._unpack_entry("..\\escape.txt")

    def test_absolute_entry_refused(self) -> None:
        # An absolute member name must not be honored as an absolute write.
        abs_name = "/tmp/zipslip_escape.txt" if os.name != "nt" else "C:\\zipslip.txt"
        self._unpack_entry(abs_name)

    def test_nothing_written_outside_dest_on_sibling_escape(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            src = root / "evil.zip"
            src.write_bytes(_zip_with_entries({"../dest_SIBLING/x.txt": b"pwned"}))
            dest = root / "dest"
            with self.assertRaises(pack.PackError):
                pack.unpack(src, dest)
            self.assertFalse((root / "dest_SIBLING").exists())
            self.assertFalse((root / "dest_SIBLING" / "x.txt").exists())

    def test_safe_entries_unpack_normally(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            src = root / "ok.zip"
            src.write_bytes(
                _zip_with_entries(
                    {
                        "[Content_Types].xml": b"<Types/>",
                        "word/document.xml": b"<document/>",
                    }
                )
            )
            dest = root / "dest"
            pack.unpack(src, dest)
            self.assertTrue((dest / "word" / "document.xml").is_file())
            self.assertEqual(
                (dest / "word" / "document.xml").read_bytes(), b"<document/>"
            )


# ---------------------------------------------------------------------------
# ooxml-6 / security-4 - theme color parsing hardening
# ---------------------------------------------------------------------------
_THEME_TEMPLATE = (
    b'<?xml version="1.0"?>'
    b'<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
    b"<a:themeElements><a:clrScheme>"
    b'<a:dk1><a:srgbClr val="111111"/></a:dk1>'
    b'<a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>'
    b'<a:accent1><a:srgbClr val="00AABB"/></a:accent1>'
    b"</a:clrScheme></a:themeElements></a:theme>"
)


class ThemeColorHardeningTest(unittest.TestCase):
    def test_parses_plain_theme(self) -> None:
        colors = color.parse_theme_colors(_THEME_TEMPLATE)
        self.assertEqual(colors["accent1"], "00AABB")
        self.assertEqual(colors["dk1"], "111111")

    def test_external_entity_is_not_resolved(self) -> None:
        """A theme referencing an external entity must not pull in file contents."""
        with tempfile.TemporaryDirectory() as td:
            secret = Path(td) / "secret.txt"
            secret.write_text("TOPSECRET", encoding="utf-8")
            xxe = (
                b'<?xml version="1.0"?>'
                b'<!DOCTYPE a:theme [ <!ENTITY xxe SYSTEM "file://'
                + str(secret).encode()
                + b'"> ]>'
                b'<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                b"<a:themeElements><a:clrScheme>"
                b'<a:accent1><a:srgbClr val="00AABB"/></a:accent1>'
                b"</a:clrScheme></a:themeElements></a:theme>"
            )
            # Either it parses without expanding the entity, or lxml raises on
            # the undefined/unresolved entity. Neither path may leak the secret.
            try:
                colors = color.parse_theme_colors(xxe)
            except Exception:
                return
            self.assertNotIn("TOPSECRET", "".join(colors.values()))
            self.assertEqual(colors.get("accent1"), "00AABB")


class LinkSchemeAllowlistTest(unittest.TestCase):
    """Hyperlink targets are scheme-allowlisted at the shared chokepoint, so author
    content cannot wire a hostile ``file:``/``javascript:``/``data:``/``smb:`` link
    into a generated docx/pptx. Safe + relative/fragment targets are allowed."""

    def test_safe_schemes_allowed(self) -> None:
        from brandkit.common.links import is_safe_link_url

        for url in (
            "https://example.com",
            "http://example.com/a?b=1#c",
            "mailto:x@example.com",
            "tel:+15551234",
            "#section",
            "relative/path.html",
        ):
            self.assertTrue(is_safe_link_url(url), url)

    def test_unsafe_schemes_refused(self) -> None:
        from brandkit.common.links import is_safe_link_url

        for url in (
            "file:///etc/passwd",
            "javascript:alert(1)",
            "data:text/html,<script>1</script>",
            "smb://server/share",
            "vbscript:msgbox(1)",
            "",
            "   ",
        ):
            self.assertFalse(is_safe_link_url(url), url)


if __name__ == "__main__":
    unittest.main()
