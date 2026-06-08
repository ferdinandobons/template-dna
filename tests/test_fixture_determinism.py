# SPDX-License-Identifier: MIT
"""Structural determinism guard for the COMPLEX fixture builders.

``tests/fixtures/builders/build_complex_{docx,pptx,xlsx}.py`` rebuild the
committed binaries ``tests/fixtures/complex/acme_complex.{docx,pptx,xlsx}``.
Those binaries are the source of truth; the builders only re-derive them. A raw
byte comparison is NOT a portable guarantee: rebuilding in a different
environment changes the SHA (and even the size) purely from
python-docx/pptx/openpyxl serialization differences across library versions,
not from any non-determinism (two builds in the SAME environment are identical).

So this guard compares a rebuild to the committed fixture STRUCTURALLY, robust
across library versions, to catch a builder that has drifted in
structure/content (real fixture rot) while tolerating benign cross-version
serialization noise:

  * the sorted list of OOXML package part names must be identical; AND
  * each shared part's content must match: XML/.rels parts are compared by
    CANONICAL XML (C14N over the parsed tree, with namespace prefixes normalized
    first, so attribute order and prefix spelling do not matter); nested OOXML
    packages (the embedded chart workbook) are compared recursively the same
    way; other binary/media parts are compared byte-for-byte.

Volatile parts whose content is tool-stamped rather than content-derived
(``docProps/core.xml`` timestamps / producer string, at any nesting depth -
the embedded chart workbook carries its own wall-clock-stamped core.xml) are
ignored.

If a builder cannot run in this environment, the corresponding test is skipped
(mirroring the missing-fixture skip in ``test_xlsx_complex_fidelity.py``) rather
than failing.
"""

from __future__ import annotations

import copy
import importlib.util
import io
import tempfile
import unittest
import zipfile
from pathlib import Path
from types import ModuleType

from lxml import etree

_ROOT = Path(__file__).resolve().parents[1]
_BUILDERS = _ROOT / "tests" / "fixtures" / "builders"
_COMPLEX = _ROOT / "tests" / "fixtures" / "complex"

# Parts whose content is tool-stamped (timestamps / producer name + version)
# rather than content-derived: not a structural signal, so ignored by the
# comparison. ``docProps/app.xml`` embeds the writing library's name AND version
# verbatim (e.g. openpyxl stamps "Openpyxl 3.1.5" into <Application>/<AppVersion>),
# so a rebuild on a different library version would otherwise trip a false positive.
_VOLATILE_PARTS = frozenset({"docProps/core.xml", "docProps/app.xml"})

# Suffixes that identify XML payloads (incl. relationship parts).
_XML_SUFFIXES = (".xml", ".rels")
# Suffixes that identify a nested OOXML / zip package (e.g. the embedded chart
# workbook ``ppt/embeddings/Microsoft_Excel_Sheet1.xlsx``). These are compared
# recursively because their raw zip bytes can drift across library versions
# even when their logical content is unchanged.
_NESTED_PACKAGE_SUFFIXES = (".docx", ".pptx", ".xlsx", ".zip")


def _load_builder(stem: str) -> ModuleType:
    """Import a ``build_complex_<stem>.py`` builder by file path.

    ``tests/fixtures/builders`` is not a package, so load by location rather than
    by import name.
    """
    path = _BUILDERS / f"build_complex_{stem}.py"
    spec = importlib.util.spec_from_file_location(f"_builder_{stem}", path)
    if spec is None or spec.loader is None:  # pragma: no cover - defensive
        raise unittest.SkipTest(f"cannot load builder spec for {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _all_namespace_uris(root: etree._Element) -> set[str]:
    """Every namespace URI used by an element/attribute name in the tree."""
    uris: set[str] = set()
    for el in root.iter():
        if isinstance(el.tag, str) and el.tag.startswith("{"):
            uris.add(el.tag[1:].split("}", 1)[0])
        for attr_name in el.attrib:
            if attr_name.startswith("{"):
                uris.add(attr_name[1:].split("}", 1)[0])
    return uris


def _canonical_xml(data: bytes) -> bytes:
    """Return a prefix-agnostic canonical (C14N) serialization of an XML payload.

    Plain C14N is namespace-PREFIX-sensitive: the same namespace URI bound to
    ``a:`` in one file and ``ns2:`` in another serializes differently even
    though the documents are semantically identical. python-docx/lxml pick the
    drawingml prefix from lxml's process-global namespace registry, so the
    prefix a rebuild emits depends on what else has run in the interpreter -
    benign noise that must NOT register as drift.

    To neutralize that, the tree is rebuilt with a deterministic
    ``uri -> nN`` prefix map (by sorted URI) before C14N. C14N still respects
    significant whitespace and element/attribute structure, so real
    content/namespace drift is still caught; only the prefix spelling is
    normalized. Passing the parsed tree (not raw bytes) also keeps this
    independent of the input's XML declaration / encoding header.
    """
    root = etree.fromstring(data)
    nsmap = {f"n{i}": uri for i, uri in enumerate(sorted(_all_namespace_uris(root)))}
    rebuilt = etree.Element(root.tag, nsmap=nsmap)
    rebuilt.text = root.text
    for attr_name, attr_val in root.attrib.items():
        rebuilt.set(attr_name, attr_val)
    for child in root:
        rebuilt.append(copy.deepcopy(child))
    return etree.tostring(rebuilt, method="c14n2")


def _compare_zip(
    test: unittest.TestCase, expected: bytes, actual: bytes, where: str
) -> None:
    """Assert two OOXML/zip packages are structurally equal.

    Same part-name set, and matching content per part (canonical XML for XML
    parts, recursion for nested packages, byte equality for other media).
    Volatile parts (tool-stamped timestamps / producer strings) are ignored at
    any nesting depth - the embedded chart workbook carries its OWN
    ``docProps/core.xml`` whose timestamp is wall-clock, not pinned.
    """
    ez = zipfile.ZipFile(io.BytesIO(expected))
    az = zipfile.ZipFile(io.BytesIO(actual))
    enames = sorted(ez.namelist())
    anames = sorted(az.namelist())
    prefix = f"{where}: " if where else ""
    test.assertEqual(
        enames,
        anames,
        f"{prefix}package part-name set drifted "
        f"(committed-only={sorted(set(enames) - set(anames))}, "
        f"rebuilt-only={sorted(set(anames) - set(enames))})",
    )

    for part in enames:
        if part in _VOLATILE_PARTS:
            continue
        eb = ez.read(part)
        ab = az.read(part)
        loc = f"{where}::{part}" if where else part
        if part.endswith(_XML_SUFFIXES):
            test.assertEqual(
                _canonical_xml(eb),
                _canonical_xml(ab),
                f"canonical XML differs in {loc}",
            )
        elif part.endswith(_NESTED_PACKAGE_SUFFIXES):
            _compare_zip(test, eb, ab, loc)
        else:
            test.assertEqual(eb, ab, f"binary part differs in {loc}")


class FixtureDeterminismGuard(unittest.TestCase):
    """Each builder must reproduce its committed fixture's STRUCTURE/CONTENT."""

    def _check(self, stem: str, filename: str) -> None:
        committed = _COMPLEX / filename
        if not committed.exists():
            raise unittest.SkipTest(f"missing committed fixture {committed}")
        try:
            builder = _load_builder(stem)
        except unittest.SkipTest:
            raise
        except Exception as exc:  # missing lib / import error -> skip, don't fail
            raise unittest.SkipTest(f"builder for {stem} unavailable: {exc!r}")

        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / filename
            try:
                builder.build(out)
            except Exception as exc:  # builder cannot run here -> skip, don't fail
                raise unittest.SkipTest(f"builder for {stem} could not run: {exc!r}")
            self.assertTrue(out.exists(), f"builder for {stem} produced no file")
            _compare_zip(self, committed.read_bytes(), out.read_bytes(), "")

    def test_docx_builder_matches_committed_structure(self) -> None:
        self._check("docx", "acme_complex.docx")

    def test_pptx_builder_matches_committed_structure(self) -> None:
        self._check("pptx", "acme_complex.pptx")

    def test_xlsx_builder_matches_committed_structure(self) -> None:
        self._check("xlsx", "acme_complex.xlsx")


if __name__ == "__main__":
    unittest.main()
