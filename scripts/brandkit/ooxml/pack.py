# SPDX-License-Identifier: MIT
"""Pure ``zipfile`` + ``lxml`` (un)packing of OOXML packages.

An OOXML file (``.docx`` / ``.pptx`` / ``.xlsx``) is a ZIP whose entries are XML
and media "parts". This module is the engine's self-contained replacement for
the Anthropic ``office/unpack.py`` + ``office/pack.py`` scripts - re-implemented,
not vendored. It does four things and nothing else:

- :func:`unpack` - explode a package into a directory tree.
- :func:`pack` - repack a directory tree into a valid OOXML file, preserving the
  ZIP entry ordering and the ``[Content_Types].xml`` / ``mimetype`` conventions.
- :func:`read_part` - read one part's raw bytes without a full unpack.
- :func:`list_parts` - list the part names in a package.

Plus small ``lxml`` conveniences (:func:`parse_xml_bytes`, :func:`serialize_xml`)
and namespace helpers so the rest of the engine never re-imports raw lxml glue.

**Determinism.** :func:`pack` writes entries in a stable order (``[Content_Types]
.xml`` first, then ``_rels/.rels``, then the remaining parts sorted) and zeroes
the per-entry timestamp, so packing the same tree twice yields a byte-identical
archive - a hard requirement for the "generate() twice is byte-identical" rule.
The shell, however, is copied **verbatim** elsewhere (never round-tripped) to
preserve themes/numbering/fonts exactly; this module is for parts we deliberately
rewrite.
"""

from __future__ import annotations

import os
import shutil
import zipfile
from pathlib import Path
from typing import Union

from lxml import etree

PathLike = Union[str, os.PathLike]

# ---------------------------------------------------------------------------
# Common OOXML namespaces, registered once for serialization niceties.
# ---------------------------------------------------------------------------
NAMESPACES: dict[str, str] = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "ssml": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
}

CONTENT_TYPES_PART = "[Content_Types].xml"
ROOT_RELS_PART = "_rels/.rels"
MIMETYPE_PART = "mimetype"  # ODF-style; rare in OOXML but handled if present.


class PackError(RuntimeError):
    """Raised when a package cannot be unpacked or packed."""


def qn(prefix_tag: str) -> str:
    """Return the Clark-notation qualified name for a ``prefix:tag`` string.

    Example: ``qn("w:p")`` -> ``"{http://...wordprocessingml...}p"``. Raises
    ``KeyError`` for an unknown prefix (use :data:`NAMESPACES`).
    """
    prefix, _, tag = prefix_tag.partition(":")
    if not tag:  # bare local name, no prefix
        return prefix_tag
    return f"{{{NAMESPACES[prefix]}}}{tag}"


# ---------------------------------------------------------------------------
# lxml conveniences
# ---------------------------------------------------------------------------
def parse_xml_bytes(data: bytes) -> etree._Element:
    """Parse XML bytes into an lxml element (whitespace preserved).

    Uses a parser that does NOT strip insignificant whitespace, because OOXML is
    whitespace-sensitive inside ``w:t``/``a:t`` runs (``xml:space="preserve"``).
    """
    parser = etree.XMLParser(remove_blank_text=False, resolve_entities=False)
    return etree.fromstring(data, parser=parser)


def serialize_xml(
    element: etree._Element, *, xml_declaration: bool = True, standalone: bool = True
) -> bytes:
    """Serialize an lxml element back to OOXML-flavoured bytes (UTF-8).

    Emits the standalone XML declaration OOXML parts expect. No pretty-printing
    (pretty-printing would inject whitespace text nodes that corrupt runs).
    """
    return etree.tostring(
        element,
        xml_declaration=xml_declaration,
        encoding="UTF-8",
        standalone=standalone,
    )


# ---------------------------------------------------------------------------
# unpack / pack
# ---------------------------------------------------------------------------
def unpack(src: PathLike, dest_dir: PathLike, *, overwrite: bool = True) -> Path:
    """Explode an OOXML package at ``src`` into ``dest_dir``.

    Args:
        src: path to a ``.docx``/``.pptx``/``.xlsx`` (any ZIP-based OOXML).
        dest_dir: directory to create and populate with the part tree. Parent
            directories are created as needed.
        overwrite: if True (default) and ``dest_dir`` exists, it is removed
            first; if False and it exists non-empty, raises :class:`PackError`.

    Returns:
        The ``Path`` to ``dest_dir`` (populated).

    Raises:
        PackError: if ``src`` is not a valid ZIP/OOXML package, or ``dest_dir``
            exists and ``overwrite`` is False.
    """
    src_path = Path(src)
    dest = Path(dest_dir)
    if dest.exists():
        if not overwrite and any(dest.iterdir()):
            raise PackError(f"destination already exists and is non-empty: {dest}")
        if overwrite:
            shutil.rmtree(dest)
    dest.mkdir(parents=True, exist_ok=True)
    dest_resolved = dest.resolve()
    try:
        with zipfile.ZipFile(src_path, "r") as zf:
            for name in zf.namelist():
                # Defend against zip-slip: reject entries that escape dest.
                # Reject any '..' segment up front (covers backslash-style
                # separators that Path won't split on POSIX), then confirm the
                # resolved target is contained in dest via relative_to (a real
                # containment check, not a broken string-prefix one that a
                # sibling dir like ``dest_SIBLING`` would defeat).
                if ".." in name.replace("\\", "/").split("/"):
                    raise PackError(f"unsafe zip entry path: {name!r}")
                target = (dest / name).resolve()
                try:
                    target.relative_to(dest_resolved)
                except ValueError:
                    raise PackError(f"unsafe zip entry path: {name!r}")
                if name.endswith("/"):
                    target.mkdir(parents=True, exist_ok=True)
                    continue
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(zf.read(name))
    except zipfile.BadZipFile as exc:
        raise PackError(f"{src_path} is not a valid OOXML/ZIP package") from exc
    return dest


def _ordered_part_names(names: list[str]) -> list[str]:
    """Return ``names`` in the canonical OOXML write order (deterministic).

    Order: an uncompressed ``mimetype`` (if present) MUST be first; then
    ``[Content_Types].xml``; then ``_rels/.rels``; then every remaining part
    sorted lexicographically. This satisfies the OOXML readers that expect
    Content_Types early and gives byte-stable output.
    """
    names_set = set(names)
    ordered: list[str] = []
    for special in (MIMETYPE_PART, CONTENT_TYPES_PART, ROOT_RELS_PART):
        if special in names_set:
            ordered.append(special)
            names_set.discard(special)
    ordered.extend(sorted(names_set))
    return ordered


def pack(src_dir: PathLike, dest: PathLike, *, compresslevel: int = 6) -> Path:
    """Repack a part tree at ``src_dir`` into a valid OOXML file at ``dest``.

    Walks every file under ``src_dir`` (recursively), writes them into a fresh
    ZIP in canonical order with a fixed timestamp (deterministic output). A
    ``mimetype`` part, if present, is stored UNCOMPRESSED and first (ODF
    convention; harmless and correct when present).

    Args:
        src_dir: the unpacked part tree (e.g. from :func:`unpack`).
        dest: output ``.docx``/``.pptx``/``.xlsx`` path (parent dirs created).
        compresslevel: DEFLATE level for normal parts.

    Returns:
        The ``Path`` to ``dest``.

    Raises:
        PackError: if ``src_dir`` does not exist or contains no parts.
    """
    src = Path(src_dir)
    if not src.is_dir():
        raise PackError(f"source part tree does not exist: {src}")
    out = Path(dest)
    out.parent.mkdir(parents=True, exist_ok=True)

    # Collect relative part names (POSIX separators inside the archive).
    names: list[str] = []
    for path in src.rglob("*"):
        if path.is_file():
            names.append(path.relative_to(src).as_posix())
    if not names:
        raise PackError(f"no parts found under {src}")

    fixed_date = (1980, 1, 1, 0, 0, 0)  # deterministic ZIP timestamp
    with zipfile.ZipFile(
        out, "w", zipfile.ZIP_DEFLATED, compresslevel=compresslevel
    ) as zf:
        for name in _ordered_part_names(names):
            data = (src / name).read_bytes()
            info = zipfile.ZipInfo(name, date_time=fixed_date)
            if name == MIMETYPE_PART:
                info.compress_type = zipfile.ZIP_STORED
            else:
                info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o600 << 16
            zf.writestr(info, data)
    return out


# ---------------------------------------------------------------------------
# Read / list without a full unpack
# ---------------------------------------------------------------------------
def read_part(src: PathLike, part_name: str) -> bytes:
    """Read a single part's raw bytes from a package without unpacking it.

    Args:
        src: the OOXML package.
        part_name: the archive-internal part name, POSIX style
            (e.g. ``"word/theme/theme1.xml"``). A leading ``/`` is tolerated.

    Returns:
        The raw bytes of that part.

    Raises:
        PackError: if the package is invalid.
        KeyError: if the part does not exist in the package.
    """
    name = part_name.lstrip("/")
    try:
        with zipfile.ZipFile(Path(src), "r") as zf:
            return zf.read(name)
    except zipfile.BadZipFile as exc:
        raise PackError(f"{src} is not a valid OOXML/ZIP package") from exc


def has_part(src: PathLike, part_name: str) -> bool:
    """Return True if ``part_name`` exists in the package."""
    name = part_name.lstrip("/")
    try:
        with zipfile.ZipFile(Path(src), "r") as zf:
            return name in zf.namelist()
    except zipfile.BadZipFile as exc:
        raise PackError(f"{src} is not a valid OOXML/ZIP package") from exc


def list_parts(src: PathLike) -> list[str]:
    """Return every part name in the package (directories excluded), sorted.

    Raises:
        PackError: if the package is invalid.
    """
    try:
        with zipfile.ZipFile(Path(src), "r") as zf:
            return sorted(n for n in zf.namelist() if not n.endswith("/"))
    except zipfile.BadZipFile as exc:
        raise PackError(f"{src} is not a valid OOXML/ZIP package") from exc


def read_part_xml(src: PathLike, part_name: str) -> etree._Element:
    """Read a part and parse it as XML in one call (see :func:`read_part`)."""
    return parse_xml_bytes(read_part(src, part_name))


def copy_verbatim(src: PathLike, dest: PathLike) -> Path:
    """Copy an OOXML file byte-for-byte to ``dest`` (the shell-copy primitive).

    The Brand Profile keeps the template shell byte-identical; this is the one
    sanctioned way to materialize it. Parent directories are created.
    """
    out = Path(dest)
    out.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(Path(src), out)
    return out
