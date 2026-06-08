# SPDX-License-Identifier: MIT
"""Byte-idempotency normalization shared by the docx/pptx/xlsx generators.

A saved OOXML package is a ZIP; python-docx / python-pptx / openpyxl all stamp the
current wall clock into every ZIP entry header at save time, so two otherwise
identical generations differ by a byte whenever the two saves straddle the DOS-time
2-second boundary. :func:`repack_fixed_timestamps` re-zips the package with a FIXED
entry timestamp (the same ``(1980, 1, 1, 0, 0, 0)`` ``ooxml.pack`` uses), preserving
part bytes and order, so re-running the generator yields an identical file.

openpyxl additionally stamps ``docProps/core.xml`` ``dcterms:modified`` at save;
the xlsx generator opts into ``pin_modified_from_created=True`` to pin it to the
package's OWN ``dcterms:created`` (a value derived from the shell, never a
code-literal date). docx/pptx do not need this: pptx pins ``dcterms:modified`` to a
constant before save and docx's ``core.xml`` is already stable, so they call with
the default ``pin_modified_from_created=False`` and keep byte-identical behavior.

A rewrite failure is tolerated (idempotency is a determinism nicety, not a
correctness invariant) and never corrupts the file: the already-saved package
stays in place.
"""

from __future__ import annotations

from pathlib import Path

_FIXED_DATE = (1980, 1, 1, 0, 0, 0)
# A fixed ISO instant used to neutralize the wall-clock ``dcterms`` timestamps a
# library stamps into a NESTED package's ``docProps/core.xml`` (e.g. the xlsx
# workbook python-pptx embeds for a native chart). A FORMAT constant matching the
# pinned ZIP epoch above, never a brand value.
_FIXED_ISO = "1980-01-01T00:00:00Z"
# Parts that are themselves OOXML/zip packages embedded in the outer package. A
# native pptx chart embeds ``ppt/embeddings/*.xlsx`` whose core.xml carries a
# wall-clock ``dcterms:created``/``modified`` - normalized recursively so two
# generations are byte-identical.
_NESTED_PACKAGE_SUFFIXES = (".xlsx", ".docx", ".pptx", ".zip")
# Real OOXML nests one level (an embedded chart workbook); this bounds a
# pathologically self-nesting archive so recursion can never overflow the stack.
_MAX_NEST_DEPTH = 4


def _pin_core_dcterms(text: str, *, both: bool) -> str:
    """Pin ``docProps/core.xml`` dcterms timestamps for byte-idempotency.

    ``both=False`` mirrors the legacy behavior: rewrite ``modified`` to equal the
    package's own ``created`` (openpyxl stamps ``modified`` at save). ``both=True``
    pins BOTH ``created`` and ``modified`` to ``_FIXED_ISO`` - used for a NESTED
    embedded package whose ``created`` is ITSELF a wall-clock value (so there is no
    stable in-package value to borrow).
    """
    import re

    # Timestamp text never contains ``<``, so ``[^<]*`` cannot run past the element
    # (unlike ``.*?``, which on a malformed core.xml could span a mismatched closing
    # tag and corrupt the XML). ``count=1`` pins the single canonical element a
    # well-formed ``core.xml`` carries, never a stray duplicate.
    if both:
        for tag in ("created", "modified"):
            text = re.sub(
                rf"(<dcterms:{tag}[^>]*>)[^<]*(</dcterms:{tag}>)",
                lambda m: m.group(1) + _FIXED_ISO + m.group(2),
                text,
                count=1,
            )
        return text
    created = re.search(r"<dcterms:created[^>]*>([^<]*)</dcterms:created>", text)
    if created:
        text = re.sub(
            r"(<dcterms:modified[^>]*>)[^<]*(</dcterms:modified>)",
            lambda m: m.group(1) + created.group(1) + m.group(2),
            text,
            count=1,
        )
    return text


def _normalize_nested_package(data: bytes, _depth: int = 0) -> bytes:
    """Return an embedded OOXML/zip package re-zipped for byte-idempotency.

    Pins its ``docProps/core.xml`` ``dcterms:created``/``modified`` to a fixed
    instant and re-zips with fixed entry timestamps (recursing into any further
    nested packages). Any failure returns the original bytes unchanged - this is a
    determinism nicety, never a correctness gate. ``_depth`` bounds the recursion
    (real OOXML nests one level - an embedded chart workbook); a pathologically
    self-nesting archive stops at ``_MAX_NEST_DEPTH`` instead of overflowing.
    """
    import io
    import zipfile

    if _depth > _MAX_NEST_DEPTH:
        return data

    try:
        with zipfile.ZipFile(io.BytesIO(data), "r") as zin:
            names = zin.namelist()
            parts = {name: zin.read(name) for name in names}
    except (OSError, zipfile.BadZipFile):
        return data

    core = parts.get("docProps/core.xml")
    if core is not None:
        try:
            parts["docProps/core.xml"] = _pin_core_dcterms(
                core.decode("utf-8"), both=True
            ).encode("utf-8")
        except UnicodeDecodeError:
            pass

    try:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for name in names:
                payload = parts[name]
                if name.endswith(_NESTED_PACKAGE_SUFFIXES):
                    payload = _normalize_nested_package(payload, _depth + 1)
                info = zipfile.ZipInfo(name, date_time=_FIXED_DATE)
                info.compress_type = zipfile.ZIP_DEFLATED
                zout.writestr(info, payload)
        return buf.getvalue()
    except OSError:
        return data


def repack_fixed_timestamps(
    path: Path, *, pin_modified_from_created: bool = False
) -> None:
    """Re-zip ``path`` with fixed entry timestamps so two saves are byte-identical.

    When ``pin_modified_from_created`` is True, also rewrite ``docProps/core.xml``
    ``dcterms:modified`` to equal the package's ``dcterms:created`` (for writers
    like openpyxl that stamp ``modified`` at save). Any NESTED package part (an
    embedded ``.xlsx`` workbook a native chart carries) is normalized recursively
    so its wall-clock ``core.xml`` timestamps do not break idempotency. Part bytes
    and order are otherwise preserved verbatim. A read or rewrite failure is
    tolerated.
    """
    import zipfile

    try:
        with zipfile.ZipFile(path, "r") as zin:
            names = zin.namelist()
            parts = {name: zin.read(name) for name in names}
    except (OSError, zipfile.BadZipFile):
        return

    if pin_modified_from_created:
        core = parts.get("docProps/core.xml")
        if core is not None:
            try:
                parts["docProps/core.xml"] = _pin_core_dcterms(
                    core.decode("utf-8"), both=False
                ).encode("utf-8")
            except UnicodeDecodeError:
                pass  # tolerate a non-UTF-8 core.xml (idempotency is a nicety)

    try:
        tmp = path.with_name(path.name + ".tmp")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            # Preserve the original part order so the central directory is stable.
            for name in names:
                payload = parts[name]
                if name.endswith(_NESTED_PACKAGE_SUFFIXES):
                    payload = _normalize_nested_package(payload)
                info = zipfile.ZipInfo(name, date_time=_FIXED_DATE)
                info.compress_type = zipfile.ZIP_DEFLATED
                zout.writestr(info, payload)
        tmp.replace(path)
    except OSError:
        return
