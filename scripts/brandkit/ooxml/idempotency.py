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


def repack_fixed_timestamps(
    path: Path, *, pin_modified_from_created: bool = False
) -> None:
    """Re-zip ``path`` with fixed entry timestamps so two saves are byte-identical.

    When ``pin_modified_from_created`` is True, also rewrite ``docProps/core.xml``
    ``dcterms:modified`` to equal the package's ``dcterms:created`` (for writers
    like openpyxl that stamp ``modified`` at save). Part bytes and order are
    otherwise preserved verbatim. A read or rewrite failure is tolerated.
    """
    import zipfile

    try:
        with zipfile.ZipFile(path, "r") as zin:
            names = zin.namelist()
            parts = {name: zin.read(name) for name in names}
    except (OSError, zipfile.BadZipFile):
        return

    if pin_modified_from_created:
        import re

        core = parts.get("docProps/core.xml")
        if core is not None:
            text = core.decode("utf-8")
            created = re.search(r"<dcterms:created[^>]*>(.*?)</dcterms:created>", text)
            if created:
                text = re.sub(
                    r"(<dcterms:modified[^>]*>).*?(</dcterms:modified>)",
                    lambda m: m.group(1) + created.group(1) + m.group(2),
                    text,
                )
                parts["docProps/core.xml"] = text.encode("utf-8")

    try:
        tmp = path.with_name(path.name + ".tmp")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            # Preserve the original part order so the central directory is stable.
            for name in names:
                info = zipfile.ZipInfo(name, date_time=_FIXED_DATE)
                info.compress_type = zipfile.ZIP_DEFLATED
                zout.writestr(info, parts[name])
        tmp.replace(path)
    except OSError:
        return
