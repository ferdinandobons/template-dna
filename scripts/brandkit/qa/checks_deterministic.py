# SPDX-License-Identifier: MIT
"""Deterministic L0 checks for M1."""
from __future__ import annotations

from pathlib import Path

from docx import Document
from openpyxl import load_workbook
from pptx import Presentation

from brandkit.common import text as textutil
from brandkit.profile import schema
from brandkit.qa.model import Finding


def check_profile(profile: dict) -> list[Finding]:
    findings: list[Finding] = []
    for problem in schema.validate(profile):
        findings.append(Finding("schema", schema.Severity.ERROR.value, problem))
    for rid in (profile.get("roles") or {}).get("_index", []):
        entry = profile.get("roles", {}).get(rid)
        if not entry or not entry.get("resolver"):
            findings.append(Finding("every_role_resolves", schema.Severity.ERROR.value, f"{rid} has no resolver"))
    return findings


def check_resolver_targets(shell, profile: dict) -> list[Finding]:
    """Verify every role's resolver target actually exists in the ``shell``.

    This is the deterministic backstop for the core promise ("apply only artifacts
    the profile *proved* exist"): it opens the shell once and confirms each role's
    concrete target is present —
      - docx: ``style_id`` or ``style_name`` ∈ ``doc.styles``;
      - pptx: ``layout`` ∈ the presentation's slide-layout names;
      - xlsx: ``name`` ∈ the workbook's defined names.
    A role whose target is missing yields an ERROR ``resolver_targets_exist``
    finding. ``shell`` may be a path or None; when None the check is skipped (the
    caller is responsible for supplying the shell at verify time).
    """
    if shell is None:
        return []
    kind = profile.get("kind")
    roles = (profile.get("roles") or {})
    role_ids = [r for r in roles if r != "_index"]
    findings: list[Finding] = []
    try:
        if kind == schema.Kind.DOCX.value:
            present = _docx_style_keys(shell)
            for rid in role_ids:
                resolver = (roles.get(rid) or {}).get("resolver") or {}
                if resolver.get("type") not in (
                    schema.ResolverType.NAMED_STYLE.value,
                    None,
                ):
                    continue
                sid = resolver.get("style_id")
                sname = resolver.get("style_name")
                if not sid and not sname:
                    continue
                if (sid and sid in present) or (sname and sname in present):
                    continue
                findings.append(
                    Finding(
                        "resolver_targets_exist",
                        schema.Severity.ERROR.value,
                        f"role {rid!r} target style {(sid or sname)!r} not found in shell",
                    )
                )
        elif kind == schema.Kind.PPTX.value:
            layout_names = _pptx_layout_names(shell)
            for rid in role_ids:
                resolver = (roles.get(rid) or {}).get("resolver") or {}
                if resolver.get("type") != schema.ResolverType.PLACEHOLDER.value:
                    continue
                layout = resolver.get("layout")
                if layout is None or layout in layout_names:
                    continue
                findings.append(
                    Finding(
                        "resolver_targets_exist",
                        schema.Severity.ERROR.value,
                        f"role {rid!r} layout {layout!r} not found in shell (have {sorted(layout_names)})",
                    )
                )
        elif kind == schema.Kind.XLSX.value:
            defined = _xlsx_defined_names(shell)
            for rid in role_ids:
                resolver = (roles.get(rid) or {}).get("resolver") or {}
                if resolver.get("type") != schema.ResolverType.NAMED_RANGE.value:
                    continue
                name = resolver.get("name")
                if name is None or name in defined:
                    continue
                findings.append(
                    Finding(
                        "resolver_targets_exist",
                        schema.Severity.ERROR.value,
                        f"role {rid!r} named range {name!r} not found in shell (have {sorted(defined)})",
                    )
                )
    except Exception as exc:  # opening the shell must never crash the gate
        findings.append(
            Finding(
                "resolver_targets_exist",
                schema.Severity.WARNING.value,
                f"could not verify resolver targets against shell: {exc}",
            )
        )
    return findings


def _docx_style_keys(shell) -> set:
    doc = Document(shell)
    keys: set = set()
    for style in doc.styles:
        sid = getattr(style, "style_id", None)
        if sid:
            keys.add(sid)
        name = getattr(style, "name", None)
        if name:
            keys.add(name)
    return keys


def _pptx_layout_names(shell) -> set:
    prs = Presentation(shell)
    return {layout.name for layout in prs.slide_layouts}


def _xlsx_defined_names(shell) -> set:
    wb = load_workbook(shell, data_only=False)
    try:
        return set(wb.defined_names.keys())
    except AttributeError:
        # Older openpyxl exposes defined_names as a list-like of DefinedName.
        return {dn.name for dn in wb.defined_names}


def check_docx(path, profile: dict) -> list[Finding]:
    findings = check_profile(profile)
    doc = Document(path)
    text = "\n".join([p.text for p in doc.paragraphs] + [cell.text for t in doc.tables for row in t.rows for cell in row.cells])
    for hit in textutil.find_markdown_literals(text):
        findings.append(
            Finding(
                "no_literal_markdown",
                schema.Severity.ERROR.value,
                f"literal markdown leaked: {hit['match']!r}",
            )
        )
    demo = ((profile.get("surface") or {}).get("docx") or {}).get("demo_region") or {}
    for marker in demo.get("instruction_markers") or []:
        if marker and marker in text:
            findings.append(
                Finding("no_residual_template_text", schema.Severity.ERROR.value, f"residual template text: {marker!r}")
            )
    return findings


def check_pptx(path, profile: dict) -> list[Finding]:
    findings = check_profile(profile)
    prs = Presentation(path)
    text = "\n".join(shape.text for slide in prs.slides for shape in slide.shapes if hasattr(shape, "text"))
    for hit in textutil.find_markdown_literals(text):
        findings.append(Finding("no_literal_markdown", schema.Severity.ERROR.value, f"literal markdown leaked: {hit['match']!r}"))
    if "Example slide instructions" in text:
        findings.append(Finding("no_residual_template_text", schema.Severity.ERROR.value, "residual template slide instructions"))
    return findings


def check_xlsx(path, profile: dict) -> list[Finding]:
    findings = check_profile(profile)
    wb = load_workbook(path, data_only=False)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    for hit in textutil.find_markdown_literals(cell.value):
                        findings.append(Finding("no_literal_markdown", schema.Severity.ERROR.value, f"literal markdown leaked: {hit['match']!r}"))
    return findings
