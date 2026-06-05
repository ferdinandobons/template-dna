# SPDX-License-Identifier: MIT
"""Deterministic L0 checks for M1."""
from __future__ import annotations

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
