# SPDX-License-Identifier: MIT
from __future__ import annotations

from dataclasses import dataclass, field

from brandkit.profile import schema

# Frozen registry of every check id the deterministic L0 library
# (``qa/checks_deterministic.py``) may stamp on a Finding. Check ids are a
# PERSISTED contract: they key ``generation_report.json`` digests, the
# cross-run regression multiset, and the ``learn`` distillation, so an id is
# append-only (never rename one in place; see CONVENTIONS.md section 9 for the
# naming convention new ids must follow). ``tests/test_check_registry.py``
# AST-scans the module and fails on any Finding whose id is not listed here,
# catching silent typos before they reach a report.
CHECK_REGISTRY: frozenset[str] = frozenset(
    {
        "appearance_geometry_targets",
        "appearance_numbering_targets",
        "appearance_table_targets",
        "appearance_targets_exist",
        "audit_targets_exist",
        "blend_shell_provenance",
        "color_token_targets_exist",
        "component_survival",
        "comprehension_targets_exist",
        "every_role_resolves",
        "formula_preservation",
        "index_matches_content",
        "no_literal_markdown",
        "no_net_structure_loss",
        "no_orphan_cover_placeholder",
        "no_residual_template_text",
        "override_applied",
        "override_targets_exist",
        "package_integrity",
        "palette_alias_targets_exist",
        "resolver_targets_exist",
        "schema",
        "shell_provenance",
        "triage_targets_exist",
    }
)


@dataclass
class Finding:
    check: str
    severity: str
    message: str
    location: str | None = None


@dataclass
class QAReport:
    verdict: str
    findings: list[Finding] = field(default_factory=list)

    @property
    def passed(self) -> bool:
        return not any(f.severity == schema.Severity.ERROR.value for f in self.findings)
