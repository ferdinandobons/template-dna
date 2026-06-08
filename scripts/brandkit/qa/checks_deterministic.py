# SPDX-License-Identifier: MIT
"""Deterministic L0 checks for M1."""

from __future__ import annotations

import hashlib
from pathlib import Path

from docx import Document
from openpyxl import load_workbook
from pptx import Presentation

from brandkit.common import text as textutil
from brandkit.profile import comprehension as comprehensionmod
from brandkit.profile import schema
from brandkit.profile.reconcile import confidence_clears_floor
from brandkit.qa.model import Finding


def check_profile(profile: dict) -> list[Finding]:
    findings: list[Finding] = []
    for problem in schema.validate(profile):
        findings.append(Finding("schema", schema.Severity.ERROR.value, problem))
    for rid in (profile.get("roles") or {}).get("_index", []):
        entry = profile.get("roles", {}).get(rid)
        if not entry or not entry.get("resolver"):
            findings.append(
                Finding(
                    "every_role_resolves",
                    schema.Severity.ERROR.value,
                    f"{rid} has no resolver",
                )
            )
    return findings


def check_shell_provenance(shell, profile: dict) -> list[Finding]:
    """ERROR when the saved shell no longer matches profile provenance.

    The Brand Profile records ``provenance.shell.sha256`` at extraction time. A
    later hand edit, corrupt copy, or tamper must be load-bearing in QA: generation
    from a drifted shell is not the same brand contract the profile verified.
    """
    if shell is None:
        return []
    recorded = ((profile.get("provenance") or {}).get("shell") or {}).get("sha256")
    if not recorded:
        return []

    shell_path = Path(shell)
    if not shell_path.is_file():
        return [
            Finding(
                "shell_provenance",
                schema.Severity.ERROR.value,
                f"recorded shell hash exists but shell is missing: {shell_path}",
                location=str(shell_path),
            )
        ]

    try:
        actual = _sha256_file(shell_path)
    except OSError as exc:
        return [
            Finding(
                "shell_provenance",
                schema.Severity.ERROR.value,
                f"could not hash shell {shell_path}: {exc}",
                location=str(shell_path),
            )
        ]

    if actual != recorded:
        return [
            Finding(
                "shell_provenance",
                schema.Severity.ERROR.value,
                f"shell hash drifted: recorded {recorded}, actual {actual}",
                location=str(shell_path),
            )
        ]
    return []


def _sha256_file(path: Path, *, chunk: int = 1 << 20) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for block in iter(lambda: fh.read(chunk), b""):
            h.update(block)
    return h.hexdigest()


def check_resolver_targets(shell, profile: dict) -> list[Finding]:
    """Verify every role's resolver target actually exists in the ``shell``.

    This is the deterministic backstop for the core promise ("apply only artifacts
    the profile *proved* exist"): it opens the shell once and confirms each role's
    concrete target is present -
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
    roles = profile.get("roles") or {}
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


# ---------------------------------------------------------------------------
# Comprehension-aware deterministic checks (model-free; bind to captured facts)
# ---------------------------------------------------------------------------
def check_comprehension_targets(profile: dict) -> list[Finding]:
    """FAIL-CLOSED membership of every load-bearing comprehension ref (§3).

    Every ``anchor_ref`` / ``index_ref`` / ``region_ref`` / ``feeds_from_role_id``
    / ``role_annotations`` key must be a verbatim id from the surfaced inventories;
    a ref whose inventory is empty/absent is itself an ERROR (this is the SOLE gate
    for anchor/index/region refs, so it must reject, never skip). No-ops when the
    comprehension is absent (status != present), keeping the model-free CI path and
    pptx/xlsx green. Delegates to the single membership definition so the gate and
    the merge writer can never disagree.
    """
    comp = profile.get("comprehension")
    if (
        not isinstance(comp, dict)
        or comp.get("status") != schema.ComprehensionStatus.PRESENT.value
    ):
        return []
    findings: list[Finding] = []
    for problem in comprehensionmod.check_membership(profile, comp):
        findings.append(
            Finding("comprehension_targets_exist", schema.Severity.ERROR.value, problem)
        )
    return findings


def _present_comprehension(profile: dict) -> dict | None:
    comp = profile.get("comprehension")
    if (
        isinstance(comp, dict)
        and comp.get("status") == schema.ComprehensionStatus.PRESENT.value
    ):
        return comp
    return None


def captured_template_texts(
    profile: dict, *, include_surface_prompts: bool = False
) -> list[str]:
    """Collect every demo/placeholder text the extractor captured for THIS template.

    Model-free and language-agnostic: the comparison strings come from the
    template's own captured facts (cover ``demo_value`` s, surfaced demo-region
    markers/text), never a fixed phrase in any language. Surface placeholder
    prompts are opt-in because some formats keep master/layout prompts inside the
    package even when they are not visibly rendered; L0 must not fail on those.
    """
    texts: list[str] = []
    comp = _present_comprehension(profile)
    if comp is not None:
        for slot in (comp.get("cover_slots") or {}).values():
            if isinstance(slot, dict) and slot.get("demo_value"):
                texts.append(str(slot["demo_value"]))
    kind = profile.get("kind")
    sub = ((profile.get("surface") or {}).get(kind) or {}) if kind else {}
    if isinstance(sub, dict):
        demo = sub.get("demo_region") or {}
        if isinstance(demo, dict):
            for m in demo.get("instruction_markers") or []:
                if m:
                    texts.append(str(m))
            if demo.get("start_text"):
                texts.append(str(demo["start_text"]))
        if not include_surface_prompts:
            return _dedup_texts(texts)
        for anchor in sub.get("cover_anchors") or []:
            if not isinstance(anchor, dict):
                continue
            for key in ("placeholder", "demo_value"):
                value = anchor.get(key)
                if value:
                    texts.append(str(value))
    return _dedup_texts(texts)


def _dedup_texts(texts: list[str]) -> list[str]:
    # De-dup preserving order.
    seen: set[str] = set()
    out: list[str] = []
    for t in texts:
        if t and t not in seen:
            seen.add(t)
            out.append(t)
    return out


def _captured_demo_texts(profile: dict) -> list[str]:
    return captured_template_texts(profile)


def check_residual_template_text(text: str, profile: dict) -> list[Finding]:
    """Generalized ``no_residual_template_text`` for ALL kinds (model-free).

    Scans the produced text for any demo/placeholder string captured at extract
    for THIS template (cover ``demo_value`` s + surfaced demo markers/text). No
    hardcoded phrases, language-agnostic, identical across formats.
    """
    findings: list[Finding] = []
    for marker in _captured_demo_texts(profile):
        if marker in text:
            findings.append(
                Finding(
                    "no_residual_template_text",
                    schema.Severity.ERROR.value,
                    f"residual template text: {marker!r}",
                )
            )
    return findings


def check_no_orphan_cover_placeholder(text: str, profile: dict) -> list[Finding]:
    """No bound cover slot still shows its captured ``demo_value`` in the output.

    A slot the comprehension marked ``fill_rule='in_place'`` (i.e. it should have
    been filled) whose captured ``demo_value`` still appears in the produced text
    is an ERROR (stale demo prompt left behind). A slot the model intentionally
    re-armed (``fill_rule='clear'``) is INFO. No-ops when comprehension is absent.
    """
    comp = _present_comprehension(profile)
    if comp is None:
        return []
    findings: list[Finding] = []
    for anchor_ref, slot in (comp.get("cover_slots") or {}).items():
        if not isinstance(slot, dict):
            continue
        demo_value = slot.get("demo_value")
        if not demo_value or str(demo_value) not in text:
            continue
        fill_rule = slot.get("fill_rule")
        if fill_rule == schema.FillRule.IN_PLACE.value:
            findings.append(
                Finding(
                    "no_orphan_cover_placeholder",
                    schema.Severity.ERROR.value,
                    f"cover slot {anchor_ref!r} still shows placeholder {demo_value!r}",
                )
            )
        else:
            findings.append(
                Finding(
                    "no_orphan_cover_placeholder",
                    schema.Severity.INFO.value,
                    f"cover slot {anchor_ref!r} retains re-armed prompt {demo_value!r}",
                )
            )
    return findings


def check_index_matches_content(
    present_seq_ids: set[str], profile: dict
) -> list[Finding]:
    """No preserved CAPTION index lacks corresponding captionable content (WARNING).

    Ruling A: the comparison keys on the DETERMINISTIC ``seq_id`` captured opaquely
    from the index's ``\\c`` field switch (plan §7), NEVER on the advisory ``kind``
    token. ``present_seq_ids`` is the set of caption SEQ identifiers the generator
    actually emitted captions for (each is a verbatim ``seq_id`` it matched a
    preserved index against). A preserved CAPTION index (one carrying a non-null
    ``seq_id``) that is kept (``reconcile`` not ``clear``) yet whose ``seq_id`` was
    not emitted is flagged WARNING - a stale caption index the reconciliation should
    have cleared. An outline TOC (``seq_id`` is null) is never flagged: it is always
    refreshed, not content-matched. No-ops when comprehension is absent.
    Generator-driven (the generator knows which SEQ classes it emitted); model-free.
    """
    comp = _present_comprehension(profile)
    if comp is None:
        return []
    findings: list[Finding] = []
    for idx in (comp.get("conventions") or {}).get("indexes") or []:
        if not isinstance(idx, dict):
            continue
        if idx.get("reconcile") == schema.Reconcile.CLEAR.value:
            continue  # the reconciliation already removed it on purpose
        seq_id = idx.get("seq_id")
        # Only a CAPTION index (a \c-switched index, identified by its opaque
        # seq_id) is content-matched; a bare outline TOC has no seq_id and is
        # always refreshed, never flagged.
        if seq_id and seq_id not in present_seq_ids:
            findings.append(
                Finding(
                    "index_matches_content",
                    schema.Severity.WARNING.value,
                    f"preserved caption index {idx.get('index_ref')!r} "
                    f"(seq_id {seq_id!r}) has no matching captionable content",
                )
            )
    return findings


def check_no_net_structure_loss(
    removed_refs: set[str],
    profile: dict,
    *,
    confidence: float | None = None,
) -> list[Finding]:
    """The destructive-action floor (§6): no preserved anchor/index removed unless
    the deterministic layer independently classified it placeholder/demo.

    ``removed_refs`` is the set of anchor/index refs the reconciliation actually
    deleted (supplied by the generator). For each, the comprehension must carry a
    deterministically-corroborated destructive verdict (``fill_rule='clear'`` for a
    cover anchor, ``reconcile='clear'`` for an index, or a ``verdict='demo'`` region);
    otherwise the deletion is an ERROR. Model-free (it reads the frozen verdicts).

    ``confidence`` (optional) is the model's single ``comprehension.confidence``, the
    SAME value the reconcile sites gate on. When supplied, the backstop also
    re-verifies the destructive-action confidence floor: a sanctioned removal whose
    confidence does NOT clear the floor is an ERROR, because under the uniform policy
    such a removal should have been downgraded to KEEP at the reconcile site (a wrong
    delete is unrecoverable). When ``confidence`` is ``None`` the floor re-check is
    skipped (additive, back-compatible) and only the verdict-corroboration gate runs.
    """
    comp = _present_comprehension(profile)
    findings: list[Finding] = []
    cleared_anchors = (
        {
            ref
            for ref, slot in (comp.get("cover_slots") or {}).items()
            if isinstance(slot, dict)
            and slot.get("fill_rule") == schema.FillRule.CLEAR.value
        }
        if comp
        else set()
    )
    cleared_indexes = (
        {
            idx.get("index_ref")
            for idx in ((comp.get("conventions") or {}).get("indexes") or [])
            if isinstance(idx, dict)
            and idx.get("reconcile") == schema.Reconcile.CLEAR.value
        }
        if comp
        else set()
    )
    demo_regions = (
        {
            r.get("region_ref")
            for r in ((comp.get("demo_classification") or {}).get("regions") or [])
            if isinstance(r, dict) and r.get("verdict") == schema.Verdict.DEMO.value
        }
        if comp
        else set()
    )
    sanctioned = cleared_anchors | cleared_indexes | demo_regions
    floor_blocks = confidence is not None and not confidence_clears_floor(confidence)
    for ref in sorted(removed_refs):
        if ref not in sanctioned:
            findings.append(
                Finding(
                    "no_net_structure_loss",
                    schema.Severity.ERROR.value,
                    f"reconciliation removed {ref!r} without a corroborated "
                    f"destructive verdict",
                )
            )
        elif floor_blocks:
            # The verdict is sanctioned, but the model's confidence is below the
            # destructive floor: under the uniform policy the reconcile site should
            # have downgraded this to KEEP. A removal that still happened is a floor
            # breach (a wrong delete is unrecoverable), so surface it LOUDLY.
            findings.append(
                Finding(
                    "no_net_structure_loss",
                    schema.Severity.ERROR.value,
                    f"reconciliation removed {ref!r} below the destructive "
                    f"confidence floor ({confidence:.2f})",
                )
            )
    return findings


# ---------------------------------------------------------------------------
# Shell-vs-output structural diffs (catch silent corruption the text scans miss)
# ---------------------------------------------------------------------------
def _xlsx_formula_map(path) -> dict[str, str]:
    """Map ``Sheet!Coord`` -> formula string for every formula cell in a workbook.

    Reads the live file (``data_only=False``) so the comparison reflects what is
    actually on disk, not a possibly-stale catalog snapshot. Only string cells
    starting with ``=`` are formulas; iterating ``_cells`` keeps this O(populated)
    on sparse corporate models.
    """
    out: dict[str, str] = {}
    wb = load_workbook(path, data_only=False)
    for ws in wb.worksheets:
        for cell in ws._cells.values():
            value = cell.value
            if isinstance(value, str) and value.startswith("="):
                out[f"{ws.title}!{cell.coordinate}"] = value
    return out


def check_formula_preservation(shell, output, profile: dict) -> list[Finding]:
    """ERROR on any formula the generation LOST or MUTATED vs the brand shell.

    The brand guarantee for xlsx is that the shell's own formulas are preserved
    verbatim (only the shell's inputs are refilled, never its formulas). A region
    fill that overwrites a formula cell is silent data corruption that no text
    scan can see, so this diffs the shell's ``address->formula`` set against the
    output's and emits an ERROR for every entry that disappeared or changed.

    Model-free and deterministic. No-ops (returns ``[]``) when either file is
    absent (e.g. verify time, when there is no output yet) so it is always safe to
    call from the gate. The extractor's ``artifact_catalog.formulas`` is the same
    baseline; here we read the live shell so the check stands even if the catalog
    drifts.
    """
    if shell is None or output is None:
        return []
    if profile.get("kind") != schema.Kind.XLSX.value:
        return []
    try:
        shell_formulas = _xlsx_formula_map(shell)
        output_formulas = _xlsx_formula_map(output)
    except Exception as exc:  # opening a workbook must never crash the gate
        return [
            Finding(
                "formula_preservation",
                schema.Severity.WARNING.value,
                f"could not verify formula preservation: {exc}",
            )
        ]
    findings: list[Finding] = []
    for address in sorted(shell_formulas):
        shell_formula = shell_formulas[address]
        out_formula = output_formulas.get(address)
        if out_formula is None:
            findings.append(
                Finding(
                    "formula_preservation",
                    schema.Severity.ERROR.value,
                    f"shell formula at {address} ({shell_formula!r}) was erased in the output",
                    location=address,
                )
            )
        elif out_formula != shell_formula:
            findings.append(
                Finding(
                    "formula_preservation",
                    schema.Severity.ERROR.value,
                    f"shell formula at {address} was mutated: "
                    f"{shell_formula!r} -> {out_formula!r}",
                    location=address,
                )
            )
    return findings


def _docx_component_counts(path) -> dict[str, int]:
    # Only count components the generator preserves rather than re-authors. Tables
    # are a meaningful survival signal (a flattened table is a real defect); raw
    # list-paragraph counts are NOT compared, because docx generation rewrites the
    # body from new content, so a different list length is expected, not a loss.
    doc = Document(path)
    return {"tables": len(doc.tables)}


def _pptx_component_counts(path) -> dict[str, int]:
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    prs = Presentation(path)
    tables = charts = pictures = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            if getattr(shape, "has_table", False):
                tables += 1
            if getattr(shape, "has_chart", False):
                charts += 1
            if getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.PICTURE:
                pictures += 1
    return {"tables": tables, "charts": charts, "pictures": pictures}


def _xlsx_component_counts(path) -> dict[str, int]:
    wb = load_workbook(path, data_only=False)
    tables = charts = 0
    for ws in wb.worksheets:
        try:
            tables += len(ws.tables)
        except Exception:
            pass
        try:
            charts += len(ws._charts)
        except Exception:
            pass
    return {"tables": tables, "charts": charts}


_COMPONENT_COUNTERS = {
    schema.Kind.DOCX.value: _docx_component_counts,
    schema.Kind.PPTX.value: _pptx_component_counts,
    schema.Kind.XLSX.value: _xlsx_component_counts,
}


def check_component_survival(shell, output, profile: dict) -> list[Finding]:
    """WARN when a native component present in the shell has no counterpart in the output.

    Catches the silent down-render class (e.g. pptx flattening a native table to
    text, or a docx list losing its numbering) that no text scan detects: it
    counts native components (tables/charts/lists/pictures, per format) in the
    shell and in the output and emits a WARNING for every component family whose
    output count dropped below the shell's. A WARNING (not ERROR) because some
    drops are legitimate (the new content genuinely has fewer of a component); the
    signal is "you may have lost a native object", surfaced rather than silent.

    Model-free, deterministic, no-ops when either file is absent.
    """
    if shell is None or output is None:
        return []
    counter = _COMPONENT_COUNTERS.get(profile.get("kind"))
    if counter is None:
        return []
    try:
        shell_counts = counter(shell)
        output_counts = counter(output)
    except Exception as exc:  # opening the package must never crash the gate
        return [
            Finding(
                "component_survival",
                schema.Severity.WARNING.value,
                f"could not verify component survival: {exc}",
            )
        ]
    findings: list[Finding] = []
    for family in sorted(shell_counts):
        before = shell_counts[family]
        after = output_counts.get(family, 0)
        if before > 0 and after < before:
            findings.append(
                Finding(
                    "component_survival",
                    schema.Severity.WARNING.value,
                    f"native {family} count dropped {before} -> {after} "
                    f"between shell and output (possible down-render)",
                )
            )
    return findings


def check_docx(path, profile: dict, shell=None) -> list[Finding]:
    findings = check_profile(profile)
    doc = Document(path)
    text = "\n".join(
        [p.text for p in doc.paragraphs]
        + [cell.text for t in doc.tables for row in t.rows for cell in row.cells]
    )
    for hit in textutil.find_markdown_literals(text):
        findings.append(
            Finding(
                "no_literal_markdown",
                schema.Severity.ERROR.value,
                f"literal markdown leaked: {hit['match']!r}",
            )
        )
    findings.extend(check_residual_template_text(text, profile))
    findings.extend(check_no_orphan_cover_placeholder(text, profile))
    findings.extend(check_component_survival(shell, path, profile))
    return findings


def check_pptx(path, profile: dict, shell=None) -> list[Finding]:
    findings = check_profile(profile)
    prs = Presentation(path)
    text = "\n".join(
        shape.text
        for slide in prs.slides
        for shape in slide.shapes
        if hasattr(shape, "text")
    )
    for hit in textutil.find_markdown_literals(text):
        findings.append(
            Finding(
                "no_literal_markdown",
                schema.Severity.ERROR.value,
                f"literal markdown leaked: {hit['match']!r}",
            )
        )
    # Uniform, model-free residual check: compare against THIS template's captured
    # demo/placeholder text, never a fixed phrase (de-literalized).
    findings.extend(check_residual_template_text(text, profile))
    findings.extend(check_no_orphan_cover_placeholder(text, profile))
    findings.extend(check_component_survival(shell, path, profile))
    return findings


def check_xlsx(path, profile: dict, shell=None) -> list[Finding]:
    findings = check_profile(profile)
    wb = load_workbook(path, data_only=False)
    cell_texts: list[str] = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    cell_texts.append(cell.value)
                    for hit in textutil.find_markdown_literals(cell.value):
                        findings.append(
                            Finding(
                                "no_literal_markdown",
                                schema.Severity.ERROR.value,
                                f"literal markdown leaked: {hit['match']!r}",
                            )
                        )
    # Uniform, model-free residual check across all kinds.
    text = "\n".join(cell_texts)
    findings.extend(check_residual_template_text(text, profile))
    findings.extend(check_no_orphan_cover_placeholder(text, profile))
    # Deterministic structural diffs that the text scan above is blind to: a fill
    # that erased a shell formula, or a native component lost in the output.
    findings.extend(check_formula_preservation(shell, path, profile))
    findings.extend(check_component_survival(shell, path, profile))
    return findings
