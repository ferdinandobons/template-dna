# SPDX-License-Identifier: MIT
"""The ONE writer of the canonical ``comprehension`` block (Ruling B).

The model *proposes* a comprehension JSON; this module *disposes*:

1. :func:`surface_inventories` extracts the deterministic, format-uniform
   inventories the model reasons over and the validator binds to - the SAME
   function the ``comprehend-input`` CLI verb and the ``comprehension_targets_exist``
   QA check call, so the model and the gate can never disagree about what ids
   exist. Each inventory entry is ``{"id": <verbatim id>, ...}``.

2. :func:`merge` validates a model-authored block against the profile (schema
   shape + fail-closed membership of every load-bearing ref) and, only on a
   clean pass, writes it into ``profile["comprehension"]`` with sorted keys /
   stable list order, stamps ``generated_by`` + ``source_shell_sha256`` (bound to
   the live ``provenance.shell.sha256``), sets ``status='present'``, and DERIVES
   the additive sinks ``roles[*].usage`` / ``structure.skeleton`` / ``anchors.*``
   from it. On any finding it refuses to write the understanding and instead
   records ``status='rejected'`` with the findings so the model can retry.

The merge boundary is structurally incapable of writing a ``roles[*].resolver``
or a ``surface`` id, so it can never widen the brand guarantee. Comprehension is
frozen here and never re-invoked at generate time (idempotency, §6).
"""
from __future__ import annotations

from typing import Any, Optional

from brandkit.profile import schema


# ---------------------------------------------------------------------------
# Inventory surfacing (the single definition the model AND the validator use)
# ---------------------------------------------------------------------------
def _entries_with_ids(seq: Any) -> list[dict]:
    """Normalize a surfaced list into ``[{"id": str, ...}]`` entries.

    Accepts a list of dicts that already carry an ``"id"`` (or ``"name"``) field,
    or a list/dict of plain strings/keys. Anything without a derivable id is
    dropped (it cannot be a binding target). The result preserves source order.
    """
    out: list[dict] = []
    if isinstance(seq, dict):
        # A mapping ``{id: descriptor}`` (e.g. surface.<kind>.named_regions).
        for key, val in seq.items():
            if isinstance(key, str) and key:
                entry = {"id": key}
                if isinstance(val, dict):
                    entry.update({k: v for k, v in val.items() if k != "id"})
                out.append(entry)
        return out
    if isinstance(seq, list):
        for item in seq:
            if isinstance(item, dict):
                rid = item.get("id") or item.get("name")
                if isinstance(rid, str) and rid:
                    entry = dict(item)
                    entry["id"] = rid
                    out.append(entry)
            elif isinstance(item, str) and item:
                out.append({"id": item})
    return out


def surface_inventories(profile: dict) -> dict:
    """Return the format-uniform inventories every comprehension ref binds to.

    The shape is identical across formats (only which deterministic facts back
    each list differs, per plan §4):

        {
          "cover_anchors": [ {"id": <anchor_ref>, ...}, ... ],
          "fields":        [ {"id": <index_ref>, ...}, ... ],
          "regions":       [ {"id": <region_ref>, ...}, ... ],
          "roles":         [ <role_id>, ... ],
        }

    ``cover_anchors`` / ``fields`` / ``regions`` are read from
    ``surface.<kind>.{cover_anchors,fields,regions}`` when the extractor has
    enriched them (docx after M-i-3; pptx/xlsx may leave them empty until their
    own fact-enrichment milestone, which is legal - the refs into them simply
    have nothing to bind to and a destructive ref is fail-closed at QA time).
    ``roles`` is the concrete role-id list (``_index`` order if present).

    This is the ONLY place the inventory shape is defined; the ``comprehend-input``
    verb and ``comprehension_targets_exist`` both call it so they cannot drift.
    """
    kind = profile.get("kind")
    surface = profile.get("surface") or {}
    sub = surface.get(kind) if isinstance(surface, dict) and kind else {}
    if not isinstance(sub, dict):
        sub = {}
    return {
        "cover_anchors": _entries_with_ids(sub.get("cover_anchors")),
        "fields": _entries_with_ids(sub.get("fields")),
        "regions": _entries_with_ids(sub.get("regions")),
        "roles": list(schema.list_role_ids(profile)),
    }


def comprehend_input_bundle(profile: dict, *, excerpt_chars: int = 8000) -> dict:
    """Build the bounded ``{facts, excerpt}`` bundle the model reasons over.

    ``facts`` is the surfaced inventory plus the most relevant ``surface`` /
    ``artifact_catalog`` slices; ``excerpt`` is an ordered, length-capped sample
    of text the extractor already collected (never raw OOXML). The agent reads
    THIS bundle, never the template.
    """
    inventories = surface_inventories(profile)
    kind = profile.get("kind")
    catalog = profile.get("artifact_catalog") or {}

    facts = {
        "kind": kind,
        "schema_version": profile.get("schema_version"),
        "inventories": inventories,
        "structure": profile.get("structure") or {},
        "anchors": profile.get("anchors") or {},
        "styles": _catalog_styles(catalog),
    }

    excerpt = _collect_excerpt(profile, catalog, excerpt_chars)
    return {"facts": facts, "excerpt": excerpt}


def _catalog_styles(catalog: Any) -> dict:
    """Return the catalog's style inventory under whichever key the extractor used.

    The docx/pptx catalogs write ``styles``; the xlsx catalog writes
    ``named_styles`` (a flat name list). This shared reader accepts either so
    ``facts.styles`` is populated for every format - previously it was ``{}`` for
    every workbook because only the docx/pptx key was read (CC-1 / Q6).
    """
    if not isinstance(catalog, dict):
        return {}
    styles = catalog.get("styles")
    if styles:
        return styles if isinstance(styles, dict) else {"styles": list(styles)}
    named = catalog.get("named_styles")
    if named:
        # A flat name list (xlsx) is surfaced under a stable key so the model can
        # reason over the brand's named cell styles.
        return {"named_styles": list(named)}
    return {}


def _cell_excerpt_text(cell: Any) -> Optional[str]:
    """Extract the most informative text from a non-empty-cell catalog entry.

    Each entry is ``{"address", "data_type", "style", "number_format", ...}`` and
    MAY carry a textual ``value``/``text`` once the xlsx extractor records it. We
    prefer the cell's own text; absent that we fall back to its address so the
    model at least sees which cells are populated (geometry), rather than nothing.
    """
    if isinstance(cell, dict):
        for key in ("text", "value"):
            v = cell.get(key)
            if isinstance(v, str) and v:
                return v
        addr = cell.get("address")
        if isinstance(addr, str) and addr:
            return addr
        return None
    if cell:
        return str(cell)
    return None


def _collect_excerpt(profile: dict, catalog: dict, cap: int) -> list[str]:
    """Collect an ordered, length-capped list of text samples (format-uniform)."""
    samples: list[str] = []
    if isinstance(catalog, dict):
        for s in catalog.get("paragraph_samples") or []:
            if isinstance(s, dict) and s.get("text"):
                samples.append(str(s["text"]))
        for slide in catalog.get("slides") or []:
            for t in (slide.get("texts") if isinstance(slide, dict) else None) or []:
                if t:
                    samples.append(str(t))
        # Top-level cells (legacy/flat catalog shape).
        for c in catalog.get("non_empty_cells") or []:
            txt = _cell_excerpt_text(c)
            if txt:
                samples.append(txt)
        # The xlsx catalog nests cells under each sheet; descend so every workbook
        # yields a real excerpt instead of [] (CC-1 / Q6). Sheets are walked in a
        # stable order so the cap truncates deterministically.
        sheets = catalog.get("sheets")
        if isinstance(sheets, dict):
            for sheet_name in sorted(sheets):
                sheet = sheets[sheet_name]
                if not isinstance(sheet, dict):
                    continue
                for c in sheet.get("non_empty_cells") or []:
                    txt = _cell_excerpt_text(c)
                    if txt:
                        samples.append(txt)
    # Length-cap deterministically by truncating the ordered list.
    out: list[str] = []
    total = 0
    for s in samples:
        if total + len(s) > cap:
            break
        out.append(s)
        total += len(s)
    return out


# ---------------------------------------------------------------------------
# Membership validation (fail-closed) - the deterministic-validation contract
# ---------------------------------------------------------------------------
def check_membership(profile: dict, comp: dict) -> list[str]:
    """Return fail-closed membership problems for a comprehension block.

    Every load-bearing ref must be a verbatim id from the surfaced inventories;
    a ref whose target inventory is EMPTY or ABSENT is itself an error (this is
    the sole gate for anchor/index/region refs, so it must reject, never skip -
    unlike the namespace-guarded resolver consistency check). Executor enums are
    validated by the schema; here we bind ids.

    Returns ``[]`` when ``comp`` is absent / status != present (nothing to bind).
    """
    if not isinstance(comp, dict):
        return []
    status = comp.get("status")
    # Only a PRESENT (or status-less, mid-merge trial) block carries load-bearing
    # refs to enforce. ABSENT and REJECTED both carry no enforceable refs: absent is
    # today's deterministic path, and a rejected block is rebuilt (its findings are
    # already recorded) so re-binding it would surface stale duplicate errors.
    if status not in (None, schema.ComprehensionStatus.PRESENT.value):
        return []
    inv = surface_inventories(profile)
    anchor_ids = {e["id"] for e in inv["cover_anchors"]}
    field_ids = {e["id"] for e in inv["fields"]}
    region_ids = {e["id"] for e in inv["regions"]}
    role_ids = set(inv["roles"])

    problems: list[str] = []

    # (a) cover_slots keys ∈ cover-anchor inventory (FAIL-CLOSED on empty).
    for anchor_ref, slot in (comp.get("cover_slots") or {}).items():
        if anchor_ref not in anchor_ids:
            problems.append(
                f"comprehension.cover_slots: anchor_ref {anchor_ref!r} not in "
                f"surfaced cover_anchors inventory {sorted(anchor_ids)}"
            )
        # (d) binds_to is a content-slot key, advisory; feeds nothing structural.

    conventions = comp.get("conventions") or {}
    # (b) indexes index_ref ∈ field inventory; feeds_from_role_id ∈ roles.
    for i, idx in enumerate(conventions.get("indexes") or []):
        if not isinstance(idx, dict):
            continue
        ref = idx.get("index_ref")
        if ref not in field_ids:
            problems.append(
                f"comprehension.conventions.indexes[{i}].index_ref: {ref!r} not in "
                f"surfaced fields inventory {sorted(field_ids)}"
            )
        feeds = idx.get("feeds_from_role_id")
        if feeds is not None and feeds not in role_ids:
            problems.append(
                f"comprehension.conventions.indexes[{i}].feeds_from_role_id: "
                f"{feeds!r} not in roles {sorted(role_ids)}"
            )

    # (c) sections region_ref ∈ region inventory.
    for i, sec in enumerate(conventions.get("sections") or []):
        if not isinstance(sec, dict):
            continue
        ref = sec.get("region_ref")
        if ref not in region_ids:
            problems.append(
                f"comprehension.conventions.sections[{i}].region_ref: {ref!r} not in "
                f"surfaced regions inventory {sorted(region_ids)}"
            )

    # (c) demo_classification region_ref ∈ region inventory.
    for i, reg in enumerate((comp.get("demo_classification") or {}).get("regions") or []):
        if not isinstance(reg, dict):
            continue
        ref = reg.get("region_ref")
        if ref not in region_ids:
            problems.append(
                f"comprehension.demo_classification.regions[{i}].region_ref: {ref!r} "
                f"not in surfaced regions inventory {sorted(region_ids)}"
            )

    # (d) role_annotations keys ∈ roles.
    for rid in (comp.get("role_annotations") or {}):
        if rid not in role_ids:
            problems.append(
                f"comprehension.role_annotations: role id {rid!r} not in roles "
                f"{sorted(role_ids)}"
            )

    return problems


# ---------------------------------------------------------------------------
# Merge (the only writer)
# ---------------------------------------------------------------------------
class MergeResult:
    """Outcome of a :func:`merge` attempt.

    Attributes:
        ok: True when the comprehension was written ``present``.
        status: the resulting ``comprehension.status`` (``present`` | ``rejected``).
        problems: the validation findings (empty iff ``ok``).
    """

    __slots__ = ("ok", "status", "problems")

    def __init__(self, ok: bool, status: str, problems: list[str]):
        self.ok = ok
        self.status = status
        self.problems = problems


def merge(
    profile: dict,
    comp: dict,
    *,
    generated_by: Optional[dict] = None,
) -> MergeResult:
    """Validate ``comp`` against ``profile`` and, if clean, write it in (Ruling B).

    Mutates ``profile`` in place: on success ``profile['comprehension']`` is the
    canonical block (sorted keys / stable list order, ``status='present'``,
    ``source_shell_sha256`` = live ``provenance.shell.sha256``) and the derived
    sinks (``roles[*].usage`` / ``structure.skeleton`` / ``anchors.*``) are
    refreshed from it. On any finding NOTHING load-bearing is written: the block
    becomes ``status='rejected'`` carrying the findings, so the model must retry.

    Args:
        profile: the loaded profile dict (mutated in place).
        comp: the model-authored comprehension JSON (may omit ``status`` - it is
            forced to ``present`` on a clean merge).
        generated_by: ``{"model","prompt_version","generated_at"}`` provenance to
            stamp; optional.

    Returns:
        A :class:`MergeResult`.
    """
    # 1) Shape validation: run the FULL schema validator on a trial profile that
    # carries this comprehension, so the same shape rules apply as on load.
    trial = dict(profile)
    trial_comp = dict(comp)
    trial_comp.setdefault("status", schema.ComprehensionStatus.PRESENT.value)
    trial["comprehension"] = trial_comp
    problems = list(schema.validate(trial))

    # 2) Fail-closed membership of every load-bearing ref.
    problems.extend(check_membership(profile, trial_comp))

    if problems:
        # Refuse to write the understanding; record the rejection + findings.
        rejected = schema.empty_comprehension()
        rejected["status"] = schema.ComprehensionStatus.REJECTED.value
        rejected["findings"] = sorted(problems)
        if generated_by is not None:
            rejected["generated_by"] = dict(generated_by)
        profile["comprehension"] = rejected
        return MergeResult(False, schema.ComprehensionStatus.REJECTED.value, sorted(problems))

    # 3) Clean: write the canonical block with stable, sorted serialization.
    shell_sha = (
        (profile.get("provenance") or {}).get("shell") or {}
    ).get("sha256")
    canonical = _canonicalize(trial_comp, shell_sha, generated_by)
    profile["comprehension"] = canonical

    # 4) Derive the additive sinks from the canonical block (never written
    # independently).
    _derive_role_usage(profile, canonical)
    _derive_skeleton_attrs(profile, canonical)
    _derive_anchors(profile, canonical)

    return MergeResult(True, schema.ComprehensionStatus.PRESENT.value, [])


def _canonicalize(comp: dict, shell_sha: Optional[str], generated_by: Optional[dict]) -> dict:
    """Return the comprehension block with stable order and stamped provenance."""
    out = schema.empty_comprehension()
    out["status"] = schema.ComprehensionStatus.PRESENT.value
    out["source_shell_sha256"] = shell_sha
    if generated_by is not None:
        out["generated_by"] = dict(generated_by)
    elif comp.get("generated_by") is not None:
        out["generated_by"] = dict(comp["generated_by"])
    conf = comp.get("confidence")
    out["confidence"] = float(conf) if isinstance(conf, (int, float)) else 0.0

    # cover_slots: sorted by anchor_ref.
    slots = comp.get("cover_slots") or {}
    out["cover_slots"] = {
        k: dict(slots[k]) for k in sorted(slots) if isinstance(slots.get(k), dict)
    }

    # conventions.indexes: sorted by index_ref; sections: sorted by region_ref.
    conventions = comp.get("conventions") or {}
    indexes = [i for i in (conventions.get("indexes") or []) if isinstance(i, dict)]
    sections = [s for s in (conventions.get("sections") or []) if isinstance(s, dict)]
    out["conventions"] = {
        "indexes": sorted((dict(i) for i in indexes), key=lambda d: str(d.get("index_ref"))),
        "sections": sorted((dict(s) for s in sections), key=lambda d: str(d.get("region_ref"))),
    }

    # role_annotations: sorted by role id.
    annotations = comp.get("role_annotations") or {}
    out["role_annotations"] = {
        k: dict(annotations[k]) for k in sorted(annotations)
        if isinstance(annotations.get(k), dict)
    }

    # demo_classification.regions: sorted by region_ref.
    regions = [
        r for r in ((comp.get("demo_classification") or {}).get("regions") or [])
        if isinstance(r, dict)
    ]
    out["demo_classification"] = {
        "regions": sorted((dict(r) for r in regions), key=lambda d: str(d.get("region_ref")))
    }
    return out


# ---------------------------------------------------------------------------
# Derived sinks (Ruling B) - usage / skeleton attrs / anchors come FROM the block
# ---------------------------------------------------------------------------
def _derive_role_usage(profile: dict, comp: dict) -> None:
    """Refresh ``roles[*].usage`` advisory annotations from ``role_annotations``.

    Only the advisory free-text ``purpose`` / ``generation_rules`` are mirrored
    onto the role's ``usage`` object; the structural ``scope`` / ``placement`` /
    ``required`` / ``order`` already derived at extract are left intact (they come
    from the role family, not the model).
    """
    roles = profile.get("roles")
    if not isinstance(roles, dict):
        return
    for rid, ann in (comp.get("role_annotations") or {}).items():
        entry = roles.get(rid)
        if not isinstance(entry, dict) or not isinstance(ann, dict):
            continue
        usage = entry.setdefault("usage", {})
        if not isinstance(usage, dict):
            continue
        if ann.get("purpose") is not None:
            usage["purpose"] = ann["purpose"]
        if ann.get("generation_rules") is not None:
            usage["generation_rules"] = ann["generation_rules"]


def _derive_skeleton_attrs(profile: dict, comp: dict) -> None:
    """Stamp ``demo``/``required`` boolean attrs onto matching skeleton regions.

    Where a ``conventions.sections`` / ``demo_classification`` ref matches a
    structural skeleton region (by its surfaced id), the model's verdict is
    reflected as a boolean attribute the generator may branch on. Region NAMES
    are never rewritten (open tokens).
    """
    structure = profile.get("structure")
    if not isinstance(structure, dict):
        return
    skeleton = structure.get("skeleton")
    if not isinstance(skeleton, list):
        return
    demo_refs = {
        r.get("region_ref")
        for r in (comp.get("demo_classification") or {}).get("regions") or []
        if isinstance(r, dict) and r.get("verdict") == schema.Verdict.DEMO.value
    }
    required_refs = {
        s.get("region_ref"): bool(s.get("required"))
        for s in (comp.get("conventions") or {}).get("sections") or []
        if isinstance(s, dict)
    }
    for region in skeleton:
        if not isinstance(region, dict):
            continue
        rid = region.get("id") or region.get("region_ref")
        if rid in demo_refs:
            region["demo"] = True
        if rid in required_refs:
            region["required"] = required_refs[rid]


def _derive_anchors(profile: dict, comp: dict) -> None:
    """Annotate ``anchors.cover`` with the comprehension's slot count.

    Additive only: records how many cover slots the model bound, so a reader of
    ``anchors`` sees the comprehension-aware count without re-deriving it.
    """
    anchors = profile.get("anchors")
    if not isinstance(anchors, dict):
        return
    cover = anchors.setdefault("cover", {})
    if isinstance(cover, dict):
        cover["comprehended_slots"] = len(comp.get("cover_slots") or {})
