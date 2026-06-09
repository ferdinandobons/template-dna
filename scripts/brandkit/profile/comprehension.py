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
          "palette":       [ <palette_key>, ... ],
        }

    ``cover_anchors`` / ``fields`` / ``regions`` are read from
    ``surface.<kind>.{cover_anchors,fields,regions}``. Empty inventories are
    legal (for example XLSX has no TOC-style field code), but refs into an empty
    inventory have nothing to bind to and are fail-closed at QA time.
    ``roles`` is the concrete role-id list (``_index`` order if present).
    ``palette`` is the template-derived ``theme.palette`` key list (a theme slot
    like ``accent1`` or ``hex:RRGGBB``); a ``palette_annotations`` key binds to it
    fail-closed, so the model can NAME only a color the deterministic capture
    actually observed. Sorted for a stable surfaced order.

    This is the ONLY place the inventory shape is defined; the ``comprehend-input``
    verb and ``comprehension_targets_exist`` both call it so they cannot drift.
    """
    kind = profile.get("kind")
    surface = profile.get("surface") or {}
    sub = surface.get(kind) if isinstance(surface, dict) and kind else {}
    if not isinstance(sub, dict):
        sub = {}
    palette = (profile.get("theme") or {}).get("palette") or {}
    return {
        "cover_anchors": _entries_with_ids(sub.get("cover_anchors")),
        "fields": _entries_with_ids(sub.get("fields")),
        "regions": _entries_with_ids(sub.get("regions")),
        "roles": list(schema.list_role_ids(profile)),
        "palette": sorted(palette) if isinstance(palette, dict) else [],
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
        # Advisory hints of recurring structures the model MAY turn into reusable
        # fragments (proposed back via comprehension.fragments). May be empty.
        "fragment_candidates": _fragment_candidates(kind, catalog),
        # The brand PALETTE the model NAMES (model-driven color): each entry's
        # captured ``ref`` / ``provenance`` / ``frequency`` (names null in the
        # deterministic path). The model writes back ``palette_annotations`` keyed by
        # these palette ids; it NEVER authors a color. May be empty.
        "palette": _palette_facts(profile),
    }

    excerpt = _collect_excerpt(profile, catalog, excerpt_chars)
    return {"facts": facts, "excerpt": excerpt}


def _palette_facts(profile: dict) -> list[dict]:
    """The brand palette the model reasons over to NAME each color (model-driven
    color). Each fact is the captured ``{key, ref, provenance, frequency, name}`` -
    ``name`` surfaced as-is (null in the deterministic path) so the model sees what
    is still unnamed. The model NEVER receives a hex to author; it writes back only
    ``palette_annotations`` keyed by ``key``. Sorted by ``key`` (deterministic);
    empty when the template carried no color.
    """
    palette = (profile.get("theme") or {}).get("palette") or {}
    if not isinstance(palette, dict):
        return []
    out: list[dict] = []
    for key in sorted(palette):
        entry = palette[key]
        if not isinstance(entry, dict):
            continue
        out.append(
            {
                "key": key,
                "ref": entry.get("ref"),
                "provenance": entry.get("provenance") or [],
                "frequency": entry.get("frequency"),
                "name": entry.get("name"),
            }
        )
    return out


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


def _fragment_candidates(kind: Optional[str], catalog: Any) -> list[dict]:
    """Advisory, deterministic hints of recurring structures (plan: fragments).

    Derived cheaply from the ALREADY-captured ``artifact_catalog`` (no re-walk of
    the template, no extractor change). Each hint is ``{"kind", "signal",
    "evidence"}``; the model may turn a recurrence into a reusable fragment, or
    ignore it, or propose fragments not listed here. The list is bounded, sorted
    (deterministic), and frequently EMPTY - it is a nudge, never a binding, and the
    deterministic path never reads it.
    """
    if not isinstance(catalog, dict):
        return []
    out: list[dict] = []
    if kind == "pptx":
        # A layout that recurs across slides is a natural reusable-section motif.
        counts: dict[str, int] = {}
        for slide in catalog.get("slides") or []:
            if isinstance(slide, dict):
                layout = slide.get("layout")
                if isinstance(layout, str) and layout:
                    counts[layout] = counts.get(layout, 0) + 1
        for layout in sorted(counts):
            if counts[layout] >= 2:
                out.append(
                    {
                        "kind": "section",
                        "signal": f"layout {layout!r} recurs on {counts[layout]} slides",
                        "evidence": {"layout": layout, "slides": counts[layout]},
                    }
                )
    elif kind == "docx":
        # A paragraph style that recurs is a candidate single-fragment motif.
        counts = {}
        for sample in catalog.get("paragraph_samples") or []:
            if isinstance(sample, dict):
                style = sample.get("style")
                if isinstance(style, str) and style:
                    counts[style] = counts.get(style, 0) + 1
        for style in sorted(counts):
            if counts[style] >= 2:
                out.append(
                    {
                        "kind": "component",
                        "signal": f"paragraph style {style!r} recurs {counts[style]}x",
                        "evidence": {"style": style, "count": counts[style]},
                    }
                )
    elif kind == "xlsx":
        sheets = catalog.get("sheets")
        if isinstance(sheets, dict):
            with_tables = sorted(
                name
                for name, s in sheets.items()
                if isinstance(s, dict) and (s.get("tables") or [])
            )
            if len(with_tables) >= 2:
                out.append(
                    {
                        "kind": "section",
                        "signal": f"{len(with_tables)} sheets carry styled tables",
                        "evidence": {"sheets": with_tables},
                    }
                )
    return out[:12]


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
    palette_ids = set(inv["palette"])

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
    for i, reg in enumerate(
        (comp.get("demo_classification") or {}).get("regions") or []
    ):
        if not isinstance(reg, dict):
            continue
        ref = reg.get("region_ref")
        if ref not in region_ids:
            problems.append(
                f"comprehension.demo_classification.regions[{i}].region_ref: {ref!r} "
                f"not in surfaced regions inventory {sorted(region_ids)}"
            )

    # (d) role_annotations keys ∈ roles.
    for rid in comp.get("role_annotations") or {}:
        if rid not in role_ids:
            problems.append(
                f"comprehension.role_annotations: role id {rid!r} not in roles "
                f"{sorted(role_ids)}"
            )

    # (e) palette_annotations keys ∈ the surfaced palette inventory (FAIL-CLOSED on
    # empty, same rule as anchor/index/region): the model can NAME only a color the
    # deterministic capture actually observed, never invent a palette key.
    for key in comp.get("palette_annotations") or {}:
        if key not in palette_ids:
            problems.append(
                f"comprehension.palette_annotations: palette key {key!r} not in "
                f"surfaced palette inventory {sorted(palette_ids)}"
            )

    # (f) audit keys ∈ the profile-derived visual checklist (Cluster C1). FAIL-CLOSED
    # on empty, same rule as anchor/index/region: the L2 model may write a verdict
    # only against a structural checklist id the engine itself derived, never invent
    # one. The id set is the SOLE definition in ``qa.visual.visual_checklist_ids`` so
    # the model bundle, this gate, and the generate-time short-circuit cannot drift.
    # Lazy import (like ``block_from_dict`` above): ``qa.visual`` imports PIL/doctor
    # at module top, so a module-level import here would create a profile<-qa cycle
    # and drag PIL into every merge.
    audit = comp.get("audit")
    if isinstance(audit, dict) and audit:
        from brandkit.qa.visual import visual_checklist_ids

        checklist_ids = set(visual_checklist_ids(profile))
        for key in audit:
            if key not in checklist_ids:
                problems.append(
                    f"comprehension.audit: checklist id {key!r} not in the profile's "
                    f"derived visual checklist {sorted(checklist_ids)}"
                )

    return problems


def check_triage(profile: dict, comp: dict) -> list[str]:
    """Fail-closed validation of ``comprehension.triage`` (Cluster C2).

    Shape is checked by ``schema._validate_comp_triage`` (each entry's ``check`` in
    the closed :data:`schema.AMBIGUOUS_TRIAGE_CHECKS`, ``disposition`` in the closed
    :class:`schema.TriageDisposition`). This enforces what shape cannot, reject-never-skip:

      - each ``check`` is in the closed eligible set (re-checked here so the gate's
        ``check_triage_targets`` and merge agree on a single membership definition;
        the eligible set is WARNING-only, so a triage entry aimed at an ERROR-emitting
        check is rejected HERE - it can never demote an ERROR);
      - ``(check, location)`` is UNIQUE across the proposal (two dispositions for one
        finding is ambiguous), mirroring ``check_fragments``' ``(kind, ref)`` uniqueness.

    The model writes only a closed disposition + advisory evidence against a closed
    ``(check, location)`` pair; it never writes a brand value, so a validated triage
    entry cannot widen the brand guarantee. ``_apply_triage`` (the SOLE consumer)
    demotes ONLY a matching WARNING to INFO, so triage can never mask an ERROR.

    Returns ``[]`` when ``comp`` is absent / status != present (same status gate as
    :func:`check_membership`): an absent block carries no triage to enforce, and a
    rejected block is rebuilt (its findings already recorded). (``profile`` is unused
    but kept for signature symmetry with :func:`check_membership`.)
    """
    del profile  # triage binds to the closed check set, not to surfaced inventories
    if not isinstance(comp, dict):
        return []
    status = comp.get("status")
    if status not in (None, schema.ComprehensionStatus.PRESENT.value):
        return []
    triage = comp.get("triage")
    if not isinstance(triage, list) or not triage:
        return []

    problems: list[str] = []
    seen: set[tuple] = set()
    for i, entry in enumerate(triage):
        if not isinstance(entry, dict):
            continue  # shape validator already flags
        path = f"comprehension.triage[{i}]"
        check = entry.get("check")
        if check not in schema.AMBIGUOUS_TRIAGE_CHECKS:
            problems.append(
                f"{path}.check: {check!r} is not an eligible triage check "
                f"{sorted(schema.AMBIGUOUS_TRIAGE_CHECKS)}"
            )
            continue  # an out-of-set check cannot have a well-formed (check, location)
        location = entry.get("location")
        key = (check, location)
        if key in seen:
            problems.append(
                f"{path}: duplicate triage entry for (check={check!r}, "
                f"location={location!r})"
            )
        seen.add(key)
    return problems


# ---------------------------------------------------------------------------
# Refinement overlay (Cluster C3) - the qualitative-answer -> comprehension delta
# ---------------------------------------------------------------------------
# The closed set of comprehension sinks a refinement may touch. A delta key that
# is not one of these is IGNORED (the model can never smuggle a new field or shadow
# a structural one). ``audit`` / ``triage`` are deliberately ABSENT: refine is over
# the qualitative-understanding sinks only, never the QA-verdict sinks.
REFINABLE_SINKS: frozenset[str] = frozenset(
    {
        "role_annotations",
        "palette_annotations",
        "demo_classification",
        "cover_slots",
        "conventions",
    }
)


def overlay_refinement(existing: dict, delta: dict) -> dict:
    """Overlay a model-authored refinement ``delta`` onto an EXISTING comprehension.

    ``merge`` is REPLACE-from-single-source: ``_canonicalize`` rebuilds a fresh
    ``empty_comprehension()`` and copies only what the proposal carries, so passing a
    raw delta straight to ``merge`` would WIPE every existing sink. This primitive
    closes that trap: it returns a NEW dict that is the existing present block with
    the delta's sinks overlaid, ready to route whole through ``merge`` (which re-runs
    the full fail-closed validation + membership binding on the combined block).

    Closed-key, per-sink semantics (pure; ``existing`` / ``delta`` are not mutated):

      - ``role_annotations`` / ``palette_annotations`` / ``cover_slots``: shallow MAP
        update (a delta key replaces the matching existing entry, a new key is added).
      - ``demo_classification.regions``: MERGE-BY ``region_ref`` (a delta region
        replaces the matching existing one, a new ``region_ref`` is appended) - never
        a naive concat, which would dup-key a region.
      - ``conventions.indexes`` / ``conventions.sections``: MERGE-BY ``index_ref`` /
        ``region_ref`` respectively (same replace-or-append-by-ref rule).

    Any delta key NOT in :data:`REFINABLE_SINKS` is ignored, so the model cannot
    introduce a new field or overwrite a structural one (e.g. ``roles[*].resolver``,
    ``source_shell_sha256``, ``audit``, ``triage``). The result still carries the
    existing block's other sinks verbatim (deep-copied), so the subsequent ``merge``
    preserves them rather than dropping them.
    """
    import copy

    out = copy.deepcopy(existing) if isinstance(existing, dict) else {}
    if not isinstance(delta, dict):
        return out

    # (a) plain map sinks: shallow update keyed by id.
    for sink in ("role_annotations", "palette_annotations", "cover_slots"):
        d = delta.get(sink)
        if not isinstance(d, dict):
            continue
        base = out.get(sink)
        merged = dict(base) if isinstance(base, dict) else {}
        for key, val in d.items():
            if isinstance(key, str) and key:
                merged[key] = copy.deepcopy(val)
        out[sink] = merged

    # (b) demo_classification.regions: merge-by region_ref.
    demo = delta.get("demo_classification")
    if isinstance(demo, dict) and isinstance(demo.get("regions"), list):
        base_dc = out.get("demo_classification")
        regions = (
            list((base_dc or {}).get("regions") or [])
            if isinstance(base_dc, dict)
            else []
        )
        out["demo_classification"] = {
            "regions": _merge_by_ref(regions, demo["regions"], "region_ref")
        }

    # (c) conventions.indexes / conventions.sections: merge-by ref.
    conv = delta.get("conventions")
    if isinstance(conv, dict):
        base_conv = out.get("conventions")
        base_conv = base_conv if isinstance(base_conv, dict) else {}
        merged_conv = {
            "indexes": list(base_conv.get("indexes") or []),
            "sections": list(base_conv.get("sections") or []),
        }
        if isinstance(conv.get("indexes"), list):
            merged_conv["indexes"] = _merge_by_ref(
                merged_conv["indexes"], conv["indexes"], "index_ref"
            )
        if isinstance(conv.get("sections"), list):
            merged_conv["sections"] = _merge_by_ref(
                merged_conv["sections"], conv["sections"], "region_ref"
            )
        out["conventions"] = merged_conv

    return out


def _merge_by_ref(base: list, delta: list, ref_key: str) -> list:
    """Merge ``delta`` entries into ``base`` keyed by ``ref_key`` (replace or append).

    A delta entry whose ``ref_key`` matches an existing entry REPLACES it in place
    (preserving order); a delta entry with a new ``ref_key`` is APPENDED. This is the
    list-sink rule the overlay needs so a refinement of one region/index does not
    drop the others and a re-stated region does not duplicate (naive concat would).
    Entries without a usable ``ref_key`` are appended verbatim (the downstream merge
    re-validates shape + membership and will reject a malformed/dangling ref).
    """
    import copy

    out = [copy.deepcopy(e) for e in base]
    index: dict = {}
    for i, e in enumerate(out):
        if isinstance(e, dict):
            ref = e.get(ref_key)
            if isinstance(ref, str) and ref:
                index[ref] = i
    for e in delta:
        if not isinstance(e, dict):
            continue
        ref = e.get(ref_key)
        if isinstance(ref, str) and ref and ref in index:
            out[index[ref]] = copy.deepcopy(e)
        else:
            if isinstance(ref, str) and ref:
                index[ref] = len(out)
            out.append(copy.deepcopy(e))
    return out


def check_fragments(profile: dict, comp: dict) -> list[str]:
    """Fail-closed validation of ``comprehension.fragments`` block CONTENTS.

    Shape is checked by ``schema._validate_comp_fragments``; this enforces the
    fail-closed contract the shape validator deliberately cannot:

      - every fragment block must be parseable by :func:`block_from_dict` (a known
        IID primitive ``type``); an unparseable block is rejected HERE, at merge,
        not deferred to the loud-but-late ``expand_components`` failure at generate;
      - ``(kind, ref)`` must be unique across the proposal (two fragments writing
        the same registry slot is ambiguous);
      - a nested ``component``/``section`` block inside a fragment's ``blocks`` must
        resolve to another proposed fragment of the matching kind or to an existing
        registry entry, so the populated registry can never carry a dangling ref
        (which would otherwise hard-fail ``expand_components`` at generate time).

    A fragment block is presentation-free IID (it names intent, never a style /
    color / layout), so a validated fragment cannot widen the brand guarantee: its
    blocks resolve through the SAME chokepoint as any inline block.

    Not gated by ``status``: merge derives the registries from the proposal
    regardless of the incoming status (and forces ``present``), so fragment
    contents must always be validated. Returns ``[]`` when ``comp`` carries no
    fragments. (``profile`` is currently unused but kept for signature symmetry
    with :func:`check_membership` and future cross-binding.)
    """
    del profile  # nested refs bind to the proposal alone (single-source rebuild)
    if not isinstance(comp, dict):
        return []
    fragments = comp.get("fragments")
    if not isinstance(fragments, list) or not fragments:
        return []

    from brandkit.ir.model import IIDParseError, block_from_dict

    # Nested refs may resolve ONLY to another fragment proposed in THIS
    # comprehension. merge rebuilds the registries from the proposal alone (the
    # single source), so a ref to a pre-existing-but-not-reproposed entry would be
    # dangling after the rebuild, and binding to prior registry state would make
    # the merge outcome depend on history (non-deterministic for the same input).
    proposed = {"component": set(), "section": set()}
    for frag in fragments:
        if isinstance(frag, dict):
            kind = frag.get("kind")
            ref = frag.get("ref")
            if kind in proposed and isinstance(ref, str) and ref:
                proposed[kind].add(ref)

    problems: list[str] = []
    seen: set[tuple] = set()
    graph: dict[tuple, set] = {}
    for i, frag in enumerate(fragments):
        if not isinstance(frag, dict):
            continue  # shape validator already flags
        path = f"comprehension.fragments[{i}]"
        kind = frag.get("kind")
        ref = frag.get("ref")
        src = None
        if kind in proposed and isinstance(ref, str) and ref:
            key = (kind, ref)
            if key in seen:
                problems.append(f"{path}: duplicate {kind} ref {ref!r}")
            seen.add(key)
            src = key
            graph.setdefault(src, set())
        blocks = frag.get("blocks")
        if not isinstance(blocks, list):
            continue  # shape validator already flags
        for j, block in enumerate(blocks):
            bpath = f"{path}.blocks[{j}]"
            if not isinstance(block, dict):
                problems.append(f"{bpath}: must be a block object")
                continue
            btype = block.get("type")
            if btype in ("component", "section"):
                nref = block.get("ref")
                if not isinstance(nref, str) or nref not in proposed.get(btype, set()):
                    problems.append(
                        f"{bpath}: nested {btype} ref {nref!r} is not defined by "
                        f"another fragment proposed in this comprehension"
                    )
                elif src is not None:
                    graph[src].add((btype, nref))
                continue
            try:
                block_from_dict(block)
            except IIDParseError as exc:
                problems.append(f"{bpath}: {exc}")

    # A cyclic fragment reference can never expand to primitives; reject it at the
    # merge (the single writer) rather than letting it fail loud-but-late at the
    # generate-time depth guard.
    cycle = _detect_fragment_cycle(graph)
    if cycle:
        problems.append(
            "comprehension.fragments: cyclic fragment reference involving "
            f"{sorted(str(node) for node in cycle)}"
        )
    return problems


def _detect_fragment_cycle(graph: dict) -> set:
    """Return the nodes on a cycle in the proposed-fragment ref graph (or empty).

    Nodes are ``(kind, ref)``; edges are nested component/section refs (only edges
    into proposed nodes are recorded, so every edge target is itself a graph key).
    Implemented as an ITERATIVE colored DFS so an adversarially huge cycle can never
    overflow the Python recursion limit (it returns ``MergeResult(ok=False)`` rather
    than raising), mirroring the bounded ``_apply_slots`` guard.
    """
    white, gray, black = 0, 1, 2
    color: dict = {}
    for root in graph:
        if color.get(root, white) != white:
            continue
        color[root] = gray
        stack = [(root, iter(graph.get(root, ())))]
        while stack:
            node, it = stack[-1]
            advanced = False
            for nxt in it:
                cv = color.get(nxt, white)
                if cv == gray:
                    # Back-edge: nxt is on the active DFS path -> report the cycle.
                    path = [n for n, _ in stack]
                    return set(path[path.index(nxt) :]) if nxt in path else {nxt}
                if cv == white and nxt in graph:
                    color[nxt] = gray
                    stack.append((nxt, iter(graph.get(nxt, ()))))
                    advanced = True
                    break
            if not advanced:
                color[node] = black
                stack.pop()
    return set()


def _canonical_blocks(blocks: Any) -> list:
    """Round-trip each block through the IID parser so the stored template is
    canonical, presentation-free, free of unknown/dead fields, and INDEPENDENT (no
    mutable refs shared with the proposal or the canonical comprehension block).

    Only called after :func:`check_fragments` has confirmed parseability; the
    defensive fallback deep-copies an unexpectedly-unparseable block rather than
    aliasing it.
    """
    import copy

    from brandkit.ir.model import IIDParseError, block_from_dict

    out: list = []
    for b in blocks if isinstance(blocks, list) else []:
        try:
            out.append(block_from_dict(b).to_dict())
        except IIDParseError:
            out.append(copy.deepcopy(b))
    return out


def _derive_fragment_registries(comp: dict) -> tuple[dict, dict]:
    """Derive ``(components, sections)`` registries from ``comprehension.fragments``.

    Each well-shaped fragment becomes a registry entry ``{'blocks': [...],
    'purpose'?}`` keyed by ``ref``; built with sorted refs for stable, idempotent
    serialization, and with blocks round-tripped to canonical IID. Only called on a
    CLEAN merge, so every fragment is already validated. Malformed entries (should
    be none here) are skipped defensively.
    """
    components: dict = {}
    sections: dict = {}
    frags = comp.get("fragments")
    if not isinstance(frags, list):
        return components, sections
    by_kind: dict[str, list[tuple[str, dict]]] = {"component": [], "section": []}
    for frag in frags:
        if not isinstance(frag, dict):
            continue
        kind = frag.get("kind")
        ref = frag.get("ref")
        blocks = frag.get("blocks")
        if kind not in by_kind or not isinstance(ref, str) or not ref:
            continue
        if not isinstance(blocks, list):
            continue
        entry: dict = {"blocks": _canonical_blocks(blocks)}
        purpose = frag.get("purpose")
        if isinstance(purpose, str) and purpose:
            entry["purpose"] = purpose
        by_kind[kind].append((ref, entry))
    for ref, entry in sorted(by_kind["component"], key=lambda kv: kv[0]):
        components[ref] = entry
    for ref, entry in sorted(by_kind["section"], key=lambda kv: kv[0]):
        sections[ref] = entry
    return components, sections


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
    # merge DISPOSES status: a model-supplied status is never trusted (it would
    # otherwise let a status='rejected'/'absent' input short-circuit the
    # membership / fragment checks while merge still derives the registries). Force
    # the trial to PRESENT so every load-bearing validation always runs.
    trial_comp["status"] = schema.ComprehensionStatus.PRESENT.value
    trial["comprehension"] = trial_comp
    problems = list(schema.validate(trial))

    # 2) Fail-closed membership of every load-bearing ref.
    problems.extend(check_membership(profile, trial_comp))

    # 2b) Fail-closed validation of any reusable-fragment proposals (block
    # parseability, ref uniqueness, nested-ref resolution). Part of the SAME
    # all-or-nothing transaction: a bad fragment rejects the whole comprehension
    # and writes nothing into the registries.
    problems.extend(check_fragments(profile, trial_comp))

    # 2c) Fail-closed validation of any model-assisted QA-triage entries (Cluster
    # C2): each names an ELIGIBLE WARNING-only check and a UNIQUE (check, location)
    # pair. Part of the SAME all-or-nothing transaction (easy to forget; without this
    # line triage would be canonicalized unguarded). A triage entry can never demote
    # an ERROR - the eligible set is WARNING-only, so an ERROR-aimed entry is rejected
    # here, and ``qa.gate._apply_triage`` independently guards on severity==WARNING.
    problems.extend(check_triage(profile, trial_comp))

    if problems:
        # Refuse to write the understanding; record the rejection + findings.
        rejected = schema.empty_comprehension()
        rejected["status"] = schema.ComprehensionStatus.REJECTED.value
        rejected["findings"] = sorted(problems)
        if generated_by is not None:
            rejected["generated_by"] = dict(generated_by)
        profile["comprehension"] = rejected
        return MergeResult(
            False, schema.ComprehensionStatus.REJECTED.value, sorted(problems)
        )

    # 3) Clean: write the canonical block with stable, sorted serialization.
    shell_sha = ((profile.get("provenance") or {}).get("shell") or {}).get("sha256")
    canonical = _canonicalize(trial_comp, shell_sha, generated_by)
    profile["comprehension"] = canonical

    # 4) Derive the additive sinks from the canonical block (never written
    # independently).
    _derive_role_usage(profile, canonical)
    _derive_skeleton_attrs(profile, canonical)
    _derive_anchors(profile, canonical)
    _derive_palette_annotations(profile, canonical)

    # 4b) Derive the reusable-fragment registries from the canonical fragments.
    # comprehend OWNS components/sections: they are rebuilt deterministically from
    # the (single-source) comprehension on every clean merge, so a re-merge of the
    # same proposal yields byte-identical registries (idempotency), and a proposal
    # with no fragments resets them to empty. A fragment is presentation-free IID,
    # so this can never widen the brand guarantee.
    components, sections = _derive_fragment_registries(canonical)
    profile["components"] = components
    profile["sections"] = sections

    return MergeResult(True, schema.ComprehensionStatus.PRESENT.value, [])


def _canonicalize(
    comp: dict, shell_sha: Optional[str], generated_by: Optional[dict]
) -> dict:
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
        "indexes": sorted(
            (dict(i) for i in indexes), key=lambda d: str(d.get("index_ref"))
        ),
        "sections": sorted(
            (dict(s) for s in sections), key=lambda d: str(d.get("region_ref"))
        ),
    }

    # role_annotations: sorted by role id.
    annotations = comp.get("role_annotations") or {}
    out["role_annotations"] = {
        k: dict(annotations[k])
        for k in sorted(annotations)
        if isinstance(annotations.get(k), dict)
    }

    # palette_annotations: sorted by palette key (the model NAMES a captured color).
    palette_ann = comp.get("palette_annotations") or {}
    out["palette_annotations"] = {
        k: dict(palette_ann[k])
        for k in sorted(palette_ann)
        if isinstance(palette_ann.get(k), dict)
    }

    # audit: sorted by checklist id (Cluster C1). REQUIRED arm - ``empty_comprehension``
    # rebuilds with ``audit={}``, so omitting this copy would silently discard the
    # persisted L2 verdict on every merge. The model wrote only closed dispositions
    # + advisory evidence + per-row shas against ids it did not author.
    audit = comp.get("audit") or {}
    out["audit"] = {
        k: dict(audit[k]) for k in sorted(audit) if isinstance(audit.get(k), dict)
    }

    # triage: sorted by (check, location-or-"") for a stable, idempotent serialization
    # (Cluster C2). REQUIRED arm - ``empty_comprehension`` rebuilds with ``triage=[]``,
    # so omitting this copy would silently discard the model's triage on every merge.
    # No derived sink: triage is consumed LIVE by ``run_qa._apply_triage``, never
    # mirrored onto roles/theme. The model wrote only a closed disposition + advisory
    # evidence against a closed (check, location) pair.
    triage = [t for t in (comp.get("triage") or []) if isinstance(t, dict)]
    out["triage"] = sorted(
        (dict(t) for t in triage),
        key=lambda d: (str(d.get("check")), str(d.get("location") or "")),
    )

    # demo_classification.regions: sorted by region_ref.
    regions = [
        r
        for r in ((comp.get("demo_classification") or {}).get("regions") or [])
        if isinstance(r, dict)
    ]
    out["demo_classification"] = {
        "regions": sorted(
            (dict(r) for r in regions), key=lambda d: str(d.get("region_ref"))
        )
    }

    # fragments: sorted by (kind, ref) for a stable, idempotent serialization.
    frags = [f for f in (comp.get("fragments") or []) if isinstance(f, dict)]
    out["fragments"] = sorted(
        (_canonical_fragment(f) for f in frags),
        key=lambda d: (str(d.get("kind")), str(d.get("ref"))),
    )
    return out


def _canonical_fragment(frag: dict) -> dict:
    """Return a canonical reusable-fragment proposal entry (stable key order).

    Blocks are round-tripped to canonical IID (independent of the derived registry
    copy, so neither aliases the other).
    """
    out: dict = {
        "ref": frag.get("ref"),
        "kind": frag.get("kind"),
        "blocks": _canonical_blocks(frag.get("blocks")),
    }
    purpose = frag.get("purpose")
    if isinstance(purpose, str) and purpose:
        out["purpose"] = purpose
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


def _derive_palette_annotations(profile: dict, comp: dict) -> None:
    """Mirror the model's palette NAMES onto ``theme.palette[key]`` (model-driven
    color), exactly like :func:`_derive_role_usage` mirrors role annotations.

    Only the advisory free-text fields (``name`` / ``purpose`` / ``use_when`` /
    ``semantic_role``) are copied onto the matching palette entry; the structural
    ``ref`` / ``provenance`` / ``frequency`` are NEVER touched (they are the
    deterministic capture's, and the model never authors a real color). A key with
    no matching palette entry is skipped (membership already gated it fail-closed).
    """
    palette = (profile.get("theme") or {}).get("palette") or {}
    if not isinstance(palette, dict):
        return
    for key, ann in (comp.get("palette_annotations") or {}).items():
        entry = palette.get(key)
        if not isinstance(entry, dict) or not isinstance(ann, dict):
            continue
        for field in schema.PALETTE_ANNOTATION_FIELDS:
            if ann.get(field) is not None:
                entry[field] = ann[field]


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
        # Skeleton regions are keyed by ``region`` (their region id, e.g.
        # ``section.toc``); match that against the comprehension's region_refs.
        # (The old ``id``/``region_ref`` lookup was always None, so demo/required
        # were never annotated onto the skeleton.)
        rid = region.get("region")
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
