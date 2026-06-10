# SPDX-License-Identifier: MIT
"""DOCX generation from an IntermediateDocument."""

from __future__ import annotations

from pathlib import Path
from typing import Optional

from docx import Document
from docx.enum.dml import MSO_THEME_COLOR
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.shared import Emu, Pt, RGBColor

from brandkit.common import appearance
from brandkit.common import color as colorutil
from brandkit.common import text as textutil
from brandkit.common.links import is_safe_link_url
from brandkit.formats.docx import cover, structure
from brandkit.formats.docx.styles import lookup_style
from brandkit.ir import components
from brandkit.ir import model as ir
from brandkit.ooxml import chart as chartlib
from brandkit.ooxml.idempotency import repack_fixed_timestamps
from brandkit.profile import schema, store
from brandkit.profile.reconcile import confidence_clears_floor
from brandkit.profile.resolver import ProfileResolver
from brandkit.qa.checks_deterministic import (
    check_index_matches_content,
    check_no_net_structure_loss,
)
from brandkit.qa.model import Finding

# Block types that carry no DOCX writer. ``component``/``section`` are EXPANDED to
# primitives before this writer runs (``expand_components``), so they should never
# reach it; this set is the defensive fail-loud backstop if one ever survives. A
# block in this set is skipped cleanly (no blank paragraph) and recorded as a
# degradation finding rather than dropped silently. ``toc`` is now a native writer
# (``_write_toc``), so it is no longer here.
_UNHANDLED_BLOCK_TYPES: frozenset[str] = frozenset({"component", "section"})


class GenerationError(ValueError):
    """Raised when a block cannot be written and must not be silently dropped."""


def generate(
    profile: dict,
    shell_path: str | Path,
    idoc: ir.IntermediateDocument,
    output: str | Path,
    *,
    findings: Optional[list[Finding]] = None,
) -> Path:
    """Generate a DOCX from ``idoc`` onto the brand ``shell_path``.

    ``findings`` (optional out-param) is appended with any degradation / style-miss
    findings recorded during writing so the QA gate can surface them. Off-brand
    output stays impossible by construction: every style comes from the profile via
    :func:`brandkit.formats.docx.styles.lookup_style`, and a role that resolves to
    no matching shell style is recorded LOUDLY (ERROR) instead of silently falling
    back to ``Normal``.
    """
    sink: list[Finding] = findings if findings is not None else []
    idoc = components.expand_components(idoc, profile)
    doc = Document(shell_path)

    # Repair any malformed section measures (non-integer twips some editors emit)
    # before any python-docx access - including python-docx internals like
    # add_table's width computation - can choke on them.
    structure.sanitize_section_measures(doc)

    # Order-aware body replacement: remove ONLY the freeform body region, keeping
    # the ordered cover and TOC regions (and the final sectPr) in place. New
    # content is appended into the now-empty body region - immediately before the
    # sectPr - which is exactly the right slot.
    struct = profile.get("structure")
    structure.clear_body_region(doc, struct, preserve_cover=True, preserve_toc=True)

    # Reconcile the preserved derived indexes (TOC/TOF/TOT) and demo regions with
    # the new content FIRST, comprehension-steered when present (no-op when absent).
    # This MUST run before the cover fill: the surfaced index ids (and the orphan-
    # index-heading lookup) are position-based over the top-level body children, and
    # ``compose_cover`` inserts a paragraph into the cover region, shifting every
    # subsequent child index. The body clear above leaves the cover/TOC positions
    # untouched (the body region is last), so resolving the index refs here - before
    # any cover insert - keeps the position ids valid. Removing a stale caption index
    # also removes its introducing heading (the orphan-index-heading fix).
    removed_refs = _reconcile_indexes_and_demo(doc, profile, idoc, sink)

    # Caption-index reconciliation can expose blank TOC separators and preserved
    # demo-body section breaks at the start of the body slot. Prune them only when
    # this generation actually removed such structure; otherwise keep historical
    # multi-section geometry byte-for-byte.
    if removed_refs:
        structure.prune_leading_empty_body_artifacts(doc)

    # Fill the preserved cover anchors in place (never recreate the cover). Returns
    # the set of cover anchor refs the reconciliation CLEARED, for the destructive
    # floor below.
    cleared_anchors = cover.compose_cover(doc, idoc.cover, profile, findings=sink)

    # Write the body blocks in the order given (the body region is freeform). The
    # caption indexer threads SEQ numbering through every caption so a KEPT caption
    # index (table-of-tables / -figures) can regenerate from the new content.
    resolver = ProfileResolver(profile)
    caption_ctx = _build_caption_indexer(profile)
    for block in idoc.blocks:
        _write_block(doc, resolver, block, sink, caption_ctx)

    # Keep the visible TOC field cache aligned for renderers that do not update
    # fields in headless export. The field itself remains dirty/updateable.
    structure.refresh_visible_outline_toc_cache(doc, _outline_headings(idoc.blocks))

    # Rebuild every KEPT caption index's visible cache from the captions just emitted
    # (the SEQ classes the indexer collected), so a headless render shows the new
    # entries instead of the template's stale ones. No-op when nothing was collected.
    if caption_ctx is not None and caption_ctx.entries:
        structure.refresh_visible_caption_index_cache(doc, caption_ctx.entries)

    # Refresh the preserved TOC (if any) so Word recomputes it on open - the new
    # headings written into the body will be picked up. No-op when there is no TOC.
    structure.refresh_toc(doc)

    # Destructive-action floor (plan §6): every cover anchor / index block the
    # reconciliation removed must carry a corroborated destructive verdict AND clear
    # the confidence floor, else ERROR (a wrong delete is not recoverable). Model-free;
    # reads frozen verdicts. The confidence threaded here is the model's single
    # comprehension confidence (the same value the reconcile sites gate on).
    if store.comprehension_is_present(profile):
        comp = profile.get("comprehension")
        confidence = (
            float(comp.get("confidence") or 0.0) if isinstance(comp, dict) else None
        )
        sink.extend(
            check_no_net_structure_loss(
                cleared_anchors | removed_refs, profile, confidence=confidence
            )
        )

    out = Path(output)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out)
    # python-docx's core.xml is already stable, so only the ZIP entry timestamps
    # need normalizing for byte-idempotent re-runs (no modified<-created pin).
    repack_fixed_timestamps(out)
    return out


def _outline_headings(blocks: list) -> list[tuple[int, str]]:
    headings: list[tuple[int, str]] = []
    for block in blocks:
        if isinstance(block, ir.Heading):
            text = textutil.runs_to_text(block.runs).strip()
            if text:
                headings.append((block.level, text))
    return headings


def _content_has_captionables(idoc: ir.IntermediateDocument) -> bool:
    """True when the new content carries any captionable item (caption/table/image).

    A derived caption index (table-of-tables / table-of-figures) has something to
    point at only when the content contains captionable blocks. This is the
    deterministic corroboration for KEEPING such an index; an empty result
    corroborates REMOVING it (plan §6 caption/index reconciliation).
    """
    for block in idoc.blocks:
        if isinstance(block, ir.Caption):
            return True
        if isinstance(block, ir.Table) and getattr(block, "caption", None):
            return True
        if isinstance(block, ir.Image) and getattr(block, "caption", None):
            return True
    return False


def _reconcile_indexes_and_demo(
    doc, profile: dict, idoc: ir.IntermediateDocument, findings: list[Finding]
) -> set[str]:
    """Reconcile preserved derived indexes + demo regions with the new content.

    Comprehension-steered (no-op when comprehension is absent - the deterministic
    body clear already removed stale body content, and the outline TOC is refreshed
    separately). For each ``conventions.indexes`` entry:

      - ``reconcile == clear``  -> REMOVE the orphan index block, but ONLY when the
        destructive floor corroborates it: the model also tagged the index's own
        region (``region.<field-id>``) ``verdict=demo``, OR the content carries no
        captionable item the index could point at. Otherwise KEEP + WARNING.
      - ``reconcile in {regenerate, preserve}``  -> KEEP (the outline TOC is
        refreshed by ``refresh_toc``; a caption index is left for Word to recompute).

    Demo regions tagged ``verdict=demo`` that map to a derived-index region are
    removed via the same index path; demo regions over the freeform body are
    already gone (the body clear). Returns the set of index refs actually removed
    (for ``no_net_structure_loss``).
    """
    comp = profile.get("comprehension")
    if not store.comprehension_is_present(profile) or not isinstance(comp, dict):
        return set()

    has_captionables = _content_has_captionables(idoc)
    # The model's single comprehension confidence - the SAME value the cover
    # reconciler gates on - so a CLEAR is honored uniformly across formats.
    confidence = float(comp.get("confidence") or 0.0)
    demo_region_refs = {
        r.get("region_ref")
        for r in (comp.get("demo_classification") or {}).get("regions") or []
        if isinstance(r, dict) and r.get("verdict") == schema.Verdict.DEMO.value
    }

    # Collect every index ruled CLEAR whose deletion the destructive floor
    # corroborates, then remove them all in one shift-safe pass (a position-based
    # field id would otherwise be invalidated by a prior removal).
    to_clear: list[str] = []
    for idx in (comp.get("conventions") or {}).get("indexes") or []:
        if not isinstance(idx, dict):
            continue
        index_ref = idx.get("index_ref")
        if not index_ref:
            continue
        if idx.get("reconcile") != schema.Reconcile.CLEAR.value:
            continue  # regenerate / preserve: keep (refresh handles the outline TOC)
        # Destructive floor: a CLEAR is honored only when (i) the model's confidence
        # clears the floor AND (ii) determinism corroborates the removal - the model
        # also tagged the index's region demo, or the content has no captionable item
        # the index could point at. Either gate failing downgrades to KEEP + WARNING
        # (a wrong delete is unrecoverable). The confidence gate mirrors the cover
        # reconcilers so the SAME confidence yields the SAME behavior per format.
        if not confidence_clears_floor(confidence):
            findings.append(
                Finding(
                    "index_clear_downgraded",
                    schema.Severity.WARNING.value,
                    f"index {index_ref!r} clear not corroborated "
                    f"(confidence {confidence:.2f}); kept",
                )
            )
            continue
        region_ref = f"region.{index_ref}"
        corroborated = (region_ref in demo_region_refs) or (not has_captionables)
        if not corroborated:
            findings.append(
                Finding(
                    "index_clear_downgraded",
                    schema.Severity.WARNING.value,
                    f"index {index_ref!r} clear not corroborated "
                    f"(no demo verdict and content has captionable items); kept",
                )
            )
            continue
        to_clear.append(index_ref)
    removed = structure.remove_index_fields(doc, to_clear)

    # Self-awareness check (plan §7): a preserved CAPTION index (one carrying a
    # deterministic ``seq_id`` from its ``\c`` switch) that is KEPT yet has no
    # captionable content to point at is flagged WARNING. The set of emitted SEQ
    # classes is keyed on the opaque ``seq_id`` (never the advisory ``kind``): a
    # caption index's seq is "emitted" only when the content carries captionable
    # items. Empty when none, so every kept-but-stale caption index is surfaced.
    present_seq_ids: set[str] = set()
    if has_captionables:
        for idx in (comp.get("conventions") or {}).get("indexes") or []:
            if not isinstance(idx, dict):
                continue
            if idx.get("index_ref") in removed:
                continue  # already cleared; not a kept index
            seq = idx.get("seq_id")
            if seq:
                present_seq_ids.add(seq)
    findings.extend(check_index_matches_content(present_seq_ids, profile))
    return removed


def _apply_run_toggles(run, r: dict) -> None:
    """Apply inline emphasis from an IR run to a docx run.

    Only character-level toggles (``w:b``/``w:i``/``w:u``/strike/super-/subscript) -
    author intent, NOT a brand style/font/color - are applied, so the brand
    guarantee (generators never write a literal style/hex/font) is untouched. The
    run still inherits the paragraph style's brand font and color. A ``code`` toggle
    carries no brand-safe typeface (a monospace family would be a literal font), so
    it is rendered as plain text rather than fabricating one.
    """
    if r.get("b"):
        run.bold = True
    if r.get("i"):
        run.italic = True
    if r.get("u"):
        run.underline = True
    if r.get("strike"):
        run.font.strike = True
    if r.get("sup"):
        run.font.superscript = True
    if r.get("sub"):
        run.font.subscript = True


def _resolve_run_color(
    resolver: Optional[ProfileResolver],
    token: Optional[str],
    findings: Optional[list[Finding]],
) -> Optional[dict]:
    """Resolve a run's ``color`` palette TOKEN to its captured color ``ref`` object.

    A thin docx alias for the format-neutral ``common.appearance.resolve_run_color``:
    returns the ``{'kind': 'theme'|'hex', ...}`` object the writers apply, or ``None``
    when there is no token, no resolver, or the token is unknown (an UNRESOLVED token
    records a graceful INFO ``color_token_unresolved`` finding and leaves the run
    inherited - the writer NEVER fabricates a color for an unknown token)."""
    return appearance.resolve_run_color(resolver, token, findings)


def _inject_hyperlink_run_color(rpr, color: Optional[dict]) -> None:
    """Inject the resolved run COLOR onto a raw-XML hyperlink ``w:rPr``, gated on an
    ABSENT ``w:color`` so re-runs stay byte-identical (exactly like the ``w:rFonts``
    injection above it).

    A hex ref writes ``w:color@w:val`` (normalized RRGGBB); a theme-token ref writes
    ``w:color@w:themeColor`` via the CLOSED :data:`_WML_TOKEN_TO_THEME_COLOR` map (the
    WordprocessingML themeColor token). A theme token outside that closed map, or a
    malformed hex, is SKIPPED (the link stays inherited) rather than risking invalid
    XML - the value comes STRICTLY from the resolved palette ref, never a literal."""
    if not color or rpr.find(qn("w:color")) is not None:
        return
    kind = color.get("kind")
    if kind == "hex":
        hexval = color.get("hex")
        if not hexval:
            return
        try:
            normalized = colorutil.normalize_hex(hexval)
        except (ValueError, AttributeError):
            return
        _link_color_element(rpr).set(qn("w:val"), normalized)
        return
    if kind == "theme":
        token = color.get("theme")
        member = _WML_TOKEN_TO_THEME_COLOR.get(token)
        hexval = color.get("hex")
        if member is None:
            # No WordprocessingML themeColor member (a clrScheme-slot token like
            # dk1/lt1/hlink): realize the brand color via the resolver-supplied hex
            # rather than dropping the link to inherited.
            if hexval:
                try:
                    normalized = colorutil.normalize_hex(hexval)
                except (ValueError, AttributeError):
                    return
                _link_color_element(rpr).set(qn("w:val"), normalized)
            return
        el = _link_color_element(rpr)
        # python-docx's themeColor attribute carries the SAME WML token the closed
        # map is keyed on; write it back verbatim. Also carry the resolved hex in
        # w:val (from the palette ref, never a literal) so renderers that ignore a
        # run themeColor still show the brand color; Word uses the themeColor. With
        # no resolved hex, emit themeColor only (no w:val) so a renderer that reads
        # w:val does not paint the link black.
        el.set(qn("w:themeColor"), member.xml_value)
        if hexval:
            try:
                el.set(qn("w:val"), colorutil.normalize_hex(hexval))
            except (ValueError, AttributeError):
                pass


def _link_color_element(rpr):
    """Return ``rpr``'s ``w:color`` inserted at the schema-correct position, cleared.

    ``CT_RPr.get_or_add_color`` places ``w:color`` BEFORE ``w:u``/``w:sz``/
    ``w:vertAlign`` (a raw ``rpr.append`` would emit non-conformant run-property
    child order on an underlined/super-scripted link). python-docx seeds the new
    element with a default ``w:val="000000"``, so clear it: the caller then writes
    ONLY the attributes it intends, and a theme-token-without-hex link is not painted
    black by a renderer that reads ``w:val``. Caller guards on an absent ``w:color``,
    so this never duplicates and re-runs stay byte-identical.
    """
    el = rpr.get_or_add_color()
    el.attrib.pop(qn("w:val"), None)
    return el


def _add_hyperlink(
    para,
    url: str,
    text: str,
    r: dict,
    latin: Optional[str] = None,
    *,
    color: Optional[dict] = None,
    findings: Optional[list[Finding]] = None,
) -> None:
    """Append a real ``w:hyperlink`` (external relationship) carrying ``text``.

    The link target is the author's URL (content, not brand), wired through a
    package relationship. The visible run keeps the author's inline emphasis; we do
    not inject a literal ``Hyperlink`` character style or color (brand guarantee).

    A hyperlink run is raw ``w:r`` XML appended under the ``w:hyperlink`` element, so
    it is NOT a python-docx ``Run`` and ``_apply_appearance`` (which iterates
    ``para.runs``) never reaches it. When the resolved op carries a captured brand
    font (``latin``), brand the run here by injecting ``w:rPr/w:rFonts`` with ONLY
    ``w:ascii`` + ``w:hAnsi`` - exactly what python-docx's ``run.font.name`` writes
    (no ``w:cs``), so the link text matches the surrounding branded runs. ``latin``
    is read STRICTLY from the resolver op (never a literal). The rFonts is injected
    only when absent, so re-runs stay byte-identical.

    ``color`` (when given) is the resolved palette ref for this run's ``color`` token.
    For the raw-XML link run it is injected as ``w:color`` (hex ``w:val`` or theme
    ``w:themeColor``) on the rPr, gated on absent like ``w:rFonts``; for an unsafe-url
    link (a python-docx run) it goes through the shared ``_brand_run_color``. The
    value comes STRICTLY from the resolved palette ref - never a literal.

    An UNSAFE scheme (``file:``/``smb:``/``javascript:``/``data:``/...) is neutralized:
    the author's TEXT is kept (emphasis preserved, and branded via the python-docx
    run) but the dangerous target is not wired, so untrusted content cannot smuggle a
    hostile link into the document.
    """
    if not is_safe_link_url(url):
        run = para.add_run(text)
        _apply_run_toggles(run, r)
        _brand_run_font(run, latin)
        _brand_run_color(run, color, findings if findings is not None else [])
        return
    r_id = para.part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    run = OxmlElement("w:r")
    rpr: Optional[object] = None
    if any(r.get(k) for k in ("b", "i", "u", "strike", "sup", "sub")):
        rpr = OxmlElement("w:rPr")
        for key, tag in (("b", "w:b"), ("i", "w:i"), ("strike", "w:strike")):
            if r.get(key):
                rpr.append(OxmlElement(tag))
        if r.get("u"):
            u = OxmlElement("w:u")
            u.set(qn("w:val"), "single")
            rpr.append(u)
        if r.get("sup") or r.get("sub"):
            va = OxmlElement("w:vertAlign")
            va.set(qn("w:val"), "superscript" if r.get("sup") else "subscript")
            rpr.append(va)
    # Brand the link run with the captured font (rFonts ascii+hAnsi only, matching
    # python-docx's run.font.name), gated on an ABSENT w:rFonts so re-runs are
    # byte-identical. The font value comes only from the resolver op (no literal).
    if latin:
        if rpr is None:
            rpr = OxmlElement("w:rPr")
        if rpr.find(qn("w:rFonts")) is None:
            rfonts = OxmlElement("w:rFonts")
            rfonts.set(qn("w:ascii"), latin)
            rfonts.set(qn("w:hAnsi"), latin)
            rpr.insert(0, rfonts)
    # Inject the resolved run color (w:color), gated on absent like w:rFonts above,
    # so a colored hyperlink token reaches the raw-XML link run too.
    if color:
        if rpr is None:
            rpr = OxmlElement("w:rPr")
        _inject_hyperlink_run_color(rpr, color)
    if rpr is not None:
        run.append(rpr)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    run.append(t)
    hyperlink.append(run)
    para._p.append(hyperlink)


def _add_runs(
    para,
    runs,
    latin: Optional[str] = None,
    *,
    resolver: Optional[ProfileResolver] = None,
    findings: Optional[list[Finding]] = None,
) -> None:
    """Write IR ``runs`` into ``para`` as real docx runs, preserving inline emphasis
    and hyperlinks, instead of flattening to a single plain run.

    ``latin`` (when given) is the captured brand font of the resolved op for this
    paragraph; it is applied to plain runs and to hyperlink runs at write time so the
    branding reaches even the raw-XML hyperlink run ``_apply_appearance`` cannot see.
    Plain (non-hyperlink) runs added here are also branded by the post-write
    ``_apply_appearance`` pass; the guard (set only when absent) makes that a no-op.

    ``resolver`` (when given) enables model-driven run COLOR: a run carrying a
    ``color`` palette TOKEN has it resolved to a captured ref via
    ``resolver.resolve_color`` and applied here through the EXISTING
    ``_brand_run_color`` (whose ``run.font.color.type is None`` guard gives the
    explicit token first-writer-wins precedence over the later ``_apply_appearance``
    body/role default, and keeps re-runs byte-identical). An UNRESOLVED token is left
    inherited + an INFO ``color_token_unresolved`` finding. No resolver -> no per-run
    color (the body/role default still applies via ``_apply_appearance``)."""
    sink = findings if findings is not None else []
    for r in runs or []:
        text = str(r.get("t", ""))
        link = r.get("link")
        color = _resolve_run_color(resolver, r.get("color"), findings)
        if link and text:
            _add_hyperlink(para, link, text, r, latin, color=color, findings=findings)
        elif text:
            run = para.add_run(text)
            _apply_run_toggles(run, r)
            _brand_run_font(run, latin)
            appearance.apply_run_color(DOCX_BACKEND, run, color, sink)


def _para_with_runs(
    doc,
    runs,
    latin: Optional[str] = None,
    *,
    resolver: Optional[ProfileResolver] = None,
    findings: Optional[list[Finding]] = None,
):
    """Add a paragraph carrying ``runs`` as real, formatting-preserving docx runs.

    ``resolver`` / ``findings`` (when given) thread model-driven run color into
    ``_add_runs`` so a run's ``color`` palette token is resolved and applied."""
    para = doc.add_paragraph()
    _add_runs(para, runs, latin, resolver=resolver, findings=findings)
    return para


class _CaptionIndexer:
    """SEQ numbering + visible-cache entries for the KEPT caption indexes.

    Maps a caption's ``target`` (table/figure) to the opaque ``seq_id`` of the
    comprehension's matching kept caption index (via its model-authored
    ``caption_target``), so the generator can emit a real Word ``SEQ`` field that the
    index's ``\\c`` switch collects - brand-agnostic, no language heuristic on the seq
    name. Accumulates the visible entry text per seq for the post-body cache refresh.
    """

    def __init__(self, target_to_seq: dict) -> None:
        self.target_to_seq = target_to_seq
        self._counts: dict = {}
        self.entries: dict = {}

    def seq_for(self, target: Optional[str]) -> Optional[str]:
        return self.target_to_seq.get(target) if target else None

    def emit(self, seq: str, text: str) -> int:
        """Record the next caption of class ``seq``; return its 1-based number."""
        n = self._counts.get(seq, 0) + 1
        self._counts[seq] = n
        self.entries.setdefault(seq, []).append(f"{seq} {n}. {text}".rstrip())
        return n


def _build_caption_indexer(profile: dict) -> Optional["_CaptionIndexer"]:
    """Build the caption indexer from the KEPT caption indexes in the comprehension.

    An index participates only when it is KEPT (``reconcile`` != ``clear``) AND carries
    a ``seq_id`` AND a model-authored ``caption_target`` (the table/figure kind that
    feeds it). A cleared index is gone; an index with no target cannot be mapped to a
    caption kind. Returns None when no kept caption index opts in - captions then render
    as plain styled text (the prior behavior), so SEQ emission is strictly additive.
    """
    comp = profile.get("comprehension")
    if not store.comprehension_is_present(profile) or not isinstance(comp, dict):
        return None
    target_to_seq: dict = {}
    for idx in (comp.get("conventions") or {}).get("indexes") or []:
        if not isinstance(idx, dict):
            continue
        if idx.get("reconcile") == schema.Reconcile.CLEAR.value:
            continue
        seq = idx.get("seq_id")
        target = idx.get("caption_target")
        if seq and target in schema.CAPTION_TARGETS:
            target_to_seq.setdefault(target, seq)
    return _CaptionIndexer(target_to_seq) if target_to_seq else None


def _plain_run_el(text: str):
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    return r


def _prepend_caption_seq(para, seq_id: str, number: int) -> None:
    r"""Prepend ``<seq_id> {SEQ <seq_id> \* ARABIC}. `` to a caption paragraph.

    Emits a REAL Word SEQ field (cached value ``number``) so the template's
    ``TOC \c "<seq_id>"`` caption index collects this caption on a field update, and so
    the visible label matches the template's own caption convention (the label IS the
    seq identifier the template author chose - read from the comprehension, never a
    literal). The body is rebuilt from scratch each generation, so the cached number is
    deterministic and re-runs stay byte-identical.
    """
    p = para._p
    pPr = p.find(qn("w:pPr"))
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), f" SEQ {seq_id} \\* ARABIC ")
    fld.append(_plain_run_el(str(number)))
    nodes = [_plain_run_el(f"{seq_id} "), fld, _plain_run_el(". ")]
    if pPr is not None:
        ref = pPr
        for node in nodes:
            ref.addnext(node)
            ref = node
    else:
        for node in reversed(nodes):
            p.insert(0, node)


def _emit_caption(doc, resolver, runs, target, caption_ctx, findings):
    """Write a caption paragraph; when a kept caption index claims ``target``, prepend
    a label + SEQ field so the index regenerates from the new content.

    The SEQ prefix is added BEFORE the role style/appearance pass, so the label and
    separator runs carry the same brand caption styling as the caption text.
    """
    op = resolver.resolve_role("caption")
    para = _para_with_runs(
        doc, runs, _op_latin(op), resolver=resolver, findings=findings
    )
    seq = caption_ctx.seq_for(target) if caption_ctx is not None else None
    if seq:
        number = caption_ctx.emit(seq, textutil.runs_to_text(runs).strip())
        _prepend_caption_seq(para, seq, number)
    _apply_resolved_style(doc, para, op, findings)
    return para


def _write_block(
    doc,
    resolver: ProfileResolver,
    block: ir.Block,
    findings: list[Finding],
    caption_ctx: Optional["_CaptionIndexer"] = None,
) -> None:
    # Compute the resolved op first so the captured brand font reaches even the raw-
    # XML hyperlink runs (threaded via _op_latin into _add_runs); plain runs are also
    # branded by the post-write _apply_appearance pass, whose set-only-when-unset
    # guard makes the double-touch a no-op (re-runs stay byte-identical).
    if isinstance(block, ir.Heading):
        op = resolver.resolve_block(block)
        para = _para_with_runs(
            doc, block.runs, _op_latin(op), resolver=resolver, findings=findings
        )
        _apply_resolved_style(doc, para, op, findings)
    elif isinstance(block, ir.Paragraph):
        op = resolver.resolve_block(block)
        para = _para_with_runs(
            doc, block.runs, _op_latin(op), resolver=resolver, findings=findings
        )
        _apply_resolved_style(doc, para, op, findings)
    elif isinstance(block, ir.Callout):
        op = resolver.resolve_block(block)
        latin = _op_latin(op)
        para = doc.add_paragraph()
        if block.title:
            _add_runs(para, block.title, latin, resolver=resolver, findings=findings)
            para.add_run().add_break()  # title above body, same callout paragraph
        _add_runs(para, block.runs, latin, resolver=resolver, findings=findings)
        _apply_resolved_style(doc, para, op, findings)
    elif isinstance(block, ir.ListBlock):
        _write_list_items(doc, resolver, block, block.items, findings)
    elif isinstance(block, ir.Table):
        _write_table(doc, resolver, block, findings, caption_ctx)
    elif isinstance(block, ir.Caption):
        _emit_caption(doc, resolver, block.runs, block.target, caption_ctx, findings)
    elif isinstance(block, ir.Quote):
        op = resolver.resolve_block(block)
        para = _para_with_runs(
            doc, block.runs, _op_latin(op), resolver=resolver, findings=findings
        )
        _apply_resolved_style(doc, para, op, findings)
        if block.attribution:
            attr_op = resolver.resolve_role("paragraph")
            attr = _para_with_runs(
                doc,
                block.attribution,
                _op_latin(attr_op),
                resolver=resolver,
                findings=findings,
            )
            _apply_resolved_style(doc, attr, attr_op, findings)
    elif isinstance(block, ir.PageBreak):
        doc.add_page_break()
    elif isinstance(block, ir.Divider):
        # Native horizontal rule: an empty paragraph carrying a bottom border. Color
        # is OOXML "auto" (resolves to the theme's text color) - never a literal hex,
        # so the brand guarantee holds without a profile-defined divider artifact.
        para = doc.add_paragraph()
        pbdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "6")
        bottom.set(qn("w:space"), "1")
        bottom.set(qn("w:color"), "auto")
        pbdr.append(bottom)
        para._p.get_or_add_pPr().append(pbdr)
    elif isinstance(block, ir.Image):
        _write_image(doc, resolver, block, findings, caption_ctx)
    elif isinstance(block, ir.Kpi):
        _write_kpi(doc, resolver, block, findings)
    elif isinstance(block, ir.Chart):
        _write_chart(doc, block, findings)
    elif isinstance(block, ir.SmartArt):
        _write_smartart(doc, resolver, block, findings)
    elif isinstance(block, ir.Toc):
        _write_toc(doc, resolver, block, findings)
    elif block.TYPE in _UNHANDLED_BLOCK_TYPES:
        # No writer for this block. Skip cleanly - NEVER emit a blank ``Normal``
        # paragraph - and record a degradation finding so the skipped content is
        # visible in QA instead of silently lost. (component/section are normally
        # expanded away before this writer; reaching here is the defensive path.)
        findings.append(
            Finding(
                "block_degraded",
                schema.Severity.WARNING.value,
                f"{block.TYPE!r} block not rendered in docx vertical (skipped, no placeholder emitted)",
            )
        )
    else:
        # A genuinely unknown block type is a programming error, not a degradation:
        # fail loudly rather than dropping authored content or injecting a blank.
        raise GenerationError(f"unhandled block type {block.TYPE!r}")


def _write_toc(doc, resolver, block, findings) -> None:
    """Render an authored ``toc`` block as a native, updateable table of contents.

    If the shell already carries a structural TOC (preserved through the body
    clear), DEFER to it: that canonical TOC is refreshed by ``refresh_toc`` and a
    second one would duplicate it (the interplay this writer is careful to avoid).
    Otherwise author a native outline TOC field at the block's flow position; its
    visible cache is filled from the generated headings by
    ``refresh_visible_outline_toc_cache`` and it is marked dirty + ``updateFields``
    by ``refresh_toc`` (both run after the body is built), so Word rebuilds it on
    open. An optional title heading uses the resolved ``toc`` role style.
    """
    if structure.is_outline_toc_present(doc):
        findings.append(
            Finding(
                "toc_deferred",
                schema.Severity.INFO.value,
                "toc block deferred to the template's preserved outline table of "
                "contents (already present; no duplicate emitted)",
            )
        )
        return
    if block.title:
        para = doc.add_paragraph()
        para.add_run(str(block.title))
        _apply_resolved_style(doc, para, resolver.resolve_role("toc"), findings)
    structure.append_outline_toc_field(doc, max_level=block.max_level)


def _write_kpi(doc, resolver, block, findings) -> None:
    """Render a KPI / metric-card group as a brand-styled table (one row per metric:
    label, value, optional delta), value bolded for prominence.

    Reuses the table writer so the metric grid carries the profile's table style (no
    fabricated KPI box style/color): a faithful native rendering instead of a dropped
    block. The delta column is emitted only when at least one metric carries a delta.
    """
    items = block.items or []
    if not items:
        findings.append(
            Finding(
                "block_degraded",
                schema.Severity.WARNING.value,
                "'kpi' block had no items; skipped",
            )
        )
        return
    has_delta = any(getattr(k, "delta", None) for k in items)
    rows = []
    for k in items:
        cells = [
            ir.TableCell(runs=textutil.normalize_runs(k.label or "")),
            ir.TableCell(runs=[{"t": str(k.value or ""), "b": True}]),
        ]
        if has_delta:
            cells.append(ir.TableCell(runs=textutil.normalize_runs(k.delta or "")))
        rows.append(cells)
    _write_table(
        doc, resolver, ir.Table(columns=[], rows=rows, role="default"), findings
    )


def _write_image(doc, resolver, block, findings, caption_ctx=None) -> None:
    """Place an ``Image`` block as an inline picture, sized to the section's content
    width when no explicit size hint is given, with its caption below.

    Only an external ``src`` file is placed natively; an unresolved ``src``/``asset``
    (file missing, or only a profile asset id that cannot be loaded here) degrades to
    a loud ``block_degraded`` WARNING rather than crashing, so authored figures are
    realized when available and never silently lost otherwise. No brand literal is
    written: the picture is the author's asset, sized by layout geometry.
    """
    src = block.src
    placed = False
    if src and Path(src).is_file():
        sec = doc.sections[-1]
        width = (
            Emu(int(block.width_emu))
            if block.width_emu
            else Emu(structure.section_content_width_emu(sec))
        )
        height = Emu(int(block.height_emu)) if block.height_emu else None
        try:
            run = doc.add_paragraph().add_run()
            if height is not None:
                run.add_picture(src, width=width, height=height)
            else:
                run.add_picture(src, width=width)
            placed = True
        except Exception:
            # Any image-decode / placement failure degrades rather than crashing.
            placed = False
    if not placed:
        findings.append(
            Finding(
                "block_degraded",
                schema.Severity.WARNING.value,
                "'image' block not placed in docx (src/asset source unavailable); skipped",
            )
        )
        return
    if block.caption:
        _emit_caption(doc, resolver, block.caption, "figure", caption_ctx, findings)


_CHART_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
)
_CHART_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
)
_NS_CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"
_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
_NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _write_chart(doc, block, findings) -> None:
    """Author an ``ir.Chart`` as a NATIVE Word chart (a real DrawingML
    ``c:chartSpace`` part referenced by an inline ``w:drawing``), not flattened to
    text. The chart carries INLINE cached data (no embedded workbook), so it stays
    byte-idempotent, and writes no literal colors, so it inherits the document
    theme's accent colors (on-brand by construction). An unknown ``chart_type``
    falls back to a clustered column chart (surfaced as INFO); an empty/all-non-
    numeric chart degrades to a loud ``block_degraded`` WARNING - never a crash or a
    silent drop. ``title`` is authoritative when present.
    """
    if not chartlib.has_plottable_data(block):
        findings.append(
            Finding(
                "block_degraded",
                schema.Severity.WARNING.value,
                "'chart' block had no plottable numeric data; skipped",
            )
        )
        return
    if not chartlib.is_known_chart_type(block.chart_type):
        findings.append(
            Finding(
                "chart_type_fallback",
                schema.Severity.INFO.value,
                f"chart_type {block.chart_type!r} unknown; "
                "rendered as a clustered column chart",
            )
        )

    series = chartlib.coerce_series(block)
    if chartlib.is_single_series_type(block.chart_type) and len(series) > 1:
        findings.append(
            Finding(
                "chart_series_truncated",
                schema.Severity.WARNING.value,
                f"a {block.chart_type!r} chart renders only the first of "
                f"{len(series)} series; the others are not shown",
            )
        )
    xml = chartlib.build_chart_xml(
        block.chart_type, series, block.categories, block.title
    )

    # New chart part with a STABLE partname (derived from the count of existing
    # chart parts, so a re-run produces the same name) + the relationship from the
    # document part. python-docx auto-registers the content-type override at save.
    existing = sum(
        1
        for p in doc.part.package.iter_parts()
        if str(p.partname).startswith("/word/charts/chart")
    )
    partname = PackURI(f"/word/charts/chart{existing + 1}.xml")
    part = Part(partname, _CHART_CONTENT_TYPE, xml, doc.part.package)
    r_id = doc.part.relate_to(part, _CHART_REL_TYPE)

    # ``wp:docPr@id`` must be UNIQUE across every drawing object in the document
    # (images use python-docx's global max+1; a chart-only counter would collide
    # with an image's id and make Word refuse the file). Derive the next id from the
    # max of all existing ``wp:docPr`` ids - deterministic given the fixed
    # generation order, so byte-idempotency holds.
    doc_pr_ids = [
        int(el.get("id"))
        for el in doc.element.iter(qn("wp:docPr"))
        if (el.get("id") or "").isdigit()
    ]
    doc_pr_id = max(doc_pr_ids, default=0) + 1

    # Inline drawing sized to the section content width with a 3:2 aspect (matches
    # the image writer's layout-driven sizing; no fabricated coordinate). The
    # ``c:chart`` graphicData references the chart part by rId. The drawing goes into
    # a run of a paragraph added via ``add_paragraph`` so it lands BEFORE the body's
    # final ``w:sectPr`` (appending to the body directly would place it after, which
    # is invalid).
    sec = doc.sections[-1]
    cx = structure.section_content_width_emu(sec)
    cy = int(cx * 2 / 3)
    drawing = (
        f'<w:drawing xmlns:w="{_NS_W}" xmlns:wp="{_NS_WP}" xmlns:a="{_NS_A}" '
        f'xmlns:r="{_NS_R}" xmlns:c="{_NS_CHART}">'
        f'<wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/>'
        f'<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{doc_pr_id}" name="Chart {existing + 1}"/>'
        f"<wp:cNvGraphicFramePr/>"
        f'<a:graphic><a:graphicData uri="{_NS_CHART}">'
        f'<c:chart r:id="{r_id}"/>'
        f"</a:graphicData></a:graphic></wp:inline></w:drawing>"
    )
    run = doc.add_paragraph().add_run()
    run._r.append(parse_xml(drawing))


# SmartArt diagram families rendered as a single-COLUMN table (one row per node) vs
# the default single-ROW process strip. A FORMAT layout choice, not a brand value.
_SMARTART_LIST_DIAGRAMS = frozenset(
    {"list", "hierarchy", "pyramid", "table", "vertical_list", "bullet_list"}
)


def _smartart_node_text(node) -> str:
    """One node's label: its text plus any child texts inline (a docx table cell has
    no soft line break, so children are joined with ' - '), so nesting is not lost."""
    if not isinstance(node, dict):
        return str(node or "").strip()
    text = str(node.get("text") or "").strip()
    kids = [
        str(c.get("text") if isinstance(c, dict) else c or "").strip()
        for c in (node.get("children") or [])
    ]
    kids = [k for k in kids if k]
    if kids:
        text = f"{text} - {'; '.join(kids)}".strip(" -")
    return text


def _write_smartart(doc, resolver, block, findings) -> None:
    """Author an ``ir.SmartArt`` as a NATIVE brand-styled table, reusing the table
    writer: a process/flow becomes a single ROW (one cell per step), a list/
    hierarchy a single COLUMN (one row per node). On-brand via the profile's table
    style - never a fabricated diagram color. An empty diagram degrades to a loud
    ``block_degraded`` WARNING, never a silent drop. (The pptx vertical renders the
    same SmartArt as native chevron/box shapes.)"""
    labels = [t for t in (_smartart_node_text(n) for n in (block.nodes or [])) if t]
    if not labels:
        findings.append(
            Finding(
                "block_degraded",
                schema.Severity.WARNING.value,
                "'smartart' block had no nodes; skipped",
            )
        )
        return
    is_list = (block.diagram or "process").lower() in _SMARTART_LIST_DIAGRAMS
    if is_list:
        rows = [[ir.TableCell(runs=[{"t": label}])] for label in labels]
    else:
        rows = [[ir.TableCell(runs=[{"t": label}]) for label in labels]]
    _write_table(
        doc, resolver, ir.Table(columns=[], rows=rows, role="default"), findings
    )


def _write_list_items(doc, resolver, block, items, findings) -> None:
    """Write a list depth-first, one paragraph per item, threading each item's
    level into the resolver so nested items get their level-specific list role.

    Applies the resolved list PARAGRAPH style AND re-asserts the numbering
    (``w:pPr/w:numPr``) on each item (D1): a style's own ``w:numPr`` is not
    inherited onto a python-docx ``add_paragraph``, so without this the list would
    render as flat un-numbered paragraphs. The ``numId`` is the verbatim id the
    role nominated from the document's numbering part; the ``ilvl`` follows the
    item's actual nesting level so deeper items indent under the same definition.
    """
    for item in items:
        op = resolver.resolve_list_item(block, item)
        para = _para_with_runs(
            doc, item.runs, _op_latin(op), resolver=resolver, findings=findings
        )
        _apply_resolved_style(doc, para, op, findings)
        _apply_list_numbering(para, op, item)
        # Cluster D3: reference/clone the shell's numbering definition by id and re-assert
        # the captured per-level facts (numFmt / lvlText / indent) onto the output's own
        # w:abstractNum, set-only-when-unset. No-op for a profile with no captured numbering
        # (op_numbering returns None), so the no-numbering path is byte-identical.
        _apply_list_numbering_appearance(doc, op)
        if item.items:
            _write_list_items(doc, resolver, block, item.items, findings)


def _apply_list_numbering(para, op, item) -> None:
    """Write ``w:pPr/w:numPr`` (numId + ilvl) on a list paragraph from a resolved
    list role that carries a verbatim ``num_id``.

    No-op when the role carries no ``num_id`` (e.g. a ``Normal``-floor list with no
    real numbering definition, or a non-list role) - the paragraph then keeps only
    its style, exactly as before. The ``ilvl`` is the item's own nesting level so a
    nested item indents correctly even when its role fell back to the level-1
    list role.
    """
    resolver = op.resolver or {}
    num_id = resolver.get("num_id")
    if not num_id:
        return
    try:
        num_id_int = int(num_id)
    except (TypeError, ValueError):
        return
    ilvl = max(0, int(getattr(item, "level", 0) or 0))
    pPr = para._p.get_or_add_pPr()
    # Remove any inherited/duplicate numPr first so re-runs stay idempotent.
    for existing in pPr.findall(qn("w:numPr")):
        pPr.remove(existing)
    numPr = OxmlElement("w:numPr")
    ilvl_el = OxmlElement("w:ilvl")
    ilvl_el.set(qn("w:val"), str(ilvl))
    num_el = OxmlElement("w:numId")
    num_el.set(qn("w:val"), str(num_id_int))
    numPr.append(ilvl_el)
    numPr.append(num_el)
    pPr.append(numPr)


# The per-level ``w:ind`` attribute names captured/re-applied for Cluster D3 (mirrors
# ``structure._NUM_INDENT_ATTRS``). Each is set-only-when-unset on the cloned level.
_NUM_INDENT_ATTRS: tuple[str, ...] = ("left", "right", "firstLine", "hanging")


def _numbering_root_of(doc):
    """The output doc's ``w:numbering`` element, or ``None``. Crash-safe.

    The output package is opened FROM the shell, so its numbering part already carries
    the shell's own ``w:abstractNum``/``w:num`` definitions verbatim (python-docx copies
    the part on load). This surfaces that root so the apply side can re-assert per-level
    facts onto the cloned definition the output already holds."""
    try:
        part = doc.part.numbering_part
    except (KeyError, AttributeError, ValueError):
        return None
    if part is None:
        return None
    return getattr(part, "element", None)


def _ensure_numbering_def_present(doc, shell_doc, abstract_num_id: str) -> None:
    """Ensure the output's numbering part carries ``w:abstractNum[@abstract_num_id]``,
    cloning the SHELL's own definition by id when it is missing (Cluster D3).

    The engine NEVER synthesizes a numbering definition: it can only CLONE the shell's
    existing ``w:abstractNum`` verbatim (:func:`structure.clone_abstract_num`). In the
    normal path the output IS the shell (opened from ``shell_path``), so the def is
    already present and this is a no-op. Idempotent: a def already in the output is left
    untouched (never duplicated). No-op when either numbering part is absent or the shell
    declares no such id (the check layer rejects an undefined reference)."""
    out_root = _numbering_root_of(doc)
    if out_root is None or not abstract_num_id:
        return
    for an in out_root.findall(qn("w:abstractNum")):
        if an.get(qn("w:abstractNumId")) == str(abstract_num_id):
            return  # already present (idempotent)
    shell_root = _numbering_root_of(shell_doc) if shell_doc is not None else None
    clone = structure.clone_abstract_num(shell_root, abstract_num_id)
    if clone is None:
        return
    # Insert the cloned abstractNum before the first w:num (abstractNum precedes num in
    # the CT_Numbering child order), else append.
    first_num = out_root.find(qn("w:num"))
    if first_num is not None:
        first_num.addprevious(clone)
    else:
        out_root.append(clone)


# The WordprocessingML ``CT_Lvl`` child sequence (spec-fixed order, ECMA-376
# ``w:lvl``). A re-asserted level fact must be inserted BEFORE the first existing
# child that follows it in this order (e.g. ``w:lvlText`` precedes ``w:pPr``), or a
# strict OOXML reader rejects the cloned definition. The peer of
# :data:`_PPR_CHILD_ORDER` / :data:`_TBLPR_CHILD_ORDER` for the numbering part.
_LVL_CHILD_ORDER: tuple[str, ...] = (
    "w:start",
    "w:numFmt",
    "w:lvlRestart",
    "w:pStyle",
    "w:isLgl",
    "w:suff",
    "w:lvlText",
    "w:lvlPicBulletId",
    "w:legacy",
    "w:lvlJc",
    "w:pPr",
    "w:rPr",
)


def _insert_lvl_child_ordered(lvl, tag: str):
    """Get-or-create ``lvl/<tag>`` at the SPEC-CORRECT position in the ``CT_Lvl`` child
    sequence (so Word accepts the cloned numbering definition). Returns the existing
    element when present, else a new one inserted before the first existing successor
    in :data:`_LVL_CHILD_ORDER` (appended only when no successor exists)."""
    existing = lvl.find(qn(tag))
    if existing is not None:
        return existing
    el = OxmlElement(tag)
    successors = _LVL_CHILD_ORDER[_LVL_CHILD_ORDER.index(tag) + 1 :]
    for child in lvl:
        for succ in successors:
            if child.tag == qn(succ):
                child.addprevious(el)
                return el
    lvl.append(el)
    return el


def _reassert_level_facts(lvl, facts: dict) -> None:
    """Re-assert one level's captured facts (numFmt / lvlText / indent) onto its
    ``w:lvl`` element, SET-ONLY-WHEN-UNSET (Cluster D3).

    Each fact is written ONLY when the level does not already carry it directly, so an
    authored value (the shell's own, or a manual edit) is never clobbered and re-runs
    stay byte-identical. The engine writes only VALUES the profile captured from the
    shell (never synthesized): ``w:numFmt@w:val`` / ``w:lvlText@w:val`` are set on the
    existing or a freshly-created child (created at its SPEC-CORRECT ``CT_Lvl``
    position via :func:`_insert_lvl_child_ordered`); each ``w:ind`` attribute is set
    on the level's ``w:pPr/w:ind`` set-only-when-unset."""
    numfmt = facts.get("numFmt")
    if numfmt is not None:
        el = _insert_lvl_child_ordered(lvl, "w:numFmt")
        if el.get(qn("w:val")) is None:
            el.set(qn("w:val"), str(numfmt))
    lvltext = facts.get("lvlText")
    if lvltext is not None:
        el = _insert_lvl_child_ordered(lvl, "w:lvlText")
        if el.get(qn("w:val")) is None:
            el.set(qn("w:val"), str(lvltext))
    indent = facts.get("indent") or {}
    if indent:
        ppr = _insert_lvl_child_ordered(lvl, "w:pPr")
        ind = ppr.find(qn("w:ind"))
        if ind is None:
            ind = OxmlElement("w:ind")
            ppr.append(ind)
        for attr in _NUM_INDENT_ATTRS:
            value = indent.get(attr)
            if value is None:
                continue
            if ind.get(qn(f"w:{attr}")) is not None:
                continue  # authored indent is never clobbered (set-only-when-unset)
            try:
                ind.set(qn(f"w:{attr}"), str(int(value)))
            except (TypeError, ValueError):
                continue


def _apply_list_numbering_appearance(doc, op, shell_doc=None) -> None:
    """Apply the captured NUMBERING facts (Cluster D3, docx-only) onto the output's own
    numbering definition: clone the shell's ``w:abstractNum`` by id when the output lacks
    it, then re-assert each per-level ``numFmt``/``lvlText``/indent SET-ONLY-WHEN-UNSET.

    The captured ``num_id``/``abstract_num_id`` are SYMBOLIC references the engine NEVER
    invents; the per-level facts came from the shell's own abstractNum and are re-asserted
    only when the output's level does not already declare them (the authored/inherited
    value wins). A profile with no captured numbering never reaches the per-level
    re-assert (``op_numbering`` returns ``None``), so the no-numbering path is a
    byte-identical no-op. The output is opened FROM the shell, so the def is normally
    present and the clone is a no-op; ``shell_doc`` lets a caller clone from a separately
    opened shell when needed (kept optional for the byte-identical common path)."""
    numbering = appearance.op_numbering(op)
    if not numbering:
        return
    abstract_num_id = numbering.get("abstract_num_id")
    if abstract_num_id:
        _ensure_numbering_def_present(doc, shell_doc, abstract_num_id)
    per_level = numbering.get("per_level_facts") or {}
    if not per_level:
        return
    out_root = _numbering_root_of(doc)
    if out_root is None:
        return
    target = None
    for an in out_root.findall(qn("w:abstractNum")):
        if an.get(qn("w:abstractNumId")) == str(abstract_num_id):
            target = an
            break
    if target is None:
        return
    levels = {}
    for lvl in target.findall(qn("w:lvl")):
        try:
            levels[int(lvl.get(qn("w:ilvl")) or 0)] = lvl
        except (TypeError, ValueError):
            continue
    for ilvl, facts in per_level.items():
        try:
            il = int(ilvl)
        except (TypeError, ValueError):
            continue
        lvl = levels.get(il)
        if lvl is not None and isinstance(facts, dict):
            _reassert_level_facts(lvl, facts)


def _write_table(doc, resolver, block, findings, caption_ctx=None) -> None:
    # Honor colspan/rowspan: size the grid by span-expanded width, then merge.
    # Header cells are a flat run list (one column each), so the header width is
    # simply the column count.
    body_widths = [sum(max(1, c.colspan) for c in row) for row in block.rows]
    header_width = len(block.columns)
    cols = max([header_width] + body_widths + [1])
    rows = len(block.rows) + (1 if block.columns else 0)
    table = doc.add_table(rows=rows, cols=cols)
    # The table.default ResolvedOp ALREADY carries appearance (the role's own font, or
    # the document body font as the resolver's fallback). A cell paragraph carries no
    # python-docx style, so its runs are never reached by ``_apply_appearance``; brand
    # them at write time from this one resolved op (never a fabricated cell font).
    # KPI / SmartArt synthetic tables route through here, so they inherit it for free.
    table_op = resolver.resolve_block(block)
    _apply_table_style(doc, table, table_op, findings)
    cell_latin = _op_latin(table_op)

    header_op = resolver.resolve_role(
        schema.role_id("table", (block.role or "default"), "header"), fallback=None
    )
    # A header cell prefers the header role's own captured font (role wins), falling
    # back to the table-body font; both come only from the resolver.
    header_latin = _op_latin(header_op) or cell_latin
    row_offset = 0
    if block.columns:
        _fill_row(
            doc,
            table,
            0,
            [_as_cell(c) for c in block.columns],
            header_op,
            findings,
            cell_latin=cell_latin,
            header_latin=header_latin,
            force_header=True,
            resolver=resolver,
        )
        row_offset = 1
    for r_idx, row in enumerate(block.rows):
        _fill_row(
            doc,
            table,
            r_idx + row_offset,
            row,
            header_op,
            findings,
            cell_latin=cell_latin,
            header_latin=header_latin,
            resolver=resolver,
        )
    if block.caption:
        _emit_caption(doc, resolver, block.caption, "table", caption_ctx, findings)


def _fill_row(
    doc,
    table,
    r_idx,
    cells,
    header_op,
    findings,
    *,
    cell_latin=None,
    header_latin=None,
    force_header=False,
    resolver: Optional[ProfileResolver] = None,
) -> None:
    """Fill one logical row, honoring colspan/rowspan by merging grid cells.

    Each cell paragraph's runs are branded with the captured font as they are written
    (``cell_latin`` for body cells, ``header_latin`` for header cells - both resolved
    from the profile by ``_write_table``, never a literal), so cell text, hyperlinks
    included, carries the brand typeface.

    CONTRACT - table cells receive the FONT axis only. Run SIZE and COLOR follow the
    resolver's family gate (``_allows_body_default``): the document body size/color
    default is intentionally NOT pushed onto the ``table.*`` family - exactly as for
    ``heading.*`` - so a table never has the body size/color forced over its own table
    style's intrinsic size/color. A ``table``/header role that itself captured a
    size/color still gets it (role-specific values apply to any role) via the header
    cell's ``_apply_resolved_style`` -> ``_apply_appearance`` pass."""
    c_cursor = 0
    ncols = len(table.columns)
    for cell in cells:
        if c_cursor >= ncols:
            break
        cspan = max(1, getattr(cell, "colspan", 1))
        rspan = max(1, getattr(cell, "rowspan", 1))
        anchor = table.cell(r_idx, c_cursor)
        # Merge the spanned rectangle so the column grid stays aligned.
        if cspan > 1 or rspan > 1:
            end_c = min(c_cursor + cspan - 1, ncols - 1)
            end_r = min(r_idx + rspan - 1, len(table.rows) - 1)
            anchor = anchor.merge(table.cell(end_r, end_c))
        is_header = force_header or getattr(cell, "header", False)
        # Run-aware cell text: a fresh cell has one empty paragraph; write the IR
        # runs into it so inline emphasis/links survive (not a flattened cell.text).
        _add_runs(
            anchor.paragraphs[0],
            cell.runs,
            header_latin if is_header else cell_latin,
            resolver=resolver,
            findings=findings,
        )
        if is_header and header_op is not None and header_op.resolver:
            for para in anchor.paragraphs:
                _apply_resolved_style(doc, para, header_op, findings)
        c_cursor += cspan


def _as_cell(col):
    """Wrap a header column in a ``TableCell`` shim. ``col`` is a run-list (the
    shape ``Table.from_dict`` produces, so multi-run header emphasis survives); a
    bare run dict / string from a direct construction is tolerated."""
    runs = col if isinstance(col, list) else [col]
    return ir.TableCell(runs=runs)


def _apply_style(
    doc,
    target_obj,
    op,
    findings: list[Finding],
    *,
    label: str,
    expect_style: bool = False,
) -> None:
    """Apply a resolved role's style to a paragraph OR a table (both expose
    ``.style``). A role that names a resolver but resolves to no shell style is a
    brand breach: be LOUD (ERROR) instead of silently leaving the default style.
    When ``expect_style`` is set (a block family that should carry a dedicated brand
    style, e.g. a table), a STUB resolver - the profile has no such role at all - is
    surfaced as an INFO ``style_fallback`` (the block still renders, with the default
    style) so the missing brand style is visible in QA rather than silently dropped.
    The single implementation behind ``_apply_resolved_style`` / ``_apply_table_style``."""
    style = lookup_style(doc, op.resolver)
    if style is not None:
        target_obj.style = style
    elif op.resolver:
        target = op.resolver.get("style_id") or op.resolver.get("style_name")
        findings.append(
            Finding(
                "resolver_targets_exist",
                schema.Severity.ERROR.value,
                f"{label}role {op.role_id!r} resolves to style {target!r} "
                "which is not in the shell",
                location=op.role_id,
            )
        )
    elif expect_style:
        findings.append(
            Finding(
                "style_fallback",
                schema.Severity.INFO.value,
                f"{label}block rendered with the default style: this profile has no "
                f"{label.strip() or 'matching'} role to brand it",
                location=op.role_id,
            )
        )
    # Apply the role's captured typography (font/size/color) on top of the style.
    # No-op for tables (no runs here) and for profiles with no captured appearance.
    _apply_appearance(target_obj, op, findings)


# The captured-axis readers now live in the format-neutral ``common.appearance``
# engine. docx still reads the font axis directly (``_op_latin`` is threaded into
# ``_add_runs`` for the raw-XML hyperlink font, the one path outside the shared
# orchestration); the size/color axes are applied only through
# ``appearance.apply_role_appearance``, so no docx-local size/color alias is needed.
_op_latin = appearance.op_latin


# The WordprocessingML themeColor tokens python-docx emits (``MSO_THEME_COLOR``
# ``.xml_value``) keyed back to their enum member, so a captured theme-color token
# can be applied via ``run.font.color.theme_color``. A CLOSED table (the same
# tokens ``_run_color`` captures); a token outside it is SKIPPED at apply time
# rather than risking a raise.
_WML_TOKEN_TO_THEME_COLOR = {
    member.xml_value: member
    for member in MSO_THEME_COLOR
    if getattr(member, "xml_value", None)
    and member is not MSO_THEME_COLOR.NOT_THEME_COLOR
}


def _brand_run_font(run, latin: Optional[str]) -> None:
    """Set ``run.font.name`` to the captured ``latin`` ONLY when the run carries no
    explicit font (the exact v1 guard).

    The IR never sets fonts, so authored runs are unfonted and inherit - we brand
    them; a run that already carries an explicit font is left alone, which also keeps
    re-runs byte-identical (the name is set only when absent). A falsy ``latin`` (a
    pre-capture profile) is a no-op."""
    if latin and run.font.name is None:
        run.font.name = latin


def _brand_run_color(run, color: Optional[dict], findings: list[Finding]) -> None:
    """Set ``run.font.color`` from the captured ``color`` object ONLY when the run
    carries no explicit color (``run.font.color.type is None``) (the per-axis guard).

    A guarded wrapper over :func:`_set_run_color` for the docx-only call sites that
    apply a color outside the shared orchestration (the unsafe-url link fallback in
    :func:`_add_hyperlink`). The set-only-when-unset guard for the orchestration-driven
    paths lives in :meth:`_DocxAppearanceBackend.color_unset`; this wrapper carries the
    SAME ``run.font.color.type is None`` guard so the two paths stay identical. The
    value is read STRICTLY from the resolver op (never a literal); a falsy ``color`` is
    a no-op."""
    if not color or run.font.color.type is not None:
        return
    _set_run_color(run, color, findings)


def _set_run_color(run, color: dict, findings: list[Finding]) -> None:
    """Write ``run.font.color`` from the captured ``color`` object UNCONDITIONALLY
    (the caller has already confirmed the run carries no explicit color).

    A hex color is applied via ``run.font.color.rgb``; a theme-token color is mapped
    through the CLOSED :data:`_WML_TOKEN_TO_THEME_COLOR` table to an
    ``MSO_THEME_COLOR`` member and applied via ``run.font.color.theme_color``. A
    token that does not map is SKIPPED with an INFO finding rather than letting a bad
    apply raise. The value is read STRICTLY from the resolver op (never a literal)."""
    kind = color.get("kind")
    if kind == "hex":
        hexval = color.get("hex")
        if hexval:
            # Normalize (#rrggbb / 3-digit / any case) exactly as the verify side
            # does, and fail CLOSED on a truly malformed hex (INFO skip, never a
            # raise that would abort generation) - mirrors the theme-token branch.
            try:
                run.font.color.rgb = RGBColor.from_string(
                    colorutil.normalize_hex(hexval)
                )
            except (ValueError, AttributeError):
                findings.append(
                    Finding(
                        "appearance_color_skipped",
                        schema.Severity.INFO.value,
                        f"captured hex color {hexval!r} is not a valid RRGGBB; "
                        "left inherited",
                    )
                )
        return
    if kind == "theme":
        token = color.get("theme")
        member = _WML_TOKEN_TO_THEME_COLOR.get(token)
        if member is None:
            # A clrScheme-slot token (dk1/lt1/hlink ...) has no WordprocessingML
            # themeColor member, but the resolver enriched it with the concrete hex:
            # realize the brand color via the hex rather than dropping it to inherited.
            hexval = color.get("hex")
            if hexval:
                try:
                    run.font.color.rgb = RGBColor.from_string(
                        colorutil.normalize_hex(hexval)
                    )
                    return
                except (ValueError, AttributeError):
                    pass
            findings.append(
                Finding(
                    "appearance_color_skipped",
                    schema.Severity.INFO.value,
                    f"captured theme color token {token!r} does not map to a "
                    "WordprocessingML theme color and has no resolvable hex; "
                    "left inherited",
                )
            )
            return
        run.font.color.theme_color = member
        # Carry the resolved hex in w:color@w:val ALONGSIDE the themeColor so a
        # renderer that ignores a run-level themeColor (headless LibreOffice) still
        # shows the brand color instead of the black fallback; Word still uses the
        # themeColor. The hex comes only from the resolved palette ref (no literal).
        hexval = color.get("hex")
        if hexval:
            try:
                normalized = colorutil.normalize_hex(hexval)
            except (ValueError, AttributeError):
                normalized = None
            if normalized:
                clr = run._element.get_or_add_rPr().find(qn("w:color"))
                if clr is not None:
                    clr.set(qn("w:val"), normalized)


def _all_para_runs(target_obj) -> list:
    """Every run in a paragraph, INCLUDING runs nested in a ``w:hyperlink``.

    A hyperlink run is raw ``w:r`` XML under ``w:hyperlink`` and is NOT in
    ``paragraph.runs`` - but modern python-docx exposes it via
    ``paragraph.hyperlinks[*].runs``. Branding those too is what gives a link the
    SAME size (and font/color) as the surrounding body text instead of leaving it at
    the inherited default. Crash-safe: an object without ``hyperlinks`` (a table or
    an older python-docx) just yields the direct runs."""
    runs = list(getattr(target_obj, "runs", None) or [])
    try:
        for link in getattr(target_obj, "hyperlinks", None) or []:
            runs.extend(getattr(link, "runs", None) or [])
    except Exception:
        pass
    return runs


# ---------------------------------------------------------------------------
# Paragraph GEOMETRY apply (Cluster D1, DOCX-ONLY).
# ---------------------------------------------------------------------------
# Geometry rides on ``w:pPr`` (the PARAGRAPH), so it is applied set-only-when-unset per
# PROPERTY: each spacing/indent attribute is written only when ``w:pPr`` does not
# already carry it directly, borders only when the side is absent, shading only when no
# ``w:shd@w:fill`` is set. The captured border element is re-parsed and DEEP-COPIED
# verbatim (never hand-built), so apply re-emits the exact structure capture recorded.
# A profile with no captured geometry never reaches ``set_geometry`` (the op carries no
# ``appearance.geometry``), so the no-geometry path is byte-identical.
_GEOMETRY_BORDER_SIDES: tuple[str, ...] = ("top", "bottom", "left", "right")

# The WordprocessingML ``CT_PPr`` child sequence (spec-fixed order). A geometry element
# must be inserted BEFORE the first existing child that follows it in this order, or
# Word rejects the document. python-docx wires this order for ``w:spacing`` / ``w:ind``
# (``get_or_add_spacing`` / ``get_or_add_ind``) but NOT for ``w:pBdr`` / ``w:shd``, so
# those two are inserted via :func:`_insert_ppr_child_ordered`.
_PPR_CHILD_ORDER: tuple[str, ...] = (
    "w:pStyle",
    "w:keepNext",
    "w:keepLines",
    "w:pageBreakBefore",
    "w:framePr",
    "w:widowControl",
    "w:numPr",
    "w:suppressLineNumbers",
    "w:pBdr",
    "w:shd",
    "w:tabs",
    "w:suppressAutoHyphens",
    "w:kinsoku",
    "w:wordWrap",
    "w:overflowPunct",
    "w:topLinePunct",
    "w:autoSpaceDE",
    "w:autoSpaceDN",
    "w:bidi",
    "w:adjustRightInd",
    "w:snapToGrid",
    "w:spacing",
    "w:ind",
    "w:contextualSpacing",
    "w:mirrorIndents",
    "w:suppressOverlap",
    "w:jc",
    "w:textDirection",
    "w:textAlignment",
    "w:textboxTightWrap",
    "w:outlineLvl",
    "w:divId",
    "w:cnfStyle",
    "w:rPr",
    "w:sectPr",
    "w:pPrChange",
)


def _insert_ppr_child_ordered(ppr, tag: str):
    """Get-or-create ``ppr/<tag>`` at the SPEC-CORRECT position in the ``CT_PPr`` child
    sequence (so Word accepts the document). Returns the existing element when present,
    else a new one inserted before the first existing successor in
    :data:`_PPR_CHILD_ORDER` (appended only when no successor exists)."""
    existing = ppr.find(qn(tag))
    if existing is not None:
        return existing
    el = OxmlElement(tag)
    successors = _PPR_CHILD_ORDER[_PPR_CHILD_ORDER.index(tag) + 1 :]
    for child in ppr:
        for succ in successors:
            if child.tag == qn(succ):
                child.addprevious(el)
                return el
    ppr.append(el)
    return el


def _set_twips_if_unset(el, attr: str, value) -> None:
    """Set ``el@w:<attr>`` to ``value`` ONLY when the attribute is currently absent.

    The set-only-when-unset guard at the single-attribute level, so an authored
    spacing/indent attribute is never clobbered. A non-int ``value`` is skipped (the
    captured value is always an int twips; this is a defensive guard)."""
    if value is None or el.get(qn(f"w:{attr}")) is not None:
        return
    try:
        el.set(qn(f"w:{attr}"), str(int(value)))
    except (TypeError, ValueError):
        return


def _apply_spacing(ppr, spacing: dict) -> None:
    """Apply the captured ``spacing`` sub-dict to ``w:pPr/w:spacing``, per-attribute
    set-only-when-unset. The ``w:spacing`` element is created (in spec order, via
    python-docx's ordered ``get_or_add_spacing``) if absent."""
    if not spacing:
        return
    el = ppr.get_or_add_spacing()
    _set_twips_if_unset(el, "before", spacing.get("before_twips"))
    _set_twips_if_unset(el, "after", spacing.get("after_twips"))
    _set_twips_if_unset(el, "line", spacing.get("line_twips"))
    line_rule = spacing.get("line_rule")
    if line_rule is not None and el.get(qn("w:lineRule")) is None:
        el.set(qn("w:lineRule"), str(line_rule))


def _apply_indent(ppr, indentation: dict) -> None:
    """Apply the captured ``indentation`` sub-dict to ``w:pPr/w:ind``, per-attribute
    set-only-when-unset. The ``w:ind`` element is created (in spec order, via
    python-docx's ordered ``get_or_add_ind``) if absent."""
    if not indentation:
        return
    el = ppr.get_or_add_ind()
    _set_twips_if_unset(el, "left", indentation.get("left_twips"))
    _set_twips_if_unset(el, "right", indentation.get("right_twips"))
    _set_twips_if_unset(el, "firstLine", indentation.get("first_line_twips"))
    _set_twips_if_unset(el, "hanging", indentation.get("hanging_twips"))


def _apply_borders(ppr, borders: dict) -> None:
    """Apply the captured ``borders`` sub-dict to ``w:pPr/w:pBdr``, per-SIDE
    set-only-when-unset. ``w:pBdr`` is created at its spec position; each side element
    is RE-PARSED from its serialized copy and appended verbatim (the exact structure
    capture recorded), only when the side is absent. A malformed serialized side is
    skipped (never crashes the write)."""
    if not borders:
        return
    pbdr = _insert_ppr_child_ordered(ppr, "w:pBdr")
    for side in _GEOMETRY_BORDER_SIDES:
        serialized = borders.get(side)
        if not serialized or pbdr.find(qn(f"w:{side}")) is not None:
            continue
        try:
            el = parse_xml(serialized)
        except Exception:
            continue
        pbdr.append(el)


def _apply_shading(ppr, shading: dict) -> None:
    """Apply the captured ``shading.fill_hex`` to ``w:pPr/w:shd@w:fill``, set-only-when-
    unset (the whole ``w:shd`` is left untouched if it already carries an explicit
    fill). The ``w:shd`` is created at its spec position; a clear shading is written
    with the spec-required ``w:val='clear'``."""
    fill = (shading or {}).get("fill_hex")
    if not fill:
        return
    shd = ppr.find(qn("w:shd"))
    if shd is not None and shd.get(qn("w:fill")) is not None:
        return
    try:
        normalized = colorutil.normalize_hex(fill)
    except (ValueError, AttributeError):
        return
    if shd is None:
        shd = _insert_ppr_child_ordered(ppr, "w:shd")
    if shd.get(qn("w:val")) is None:
        shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), normalized)


def _apply_paragraph_geometry(para, geometry: dict) -> None:
    """Apply the captured ``geometry`` dict to ``para``'s ``w:pPr``, set-only-when-unset
    per property (Cluster D1). The ``confidence`` sub-dict (capture provenance) is not
    applied. A paragraph with no ``w:pPr`` gets one created lazily by the first axis
    that has a captured value."""
    if not geometry:
        return
    ppr = para._p.get_or_add_pPr()
    _apply_spacing(ppr, geometry.get("spacing") or {})
    _apply_indent(ppr, geometry.get("indentation") or {})
    _apply_borders(ppr, geometry.get("borders") or {})
    _apply_shading(ppr, geometry.get("shading") or {})


def _paragraphs_of(target_obj) -> list:
    """The paragraph(s) of ``target_obj`` that carry brand geometry.

    A docx paragraph IS the geometry target (it exposes ``_p``); a table here exposes
    no paragraph-geometry target, so it yields nothing (geometry is a paragraph axis).
    Crash-safe: an object without a raw ``_p`` element yields nothing."""
    if getattr(target_obj, "_p", None) is not None:
        return [target_obj]
    return []


class _DocxAppearanceBackend:
    """The docx hook set the format-neutral ``common.appearance`` orchestration drives.

    It wraps the docx-specific run reader and per-axis probes/writers VERBATIM -
    ``_all_para_runs`` (paragraph runs incl. raw-XML hyperlink runs), the
    ``run.font.<axis> is None`` set-only-when-unset probes, and the unconditional
    per-axis writes (``run.font.name = latin`` / ``Pt(half_pts / 2)`` /
    :func:`_set_run_color`) - so the shared control flow (probe, then write only when
    unset) produces byte-identical docx output, exactly the inlined v1 behavior. The
    docx-only raw-XML hyperlink injection (``_inject_hyperlink_run_color``) and the
    WordprocessingML token map stay outside this backend.

    ``realized_axes`` declares the parity ledger (Cluster E3): docx realizes ALL six
    appearance axes - font/size/color here, geometry via ``set_geometry``, and
    table/numbering via the docx writers outside this orchestration
    (``_apply_table_style`` / ``_apply_list_numbering_appearance``) - so no
    ``appearance_apply_degraded`` finding ever fires for docx today.

    Cluster D1 adds the PARAGRAPH geometry hooks (``paragraphs_of`` + ``set_geometry``):
    geometry rides on ``w:pPr`` (the paragraph, not the run), applied set-only-when-
    unset per property. A profile with no captured geometry never reaches them (the op
    carries no ``appearance.geometry``), so the no-geometry path is byte-identical."""

    # Parity ledger (Cluster E3): every appearance axis is realized for docx.
    realized_axes = frozenset(
        {"font", "size_hp", "color", "geometry", "table", "numbering"}
    )

    def runs_of(self, target):
        return _all_para_runs(target)

    def font_unset(self, run) -> bool:
        return run.font.name is None

    def set_font(self, run, latin: str) -> None:
        run.font.name = latin

    def size_unset(self, run) -> bool:
        return run.font.size is None

    def set_size(self, run, half_pts: int) -> None:
        run.font.size = Pt(half_pts / 2)

    def color_unset(self, run) -> bool:
        return run.font.color.type is None

    def set_color(self, run, ref: dict, findings) -> None:
        _set_run_color(run, ref, findings)

    def paragraphs_of(self, target):
        return _paragraphs_of(target)

    def set_geometry(self, para, geometry: dict) -> None:
        _apply_paragraph_geometry(para, geometry)


DOCX_BACKEND = _DocxAppearanceBackend()


def _apply_appearance(target_obj, op, findings: list[Finding]) -> None:
    """Apply captured brand typography (font, size, color) from the profile as direct
    run formatting on a paragraph's runs (hyperlink runs included).

    A thin docx wrapper over the format-neutral
    ``common.appearance.apply_role_appearance`` driven by :data:`DOCX_BACKEND`. The
    three axes stay INDEPENDENT (each applied only when the run's attribute is unset,
    via the backend's ``_brand_run_*`` guards), hyperlink runs are reached via
    :func:`_all_para_runs`, and a target with no runs (a docx table here) is skipped -
    output is byte-identical to the inlined v1 loop."""
    appearance.apply_role_appearance(DOCX_BACKEND, target_obj, op, findings)


def _apply_resolved_style(doc, para, op, findings: list[Finding]) -> None:
    _apply_style(doc, para, op, findings, label="")


# The spec child order of ``CT_TblPr`` (``w:tblPr``), enough of it to place the three
# table-appearance elements (``w:tblStyle`` first, ``w:tblCellMar`` then ``w:tblLook``
# near the end) at a position Word accepts. Only the elements we (or the template) might
# carry need to be listed before/after the ones we insert.
_TBLPR_CHILD_ORDER: tuple[str, ...] = (
    "w:tblStyle",
    "w:tblpPr",
    "w:tblOverlap",
    "w:bidiVisual",
    "w:tblStyleRowBandSize",
    "w:tblStyleColBandSize",
    "w:tblW",
    "w:jc",
    "w:tblCellSpacing",
    "w:tblInd",
    "w:tblBorders",
    "w:shd",
    "w:tblLayout",
    "w:tblCellMar",
    "w:tblLook",
    "w:tblCaption",
    "w:tblDescription",
)


def _insert_tblpr_child_ordered(tblpr, tag: str):
    """Get-or-create ``tblpr/<tag>`` at the SPEC-CORRECT position in the ``CT_TblPr``
    child sequence (so Word accepts the document). Returns the existing element when
    present, else a new one inserted before the first existing successor in
    :data:`_TBLPR_CHILD_ORDER` (appended only when no successor exists)."""
    existing = tblpr.find(qn(tag))
    if existing is not None:
        return existing
    el = OxmlElement(tag)
    successors = _TBLPR_CHILD_ORDER[_TBLPR_CHILD_ORDER.index(tag) + 1 :]
    for child in tblpr:
        for succ in successors:
            if child.tag == qn(succ):
                child.addprevious(el)
                return el
    tblpr.append(el)
    return el


# python-docx's SYNTHETIC default ``w:tblLook@w:val`` (Word's "firstRow|firstColumn|
# noVBand" template default), injected on EVERY ``doc.add_table`` regardless of the brand.
# It is NOT an authored brand value, so the captured tblLook may replace it; an authored
# value that DIFFERS from this default is treated as set and never clobbered.
_TBLLOOK_DOCX_DEFAULT = 0x04A0

# The six ``w:tblLook`` per-flag attribute names and their bits in the ``w:val`` bitmask
# (the same mapping the capture reader uses). Word writes BOTH forms; emitting the legacy
# attributes alongside ``w:val`` keeps older readers in sync.
_TBLLOOK_FLAG_ATTRS: tuple[tuple[str, int], ...] = (
    ("firstRow", 0x0020),
    ("lastRow", 0x0040),
    ("firstColumn", 0x0080),
    ("lastColumn", 0x0100),
    ("noHBand", 0x0200),
    ("noVBand", 0x0400),
)


def _tbllook_is_unset(tblpr) -> bool:
    """``True`` when the table carries NO authored ``w:tblLook`` - i.e. it is absent OR
    is exactly python-docx's synthetic :data:`_TBLLOOK_DOCX_DEFAULT` (which the library
    injects on every fresh ``add_table``, never a brand decision). An authored bitmask
    that differs from the synthetic default is treated as SET (set-only-when-unset)."""
    look = tblpr.find(qn("w:tblLook"))
    if look is None:
        return True
    val = look.get(qn("w:val"))
    if val is None:
        return True
    try:
        return int(val, 16) == _TBLLOOK_DOCX_DEFAULT
    except (TypeError, ValueError):
        return False


# The ``w:tblCellMar`` side -> captured ``cell_margins`` field, the same naming the
# capture reader (``formats/docx/typography._TABLE_CELL_MARGIN_SIDES``) uses.
_TABLE_CELL_MARGIN_SIDES: tuple[tuple[str, str], ...] = (
    ("top", "top_twips"),
    ("bottom", "bottom_twips"),
    ("left", "left_twips"),
    ("right", "right_twips"),
)


def _apply_table_cell_margins(tblpr, cell_margins: dict) -> None:
    """Apply the captured ``cell_margins`` twips to ``w:tblPr/w:tblCellMar``, per-side
    set-only-when-unset. Each side (``w:{top,bottom,left,right}``) is written ONLY when
    the table does not already carry that side's margin, so an authored margin is never
    clobbered. The ``w:tblCellMar`` (and each side element) is created in spec order."""
    if not cell_margins:
        return
    cell_mar = None
    for side, field in _TABLE_CELL_MARGIN_SIDES:
        value = cell_margins.get(field)
        if value is None:
            continue
        if cell_mar is None:
            cell_mar = _insert_tblpr_child_ordered(tblpr, "w:tblCellMar")
        side_el = cell_mar.find(qn(f"w:{side}"))
        if side_el is not None:
            # An authored side margin (carries an explicit width) is never clobbered.
            if side_el.get(qn("w:w")) is not None:
                continue
        else:
            side_el = OxmlElement(f"w:{side}")
            cell_mar.append(side_el)
        try:
            side_el.set(qn("w:w"), str(int(value)))
        except (TypeError, ValueError):
            continue
        # ``w:tblCellMar`` side widths are twips (``w:type='dxa'``), set only when the
        # element did not already declare its own type.
        if side_el.get(qn("w:type")) is None:
            side_el.set(qn("w:type"), "dxa")


def _apply_table_appearance(table, op) -> None:
    """Apply the captured TABLE conditional-format facts (Cluster D2, docx-only) onto a
    generated table's ``w:tblPr``, every fact SET-ONLY-WHEN-UNSET.

    - ``style_id``: the table's ``w:tblStyle@w:val`` is written ONLY when the table does
      not already reference a style. (``_apply_style`` already set ``table.style`` from
      the resolver's named style; this is a defensive re-assert that never overrides an
      authored reference, so it is normally a no-op.)
    - ``tblLook``: the captured bitmask is written to ``w:tblLook@w:val`` (hex) ONLY when
      the table carries no ``w:tblLook`` - it merely ENABLES the shell style's OWN
      ``w:tblStylePr`` conditional formats (banding / first-last emphasis), never a fill.
    - ``cell_margins``: each side is written set-only-when-unset (see
      :func:`_apply_table_cell_margins`).

    The engine NEVER authors a ``w:tblStylePr``, a fill, or a border here: the band fills
    live in the shell's styles part and are toggled via the bitmask + the style
    reference. A table that already carries an authored ``tblLook`` / style / margin is
    left untouched. A profile with no captured table appearance never reaches this
    function (``op_table`` returns ``None``), so the no-table path is byte-identical."""
    table_appearance = appearance.op_table(op)
    if not table_appearance:
        return
    # ``CT_Tbl.tblPr`` is a python-docx auto-creating property (the ``w:tblPr`` is the
    # first child of ``w:tbl`` and is created on first access if absent).
    tblpr = table._tbl.tblPr

    style_id = table_appearance.get("style_id")
    if style_id and tblpr.find(qn("w:tblStyle")) is None:
        style_el = _insert_tblpr_child_ordered(tblpr, "w:tblStyle")
        style_el.set(qn("w:val"), str(style_id))

    tbllook = table_appearance.get("tblLook")
    if tbllook is not None and _tbllook_is_unset(tblpr):
        try:
            bits = int(tbllook)
        except (TypeError, ValueError):
            bits = None
        if bits is not None:
            # Drop python-docx's synthetic-default ``w:tblLook`` (it carries stale legacy
            # per-flag attributes) so the captured bitmask is the single source of truth.
            stale = tblpr.find(qn("w:tblLook"))
            if stale is not None:
                tblpr.remove(stale)
            look_el = _insert_tblpr_child_ordered(tblpr, "w:tblLook")
            # Write BOTH the modern ``w:val`` (hex) and the legacy per-flag attributes,
            # exactly as Word does, so old and new readers agree on the toggles.
            look_el.set(qn("w:val"), format(bits, "04X"))
            for attr, bit in _TBLLOOK_FLAG_ATTRS:
                look_el.set(qn(f"w:{attr}"), "1" if (bits & bit) else "0")

    _apply_table_cell_margins(tblpr, table_appearance.get("cell_margins") or {})


def _apply_table_style(doc, table, op, findings: list[Finding]) -> None:
    _apply_style(doc, table, op, findings, label="table ", expect_style=True)
    # Cluster D2: after the resolver's named table style is applied, re-emit the captured
    # table conditional-format facts (tblLook bitmask / style reference / cell margins),
    # set-only-when-unset, so the shell style's own banding/first-last emphasis renders.
    # No-op for a profile with no captured table appearance (byte-identical).
    _apply_table_appearance(table, op)
