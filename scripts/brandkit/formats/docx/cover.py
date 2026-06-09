# SPDX-License-Identifier: MIT
"""DOCX cover-anchor discovery and composition.

Real company covers seldom expose their title as a plain paragraph: the title
usually lives in a block-level ``w:sdt`` (a content control, with an ``alias`` /
``dataBinding`` / ``docPartGallery`` / placeholder prompt). python-docx's
``doc.paragraphs`` cannot see block-level SDTs, so discovery and composition both
work on the lxml tree here. The cover title is filled IN PLACE - only the inner
placeholder run's text is overwritten so the run-level formatting (``w:rPr``) the
brand defined survives - and a brand-new title paragraph is appended only when the
shell genuinely has no cover region (and then it is inserted before the first
toc/body child so it lands on the cover, never after the TOC).
"""

from __future__ import annotations

from typing import Optional

from brandkit.common import text as textutil
from brandkit.formats.docx.styles import lookup_style
from brandkit.formats.docx.structure import (
    _element_holds_strong_toc,
    _local_name,
    _p_style_val,
    classify_body_children,
    w,
)
from brandkit.ir.model import Cover
from brandkit.qa.model import Finding
from brandkit.profile import schema
from brandkit.profile.reconcile import confidence_clears_floor
from brandkit.profile.resolver import ProfileResolver


PLACEHOLDER_TITLE = "{{title}}"
# WEAK PRIOR ONLY (plan §5). These multilingual title prompts are a LAST-RESORT
# tiebreaker for naming the title slot when no structural SDT metadata is present;
# they are never the primary signal and never a matching rule that gates output.
# Cover-slot DISCOVERY is structural (block-level SDTs + the cover-region
# placeholder paragraphs); the model names each slot via ``comprehension``.
_TITLE_PROMPT_TOKENS: tuple[str, ...] = (
    "insert title",
    "title",
    "titolo",
    "titel",
    "titre",
    "titulo",
)


def _sdt_props(sdt):
    """Return the ``w:sdtPr`` element of a ``w:sdt``, or None."""
    return sdt.find(w("sdtPr"))


def _sdt_is_title(sdt) -> bool:
    """Heuristically decide whether a block-level ``w:sdt`` is the cover title.

    Evidence, in order of strength:
      - ``w:sdtPr/w:alias/@w:val`` or ``w:dataBinding`` xpath mentions a title
        token (``title``/``titolo``/``titel``/``titre``/``titulo``);
      - ``w:docPartGallery/@w:val`` mentions a title token;
      - the inner text is/contains a placeholder prompt (``{{title}}`` / "insert
        title" / a title token). Brand-agnostic and multilingual.

    A Table-of-Contents content control is **never** the cover title, no matter
    what title-like words its rendered entries contain. A TOC SDT
    (``docPartGallery='Table of Contents'`` or an inner ``w:instrText`` ``TOC``
    field) is excluded up front so the weak inner-text token match below can never
    misclassify the TOC as the title - which would dump the user's title into the
    TOC content control and blank every TOC entry. The author's own title slot
    (an ``alias``/``dataBinding`` SDT) still matches via the strong checks.
    """
    if _element_holds_strong_toc(sdt):
        return False
    props = _sdt_props(sdt)
    if props is not None:
        alias = props.find(w("alias"))
        if alias is not None:
            val = (alias.get(w("val")) or "").lower()
            if any(tok in val for tok in _TITLE_PROMPT_TOKENS):
                return True
        binding = props.find(w("dataBinding"))
        if binding is not None:
            xpath = (binding.get(w("xpath")) or "").lower()
            if any(tok in xpath for tok in _TITLE_PROMPT_TOKENS):
                return True
        for gallery in props.iter(w("docPartGallery")):
            val = (gallery.get(w("val")) or "").lower()
            if any(tok in val for tok in _TITLE_PROMPT_TOKENS):
                return True
    text = _sdt_text(sdt).strip().lower()
    if PLACEHOLDER_TITLE in text:
        return True
    if any(tok in text for tok in _TITLE_PROMPT_TOKENS):
        return True
    return False


def _sdt_text(sdt) -> str:
    return "".join(t.text for t in sdt.iter(w("t")) if t.text)


def _sdt_alias(sdt) -> Optional[str]:
    props = _sdt_props(sdt)
    if props is None:
        return None
    alias = props.find(w("alias"))
    return alias.get(w("val")) if alias is not None else None


def _sdt_databinding(sdt) -> Optional[str]:
    props = _sdt_props(sdt)
    if props is None:
        return None
    binding = props.find(w("dataBinding"))
    return binding.get(w("xpath")) if binding is not None else None


def _sdt_gallery(sdt) -> Optional[str]:
    props = _sdt_props(sdt)
    if props is None:
        return None
    for gallery in props.iter(w("docPartGallery")):
        return gallery.get(w("val"))
    return None


def _sdt_showing_placeholder(sdt) -> bool:
    props = _sdt_props(sdt)
    return props is not None and props.find(w("showingPlcHdr")) is not None


def _iter_block_sdts(doc):
    """Yield top-level block-level ``w:sdt`` elements in document order."""
    for child in doc.element.body:
        if _local_name(child.tag) == "sdt":
            yield child


def _cover_child_indices(doc) -> list[int]:
    """Return the top-level body-child indices that belong to the cover region."""
    return [
        c["index"] for c in classify_body_children(doc) if c.get("region") == "cover"
    ]


def _paragraph_is_placeholder_slot(p) -> bool:
    """Decide STRUCTURALLY whether a cover-region paragraph is a fillable slot.

    Evidence (language-invariant, no brand words):
      - it carries a literal ``{{...}}`` placeholder token, OR
      - it is a single short text run on the default (no explicit ``pStyle``, i.e.
        the run carries the cover's own direct formatting) - the shape a hand-typed
        cover prompt takes (a title line, an identifier line, a date line).

    A blank paragraph (empty text) is NOT a slot. The weak title-prompt tokens are
    consulted only to break ties when several short paragraphs qualify (see
    :func:`discover_cover`), never as the primary gate here.
    """
    text = _p_text_local(p).strip()
    if not text:
        return False
    if "{{" in text and "}}" in text:
        return True
    # A short, single-line cover prompt with no list/heading structure.
    if "\n" in text or len(text) > 120:
        return False
    return True


def _p_text_local(p) -> str:
    return "".join(t.text for t in p.iter(w("t")) if t.text)


def _fill_sdt_title(sdt, title: str) -> bool:
    """Overwrite the SDT's first text run with ``title`` IN PLACE.

    Writes into the first ``w:t`` inside the SDT content (``w:sdtContent``),
    preserving its run formatting, and clears any other ``w:t`` in the content so
    the placeholder prompt does not linger. Returns True on success.
    """
    content = sdt.find(w("sdtContent"))
    scope = content if content is not None else sdt
    texts = list(scope.iter(w("t")))
    if not texts:
        return False
    texts[0].text = title
    # xml:space=preserve so leading/trailing spaces in the title are not trimmed.
    texts[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    for extra in texts[1:]:
        extra.text = ""
    # Drop the "showingPlcHdr" flag so Word treats the content as real, not prompt.
    props = _sdt_props(sdt)
    if props is not None:
        for plc in props.findall(w("showingPlcHdr")):
            props.remove(plc)
    return True


def discover_cover(doc) -> tuple[list[dict], dict]:
    """Enumerate EVERY fillable cover slot as one anchor (plan §4 fact inventory).

    A real company cover is multi-slot: a title, a subtitle / object, a document
    id, a date, each its own fillable element. Discovery surfaces one anchor per
    slot - never just the title - so the comprehension can bind each slot and
    generation can FILL/CLEAR each in place. The two structural slot containers:

      - **block-level ``w:sdt``** (a content control) that is NOT a TOC/index field
        - keyed ``sdt.<body-child-index>``. The structural metadata (alias /
          dataBinding / docPartGallery / placeholder text / ``showingPlcHdr``) is
          captured as evidence; the model assigns the semantic role.
      - **cover-region placeholder paragraph** - a non-blank, short, single-line
        paragraph in the cover region (the body content before the first TOC /
        Heading-1) - keyed ``para.<body-child-index>``.

    The body-child index is a STABLE id: generation rewrites only the freeform body
    region, so the cover front matter keeps its positions, and the generator
    recomputes the same ids from the live tree. No anchor carries a brand word; the
    captured ``placeholder`` text is data (the template's own demo value), used for
    in-place fill and the residual-text check, never as a matching rule.

    Returns ``(anchors, anchor_block)`` where each anchor is::

        {"id": "sdt.8", "container": "sdt", "child_index": 8,
         "placeholder": "Descrizione", "alias": "Oggetto", "data_binding": None,
         "gallery": None, "showing_placeholder": False, "branches": None}

    and ``anchor_block`` is the legacy ``{"cover": {"kind","slots_found"}}`` summary.
    """
    anchors: list[dict] = []
    children = list(doc.element.body)
    cover_indices = set(_cover_child_indices(doc))

    for i, child in enumerate(children):
        ln = _local_name(child.tag)
        if ln == "sdt":
            # A TOC / index content control is never a cover slot.
            if _element_holds_strong_toc(child):
                continue
            anchors.append(
                {
                    "id": f"sdt.{i}",
                    "container": "sdt",
                    "child_index": i,
                    "placeholder": _sdt_text(child)[:200],
                    "alias": _sdt_alias(child),
                    "data_binding": _sdt_databinding(child),
                    "gallery": _sdt_gallery(child),
                    "showing_placeholder": _sdt_showing_placeholder(child),
                    "branches": None,
                }
            )
        elif ln == "p" and i in cover_indices and _paragraph_is_placeholder_slot(child):
            anchors.append(
                {
                    "id": f"para.{i}",
                    "container": "paragraph",
                    "child_index": i,
                    "placeholder": _p_text_local(child)[:200],
                    "style_id": _p_style_val(child),
                    "branches": None,
                }
            )

    anchor_block = {
        "cover": {
            "kind": schema.AnchorKind.SDT_ANCHORED.value
            if anchors
            else schema.AnchorKind.NONE.value,
            "slots_found": len(anchors),
        }
    }
    return anchors, anchor_block


def compose_cover(
    doc,
    cover: Cover | None,
    profile: dict,
    *,
    findings: Optional[list[Finding]] = None,
) -> set[str]:
    """Reconcile the PRESERVED cover slots with the new content (plan §6).

    When the profile carries a present comprehension, every cover slot is
    reconciled in place by its frozen ``fill_rule`` (multi-slot, never just the
    title): ``in_place`` FILLs the bound ``idoc.cover`` content preserving run
    formatting; ``clear`` empties / re-arms the slot, but ONLY when determinism
    corroborates the slot is a placeholder (``showingPlcHdr`` set, run==style
    default) and ``comprehension.confidence`` clears the floor - otherwise it is
    downgraded to KEEP + WARNING (the destructive-action floor); ``leave`` keeps
    the slot untouched. The append-a-new-title fallback fires ONLY when no
    title-bearing slot exists at all, so a duplicate title is never appended.

    When comprehension is ABSENT this falls back to today's deterministic single-
    title behavior (SDT title, then ``{{title}}``/"Insert title" placeholder, then
    append-before-TOC).

    Cluster E4: when extraction recorded the STRUCTURAL FACT that the shell has no
    cover anchor at all (``anchors.cover.kind == NONE`` - so there is no slot the
    in-place machinery could ever fill) a minimal cover is SYNTHESIZED from the
    profile's own RESOLVABLE ``cover.*`` roles through the shared resolver spine
    (:func:`_synthesize_cover_from_roles`). The path is completely DISJOINT from
    the anchored reconciliation above and falls through to the unchanged
    deterministic fill (byte-identical) whenever the fact is absent or no cover
    role resolves.

    Returns the set of cover anchor refs the reconciliation actually CLEARED
    (emptied/re-armed), so the caller can feed ``no_net_structure_loss``.
    """
    sink = findings if findings is not None else []
    if cover is None:
        return set()

    # ORDERING (intentional): _sync_core_properties runs FIRST, and its
    # _refresh_databound_sdt_caches fills + STRIPS w:showingPlcHdr on every cover SDT
    # bound to a core title/subject leaf (it has just written the brand text into the
    # binding, so the control is no longer a prompt). The comprehended path below
    # therefore reads POST-SYNC placeholder state: for such a slot _fill_anchor_in_place
    # sees was_placeholder=False and KEEPS the author's formatting (no role-style
    # reassert), and _clear_is_corroborated returns False so a model CLEAR downgrades to
    # KEEP+WARNING. Both are by design - the binding fill already wrote the brand text
    # preserving the run rPr, and clearing a freshly bound slot would be
    # self-contradictory. The placeholder gate reflecting post-sync state is the
    # contract, exercised end-to-end by CoverDataBoundSdtComposeOrderingTest.
    _sync_core_properties(doc, cover)

    comp = _present_comprehension(profile)
    if comp is not None and comp.get("cover_slots"):
        return _compose_cover_comprehended(doc, cover, profile, comp, sink)
    # E4 synthesis trigger: ONLY the recorded kind==NONE fact opens this branch
    # (a kind==NONE shell surfaces an empty cover_anchors inventory, so a present
    # comprehension can have bound no cover_slots and never reaches it). A False
    # return (no resolvable cover.* role) falls through byte-identically.
    if anchor_kind_is_none(profile) and _synthesize_cover_from_roles(
        doc, cover, profile, sink
    ):
        return set()
    return _compose_cover_deterministic(doc, cover, profile, sink)


def _sync_core_properties(doc, cover: Cover) -> None:
    """Keep Word document properties aligned with generated cover content.

    Corporate templates often render cover text through DOCPROPERTY fields or
    SDTs bound to ``docProps/core.xml``. If generation updates only the visible
    body text, LibreOffice/Word can refresh those fields back to stale values like
    "Titolo" / "Descrizione". Syncing the core properties makes field refreshes
    render the generated cover.
    """
    props = getattr(doc, "core_properties", None)
    title = textutil.runs_to_text(cover.title or []) or str(
        cover.fields.get("title", "")
    )
    subtitle = textutil.runs_to_text(cover.subtitle or []) or str(
        cover.fields.get("subtitle", "")
    )
    if props is not None:
        if title:
            props.title = title
        if subtitle:
            props.subject = subtitle
    # Syncing core.xml is not enough for headless renderers: a content control
    # bound to a core property renders its CACHED run text, not the re-resolved
    # binding, so a stale prompt would survive in the body and in running
    # headers/footers. Align every bound SDT's cache with the value we just wrote.
    _refresh_databound_sdt_caches(doc, title, subtitle)


def _iter_story_roots(doc):
    """Yield the body element plus every header/footer part root (``w:hdr``/``w:ftr``).

    Cover-bound content controls can live in running headers/footers (a company
    template often repeats the document subject in the page header), not only in
    the body. Returns each story's XML root so a single pass can reach every SDT.
    """
    yield doc.element.body
    try:
        parts = list(doc.part.package.iter_parts())
    except Exception:  # pragma: no cover - defensive; package always iterable
        return
    for part in parts:
        name = str(getattr(part, "partname", ""))
        if "header" in name or "footer" in name:
            el = getattr(part, "element", None)
            if el is not None:
                yield el


def _refresh_databound_sdt_caches(doc, title: str, subtitle: str) -> None:
    """Align the visible cache of every data-bound cover SDT with the synced value.

    A content control whose ``w:dataBinding`` xpath targets a ``docProps/core.xml``
    leaf displays the BOUND value. :func:`_sync_core_properties` updates that core
    property, but headless LibreOffice renders the SDT's cached run text instead of
    re-resolving the binding on export, so a stale prompt ("Titolo" / "Descrizione")
    would linger in the body AND in repeated headers/footers. Overwrite each bound
    SDT's cached text in place (dropping ``showingPlcHdr``) so the render matches
    the binding. Brand-agnostic: the SDT is matched on the binding xpath's
    core-property leaf (``title`` -> cover title, ``subject`` -> cover subtitle),
    never on the placeholder words a particular template happens to use.
    """
    bindings: dict[str, str] = {}
    if subtitle:
        bindings["subject"] = subtitle
    if title:
        bindings["title"] = title
    if not bindings:
        return
    for root in _iter_story_roots(doc):
        for sdt in root.iter(w("sdt")):
            leaf = _databinding_leaf(_sdt_databinding(sdt))
            value = bindings.get(leaf) if leaf else None
            if value is not None:
                _fill_sdt_title(sdt, value)


def _databinding_leaf(xpath: Optional[str]) -> Optional[str]:
    """Return the lowercased local element name of a data-binding xpath's last step.

    The binding xpath targets a core-property leaf, e.g.
    ``/ns1:coreProperties[1]/ns0:subject[1]`` -> ``subject``. Matching on the leaf
    name (not a raw substring) keeps ``subject``/``title`` disjoint - a substring
    test would wrongly fire ``title`` for a custom leaf named ``subtitle`` (which
    contains ``title``). Returns None for an empty/absent xpath.
    """
    if not xpath:
        return None
    step = xpath.rsplit("/", 1)[-1]  # last path step, e.g. "ns0:subject[1]"
    step = step.split("[", 1)[0]  # drop a positional predicate
    return step.rsplit(":", 1)[-1].strip().lower()  # drop the namespace prefix


def _present_comprehension(profile: dict) -> Optional[dict]:
    """Return the comprehension block only when it is present AND sha-current."""
    from brandkit.profile import store

    if not store.comprehension_is_present(profile):
        return None
    return profile.get("comprehension")


def _cover_content_for(cover: Cover, binds_to: Optional[str]) -> Optional[str]:
    """Resolve the new content string for a slot's ``binds_to`` key (advisory).

    ``binds_to`` is the model's advisory hint for WHICH content slot fills this
    anchor. ``title`` / ``subtitle`` map to the structured ``ir.Cover`` fields;
    anything else is looked up in ``cover.fields`` (the free-form ``{slot: value}``
    map). Returns None when the content has nothing for this slot (⇒ CLEAR if the
    model so ruled, else LEAVE).
    """
    if not binds_to:
        return None
    if binds_to == "title":
        return (
            textutil.runs_to_text(cover.title or [])
            or str(cover.fields.get("title", ""))
            or None
        )
    if binds_to == "subtitle":
        return (
            textutil.runs_to_text(cover.subtitle or [])
            or str(cover.fields.get("subtitle", ""))
            or None
        )
    val = cover.fields.get(binds_to)
    return str(val) if val not in (None, "") else None


def _compose_cover_comprehended(
    doc, cover: Cover, profile: dict, comp: dict, sink: list
) -> set[str]:
    """Multi-slot, comprehension-steered cover reconciliation."""
    confidence = float(comp.get("confidence") or 0.0)
    slots = comp.get("cover_slots") or {}
    cleared: set[str] = set()
    title_slot_filled = False
    has_title_slot = False

    # Resolve each slot's live element by its stable id, then act on fill_rule.
    for anchor_ref in sorted(slots):
        slot = slots[anchor_ref]
        if not isinstance(slot, dict):
            continue
        el = _resolve_anchor_element(doc, anchor_ref)
        if el is None:
            continue  # the slot is gone (already cleared / never existed); skip.
        fill_rule = slot.get("fill_rule")
        binds_to = slot.get("binds_to")
        if binds_to == "title":
            has_title_slot = True
        content = _cover_content_for(cover, binds_to)

        if fill_rule == schema.FillRule.IN_PLACE.value and content:
            _fill_anchor_in_place(doc, el, anchor_ref, content, profile, binds_to)
            if binds_to == "title":
                title_slot_filled = True
            continue

        if fill_rule == schema.FillRule.CLEAR.value or (
            fill_rule == schema.FillRule.IN_PLACE.value and not content
        ):
            # A CLEAR (or an in_place slot the content does not fill) is honored
            # only when determinism corroborates the slot is a placeholder AND the
            # confidence clears the floor; otherwise KEEP + WARNING.
            if _clear_is_corroborated(el, confidence):
                _clear_anchor(doc, el)
                cleared.add(anchor_ref)
            else:
                sink.append(
                    Finding(
                        "cover_clear_downgraded",
                        schema.Severity.WARNING.value,
                        f"cover slot {anchor_ref!r} clear not corroborated "
                        f"(confidence {confidence:.2f}); kept",
                    )
                )
            continue
        # fill_rule == leave (or unknown): leave the slot untouched.

    # Append a brand-new title ONLY when there is no title-bearing slot at all.
    title = textutil.runs_to_text(cover.title or []) or str(
        cover.fields.get("title", "")
    )
    if title and not has_title_slot and not title_slot_filled:
        para = doc.add_paragraph(title)
        _apply_role_style(doc, para, profile, "cover.title")
        _move_before_first_toc_or_body(doc, para)
        sink.append(
            Finding(
                "cover_degraded",
                schema.Severity.WARNING.value,
                "no title-bearing cover slot in shell; title paragraph appended "
                "before the first toc/body child",
            )
        )
    return cleared


def _unplaced_cover_extras(cover: Cover) -> list[str]:
    """Authored cover slots the deterministic (single-title) fill cannot place.

    The deterministic path fills only the title; a subtitle or extra ``fields`` the
    author supplied would otherwise vanish with no trace. Returns the slot names so
    the caller can surface them (the comprehension path places them in full)."""
    extras: list[str] = []
    subtitle = textutil.runs_to_text(cover.subtitle or []) or str(
        cover.fields.get("subtitle", "")
    )
    if subtitle:
        extras.append("subtitle")
    for key, val in (cover.fields or {}).items():
        if key not in ("title", "subtitle") and val:
            extras.append(key)
    return extras


def _note_unplaced_cover_extras(sink: list, extras: list[str]) -> None:
    """INFO (not silent) when the deterministic fill leaves authored slots unplaced."""
    if extras:
        sink.append(
            Finding(
                "cover_degraded",
                schema.Severity.INFO.value,
                "deterministic cover fill placed only the title; authored cover "
                f"slot(s) not placed: {', '.join(extras)} "
                "(comprehension fills the multi-slot cover in place)",
            )
        )


def _compose_cover_deterministic(
    doc, cover: Cover, profile: dict, sink: list
) -> set[str]:
    """Deterministic cover fill (comprehension absent): place the title, then a
    style-identified subtitle slot. Any remaining authored slots (extra fields) are
    surfaced as unplaced - the comprehension path fills the full multi-slot cover.
    """
    title = textutil.runs_to_text(cover.title or []) or str(
        cover.fields.get("title", "")
    )
    if not title:
        return set()
    extras = _unplaced_cover_extras(cover)
    subtitle = textutil.runs_to_text(cover.subtitle or []) or str(
        cover.fields.get("subtitle", "")
    )

    _fill_cover_title_deterministic(doc, profile, title, sink)
    if (
        subtitle
        and _fill_cover_subtitle_deterministic(doc, profile, subtitle)
        and "subtitle" in extras
    ):
        extras.remove("subtitle")
    _note_unplaced_cover_extras(sink, extras)
    return set()


def anchor_kind_is_none(profile: dict) -> bool:
    """True iff extraction RECORDED that the shell carries no cover anchor (E4).

    The trigger is the STRUCTURAL FACT ``anchors.cover.kind == AnchorKind.NONE``
    stamped by :func:`discover_cover` at extract time - never a guess recomputed at
    generate time. A profile with no ``anchors.cover`` block at all (hand-built
    profiles, pre-anchor envelopes) is NOT kind==NONE: the fact was never recorded,
    so the synthesis path stays cold and output is byte-identical to today.
    """
    cover_anchor = (profile.get("anchors") or {}).get("cover") or {}
    return cover_anchor.get("kind") == schema.AnchorKind.NONE.value


def _synthesize_cover_from_roles(doc, cover: Cover, profile: dict, sink: list) -> bool:
    """Cluster E4: build a MINIMAL cover from the profile's RESOLVABLE cover.* roles.

    Fires only from :func:`compose_cover` when ``anchors.cover.kind == NONE`` (the
    shell exposes no SDT / placeholder-paragraph cover slot the in-place machinery
    could fill). Each authored slot - title first, then subtitle, the canonical
    order - is resolved through the SHARED resolver spine
    (:meth:`ProfileResolver.resolve_role`, the single brand chokepoint): a slot is
    synthesized ONLY when its role resolves to a style the live shell actually
    carries (:func:`lookup_style`); a stub/unresolvable role contributes NOTHING
    (never a fabricated style/font/hex). The synthesized paragraphs are ordinary
    role-resolved paragraphs moved before the first toc/body child - the same
    position the deterministic append fallback uses - and are NEVER touched by the
    anchored-cover machinery: ``_sync_core_properties`` already ran (it rewrites
    only existing data-bound SDTs) and no SDT/``showingPlcHdr`` state exists on a
    fresh paragraph, so the D4 clobber class is structurally unreachable here.

    Returns True when at least one slot was synthesized; the INFO
    ``cover_synthesized`` finding then audits the structural fact + the role ids
    used (mirroring ``override_applied`` - never brand text). Returns False to
    fall through to the UNCHANGED deterministic path, so a kind==NONE shell whose
    cover.* roles do not resolve stays byte-identical to today. An authored title
    is never dropped: when title content exists but ``cover.title`` does not
    resolve, synthesis declines entirely (the deterministic fallback still places
    the title).
    """
    title = textutil.runs_to_text(cover.title or []) or str(
        cover.fields.get("title", "")
    )
    subtitle = textutil.runs_to_text(cover.subtitle or []) or str(
        cover.fields.get("subtitle", "")
    )
    resolver = ProfileResolver(profile)
    slots: list[tuple[str, str, object]] = []
    for role_id, content in (("cover.title", title), ("cover.subtitle", subtitle)):
        if not content:
            continue
        op = resolver.resolve_role(role_id, fallback=None)
        style = lookup_style(doc, op.resolver) if op.resolver else None
        if style is not None:
            slots.append((role_id, content, style))
    if not slots:
        return False  # nothing resolves: byte-identical fall-through
    if title and slots[0][0] != "cover.title":
        return False  # never place a subtitle while the authored title would drop
    paras = []
    for _role_id, content, style in slots:
        para = doc.add_paragraph(content)
        para.style = style
        paras.append(para)
    _move_paras_before_first_toc_or_body(doc, paras)
    placed = [role_id for role_id, _content, _style in slots]
    placed_slot_names = {role_id.split(".", 1)[1] for role_id in placed}
    extras = [e for e in _unplaced_cover_extras(cover) if e not in placed_slot_names]
    if extras:
        sink.append(
            Finding(
                "cover_degraded",
                schema.Severity.INFO.value,
                "synthesized cover placed only the resolvable cover role(s); "
                f"authored cover slot(s) not placed: {', '.join(extras)}",
            )
        )
    sink.append(
        Finding(
            "cover_synthesized",
            schema.Severity.INFO.value,
            "no cover anchor in shell (anchors.cover.kind=NONE); cover "
            f"synthesized from resolved role(s): {', '.join(placed)}",
        )
    )
    return True


def _move_paras_before_first_toc_or_body(doc, paras: list) -> None:
    """Move freshly appended paragraphs - keeping their order - before the first
    toc/body child.

    The group form of :func:`_move_before_first_toc_or_body` (which is kept
    untouched for the existing single-title paths): the whole synthesized group is
    excluded when picking the move target, so the title/subtitle land on the cover
    in authored order even if an already-moved sibling were itself classified as a
    body child. No toc/body child at all leaves the paragraphs appended (same
    last-resort as the single-paragraph helper).
    """
    if not paras:
        return
    body = doc.element.body
    own = {p._p for p in paras}
    children = list(body)
    target = None
    for c in classify_body_children(doc):
        if c["index"] >= len(children):
            continue
        el = children[c["index"]]
        if el in own:
            continue
        if c["region"] in ("toc", "body"):
            target = el
            break
    if target is None:
        return
    for p in paras:
        body.remove(p._p)
        target.addprevious(p._p)


def _fill_cover_title_deterministic(doc, profile: dict, title: str, sink: list) -> None:
    """Place the cover title IN PLACE: a title SDT, then a ``{{title}}``/"Insert
    title" placeholder paragraph, then (last resort) a new title paragraph moved
    before the first toc/body child. Unchanged behavior, extracted so the subtitle
    fill can run after whichever title branch fired."""
    # 1) Block-level SDT cover title.
    for sdt in _iter_block_sdts(doc):
        if _sdt_is_title(sdt) and _fill_sdt_title(sdt, title):
            _apply_role_style_sdt(doc, sdt, profile, "cover.title")
            return

    # 2) Placeholder paragraph - overwrite only the matching run's text in place.
    for para in doc.paragraphs[:8]:
        if PLACEHOLDER_TITLE in para.text or "Insert title" in para.text:
            _fill_paragraph_in_place(para, title)
            _apply_role_style(doc, para, profile, "cover.title")
            return

    # 3) No cover anchor: append a title paragraph but place it on the cover, i.e.
    # BEFORE the first toc/body child, never after the TOC.
    para = doc.add_paragraph(title)
    _apply_role_style(doc, para, profile, "cover.title")
    _move_before_first_toc_or_body(doc, para)
    sink.append(
        Finding(
            "cover_degraded",
            schema.Severity.WARNING.value,
            "no cover anchor in shell; title paragraph appended before the first toc/body child",
        )
    )


def _fill_cover_subtitle_deterministic(doc, profile: dict, subtitle: str) -> bool:
    """Fill a cover-region subtitle slot IN PLACE, identified by STYLE identity.

    Correct-by-construction: the slot is the first cover-region paragraph whose
    paragraph style is the profile's resolved ``cover.subtitle`` style - never a
    guess from the template's placeholder text. This handles the common plain
    styled-subtitle line (e.g. a "Cover Subtitle" paragraph) that the single-title
    fill used to leave showing the template's stale demo subtitle. A template whose
    subtitle is a databound SDT (bound to e.g. ``core/subject``) is already covered
    by :func:`_sync_core_properties`; extra fields (date/id/author) stay the
    comprehension path's job. Returns True when a slot was filled."""
    resolver = ((profile.get("roles") or {}).get("cover.subtitle") or {}).get(
        "resolver"
    ) or {}
    style_id = resolver.get("style_id")
    if not style_id:
        return False
    children = list(doc.element.body)
    for i in sorted(_cover_child_indices(doc)):
        if i >= len(children):
            continue
        child = children[i]
        if _local_name(child.tag) != "p" or _p_style_val(child) != style_id:
            continue
        _fill_p_element_in_place(child, subtitle)
        _apply_role_style_p_element(doc, child, profile, "cover.subtitle")
        return True
    return False


def _resolve_anchor_element(doc, anchor_ref: str):
    """Resolve a stable anchor id (``sdt.<i>`` / ``para.<i>``) to its live element.

    The body-child index in the id is recomputed against the live tree; cover
    front matter keeps its positions across generation (only the freeform body is
    rewritten), so the same id resolves to the same element at generate time.
    Returns the element, or None when the id no longer maps (defensive).
    """
    kind, _, idx_s = anchor_ref.partition(".")
    try:
        idx = int(idx_s)
    except ValueError:
        return None
    children = list(doc.element.body)
    if idx < 0 or idx >= len(children):
        return None
    el = children[idx]
    ln = _local_name(el.tag)
    if kind == "sdt" and ln == "sdt":
        return el
    if kind == "para" and ln == "p":
        return el
    return None


def _fill_anchor_in_place(
    doc,
    el,
    anchor_ref: str,
    content: str,
    profile: dict,
    binds_to: Optional[str] = None,
) -> None:
    """FILL a cover anchor element in place, preserving run formatting.

    After the in-place fill, the slot's bound role style is re-asserted so a filled
    cover slot is brand-styled rather than carrying whatever incidental style a bare
    prompt had (D4). The role is derived from the slot's ``binds_to``
    (``title``->``cover.title``, ``subtitle``->``cover.subtitle``), a verbatim
    resolver target only - no literal.

    The re-assertion is gated for an SDT: it fires only when the control was a bare
    PLACEHOLDER before the fill (``showingPlcHdr``). A real, author-formatted slot
    (not showing a placeholder) already carries the template's intended cover
    styling, and stamping the role's auto-mapped paragraph style over it can replace
    working direct formatting with a style that renders differently - e.g. a builtin
    ``Subtitle`` whose color is a near-white ``text1`` tint, which blanks a filled
    slot on a white page. For such a slot we fill text only and keep the author's
    formatting (``_fill_sdt_title`` already preserves the run ``rPr``).
    """
    ln = _local_name(el.tag)
    role_id = _cover_role_for(binds_to)
    if ln == "sdt":
        was_placeholder = _sdt_showing_placeholder(
            el
        )  # capture BEFORE the fill drops it
        _fill_sdt_title(el, content)
        if role_id and was_placeholder:
            _apply_role_style_sdt(doc, el, profile, role_id)
    elif ln == "p":
        _fill_p_element_in_place(el, content)
        if role_id:
            _apply_role_style_p_element(doc, el, profile, role_id)


def _cover_role_for(binds_to: Optional[str]) -> Optional[str]:
    """Map a slot's ``binds_to`` hint to its cover role id (verbatim, no literal).

    ``title`` -> ``cover.title``; ``subtitle`` -> ``cover.subtitle``. Any other (or
    absent) binding defaults to ``cover.title`` so a filled, un-annotated slot
    still re-asserts the cover title style rather than keeping a prompt's incidental
    style. Returns the role id, or None only when there is nothing to bind.
    """
    if binds_to == "subtitle":
        return "cover.subtitle"
    return "cover.title"


def _clear_is_corroborated(el, confidence: float) -> bool:
    """Destructive-action floor for a cover CLEAR (plan §6).

    A CLEAR is corroborated only when determinism agrees the slot is a placeholder
    - an SDT still ``showingPlcHdr``, or a paragraph whose runs carry no explicit
    paragraph style (the cover's own default-styled prompt) - AND the model's
    confidence clears the floor. Both conditions are required; otherwise the slot
    is kept.
    """
    if not confidence_clears_floor(confidence):
        return False
    ln = _local_name(el.tag)
    if ln == "sdt":
        return _sdt_showing_placeholder(el)
    if ln == "p":
        # A cover prompt paragraph carries no heading/list style of its own.
        return _p_style_val(el) is None
    return False


def _clear_anchor(doc, el) -> None:
    """CLEAR a cover anchor in place: empty its text, re-arming an SDT placeholder."""
    ln = _local_name(el.tag)
    if ln == "sdt":
        content = el.find(w("sdtContent"))
        scope = content if content is not None else el
        for t in scope.iter(w("t")):
            t.text = ""
    elif ln == "p":
        for t in el.iter(w("t")):
            t.text = ""


def _fill_p_element_in_place(p_el, content: str) -> None:
    """Set ``content`` on a lxml ``w:p`` placeholder, preserving the first run rPr."""
    texts = list(p_el.iter(w("t")))
    if not texts:
        # No run/text node: create a minimal run so the content is not lost.
        from lxml import etree

        r = etree.SubElement(p_el, w("r"))
        t = etree.SubElement(r, w("t"))
        t.text = content
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        return
    texts[0].text = content
    texts[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    for extra in texts[1:]:
        extra.text = ""


def _fill_paragraph_in_place(para, title: str) -> None:
    """Set the title on a placeholder paragraph without destroying run rPr.

    Writes ``title`` into the first run and clears the remaining runs, so the
    run-level formatting the brand defined on the first run is preserved (unlike
    ``para.text = title``, which rebuilds the runs from scratch).
    """
    runs = para.runs
    if not runs:
        para.add_run(title)
        return
    runs[0].text = title
    for r in runs[1:]:
        r.text = ""


def _move_before_first_toc_or_body(doc, para) -> None:
    """Move ``para``'s ``w:p`` element so it precedes the first toc/body child."""
    body = doc.element.body
    p_el = para._p
    classes = classify_body_children(doc)
    children = list(body)
    target = None
    for c in classes:
        if c["index"] >= len(children):
            continue
        el = children[c["index"]]
        if el is p_el:
            continue
        if c["region"] in ("toc", "body"):
            target = el
            break
    if target is not None:
        body.remove(p_el)
        target.addprevious(p_el)


def _apply_role_style(doc, para, profile: dict, role_id: str) -> None:
    entry = (profile.get("roles") or {}).get(role_id) or {}
    resolver = entry.get("resolver") or {}
    style = lookup_style(doc, resolver)
    if style is not None:
        para.style = style


def _apply_role_style_sdt(doc, sdt, profile: dict, role_id: str) -> None:
    """Apply a role's paragraph style to the first paragraph inside the SDT."""
    content = sdt.find(w("sdtContent"))
    scope = content if content is not None else sdt
    p = next(iter(scope.iter(w("p"))), None)
    if p is not None:
        _apply_role_style_p_element(doc, p, profile, role_id)


def _apply_role_style_p_element(doc, p_el, profile: dict, role_id: str) -> None:
    """Stamp a role's ``w:pStyle`` onto a bare lxml ``w:p`` element.

    The single place cover style re-assertion writes a paragraph style id onto an
    lxml element (used by both the SDT branch and the plain-paragraph branch, D4).
    No-op when the role resolves to no shell style or the style has no id, so a
    missing/absent role never crashes the in-place fill.
    """
    if _local_name(p_el.tag) != "p":
        return
    entry = (profile.get("roles") or {}).get(role_id) or {}
    resolver = entry.get("resolver") or {}
    style = lookup_style(doc, resolver)
    if style is None:
        return
    style_id = getattr(style, "style_id", None)
    if not style_id:
        return
    pPr = p_el.find(w("pPr"))
    if pPr is None:
        from lxml import etree

        pPr = etree.SubElement(p_el, w("pPr"))
        p_el.insert(0, pPr)
    pStyle = pPr.find(w("pStyle"))
    if pStyle is None:
        from lxml import etree

        pStyle = etree.SubElement(pPr, w("pStyle"))
        pPr.insert(0, pStyle)
    pStyle.set(w("val"), style_id)
