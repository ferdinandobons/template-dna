# SPDX-License-Identifier: MIT
"""PPTX structure helpers - the format-uniform inventory peer of docx structure.

The deck's *ordered structure* (cover -> section list / agenda -> body slides) is
first-class, exactly as the docx cover/TOC/body trio is. Generation must preserve
the cover and the structural slides and only regenerate the freeform body region
(and re-derive the agenda/section-list from the new headings). Everything here
works on the parsed ``python-pptx`` presentation plus, where python-pptx does not
expose it, the raw lxml presentation element (the PowerPoint *section list* lives
in a ``p14:sectionLst`` extension that python-pptx 1.x does not surface, just as
block-level ``w:sdt`` is invisible to python-docx).

Detection is grounded in evidence, never in brand-specific names:

- **cover_anchors** - every placeholder on a cover/section-layout, keyed by its
  layout name + placeholder ``idx`` (``ph.<layout-idx>.<ph-idx>``). The placeholder
  TYPE (TITLE/SUBTITLE/BODY/DATE/FOOTER/PICTURE/...) is the language-invariant
  structural signal; the layout placeholder's own prompt text is captured as the
  demo value (used for in-place fill and the residual-text check), never as a
  matching rule. This is the multi-placeholder cover analogue of the docx
  multi-slot cover.
- **fields** (derived indexes) - the deck's real *section list* (a
  ``p14:sectionLst``) surfaces as an agenda/section-list index that ``generate``
  regenerates from the new headings, the PPTX peer of the docx outline TOC. A deck
  with no section list surfaces no field (honest), just as a docx with no TOC field
  surfaces none.
- **regions** - every slide classified ``cover`` / ``structural`` / ``demo``. A
  *demo* slide is one whose body text **equals a layout placeholder prompt** (the
  author-facing "Click to edit..." sample the template ships with) - real,
  structural slides keep their own authored text. The cover region is the cover
  slide(s); section-list/agenda slides are structural; the rest is body.

The ids are stable and recomputable from the live deck at generate time, so the
model binds a ref once and the generator resolves the same element later.
"""

from __future__ import annotations

from typing import Optional

from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER

from brandkit.ooxml.names import local_name as _local_name

# ---------------------------------------------------------------------------
# Namespaces. The presentation section list is a Microsoft 2010 extension that
# python-pptx does not model; we read it off the raw lxml tree (peer of how docx
# structure walks ``w:sdt`` / ``w:instrText`` the high-level API hides). The
# ``_local_name`` helper comes from the shared :mod:`brandkit.ooxml.names` layer.
# ---------------------------------------------------------------------------
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"
# The frozen URI of the section-list extension (ECMA/MS-PPTX). This is a FORMAT
# constant, never a brand word.
SECTION_EXT_URI = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"


# ---------------------------------------------------------------------------
# Placeholder-type families (language-invariant structural signal). A *title*
# slot is any TITLE/CENTER_TITLE/VERTICAL_TITLE; a *subtitle* slot is SUBTITLE; a
# *body* slot is any text-bearing content placeholder. These are the PRIMARY
# signal the model reasons over - the same role a placeholder plays whatever the
# template's prompt text reads.
# ---------------------------------------------------------------------------
TITLE_TYPES = frozenset(
    {
        PP_PLACEHOLDER.TITLE,
        PP_PLACEHOLDER.CENTER_TITLE,
        PP_PLACEHOLDER.VERTICAL_TITLE,
    }
)
SUBTITLE_TYPES = frozenset({PP_PLACEHOLDER.SUBTITLE})
BODY_TYPES = frozenset(
    {
        PP_PLACEHOLDER.BODY,
        PP_PLACEHOLDER.OBJECT,
        PP_PLACEHOLDER.VERTICAL_BODY,
        PP_PLACEHOLDER.VERTICAL_OBJECT,
    }
)


def _ph_family(ptype) -> str:
    """Map a placeholder TYPE to an OPEN advisory family token.

    The token (``title`` / ``subtitle`` / ``body`` / ``date`` / ``footer`` /
    ``picture`` / ``other``) is carried into the inventory as evidence; the model
    assigns the actual semantic role. It is derived from the structural TYPE, never
    from the prompt words, so it is language-invariant.
    """
    if ptype in TITLE_TYPES:
        return "title"
    if ptype in SUBTITLE_TYPES:
        return "subtitle"
    if ptype in BODY_TYPES:
        return "body"
    name = str(ptype or "").rsplit(".", 1)[-1].split(" ")[0].lower()
    return name or "other"


# ---------------------------------------------------------------------------
# Layout classification (the only source of truth for cover/content selection)
# ---------------------------------------------------------------------------
def classify_layouts(prs) -> list[dict]:
    """Describe each REAL slide layout by the slots it actually exposes.

    Returns a list (in deck order) of descriptors::

        {"name", "idx", "title_idx", "subtitle_idx", "body_idx",
         "placeholders": [{"idx", "ptype", "family", "prompt"}, ...]}

    where each ``*_idx`` is the placeholder ``idx`` of the first slot of that
    family present in the layout, or ``None``. ``placeholders`` lists every
    placeholder with its captured layout *prompt* text (the demo value). Nothing
    here is fabricated: every entry points at a placeholder python-pptx parsed.
    """
    described: list[dict] = []
    for pos, layout in enumerate(prs.slide_layouts):
        title_idx = subtitle_idx = body_idx = None
        phs: list[dict] = []
        for ph in layout.placeholders:
            fmt = ph.placeholder_format
            ptype = fmt.type
            family = _ph_family(ptype)
            if ptype in TITLE_TYPES and title_idx is None:
                title_idx = fmt.idx
            elif ptype in SUBTITLE_TYPES and subtitle_idx is None:
                subtitle_idx = fmt.idx
            elif ptype in BODY_TYPES and body_idx is None:
                body_idx = fmt.idx
            prompt = (
                ph.text[:200]
                if getattr(ph, "has_text_frame", False) and ph.text
                else ""
            )
            phs.append(
                {
                    "idx": fmt.idx,
                    "ptype": str(ptype),
                    "family": family,
                    "prompt": prompt,
                }
            )
        described.append(
            {
                "name": layout.name,
                "idx": pos,
                "title_idx": title_idx,
                "subtitle_idx": subtitle_idx,
                "body_idx": body_idx,
                "placeholders": phs,
            }
        )
    return described


def pick_cover(described: list[dict]) -> Optional[dict]:
    """Pick the layout that best reads as a cover/title slide.

    Strongest signal: a title slot paired with a subtitle slot (the canonical
    cover shape). Falls back to the first layout exposing any title slot.
    """
    for d in described:
        if d["title_idx"] is not None and d["subtitle_idx"] is not None:
            return d
    for d in described:
        if d["title_idx"] is not None:
            return d
    return None


def pick_content(
    described: list[dict], *, exclude_idx: Optional[int] = None
) -> Optional[dict]:
    """Pick the layout that best reads as a title+body content slide.

    Prefers a layout with BOTH a title and a body slot, skipping ``exclude_idx``
    (the cover) when an alternative exists. Falls back to any title-bearing layout,
    then any body-bearing layout.
    """
    title_body = [
        d for d in described if d["title_idx"] is not None and d["body_idx"] is not None
    ]
    for d in title_body:
        if d["idx"] != exclude_idx:
            return d
    if title_body:
        return title_body[0]
    for d in described:
        if d["title_idx"] is not None:
            return d
    for d in described:
        if d["body_idx"] is not None:
            return d
    return None


# ---------------------------------------------------------------------------
# Cover-anchor inventory (the multi-placeholder cover; plan §4 fact inventory)
# ---------------------------------------------------------------------------
def inventory_cover_anchors(prs, described: list[dict] | None = None) -> list[dict]:
    """Surface EVERY cover-layout placeholder as one anchor (multi-placeholder cover).

    A real cover is multi-slot: a title, a subtitle, one or more body/overlay
    slots, a date, a footer - each its own fillable placeholder. Discovery surfaces
    one anchor per placeholder of the chosen cover layout (never just the title),
    keyed ``ph.<layout-idx>.<ph-idx>`` - a STABLE id the generator recomputes from
    the same layout at generate time. The placeholder ``family`` (from its TYPE)
    and the layout's own prompt text are captured as evidence / demo value; the
    model assigns the semantic role and the fill rule.

    Returns ``[]`` when no layout reads as a cover (honest - nothing to bind).
    Each anchor::

        {"id": "ph.0.0", "container": "placeholder", "layout": "Cover",
         "layout_idx": 0, "ph_idx": 0, "ph_type": "TITLE (1)", "family": "title",
         "placeholder": "Click to edit title"}

    ``described`` optionally injects a pre-computed :func:`classify_layouts`
    result (a pure function of the unmutated deck) so one extract pass can share
    a single classification; ``None`` (the default) recomputes as before.
    """
    if described is None:
        described = classify_layouts(prs)
    cover = pick_cover(described)
    if cover is None:
        return []
    anchors: list[dict] = []
    for ph in cover["placeholders"]:
        anchors.append(
            {
                "id": f"ph.{cover['idx']}.{ph['idx']}",
                "container": "placeholder",
                "layout": cover["name"],
                "layout_idx": cover["idx"],
                "ph_idx": ph["idx"],
                "ph_type": ph["ptype"],
                "family": ph["family"],
                "placeholder": ph["prompt"],
            }
        )
    return anchors


# ---------------------------------------------------------------------------
# Section list (the deck's real agenda/section index; plan §4 fields inventory)
# ---------------------------------------------------------------------------
def _section_list_element(prs):
    """Return the ``p14:sectionLst`` lxml element, or None when the deck has none."""
    pres = prs.part._element
    for ext in pres.iter():
        if _local_name(ext.tag) != "ext":
            continue
        if (ext.get("uri") or "") != SECTION_EXT_URI:
            continue
        for child in ext.iter():
            if _local_name(child.tag) == "sectionLst":
                return child
    return None


def detect_sections(prs) -> list[dict]:
    """Return the deck's REAL section list (empty when the deck has none).

    Each entry is ``{"name": <section name>, "slide_ids": [<sldId>, ...]}`` in
    deck order, read off the ``p14:sectionLst`` extension. This is the structural
    backing for the agenda/section-list derived index - a deck with no sections
    has no agenda field to regenerate (honest), exactly as a docx with no TOC field
    surfaces no index. The section NAMES are data carried into the profile, never a
    matching rule.
    """
    sect_lst = _section_list_element(prs)
    if sect_lst is None:
        return []
    out: list[dict] = []
    for sec in sect_lst:
        if _local_name(sec.tag) != "section":
            continue
        slide_ids = [
            c.get("id")
            for c in sec.iter()
            if _local_name(c.tag) == "sldId" and c.get("id")
        ]
        out.append({"name": sec.get("name") or "", "slide_ids": slide_ids})
    return out


def inventory_fields(prs) -> list[dict]:
    """Surface the deck's derived indexes as a stable-id inventory (plan §4).

    Today the single PPTX derived index is the *agenda / section list*: a deck that
    carries a ``p14:sectionLst`` has a list it regenerates from its own section
    headings, the PPTX peer of the docx outline TOC. It surfaces as one field id
    ``field.sections`` with ``seq_id=None`` (an outline-style index, not a ``\\c``
    caption index). A deck with no section list surfaces no field, so a comprehension
    ref into the (then empty) fields inventory is fail-closed at QA time.

    Each entry: ``{"id": "field.sections", "seq_id": None, "kind": "section_list",
    "section_count": <n>}``.
    """
    sections = detect_sections(prs)
    if not sections:
        return []
    return [
        {
            "id": "field.sections",
            "seq_id": None,
            "kind": "section_list",
            "section_count": len(sections),
        }
    ]


# ---------------------------------------------------------------------------
# Demo-slide detection + region inventory (plan §4 region/demo candidates)
# ---------------------------------------------------------------------------
def _layout_prompts(prs) -> set[str]:
    """Return every non-empty layout placeholder prompt text in the deck.

    These are the author-facing sample strings the template ships with ("Click to
    edit title", "Click to enter/edit subtitle text", ...). A slide whose body text
    EQUALS one of these is showing an unedited template prompt - a demo slide. The
    set is the template's OWN captured text, never a fixed phrase, so the detection
    is language-invariant.
    """
    prompts: set[str] = set()
    for layout in prs.slide_layouts:
        for ph in layout.placeholders:
            if getattr(ph, "has_text_frame", False) and ph.text:
                t = ph.text.strip()
                if t:
                    prompts.add(t)
    for master in prs.slide_masters:
        for ph in master.placeholders:
            if getattr(ph, "has_text_frame", False) and ph.text:
                t = ph.text.strip()
                if t:
                    prompts.add(t)
    return prompts


def _slide_texts(slide) -> list[str]:
    """Return the stripped, non-empty text of each shape on a slide."""
    out: list[str] = []
    for shape in slide.shapes:
        if getattr(shape, "has_text_frame", False) and shape.text:
            t = shape.text.strip()
            if t:
                out.append(t)
    return out


def _slide_is_demo(slide, prompts: set[str]) -> bool:
    """A slide is demo when it has text AND all of its text equals layout prompts.

    Evidence-based, language-invariant: a slide whose every text run is an unedited
    template placeholder prompt is sample/demo content the author is meant to
    replace. A slide carrying any authored text (text that is NOT a prompt) is real
    structural content and is kept. An empty slide is not demo (nothing to clear).
    """
    texts = _slide_texts(slide)
    if not texts:
        return False
    return all(t in prompts for t in texts)


def _cover_layout_names(prs, described: list[dict] | None = None) -> set[str]:
    """Return the set of layout names that read as a cover (title-bearing cover)."""
    if described is None:
        described = classify_layouts(prs)
    cover = pick_cover(described)
    return {cover["name"]} if cover is not None else set()


def classify_slides(prs, described: list[dict] | None = None) -> list[dict]:
    """Classify each slide into a region (``cover`` / ``structural`` / ``demo``).

    Returns one descriptor per slide (in deck order)::

        {"index", "slide_id", "layout", "region"}

    Boundaries (evidence-based, never positional):
      - a slide built on the cover layout is ``cover``;
      - a slide whose body text is entirely layout placeholder prompts is ``demo``
        (unedited sample content to clear);
      - everything else is ``structural`` (real authored slides to KEEP).

    A demo classification wins over structural but never over cover (a cover slide
    showing its own prompts is still the cover, filled in place, not cleared).

    ``described`` optionally injects a pre-computed :func:`classify_layouts`
    result (a pure function of the unmutated deck) so one extract pass can share
    a single classification; ``None`` (the default) recomputes as before.
    """
    cover_layouts = _cover_layout_names(prs, described)
    prompts = _layout_prompts(prs)
    out: list[dict] = []
    sld_ids = [sid.get("id") for sid in prs.slides._sldIdLst]
    for i, slide in enumerate(prs.slides):
        layout_name = slide.slide_layout.name
        if layout_name in cover_layouts:
            region = "cover"
        elif _slide_is_demo(slide, prompts):
            region = "demo"
        else:
            region = "structural"
        out.append(
            {
                "index": i,
                "slide_id": sld_ids[i] if i < len(sld_ids) else None,
                "layout": layout_name,
                "region": region,
            }
        )
    return out


def inventory_regions(prs, classes: list[dict] | None = None) -> list[dict]:
    """Surface the deck's region inventory (stable ids the model binds to).

    Combines a per-slide region id (``region.slide.<i>`` carrying its
    classification) with the conceptual top-level regions actually present
    (``region.cover`` / ``region.body`` / ``region.sections`` when there is a
    section list). Each entry is ``{"id": <region_ref>, "kind": <open token>}``.

    Deterministic and recomputable at generate time; the ids encode structural
    positions / classifications, never brand words. ``classes`` optionally
    injects a pre-computed :func:`classify_slides` result (a pure function of
    the unmutated deck) so one extract pass can share a single classification;
    ``None`` (the default) recomputes as before.
    """
    out: list[dict] = []
    if classes is None:
        classes = classify_slides(prs)
    seen_top: set[str] = set()
    for c in classes:
        out.append({"id": f"region.slide.{c['index']}", "kind": c["region"]})
        top = "cover" if c["region"] == "cover" else "body"
        if top not in seen_top:
            seen_top.add(top)
            out.append({"id": f"region.{top}", "kind": top})
    if detect_sections(prs):
        out.append({"id": "region.sections", "kind": "section_list"})
    return out


# ---------------------------------------------------------------------------
# Ordered skeleton (the profile["structure"] payload; peer of docx)
# ---------------------------------------------------------------------------
def detect_skeleton(
    prs,
    described: list[dict] | None = None,
    classes: list[dict] | None = None,
) -> dict:
    """Detect the deck's ordered top-level skeleton (peer of docx detect_skeleton).

    Returns the ``structure`` section for ``profile.json``::

        {"ordered": True,
         "skeleton": [ {"region","order","role","required","repeatable",...}, ...]}

    Only regions actually present are included. The cover region exists when a
    cover layout exists; the sections/agenda region exists when the deck carries a
    section list; the body region is ``freeform`` (slide order inside it is not
    prescribed). ``ordered`` means the top-level region order must be respected.

    ``described`` / ``classes`` optionally inject pre-computed
    :func:`classify_layouts` / :func:`classify_slides` results (pure functions
    of the unmutated deck) so one extract pass can share a single
    classification; ``None`` (the default) recomputes as before.
    """
    if described is None:
        described = classify_layouts(prs)
    has_cover = pick_cover(described) is not None
    has_sections = bool(detect_sections(prs))
    if classes is None:
        classes = classify_slides(prs, described)
    has_body = any(c["region"] in ("structural", "demo") for c in classes)

    skeleton: list[dict] = []
    order = 0
    if has_cover:
        skeleton.append(
            {
                "region": "cover",
                "order": order,
                "role": "section.cover",
                "required": True,
                "repeatable": False,
                "evidence": "a slide layout exposes a title placeholder (the cover layout)",
            }
        )
        order += 1
    if has_sections:
        skeleton.append(
            {
                "region": "sections",
                "order": order,
                "role": "section.agenda",
                "required": False,
                "repeatable": False,
                "evidence": "the presentation carries a p14:sectionLst (the agenda/section list)",
            }
        )
        order += 1
    if has_body or not skeleton:
        skeleton.append(
            {
                "region": "body",
                "order": order,
                "role": "section.body",
                "required": True,
                "repeatable": True,
                "freeform": True,
                "evidence": "the content slides after the cover (one section per heading)",
            }
        )

    return {"ordered": True, "skeleton": skeleton}


# ---------------------------------------------------------------------------
# Native-component inventory (typed counts for the survival check; plan P5)
# ---------------------------------------------------------------------------
def _slide_component_counts(slide) -> dict:
    """Count the native components on one slide by typed family.

    Families are structural, language-invariant python-pptx signals:
      - ``table``: a ``graphicFrame`` carrying an ``a:tbl`` (``shape.has_table``);
      - ``chart``: a ``graphicFrame`` carrying a ``c:chart`` (``shape.has_chart``);
      - ``picture``: a PICTURE shape (``shape_type == MSO_SHAPE_TYPE.PICTURE``).

    Placeholder pictures are counted too (a picture placeholder is still a native
    image). Text placeholders are NOT components - they are the body the generator
    fills - so they are excluded. The counts back the component-survival check that
    fires when a native object present in the shell has no counterpart in output.
    """
    counts = {"table": 0, "chart": 0, "picture": 0}
    for shape in slide.shapes:
        if getattr(shape, "has_table", False):
            counts["table"] += 1
        elif getattr(shape, "has_chart", False):
            counts["chart"] += 1
        elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            counts["picture"] += 1
    return counts


def inventory_components(prs) -> dict:
    """Return the deck-wide native-component totals (``table``/``chart``/``picture``).

    Deterministic and recomputable from the live deck at generate time, so a
    shell-vs-output diff (the component-survival check) is model-free. The totals are
    the sum across all slides; a per-slide breakdown is available via
    :func:`slide_component_inventory` for the extractor catalog.
    """
    totals = {"table": 0, "chart": 0, "picture": 0}
    for slide in prs.slides:
        for family, n in _slide_component_counts(slide).items():
            totals[family] += n
    return totals


def slide_component_inventory(prs) -> list[dict]:
    """Return a per-slide typed native-component inventory (extractor catalog peer).

    One entry per slide (in deck order) ``{"index", "layout", "components":
    {"table","chart","picture"}}`` - the typed inventory the survival check reads as
    its shell baseline. A slide with no native object reports all-zero counts (its
    text placeholders are not components).
    """
    out: list[dict] = []
    for i, slide in enumerate(prs.slides):
        out.append(
            {
                "index": i,
                "layout": slide.slide_layout.name,
                "components": _slide_component_counts(slide),
            }
        )
    return out


# ---------------------------------------------------------------------------
# Live-deck resolution helpers (used by generate to reconcile, not rebuild)
# ---------------------------------------------------------------------------
def structural_slide_indices(prs) -> list[int]:
    """Return the indices of slides that must be KEPT (cover + structural).

    Demo slides (every text run is an unedited layout prompt) are the only slides a
    reconcile may clear; cover and structural slides are preserved. Recomputed from
    the live deck so it is robust at generate time.
    """
    return [c["index"] for c in classify_slides(prs) if c["region"] != "demo"]


def demo_slide_indices(prs) -> list[int]:
    """Return the indices of demo slides (unedited template-prompt slides)."""
    return [c["index"] for c in classify_slides(prs) if c["region"] == "demo"]
