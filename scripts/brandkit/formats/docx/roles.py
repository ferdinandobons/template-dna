# SPDX-License-Identifier: MIT
"""DOCX role inference for the M1 vertical slice."""
from __future__ import annotations

from typing import Iterable, Optional

from docx.enum.style import WD_STYLE_TYPE

from brandkit.common import text as textutil
from brandkit.formats.docx import structure
from brandkit.profile import schema

# Builtin style ids the role-inference floor must NEVER nominate as a *brand*
# artifact. These are Word's own out-of-the-box styles: matching one (e.g. the
# builtin ``Footnote Text`` carrying the ``note`` lexicon token, or ``Table
# Normal``/``Table Grid`` carrying ``table``/``grid``) would let a builtin win
# over a real custom brand style. Membership is by STYLE ID (language-invariant:
# Word stores builtin ids identically in every UI language), so this is not a
# brand/language matching rule - it is the structural "is this Word's builtin"
# test. Lowercased, space-stripped for robust comparison.
_BUILTIN_FOOTNOTE_ENDNOTE_IDS: frozenset[str] = frozenset(
    {"footnotetext", "endnotetext", "footnotereference", "endnotereference"}
)
_BUILTIN_TABLE_IDS: frozenset[str] = frozenset(
    {"tablenormal", "tablegrid", "normaltable", "plaintable"}
)


def _style_id(style) -> str:
    return getattr(style, "style_id", None) or style.name


def _norm_id(style) -> str:
    """Lowercased, space-stripped style id for builtin-membership tests."""
    sid = getattr(style, "style_id", None) or getattr(style, "name", None) or ""
    return sid.replace(" ", "").lower()


def _is_custom_style(style) -> bool:
    """True when a style is an author-defined (``w:customStyle="1"``) style.

    A custom style is the structural proof that the template author added it (a
    brand style), as opposed to a Word builtin. Read straight off the style XML;
    falls back to False (treat as builtin) when the attribute is absent.
    """
    el = getattr(style, "element", None)
    if el is None:
        return False
    val = el.get(structure.w("customStyle"))
    return val in ("1", "true")


def _paragraph_has_box(style) -> bool:
    """True when a paragraph style carries callout-shaped direct formatting:
    a ``w:pPr/w:shd`` fill AND/OR a ``w:pPr/w:pBdr`` box border.

    This is the STRUCTURAL callout signal (a shaded/bordered box), language- and
    brand-invariant, preferred over the name-token lexicon. Either signal alone
    qualifies (a shaded note, or a bordered box)."""
    el = getattr(style, "element", None)
    if el is None:
        return False
    pPr = el.find(structure.w("pPr"))
    if pPr is None:
        return False
    has_shd = pPr.find(structure.w("shd")) is not None
    has_border = pPr.find(structure.w("pBdr")) is not None
    return has_shd or has_border


def role_usage(role_id: str) -> dict:
    """Derive the per-artifact ``usage`` annotation from a role id.

    Brand-agnostic and evidence-driven (uses the *family* of the role id, which is
    already inferred from style placement / OOXML, never a brand-specific name):

      - ``cover.*``          -> scope=cover, structural, required, order=0
      - ``toc``              -> scope=toc,   structural, required, order=1
      - everything else      -> scope=body,  freeform,   not required, order=null

    ``placement="structural"`` means the artifact is part of the ordered skeleton
    and must appear in its slot; ``placement="freeform"`` means it is used on demand
    inside the freeform body region.
    """
    family, _ = schema.parse_role_id(role_id)
    if family == "cover":
        return {"scope": "cover", "placement": "structural", "required": True, "order": 0}
    if family == "toc":
        return {"scope": "toc", "placement": "structural", "required": True, "order": 1}
    return {"scope": "body", "placement": "freeform", "required": False, "order": None}


def _role_entry(
    role_id: str,
    style,
    *,
    confidence: float,
    status: str,
    signal: str,
    resolver_extra: Optional[dict] = None,
) -> dict:
    resolver = {
        "type": schema.ResolverType.NAMED_STYLE.value,
        "style_id": _style_id(style),
        "style_name": style.name,
        "style_type": _style_type(style),
    }
    if resolver_extra:
        # Verbatim structural facts the generator needs in addition to the
        # load-bearing style id (e.g. a list role's ``num_id``/``ilvl`` from the
        # document's own numbering part). They are data, never matching rules; the
        # only load-bearing ref the deterministic guards bind to is ``style_id``.
        resolver.update(resolver_extra)
    return {
        "resolver": resolver,
        "appearance": {},
        "usage": role_usage(role_id),
        "verified": True,
        "confidence": confidence,
        "status": status,
        "evidence": {"signal": signal},
    }


def _style_type(style) -> str:
    if style.type == WD_STYLE_TYPE.PARAGRAPH:
        return "paragraph"
    if style.type == WD_STYLE_TYPE.TABLE:
        return "table"
    if style.type == WD_STYLE_TYPE.CHARACTER:
        return "character"
    return str(style.type)


def infer_roles(doc) -> dict:
    """Infer the template's role registry.

    Evidence order (plan §5 / M-i-8 lexicon demotion) is STRICT: a role is bound by
    a **structural** signal whenever one exists, and the multilingual name-token
    lexicon (:data:`text.NAME_TOKEN_LEXICON`) is consulted ONLY as a last-resort
    weak prior, for the handful of roles that have no builtin style id at all.

      PRIMARY (structural, language-invariant; high confidence / robust):
        builtin style ids - ``Normal`` -> paragraph, ``Heading N`` -> heading.N,
        ``Title`` -> cover.title, ``Caption`` -> caption, ``Quote`` -> quote,
        ``TOCHeading`` -> toc, ``Table Grid`` -> table.default fallback.
      WEAK PRIOR (lexicon name-token containment; best_effort / low confidence so
        it can NEVER gate output): only ``callout.info`` (no builtin style exists)
        and a *custom* brand table style. A role bound solely by the lexicon is
        stamped best_effort and its load-bearing resolver ref is still a verbatim
        structural id the deterministic guards bind to.
    """
    roles: dict = {"_index": []}

    def add(
        role_id: str,
        style,
        confidence: float,
        status: str,
        signal: str,
        *,
        resolver_extra: Optional[dict] = None,
    ) -> None:
        if role_id in roles:
            return
        roles[role_id] = _role_entry(
            role_id,
            style,
            confidence=confidence,
            status=status,
            signal=signal,
            resolver_extra=resolver_extra,
        )
        roles["_index"].append(role_id)

    paragraph_styles = [s for s in doc.styles if s.type == WD_STYLE_TYPE.PARAGRAPH]
    table_styles = [s for s in doc.styles if s.type == WD_STYLE_TYPE.TABLE]

    normal = _find_style(paragraph_styles, "Normal")
    if normal is not None:
        add("paragraph", normal, 1.0, schema.Status.ROBUST.value, "builtin Normal")

    for level in (1, 2, 3):
        style = _find_style(paragraph_styles, f"Heading {level}")
        if style is not None:
            add(schema.role_id("heading", level), style, 1.0, schema.Status.ROBUST.value, f"builtin Heading {level}")

    title = _find_style(paragraph_styles, "Title")
    if title is not None:
        add("cover.title", title, 0.8, schema.Status.BEST_EFFORT.value, "builtin Title style")

    toc = _toc_heading_style(paragraph_styles)
    if toc is not None:
        add("toc", toc, 0.7, schema.Status.BEST_EFFORT.value, "TOC/contents-named paragraph style")

    # Callout (D6). PRIMARY (structural): a paragraph style carrying a shaded /
    # bordered BOX (``w:pPr/w:shd`` and/or ``w:pBdr``) is a real callout container,
    # language- and brand-invariant. Builtin Footnote/Endnote styles are EXCLUDED
    # up front so the weak ``note`` lexicon token can never let Word's builtin
    # ``Footnote Text`` win over a real brand callout style. WEAK PRIOR (lexicon)
    # is the last-resort tiebreaker, over NON-builtin styles only.
    callout_candidates = [
        s for s in paragraph_styles
        if _norm_id(s) not in _BUILTIN_FOOTNOTE_ENDNOTE_IDS
    ]
    callout = _first_boxed_style(callout_candidates)
    if callout is not None:
        add("callout.info", callout, 0.7, schema.Status.BEST_EFFORT.value,
            "paragraph shd/border box (structural callout signal)")
    else:
        callout = _best_name_token_style(callout_candidates, "callout")
        if callout is not None:
            add("callout.info", callout, 0.62, schema.Status.BEST_EFFORT.value,
                "name-token callout over non-builtin styles (weak prior)")

    caption = _find_style(paragraph_styles, "Caption")
    if caption is not None:
        add("caption", caption, 0.9, schema.Status.ROBUST.value, "builtin Caption")

    quote = _find_style(paragraph_styles, "Quote")
    if quote is not None:
        add("quote", quote, 0.9, schema.Status.ROBUST.value, "builtin Quote")

    # table.default (D2). PRIMARY (structural): a *custom* (``w:customStyle="1"``)
    # ``w:type="table"`` style is the author's own brand table style and is bound
    # ahead of the lexicon/builtin floor. Only when there is no custom table style
    # at all does the WEAK PRIOR name-token lexicon, then the builtin ``Table
    # Grid`` floor, apply. The load-bearing resolver ref is the verbatim style id
    # either way, so the lexicon can never widen the brand guarantee.
    custom_table = next((s for s in table_styles if _is_custom_style(s)), None)
    if custom_table is not None:
        add("table.default", custom_table, 0.85, schema.Status.ROBUST.value,
            "custom w:type='table' style (structural brand table)")
    else:
        table = _best_name_token_style(table_styles, "table") or _find_style(table_styles, "Table Grid")
        if table is not None:
            add("table.default", table, 0.72, schema.Status.BEST_EFFORT.value,
                "table style candidate (weak prior)")

    # List roles (D1). PRIMARY (structural): a paragraph style whose definition
    # carries a ``w:pPr/w:numPr`` references a real numbering definition; resolve
    # its FAMILY (bullet vs number) from the abstractNum ``w:numFmt`` field code
    # and bind it to ``list.<family>.<level>`` (1-based level = ``ilvl + 1``). The
    # resolver carries the verbatim ``num_id``/``ilvl`` so generation can re-assert
    # ``w:numPr`` (a style's numPr is not inherited onto an ``add_paragraph``). The
    # load-bearing ref is the verbatim style id. The old ``Normal``-as-list floor
    # remains ONLY when the template has no real numbered list style at all.
    _nominate_list_styles(doc, paragraph_styles, add)
    if normal is not None and not _has_any_list_role(roles):
        add("list.bullet.1", normal, 0.45, schema.Status.BEST_EFFORT.value, "M1 fallback list paragraph")

    return roles


def _has_any_list_role(roles: dict) -> bool:
    """True if any ``list.*`` role has already been nominated."""
    return any(rid.startswith("list.") for rid in roles if rid != "_index")


def _nominate_list_styles(doc, paragraph_styles, add) -> None:
    """Bind ``list.{bullet,number}.{level}`` to the document's real list styles.

    For each paragraph style whose definition carries a ``w:numPr``, read its
    ``(num_id, ilvl)`` and resolve the list family from the numbering part's
    ``w:numFmt`` field code (NEVER the style name). The resolver carries the
    verbatim ``num_id`` / ``ilvl`` the generator re-asserts.

    CUSTOM-FIRST (CC-2): a template's brand list styles are author-defined
    (``w:customStyle="1"``) while Word's default template ALSO ships builtin
    latent list styles (``List Bullet``, ``List Number`` ...) that reference the
    SAME numbering ids and iterate BEFORE the brand styles. So custom styles are
    nominated first; a builtin list style fills a ``list.<family>.<level>`` slot
    only when no custom style claimed it. The load-bearing ref stays the verbatim
    style id, so this only ever picks a *present* style, never invents one.
    """
    styles = list(paragraph_styles)
    custom = [s for s in styles if _is_custom_style(s)]
    builtin = [s for s in styles if not _is_custom_style(s)]
    for style in custom + builtin:  # custom wins; builtin is the fallback floor
        binding = structure.style_num_binding(style)
        if binding is None:
            continue
        num_id, ilvl = binding
        family = structure.num_family_for(doc, num_id, ilvl)
        if family is None:
            continue
        level = ilvl + 1  # role ids are 1-based; ilvl is 0-based
        rid = schema.role_id("list", family, level)
        add(
            rid,
            style,
            0.85,
            schema.Status.ROBUST.value,
            f"paragraph style w:numPr -> numbering numFmt {family} (structural)",
            resolver_extra={"num_id": str(num_id), "ilvl": int(ilvl)},
        )


def _find_style(styles: Iterable, name: str):
    for style in styles:
        if style.name == name:
            return style
    return None


def _first_boxed_style(styles: Iterable):
    """Return the first CUSTOM paragraph style carrying a callout-shaped box.

    STRUCTURAL signal only (``w:pPr/w:shd`` fill and/or ``w:pBdr`` border); brand-
    and language-invariant. Restricted to author-defined (``w:customStyle="1"``)
    styles so Word builtins that incidentally carry a box (``Title``, ``Intense
    Quote``) never masquerade as a callout - those are not callout containers, and
    a custom boxed style is the real structural proof the author authored one.
    Iterates in style order for deterministic selection; returns None when no
    custom boxed style exists (the caller then falls back to the lexicon).
    """
    for style in styles:
        if _is_custom_style(style) and _paragraph_has_box(style):
            return style
    return None


def _best_name_token_style(styles: Iterable, family: str):
    """Return the first style whose display NAME contains a family lexicon token.

    WEAK PRIOR ONLY (plan §5 / M-i-8): the multilingual name-token lexicon is the
    engine's weakest, last-resort signal. It is consulted only for roles with no
    builtin structural style id (``callout``, a custom brand ``table`` style) and
    its result is always stamped best_effort / low confidence by the caller so it
    can never gate output. It matches a style DISPLAY NAME only, never rendered
    body text, so it cannot leak a brand word into the matching of content.
    """
    tokens = textutil.NAME_TOKEN_LEXICON.get(family, frozenset())
    for style in styles:
        lname = style.name.lower()
        if any(token in lname for token in tokens):
            return style
    return None


def _toc_heading_style(styles: Iterable):
    """Return the paragraph style that heads a Table of Contents, if any.

    Prefers an explicit ``TOCHeading``-style id/name; otherwise falls back to any
    style whose name carries a multilingual contents token (``toc``/``contents``/
    ``sommario``/``indice``/``inhalt``/``contenido`` …). Brand-agnostic: the match
    is on the style *family* token, never on a brand-specific label. TOC-entry
    styles (``TOC 1``, ``TOC 2`` …) are skipped in favour of the heading style.
    """
    tokens = textutil.NAME_TOKEN_LEXICON.get("toc", frozenset())
    fallback = None
    for style in styles:
        sid = (getattr(style, "style_id", None) or "").lower()
        lname = (style.name or "").lower()
        if sid in ("tocheading", "toc heading") or "tocheading" in sid:
            return style
        if "heading" in lname and any(tok in lname for tok in tokens):
            return style
        if fallback is None and any(tok in lname for tok in tokens):
            # A plain TOC-entry style (e.g. "TOC 1"); keep as a last resort.
            fallback = style
    return fallback
