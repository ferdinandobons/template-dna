# SPDX-License-Identifier: MIT
"""The IntermediateDocument (IID) - the brand-agnostic flow content model.

This is "semantic space" (§2, §5.3): the user/LLM authors an ordered list of
typed *blocks* carrying **intent**, never presentation. The resolver later maps
each block to a concrete brand artifact. **No block ever names a style, hex,
font, or layout** - that is the whole point, and the reason off-brand output is
impossible by construction.

The flow model (this module) serves ``docx`` and ``pptx``. The grid model for
``xlsx`` lives separately in ``brandkit.grid.model`` (§2 "two content models,
one spine").

Block catalog (§5.3) - ~20 types::

    heading paragraph list table callout kpi chart smartart component section
    cover caption toc image quote divider pagebreak

Inline text is a **rich-run array** (``runs: [{t, b?, i?, u?, code?, link?}]``);
a bare ``text: "..."`` is accepted sugar and normalized to runs on parse.

Public surface
--------------
- The ``Block`` dataclasses + the :data:`BLOCK_TYPES` registry.
- :class:`IntermediateDocument` with ``cover``, ``blocks``, and ``meta``.
- ``from_dict`` / ``to_dict`` round-trips on every block and on the document.
- :func:`parse_idoc` - accept a loose JSON dict (the on-the-wire form) and
  return a validated :class:`IntermediateDocument`.

Every block subclasses :class:`Block`, which carries the discriminator
``type`` and an optional free-form ``id`` and ``meta``. ``from_dict`` dispatches
on ``type`` via :data:`BLOCK_TYPES`.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, ClassVar, Optional

from brandkit.common import text as textutil

Run = dict  # see brandkit.common.text for the canonical run shape


# ---------------------------------------------------------------------------
# Block base
# ---------------------------------------------------------------------------
@dataclass
class Block:
    """Base class for every IID block.

    Attributes:
        id: optional author-supplied identifier (stable across regenerations;
            used by repairs to locate "the offending block").
        meta: optional free-form annotations (e.g. ``{"source_line": 12}``) that
            the resolver ignores but QA/repair may read.
    """

    #: The discriminator string written to ``type`` in JSON. Overridden per
    #: subclass. The base class itself is abstract (never serialized directly).
    TYPE: ClassVar[str] = "block"

    id: Optional[str] = None
    meta: dict = field(default_factory=dict)

    # -- serialization -----------------------------------------------------
    def to_dict(self) -> dict:
        """Serialize to a JSON-ready dict (subclasses extend ``_payload``)."""
        out: dict = {"type": self.TYPE}
        if self.id is not None:
            out["id"] = self.id
        if self.meta:
            out["meta"] = dict(self.meta)
        out.update(self._payload())
        return out

    def _payload(self) -> dict:
        """Return the subclass-specific fields (excluding base fields)."""
        return {}

    @classmethod
    def _common_kwargs(cls, data: dict) -> dict:
        """Extract the base-class kwargs (``id``, ``meta``) from a dict."""
        return {"id": data.get("id"), "meta": dict(data.get("meta") or {})}


# ---------------------------------------------------------------------------
# Concrete blocks
# ---------------------------------------------------------------------------
@dataclass
class Heading(Block):
    """A section heading. Resolves to ``heading.{level}``."""

    TYPE: ClassVar[str] = "heading"
    level: int = 1
    runs: list[Run] = field(default_factory=list)

    def _payload(self) -> dict:
        return {"level": self.level, "runs": self.runs}

    @classmethod
    def from_dict(cls, data: dict) -> "Heading":
        return cls(
            level=int(data.get("level", 1)),
            runs=textutil.normalize_runs(data.get("runs"), text=data.get("text")),
            **cls._common_kwargs(data),
        )


@dataclass
class Paragraph(Block):
    """A body paragraph. Resolves to ``paragraph.{variant or default}``."""

    TYPE: ClassVar[str] = "paragraph"
    runs: list[Run] = field(default_factory=list)
    variant: Optional[str] = None

    def _payload(self) -> dict:
        out: dict = {"runs": self.runs}
        if self.variant:
            out["variant"] = self.variant
        return out

    @classmethod
    def from_dict(cls, data: dict) -> "Paragraph":
        return cls(
            runs=textutil.normalize_runs(data.get("runs"), text=data.get("text")),
            variant=data.get("variant"),
            **cls._common_kwargs(data),
        )


# Bounds list-item nesting so a pathologically deep ``items`` chain raises the
# contracted IIDParseError instead of an uncaught RecursionError.
_MAX_LIST_NESTING = 64


@dataclass
class ListItem:
    """One item in a :class:`ListBlock` (may nest sub-items)."""

    runs: list[Run] = field(default_factory=list)
    level: int = 0
    items: list["ListItem"] = field(default_factory=list)

    def to_dict(self) -> dict:
        out: dict = {"runs": self.runs, "level": self.level}
        if self.items:
            out["items"] = [i.to_dict() for i in self.items]
        return out

    @classmethod
    def from_dict(
        cls, data: dict | str, *, level: int = 0, _depth: int = 0
    ) -> "ListItem":
        if _depth > _MAX_LIST_NESTING:
            raise IIDParseError(
                f"list item nesting exceeded max depth ({_MAX_LIST_NESTING})"
            )
        if isinstance(data, str):
            return cls(runs=textutil.plain_run(data), level=level)
        runs = textutil.normalize_runs(data.get("runs"), text=data.get("text"))
        lvl = int(data.get("level", level))
        kids = [
            cls.from_dict(c, level=lvl + 1, _depth=_depth + 1)
            for c in (data.get("items") or [])
        ]
        return cls(runs=runs, level=lvl, items=kids)


@dataclass
class ListBlock(Block):
    """An ordered or unordered list. Resolves to ``list.{bullet|number}.{level}``."""

    TYPE: ClassVar[str] = "list"
    ordered: bool = False
    items: list[ListItem] = field(default_factory=list)

    def _payload(self) -> dict:
        return {"ordered": self.ordered, "items": [i.to_dict() for i in self.items]}

    @classmethod
    def from_dict(cls, data: dict) -> "ListBlock":
        return cls(
            ordered=bool(data.get("ordered", False)),
            items=[ListItem.from_dict(i) for i in (data.get("items") or [])],
            **cls._common_kwargs(data),
        )


@dataclass
class TableCell:
    """A single table cell (rich-run content, optional header flag/span)."""

    runs: list[Run] = field(default_factory=list)
    header: bool = False
    colspan: int = 1
    rowspan: int = 1

    def to_dict(self) -> dict:
        out: dict = {"runs": self.runs}
        if self.header:
            out["header"] = True
        if self.colspan != 1:
            out["colspan"] = self.colspan
        if self.rowspan != 1:
            out["rowspan"] = self.rowspan
        return out

    @classmethod
    def from_dict(cls, data: Any) -> "TableCell":
        if isinstance(data, str):
            return cls(runs=textutil.plain_run(data))
        return cls(
            runs=textutil.normalize_runs(data.get("runs"), text=data.get("text")),
            header=bool(data.get("header", False)),
            colspan=int(data.get("colspan", 1)),
            rowspan=int(data.get("rowspan", 1)),
        )


@dataclass
class Table(Block):
    """A table. Resolves to ``table.{role}`` + a header style."""

    TYPE: ClassVar[str] = "table"
    # Header-row cells, each a RICH run-list (``from_dict`` normalizes every column
    # to one, preserving multi-run emphasis). A bare run dict / string from a direct
    # construction is tolerated by the writers. ``[]`` means no header row.
    columns: list = field(default_factory=list)
    rows: list[list[TableCell]] = field(default_factory=list)
    caption: Optional[list[Run]] = None
    role: str = "default"

    def _payload(self) -> dict:
        out: dict = {
            "columns": self.columns,
            "rows": [[c.to_dict() for c in row] for row in self.rows],
            "role": self.role,
        }
        if self.caption:
            out["caption"] = self.caption
        return out

    @classmethod
    def from_dict(cls, data: dict) -> "Table":
        # Keep EVERY run of each column header (not just the first) so multi-run
        # emphasis in a header survives the round-trip; each column becomes a
        # run-list, mirroring how a body TableCell holds its runs. The same loose
        # shapes a cell accepts are accepted: a string, a single run dict, a
        # run-list, or a ``{"runs": [...]}`` / ``{"text": "..."}`` wrapper.
        columns = [
            textutil.normalize_runs(c.get("runs"), text=c.get("text"))
            if isinstance(c, dict) and ("runs" in c or "text" in c)
            else textutil.normalize_runs(c)
            for c in (data.get("columns") or [])
        ]
        rows = [
            [TableCell.from_dict(c) for c in row] for row in (data.get("rows") or [])
        ]
        caption = (
            textutil.normalize_runs(data.get("caption"))
            if data.get("caption")
            else None
        )
        return cls(
            columns=columns,
            rows=rows,
            caption=caption,
            role=data.get("role", "default"),
            **cls._common_kwargs(data),
        )


@dataclass
class Callout(Block):
    """A highlighted box. ``intent`` is semantic; the BRAND picks the color.

    Resolves to ``callout.{intent}``. ``intent`` is one of the semantic values
    (``info`` | ``warning`` | ``danger`` | ``success`` | ``note``); the profile
    decides whether that maps to blue/red/etc.
    """

    TYPE: ClassVar[str] = "callout"
    intent: str = "info"
    runs: list[Run] = field(default_factory=list)
    title: Optional[list[Run]] = None

    def _payload(self) -> dict:
        out: dict = {"intent": self.intent, "runs": self.runs}
        if self.title:
            out["title"] = self.title
        return out

    @classmethod
    def from_dict(cls, data: dict) -> "Callout":
        return cls(
            intent=data.get("intent", "info"),
            runs=textutil.normalize_runs(data.get("runs"), text=data.get("text")),
            title=textutil.normalize_runs(data["title"]) if data.get("title") else None,
            **cls._common_kwargs(data),
        )


@dataclass
class KpiItem:
    """One metric in a :class:`Kpi` block."""

    label: str = ""
    value: str = ""
    delta: Optional[str] = None

    def to_dict(self) -> dict:
        out = {"label": self.label, "value": self.value}
        if self.delta is not None:
            out["delta"] = self.delta
        return out

    @classmethod
    def from_dict(cls, data: dict) -> "KpiItem":
        return cls(
            label=str(data.get("label", "")),
            value=str(data.get("value", "")),
            delta=data.get("delta"),
        )


@dataclass
class Kpi(Block):
    """A KPI / metric card group. Resolves to ``component:kpi.{layout}``."""

    TYPE: ClassVar[str] = "kpi"
    items: list[KpiItem] = field(default_factory=list)
    layout: Optional[str] = None

    def _payload(self) -> dict:
        out: dict = {"items": [i.to_dict() for i in self.items]}
        if self.layout:
            out["layout"] = self.layout
        return out

    @classmethod
    def from_dict(cls, data: dict) -> "Kpi":
        return cls(
            items=[KpiItem.from_dict(i) for i in (data.get("items") or [])],
            layout=data.get("layout"),
            **cls._common_kwargs(data),
        )


@dataclass
class Chart(Block):
    """A chart. Resolves to ``chart.{type}`` (clone-fill a template part)."""

    TYPE: ClassVar[str] = "chart"
    chart_type: str = "bar"
    series: list[dict] = field(default_factory=list)  # [{"name", "values":[...]}]
    categories: list[str] = field(default_factory=list)
    title: Optional[str] = None

    def _payload(self) -> dict:
        out: dict = {
            "chart_type": self.chart_type,
            "series": self.series,
            "categories": self.categories,
        }
        if self.title:
            out["title"] = self.title
        return out

    @classmethod
    def from_dict(cls, data: dict) -> "Chart":
        return cls(
            chart_type=data.get("chart_type", "bar"),
            series=list(data.get("series") or []),
            categories=list(data.get("categories") or []),
            title=data.get("title"),
            **cls._common_kwargs(data),
        )


@dataclass
class SmartArt(Block):
    """A diagram. Resolves to ``smartart.{diagram}`` (clone-fill / render-img)."""

    TYPE: ClassVar[str] = "smartart"
    diagram: str = "process"
    nodes: list[dict] = field(default_factory=list)  # [{"text", "children":[...]}]

    def _payload(self) -> dict:
        return {"diagram": self.diagram, "nodes": self.nodes}

    @classmethod
    def from_dict(cls, data: dict) -> "SmartArt":
        return cls(
            diagram=data.get("diagram", "process"),
            nodes=list(data.get("nodes") or []),
            **cls._common_kwargs(data),
        )


@dataclass
class Component(Block):
    """A reference to a reusable single fragment. Expands pre-resolve.

    ``ref`` names a ``components/<ref>`` entry in the profile; ``slots`` fills
    its render contract. ``expand_components`` (the generate leg) replaces this
    with primitive sub-blocks before resolution.
    """

    TYPE: ClassVar[str] = "component"
    ref: str = ""
    slots: dict = field(default_factory=dict)

    def _payload(self) -> dict:
        return {"ref": self.ref, "slots": self.slots}

    @classmethod
    def from_dict(cls, data: dict) -> "Component":
        return cls(
            ref=data.get("ref", ""),
            slots=dict(data.get("slots") or {}),
            **cls._common_kwargs(data),
        )


@dataclass
class Section(Block):
    """A reference to a multi-block reusable unit. Expands pre-resolve.

    ``ref`` names a ``sections/<ref>`` entry; ``slots`` fills it. Like
    :class:`Component`, expanded to primitives before resolution.
    """

    TYPE: ClassVar[str] = "section"
    ref: str = ""
    slots: dict = field(default_factory=dict)

    def _payload(self) -> dict:
        return {"ref": self.ref, "slots": self.slots}

    @classmethod
    def from_dict(cls, data: dict) -> "Section":
        return cls(
            ref=data.get("ref", ""),
            slots=dict(data.get("slots") or {}),
            **cls._common_kwargs(data),
        )


@dataclass
class Caption(Block):
    """A figure/table caption line. Resolves to ``caption``."""

    TYPE: ClassVar[str] = "caption"
    runs: list[Run] = field(default_factory=list)
    target: Optional[str] = None  # "figure" | "table" | None

    def _payload(self) -> dict:
        out: dict = {"runs": self.runs}
        if self.target:
            out["target"] = self.target
        return out

    @classmethod
    def from_dict(cls, data: dict) -> "Caption":
        return cls(
            runs=textutil.normalize_runs(data.get("runs"), text=data.get("text")),
            target=data.get("target"),
            **cls._common_kwargs(data),
        )


@dataclass
class Toc(Block):
    """A table-of-contents marker. The generator refreshes the live TOC field."""

    TYPE: ClassVar[str] = "toc"
    title: Optional[str] = None
    max_level: int = 3

    def _payload(self) -> dict:
        out: dict = {"max_level": self.max_level}
        if self.title:
            out["title"] = self.title
        return out

    @classmethod
    def from_dict(cls, data: dict) -> "Toc":
        return cls(
            title=data.get("title"),
            max_level=int(data.get("max_level", 3)),
            **cls._common_kwargs(data),
        )


@dataclass
class Image(Block):
    """An image reference. Resolves to an image placement (asset or inline path).

    ``asset`` names a profile asset id; ``src`` is an external file path. Exactly
    one should be set. ``width_emu``/``height_emu`` are optional sizing hints.
    """

    TYPE: ClassVar[str] = "image"
    asset: Optional[str] = None
    src: Optional[str] = None
    alt: Optional[str] = None
    caption: Optional[list[Run]] = None
    width_emu: Optional[int] = None
    height_emu: Optional[int] = None

    def _payload(self) -> dict:
        out: dict = {}
        for k in ("asset", "src", "alt"):
            v = getattr(self, k)
            if v is not None:
                out[k] = v
        if self.caption:
            out["caption"] = self.caption
        if self.width_emu is not None:
            out["width_emu"] = self.width_emu
        if self.height_emu is not None:
            out["height_emu"] = self.height_emu
        return out

    @classmethod
    def from_dict(cls, data: dict) -> "Image":
        return cls(
            asset=data.get("asset"),
            src=data.get("src"),
            alt=data.get("alt"),
            caption=textutil.normalize_runs(data["caption"])
            if data.get("caption")
            else None,
            width_emu=data.get("width_emu"),
            height_emu=data.get("height_emu"),
            **cls._common_kwargs(data),
        )


@dataclass
class Quote(Block):
    """A block quotation. Resolves to ``quote`` (falls back to body)."""

    TYPE: ClassVar[str] = "quote"
    runs: list[Run] = field(default_factory=list)
    attribution: Optional[list[Run]] = None

    def _payload(self) -> dict:
        out: dict = {"runs": self.runs}
        if self.attribution:
            out["attribution"] = self.attribution
        return out

    @classmethod
    def from_dict(cls, data: dict) -> "Quote":
        return cls(
            runs=textutil.normalize_runs(data.get("runs"), text=data.get("text")),
            attribution=textutil.normalize_runs(data["attribution"])
            if data.get("attribution")
            else None,
            **cls._common_kwargs(data),
        )


@dataclass
class Divider(Block):
    """A horizontal rule / separator. Resolves to a brand divider."""

    TYPE: ClassVar[str] = "divider"

    @classmethod
    def from_dict(cls, data: dict) -> "Divider":
        return cls(**cls._common_kwargs(data))


@dataclass
class PageBreak(Block):
    """An explicit page (docx) or slide (pptx) break."""

    TYPE: ClassVar[str] = "pagebreak"

    @classmethod
    def from_dict(cls, data: dict) -> "PageBreak":
        return cls(**cls._common_kwargs(data))


# ---------------------------------------------------------------------------
# Cover (special: not in the flow; sits in IntermediateDocument.cover)
# ---------------------------------------------------------------------------
@dataclass
class Cover:
    """The semantic cover slots. Filled into the profile's cover anchors.

    The author supplies semantic slots only (``title``, ``subtitle``,
    ``fields``); the resolver maps them to whatever anchors the shell has. An
    absent cover (``None`` on the document) means "leave the shell cover as-is".

    Attributes:
        title: rich-run title (or None).
        subtitle: rich-run subtitle (or None).
        fields: a free-form ``{slot: value}`` map for extra cover anchors
            (e.g. ``{"doc_id": "RPT-2026-014", "date": "2026-06-04"}``). Values
            are plain strings; the resolver matches keys to cover anchor ids.
    """

    title: Optional[list[Run]] = None
    subtitle: Optional[list[Run]] = None
    fields: dict = field(default_factory=dict)

    def to_dict(self) -> dict:
        out: dict = {}
        if self.title:
            out["title"] = self.title
        if self.subtitle:
            out["subtitle"] = self.subtitle
        if self.fields:
            out["fields"] = dict(self.fields)
        return out

    @classmethod
    def from_dict(cls, data: Optional[dict]) -> Optional["Cover"]:
        if not data:
            return None
        return cls(
            title=textutil.normalize_runs(data["title"]) if data.get("title") else None,
            subtitle=textutil.normalize_runs(data["subtitle"])
            if data.get("subtitle")
            else None,
            fields=dict(data.get("fields") or {}),
        )


# ---------------------------------------------------------------------------
# Block registry (discriminator -> class)
# ---------------------------------------------------------------------------
BLOCK_TYPES: dict[str, type] = {
    Heading.TYPE: Heading,
    Paragraph.TYPE: Paragraph,
    ListBlock.TYPE: ListBlock,
    Table.TYPE: Table,
    Callout.TYPE: Callout,
    Kpi.TYPE: Kpi,
    Chart.TYPE: Chart,
    SmartArt.TYPE: SmartArt,
    Component.TYPE: Component,
    Section.TYPE: Section,
    Caption.TYPE: Caption,
    Toc.TYPE: Toc,
    Image.TYPE: Image,
    Quote.TYPE: Quote,
    Divider.TYPE: Divider,
    PageBreak.TYPE: PageBreak,
}

#: Every flow-block type string, for callers that need the closed set.
BLOCK_TYPE_NAMES: frozenset[str] = frozenset(BLOCK_TYPES)


class IIDParseError(ValueError):
    """Raised when an IID dict cannot be parsed into the typed model."""


def block_from_dict(data: dict) -> Block:
    """Build the right :class:`Block` subclass from a dict with a ``type`` key.

    Raises:
        IIDParseError: if ``type`` is missing or unknown.
    """
    if not isinstance(data, dict):
        raise IIDParseError(f"block must be an object, got {type(data).__name__}")
    btype = data.get("type")
    cls = BLOCK_TYPES.get(btype)
    if cls is None:
        raise IIDParseError(
            f"unknown block type {btype!r}; expected one of {sorted(BLOCK_TYPE_NAMES)}"
        )
    return cls.from_dict(data)  # type: ignore[attr-defined]


# ---------------------------------------------------------------------------
# The document
# ---------------------------------------------------------------------------
@dataclass
class IntermediateDocument:
    """An ordered flow of typed blocks plus an optional cover and metadata.

    Attributes:
        blocks: the body flow, in document order.
        cover: optional :class:`Cover` slots (None = leave shell cover as-is).
        meta: free-form document metadata (``title``, ``author``, ``lang``…)
            that the generator may use for document properties but never for
            styling decisions.
    """

    blocks: list[Block] = field(default_factory=list)
    cover: Optional[Cover] = None
    meta: dict = field(default_factory=dict)

    # -- serialization -----------------------------------------------------
    def to_dict(self) -> dict:
        out: dict = {"blocks": [b.to_dict() for b in self.blocks]}
        if self.cover is not None:
            out["cover"] = self.cover.to_dict()
        if self.meta:
            out["meta"] = dict(self.meta)
        return out

    @classmethod
    def from_dict(cls, data: dict) -> "IntermediateDocument":
        """Build a document from a loose dict (see :func:`parse_idoc`)."""
        if not isinstance(data, dict):
            raise IIDParseError("idoc must be a JSON object")
        raw_blocks = data.get("blocks")
        if raw_blocks is None:
            raise IIDParseError("idoc.blocks: required list of blocks")
        if not isinstance(raw_blocks, list):
            raise IIDParseError("idoc.blocks: must be a list")
        blocks = [block_from_dict(b) for b in raw_blocks]
        return cls(
            blocks=blocks,
            cover=Cover.from_dict(data.get("cover")),
            meta=dict(data.get("meta") or {}),
        )

    # -- convenience -------------------------------------------------------
    def iter_blocks(self, *, types: Optional[frozenset[str]] = None):
        """Yield blocks, optionally filtered to the given ``type`` strings."""
        for b in self.blocks:
            if types is None or b.TYPE in types:
                yield b


def parse_idoc(data: dict) -> IntermediateDocument:
    """Parse and validate a JSON IID dict into an :class:`IntermediateDocument`.

    This is the public entry point the generate leg calls on author input. It
    normalizes the ``text:`` sugar to rich runs, dispatches every block on its
    ``type`` discriminator, and raises :class:`IIDParseError` with a precise
    message on the first malformed block.
    """
    return IntermediateDocument.from_dict(data)
