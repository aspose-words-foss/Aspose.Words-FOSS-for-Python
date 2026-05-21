"""
Lightweight document model for PDF rendering via fpdf2.

The model is genuinely recursive (Cell → list[Table] → list[Row] →
list[Cell]), so this is one of the rare modules that legitimately
needs ``from __future__ import annotations`` to keep all field
type-hints lazy — pydantic's ``model_rebuild`` then stitches the
forward references at the bottom of the file.
"""

from __future__ import annotations

from typing import Annotated, Any, Literal, Optional, Union

from pydantic import BaseModel, BeforeValidator, Field, PrivateAttr, model_serializer, model_validator

# ─────────────────────────────────────────────
# Primitives
# ─────────────────────────────────────────────


class Border(BaseModel):
    line_style: int = 0
    line_width: float = 0.0
    color: str = ""
    # REMOVED: distance_from_text, shadow


class Shading(BaseModel):
    background_pattern_color: str = ""
    foreground_pattern_color: str = ""
    # REMOVED: theme_color, theme_shade, theme_tint, theme_fill,
    #          theme_fill_shade, theme_fill_tint, foreground_tint_and_shade,
    #          background_tint_and_shade, texture


# ─────────────────────────────────────────────
# Tab stops
# ─────────────────────────────────────────────


class TabStop(BaseModel):
    position: float = 0.0
    alignment: int = 0  # TabAlignment: 0=Left,1=Center,2=Right,3=Decimal,4=Bar,5=List,6=Clear
    leader: int = 0  # TabLeader: 0=None,1=Dot,2=Dash,3=Line,4=Heavy,5=MiddleDot
    is_clear: bool = False


class TabStopCollection(BaseModel):
    tab_stops: list[TabStop] = Field(default_factory=list)

    def clear(self) -> None:
        self.tab_stops.clear()

    def add(self, position: float, alignment: int = 0, leader: int = 0) -> None:
        self.tab_stops.append(TabStop(position=position, alignment=alignment, leader=leader))
        self.tab_stops.sort(key=lambda t: t.position)

    def remove_by_position(self, position: float) -> None:
        self.tab_stops = [t for t in self.tab_stops if abs(t.position - position) > 0.01]

    def before(self, pos: float) -> TabStop | None:
        result = None
        for t in self.tab_stops:
            if t.position < pos and not t.is_clear:
                result = t
        return result

    def after(self, pos: float) -> TabStop | None:
        for t in self.tab_stops:
            if t.position > pos and not t.is_clear:
                return t
        return None

    def __len__(self) -> int:
        return len(self.tab_stops)

    def __iter__(self):
        return iter(self.tab_stops)

    def __bool__(self) -> bool:
        return bool(self.tab_stops)


# ─────────────────────────────────────────────
# Font  (was 30+ fields → 15)
# ─────────────────────────────────────────────


class Font(BaseModel):
    name: str = ""
    size: float = 0.0
    bold: bool = False
    italic: bool = False
    underline: int = 0
    color: str = ""
    strike_through: bool = False
    superscript: bool = False
    subscript: bool = False
    highlight_color: str = ""
    all_caps: bool = False
    small_caps: bool = False
    hidden: bool = False
    style_name: str = ""
    style_identifier: int = 0
    shading: Shading = Field(default_factory=Shading)
    emboss: bool = False
    engrave: bool = False
    outline: bool = False
    shadow: bool = False
    # 0=None, 1=OverSolidCircle (dot), 2=OverComma, 3=OverWhiteCircle,
    # 4=UnderSolidCircle (underDot).
    emphasis_mark: int = 0
    # 0=None, 1=LasVegasLights, 2=BlinkingBackground, 3=SparkleText,
    # 4=MarchingBlackAnts, 5=MarchingRedAnts, 6=Shimmer.
    text_effect: int = 0
    # Minimum font size in points to apply kerning; 0 disables.
    kerning: float = 0.0
    # REMOVED: name_bi, name_far_east, size_bi, bold_bi, italic_bi,
    #          double_strike_through, underline_color, scaling, spacing,
    #          position, no_proofing, locale_id, complex_script, border


# ─────────────────────────────────────────────
# Paragraph format  (was 25+ fields → 18)
# ─────────────────────────────────────────────


class ParagraphFormat(BaseModel):
    style_name: str = ""
    alignment: int = 0  # 0=Left, 1=Center, 2=Right, 3=Justify
    left_indent: float = 0.0
    right_indent: float = 0.0
    first_line_indent: float = 0.0
    space_before: float = 0.0
    space_after: float = 0.0
    space_before_auto: bool = False
    space_after_auto: bool = False
    line_spacing: float = 0.0
    line_spacing_rule: int = 0  # 0=AtLeast, 1=Exactly, 2=Multiple
    keep_with_next: bool = False
    page_break_before: bool = False
    no_space_between_paragraphs_of_same_style: bool = False
    outline_level: int = 9
    is_heading: bool = False
    is_list_item: bool = False
    shading: Shading = Field(default_factory=Shading)
    borders: list[Border] = Field(default_factory=list)
    paragraph_mark_font: Optional[Font] = None
    style_identifier: int = 0
    keep_together: bool = False
    widow_control: bool = True
    suppress_auto_hyphens: bool = False
    suppress_line_numbers: bool = False
    snap_to_grid: bool = True
    add_space_between_far_east_and_alpha: bool = True
    add_space_between_far_east_and_digit: bool = True
    auto_adjust_right_indent: bool = True
    # 0=Auto, 1=Top, 2=Center, 3=Baseline, 4=Bottom.
    baseline_alignment: int = 0
    # 12-char banded-table mask from ``<w:cnfStyle/>``.
    conditional_style: str = ""
    tab_stops: TabStopCollection = Field(default_factory=TabStopCollection)
    frame: Optional["FrameFormat"] = None
    lines_to_drop: int = 0
    # 0=None, 1=Normal, 2=Margin.
    drop_cap_position: int = 0
    # REMOVED: suppress_auto_hyphens, suppress_line_numbers,
    #          no_space_between_paragraphs_of_same_style, bidi,
    #          character_unit_*, line_unit_*, snap_to_grid


class FrameFormat(BaseModel):
    """Floating text-frame definition.  Numeric dimensions are in points."""

    width: float = 0.0
    height: float = 0.0
    # 0=AtLeast, 1=Exactly, 2=Auto.
    height_rule: int = 0
    horizontal_position: float = 0.0
    vertical_position: float = 0.0
    # 0=None, 1=Left, 2=Center, 3=Right, 4=Inside, 5=Outside.
    horizontal_alignment: int = 0
    # -1=Inline, 0=None, 1=Top, 2=Center, 3=Bottom, 4=Inside, 5=Outside.
    vertical_alignment: int = 0
    # 0=Margin, 1=Page, 3=Character.
    relative_horizontal_position: int = 0
    # 0=Margin, 1=Page, 2=Paragraph.
    relative_vertical_position: int = 0
    horizontal_distance_from_text: float = 0.0
    vertical_distance_from_text: float = 0.0
    lock_anchor: bool = False
    # 0=Inline, 1=TopBottom, 2=Square, 3=None, 4=Tight, 5=Through.
    wrap_type: int = 0


class ListFormat(BaseModel):
    is_list_item: bool = False
    list_level_number: int = 0
    list_id: int = 0


class ListLabel(BaseModel):
    """Snapshot of a list-item's rendered bullet/number label."""

    # Rendered label text (e.g. ``"1."`` / ``"-"`` / ``"a)"``).
    label_string: str = ""
    # Integer counter behind the label.
    label_value: int = 0
    font: Font = Field(default_factory=Font)


# ─────────────────────────────────────────────
# Image data
# ─────────────────────────────────────────────


class ImageData(BaseModel):
    source_filename: str = ""
    content_type: str = ""
    image_bytes: bytes = b""
    crop_left: int = 0
    crop_top: int = 0
    crop_right: int = 0
    crop_bottom: int = 0


# ─────────────────────────────────────────────
# Inline nodes
# ─────────────────────────────────────────────


class Run(BaseModel):
    type: str = Field(default="Run", alias="_type")
    text: str = ""
    font: Font = Field(default_factory=Font)

    model_config = {"populate_by_name": True}


class ShapeNode(BaseModel):
    type: str = Field(default="Shape", alias="_type")
    shape_type: int | None = None
    name: str = ""
    width: float | None = None
    height: float | None = None
    # Absolute page position of the shape's top-left corner.
    # Populated by the reader for anchored (``wp:anchor``) shapes.
    left: float = 0.0
    top: float = 0.0
    is_inline: bool | None = None
    has_image: bool | None = None
    image_data: Optional[ImageData] = None
    text_box: dict[str, Any] | None = None  # textbox paragraph content
    # Shape fill, reusing the Shading primitive that Paragraph / Cell
    # already carry.  Only ``background_color`` is populated — solid
    # fills are the only kind supported by the PDF writer.
    shading: Shading = Field(default_factory=Shading)
    # Shape outline ("stroke"), reusing the Border primitive.
    # Single-element list when the shape has a plain rectangular outline.
    borders: list[Border] = Field(default_factory=list)
    # Vertical alignment of the text-box content inside the shape's
    # bounding box: 0=Top (default), 1=Center, 2=Bottom.  Follows the
    # same integer convention as ``CellFormat.vertical_alignment``.
    vertical_alignment: int = 0
    # WrapType (see drawing.WrapType):
    # 0=Inline, 1=TopBottom, 2=Square, 3=None, 4=Tight, 5=Through.
    wrap_type: int = 0  # WrapType.INLINE
    relative_horizontal_position: int = 0
    relative_vertical_position: int = 0
    horizontal_position: float = 0.0
    vertical_position: float = 0.0
    horizontal_alignment: int = 0
    vertical_anchor_alignment: int = 0
    behind_text: bool = False
    allow_overlap: bool = True
    layout_in_cell: bool = True
    is_locked: bool = False
    # REMOVED: wrap_side, rotation, z_order

    # Runtime-only flag (not part of the JSON schema / not serialised)
    # set by the reader when the shape carries absolute page
    # coordinates extracted from an anchored group.  The PDF writer
    # uses it to branch into the absolute-positioning code path.
    _is_positioned: bool = PrivateAttr(default=False)

    model_config = {"populate_by_name": True}


class FieldStart(BaseModel):
    type: str = Field(default="FieldStart", alias="_type")
    field_type: int | None = None

    model_config = {"populate_by_name": True}


class FieldSeparator(BaseModel):
    type: str = Field(default="FieldSeparator", alias="_type")
    field_type: int | None = None

    model_config = {"populate_by_name": True}


class FieldEnd(BaseModel):
    type: str = Field(default="FieldEnd", alias="_type")
    field_type: int | None = None
    has_separator: bool = False

    model_config = {"populate_by_name": True}


class BookmarkStart(BaseModel):
    """Marks the beginning of a Word bookmark (``<w:bookmarkStart>``).

    The reader emits one of these for every named bookmark so the PDF
    writer can register the anchor's page + Y position, turning
    ``#name`` run-text references (hyperlinks, TOC entries) into real
    clickable internal links.
    """

    type: str = Field(default="BookmarkStart", alias="_type")
    name: str = ""

    model_config = {"populate_by_name": True}


class BookmarkEnd(BaseModel):
    type: str = Field(default="BookmarkEnd", alias="_type")
    name: str = ""

    model_config = {"populate_by_name": True}


# REMOVED entirely: CommentNode, FootnoteNode

ChildNode = Union[Run, ShapeNode, FieldStart, FieldSeparator, FieldEnd, BookmarkStart, BookmarkEnd]

_CHILD_NODE_CLASSES: tuple[type, ...] = (
    Run, ShapeNode, FieldStart, FieldSeparator, FieldEnd, BookmarkStart, BookmarkEnd,
)
_CHILD_NODE_TYPE_TAGS = {"Run", "Shape", "FieldStart", "FieldSeparator", "FieldEnd", "BookmarkStart", "BookmarkEnd"}


def _coerce_child_node(item: Any) -> ChildNode | None:
    """Materialise a dict-or-instance into a typed child node.

    Returns ``None`` for anything that isn't a known child kind (e.g.
    stale ``CommentNode`` payloads) — the caller drops these on the floor.
    """
    if isinstance(item, _CHILD_NODE_CLASSES):
        return item
    if isinstance(item, dict):
        t = item.get("_type", "")
        if t == "Run":
            return Run.model_validate(item)
        if t == "Shape":
            return ShapeNode.model_validate(item)
        if t == "FieldStart":
            return FieldStart.model_validate(item)
        if t == "FieldSeparator":
            return FieldSeparator.model_validate(item)
        if t == "FieldEnd":
            return FieldEnd.model_validate(item)
        if t == "BookmarkStart":
            return BookmarkStart.model_validate(item)
        if t == "BookmarkEnd":
            return BookmarkEnd.model_validate(item)
    return None


# ─────────────────────────────────────────────
# Paragraph
# ─────────────────────────────────────────────


class Paragraph(BaseModel):
    """A paragraph whose children — ``Run``, ``BookmarkStart`` / ``End``,
    ``FieldStart`` / ``Separator`` / ``End`` and inline ``ShapeNode`` —
    sit in a single ordered collection in document order.

    The collection itself is intentionally not part of the public attribute
    surface: it lives on the private ``_children`` slot.  Read access goes
    through the typed view :attr:`runs` (the analogue of Aspose.Words'
    ``Paragraph.Runs``); writers inside the library mutate ``_children``
    directly.  In JSON the children are emitted under the ``children`` key.
    """

    type: Literal["Paragraph"] = Field(default="Paragraph", alias="_type")
    paragraph_format: ParagraphFormat = Field(default_factory=ParagraphFormat)
    list_format: ListFormat | None = None
    list_label: Optional[ListLabel] = None
    text: str = Field(default="", alias="_text")

    _children: list[ChildNode] = PrivateAttr(default_factory=list)

    model_config = {"populate_by_name": True}

    @model_validator(mode="wrap")
    @classmethod
    def _absorb_children(cls, data: Any, handler) -> "Paragraph":
        """Pull the ``children`` array out of the input dict and install
        it on the private slot.  Anything that isn't a known child kind
        (unknown ``_type``, malformed entries, ``CommentNode`` …) is
        silently dropped — keeps stale JSON from rejecting the whole
        paragraph.
        """
        raw: Any = None
        if isinstance(data, dict):
            raw = data.pop("children", None)
        instance: "Paragraph" = handler(data)
        if raw:
            kept: list[ChildNode] = []
            for entry in raw:
                node = _coerce_child_node(entry)
                if node is not None:
                    kept.append(node)
            instance._children = kept
        return instance

    @model_serializer(mode="wrap")
    def _emit_children(self, handler) -> dict[str, Any]:
        """Always emit the ``children`` key in the serialised form so the
        full document round-trips through ``model_dump`` /
        ``model_validate`` with no manual plumbing on the caller side.
        """
        data = handler(self)
        data["children"] = [c.model_dump(by_alias=True) for c in self._children]
        return data

    @property
    def runs(self) -> list[Run]:
        """Typed view over the paragraph's ``Run`` children, in document
        order.  Direct counterpart of Aspose.Words' ``Paragraph.Runs``.
        """
        return [c for c in self._children if isinstance(c, Run)]


# ─────────────────────────────────────────────
# Table → Row → Cell
# ─────────────────────────────────────────────


class CellFormat(BaseModel):
    width: float = 0.0
    preferred_width: str = "Auto"  # fallback when width=0
    vertical_alignment: int = 0
    vertical_merge: int = 0  # 0=None, 1=First, 2=Previous
    horizontal_merge: int = 0  # 0=None, 1=First, 2=Previous
    top_padding: float = 0.0
    bottom_padding: float = 0.0
    left_padding: float = 0.0
    right_padding: float = 0.0
    shading: Shading = Field(default_factory=Shading)
    borders: list[Border] = Field(default_factory=list)
    orientation: int = 0
    wrap_text: bool = True
    conditional_style: str = ""
    # REMOVED: fit_text


class Cell(BaseModel):
    type: str = Field(default="Cell", alias="_type")
    cell_format: CellFormat = Field(default_factory=CellFormat)
    paragraphs: list[Paragraph] = Field(default_factory=list)
    tables: list[Table] = Field(default_factory=list)

    model_config = {"populate_by_name": True}


class RowFormat(BaseModel):
    height: float = 0.0
    height_rule: int = 0
    heading_format: bool = False
    allow_break_across_pages: bool = True  # row split control
    conditional_style: str = ""
    borders: list[Border] = Field(default_factory=list)
    # Per-row override (``<w:tblPrEx>``).  Only the width slot for now;
    # CT_TblPrExBase allows more fields, add as needed.
    preferred_width: str = "Auto"


class Row(BaseModel):
    type: str = Field(default="Row", alias="_type")
    row_format: RowFormat = Field(default_factory=RowFormat)
    cells: list[Cell] = Field(default_factory=list)

    model_config = {"populate_by_name": True}


class Table(BaseModel):
    type: Literal["Table"] = Field(default="Table", alias="_type")
    alignment: int = 0
    preferred_width: str = "Auto"
    left_indent: float = 0.0
    left_padding: float = 0.0
    right_padding: float = 0.0
    top_padding: float = 0.0
    bottom_padding: float = 0.0
    style_name: str = ""
    text_wrapping: int = 0  # 0=None, 1=Default/Around
    rows: list[Row] = Field(default_factory=list)
    # REMOVED: style_identifier, bidi, allow_auto_fit,
    #          allow_cell_spacing, cell_spacing

    _tblp_pr_attrs: dict[str, str] = PrivateAttr(default_factory=dict)

    model_config = {"populate_by_name": True}


Cell.model_rebuild()


# ─────────────────────────────────────────────
# Body children
# ─────────────────────────────────────────────


class UnknownNode(BaseModel):
    type: str = Field(default="", alias="_type")
    model_config = {"extra": "allow", "populate_by_name": True}


def _parse_body_child(v: Any) -> Paragraph | Table | UnknownNode:
    if isinstance(v, (Paragraph, Table, UnknownNode)):
        return v
    if isinstance(v, dict):
        t = v.get("_type", "")
        if t == "Paragraph":
            return Paragraph.model_validate(v)
        if t == "Table":
            return Table.model_validate(v)
        # Infer type from structure when _type is absent
        if not t:
            if "rows" in v:
                return Table.model_validate(v)
            if "runs" in v or "paragraph_format" in v or "_text" in v:
                return Paragraph.model_validate(v)
        return UnknownNode.model_validate(v)
    raise ValueError(f"Cannot parse body child: {v!r}")


BodyChild = Annotated[
    Union[Paragraph, Table, UnknownNode],
    BeforeValidator(_parse_body_child),
]


# ─────────────────────────────────────────────
# Text columns
# ─────────────────────────────────────────────


class TextColumn(BaseModel):
    width: float = 0.0
    space_after: float = 0.0


class TextColumns(BaseModel):
    count: int = 1
    evenly_spaced: bool = True
    spacing: float = 36.0
    line_between: bool = False
    columns: list[TextColumn] = Field(default_factory=list)


# ─────────────────────────────────────────────
# Section structure
# ─────────────────────────────────────────────


class PageSetup(BaseModel):
    paper_size: int = 0
    orientation: int = 0  # 0=Portrait, 1=Landscape
    top_margin: float = 0.0
    bottom_margin: float = 0.0
    left_margin: float = 0.0
    right_margin: float = 0.0
    header_distance: float = 0.0
    footer_distance: float = 0.0
    gutter: float = 0.0  # Aspose.Words: PageSetup.Gutter — extra binding margin
    page_width: float = 0.0
    page_height: float = 0.0
    page_number_style: int = 0  # Roman, Arabic, letters...
    page_starting_number: int = 1  # section starts at page N
    restart_page_numbering: bool = False  # reset counter at section
    section_start: int = 2  # 0=Continuous,1=NewColumn,2=NextPage,3=EvenPage,4=OddPage
    different_first_page_header_footer: bool = False
    odd_and_even_pages_header_footer: bool = False  # enable odd/even headers
    text_columns: Optional[TextColumns] = None
    # REMOVED: multiple_pages, sheets_per_booklet,
    #          vertical_alignment, line_number_*, bidi, borders


class HeaderFooter(BaseModel):
    type: str = Field(default="", alias="_type")
    header_footer_type: int | None = None
    children: list[BodyChild] = Field(default_factory=list)

    model_config = {"populate_by_name": True}

    @property
    def paragraphs(self) -> list[Paragraph]:
        return [c for c in self.children if isinstance(c, Paragraph)]

    @property
    def tables(self) -> list[Table]:
        return [c for c in self.children if isinstance(c, Table)]


class Body(BaseModel):
    type: str = Field(default="Body", alias="_type")
    children: list[BodyChild] = Field(default_factory=list)

    model_config = {"populate_by_name": True}


class Section(BaseModel):
    type: str = Field(default="Section", alias="_type")
    page_setup: PageSetup = Field(default_factory=PageSetup)
    body: Body = Field(default_factory=Body)
    headers_footers: list[HeaderFooter] = Field(default_factory=list)

    model_config = {"populate_by_name": True}


# ─────────────────────────────────────────────
# Styles
# ─────────────────────────────────────────────


class TableStyleFormat(BaseModel):
    """Table-level properties stored on table styles (``w:tblPr`` inside ``w:style``)."""

    borders: list[Border] = Field(default_factory=list)
    left_padding: float = 0.0
    right_padding: float = 0.0
    top_padding: float = 0.0
    bottom_padding: float = 0.0


class Style(BaseModel):
    name: str = ""
    type: int = 0  # 1=paragraph, 2=character, 3=table
    is_heading: bool = False
    base_style_name: str = ""
    next_paragraph_style_name: str = ""
    paragraph_format: ParagraphFormat | None = None
    font: Font | None = None
    table_style_format: TableStyleFormat | None = None
    style_identifier: int = 0
    built_in: bool = False
    priority: int = 99
    # REMOVED: is_quick_style, linked_style_name, aliases, automatically_update


# ─────────────────────────────────────────────
# Lists
# ─────────────────────────────────────────────


class ListLevel(BaseModel):
    number_format: str = ""
    number_style: int = 0
    start_at: int = 1
    alignment: int = 0
    number_position: float = 0.0
    text_position: float = 0.0
    # ``<w:lvl><w:rPr>`` — formatting of the bullet / number glyph itself
    # (font name, size, italic, color).  Distinct from the paragraph
    # mark font carried on each individual list-item paragraph.
    font: Optional[Font] = None
    # Restart this level's counter when a higher level resets past N;
    # -1 disables restart entirely.
    restart_after_level: int = -1
    # 0=Tab (default), 1=Space, 2=Nothing.
    trailing_character: int = 0


class ListLevelOverride(BaseModel):
    """One ``<w:lvlOverride>`` inside a concrete ``<w:num>``."""

    ilvl: int = 0
    is_start_at: bool = False
    start_at_raw: int = 1
    is_formatting: bool = False
    list_level: Optional[ListLevel] = None


class DocList(BaseModel):
    list_id: int = 0
    is_multi_level: bool = False
    levels: list[ListLevel] = Field(default_factory=list)
    # ``<w:lvlOverride>`` entries on the concrete ``<w:num>`` element.
    overrides: list[ListLevelOverride] = Field(default_factory=list)
    # REMOVED: is_list_style_definition, is_list_style_reference,
    #          is_restart_at_each_section


# ─────────────────────────────────────────────
# Root Document
# ─────────────────────────────────────────────

# REMOVED entirely (17 classes):
#   CompatibilityOptions, EndnoteOptions, FootnoteOptions,
#   HyphenationOptions, ViewOptions, WriteProtection, Watermark,
#   Theme, ThemeColors, MailMergeSettings, FontInfo,
#   BuiltInProperties, TextColumns


class Document(BaseModel):
    type: str = Field(default="Document", alias="_type")
    default_tab_stop: float = 36.0
    page_color: str = ""
    page_count: int = 0

    styles: list[Style] = Field(default_factory=list)
    lists: list[DocList] = Field(default_factory=list)
    sections: list[Section] = Field(default_factory=list)

    # Header/footer paragraphs (populated by the reader)
    header_paragraphs: list[Paragraph] = Field(default_factory=list)
    footer_paragraphs: list[Paragraph] = Field(default_factory=list)

    model_config = {"populate_by_name": True}

    # REMOVED root fields: node_type, attached_template,
    #   automatically_update_styles, compliance, custom_node_id,
    #   grammar_checked, has_macros, has_revisions, justification_mode,
    #   original_load_format, protection_type, punctuation_kerning,
    #   remove_personal_information, revisions_view, shade_form_data,
    #   show_grammatical_errors, show_spelling_errors, spelling_checked,
    #   track_revisions, versions_count,
    #   include_textboxes_footnotes_endnotes_in_stat,
    #   compatibility_options, endnote_options, footnote_options,
    #   hyphenation_options, view_options, write_protection, watermark,
    #   theme, mail_merge_settings, font_infos, built_in_properties

    # ── convenience helpers ──

    @property
    def all_paragraphs(self) -> list[Paragraph]:
        """Flat list of direct body paragraphs only (excludes table cells)."""
        result: list[Paragraph] = []
        for sec in self.sections:
            for child in sec.body.children:
                if isinstance(child, Paragraph):
                    result.append(child)
        return result

    @property
    def tables(self) -> list[Table]:
        """Flat list of all top-level tables in the body."""
        return [
            child
            for sec in self.sections
            for child in sec.body.children
            if isinstance(child, Table)
        ]

    @property
    def all_tables(self) -> list[Table]:
        """Alias for tables (backwards compatibility)."""
        return self.tables

    @property
    def text(self) -> str:
        """Plain text of the whole document (body paragraphs only)."""
        return "\n".join(p.text for p in self.all_paragraphs if p.text)

    def find_style(self, name: str) -> Style | None:
        """Look up a style by its exact name."""
        for s in self.styles:
            if s.name == name:
                return s
        return None

    def headings(self, max_level: int = 9) -> list[Paragraph]:
        """All heading paragraphs up to a given outline level."""
        return [
            p
            for p in self.all_paragraphs
            if p.paragraph_format.is_heading and p.paragraph_format.outline_level < max_level
        ]
