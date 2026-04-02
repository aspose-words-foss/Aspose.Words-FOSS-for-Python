"""
Lightweight document model for PDF rendering via fpdf2.

Strict subset of the full Aspose.Words model — every class and field
here exists in the original; nothing is added, only removed.

Usage:
    from models_pdf import Document
    doc = Document.model_validate(data)  # same JSON, unknown fields ignored
"""

from __future__ import annotations

from typing import Annotated, Any, Literal, Optional, Union

from pydantic import BaseModel, BeforeValidator, Field, field_validator

# ─────────────────────────────────────────────
# Primitives
# ─────────────────────────────────────────────


class Border(BaseModel):
    line_style: int = 0
    line_width: float = 0.0
    color: str = ""
    # REMOVED: distance_from_text, shadow


class Shading(BaseModel):
    background_color: str = ""
    # REMOVED: foreground_color, foreground_pattern_color,
    #          background_pattern_color, texture


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
    shading: Shading = Field(default_factory=Shading)  # alternate highlight mechanism
    # REMOVED: name_bi, name_far_east, size_bi, bold_bi, italic_bi,
    #          double_strike_through, underline_color, scaling, spacing,
    #          position, kerning, emboss, engrave, outline, shadow,
    #          no_proofing, locale_id, complex_script, text_effect, border


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
    outline_level: int = 9
    is_heading: bool = False
    is_list_item: bool = False
    shading: Shading = Field(default_factory=Shading)
    borders: list[Border] = Field(default_factory=list)  # paragraph frame/box
    # REMOVED: style_identifier, keep_together, widow_control,
    #          suppress_auto_hyphens, suppress_line_numbers,
    #          no_space_between_paragraphs_of_same_style, bidi,
    #          character_unit_*, line_unit_*, snap_to_grid


class ListFormat(BaseModel):
    is_list_item: bool = False
    list_level_number: int = 0
    list_id: int = 0
    list_label: str = ""


# ─────────────────────────────────────────────
# Image data
# ─────────────────────────────────────────────


class ImageData(BaseModel):
    source_filename: str = ""
    content_type: str = ""
    image_bytes: bytes = b""


# ─────────────────────────────────────────────
# Inline nodes
# ─────────────────────────────────────────────


class Run(BaseModel):
    type: str = Field("Run", alias="_type")
    text: str = ""
    font: Font = Field(default_factory=Font)

    model_config = {"populate_by_name": True}


class ShapeNode(BaseModel):
    type: str = Field("Shape", alias="_type")
    shape_type: int | None = None
    name: str = ""
    width: float | None = None
    height: float | None = None
    is_inline: bool | None = None
    has_image: bool | None = None
    image_data: Optional[ImageData] = None
    text_box: dict[str, Any] | None = None  # textbox paragraph content
    # REMOVED: left, top, wrap_type, wrap_side, relative_horizontal_position,
    #          relative_vertical_position, horizontal_alignment,
    #          vertical_alignment, rotation, z_order, behind_text

    model_config = {"populate_by_name": True}


class FieldStart(BaseModel):
    type: str = Field("FieldStart", alias="_type")
    field_type: int | None = None

    model_config = {"populate_by_name": True}


# REMOVED entirely: BookmarkStart, BookmarkEnd, CommentNode, FootnoteNode

InlineExtra = Union[ShapeNode, FieldStart]

_KEPT_INLINE_TYPES = {"Shape", "FieldStart"}


# ─────────────────────────────────────────────
# Paragraph
# ─────────────────────────────────────────────


class Paragraph(BaseModel):
    type: Literal["Paragraph"] = Field("Paragraph", alias="_type")
    paragraph_format: ParagraphFormat = Field(default_factory=ParagraphFormat)
    list_format: ListFormat | None = None
    runs: list[Run] = Field(default_factory=list)
    inline_extras: list[InlineExtra] = Field(default_factory=list)
    # Unified ordered sequence of Run and ShapeNode in XML element order.
    # Excluded from serialization; populated by the reader when images are present.
    content_sequence: list[Union["Run", "ShapeNode"]] = Field(default_factory=list, exclude=True)
    text: str = Field("", alias="_text")

    model_config = {"populate_by_name": True}

    @field_validator("inline_extras", mode="before")
    @classmethod
    def _keep_renderable_extras(cls, v: Any) -> list:
        """Filter out non-renderable inline nodes (bookmarks, comments, etc.)."""
        if not isinstance(v, list):
            return v
        return [
            item
            for item in v
            if (isinstance(item, dict) and item.get("_type") in _KEPT_INLINE_TYPES)
            or isinstance(item, (ShapeNode, FieldStart))
        ]


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
    # REMOVED: orientation, wrap_text, fit_text


class Cell(BaseModel):
    type: str = Field("Cell", alias="_type")
    cell_format: CellFormat = Field(default_factory=CellFormat)
    paragraphs: list[Paragraph] = Field(default_factory=list)
    tables: list[Table] = Field(default_factory=list)

    model_config = {"populate_by_name": True}


class RowFormat(BaseModel):
    height: float = 0.0
    height_rule: int = 0
    heading_format: bool = False
    allow_break_across_pages: bool = True  # row split control
    # REMOVED: borders


class Row(BaseModel):
    type: str = Field("Row", alias="_type")
    row_format: RowFormat = Field(default_factory=RowFormat)
    cells: list[Cell] = Field(default_factory=list)

    model_config = {"populate_by_name": True}


class Table(BaseModel):
    type: Literal["Table"] = Field("Table", alias="_type")
    alignment: int = 0
    preferred_width: str = ""  # table-level width hint
    left_indent: float = 0.0
    left_padding: float = 0.0
    right_padding: float = 0.0
    top_padding: float = 0.0
    bottom_padding: float = 0.0
    rows: list[Row] = Field(default_factory=list)
    # REMOVED: style_name, style_identifier, bidi, allow_auto_fit,
    #          allow_cell_spacing, cell_spacing, text_wrapping

    model_config = {"populate_by_name": True}


Cell.model_rebuild()


# ─────────────────────────────────────────────
# Body children
# ─────────────────────────────────────────────


class UnknownNode(BaseModel):
    type: str = Field("", alias="_type")
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
    page_width: float = 0.0
    page_height: float = 0.0
    page_number_style: int = 0  # Roman, Arabic, letters...
    page_starting_number: int = 1  # section starts at page N
    restart_page_numbering: bool = False  # reset counter at section
    section_start: int = 0
    different_first_page_header_footer: bool = False
    odd_and_even_pages_header_footer: bool = False  # enable odd/even headers
    # REMOVED: gutter, multiple_pages, sheets_per_booklet,
    #          vertical_alignment, line_number_*, bidi, borders,
    #          text_columns


class HeaderFooter(BaseModel):
    type: str = Field("", alias="_type")
    header_footer_type: int | None = None
    paragraphs: list[Paragraph] = Field(default_factory=list)
    tables: list[Table] = Field(default_factory=list)

    model_config = {"populate_by_name": True}


class Body(BaseModel):
    type: str = Field("Body", alias="_type")
    children: list[BodyChild] = Field(default_factory=list)

    model_config = {"populate_by_name": True}


class Section(BaseModel):
    type: str = Field("Section", alias="_type")
    page_setup: PageSetup = Field(default_factory=PageSetup)
    body: Body = Field(default_factory=Body)
    headers_footers: list[HeaderFooter] = Field(default_factory=list)

    model_config = {"populate_by_name": True}


# ─────────────────────────────────────────────
# Styles
# ─────────────────────────────────────────────


class Style(BaseModel):
    name: str = ""
    type: int = 0  # 1=paragraph, 2=character, 3=table
    is_heading: bool = False
    base_style_name: str = ""
    next_paragraph_style_name: str = ""
    paragraph_format: ParagraphFormat | None = None
    font: Font | None = None
    # REMOVED: style_identifier, built_in, is_quick_style,
    #          linked_style_name, aliases, automatically_update, priority


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
    # REMOVED: tab_position, restart_after_level, trailing_character, is_legal


class DocList(BaseModel):
    list_id: int = 0
    is_multi_level: bool = False
    levels: list[ListLevel] = Field(default_factory=list)
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
    type: str = Field("Document", alias="_type")
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


