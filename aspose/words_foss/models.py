"""Data models for DOCX to Markdown conversion."""

from dataclasses import dataclass, field
from enum import Enum
from typing import Optional

from aspose.words_foss.saving import (
    MarkdownEmptyParagraphExportMode,
    MarkdownExportAsHtml,
    MarkdownLinkExportMode,
    MarkdownListExportMode,
    TableContentAlignment,
    coerce_enum,
)


class HeadingStyle(Enum):
    """Heading export style preference."""

    ATX = "atx"  # # Heading
    SETEXT = "setext"  # Heading\n=======


class ListMarker(Enum):
    """Bullet list marker style."""

    DASH = "-"
    ASTERISK = "*"
    PLUS = "+"


class CodeBlockStyle(Enum):
    """Code block style preference."""

    FENCED = "fenced"  # ```code```
    INDENTED = "indented"  # 4-space indent


@dataclass
class ConversionOptions:
    """Options for controlling DOCX to Markdown conversion."""

    heading_style: HeadingStyle = HeadingStyle.ATX
    list_marker: ListMarker = ListMarker.DASH
    code_block_style: CodeBlockStyle = CodeBlockStyle.FENCED
    export_underline: bool = False
    export_strikethrough: bool = True
    export_headers_footers: bool = False
    preserve_emphasis: bool = True
    table_pipe_style: bool = True
    wrap_width: Optional[int] = None
    escape_special_chars: bool = True
    table_content_alignment: TableContentAlignment = TableContentAlignment.AUTO
    list_export_mode: MarkdownListExportMode = MarkdownListExportMode.MARKDOWN_SYNTAX
    link_export_mode: MarkdownLinkExportMode = MarkdownLinkExportMode.AUTO
    export_as_html: MarkdownExportAsHtml = MarkdownExportAsHtml.NONE
    empty_paragraph_export_mode: MarkdownEmptyParagraphExportMode = (
        MarkdownEmptyParagraphExportMode.EMPTY_LINE
    )
    export_images_as_base64: bool = False
    images_folder: str = ""
    images_folder_alias: str = ""
    paragraph_break: str = "\n"

    def __post_init__(self) -> None:
        # Accept the legacy lowercase string spellings as well as members.
        for name, enum_cls in (
            ("table_content_alignment", TableContentAlignment),
            ("list_export_mode", MarkdownListExportMode),
            ("link_export_mode", MarkdownLinkExportMode),
            ("export_as_html", MarkdownExportAsHtml),
            ("empty_paragraph_export_mode", MarkdownEmptyParagraphExportMode),
        ):
            setattr(self, name, coerce_enum(enum_cls, getattr(self, name)))


@dataclass
class RunFormatting:
    """Text run formatting properties."""

    bold: bool = False
    italic: bool = False
    underline: bool = False
    strikethrough: bool = False
    code: bool = False
    superscript: bool = False
    subscript: bool = False


@dataclass
class ParagraphInfo:
    """Information about a paragraph's style and context."""

    style_name: str = "Normal"
    heading_level: int = 0
    is_quote: bool = False
    quote_level: int = 0
    is_list_item: bool = False
    list_level: int = 0
    list_type: str = ""  # "bullet" or "ordered"
    list_marker: str = ""
    is_code_block: bool = False
    code_language: str = ""
    alignment: str = "left"


@dataclass
class TableCell:
    """Represents a table cell."""

    text: str
    alignment: str = "left"
    row_span: int = 1
    col_span: int = 1
    formatting: RunFormatting = field(default_factory=RunFormatting)


@dataclass
class TableRow:
    """Represents a table row."""

    cells: list[TableCell] = field(default_factory=list)
    is_header: bool = False


@dataclass
class Table:
    """Represents a table structure."""

    rows: list[TableRow] = field(default_factory=list)
    column_alignments: list[str] = field(default_factory=list)
