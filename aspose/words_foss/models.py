"""Data models for DOCX to Markdown conversion."""

from dataclasses import dataclass, field
from enum import Enum
from typing import Optional


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
    table_content_alignment: str = "auto"
    list_export_mode: str = "markdown_syntax"
    link_export_mode: str = "auto"
    export_as_html: str = "none"
    empty_paragraph_export_mode: str = "empty_line"
    export_images_as_base64: bool = False
    images_folder: str = ""
    images_folder_alias: str = ""
    paragraph_break: str = "\n"


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
