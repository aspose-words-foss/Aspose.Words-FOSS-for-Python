"""
Data classes for DOCX document elements.

These lightweight dataclasses represent parsed DOCX content
(paragraphs, runs, tables, numbering) in a format-agnostic way,
decoupled from XML parsing and LDM construction.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional


@dataclass
class RunData:
    """Text run with formatting."""

    text: str = ""
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strikethrough: bool = False
    style_name: str = ""
    is_code_style: bool = False


@dataclass
class ParagraphData:
    """Paragraph with style and content."""

    text: str = ""
    style_name: str = "Normal"
    runs: list[RunData] = field(default_factory=list)
    is_list_item: bool = False
    list_level: int = 0
    list_id: Optional[int] = None
    alignment: str = "left"
    has_bottom_border: bool = False
    border_size: int = 0


@dataclass
class CellData:
    """Table cell."""

    paragraphs: list[ParagraphData] = field(default_factory=list)
    alignment: str = "left"


@dataclass
class RowData:
    """Table row."""

    cells: list[CellData] = field(default_factory=list)


@dataclass
class TableData:
    """Table structure."""

    rows: list[RowData] = field(default_factory=list)


@dataclass
class NumberingLevel:
    """List level definition."""

    format: str = "bullet"
    start: int = 1
    text: str = ""


@dataclass
class NumberingInfo:
    """Numbering definition."""

    num_id: int
    abstract_num_id: int
    levels: dict[int, NumberingLevel] = field(default_factory=dict)
