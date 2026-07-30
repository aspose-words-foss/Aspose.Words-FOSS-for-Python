"""
Markdown block-model import infrastructure.

``blocks`` defines the internal ``Block`` / ``BlockType`` tree that
mirrors Aspose.Words' Markdig-based block model; ``block_parser`` turns
raw Markdown text into that tree; ``inline_parser`` handles the inline
(emphasis, links, code spans, ...) content of leaf blocks; ``document_builder``
and ``reader_context`` implement the ``MarkdownReaderContext``-equivalent
that walks the block tree and populates the light document model;
``traversal`` drives that walk end-to-end (``parse_and_build``), which
``markdown_reader.MarkdownReader.to_light_document`` calls for real ``.md``
imports.
"""

from .block_parser import is_block_start, is_ordered_list_item, parse_document
from .inline_parser import is_valid_autolink
from .blocks import (
    AtxHeadingBlock,
    AutolinkBlock,
    Block,
    BlockType,
    BoldInlineBlock,
    BulletListItemBlock,
    CellBlock,
    DocumentBlock,
    FencedCodeBlock,
    FootnoteDefinitionBlock,
    FootnoteReferenceBlock,
    HeadingBlock,
    HorizontalRuleBlock,
    HtmlTagBlock,
    IndentedCodeBlock,
    InlineCodeBlock,
    ItalicInlineBlock,
    LineBreakBlock,
    LinkTextBlock,
    ListBlock,
    ListItemBlock,
    ListMarker,
    MarkdownBlockLevel,
    OrderedListItemBlock,
    ParagraphBlock,
    QuoteBlock,
    RowBlock,
    SetextHeadingBlock,
    StrikethroughBlock,
    TableBlock,
    TextBlock,
    UnderlineBlock,
)
from .document_builder import HtmlInsertOptions, MarkdownDocumentBuilder
from .inline_parser import parse_inline
from .reader_context import MarkdownReaderContext
from .traversal import build_document, drive_block, parse_and_build

__all__ = [
    "is_block_start",
    "is_valid_autolink",
    "is_ordered_list_item",
    "parse_document",
    "parse_inline",
    "MarkdownReaderContext",
    "MarkdownDocumentBuilder",
    "HtmlInsertOptions",
    "build_document",
    "drive_block",
    "parse_and_build",
    "Block",
    "BlockType",
    "MarkdownBlockLevel",
    "ListMarker",
    "DocumentBlock",
    "TextBlock",
    "LineBreakBlock",
    "ParagraphBlock",
    "HeadingBlock",
    "AtxHeadingBlock",
    "SetextHeadingBlock",
    "QuoteBlock",
    "IndentedCodeBlock",
    "FencedCodeBlock",
    "HorizontalRuleBlock",
    "ListBlock",
    "ListItemBlock",
    "BulletListItemBlock",
    "OrderedListItemBlock",
    "TableBlock",
    "RowBlock",
    "CellBlock",
    "HtmlTagBlock",
    "LinkTextBlock",
    "AutolinkBlock",
    "BoldInlineBlock",
    "ItalicInlineBlock",
    "StrikethroughBlock",
    "UnderlineBlock",
    "InlineCodeBlock",
    "FootnoteReferenceBlock",
    "FootnoteDefinitionBlock",
    "is_block_start",
    "is_valid_autolink",
    "is_ordered_list_item",
]
