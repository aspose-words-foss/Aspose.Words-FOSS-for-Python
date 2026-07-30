"""
Internal Markdown block model.

Mirrors the block tree Aspose.Words builds from Markdig before feeding it
to ``MarkdownReaderContext``: every node is a ``Block`` with a ``type``
(``BlockType``), a parent, and ordered children, plus the sibling/ancestor
navigation (``is_first_child``, ``next_sibling``, ``get_parent`` ...)
``MarkdownReaderContext``-style consumers rely on to decide when to reuse
vs. create a paragraph style. ``block_parser.parse_document`` builds this
tree from raw Markdown text; ``MarkdownReaderContext`` (a later step) will
walk it to populate the light document model.
"""

from __future__ import annotations

from enum import Enum, auto
from typing import Optional


class MarkdownBlockLevel(Enum):
    """Where a block sits in the tree.

    ``DOCUMENT`` is the root; ``BLOCK`` is a container that only holds
    other blocks (list, list item, quote, table, row, cell); ``LEAF``
    holds inline content directly (paragraph, heading, code block); and
    ``INLINE`` is a formatting span or text run inside a leaf.
    """

    DOCUMENT = auto()
    BLOCK = auto()
    LEAF = auto()
    INLINE = auto()


class BlockType(Enum):
    DOCUMENT = auto()
    PARAGRAPH = auto()
    ATX_HEADING = auto()
    SETEXT_HEADING = auto()
    QUOTE = auto()
    INDENTED_CODE = auto()
    FENCED_CODE = auto()
    HORIZONTAL_RULE = auto()
    LIST = auto()
    BULLET_LIST_ITEM = auto()
    ORDERED_LIST_ITEM = auto()
    TABLE = auto()
    ROW = auto()
    CELL = auto()
    HTML_TAG = auto()
    FOOTNOTE_DEFINITION = auto()
    TEXT = auto()
    LINE_BREAK = auto()
    LINK_TEXT = auto()
    AUTOLINK = auto()
    BOLD_INLINE = auto()
    ITALIC_INLINE = auto()
    STRIKETHROUGH = auto()
    UNDERLINE = auto()
    INLINE_CODE = auto()
    FOOTNOTE_REFERENCE = auto()


class ListMarker(Enum):
    """Bullet/ordered list marker, mirroring Aspose.Words' ``ListMarker``."""

    NONE = auto()
    DASH = auto()
    STAR = auto()
    PLUS = auto()
    DOT = auto()
    PARENTHESIS = auto()


class Block:
    """Base class for every node in the Markdown block tree."""

    block_level: MarkdownBlockLevel = MarkdownBlockLevel.BLOCK

    def __init__(self, block_type: BlockType) -> None:
        self.type = block_type
        self.parent: Optional["Block"] = None
        self.children: list["Block"] = []

    def add_child(self, child: "Block") -> "Block":
        child.parent = self
        self.children.append(child)
        return child

    @property
    def is_first_child(self) -> bool:
        return self.parent is not None and bool(self.parent.children) and self.parent.children[0] is self

    @property
    def is_last_child(self) -> bool:
        return self.parent is not None and bool(self.parent.children) and self.parent.children[-1] is self

    @property
    def next_sibling(self) -> Optional["Block"]:
        if self.parent is None:
            return None
        siblings = self.parent.children
        idx = siblings.index(self)
        return siblings[idx + 1] if idx + 1 < len(siblings) else None

    @property
    def previous_sibling(self) -> Optional["Block"]:
        if self.parent is None:
            return None
        siblings = self.parent.children
        idx = siblings.index(self)
        return siblings[idx - 1] if idx > 0 else None

    def get_parent(self, *types: BlockType) -> Optional["Block"]:
        """Walk up the ancestor chain and return the first block whose
        ``type`` is one of *types*, or ``None`` if none matches."""
        node = self.parent
        while node is not None:
            if node.type in types:
                return node
            node = node.parent
        return None

    def __repr__(self) -> str:  # pragma: no cover - debugging aid
        return f"<{type(self).__name__} {self.type.name}>"


class DocumentBlock(Block):
    block_level = MarkdownBlockLevel.DOCUMENT

    def __init__(self) -> None:
        super().__init__(BlockType.DOCUMENT)


class TextBlock(Block):
    """A literal run of text (Markdig's ``LiteralInline``)."""

    block_level = MarkdownBlockLevel.INLINE

    def __init__(self, text: str) -> None:
        super().__init__(BlockType.TEXT)
        self.text = text


class LineBreakBlock(Block):
    block_level = MarkdownBlockLevel.INLINE

    def __init__(self, is_hard: bool = True) -> None:
        super().__init__(BlockType.LINE_BREAK)
        self.is_hard = is_hard


class ParagraphBlock(Block):
    block_level = MarkdownBlockLevel.LEAF

    def __init__(self) -> None:
        super().__init__(BlockType.PARAGRAPH)


class HeadingBlock(Block):
    block_level = MarkdownBlockLevel.LEAF

    def __init__(self, block_type: BlockType, level: int) -> None:
        super().__init__(block_type)
        self.level = level


class AtxHeadingBlock(HeadingBlock):
    def __init__(self, level: int) -> None:
        super().__init__(BlockType.ATX_HEADING, level)


class SetextHeadingBlock(HeadingBlock):
    def __init__(self, level: int) -> None:
        super().__init__(BlockType.SETEXT_HEADING, level)


class QuoteBlock(Block):
    block_level = MarkdownBlockLevel.BLOCK

    def __init__(self) -> None:
        super().__init__(BlockType.QUOTE)


class IndentedCodeBlock(Block):
    block_level = MarkdownBlockLevel.LEAF

    def __init__(self, code: str) -> None:
        super().__init__(BlockType.INDENTED_CODE)
        self.code = code


class FencedCodeBlock(Block):
    block_level = MarkdownBlockLevel.LEAF

    def __init__(self, code: str, info: str = "", fence_char: str = "`") -> None:
        super().__init__(BlockType.FENCED_CODE)
        self.code = code
        self.info = info
        self.fence_char = fence_char


class HorizontalRuleBlock(Block):
    block_level = MarkdownBlockLevel.LEAF

    def __init__(self) -> None:
        super().__init__(BlockType.HORIZONTAL_RULE)


class ListBlock(Block):
    """Container grouping sibling ``ListItemBlock``s (Markdig's ``ListBlock``)."""

    block_level = MarkdownBlockLevel.BLOCK

    def __init__(self, ordered: bool, marker: ListMarker) -> None:
        super().__init__(BlockType.LIST)
        self.ordered = ordered
        self.marker = marker


class ListItemBlock(Block):
    block_level = MarkdownBlockLevel.BLOCK

    def __init__(self, block_type: BlockType, marker: ListMarker, start_at: int = 1) -> None:
        super().__init__(block_type)
        self.marker = marker
        self.start_at = start_at
        self.is_level_start = False

    def get_level(self) -> int:
        """0-based nesting depth: how many ancestor list items enclose this one."""
        level = 0
        node = self.parent
        while node is not None:
            if node.type in (BlockType.BULLET_LIST_ITEM, BlockType.ORDERED_LIST_ITEM):
                level += 1
            node = node.parent
        return level

    def get_list_container(self) -> Optional["ListBlock"]:
        container = self.get_parent(BlockType.LIST)
        return container  # type: ignore[return-value]


class BulletListItemBlock(ListItemBlock):
    def __init__(self, marker: ListMarker) -> None:
        super().__init__(BlockType.BULLET_LIST_ITEM, marker)


class OrderedListItemBlock(ListItemBlock):
    def __init__(self, start_at: int = 1, marker: ListMarker = ListMarker.DOT) -> None:
        super().__init__(BlockType.ORDERED_LIST_ITEM, marker, start_at)


class TableBlock(Block):
    block_level = MarkdownBlockLevel.BLOCK

    def __init__(self, column_alignments: list[Optional[str]]) -> None:
        super().__init__(BlockType.TABLE)
        self.column_alignments = column_alignments

    @property
    def columns_count(self) -> int:
        return len(self.column_alignments)


class RowBlock(Block):
    block_level = MarkdownBlockLevel.BLOCK

    def __init__(self, is_header: bool = False) -> None:
        super().__init__(BlockType.ROW)
        self.is_header = is_header

    @property
    def count(self) -> int:
        return len(self.children)


class CellBlock(Block):
    block_level = MarkdownBlockLevel.BLOCK

    def __init__(self, alignment: Optional[str] = None) -> None:
        super().__init__(BlockType.CELL)
        self.alignment = alignment


KNOWN_HTML_TAG_NAMES = {
    "p", "h1", "h2", "h3", "h4", "h5", "h6", "i", "b", "u", "sup", "sub",
    "strong", "strike", "del", "s", "br", "table", "tr", "td", "code",
    "html", "head", "meta", "title", "div", "span", "body",
}

# CommonMark's HTML-block type-6/type-1 tag whitelist: only these (plus script/pre/style/
# textarea) can start a raw HTML block; inline-only tags like <del>/<b>/<code> never do.
HTML_BLOCK_TAG_NAMES = {
    "address", "article", "aside", "base", "basefont", "blockquote", "body",
    "caption", "center", "col", "colgroup", "dd", "details", "dialog",
    "dir", "div", "dl", "dt", "fieldset", "figcaption", "figure",
    "footer", "form", "frame", "frameset",
    "h1", "h2", "h3", "h4", "h5", "h6", "head", "header", "hr",
    "html", "iframe", "legend", "li", "link", "main", "menu", "menuitem",
    "nav", "noframes", "ol", "optgroup", "option", "p", "param",
    "search", "section", "summary", "table", "tbody", "td",
    "tfoot", "th", "thead", "title", "tr", "track", "ul",
    "script", "pre", "style", "textarea",
}

_PARA_OR_HEADING_TAGS = {"p", "h1", "h2", "h3", "h4", "h5", "h6"}


class HtmlTagBlock(Block):
    """A raw HTML tag, block-level or inline (Markdig's ``HtmlBlock`` /
    ``HtmlInline``, folded into one node here)."""

    block_level = MarkdownBlockLevel.INLINE

    def __init__(
        self,
        tag_name: str,
        raw_text: str,
        is_self_closing: bool = False,
        is_closing: bool = False,
        is_self_contained: bool = False,
    ) -> None:
        super().__init__(BlockType.HTML_TAG)
        self.tag_name = tag_name.lower()
        self.raw_text = raw_text
        self.is_self_closing = is_self_closing
        self.is_closing = is_closing
        self.is_self_contained = is_self_contained  # whole "<tag>...</tag>" pair vs. a flat open/close event

    @property
    def is_known(self) -> bool:
        return self.tag_name in KNOWN_HTML_TAG_NAMES

    @property
    def is_para_or_heading(self) -> bool:
        return self.tag_name in _PARA_OR_HEADING_TAGS

    @property
    def is_table(self) -> bool:
        return self.tag_name == "table"

    @property
    def is_cell(self) -> bool:
        return self.tag_name == "td"

    @property
    def heading_tag_level(self) -> int:
        if len(self.tag_name) == 2 and self.tag_name[0] == "h" and self.tag_name[1].isdigit():
            return int(self.tag_name[1])
        return -1


class LinkTextBlock(Block):
    block_level = MarkdownBlockLevel.INLINE

    def __init__(
        self,
        uri: Optional[str] = None,
        title: str = "",
        is_image: bool = False,
        definition_label: Optional[str] = None,
        raw_text: str = "",
    ) -> None:
        super().__init__(BlockType.LINK_TEXT)
        self.uri = uri  # None = unresolved; "" = explicit-but-empty "()", a real link
        self.title = title
        self.is_image = is_image
        self.definition_label = definition_label
        self.raw_text = raw_text  # literal fallback if never resolved, see _resolve_references

    @property
    def has_destination(self) -> bool:
        return self.uri is not None or self.definition_label is not None


class AutolinkBlock(Block):
    block_level = MarkdownBlockLevel.INLINE

    def __init__(self, text: str, uri: str) -> None:
        super().__init__(BlockType.AUTOLINK)
        self.text = text  # visible label, e.g. "email@e.com" (no mailto:)
        self.uri = uri  # hyperlink target, e.g. "mailto:email@e.com"


class BoldInlineBlock(Block):
    block_level = MarkdownBlockLevel.INLINE

    def __init__(self) -> None:
        super().__init__(BlockType.BOLD_INLINE)


class ItalicInlineBlock(Block):
    block_level = MarkdownBlockLevel.INLINE

    def __init__(self) -> None:
        super().__init__(BlockType.ITALIC_INLINE)


class StrikethroughBlock(Block):
    block_level = MarkdownBlockLevel.INLINE

    def __init__(self) -> None:
        super().__init__(BlockType.STRIKETHROUGH)


class UnderlineBlock(Block):
    block_level = MarkdownBlockLevel.INLINE

    def __init__(self) -> None:
        super().__init__(BlockType.UNDERLINE)


class InlineCodeBlock(Block):
    block_level = MarkdownBlockLevel.INLINE

    def __init__(self, code: str, delimiter_length: int = 1) -> None:
        super().__init__(BlockType.INLINE_CODE)
        self.code = code
        self.delimiter_length = delimiter_length


class FootnoteReferenceBlock(Block):
    block_level = MarkdownBlockLevel.INLINE

    def __init__(self, label: str) -> None:
        super().__init__(BlockType.FOOTNOTE_REFERENCE)
        self.label = label


class FootnoteDefinitionBlock(Block):
    block_level = MarkdownBlockLevel.BLOCK

    def __init__(self, label: str) -> None:
        super().__init__(BlockType.FOOTNOTE_DEFINITION)
        self.label = label
