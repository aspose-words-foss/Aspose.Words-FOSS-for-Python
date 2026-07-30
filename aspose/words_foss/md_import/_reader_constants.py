"""Shared constants for the ``MarkdownReaderContext`` mixins.

Style-name strings mirror the reference's ``MarkdownUtil.*StyleName``
constants (whose literal values aren't part of the task's C# excerpt, so
these are chosen to (a) read clearly and (b) match ``md_writer.py``'s
export-side name heuristics — see the comment on
``_INDENTED_CODE_STYLE_NAME`` below).
"""

import re

from aspose.words_foss.model.enums import StyleIdentifier

from .blocks import BlockType, ListMarker

_NORMAL_STYLE_NAME = "Normal"
_HEADING_STYLE_NAME = "Heading"
_SETEXT_HEADING_STYLE_NAME = "SetextHeading"
_QUOTE_STYLE_NAME = "Quote"
_LIST_STYLE_NAME = "List"
_FOOTNOTE_STYLE_NAME = "Footnote Text"
_FOOTNOTE_TEXT_STYLE_NAME = "Footnote Text"
_FOOTNOTE_REFERENCE_STYLE_NAME = "Footnote Reference"
# FootnoteReferenceBlock.OpeningDelimiter/ClosingDelimiter: the mark a
# footnote reference and its definition share.
_FOOTNOTE_REFERENCE_OPENING = "[^"
_FOOTNOTE_REFERENCE_CLOSING = "]"
_HYPERLINK_STYLE_NAME = "Hyperlink"
_INDENTED_CODE_STYLE_NAME = "IndentedCode"
_FENCED_CODE_STYLE_NAME = "FencedCode"
# The delimiter a fenced block was written with, kept in the style name
# next to its info string -- the only place the model can carry it.
_TILDE_FENCE_MARKER = "~"
_INLINE_CODE_STYLE_NAME = "InlineCode"

# Style-name -> StyleIdentifier overrides for names the shared
# ``model.style_identifiers.resolve_style_identifier`` table can't
# resolve on its own. A Setext-sourced heading is semantically a
# "Heading N" (only its *name* differs, to preserve which Markdown
# syntax produced it) — reuse the matching Heading identifier so
# _create_style's "style_identifier != NORMAL" base-style-linking
# check fires. Every other built-in name (Normal, Heading N, Quote,
# List Paragraph, Hyperlink, Footnote Text/Reference, …) resolves
# through the shared table instead, so it isn't duplicated here.
_SPECIAL_STYLE_IDENTIFIERS = {
    f"{_SETEXT_HEADING_STYLE_NAME}1": StyleIdentifier.HEADING_1,
    f"{_SETEXT_HEADING_STYLE_NAME}2": StyleIdentifier.HEADING_2,
}

_UNIQUE_STYLE_BASE_NAMES = {
    BlockType.QUOTE: _QUOTE_STYLE_NAME,
    BlockType.BULLET_LIST_ITEM: _LIST_STYLE_NAME,
    BlockType.ORDERED_LIST_ITEM: _LIST_STYLE_NAME,
    BlockType.FOOTNOTE_DEFINITION: _FOOTNOTE_STYLE_NAME,
}

_MARKER_CHAR = {
    ListMarker.DASH: "-",
    ListMarker.STAR: "*",
    ListMarker.PLUS: "+",
    ListMarker.DOT: ".",
    ListMarker.PARENTHESIS: ")",
    ListMarker.NONE: "",
}
_CHAR_TO_MARKER = {v: k for k, v in _MARKER_CHAR.items() if v}

_INLINE_HTML_HEADING_TAGS = frozenset({"p", "h1", "h2", "h3", "h4", "h5", "h6"})

_LEAF_CLOSE_TYPES = frozenset({
    BlockType.PARAGRAPH, BlockType.HORIZONTAL_RULE, BlockType.ATX_HEADING,
    BlockType.SETEXT_HEADING, BlockType.INDENTED_CODE, BlockType.FENCED_CODE,
})

_FOOTNOTE_PARAGRAPH_TYPES = frozenset({
    BlockType.PARAGRAPH, BlockType.INDENTED_CODE, BlockType.HORIZONTAL_RULE,
    BlockType.FENCED_CODE, BlockType.SETEXT_HEADING, BlockType.ATX_HEADING,
})

# Matches a single node holding a full open+body+close HTML tag pair.
_SELF_CONTAINED_HTML_RE = re.compile(r"^<([a-zA-Z][\w-]*)[^>]*>(.*)</\1>$", re.DOTALL | re.IGNORECASE)
