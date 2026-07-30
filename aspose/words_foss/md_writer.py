"""Markdown writer operating on the light document model.

Converts an ldm.Document into a Markdown string.  All formatting
decisions (headings, lists, tables, inline emphasis) are driven
exclusively by the Pydantic model classes defined in
``light_document_model.py``.
"""

import base64
import html
import os
import re
from pathlib import Path
from typing import NamedTuple, Optional

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss._visible_runs import is_horizontal_rule_shape, visible_runs
from aspose.words_foss.md_import.document_builder import (
    _link_destination,
    _link_title,
    _unescape_markup_symbols,
)
from aspose.words_foss.md_import import (
    is_block_start,
    is_ordered_list_item,
    is_valid_autolink,
)
from aspose.words_foss.model.enums import NumberStyle
from aspose.words_foss.models import (
    CodeBlockStyle,
    ConversionOptions,
    HeadingStyle,
    RunFormatting,
)
from aspose.words_foss.saving import (
    MarkdownEmptyParagraphExportMode,
    MarkdownExportAsHtml,
    MarkdownLinkExportMode,
    MarkdownListExportMode,
    TableContentAlignment,
)

# Alignment int -> string (matches light_document_model conventions)
_ALIGN_STR = {0: "left", 1: "center", 2: "right", 3: "left"}

# Style-name prefixes the reader assigns (see md_import/_reader_constants.py).
_QUOTE_STYLE_NAME = "Quote"
_INDENTED_CODE_STYLE_NAME = "IndentedCode"
_FENCED_CODE_STYLE_NAME = "FencedCode"
_INLINE_CODE_STYLE_NAME = "InlineCode"
_CODE_STYLE_NAMES = (_INDENTED_CODE_STYLE_NAME, _FENCED_CODE_STYLE_NAME, "Code")
_FOOTNOTE_STYLE_NAME = "Footnote"
_FOOTNOTE_REFERENCE_OPENING = "[^"
_FOOTNOTE_REFERENCE_STYLE_NAME = "Footnote Reference"
_HYPERLINK_STYLE_NAME = "Hyperlink"
_FOOTNOTE_MARK_RE = re.compile(r"(?<!\\)\[\^[^\]\s]+\]")
# Block.MaxIndentationLength + 1 -- what a footnote's later lines are indented by.
_FOOTNOTE_CONTINUATION_INDENT = 4
_HEADING_STYLE_NAME = "Heading"
_SETEXT_HEADING_STYLE_NAME = "SetextHeading"
# IndentedCodeBlock.OpeningLength: Block.MaxIndentationLength + 1.
_INDENTED_CODE_INDENT = 4

# A fenced block keeps the delimiter it was written with; the style name is
# where the reader put it, next to the info string.
_BACKTICK_FENCE_CHAR = "`"
_TILDE_FENCE_CHAR = "~"
_FENCED_CODE_DELIMITER_LENGTH = 3
_CLOSING_FENCE_RES = {
    char: re.compile(rf"^ {{0,3}}({re.escape(char)}+)[ \t]*$")
    for char in (_BACKTICK_FENCE_CHAR, _TILDE_FENCE_CHAR)
}
_BACKTICK_RUN_RE = re.compile(r"`+")
# MarkdownUtil.gEscapableMarkupCharacters, minus intraword "_" (never a delimiter)
# and minus what already carries a backslash (escaping twice grows without end).
# A "<" that opens what would read back as an autolink joins them: the run text
# holds it as text, because the reader turns real autolinks into "[text](uri)".
_ESCAPABLE_MARKUP_RE = re.compile(
    r"(?<!\\)(?:\*|(?<![^\W_])_|_(?![^\W_])|<(?=[A-Za-z][A-Za-z0-9+.-]{1,31}:))"
)
# MarkdownWriter.HorizontalRule, at the shortest length that is one
_HORIZONTAL_RULE = "---"
_TRAILING_HASHES_RE = re.compile(r"(?:^|[ \t])#+$")
# A newline already spelled with a backslash keeps the one it has.
_HARD_BREAK_RE = re.compile(r"(?<!\\)\n")
# MarkdownEmphasesWriter's writers, each with the delimiter it opens on.
# Markdown has no syntax of its own for the last two, so they stay HTML.
_BOLD_MARKER = "**"
_ITALIC_MARKER = "*"
_STRIKETHROUGH_MARKER = "~~"
_UNDERLINE_MARKER = "++"
_SUBSCRIPT_MARKER = "<sub>"
_SUPERSCRIPT_MARKER = "<sup>"
_CLOSING_MARKERS = {_SUBSCRIPT_MARKER: "</sub>", _SUPERSCRIPT_MARKER: "</sup>"}
# BoldInlineBlock/ItalicInlineBlock.UnderscoreDelimiter: what
# SwitchToUnderscore writes instead when two delimiters would collide.
_UNDERSCORE_MARKERS = {_BOLD_MARKER: "__", _ITALIC_MARKER: "_"}
# The ASCII punctuation a backslash escapes, per the spec.
_ESCAPABLE_PUNCTUATION = set("!\"#$%&'()*+,-./:;<=>?@[\\]^_`{|}~")
# AutolinkInlineBlock.OpeningDelimiter/ClosingDelimiter, and the scheme
# MarkdownHyperlinkWriter lets an autolink target carry on its own.
_AUTOLINK_OPENING = "<"
_AUTOLINK_CLOSING = ">"
_MAILTO_PREFIX = "mailto:"
_HYPERLINK_FIELD_NAME = "HYPERLINK"
_FIELD_SWITCH_LOCATION = "l"
_FIELD_SWITCH_SCREEN_TIP = "o"
_FIELD_ARGUMENT_RE = re.compile(r'\\+([a-zA-Z])|"([^"]*)"|(\S+)')
_ORDERED_MARKER_CHARS = (".", ")")
_BULLET_MARKER_CHARS = ("-", "*", "+")


class _BlockTag(NamedTuple):
    """What ``_join_blocks`` needs to know about a block to separate it.

    Mirrors the questions ``MarkdownParagraphWriter.IsNeedParagraphBreak`` asks
    of the previous paragraph.
    """

    kind: str
    quote_level: int = 0
    is_code: bool = False
    list_level: int = -1
    ordered_start: int = 0  # 0 when the item is not numbered


def _get_number_after_substring(text: str, substring: str) -> int:
    """Port of ``MarkdownUtil.GetNumberAfterSubstring``: the number a style
    name carries after *substring* (one dot may separate the two)."""
    start = text.lower().find(substring.lower())
    if start == -1:
        return -1
    start += len(substring)
    if start < len(text) and text[start] == ".":
        start += 1
    end = start
    while end < len(text) and text[end].isdigit():
        end += 1
    return int(text[start:end]) if end > start else 0


class LdmMarkdownWriter:
    """Converts a ``light_document_model.Document`` to a Markdown string."""

    # Block-type tags used by _join_blocks to decide blank-line insertion.
    _LIST = "list"
    _BLOCK = "block"
    _BLANK = "blank"
    _HEADING = "heading"
    _SETEXT = "setext"
    _RULE = "rule"
    _HEADINGS = (_HEADING, _SETEXT)

    def __init__(self, options: Optional[ConversionOptions] = None):
        self.options = options or ConversionOptions()
        self._list_indents: dict[int, int] = {}
        self._doc: Optional[ldm.Document] = None
        self._output_path: Optional[Path] = None
        self._image_counter: int = 0
        self._reference_links: list[tuple[str, str]] = []  # (label, url)
        self._in_footnotes = False
        self._trailing_list_indent = 0

    # ------------------------------------------------------------------
    # Public entry point
    # ------------------------------------------------------------------

    def write(
        self,
        doc: ldm.Document,
        output_path: Optional[Path] = None,
    ) -> str:
        """Convert *doc* to Markdown and return the result string."""
        self._list_indents.clear()
        self._reference_links.clear()
        self._image_counter = 0
        self._doc = doc
        self._output_path = output_path
        blocks: list[tuple[_BlockTag, str]] = []  # (tag, markdown_text)

        # Header images
        for para in doc.header_paragraphs:
            for item in para._children:
                if isinstance(item, ldm.Shape) and item.has_image and item.image_data:
                    blocks.append((_BlockTag(self._BLOCK), self._render_image(item)))

        # Body
        for section in doc.sections:
            for child in section.body.children:
                if isinstance(child, ldm.Paragraph):
                    tag, md = self._convert_paragraph_tagged(child)
                    if md is not None:
                        blocks.append((tag, md))
                elif isinstance(child, ldm.Table):
                    md = self._convert_table(child)
                    blocks.append((_BlockTag(self._BLOCK), md))
                # UnknownNode is silently skipped

        # Footer images
        for para in doc.footer_paragraphs:
            for item in para._children:
                if isinstance(item, ldm.Shape) and item.has_image and item.image_data:
                    blocks.append((_BlockTag(self._BLOCK), self._render_image(item)))

        self._apply_link_export_mode(blocks)
        result = self._join_blocks(blocks)

        # Append reference-style link definitions if any were collected
        if self._reference_links:
            pb = self.options.paragraph_break
            defs = pb.join(
                f"[{label}]: {self._definition_destination(url)}"
                for label, url in self._reference_links
            )
            result = result.rstrip(pb) + pb * 2 + defs + pb

        return result

    # ------------------------------------------------------------------
    # Block joining
    # ------------------------------------------------------------------

    @staticmethod
    def _block_separator(prev_tag: _BlockTag, tag: _BlockTag) -> str:
        """The line that goes between two blocks.

        Two paragraphs of one block quote are separated by a quoted blank line;
        a plain one would end the quote and start a second one instead.
        """
        return ">" * min(prev_tag.quote_level, tag.quote_level)

    def _join_blocks(self, blocks: list[tuple[_BlockTag, str]]) -> str:
        """Join block-level elements with blank lines between them.

        Per the CommonMark spec, consecutive non-blank lines that are not
        otherwise recognised as block-level constructs form a single
        paragraph (soft line breaks become spaces).  Every block element
        must therefore be separated by a blank line so that each one is
        rendered as its own block.

        The one exception is consecutive list items: they must stay on
        adjacent lines to form a *tight* list.
        """
        result: list[str] = []
        prev_tag: Optional[_BlockTag] = None
        for tag, text in blocks:
            if not text.strip():
                # Explicit blank -- collapse consecutive blanks, and drop one
                # that opens the document: nothing reads it back.
                if prev_tag is not None and prev_tag.kind != self._BLANK:
                    result.append("")
                    prev_tag = _BlockTag(self._BLANK)
                continue

            # Decide whether a blank line is needed before this block.
            need_blank = False
            if prev_tag is not None and prev_tag.kind != self._BLANK:
                if tag.kind == self._LIST and prev_tag.kind == self._LIST:
                    # Consecutive list items -> tight list, no blank line.
                    # "When appending an ordered list that does not start at
                    # '1' we need a blank line.  Otherwise, it can break the
                    # paragraph itself": a marker that cannot interrupt a
                    # paragraph reads as more of the item above instead.
                    need_blank = (
                        tag.ordered_start > 1
                        and prev_tag.ordered_start > 0
                        and tag.list_level > prev_tag.list_level
                    )
                elif tag.kind == self._SETEXT:
                    # A setext underline makes a heading of whatever paragraph
                    # precedes it, so it can only join another heading.
                    need_blank = prev_tag.kind not in self._HEADINGS
                elif tag.kind in self._HEADINGS or prev_tag.kind in self._HEADINGS:
                    # "No need to add a blank line before or after Heading
                    # because it is not appendable by spec."
                    need_blank = False
                elif prev_tag.quote_level and tag.quote_level > prev_tag.quote_level:
                    # "No need to separate Quotes when previous quote has less
                    # level" -- when there is a previous quote at all; a quote
                    # left against a table becomes one more row of it.
                    need_blank = False
                elif prev_tag.kind == tag.kind == self._RULE:
                    # "No need to add a blank line after HorizontalRule." Only
                    # before another rule: a paragraph left against a rule that
                    # a second rule follows would read back as a setext heading.
                    need_blank = False
                else:
                    need_blank = True

            if need_blank:
                result.append(self._block_separator(prev_tag, tag))
            result.append(text)
            prev_tag = tag

        pb = self.options.paragraph_break
        output = pb.join(result)
        return output.rstrip() + pb if output else ""

    # ------------------------------------------------------------------
    # Paragraph conversion
    # ------------------------------------------------------------------

    def _render_image(self, shape: ldm.Shape) -> str:
        """Render a Shape with image data as a Markdown inline image tag."""
        img = shape.image_data
        assert img is not None
        if not img.image_bytes:
            # A linked picture has no bytes of its own; its source is the link,
            # and only its alternative text names it.
            source = _link_destination(img.source_full_name)
            return f"![{shape.alternative_text}]({source})"
        alt = shape.alternative_text or img.source_full_name

        # If images_folder is set, save to file instead of base64
        if self.options.images_folder and not self.options.export_images_as_base64:
            self._image_counter += 1
            ext = self._guess_image_extension(img.content_type, img.source_full_name)
            # Sanitize filename to prevent path traversal
            filename = Path(img.source_full_name).name if img.source_full_name else ""
            filename = filename or f"image{self._image_counter}{ext}"
            folder = Path(self.options.images_folder)
            folder.mkdir(parents=True, exist_ok=True)
            filepath = folder / filename
            filepath.write_bytes(img.image_bytes)

            # Use alias if set, otherwise compute relative path
            if self.options.images_folder_alias:
                url = f"{self.options.images_folder_alias.rstrip('/')}/{filename}"
            elif self._output_path is not None:
                url = os.path.relpath(filepath, self._output_path.parent)
            else:
                url = str(filepath)
            return f"![{alt}]({url})"

        # Default: inline base64 data URI
        b64 = base64.b64encode(img.image_bytes).decode("ascii")
        return f"![{alt}](data:{img.content_type};base64,{b64})"

    @staticmethod
    def _guess_image_extension(content_type: str, filename: str) -> str:
        """Determine file extension from content type or filename."""
        if filename and "." in filename:
            return ""  # filename already has extension
        mime_map = {
            "image/png": ".png",
            "image/jpeg": ".jpg",
            "image/gif": ".gif",
            "image/bmp": ".bmp",
            "image/svg+xml": ".svg",
            "image/tiff": ".tiff",
        }
        return mime_map.get(content_type, ".png")

    def _is_list_paragraph(self, para: ldm.Paragraph) -> bool:
        """Return True if this paragraph is a list item."""
        if not (para.list_format and para.list_format.is_list_item):
            return False
        pf = para.paragraph_format
        quote_level, code_style, heading_style = self._resolve_block_styles(pf.style_name or "")
        del quote_level  # a list item inside a quote still carries its marker
        return not (
            pf.is_heading or heading_style or code_style or self._is_horizontal_rule(para)
        )

    def _is_list_continuation(self, para: ldm.Paragraph) -> bool:
        """Return True for a paragraph that belongs to a list item but carries no marker."""
        lf = para.list_format
        if self._doc is None or lf is None or not lf.is_list_item:
            return False
        ll = self._doc.resolve_list_level(lf.list_id, lf.list_level_number)
        return ll is not None and ll.number_style == NumberStyle.BULLET and not ll.number_format

    @staticmethod
    def _list_level_of(para: ldm.Paragraph) -> int:
        lf = para.list_format
        return lf.list_level_number if lf and lf.is_list_item else -1

    def _ordered_start_of(self, para: ldm.Paragraph) -> int:
        """The level's StartAt when the item is numbered, else 0."""
        lf = para.list_format
        if self._doc is None or lf is None or not lf.is_list_item:
            return 0
        level = self._doc.resolve_list_level(lf.list_id, lf.list_level_number)
        if level is None or level.number_style in (NumberStyle.BULLET, NumberStyle.NONE):
            return 0
        return level.start_at

    def _convert_paragraph_tagged(self, para: ldm.Paragraph) -> tuple[_BlockTag, Optional[str]]:
        """Convert a paragraph and return (block_tag, markdown_text)."""
        is_list = self._is_list_paragraph(para)
        quote_level, code_style, heading_style = self._resolve_block_styles(
            para.paragraph_format.style_name or ""
        )
        if (para.list_format and para.list_format.is_list_item) or quote_level:
            self._trailing_list_indent = 0
        else:
            self._trailing_list_indent = max(self._list_indents.values(), default=0)
        if not (para.list_format and para.list_format.is_list_item):
            self._list_indents.clear()
        # In plain_text list mode, list items are regular blocks (no tight list)
        plain_text_list = self.options.list_export_mode == MarkdownListExportMode.PLAIN_TEXT
        if is_list and not (plain_text_list or self._is_list_continuation(para)):
            kind = self._LIST
        else:
            kind = self._BLOCK
        if kind == self._BLOCK and (heading_style or para.paragraph_format.is_heading):
            level = self._heading_level(para.paragraph_format, heading_style)
            kind = (
                self._SETEXT
                if self._is_setext_heading(level, heading_style)
                else self._HEADING
            )
        elif kind == self._BLOCK and self._is_horizontal_rule(para):
            kind = self._RULE
        tag = _BlockTag(
            kind,
            quote_level,
            is_code=bool(code_style),
            list_level=self._list_level_of(para),
            ordered_start=self._ordered_start_of(para),
        )
        md = self._convert_paragraph(para)
        return tag, md

    def _convert_paragraph(self, para: ldm.Paragraph) -> Optional[str]:
        pf = para.paragraph_format
        style_name = pf.style_name

        # Horizontal rule takes precedence over everything else
        if self._is_horizontal_rule(para):
            quote_level = self._resolve_block_styles(style_name or "")[0]
            rule = (
                self._format_quote(_HORIZONTAL_RULE, quote_level)
                if quote_level
                else _HORIZONTAL_RULE
            )
            return self._indent_within_list(rule, para)

        is_code_block = bool(self._resolve_block_styles(style_name)[1])

        # Handle empty paragraphs according to empty_paragraph_export_mode
        if self._is_empty_paragraph(para) and not (is_code_block and para.text):
            mode = self.options.empty_paragraph_export_mode
            if mode == MarkdownEmptyParagraphExportMode.NONE:
                return None
            markers = self._format_empty_block(para, style_name)
            if markers is not None:
                return markers
            if mode == MarkdownEmptyParagraphExportMode.MARKDOWN_HARD_LINE_BREAK:
                return "\\"
            # "empty_line" is the default — returns ""

        # Use content_sequence only when the paragraph actually mixes images with
        # text runs.  Every DOCX-reader paragraph has a non-empty content_sequence,
        # so checking merely for non-emptiness would send image-free paragraphs
        # (code-blocks, quotes, list items) down this path unnecessarily.
        if any(isinstance(i, ldm.Shape) for i in para._children):
            output_parts: list[str] = []
            pending_runs: list[ldm.Run] = []

            def _flush_runs() -> None:
                if not pending_runs:
                    return
                text = self._convert_runs(pending_runs, is_code_block, para)
                pending_runs.clear()
                if text:
                    formatted = self._format_text_part(text, pf, style_name, is_code_block, para)
                    if formatted is not None:
                        output_parts.append(formatted)

            for item in para._children:
                if (
                    isinstance(item, ldm.Shape)
                    and item.has_image
                    and item.image_data is not None
                ):
                    _flush_runs()
                    output_parts.append(self._render_image(item))
                elif isinstance(item, ldm.Run):
                    pending_runs.append(item)
            _flush_runs()

            if not output_parts:
                return ""
            # Images and runs are inline neighbours: any line break between them
            # already came out of a run, so adding one here would double it.
            return "".join(output_parts)

        # Legacy path: no content_sequence (non-reader sources or text-only paragraphs).
        # Images from inline_extras come first, then the paragraph text.
        image_parts = [
            self._render_image(item)
            for item in para._children
            if isinstance(item, ldm.Shape) and item.has_image and item.image_data is not None
        ]

        text = self._convert_runs(self._runs_with_fields(para, is_code_block), is_code_block, para)
        text_part = (
            self._format_text_part(text, pf, style_name, is_code_block, para) if text else None
        )

        parts = image_parts[:]
        if text_part is not None:
            parts.append(text_part)

        if not parts:
            return ""
        return "\n".join(parts)

    def _format_text_part(
        self,
        text: str,
        pf: ldm.ParagraphFormat,
        style_name: str,
        is_code_block: bool,
        para: ldm.Paragraph,
    ) -> Optional[str]:
        """Apply paragraph-level formatting to an already-assembled text string."""
        quote_level, code_style, heading_style = self._resolve_block_styles(style_name)
        # Leading whitespace behind a quote mark or a list marker is stripped on re-read.
        if not code_style and (
            _FOOTNOTE_STYLE_NAME in style_name
            or quote_level
            or (para.list_format and para.list_format.is_list_item)
        ):
            text = text.lstrip(" \t")
        if not text:
            return None
        is_list_item = bool(para.list_format and para.list_format.is_list_item)
        writes_markers = bool(
            heading_style
            or code_style
            or quote_level
            or is_list_item
            or self._in_footnotes
            or _FOOTNOTE_STYLE_NAME in style_name
        )
        if not writes_markers:
            text = self._escape_content_lines(text)
        if not (code_style or is_code_block):
            text = self._wrap_leading_whitespace(text)
        if heading_style:
            body = self._format_heading(text, self._heading_level(pf, heading_style), heading_style)
        elif code_style:
            body = self._format_code_block(text, self._extract_code_language(code_style), code_style)
        else:
            body = text
        if is_list_item and not (code_style and self._is_list_continuation(para)):
            body = self._format_list_item(body, para)
            if not quote_level:
                return self._continue_footnote(body, style_name)
        if quote_level:
            body = self._format_quote(body, quote_level)
        elif body is text and not (self._in_footnotes or _FOOTNOTE_STYLE_NAME in style_name):
            return text
        return self._indent_within_list(self._continue_footnote(body, style_name), para)

    @staticmethod
    def _wrap_leading_whitespace(text: str) -> str:
        """Port of ``MarkdownParagraphWriter.WrapIntoInlineCode``.

        Whitespace a line opens with is stripped when it is read back, so it
        goes inside inline-code delimiters, which keep it.  Enough of it to
        make the line a code block is left alone: that comes back as it was
        written, and wrapping it would only bury the block.
        """
        out = []
        for line in text.split("\n"):
            stripped = line.lstrip(" \t")
            lead = len(line) - len(stripped)
            wrap = stripped and 0 < lead < _INDENTED_CODE_INDENT
            out.append(f"`{line[:lead]}`{stripped}" if wrap else line)
        return "\n".join(out)

    @staticmethod
    def _escape_content_lines(text: str) -> str:
        """Port of ``MarkdownParagraphWriter.EscapeContentLine``.

        A paragraph that writes no markers of its own but whose text opens
        the way a block does reads back as that block, so the opening is
        spelled out: a backslash before it, or before the dot of what would
        be an ordered list item, which no backslash can precede.
        """
        out = []
        for line in text.split("\n"):
            stripped = line.lstrip(" \t")
            if stripped and is_block_start(stripped):
                if is_ordered_list_item(stripped):
                    head, _, rest = stripped.partition(".")
                    line = f"{head}\\.{rest}" if rest or "." in stripped else line
                elif stripped[0] in _ESCAPABLE_PUNCTUATION:
                    line = f"\\{stripped}"
            out.append(line)
        return "\n".join(out)

    def _continue_footnote(self, body: str, style_name: str) -> str:
        """Indent a footnote definition's later paragraphs to stay inside it.

        "Lazy continuation in footnotes is blank line + indentation spaces"
        (MarkdownParagraphWriter.WriteContentLines): every line but the very
        first of the definition is indented, or the blank line before it ends
        the definition instead of continuing it.  Definitions are collected at
        the end of the document, so the first one opens the region the way
        MarkdownWriter.WriteDefinitions does -- a paragraph in it need not
        carry the Footnote style itself to be part of a definition.
        """
        self._in_footnotes = self._in_footnotes or _FOOTNOTE_STYLE_NAME in style_name
        if not self._in_footnotes or body.startswith(_FOOTNOTE_REFERENCE_OPENING):
            return body
        pad = " " * _FOOTNOTE_CONTINUATION_INDENT
        return "\n".join(pad + line if line else line for line in body.split("\n"))

    def _format_empty_block(self, para: ldm.Paragraph, style_name: str) -> Optional[str]:
        """The markers a paragraph owns even with no content of its own.

        MarkdownParagraphWriter writes its style markers before the content
        lines, so an empty fenced block still gets its two fences and an empty
        list item still gets its label.  Dropping them merges the blocks that
        surrounded them into one when the export is read back.
        """
        quote_level, code_style, _ = self._resolve_block_styles(style_name)
        if code_style:
            body = self._format_code_block("", self._extract_code_language(code_style), code_style)
            if not body:
                return None
            if quote_level:
                body = self._format_quote(body, quote_level)
            return self._indent_within_list(body, para)
        if self._is_list_paragraph(para):
            return self._format_list_item("", para).rstrip()
        if quote_level:
            return self._indent_within_list(self._format_quote("", quote_level).rstrip(), para)
        return None

    def _style_chain(self, style_name: str) -> list[str]:
        """A paragraph's own style name followed by its base styles."""
        names: list[str] = []
        seen: set[str] = set()
        name = style_name
        while name and name not in seen:
            seen.add(name)
            names.append(name)
            style = self._doc.find_style(name) if self._doc is not None else None
            name = style.base_style_name if style is not None else ""
        return names

    def _resolve_block_styles(self, style_name: str) -> tuple[int, str, str]:
        """(quote level, code style, heading style) collected along the base-style
        chain -- how MarkdownParagraphWriter reads a paragraph's block types."""
        quote_level = 0
        code_style = ""
        heading_style = ""
        for name in self._style_chain(style_name):
            if name.startswith(_QUOTE_STYLE_NAME):
                quote_level += 1
            elif name.startswith(_CODE_STYLE_NAMES):
                code_style = code_style or name
            elif name.startswith((_HEADING_STYLE_NAME, _SETEXT_HEADING_STYLE_NAME)):
                heading_style = heading_style or name
        return quote_level, code_style, heading_style

    @staticmethod
    def _heading_level(pf: ldm.ParagraphFormat, heading_style: str) -> int:
        if pf.is_heading:
            return min(pf.outline_level + 1, 6)
        digits = "".join(c for c in heading_style if c.isdigit())
        return min(int(digits), 6) if digits else 1

    def _indent_within_list(self, text: str, para: ldm.Paragraph) -> str:
        """Indent a heading/code/quote block to its enclosing list item's content column."""
        lf = para.list_format
        if lf is None or not lf.is_list_item:
            return text
        level = lf.list_level_number
        if self._is_list_continuation(para):
            width = self._list_indents.get(level + 1, 0)
        else:
            width = self._list_indents.get(level, 0)
            self._drop_indents_below(level)
        if not width:
            return text
        pad = " " * width
        return "\n".join(pad + line if line else line for line in text.split("\n"))

    def _drop_indents_below(self, level: int) -> None:
        for deeper in [key for key in self._list_indents if key > level]:
            del self._list_indents[deeper]

    # ------------------------------------------------------------------
    # Horizontal rule detection
    # ------------------------------------------------------------------

    @staticmethod
    def _is_empty_paragraph(para: ldm.Paragraph) -> bool:
        """Return True if the paragraph has no visible content."""
        if visible_runs(para) and any(r.text and r.text.strip() for r in visible_runs(para)):
            return False
        if para._children and any(
            isinstance(i, ldm.Shape) for i in para._children
        ):
            return False
        if para._children and any(
            isinstance(i, ldm.Shape) and i.has_image for i in para._children
        ):
            return False
        text = para.text.strip() if para.text else ""
        return not text

    def _is_horizontal_rule(self, para: ldm.Paragraph) -> bool:
        first = para._children[0] if para._children else None
        if isinstance(first, ldm.Shape) and is_horizontal_rule_shape(first):
            return True

        pf = para.paragraph_format

        # Bottom-border based detection
        if pf.borders:
            bottom = pf.borders[0]  # BorderType.Bottom (slot 0)
            if bottom.line_style > 0 and bottom.line_width >= 1.5:
                return True

        _, code_style, heading_style = self._resolve_block_styles(pf.style_name or "")
        if code_style or heading_style or pf.is_heading:
            return False
        text = para.text.strip()
        if re.match(r"^[-*_]{3,}$", text.replace(" ", "")):
            return True

        return False

    # ------------------------------------------------------------------
    # Run conversion
    # ------------------------------------------------------------------

    # Regex for pre-rendered inline links: [text](url)
    # Destination and title exactly as the spec spells them: a bracketed
    # destination or one that holds no whitespace, then at most one quoted
    # title.  Anything looser turns text that only looks like a link -- an
    # unquoted title, a stray paren -- into a reference and tears it in half.
    # An image is left inline: its reference form would need a definition of
    # its own, and MarkdownImageWriter writes the source in place instead.
    _INLINE_LINK_RE = re.compile(
        r"(?<!!)\[((?:\\.|[^\[\]\\])+)\]"
        r"\((<[^<>\n]*>|(?:\\.|[^\s()\\]|\((?:\\.|[^()\\])*\))*)"
        r"((?:[ \t]+(?:\"(?:\\.|[^\"\\])*\"|'(?:\\.|[^'\\])*'|\((?:\\.|[^()\\])*\)))?)"
        r"[ \t]*\)"
    )

    def _convert_runs(
        self,
        runs: list[ldm.Run],
        is_code_block: bool,
        para: Optional[ldm.Paragraph] = None,
    ) -> str:
        # When the paragraph style itself defines bold/italic, suppress those
        # markers on runs so Markdown doesn't double-apply emphasis.
        style_bold, style_italic = self._get_style_emphasis(para)

        pieces: list[tuple[str, RunFormatting, list[str], str]] = []
        for run in runs:
            text = run.text or ""
            if not text:
                continue
            fmt = self._get_run_formatting(run)
            if style_bold and fmt.bold:
                fmt.bold = False
            if style_italic and fmt.italic:
                fmt.italic = False
            if "\n" in text and not is_code_block and not fmt.code:  # hard break, not literal code newline
                text = _HARD_BREAK_RE.sub("\\\\\n", text)
            pieces.append(
                (text, fmt, self._emphasis_markers(fmt, is_code_block), run.font.style_name)
            )

        parts: list[str] = []
        open_markers: list[tuple[str, int]] = []
        for index, (text, fmt, needed, style_name) in enumerate(pieces):
            keep = next(
                (i for i, (m, _) in enumerate(open_markers) if m not in needed),
                len(open_markers),
            )
            closing = open_markers[keep:]
            del open_markers[keep:]
            still_open = [m for m, _ in open_markers]
            # Delimiters nest, so the one that stays open longest has to open
            # first: "***foo** bar*" is italic wrapping bold, not the reverse.
            opening = sorted(
                (m for m in needed if m not in still_open),
                key=lambda m: -self._marker_lifetime(pieces, index, m),
            )
            self._close_markers(parts, closing, opening)
            for marker in opening:
                open_markers.append((marker, len(parts)))
                parts.append(marker)
            parts.append(
                self._apply_formatting(
                    text, fmt, is_code_block, emphasis=False, style_name=style_name
                )
            )
        self._close_markers(parts, open_markers, [])
        return "".join(parts)

    def _runs_with_fields(self, para: ldm.Paragraph, is_code_block: bool) -> list[ldm.Run]:
        """The paragraph's runs with its HYPERLINK fields folded into links.

        ``MarkdownHyperlinkWriter`` builds "[result](target)" out of a field's
        code and its result; the light document model keeps the field as
        ``FieldStart``/``FieldSeparator``/``FieldEnd`` around ordinary runs, so
        the pair is folded into one run of markup here and everything after
        this point sees the same shape a Markdown link is read into.
        """
        out: list[ldm.Run] = []
        stack: list[tuple[list[ldm.Run], list[ldm.Run]]] = []
        for child in para._children:
            if isinstance(child, ldm.FieldStart):
                stack.append(([], []))
            elif isinstance(child, ldm.FieldSeparator):
                if stack:
                    stack[-1][0].append(None)  # marks the end of the code section
            elif isinstance(child, ldm.FieldEnd):
                if not stack:
                    continue
                code, result = stack.pop()
                target = self._hyperlink_target([r for r in code if r is not None])
                folded = self._fold_hyperlink(target, result, is_code_block)
                (stack[-1][1] if stack else out).extend(folded)
            elif isinstance(child, ldm.Run):
                if not stack:
                    out.append(child)
                elif None in stack[-1][0]:
                    stack[-1][1].append(child)
                else:
                    stack[-1][0].append(child)
        return out

    @staticmethod
    def _hyperlink_target(code_runs: list[ldm.Run]) -> tuple[str, str]:
        """The address and screen tip of a HYPERLINK field code.

        ``FieldCodeHyperlink``: the address is the first bare argument, ``\\l``
        names a location within it and ``\\o`` the tip Markdown writes as the
        link title.  Both come back empty for any other field.
        """
        code = "".join(run.text or "" for run in code_runs).strip()
        if not code.upper().startswith(_HYPERLINK_FIELD_NAME):
            return "", ""
        address = sub_address = title = ""
        pending = ""
        for switch, quoted, bare in _FIELD_ARGUMENT_RE.findall(
            code[len(_HYPERLINK_FIELD_NAME):]
        ):
            if switch:
                pending = switch
                continue
            value = (quoted or bare).replace("\\\\", "\\")
            if pending == _FIELD_SWITCH_LOCATION:
                sub_address = value
            elif pending == _FIELD_SWITCH_SCREEN_TIP:
                title = value
            elif not address:
                address = value
            pending = ""
        if sub_address:
            address = f"{address}#{sub_address}"
        return address, title

    def _fold_hyperlink(
        self, target: tuple[str, str], result: list[ldm.Run], is_code_block: bool
    ) -> list[ldm.Run]:
        address, title = target
        label = self._convert_runs(result, is_code_block)
        if not address or not label:
            return result
        destination = f"{_link_destination(address)}{_link_title(title)}"
        return [
            ldm.Run(
                text=f"[{label}]({destination})",
                font=ldm.Font(style_name=_HYPERLINK_STYLE_NAME),
            )
        ]

    def _close_markers(
        self, parts: list[str], closing: list[tuple[str, int]], opening: list[str]
    ) -> None:
        """Close *closing*, moving the delimiters in front of any trailing
        whitespace and off the delimiters *opening* right after them.

        ``MarkdownEmphasisWriterBase.Close`` inserts a closing delimiter at
        the last non-whitespace position: one written after a space cannot
        close the emphasis, and both delimiters read back as literal text.
        Where a closing delimiter would run into an opening one of the same
        character, ``SwitchToUnderscore`` writes the underscore form of it
        instead -- the two would otherwise read as one longer delimiter.
        """
        if not closing:
            return
        for index, (marker, position) in enumerate(closing):
            alt = _UNDERSCORE_MARKERS.get(marker)
            # closing[0] is the outermost, so its delimiter is written last --
            # the one an opening delimiter can run into.
            abuts = index == 0 and any(o[0] == marker[0] for o in opening)
            if alt and abuts and self._can_open(parts, position):
                parts[position] = alt
                closing[index] = (alt, position)
        trailing = ""
        while parts and not parts[-1].strip():
            trailing = parts.pop() + trailing
        if parts:
            body = parts[-1]
            stripped = body.rstrip()
            # A hard line break is a backslash and the newline together, so the
            # delimiter goes in front of both.
            if stripped.endswith("\\") and body[len(stripped):].startswith("\n"):
                stripped = stripped[:-1]
            trailing = body[len(stripped):] + trailing
            parts[-1] = stripped
        parts.extend(self._closing_marker(m) for m, _ in reversed(closing))
        if trailing:
            parts.append(trailing)

    @staticmethod
    def _can_open(parts: list[str], position: int) -> bool:
        """``OpeningFlanking == FlankingType.Left``: an underscore delimiter
        only opens where the character before it is not a word character."""
        before = "".join(parts[:position])
        return not before or not (before[-1].isalnum() or before[-1] == "_")

    def _get_style_emphasis(self, para: Optional[ldm.Paragraph]) -> tuple[bool, bool]:
        """Return (bold, italic) defined by the paragraph's style font.

        Looks up the style in the document's styles list and checks its
        font properties.  This lets us suppress emphasis inherited from the
        style rather than applied as direct run formatting.
        """
        if para is None or self._doc is None:
            return False, False
        style_name = para.paragraph_format.style_name
        if not style_name:
            return False, False
        style = self._doc.find_style(style_name)
        if style is None or style.font is None:
            return False, False
        return style.font.bold, style.font.italic

    @staticmethod
    def _definition_destination(destination: str) -> str:
        """A definition ends at whitespace, so bracket one that holds any.

        MarkdownLinkDefinitionWriter.WriteDefinitions: "spaces ... in link
        definitions are detected as the end of the link".
        """
        target, _, title = destination.partition(' "')
        if not target:
            target = "<>"
        elif re.search(r"\s", target) and not target.startswith("<"):
            target = f"<{target}>"
        return f'{target} "{title}' if title else target

    @staticmethod
    def _written_destination(match: re.Match) -> str:
        """The destination and title of a link, as they were written."""
        return f"{match.group(2)}{match.group(3)}"

    @staticmethod
    def _destination_key(destination: str) -> str:
        """What two written destinations have to share to be the same link."""
        inner = destination[1:-1] if destination.startswith("<") and destination.endswith(">") else destination
        return inner.replace("\\", "")

    def _apply_link_export_mode(self, blocks: list[tuple[_BlockTag, str]]) -> None:
        """Turn inline links into reference blocks, in place.

        MarkdownLinkExportMode.Auto -- the default -- keeps a link inline
        unless it "is mentioned more than once in a document"; Reference sends
        every one to a definition. Code blocks are left alone: a "[x](y)" in
        there is code, not a link.
        """
        mode = self.options.link_export_mode
        if mode == MarkdownLinkExportMode.INLINE:
            return
        seen: dict[str, int] = {}
        for tag, text in blocks:
            if not tag.is_code:
                for match in self._INLINE_LINK_RE.finditer(text):
                    key = self._destination_key(self._written_destination(match))
                    seen[key] = seen.get(key, 0) + 1
        repeated = {dest for dest, count in seen.items() if count > 1}
        if mode == MarkdownLinkExportMode.AUTO and not repeated:
            return

        labels: dict[str, str] = {}

        def to_reference(match: re.Match) -> str:
            destination = self._written_destination(match)
            key = self._destination_key(destination)
            if mode == MarkdownLinkExportMode.AUTO and key not in repeated:
                return match.group(0)
            label = labels.get(key)
            if label is None:
                label = str(len(labels) + 1)
                labels[key] = label
                self._reference_links.append((label, destination))
            return f"[{match.group(1)}][{label}]"

        for index, (tag, text) in enumerate(blocks):
            if not tag.is_code:
                blocks[index] = (tag, self._INLINE_LINK_RE.sub(to_reference, text))

    def _get_run_formatting(self, run: ldm.Run) -> RunFormatting:
        f = run.font
        fmt = RunFormatting()
        fmt.bold = f.bold
        fmt.italic = f.italic
        fmt.underline = f.underline > 0
        fmt.strikethrough = f.strike_through
        fmt.code = bool(f.style_name and f.style_name.startswith(_INLINE_CODE_STYLE_NAME))
        if f.style_name != _FOOTNOTE_REFERENCE_STYLE_NAME:
            fmt.superscript = f.superscript
            fmt.subscript = f.subscript
        return fmt

    @staticmethod
    def _format_inline_code(text: str, style_name: str) -> str:
        """Wrap *text* in the delimiter the character style asks for, widened
        past any backtick run it contains and spaced off an adjacent backtick."""
        width = max(
            [max(_get_number_after_substring(style_name, _INLINE_CODE_STYLE_NAME), 1)]
            + [len(run) + 1 for run in _BACKTICK_RUN_RE.findall(text)]
        )
        pad = (
            " "
            if text.startswith("`")
            or text.endswith("`")
            or (text.startswith(" ") and text.endswith(" ") and text.strip())
            else ""
        )
        return f"{'`' * width}{pad}{text}{pad}{'`' * width}"

    @staticmethod
    def _marker_lifetime(pieces: list, index: int, marker: str) -> int:
        """How many further runs *marker* stays open across, without a break."""
        span = 0
        while index + span + 1 < len(pieces) and marker in pieces[index + span + 1][2]:
            span += 1
        return span

    def _emphasis_markers(self, fmt: RunFormatting, is_code_block: bool = False) -> list[str]:
        """Emphasis delimiters wrapping a run, in the order they are opened."""
        if is_code_block:
            return []
        markers: list[str] = []
        if fmt.underline and self.options.export_underline:
            markers.append(_UNDERLINE_MARKER)
        if fmt.subscript:
            markers.append(_SUBSCRIPT_MARKER)
        if fmt.superscript:
            markers.append(_SUPERSCRIPT_MARKER)
        if fmt.strikethrough and self.options.export_strikethrough:
            markers.append(_STRIKETHROUGH_MARKER)
        if fmt.bold:
            markers.append(_BOLD_MARKER)
        if fmt.italic:
            markers.append(_ITALIC_MARKER)
        return markers

    @classmethod
    def _to_autolink(cls, match: re.Match) -> str:
        """Port of ``MarkdownHyperlinkWriter.IsValidAutolink``.

        A link whose destination is its own text is written between angle
        brackets, and a mail destination may carry the scheme the text drops.
        """
        label, destination, title = match.group(1), match.group(2), match.group(3)
        if title.strip():
            return match.group(0)
        target = _unescape_markup_symbols(destination)
        bare = target[len(_MAILTO_PREFIX):] if target.lower().startswith(_MAILTO_PREFIX) else target
        if label not in (target, bare) or not is_valid_autolink(target):
            return match.group(0)
        return f"{_AUTOLINK_OPENING}{target}{_AUTOLINK_CLOSING}"

    @staticmethod
    def _closing_marker(marker: str) -> str:
        """What closes *marker*: a tag closes with its own, the rest repeat."""
        return _CLOSING_MARKERS.get(marker, marker)

    def _apply_formatting(
        self,
        text: str,
        fmt: RunFormatting,
        is_code_block: bool,
        emphasis: bool = True,
        style_name: str = "",
    ) -> str:
        if not text:
            return text

        if style_name == _HYPERLINK_STYLE_NAME:
            text = self._INLINE_LINK_RE.sub(self._to_autolink, text)

        escapable = not fmt.code and not is_code_block and style_name != _HYPERLINK_STYLE_NAME
        if self.options.escape_special_chars and escapable:
            text = self._escape_markdown(text)
            if style_name != _FOOTNOTE_REFERENCE_STYLE_NAME:
                text = _FOOTNOTE_MARK_RE.sub(r"\\\g<0>", text)

        if fmt.code:
            return self._format_inline_code(text, style_name)

        result = text
        if not emphasis:
            return result
        if fmt.bold and fmt.italic:
            result = f"***{result}***"
        elif fmt.bold:
            result = f"**{result}**"
        elif fmt.italic:
            result = f"*{result}*"

        if fmt.strikethrough and self.options.export_strikethrough:
            result = f"{_STRIKETHROUGH_MARKER}{result}{_STRIKETHROUGH_MARKER}"

        for marker in (_SUPERSCRIPT_MARKER, _SUBSCRIPT_MARKER):
            if getattr(fmt, "superscript" if marker == _SUPERSCRIPT_MARKER else "subscript"):
                result = f"{marker}{result}{self._closing_marker(marker)}"

        if fmt.underline and self.options.export_underline:
            result = f"{_UNDERLINE_MARKER}{result}{_UNDERLINE_MARKER}"

        return result

    @staticmethod
    def _escape_markdown(text: str) -> str:
        """Backslash the emphasis delimiters so text reads back as text.

        MarkdownParagraphWriter.AppendText escapes exactly these two, and only
        outside code (MarkdownUtil.gEscapableMarkupCharacters).  An underscore
        between two word characters is left alone: it cannot open or close
        emphasis there, so a backslash would only add noise to names like
        ``order_count``.
        """
        return _ESCAPABLE_MARKUP_RE.sub(r"\\\g<0>", text)

    # ------------------------------------------------------------------
    # Heading
    # ------------------------------------------------------------------

    def _is_setext_heading(self, level: int, style_name: str) -> bool:
        """Whether this heading is written as a text line plus an underline."""
        wanted = (
            self.options.heading_style == HeadingStyle.SETEXT
            or _SETEXT_HEADING_STYLE_NAME in style_name
        )
        return wanted and min(max(level, 1), 6) <= 2

    def _format_heading(self, text: str, level: int, style_name: str) -> str:
        level = min(max(level, 1), 6)
        use_setext = self._is_setext_heading(level, style_name)
        if use_setext and level <= 2:
            underline = "=" if level == 1 else "-"
            width = max(len(line) for line in text.split("\n"))
            return f"{text}\n{underline * width}"
        # Any run of '#' the text ends on reads back as a closing sequence and
        # is eaten (WriteAtxHeadingClosing); writing one of our own keeps it.
        closing = f" {'#' * level}"
        suffix = closing if _TRAILING_HASHES_RE.search(text) else ""
        return f"{'#' * level} {text}{suffix}"

    # ------------------------------------------------------------------
    # Code block
    # ------------------------------------------------------------------

    def _extract_code_language(self, style_name: str) -> str:
        if "." in style_name:
            return style_name.split(".", 1)[1]
        return ""

    @staticmethod
    def _fence_char_for(style_name: str) -> str:
        """The delimiter the block was fenced with, as the style name kept it."""
        return _TILDE_FENCE_CHAR if _TILDE_FENCE_CHAR in style_name else _BACKTICK_FENCE_CHAR

    @staticmethod
    def _code_fence_for(text: str, char: str) -> str:
        closing = _CLOSING_FENCE_RES[char]
        longest = max(
            (len(m.group(1)) for m in map(closing.match, text.split("\n")) if m),
            default=0,
        )
        return char * max(_FENCED_CODE_DELIMITER_LENGTH, longest + 1)

    def _is_indented_code(self, style_name: str) -> bool:
        return (
            self.options.code_block_style.value == CodeBlockStyle.INDENTED.value
            or style_name.startswith(_INDENTED_CODE_STYLE_NAME)
        )

    def _format_code_block(self, text: str, language: str, style_name: str = "") -> str:
        is_indented = self._is_indented_code(style_name)
        if not is_indented:
            fence = self._code_fence_for(text, self._fence_char_for(style_name))
            return f"{fence}{language}\n{text}\n{fence}"
        pad = " " * (_INDENTED_CODE_INDENT + self._trailing_list_indent)
        return "\n".join(f"{pad}{line}" if line else line for line in text.split("\n"))

    # ------------------------------------------------------------------
    # Block quote
    # ------------------------------------------------------------------

    def _extract_quote_level(self, style_name: str) -> int:
        level = self._resolve_block_styles(style_name)[0]
        if level:
            return level
        match = re.search(rf"{_QUOTE_STYLE_NAME}(\d+)?", style_name)
        return int(match.group(1)) if match and match.group(1) else 1

    def _format_quote(self, text: str, level: int) -> str:
        # Quote marks of an opening sequence are not separated with spaces.
        prefix = ">" * level + " "
        lines = text.split("\n")
        return "\n".join(f"{prefix}{line}" for line in lines)

    # ------------------------------------------------------------------
    # List item
    # ------------------------------------------------------------------

    def _format_list_item(self, text: str, para: ldm.Paragraph) -> str:
        lf = para.list_format
        assert lf is not None
        level = lf.list_level_number
        list_id = lf.list_id

        # Plain text list mode: render as indented text without markers
        if self.options.list_export_mode == MarkdownListExportMode.PLAIN_TEXT:
            indent = "    " * level
            return f"{indent}{text}"

        list_type, marker = self._get_list_type(list_id, level)
        if not marker:
            return f"{' ' * self._list_indents.get(level + 1, 0)}{text}"
        width = self._list_indents.get(level, 0)
        missed = self._missed_list_levels(list_id, level)
        self._drop_indents_below(level)
        content = width + len(missed) + len(marker) + 1
        self._list_indents[level + 1] = content
        first, _, rest = text.partition("\n")
        if rest:
            pad = " " * content
            rest = "\n".join(pad + line if line else line for line in rest.split("\n"))
            first = f"{first}\n{rest}"
        return f"{' ' * width}{missed}{marker} {first}"

    def _missed_list_levels(self, list_id: int, level: int) -> str:
        """Labels of the levels this item nests under that own no paragraph.

        ``InsertMissedListLevels``: "- - foo" is a single item at level 1, and
        dropping the level it sits in makes it a level-0 item on re-read.
        """
        if level == 0 or level in self._list_indents:
            return ""
        return "".join(
            f"{self._get_list_type(list_id, ancestor)[1]} " for ancestor in range(level)
        )

    def _get_list_type(self, list_id: int, level: int) -> tuple[str, str]:
        """Determine list type and marker from the document's list definitions."""
        if self._doc is not None and list_id > 0:
            ll = self._doc.resolve_list_level(list_id, level)
            if ll is not None and ll.number_style == NumberStyle.BULLET and not ll.number_format:
                return ("continuation", "")
            if ll is not None and ll.number_style not in (NumberStyle.BULLET, NumberStyle.NONE):
                return self._get_ordered_marker(ll.start_at, ll.number_format)
            if ll is not None and ll.number_format in _BULLET_MARKER_CHARS:
                return ("bullet", ll.number_format)
        return ("bullet", self.options.list_marker.value)

    @staticmethod
    def _get_ordered_marker(start: int, number_format: str = "") -> tuple[str, str]:
        # GetListLabel writes the level's StartAt on every item, not a running
        # count: renderers number the list themselves, and a counter drifts.
        suffix = number_format[-1] if number_format[-1:] in _ORDERED_MARKER_CHARS else "."
        return ("ordered", f"{start}{suffix}")

    # ------------------------------------------------------------------
    # Table conversion
    # ------------------------------------------------------------------

    def _convert_table(self, table: ldm.Table) -> str:
        if not table.rows:
            return ""

        # export_as_html: render table as raw HTML
        if self.options.export_as_html == MarkdownExportAsHtml.TABLES:
            return self._convert_table_as_html(table)

        # Determine number of columns
        num_cols = max(len(row.cells) for row in table.rows)

        # Extract cell texts and alignments
        cell_data: list[list[tuple[str, str]]] = []
        for row in table.rows:
            row_cells: list[tuple[str, str]] = []
            for cell in row.cells:
                text = self._extract_cell_text(cell)
                align = self._resolve_cell_alignment(cell)
                row_cells.append((text, align))
            cell_data.append(row_cells)

        # Column widths
        col_widths = [3] * num_cols
        for row_cells in cell_data:
            for j, (text, _) in enumerate(row_cells):
                if j < num_cols:
                    col_widths[j] = max(col_widths[j], len(text))

        # Render rows
        lines: list[str] = []
        for i, row_cells in enumerate(cell_data):
            cells: list[str] = []
            for j in range(num_cols):
                if j < len(row_cells):
                    text = row_cells[j][0]
                else:
                    text = ""
                cells.append(text.ljust(col_widths[j]))
            lines.append("| " + " | ".join(cells) + " |")

            # Header separator after the first row
            if i == 0:
                seps: list[str] = []
                for j in range(num_cols):
                    width = col_widths[j]
                    align = row_cells[j][1] if j < len(row_cells) else "left"
                    if align == "center":
                        sep = ":" + "-" * (width - 2) + ":"
                    elif align == "right":
                        sep = "-" * (width - 1) + ":"
                    else:
                        sep = "-" * width
                    seps.append(sep)
                lines.append("| " + " | ".join(seps) + " |")

        return "\n".join(lines)

    def _extract_cell_text(self, cell: ldm.Cell) -> str:
        parts: list[str] = []
        for para in cell.paragraphs:
            if para._children:
                # Preserve XML element order using content_sequence
                para_parts: list[str] = []
                for item in para._children:
                    if isinstance(item, ldm.Shape) and item.has_image and item.image_data:
                        para_parts.append(self._render_image(item))
                    elif isinstance(item, ldm.Run):
                        text = item.text or ""
                        if text:
                            fmt = self._get_run_formatting(item)
                            para_parts.append(
                                self._apply_formatting(
                                    text, fmt, False, style_name=item.font.style_name
                                )
                            )
                if para_parts:
                    parts.append("".join(para_parts))
            else:
                # Legacy path: images first, then text runs
                for item in para._children:
                    if isinstance(item, ldm.Shape) and item.has_image and item.image_data:
                        parts.append(self._render_image(item))
                run_parts: list[str] = []
                for run in visible_runs(para):
                    text = run.text or ""
                    if text:
                        fmt = self._get_run_formatting(run)
                        text = self._apply_formatting(
                            text, fmt, False, style_name=run.font.style_name
                        )
                        run_parts.append(text)
                if run_parts:
                    parts.append("".join(run_parts))

        result = " ".join(parts)
        result = result.replace("|", "\\|")
        result = result.replace("\n", " ").replace("\r", " ")
        return result.strip()

    @staticmethod
    def _cell_alignment(cell: ldm.Cell) -> str:
        if cell.paragraphs:
            a = cell.paragraphs[0].paragraph_format.alignment
            return _ALIGN_STR.get(a, "left")
        return "left"

    # Explicit since the enum went integral: its string value used to double
    # as the alignment token consumed downstream.
    _ALIGNMENT_TOKENS = {
        TableContentAlignment.LEFT: "left",
        TableContentAlignment.CENTER: "center",
        TableContentAlignment.RIGHT: "right",
    }

    def _resolve_cell_alignment(self, cell: ldm.Cell) -> str:
        """Return cell alignment, respecting table_content_alignment override."""
        override = self._ALIGNMENT_TOKENS.get(self.options.table_content_alignment)
        if override is not None:
            return override
        return self._cell_alignment(cell)

    def _convert_table_as_html(self, table: ldm.Table) -> str:
        """Render a table as raw HTML."""
        lines: list[str] = ["<table>"]
        for i, row in enumerate(table.rows):
            lines.append("<tr>")
            tag = "th" if i == 0 else "td"
            for cell in row.cells:
                text = html.escape(self._extract_cell_text(cell))
                align = self._resolve_cell_alignment(cell)
                if align != "left":
                    lines.append(f'<{tag} style="text-align: {align}">{text}</{tag}>')
                else:
                    lines.append(f"<{tag}>{text}</{tag}>")
            lines.append("</tr>")
        lines.append("</table>")
        return "\n".join(lines)
