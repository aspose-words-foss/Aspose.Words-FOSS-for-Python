"""Style-chain resolution for :class:`MarkdownReaderContext`.

Adapts the reference's istd/``ParaPr`` chain cache (``mBlocksChainHashcodeToStyle``,
``GetStyle``, ``FetchStyle``, ``Decorate*``) to plain named LDM
:class:`~aspose.words_foss.light_document_model.Style` objects — see the
module docstring on ``reader_context.py`` for the two structural
differences this adaptation has to work around.
"""

from __future__ import annotations

from typing import Optional

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.model.enums import LineStyle, StyleIdentifier, StyleType, Underline
from aspose.words_foss.model.style_identifiers import resolve_style_identifier

from ._reader_constants import (
    _FENCED_CODE_STYLE_NAME,
    _TILDE_FENCE_MARKER,
    _HEADING_STYLE_NAME,
    _INDENTED_CODE_STYLE_NAME,
    _INLINE_CODE_STYLE_NAME,
    _LIST_STYLE_NAME,
    _NORMAL_STYLE_NAME,
    _QUOTE_STYLE_NAME,
    _SETEXT_HEADING_STYLE_NAME,
    _SPECIAL_STYLE_IDENTIFIERS,
    _UNIQUE_STYLE_BASE_NAMES,
)
from .blocks import AtxHeadingBlock, Block, BlockType, MarkdownBlockLevel


# Style.ParaPr.BorderLeft = new Border(LineStyle.Single, 18, 0x9F9F9F): the
# reference gives the width in eighths of a point, as OOXML's w:sz does.
_LEFT_BORDER_SLOT = 1
_QUOTE_BORDER_WIDTH = 18 / 8
_QUOTE_BORDER_COLOR = "9F9F9F"


class _StyleMixin:
    """Style fetch/create/decorate + the block-chain cache that decides
    when two structurally-identical blocks should reuse a style."""

    @staticmethod
    def _is_para_or_heading(block: Optional[Block]) -> bool:
        if block is None:
            return False
        if block.type == BlockType.PARAGRAPH:
            return True
        return block.type == BlockType.HTML_TAG and block.is_para_or_heading

    @staticmethod
    def _is_table(block: Optional[Block]) -> bool:
        if block is None:
            return False
        if block.type == BlockType.TABLE:
            return True
        return block.type == BlockType.HTML_TAG and block.is_table

    # ------------------------------------------------------------------
    # Style chain resolution
    # ------------------------------------------------------------------

    @staticmethod
    def _block_hash(block: Block, seed: int) -> int:
        if block.type in (BlockType.ATX_HEADING, BlockType.SETEXT_HEADING):
            h = (block.type.value * 397) ^ block.level
        elif block.type == BlockType.HTML_TAG:
            h = (block.type.value * 397) ^ block.heading_tag_level
        elif block.type in (BlockType.BULLET_LIST_ITEM, BlockType.ORDERED_LIST_ITEM):
            h = block.type.value
            h = (h * 397) ^ block.start_at
            h = (h * 397) ^ block.marker.value
            h = (h * 397) ^ block.get_level()
        elif block.type == BlockType.FENCED_CODE:
            # Everything the style name spells out has to be in the hash, or
            # two blocks that differ only there share the first one's style.
            h = (block.type.value * 397) ^ hash(block.info)
            h = (h * 397) ^ hash(block.fence_char)
        elif block.type == BlockType.HORIZONTAL_RULE:
            h = BlockType.PARAGRAPH.value
        else:
            h = block.type.value
        return (seed * 397) ^ h

    def _get_style(self, block: Block) -> ldm.Style:
        style: Optional[ldm.Style] = None
        prev_block: Optional[Block] = None
        cur_block: Optional[Block] = block
        cur_hash = 0
        can_skip_block = False

        while cur_block is not None and cur_block.type != BlockType.DOCUMENT:
            base_style = style
            cur_hash = self._block_hash(cur_block, cur_hash)
            style = self._chain_style_cache.get(cur_hash)

            if style is None:
                if cur_block.type == BlockType.PARAGRAPH:
                    if self._is_para_or_heading(prev_block):
                        can_skip_block = True
                elif cur_block.type == BlockType.HTML_TAG:
                    if self._is_para_or_heading(prev_block) and cur_block.is_para_or_heading:
                        can_skip_block = True
                elif cur_block.type == BlockType.SETEXT_HEADING:
                    base_style = self._get_style(AtxHeadingBlock(cur_block.level))
                elif cur_block.type == BlockType.LIST:
                    can_skip_block = True
                elif cur_block.type in (BlockType.BULLET_LIST_ITEM, BlockType.ORDERED_LIST_ITEM):
                    if prev_block is not None:
                        if prev_block.block_level == MarkdownBlockLevel.LEAF or not prev_block.is_first_child:
                            can_skip_block = True

                if can_skip_block:
                    style = base_style
                    can_skip_block = False
                else:
                    style = self._get_existing_style(cur_block, base_style)
                    if style is None:
                        style = self._create_style(cur_block, base_style, cur_hash)

            prev_block = cur_block
            cur_block = cur_block.parent

        assert style is not None
        return style

    def _get_existing_style(self, block: Block, base_style: Optional[ldm.Style]) -> Optional[ldm.Style]:
        if block.type == BlockType.QUOTE:
            prefix = _QUOTE_STYLE_NAME
        elif block.type in (BlockType.BULLET_LIST_ITEM, BlockType.ORDERED_LIST_ITEM):
            prefix = _LIST_STYLE_NAME
        else:
            return None
        base_name = base_style.name if base_style is not None else ""
        for style in self._chain_style_cache.values():
            if style.name.startswith(prefix) and style.base_style_name == base_name:
                return style
        return None

    def _create_style(self, block: Block, base_style: Optional[ldm.Style], hash_code: int) -> ldm.Style:
        name = self._get_style_name(block)
        style = self._fetch_style(name, StyleType.PARAGRAPH)
        if style.style_identifier != StyleIdentifier.NORMAL and base_style is not None:
            style.base_style_name = base_style.name
        self._decorate(style, block)
        self._chain_style_cache[hash_code] = style
        return style

    def _get_style_name(self, block: Block) -> str:
        if block.type in _UNIQUE_STYLE_BASE_NAMES:
            return self._get_unique_style_name(block)
        if block.type == BlockType.ATX_HEADING:
            return f"{_HEADING_STYLE_NAME} {block.level}"
        if block.type == BlockType.HTML_TAG:
            level = block.heading_tag_level
            return _NORMAL_STYLE_NAME if level == -1 else f"{_HEADING_STYLE_NAME} {level}"
        if block.type == BlockType.SETEXT_HEADING:
            return f"{_SETEXT_HEADING_STYLE_NAME}{block.level}"
        if block.type == BlockType.INDENTED_CODE:
            return _INDENTED_CODE_STYLE_NAME
        if block.type == BlockType.FENCED_CODE:
            info = block.info
            fence = _TILDE_FENCE_MARKER if block.fence_char == _TILDE_FENCE_MARKER else ""
            return f"{_FENCED_CODE_STYLE_NAME}{fence}{'.' if info else ''}{info}"
        return _NORMAL_STYLE_NAME

    def _get_unique_style_name(self, block: Block) -> str:
        base_name = _UNIQUE_STYLE_BASE_NAMES[block.type]
        used_count = self._used_block_styles.get(base_name, 0)
        if used_count == 0:
            unique_name = base_name
            used_count = 1
        else:
            unique_name = f"{base_name}{used_count}"
            used_count += 1
        self._used_block_styles[base_name] = used_count
        return unique_name

    # ------------------------------------------------------------------
    # Style fetch/decorate
    # ------------------------------------------------------------------

    def _find_style(self, name: str, style_type: int) -> Optional[ldm.Style]:
        for style in self._builder.document.styles:
            if style.name == name and style.type == style_type:
                return style
        return None

    def _fetch_style(self, name: str, style_type: int) -> ldm.Style:
        style = self._find_style(name, style_type)
        if style is not None:
            return style
        style = ldm.Style(name=name, type=style_type)
        sid = _SPECIAL_STYLE_IDENTIFIERS.get(name)
        if sid is None:
            sid = resolve_style_identifier(name, name)
        if sid >= 0:
            style.style_identifier = sid
            style.built_in = True
        else:
            # Styles.Add() reports an unrecognised name as user-defined, not Normal.
            style.style_identifier = StyleIdentifier.USER
        self._builder.document.styles.append(style)
        return style

    def _reset_char_style(self) -> None:
        self._builder.font.reset_style()

    def _reset_hyperlink_underline(self) -> None:
        # reset_style() deliberately keeps the underline toggle intact (so a genuine
        # "__underline__" span wrapping a link stays underlined past the link's own
        # end) -- but that means the Hyperlink style's own underline, merged onto the
        # font state by Font.style = ..., would otherwise never turn back off. Restore
        # it to whatever explicit "__underline__" nesting (if any) actually calls for.
        self._builder.font.underline = Underline.SINGLE if self._underline_count > 0 else Underline.NONE

    def _decorate(self, style: ldm.Style, block: Block) -> None:
        if block.type == BlockType.QUOTE:
            self._decorate_quote(style)
        elif block.type in (BlockType.INDENTED_CODE, BlockType.FENCED_CODE):
            self._decorate_code(style)
        elif block.type == BlockType.INLINE_CODE:
            self._decorate_inline_code(style)
        # Bullet/ordered list decoration is intentionally omitted: LDM's
        # Style has no list association field, so list membership is
        # carried entirely on Paragraph.list_format (see _ListMixin).

    @staticmethod
    def _decorate_quote(style: ldm.Style) -> None:
        pf = style.paragraph_format if style.paragraph_format is not None else ldm.ParagraphFormat()
        borders = [ldm.Border() for _ in range(8)]
        borders[_LEFT_BORDER_SLOT] = ldm.Border(
            line_style=LineStyle.SINGLE, line_width=_QUOTE_BORDER_WIDTH, color=_QUOTE_BORDER_COLOR
        )
        pf.borders = borders
        style.paragraph_format = pf

    @staticmethod
    def _decorate_code(style: ldm.Style) -> None:
        pf = style.paragraph_format if style.paragraph_format is not None else ldm.ParagraphFormat()
        pf.shading = ldm.Shading(foreground_pattern_color="E2E2E2")
        style.paragraph_format = pf
        font = style.font if style.font is not None else ldm.Font()
        font.name = "Consolas"
        font.size = 10.0
        style.font = font

    @staticmethod
    def _decorate_inline_code(style: ldm.Style) -> None:
        font = style.font if style.font is not None else ldm.Font()
        font.name = "Consolas"
        font.color = "C7254E"
        font.highlight_color = "F9F2F4"
        font.size = 10.0
        style.font = font

    @staticmethod
    def _decorate_hyperlink(style: ldm.Style) -> None:
        # The built-in "Hyperlink" character style has no formatting of its own unless
        # explicitly given one -- an empty <w:style> definition in styles.xml leaves a
        # viewer with nothing to fall back to, so the link renders as plain body text.
        # Matches pdf_writer's HYPERLINK_TEXT_RGB (0, 0, 238) so both outputs agree.
        font = style.font if style.font is not None else ldm.Font()
        font.color = "0000EE"
        font.underline = Underline.SINGLE
        style.font = font

    @staticmethod
    def _get_inline_code_style_name(block: Block) -> str:
        return f"{_INLINE_CODE_STYLE_NAME}{block.delimiter_length}"

    # -- Type-only declarations of host-provided state (see reader_context.py) --
    _chain_style_cache: dict
    _used_block_styles: dict
    _underline_count: int
    _builder: object
