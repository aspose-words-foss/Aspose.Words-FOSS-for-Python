"""
Drives a parsed Markdown block tree into ``MarkdownReaderContext`` to
produce a finished LDM ``Document``.

Mirrors the event order Markdig uses when it calls the reference
``MarkdownReaderContext.Open``/``WriteText``/``Close``: a depth-first
walk of the block tree, with one twist this project's parser needs that
Markdig's doesn't — a ``FootnoteDefinitionBlock``'s children live as a
separate top-level block here, so :func:`drive_block` relocates them into
the matching ``FootnoteReferenceBlock``'s open/close window instead of
driving them at their own tree position.
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional

from aspose.words_foss import light_document_model as ldm

from .block_parser import parse_document
from .blocks import Block, BlockType
from .reader_context import MarkdownReaderContext


def _collect_footnote_definitions(root: Block) -> dict[str, Block]:
    """Map footnote label to its first ``FootnoteDefinitionBlock`` (duplicates lose)."""
    definitions: dict[str, Block] = {}

    def walk(block: Block) -> None:
        if block.type == BlockType.FOOTNOTE_DEFINITION and block.label not in definitions:
            definitions[block.label] = block
        for child in block.children:
            walk(child)

    walk(root)
    return definitions


def drive_block(ctx: MarkdownReaderContext, block: Block, footnote_defs: dict[str, Block]) -> None:
    """Feed *block* and its subtree into *ctx* in document order.

    *footnote_defs* maps footnote label to its ``FootnoteDefinitionBlock``,
    used to resolve a ``FootnoteReferenceBlock`` to the body it relocates
    inline.
    """
    if block.type == BlockType.FOOTNOTE_DEFINITION:
        return

    if block.type == BlockType.CELL:
        table = block.get_parent(BlockType.TABLE)
        if table is not None and block.parent.children.index(block) >= table.columns_count:
            return  # excess cell beyond the header's column count is ignored, per spec

    if block.type == BlockType.HTML_TAG:
        if block.is_self_contained:
            ctx.open(block)
            ctx.close(block)
        else:
            ctx.close(block) if block.is_closing else ctx.open(block)
        return

    if block.type == BlockType.LINK_TEXT and block.is_image:
        ctx.insert_image(block)
        return

    if block.type == BlockType.FOOTNOTE_REFERENCE:
        definition = footnote_defs.get(block.label)
        if definition is None:
            ctx.write_text(f"[^{block.label}]")
            return
        ctx.open(block)
        for child in definition.children:
            drive_block(ctx, child, footnote_defs)
        ctx.close(block)
        return

    ctx.open(block)

    if block.type == BlockType.TEXT:
        ctx.write_text(block.text)
    elif block.type in (BlockType.INDENTED_CODE, BlockType.FENCED_CODE, BlockType.INLINE_CODE):
        ctx.write_text(block.code)
    elif block.type == BlockType.AUTOLINK:
        ctx.write_text(block.text)
    elif block.type == BlockType.LINE_BREAK:
        # MarkdownLoadOptions.DefaultSoftLineBreakCharacter is a space: only a
        # hard break becomes a line break in the document.
        ctx.write_text("\n" if block.is_hard else " ")
    else:
        for child in block.children:
            drive_block(ctx, child, footnote_defs)

    ctx.close(block)


def build_document(root: Block, base_dir: Optional[Path] = None) -> ldm.Document:
    """Drive a pre-parsed block tree (from :func:`~.block_parser.parse_document`)
    through a fresh :class:`MarkdownReaderContext` and return the resulting LDM."""
    footnote_defs = _collect_footnote_definitions(root)
    ctx = MarkdownReaderContext(base_dir=base_dir)
    drive_block(ctx, root, footnote_defs)
    return ctx.document


def parse_and_build(
    text: str, preserve_empty_lines: bool = False, base_dir: Optional[Path] = None
) -> ldm.Document:
    """Parse *text* as Markdown; *base_dir* resolves local ``![]()`` images into real Shapes."""
    return build_document(parse_document(text, preserve_empty_lines), base_dir=base_dir)
