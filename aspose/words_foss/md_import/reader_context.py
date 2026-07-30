"""
``MarkdownReaderContext`` — walks the block tree from ``block_parser`` and
populates a :class:`~aspose.words_foss.light_document_model.Document`.

This is a line-for-line adaptation of Aspose.Words' internal
``MarkdownReaderContext`` (the reference implementation supplied for this
task) onto the light document model (LDM) instead of Word's full
``Document``/``DocumentBuilder``/``Style`` object graph. Two structural
differences from the reference are unavoidable and documented at the call
site where they matter:

* The LDM has no footnote node (see the "REMOVED entirely: CommentNode,
  FootnoteNode" note in ``light_document_model.py``) — footnote bodies are
  built as ordinary paragraphs and appended to the end of the document body
  once parsing finishes; see :meth:`MarkdownDocumentBuilder.flush_footnotes`.
* ``Style`` has no list association field (unlike Aspose.Words' ``ParaPr``),
  so list membership is carried entirely on ``Paragraph.list_format`` rather
  than also being decorated onto the paragraph style object (see
  ``_list_mixin.py``).

Style-chain resolution (``GetStyle``/``FetchStyle``/``Decorate*``) lives in
``_style_mixin.py``; list-format resolution (``ApplyList``/``SetList``/
``FetchBulletList``/``FetchOrderedList``) lives in ``_list_mixin.py``. This
module owns everything else: construction, the ``open``/``close``/
``write_text`` dispatch, raw-HTML tag-state management, and footnotes.

The traversal that drives ``open``/``write_text``/``close`` in the right
order — including feeding a ``FootnoteDefinitionBlock``'s children through
this context between the matching ``FootnoteReferenceBlock``'s ``open`` and
``close`` (Markdig performs that relocation automatically before the real
Aspose.Words class ever runs; this project's parser keeps them as separate
top-level blocks) — is a later step's responsibility.
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.model.enums import ParagraphAlignment, StyleIdentifier, StyleType, Underline

from ._list_mixin import _ListMixin
from ._reader_constants import (
    _FOOTNOTE_PARAGRAPH_TYPES,
    _FOOTNOTE_REFERENCE_CLOSING,
    _FOOTNOTE_REFERENCE_OPENING,
    _FOOTNOTE_REFERENCE_STYLE_NAME,
    _FOOTNOTE_TEXT_STYLE_NAME,
    _HYPERLINK_STYLE_NAME,
    _INLINE_CODE_STYLE_NAME,
    _INLINE_HTML_HEADING_TAGS,
    _LEAF_CLOSE_TYPES,
    _NORMAL_STYLE_NAME,
    _SELF_CONTAINED_HTML_RE,
)
from ._image_loader import load_data_uri_image, load_local_image
from ._style_mixin import _StyleMixin
from .blocks import Block, BlockType, HtmlTagBlock, LinkTextBlock, RowBlock, TextBlock
from .document_builder import HtmlInsertOptions, MarkdownDocumentBuilder


# A hyperlink label becomes part of one literal "[label](uri)" run, which no
# formatting survives, so the markup the parser consumed is written back into
# the label instead of being dropped.
_LABEL_MARKUP = {
    BlockType.BOLD_INLINE: "**",
    BlockType.ITALIC_INLINE: "*",
    BlockType.STRIKETHROUGH: "~~",
}


def _flatten_text(block: Block) -> str:
    """Concatenate literal text under *block*, ignoring formatting (an image's alt text)."""
    if isinstance(block, TextBlock):
        return block.text
    return "".join(_flatten_text(c) for c in block.children)


class MarkdownReaderContext(_StyleMixin, _ListMixin):
    """Drives Markdown block-tree events into a fresh LDM ``Document``.

    Usage mirrors the reference: call :meth:`open` when entering a block,
    :meth:`write_text` for literal inline text, and :meth:`close` when
    leaving a block, in the same nesting order the block tree implies.
    ``document`` holds the result once the ``DocumentBlock`` root has been
    closed.
    """

    def __init__(self, base_dir: Optional[Path] = None) -> None:
        self._base_dir = base_dir
        self._builder = MarkdownDocumentBuilder()

        # Disable bold/italic/strike explicitly so nothing here inherits
        # stray formatting from a parent context (mirrors the reference
        # constructor's ``mBuilder.Bold = false`` block).
        self._builder.bold = False
        self._builder.italic = False
        self._builder.font.strike_through = False

        # "Cute looking" defaults (reference constructor).
        self._builder.document.doc_defaults_font = ldm.Font(name="Calibri", size=11.0)
        normal_style = self._fetch_style(_NORMAL_STYLE_NAME, StyleType.PARAGRAPH)
        normal_style.style_identifier = StyleIdentifier.NORMAL
        if normal_style.paragraph_format is None:
            normal_style.paragraph_format = ldm.ParagraphFormat()
        normal_style.paragraph_format.space_after = 8.0

        self._chain_style_cache: dict[int, ldm.Style] = {}
        self._used_block_styles: dict[str, int] = {}

        self._bullet_lists: dict[int, ldm.DocList] = {}
        self._list_containers: dict[int, Optional[Block]] = {}
        self._level_lists: list[Optional[ldm.DocList]] = [None] * 9
        self._list_id_counter = 0

        self._bold_count = 0
        self._italic_count = 0
        self._strikethrough_count = 0
        self._underline_count = 0
        self._superscript_count = 0
        self._subscript_count = 0

        self._footnote: Optional[ldm.Paragraph] = None
        self._footnote_counter = 0
        self._saved_bold_count = 0
        self._saved_italic_count = 0
        self._saved_strikethrough_count = 0
        self._saved_underline_count = 0
        self._saved_superscript_count = 0
        self._saved_subscript_count = 0

        self._html_builder: list[str] = []
        self._html_tag_starts: dict[str, int] = {}

        self._last_opened_block: Optional[Block] = None
        self._last_closed_block: Optional[Block] = None

    @property
    def document(self) -> ldm.Document:
        return self._builder.document

    @property
    def builder(self) -> MarkdownDocumentBuilder:
        return self._builder

    # ------------------------------------------------------------------
    # Open / Close / WriteText
    # ------------------------------------------------------------------

    def _end_hyperlink(self) -> None:
        """Close the capture and give the literal "[label](uri)" its style.

        Inline markup inside the label may have reset the character style on
        its way out, and the run that carries the whole link has to say it is
        a link -- it is markup, not text, so nothing in it gets escaped.
        """
        self._builder.font.style = self._fetch_style(
            _HYPERLINK_STYLE_NAME, StyleType.CHARACTER
        )
        self._builder.end_hyperlink()
        self._reset_char_style()
        self._reset_hyperlink_underline()

    def _write_label_markup(self, block: Block) -> None:
        """Put back the inline markup a hyperlink label is made of."""
        if not self._builder.is_capturing_hyperlink:
            return
        if block.type == BlockType.INLINE_CODE:
            self._builder.write("`" * max(1, block.delimiter_length))
        elif block.type in _LABEL_MARKUP:
            self._builder.write(_LABEL_MARKUP[block.type])

    def open(self, block: Block) -> None:
        t = block.type
        self._write_label_markup(block)

        if t == BlockType.HTML_TAG:
            self._open_html_tag(block)
        elif t == BlockType.QUOTE:
            if (
                block.parent is not None and block.parent.type == BlockType.DOCUMENT
                and self._last_closed_block is not None
                and self._last_closed_block.type == BlockType.QUOTE
            ):
                self._builder.writeln()
                self._set_paragraph_style(self._fetch_style(_NORMAL_STYLE_NAME, StyleType.PARAGRAPH))
        elif t in _FOOTNOTE_PARAGRAPH_TYPES:
            if self._footnote is not None:
                self._start_footnote_paragraph()
        elif t == BlockType.TABLE:
            if self._is_table(self._last_closed_block):
                self._builder.writeln()
            self._set_paragraph_style(self._fetch_style(_NORMAL_STYLE_NAME, StyleType.PARAGRAPH))
        elif t == BlockType.CELL:
            cell = self._builder.insert_cell()
            self._apply_cell_alignment(cell, block.alignment)
            self._remove_list()
        elif t == BlockType.INLINE_CODE:
            style = self._fetch_style(self._get_inline_code_style_name(block), StyleType.CHARACTER)
            self._decorate_inline_code(style)
            self._builder.font.style = style
        elif t == BlockType.LINK_TEXT:
            if block.uri is not None:
                style = self._fetch_style(_HYPERLINK_STYLE_NAME, StyleType.CHARACTER)
                self._decorate_hyperlink(style)
                self._builder.font.style = style
                self._builder.start_hyperlink(block.uri, block.title)
        elif t == BlockType.AUTOLINK:
            style = self._fetch_style(_HYPERLINK_STYLE_NAME, StyleType.CHARACTER)
            self._decorate_hyperlink(style)
            self._builder.font.style = style
            self._builder.start_hyperlink(block.uri)
        elif t == BlockType.BOLD_INLINE:
            self._builder.bold = True
            self._bold_count += 1
        elif t == BlockType.ITALIC_INLINE:
            self._builder.italic = True
            self._italic_count += 1
        elif t == BlockType.STRIKETHROUGH:
            self._builder.font.strike_through = True
            self._strikethrough_count += 1
        elif t == BlockType.UNDERLINE:
            self._builder.font.underline = Underline.SINGLE
            self._underline_count += 1
        elif t == BlockType.FOOTNOTE_REFERENCE:
            self._open_footnote_reference(block)

        self._last_opened_block = block

    def close(self, block: Block) -> None:
        t = block.type

        if t == BlockType.DOCUMENT:
            self._close_document()
        elif t == BlockType.QUOTE:
            self._set_paragraph_style(self._fetch_style(_NORMAL_STYLE_NAME, StyleType.PARAGRAPH))
        elif t in _LEAF_CLOSE_TYPES:
            self._close_leaf_block(block)
        elif t == BlockType.TABLE:
            self._builder.end_table()
            self._remove_list()
        elif t == BlockType.ROW:
            self._close_row(block)
        elif t == BlockType.LINK_TEXT:
            if block.uri is not None:
                self._end_hyperlink()
        elif t == BlockType.AUTOLINK:
            self._end_hyperlink()
        elif t == BlockType.INLINE_CODE:
            self._reset_char_style()
        elif t == BlockType.BOLD_INLINE:
            self._bold_count -= 1
            self._builder.bold = self._bold_count > 0
        elif t == BlockType.ITALIC_INLINE:
            self._italic_count -= 1
            self._builder.italic = self._italic_count > 0
        elif t == BlockType.STRIKETHROUGH:
            self._strikethrough_count -= 1
            self._builder.font.strike_through = self._strikethrough_count > 0
        elif t == BlockType.UNDERLINE:
            self._underline_count -= 1
            if self._underline_count == 0:
                self._builder.font.underline = Underline.NONE
        elif t == BlockType.FOOTNOTE_REFERENCE:
            self._close_footnote_reference()
        elif t == BlockType.HTML_TAG:
            self._close_html_tag(block)

        self._write_label_markup(block)
        self._last_closed_block = block

    def write_text(self, text: str) -> None:
        html_block = (
            self._last_opened_block
            if self._last_opened_block is not None and self._last_opened_block.type == BlockType.HTML_TAG
            else None
        )
        if self._is_in_html and html_block is not None:
            if html_block.tag_name in ("tr", "table"):
                if text and text.strip():
                    self._builder.write(text)
            elif html_block.tag_name == "td":
                if self._is_in_html_cell:
                    self._html_builder.append(text)
                elif text and text.strip():
                    self._builder.write(text)
            else:
                self._html_builder.append(text)
        else:
            self._builder.write(text)

    def insert_image(self, block: LinkTextBlock) -> None:
        """Embed a real Shape for a local, readable image or an inline base64 data URI;
        otherwise fall back to raw source text. Data URIs carry their own bytes right in
        the source text, so decoding one is a pure local operation -- no network access,
        same trust boundary as a local file already read from disk."""
        loaded = None
        if block.uri:
            loaded = load_local_image(block.uri, self._base_dir)
            if loaded is None:
                loaded = load_data_uri_image(block.uri)
        if loaded is None:
            self._builder.write(block.raw_text)
            return
        data, image_type, width, height = loaded
        alt_text = _flatten_text(block)
        image_data = ldm.ImageData(source_full_name=alt_text, image_type=image_type, image_bytes=data)
        self._builder.insert_image(image_data, width=width, height=height, name=alt_text)

    # ------------------------------------------------------------------
    # Leaf-block close / document close
    # ------------------------------------------------------------------

    def _close_leaf_block(self, block: Block) -> None:
        if block.type == BlockType.HORIZONTAL_RULE and not self._builder.current_paragraph._children:
            self._builder.insert_horizontal_rule()
        if self._last_opened_block is not None:
            html_block = self._last_opened_block if self._last_opened_block.type == BlockType.HTML_TAG else None
            if html_block is None or not html_block.is_para_or_heading:
                self._apply_formatting(block)
        self._reset_char_style()
        self._start_new_para()

    def _close_document(self) -> None:
        if self._html_builder:
            self._flush_html_builder()

        body_children = self._builder.document.sections[-1].body.children
        last = body_children[-1] if body_children else None
        if isinstance(last, ldm.Paragraph) and not last._children:
            prev = body_children[-2] if len(body_children) > 1 else None
            if isinstance(prev, ldm.Paragraph):
                self._builder.remove_body_child(last)

        self._builder.flush_footnotes(_FOOTNOTE_TEXT_STYLE_NAME)

    def _close_row(self, block: RowBlock) -> None:
        table = block.get_parent(BlockType.TABLE)
        if table is not None:
            for i in range(block.count, table.columns_count):
                cell = self._builder.insert_cell()
                alignment = table.column_alignments[i] if i < len(table.column_alignments) else None
                self._apply_cell_alignment(cell, alignment)
        self._builder.end_row()

    def _start_new_para(self) -> None:
        if self._footnote is None:
            self._builder.writeln()

    # ------------------------------------------------------------------
    # Raw HTML state management
    # ------------------------------------------------------------------

    def _open_html_tag(self, block: HtmlTagBlock) -> None:
        # open()+close() fire back-to-back on the same block; handle its whole raw_text here, once.
        if block.is_self_contained:
            if not self._is_in_html:
                match = self._match_self_contained_html(block)
                if match is not None:
                    self._open_self_contained_html(block, match)
                    return
                self._builder.insert_html(block.raw_text, HtmlInsertOptions.REMOVE_LAST_EMPTY_PARAGRAPH)
                self._start_new_para()
                return
            self._html_builder.append(block.raw_text)
            return

        if self._is_in_html or not self._open_html_inline(block):
            tag_name = block.tag_name
            if block.is_known:
                if not block.is_self_closing:
                    self._inc_html_tag_starts(tag_name)
                self._html_builder.append(block.raw_text)
            else:
                text = block.raw_text.replace("\n", " ")
                self._builder.write(text)

    def _close_html_tag(self, block: HtmlTagBlock) -> None:
        if block.is_self_contained:
            return

        if self._is_in_html:
            self._html_builder.append(block.raw_text)
        elif not self._close_html_inline(block):
            self._builder.write(block.raw_text)

        self._dec_html_tag_starts(block.tag_name)

        if self._html_tag_starts_count == 0 and self._html_builder:
            html_text = "".join(self._html_builder)
            self._html_builder.clear()
            self._builder.insert_html(html_text, HtmlInsertOptions.REMOVE_LAST_EMPTY_PARAGRAPH)

    @staticmethod
    def _match_self_contained_html(block: HtmlTagBlock):
        if not block.is_known or block.is_self_closing or block.is_closing:
            return None
        return _SELF_CONTAINED_HTML_RE.match(block.raw_text.strip())

    def _open_self_contained_html(self, block: HtmlTagBlock, match) -> None:
        inner = match.group(2)
        if block.tag_name in _INLINE_HTML_HEADING_TAGS and "<" not in inner:
            # Simple case: apply heading/paragraph style, collapse whitespace.
            self._apply_formatting(block)
            self._builder.write(" ".join(inner.split()))
            self._start_new_para()
        else:
            # insert_html() never starts a fresh paragraph on its own.
            self._builder.insert_html(block.raw_text, HtmlInsertOptions.REMOVE_LAST_EMPTY_PARAGRAPH)
            self._start_new_para()

    def _flush_html_builder(self) -> None:
        html_text = "".join(self._html_builder)
        self._html_builder.clear()
        self._builder.insert_html(html_text)

    def _open_html_inline(self, block: HtmlTagBlock) -> bool:
        tag = block.tag_name

        if tag in _INLINE_HTML_HEADING_TAGS:
            if not block.is_first_child:
                last_closed = self._last_closed_block
                if last_closed is None or not (
                    last_closed.type == BlockType.HTML_TAG and last_closed.is_para_or_heading
                ):
                    self._apply_formatting(self._last_opened_block or block)
                    in_cell = (
                        self._last_opened_block is not None
                        and self._last_opened_block.type == BlockType.HTML_TAG
                        and self._last_opened_block.is_cell
                    )
                    if not in_cell:
                        self._start_new_para()
            self._apply_formatting(block)
            return True

        if tag == "i":
            self._builder.italic = True
            self._italic_count += 1
            return True
        if tag == "sup":
            self._builder.font.superscript = True
            self._superscript_count += 1
            return True
        if tag == "sub":
            self._builder.font.subscript = True
            self._subscript_count += 1
            return True
        if tag in ("strong", "b"):
            self._builder.bold = True
            self._bold_count += 1
            return True
        if tag == "u":
            self._builder.font.underline = Underline.SINGLE
            self._underline_count += 1
            return True
        if tag in ("strike", "del", "s"):
            self._builder.font.strike_through = True
            self._strikethrough_count += 1
            return True
        if tag == "br":
            self._apply_formatting(block)
            self._start_new_para()
            return True
        if tag == "code":
            style = self._fetch_style(_INLINE_CODE_STYLE_NAME, StyleType.CHARACTER)
            self._decorate_inline_code(style)
            self._builder.font.style = style
            return True
        return False

    def _close_html_inline(self, block: HtmlTagBlock) -> bool:
        tag = block.tag_name

        if tag in _INLINE_HTML_HEADING_TAGS:
            if not block.is_last_child:
                self._start_new_para()
            return True
        if tag == "i":
            self._italic_count -= 1
            self._builder.italic = self._italic_count > 0
            return True
        if tag in ("b", "strong"):
            self._bold_count -= 1
            self._builder.bold = self._bold_count > 0
            return True
        if tag == "sup":
            self._superscript_count -= 1
            self._builder.font.superscript = self._superscript_count > 0
            return True
        if tag == "sub":
            self._subscript_count -= 1
            self._builder.font.subscript = self._subscript_count > 0
            return True
        if tag == "u":
            self._underline_count -= 1
            if self._underline_count == 0:
                self._builder.font.underline = Underline.NONE
            return True
        if tag in ("strike", "del", "s"):
            self._strikethrough_count -= 1
            self._builder.font.strike_through = self._strikethrough_count > 0
            return True
        if tag == "br":
            self._apply_formatting(block)
            self._start_new_para()
            return True
        if tag == "code":
            self._reset_char_style()
            return True
        return False

    def _inc_html_tag_starts(self, name: str) -> None:
        self._html_tag_starts[name] = self._html_tag_starts.get(name, 0) + 1

    def _dec_html_tag_starts(self, name: str) -> None:
        count = self._html_tag_starts.get(name, 0)
        if count <= 0:
            return
        self._html_tag_starts[name] = count - 1

    @property
    def _html_tag_starts_count(self) -> int:
        return sum(self._html_tag_starts.values())

    @property
    def _is_in_html(self) -> bool:
        return self._html_tag_starts_count > 0

    @property
    def _is_in_html_cell(self) -> bool:
        start = self._last_opened_block
        if start is None or start.type != BlockType.HTML_TAG or start.tag_name != "td":
            return False
        end = self._last_closed_block
        if end is None or end.type != BlockType.HTML_TAG or end.tag_name != "td":
            return True
        node = start.next_sibling
        while node is not None:
            if node is end:
                return False
            node = node.next_sibling
        return True

    # ------------------------------------------------------------------
    # Footnotes
    # ------------------------------------------------------------------

    def _footnote_reference_mark(self) -> str:
        """``MarkdownFootnoteWriter.GetReference``: the mark of footnote N."""
        return (
            f"{_FOOTNOTE_REFERENCE_OPENING}{self._footnote_counter}"
            f"{_FOOTNOTE_REFERENCE_CLOSING}"
        )

    def _open_footnote_reference(self, block: Block) -> None:
        self._footnote_counter += 1
        ref_style = self._fetch_style(_FOOTNOTE_REFERENCE_STYLE_NAME, StyleType.CHARACTER)

        saved_superscript = self._builder.font.superscript
        self._builder.font.style = ref_style
        self._builder.font.superscript = True
        self._builder.write(self._footnote_reference_mark())
        self._builder.font.superscript = saved_superscript
        self._reset_char_style()

        parent_paragraph = self._builder.insert_footnote()
        self._footnote = parent_paragraph
        self._save_current_formatting()
        self._clear_formatting()
        self._write_footnote_definition_mark()

    def _close_footnote_reference(self) -> None:
        assert self._footnote is not None
        self._builder.move_to(self._footnote)
        self._footnote = None
        self._restore_formatting()

    def _write_footnote_definition_mark(self) -> None:
        """Open the body with "<mark>: ", the way ``WriteDefinitions`` does.

        The LDM has no footnote node to carry the mark, so it lives in the
        body's first line, under the character style the reference mark
        carries -- that is what tells the writer it is a mark and not text
        that merely looks like one.  An empty definition keeps it too, or it
        would leave nothing behind for its reference to find.
        """
        self._builder.font.style = self._fetch_style(
            _FOOTNOTE_REFERENCE_STYLE_NAME, StyleType.CHARACTER
        )
        self._builder.write(f"{self._footnote_reference_mark()}: ")
        self._reset_char_style()

    def _start_footnote_paragraph(self) -> None:
        body = self._builder.current_footnote_body()
        reuse = len(body) == 1 and len(body[0]._children) <= 1
        footnote_para = body[0] if reuse else self._builder.new_paragraph()
        self._builder.current_paragraph = footnote_para

    def _save_current_formatting(self) -> None:
        self._builder.push_font()
        self._builder.push_para_pr()
        self._saved_bold_count = self._bold_count
        self._saved_italic_count = self._italic_count
        self._saved_strikethrough_count = self._strikethrough_count
        self._saved_underline_count = self._underline_count
        self._saved_superscript_count = self._superscript_count
        self._saved_subscript_count = self._subscript_count

    def _restore_formatting(self) -> None:
        self._builder.pop_font()
        self._builder.pop_para_pr()
        self._bold_count = self._saved_bold_count
        self._italic_count = self._saved_italic_count
        self._strikethrough_count = self._saved_strikethrough_count
        self._underline_count = self._saved_underline_count
        self._superscript_count = self._saved_superscript_count
        self._subscript_count = self._saved_subscript_count

    def _clear_formatting(self) -> None:
        self._builder.clear_font()
        self._builder.current_paragraph.paragraph_format = ldm.ParagraphFormat()
        self._bold_count = 0
        self._italic_count = 0
        self._strikethrough_count = 0
        self._underline_count = 0

    # ------------------------------------------------------------------
    # Formatting application
    # ------------------------------------------------------------------

    def _apply_formatting(self, block: Block) -> None:
        style = self._get_style(block)
        para = self._builder.current_paragraph
        para.paragraph_format.style_name = style.name
        para.paragraph_format.style_identifier = style.style_identifier
        self._apply_heading_meta(para.paragraph_format, block)
        self._apply_list(block)

    def _set_paragraph_style(self, style: ldm.Style) -> None:
        para = self._builder.current_paragraph
        para.paragraph_format.style_name = style.name
        para.paragraph_format.style_identifier = style.style_identifier

    @staticmethod
    def _apply_heading_meta(pf: ldm.ParagraphFormat, block: Block) -> None:
        level = MarkdownReaderContext._heading_level(block)
        pf.is_heading = level is not None
        if level is not None:
            pf.outline_level = level - 1

    @staticmethod
    def _heading_level(block: Block) -> Optional[int]:
        if block.type in (BlockType.ATX_HEADING, BlockType.SETEXT_HEADING):
            return block.level
        if block.type == BlockType.HTML_TAG and block.heading_tag_level != -1:
            return block.heading_tag_level
        return None

    _CELL_ALIGNMENT = {
        "left": ParagraphAlignment.LEFT,
        "center": ParagraphAlignment.CENTER,
        "right": ParagraphAlignment.RIGHT,
    }

    def _apply_cell_alignment(self, cell: ldm.Cell, alignment: Optional[str]) -> None:
        if not cell.paragraphs:
            return
        cell.paragraphs[0].paragraph_format.alignment = self._CELL_ALIGNMENT.get(alignment or "", ParagraphAlignment.LEFT)
