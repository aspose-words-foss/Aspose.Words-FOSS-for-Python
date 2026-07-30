"""
Minimal LDM-facing analogue of Aspose.Words' ``DocumentBuilder``.

``MarkdownReaderContext`` (the next module) drives the reference
``MarkdownReaderContext`` algorithm against a real ``DocumentBuilder`` /
``Document`` (with ``Story``, ``ParaPr``/``RunPr`` istd chains, etc.).  The
light document model has none of that machinery — paragraphs live in flat
``list[Paragraph]`` containers and styles are plain named records — so this
class provides just the surface ``MarkdownReaderContext`` actually calls:
a cursor (``current_paragraph``) into one of those containers, toggleable
bold/italic/font state, run writing with automatic merging, table/hyperlink/
footnote cursor management, and formatting push/pop.
"""

from __future__ import annotations

import html as _html_module
import re
from typing import Optional

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss._visible_runs import HORIZONTAL_RULE_SHAPE_TYPE


# MarkdownUtil.EscapeMarkupSymbols' symbol set, and gSquareBrackets.
_MARKUP_SYMBOLS_FOR_ESCAPING = ("(", ")", "<", ">")
_SQUARE_BRACKETS = ("[", "]")
# Shape.CreateHorizontalRule: 1.5pt tall, gray, page content width (6in
# for the default page setup the markdown reader leaves untouched).
_HORIZONTAL_RULE_WIDTH = 432.0
_HORIZONTAL_RULE_HEIGHT = 1.5
_HORIZONTAL_RULE_COLOR = "808080"


def _escape_markup_symbols(text: str) -> str:
    for symbol in _MARKUP_SYMBOLS_FOR_ESCAPING:
        text = text.replace(symbol, "\\" + symbol)
    return text


_UNESCAPED_SQUARE_BRACKET_RE = re.compile(r"(?<!\\)[" + re.escape("".join(_SQUARE_BRACKETS)) + r"]")


def _unescape_markup_symbols(text: str) -> str:
    """``MarkdownUtil.UnescapeMarkupSymbols``: the inverse of the above."""
    for symbol in _MARKUP_SYMBOLS_FOR_ESCAPING:
        text = text.replace("\\" + symbol, symbol)
    return text


def _escape_square_brackets(text: str) -> str:
    """MarkdownUtil.EscapeSquareBrackets: a bracket in a link label would end
    the label where it stands, so the rest of the link stops being one.  One
    that already carries a backslash is left alone -- escaping it again would
    add a backslash on every export."""
    return _UNESCAPED_SQUARE_BRACKET_RE.sub(r"\\\g<0>", text)


def _link_destination(uri: str) -> str:
    """Render *uri* so the whole of it reads back as the destination.

    Escaping alone cannot carry a space -- the destination would end at it --
    so a destination holding whitespace goes inside angle brackets instead.
    """
    if not uri or not re.search(r"\s", uri):
        return _escape_markup_symbols(uri)
    # Only the brackets themselves need escaping in here; escaping the
    # backslashes too would double them on every export.
    return "<" + uri.replace("<", "\\<").replace(">", "\\>") + ">"


def _link_title(title: str) -> str:
    if not title:
        return ""
    escaped = title.replace('\\"', '"').replace('"', '\\"')
    return f' "{escaped}"'


class HtmlInsertOptions:
    """Flags for :meth:`MarkdownDocumentBuilder.insert_html` (mirrors
    Aspose.Words' ``HtmlInsertOptions``)."""

    NONE = 0
    REMOVE_LAST_EMPTY_PARAGRAPH = 1


class _FontState:
    """Mutable current-font state (Aspose.Words' ``DocumentBuilder.Font``)."""

    __slots__ = (
        "strike_through", "underline", "superscript", "subscript",
        "name", "size", "color", "highlight_color",
        "style_name", "style_identifier",
    )

    def __init__(self) -> None:
        self.clear()

    def clear(self) -> None:
        self.strike_through = False
        self.underline = 0
        self.superscript = False
        self.subscript = False
        self.name = ""
        self.size = 0.0
        self.color = ""
        self.highlight_color = ""
        self.style_name = ""
        self.style_identifier = 0

    def apply_style(self, style: ldm.Style) -> None:
        """Merge a character style's template font onto the current state
        (Aspose.Words' ``Font.Style = someCharacterStyle``)."""
        self.style_name = style.name
        self.style_identifier = style.style_identifier
        if style.font is None:
            return
        f = style.font
        if f.name:
            self.name = f.name
        if f.size:
            self.size = f.size
        if f.color:
            self.color = f.color
        if f.highlight_color:
            self.highlight_color = f.highlight_color
        if f.underline:
            self.underline = f.underline
        self.strike_through = self.strike_through or f.strike_through

    def snapshot(self) -> tuple:
        return (
            self.strike_through, self.underline, self.superscript, self.subscript,
            self.name, self.size, self.color, self.highlight_color,
            self.style_name, self.style_identifier,
        )

    def restore(self, snap: tuple) -> None:
        (
            self.strike_through, self.underline, self.superscript, self.subscript,
            self.name, self.size, self.color, self.highlight_color,
            self.style_name, self.style_identifier,
        ) = snap


class _FontProxy:
    """``builder.font.strike_through = True`` façade over ``_FontState``,
    matching the ``mBuilder.Font.X`` call sites in the reference."""

    def __init__(self, state: _FontState) -> None:
        self._state = state

    @property
    def strike_through(self) -> bool:
        return self._state.strike_through

    @strike_through.setter
    def strike_through(self, value: bool) -> None:
        self._state.strike_through = value

    @property
    def underline(self) -> int:
        return self._state.underline

    @underline.setter
    def underline(self, value: int) -> None:
        self._state.underline = value

    @property
    def superscript(self) -> bool:
        return self._state.superscript

    @superscript.setter
    def superscript(self, value: bool) -> None:
        self._state.superscript = value

    @property
    def subscript(self) -> bool:
        return self._state.subscript

    @subscript.setter
    def subscript(self, value: bool) -> None:
        self._state.subscript = value

    @property
    def style(self) -> None:  # write-only, matches Aspose's Font.Style setter usage here
        raise AttributeError("Font.style is write-only")

    @style.setter
    def style(self, style: ldm.Style) -> None:
        self._state.apply_style(style)

    def reset_style(self) -> None:
        """``Font.Style = DefaultParagraphFont`` — drop the character-style
        overlay while keeping bold/italic/strike/underline toggles intact."""
        self._state.style_name = ""
        self._state.style_identifier = 0
        self._state.name = ""
        self._state.size = 0.0
        self._state.color = ""
        self._state.highlight_color = ""


class _HyperlinkCapture:
    __slots__ = ("uri", "title", "buffer")

    def __init__(self, uri: str, title: str = "") -> None:
        self.uri = uri
        self.title = title
        self.buffer: list[str] = []


class _TableState:
    __slots__ = ("table", "row", "in_cell")

    def __init__(self, table: ldm.Table) -> None:
        self.table = table
        self.row: Optional[ldm.Row] = None
        self.in_cell = False


class MarkdownDocumentBuilder:
    """LDM cursor/writer used by :class:`MarkdownReaderContext`.

    Owns a fresh single-section :class:`ldm.Document` and a "current
    paragraph" cursor that write/writeln/table/footnote operations move
    around, the same role Aspose.Words' ``DocumentBuilder`` plays for the
    real reader.
    """

    def __init__(self) -> None:
        self.document = ldm.Document()
        section = ldm.Section()
        self.document.sections = [section]
        self._body_children: list = section.body.children

        self.bold = False
        self.italic = False
        self.font = _FontProxy(_FontState())

        self._container_stack: list[list] = [self._body_children]
        self.current_paragraph: ldm.Paragraph = self._append_new_paragraph()

        self._table_stack: list[_TableState] = []
        self._hyperlink_stack: list[_HyperlinkCapture] = []
        self._footnote_bodies: list[list[ldm.Paragraph]] = []

        self._font_stack: list[tuple] = []
        self._para_pr_stack: list[ldm.ParagraphFormat] = []

    # ------------------------------------------------------------------
    # Cursor / container plumbing
    # ------------------------------------------------------------------

    def _append_new_paragraph(self, container: Optional[list] = None) -> ldm.Paragraph:
        container = self._container_stack[-1] if container is None else container
        para = ldm.Paragraph()
        container.append(para)
        return para

    def push_container(self, container: list) -> None:
        self._container_stack.append(container)

    def pop_container(self) -> None:
        self._container_stack.pop()

    def new_paragraph(self) -> ldm.Paragraph:
        """Append a fresh paragraph to the current container without
        moving the cursor there (used by footnote-body management)."""
        return self._append_new_paragraph()

    @property
    def last_section_last_child(self):
        """``Document.LastSection.Body.LastChild`` — the last top-level body node."""
        return self._body_children[-1] if self._body_children else None

    def remove_body_child(self, node) -> None:
        # Identity, not `list.remove()`'s equality: two blank paragraphs
        # compare equal under Pydantic's field-based __eq__ (``_children``
        # is a PrivateAttr and isn't part of that comparison), so
        # equality-based removal can delete the wrong — structurally
        # identical but semantically different — paragraph.
        for i, child in enumerate(self._body_children):
            if child is node:
                del self._body_children[i]
                return

    # ------------------------------------------------------------------
    # Writing
    # ------------------------------------------------------------------

    def write(self, text: str) -> None:
        if not text:
            return
        if self._hyperlink_stack:
            self._hyperlink_stack[-1].buffer.append(_escape_square_brackets(text))
            return
        font = self._snapshot_font()
        children = self.current_paragraph._children
        if children and isinstance(children[-1], ldm.Run) and children[-1].font == font:
            children[-1].text += text
            return
        children.append(ldm.Run(text=text, font=font))

    def writeln(self) -> None:
        """Insert a paragraph break: finish the current paragraph and move
        the cursor to a fresh one in the same container."""
        self.current_paragraph = self._append_new_paragraph()

    def insert_image(
        self,
        image_data: ldm.ImageData,
        width: Optional[float] = None,
        height: Optional[float] = None,
        name: str = "",
    ) -> Optional[ldm.Shape]:
        """Insert an inline image Shape (Aspose.Words' ``DocumentBuilder.InsertImage``)."""
        if self._hyperlink_stack:
            self._hyperlink_stack[-1].buffer.append(name)  # no way to nest a real Shape in literal hyperlink text
            return None
        shape = ldm.Shape(
            shape_type=75,  # ShapeType.Image
            name=name,
            width=width,
            height=height,
            is_inline=True,
            has_image=True,
            image_data=image_data,
        )
        self.current_paragraph._children.append(shape)
        return shape

    def _snapshot_font(self) -> ldm.Font:
        state = self.font._state
        return ldm.Font(
            bold=self.bold,
            italic=self.italic,
            strike_through=state.strike_through,
            underline=state.underline,
            superscript=state.superscript,
            subscript=state.subscript,
            name=state.name,
            size=state.size,
            color=state.color,
            highlight_color=state.highlight_color,
            style_name=state.style_name,
            style_identifier=state.style_identifier,
        )

    # ------------------------------------------------------------------
    # Formatting stacks
    # ------------------------------------------------------------------

    def push_font(self) -> None:
        self._font_stack.append((self.bold, self.italic, self.font._state.snapshot()))

    def pop_font(self) -> None:
        self.bold, self.italic, snap = self._font_stack.pop()
        self.font._state.restore(snap)

    def clear_font(self) -> None:
        self.bold = False
        self.italic = False
        self.font._state.clear()

    def push_para_pr(self) -> None:
        self._para_pr_stack.append(self.current_paragraph.paragraph_format.model_copy(deep=True))

    def pop_para_pr(self) -> None:
        self.current_paragraph.paragraph_format = self._para_pr_stack.pop()

    # ------------------------------------------------------------------
    # Tables
    # ------------------------------------------------------------------

    def insert_cell(self) -> ldm.Cell:
        state = self._table_stack[-1] if self._table_stack else None
        if state is not None and state.in_cell:
            self.pop_container()
            state.in_cell = False
        if state is None:
            table = ldm.Table()
            self._container_stack[-1].append(table)
            state = _TableState(table)
            self._table_stack.append(state)
        if state.row is None:
            state.row = ldm.Row()
            state.table.rows.append(state.row)
        cell = ldm.Cell()
        state.row.cells.append(cell)
        self.push_container(cell.paragraphs)
        state.in_cell = True
        self.current_paragraph = self._append_new_paragraph()
        return cell

    def end_row(self) -> None:
        if not self._table_stack:
            return
        state = self._table_stack[-1]
        if state.in_cell:
            self.pop_container()
            state.in_cell = False
        state.row = None

    def end_table(self) -> None:
        if self._table_stack:
            self._table_stack.pop()
        self.current_paragraph = self._append_new_paragraph()

    # ------------------------------------------------------------------
    # Hyperlinks
    #
    # The LDM has no field-code hyperlink node, so — matching the
    # convention already used by the DOCX reader (see
    # ``ldm_builder/paragraphs.py:_append_link_run``) — a hyperlink is
    # stored as a single literal ``[text](uri)`` Run.
    # ------------------------------------------------------------------

    def start_hyperlink(self, uri: str = "", title: str = "") -> None:
        self._hyperlink_stack.append(_HyperlinkCapture(uri, title))

    def insert_horizontal_rule(self) -> None:
        """``DocumentBuilder.InsertHorizontalRule``."""
        self.current_paragraph._children.append(
            ldm.Shape(
                shape_type=HORIZONTAL_RULE_SHAPE_TYPE,
                width=_HORIZONTAL_RULE_WIDTH,
                height=_HORIZONTAL_RULE_HEIGHT,
                is_inline=True,
                fill_color=_HORIZONTAL_RULE_COLOR,
            )
        )

    @property
    def is_capturing_hyperlink(self) -> bool:
        """Whether text written now lands in a hyperlink label, not in a run."""
        return bool(self._hyperlink_stack)

    def end_hyperlink(self) -> None:
        capture = self._hyperlink_stack.pop()
        label = "".join(capture.buffer)
        display = f"[{label}]({_link_destination(capture.uri)}{_link_title(capture.title)})"
        if self._hyperlink_stack:
            self._hyperlink_stack[-1].buffer.append(display)
            return
        font = self._snapshot_font()
        children = self.current_paragraph._children
        if children and isinstance(children[-1], ldm.Run) and children[-1].font == font:
            children[-1].text += display
        else:
            children.append(ldm.Run(text=display, font=font))

    # ------------------------------------------------------------------
    # Footnotes
    #
    # The LDM has no footnote node either (see the "REMOVED entirely:
    # CommentNode, FootnoteNode" note in light_document_model.py), so a
    # footnote body is built as an ordinary side list of paragraphs and
    # spliced onto the end of the document body once parsing finishes
    # (``flush_footnotes``), the same way many single-file converters
    # render footnotes as trailing endnote-style paragraphs.
    # ------------------------------------------------------------------

    def insert_footnote(self) -> ldm.Paragraph:
        """Start a footnote body; returns the paragraph the reference to
        it should resume at once the footnote is closed."""
        parent_paragraph = self.current_paragraph
        body: list[ldm.Paragraph] = []
        self._footnote_bodies.append(body)
        self.push_container(body)
        self.current_paragraph = self._append_new_paragraph()
        return parent_paragraph

    def current_footnote_body(self) -> list[ldm.Paragraph]:
        return self._footnote_bodies[-1]

    def move_to(self, paragraph: ldm.Paragraph) -> None:
        """Return the cursor to *paragraph* after finishing a footnote body."""
        self.pop_container()
        self.current_paragraph = paragraph

    def flush_footnotes(self, style_name: str) -> None:
        """Append every footnote body collected so far to the end of the
        document body, tagging each paragraph with *style_name*."""
        for body in self._footnote_bodies:
            for para in body:
                if not para.paragraph_format.style_name:
                    para.paragraph_format.style_name = style_name
                self._body_children.append(para)
        self._footnote_bodies.clear()

    # ------------------------------------------------------------------
    # Raw HTML
    #
    # A minimal, tag-stripping conversion: block-level tags become
    # paragraph breaks, everything else is written as plain text. Full
    # tag-aware conversion (tables, nested inline formatting) is a later
    # step's job — this is enough to not silently drop raw HTML content.
    # ------------------------------------------------------------------

    _HTML_TAG_RE_SOURCE = r"<[^>]+>"

    def insert_html(self, html: str, options: int = HtmlInsertOptions.NONE) -> None:
        if not html:
            return
        blocks = re.split(r"(?i)</?(?:p|div|table|tr|td|th|h[1-6]|br)\s*/?>", html)
        wrote_any = False
        for block in blocks:
            text = re.sub(self._HTML_TAG_RE_SOURCE, "", block)
            text = _html_module.unescape(text).strip()
            if not text:
                continue
            if wrote_any:
                self.writeln()
            self.write(text)
            wrote_any = True
        if wrote_any and options & HtmlInsertOptions.REMOVE_LAST_EMPTY_PARAGRAPH:
            last = self.last_section_last_child
            if isinstance(last, ldm.Paragraph) and not last._children:
                self.remove_body_child(last)
                self.current_paragraph = self._append_new_paragraph()
