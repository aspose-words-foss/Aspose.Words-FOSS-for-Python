"""Paragraph rendering for the PDF writer.

Contains the unified styled-block rendering logic used by both the
legacy path and the content-sequence (mixed image+text) path,
eliminating the previous code duplication.
"""


import warnings
from typing import Optional

from fpdf import FPDF

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss._visible_runs import visible_runs
from aspose.words_foss.model.wrap_type import WrapType
from aspose.words_foss.pdf_writer.constants import (
    CODE_BLOCK_BG_RGB,
    DEFAULT_FONT_NAME,
    DEFAULT_FONT_SIZE_PT,
    DEFAULT_QUOTE_INDENT_MM,
    FPDF_ALIGN,
    LINE_HEIGHT_FACTOR,
    LIST_INDENT_PER_LEVEL_MM,
    PT_TO_MM,
    QUOTE_TEXT_RGB,
    WORD_LINE_SPACING_DIVISOR,
)
from aspose.words_foss.pdf_writer._context import PDFWriterContext
from aspose.words_foss.pdf_writer.font import apply_run_font, reset_font
from aspose.words_foss.pdf_writer.page_bands import register_bookmarks
from aspose.words_foss.model.enums import LineSpacingRule
from aspose.words_foss.pdf_writer.text import (
    apply_caps,
    extract_link_segments,
    get_dominant_color,
    get_dominant_font_size,
    is_pure_page_break,
    is_toc_style,
    safe_text,
)


# Mirror of ``NumberStyle``; unknown styles fall back to decimal.
def _format_number(num: int, number_style: int) -> str:
    if number_style == 0:  # ARABIC
        return str(num)
    if number_style == 1:  # UPPERCASE_ROMAN
        return _to_roman(num).upper()
    if number_style == 2:  # LOWERCASE_ROMAN
        return _to_roman(num).lower()
    if number_style == 3:  # UPPERCASE_LETTER
        return _to_alpha(num).upper()
    if number_style == 4:  # LOWERCASE_LETTER
        return _to_alpha(num).lower()
    return str(num)


def _to_roman(num: int) -> str:
    if num <= 0:
        return ""
    pairs = (
        (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
        (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
        (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"),
    )
    out = []
    for value, sym in pairs:
        while num >= value:
            out.append(sym)
            num -= value
    return "".join(out)


def _to_alpha(num: int) -> str:
    if num <= 0:
        return ""
    out = ""
    while num > 0:
        num, rem = divmod(num - 1, 26)
        out = chr(ord("a") + rem) + out
    return out


def _expand_number_format(
    fmt: str,
    lst: ldm.DocList,
    counters: dict[tuple[int, int], int],
    current_level: int,
    current_style: int,
) -> str:
    """Replace ``%N`` placeholders with the counter for level ``N-1``."""
    out = []
    i = 0
    while i < len(fmt):
        ch = fmt[i]
        if ch == "%" and i + 1 < len(fmt) and fmt[i + 1].isdigit():
            lvl_idx = int(fmt[i + 1]) - 1
            if 0 <= lvl_idx < len(lst.levels):
                style = lst.levels[lvl_idx].number_style
            else:
                style = current_style
            counter = counters.get((lst.list_id, lvl_idx), 0)
            if counter == 0 and lvl_idx == current_level:
                counter = 1
            out.append(_format_number(counter, style))
            i += 2
            continue
        out.append(ch)
        i += 1
    return "".join(out)


class ParagraphRenderer:
    """Renders LDM paragraphs into PDF."""

    def __init__(self, writer: PDFWriterContext) -> None:
        self._writer = writer
        self._warned_expanded: bool = False

    # ------------------------------------------------------------------
    # Line height computation
    # ------------------------------------------------------------------

    @staticmethod
    def line_height_mm(size_pt: float, pf: ldm.ParagraphFormat) -> float:
        """Return the line height in mm for a run of *size_pt* inside *pf*.

        Honours Word's ``w:spacing w:lineRule="exact"`` (``rule==1``), which
        pins the line advance to a fixed point value.
        ``atLeast`` (0) acts as a floor above the natural leading;
        ``multiple`` (2) scales the natural leading by ``line_spacing/240``.
        """

        natural = size_pt * PT_TO_MM * LINE_HEIGHT_FACTOR
        ls = pf.line_spacing
        if ls <= 0:
            return natural
        if pf.line_spacing_rule == LineSpacingRule.EXACTLY:
            return ls * PT_TO_MM
        if pf.line_spacing_rule == LineSpacingRule.AT_LEAST:
            return max(natural, ls * PT_TO_MM)
        if pf.line_spacing_rule == LineSpacingRule.MULTIPLE:
            return natural * (ls / WORD_LINE_SPACING_DIVISOR) if ls else natural
        return natural

    # ------------------------------------------------------------------
    # Public entry point
    # ------------------------------------------------------------------

    def render_paragraph(self, pdf: FPDF, para: ldm.Paragraph) -> None:
        """Render a single paragraph into the PDF."""
        if is_pure_page_break(para):
            pdf.add_page()
            return

        if para.paragraph_format.keep_together:
            self._maybe_keep_together(pdf, para)

        # Capture the paragraph's top Y before body rendering — needed
        # to position anchored shapes with
        # ``relative_vertical_position == Paragraph`` ().
        paragraph_top_y = pdf.get_y()

        register_bookmarks(pdf, para, self._writer._anchor_links)
        self._emit_bookmark_outlines(pdf, para)
        self._render_paragraph_body(pdf, para)

        self._render_paragraph_relative_shapes(pdf, para, paragraph_top_y)

        if any("\f" in (run.text or "") for run in visible_runs(para)):
            pdf.add_page()

    def _render_paragraph_relative_shapes(
        self, pdf: FPDF, para: ldm.Paragraph, paragraph_top_y: float
    ) -> None:
        """Draw shapes anchored ``relative_vertical_position == Paragraph``.

        The shape's ``vertical_position`` field carries the raw OOXML
        ``<wp:positionV/posOffset>`` in points; convert to mm and add
        the paragraph's actual top Y so the shape lands where Word
        places it (e.g. an absolutely-positioned date stamp
        18.81 cm below the cover-page paragraph).
        """
        extras = [
            e for e in para._children
            if isinstance(e, ldm.ShapeNode) and e._is_positioned
            and e.relative_vertical_position == 2
        ]
        if not extras:
            return
        saved_x, saved_y = pdf.get_x(), pdf.get_y()
        for shape in extras:
            y_mm = paragraph_top_y + shape.vertical_position * PT_TO_MM
            self._writer._shape_renderer.render_positioned_shape(
                pdf, shape, y_override=y_mm,
            )
        pdf.set_xy(saved_x, saved_y)

    # ------------------------------------------------------------------
    # Keep-together support
    # ------------------------------------------------------------------

    def _maybe_keep_together(self, pdf: FPDF, para: ldm.Paragraph) -> None:
        """If the paragraph won't fit on this page, move to the next one."""
        w = self._writer
        pf = para.paragraph_format
        runs = visible_runs(para)
        text = self._plain_text_from_runs(runs)
        if not text.strip():
            return
        fs = get_dominant_font_size(runs) or DEFAULT_FONT_SIZE_PT
        line_h = self.line_height_mm(fs, pf)
        usable_w = w._page_width - w._page_margin_left - w._page_margin_right
        char_w = fs * PT_TO_MM * 0.5
        chars_per_line = max(1, int(usable_w / char_w))
        num_lines = 0
        for line in text.split("\n") or [""]:
            num_lines += max(1, (len(line) // chars_per_line) + 1)
        est_height = num_lines * line_h + pf.space_before * PT_TO_MM + pf.space_after * PT_TO_MM
        remaining = w._page_height - w._page_margin_bottom - pdf.get_y()
        if est_height > remaining and est_height < (w._page_height - w._page_margin_bottom - pdf.t_margin):
            pdf.add_page()

    # ------------------------------------------------------------------
    # Outline helpers
    # ------------------------------------------------------------------

    def _effective_headings_outline_levels(self) -> int:
        """Return the effective headings outline depth."""
        oo = self._writer.options.outline_options
        if oo.headings_outline_levels > 0:
            return oo.headings_outline_levels
        if self._writer.options.export_document_structure:
            return 9
        return 0

    def _effective_bookmarks_outline_level(self) -> int:
        """Return the effective default bookmark outline level."""
        oo = self._writer.options.outline_options
        if oo.default_bookmarks_outline_level > 0 or oo.bookmarks_outline_levels:
            return oo.default_bookmarks_outline_level
        return 1 if self._writer.options.export_bookmarks_outline else 0

    def _start_section_safe(self, pdf: FPDF, text: str, level: int) -> None:
        """Call ``pdf.start_section`` filling any gaps fpdf2 would reject."""
        w = self._writer
        for gap in range(w._last_outline_level + 1, level):
            pdf.start_section("", level=gap - 1)
        w._last_outline_level = level
        pdf.start_section(text, level=level - 1)

    def _emit_heading_outline(self, pdf: FPDF, text: str, level: int) -> None:
        """Emit a PDF outline entry for a heading if allowed by OutlineOptions."""
        w = self._writer
        oo = w.options.outline_options
        # TODO: implement expanded_outline_levels via PDF /Count post-processing
        if oo.expanded_outline_levels > 0 and not self._warned_expanded:
            warnings.warn(
                "OutlineOptions.expanded_outline_levels is not yet implemented; "
                "the value will be ignored",
                stacklevel=2,
            )
            self._warned_expanded = True
        max_level = self._effective_headings_outline_levels()
        if level > max_level:
            return
        if w._in_table and not oo.create_outlines_for_headings_in_tables:
            return
        self._start_section_safe(pdf, text, level)

    def _emit_bookmark_outlines(self, pdf: FPDF, para: ldm.Paragraph) -> None:
        """Emit outline entries for bookmarks in this paragraph."""
        if para.paragraph_format.is_heading:
            return
        oo = self._writer.options.outline_options
        default_level = self._effective_bookmarks_outline_level()
        for extra in para._children:
            if not isinstance(extra, ldm.BookmarkStart) or not extra.name:
                continue
            if extra.name.startswith("_"):
                continue
            bm_level = oo.bookmarks_outline_levels.get(extra.name, default_level)
            if bm_level <= 0:
                continue
            self._start_section_safe(pdf, extra.name, bm_level)

    # ------------------------------------------------------------------
    # Body dispatcher
    # ------------------------------------------------------------------

    def _render_paragraph_body(self, pdf: FPDF, para: ldm.Paragraph) -> None:
        """Render paragraph body: floating images, then styled content."""
        w = self._writer

        # Draw floating (wrapNone) images at their anchor coordinates
        w._shape_renderer.render_floating_images(pdf, para)

        # Anchored (is_inline=False) wrapped images float — don't take inline space.
        w._shape_renderer.render_anchored_wrapped_shapes(pdf, para)

        # ``is_inline=None`` (untagged shapes from older fixtures) is treated as inline.
        has_mixed = any(
            isinstance(i, ldm.ShapeNode)
            and i.has_image
            and i.is_inline is not False
            and not i._is_positioned
            and i.wrap_type != WrapType.NONE
            for i in para._children
        )
        if has_mixed:
            self._render_content_sequence(pdf, para)
            return

        # Legacy path: no inline images — emit any standalone shapes first.
        for item in para._children:
            if isinstance(item, ldm.ShapeNode):
                if item._is_positioned or item.wrap_type == WrapType.NONE:
                    continue
                if item.is_inline is False:
                    # Anchored/floating; already drawn above.
                    continue
                if item.has_image and item.image_data is not None:
                    w._shape_renderer.render_shape(pdf, item)
                elif item.text_box:
                    w._shape_renderer.render_shape(pdf, item)

        pf = para.paragraph_format
        align = FPDF_ALIGN.get(pf.alignment, "L")

        # Ensure each paragraph starts at the left margin.
        pdf.set_x(w._page_margin_left)

        self._render_styled_block(
            pdf, visible_runs(para), pf, para.list_format, para.list_label, align
        )

    # ------------------------------------------------------------------
    # Unified styled-block rendering (DRY fix)
    # ------------------------------------------------------------------

    def _render_styled_block(
        self,
        pdf: FPDF,
        runs: list[ldm.Run],
        pf: ldm.ParagraphFormat,
        list_format: Optional[ldm.ListFormat],
        list_label: Optional[ldm.ListLabel],
        align: str,
    ) -> None:
        """Shared heading/code/quote/list/normal rendering logic.

        This single method replaces the duplicated branching that
        previously existed in both ``_render_paragraph_body`` and
        ``_render_content_sequence``'s ``_flush_styled`` closure.
        """
        w = self._writer
        fs = DEFAULT_FONT_SIZE_PT
        style_name = pf.style_name

        # Apply space before
        self._apply_space_before(pdf, pf)

        # Heading
        if pf.is_heading:
            level = min(pf.outline_level + 1, 6)
            run_size = get_dominant_font_size(runs)
            size = run_size if run_size > 0 else fs + (6 - level) * 2
            line_h = self.line_height_mm(size, pf)
            text = self._plain_text_from_runs(runs)
            if text.strip():
                self._emit_heading_outline(pdf, text.strip(), level)
            with w._tag(pdf, f"/H{level}", title=text.strip()):
                heading_color = get_dominant_color(runs)
                if runs:
                    apply_run_font(pdf, runs[0].font, default_size=size)
                else:
                    pdf.set_font(DEFAULT_FONT_NAME, style="B", size=size)
                if heading_color:
                    pdf.set_text_color(*heading_color)
                pdf.multi_cell(w=0, h=line_h, text=safe_text(text), align=align)
            self._apply_space_after(pdf, pf)
            reset_font(pdf)
            return

        # Code block
        is_code = bool(style_name and ("Code" in style_name or "code" in style_name))
        if is_code:
            code_size = fs - 1
            line_h = self.line_height_mm(code_size, pf)
            text = self._plain_text_from_runs(runs)
            with w._tag(pdf, "/Code"):
                pdf.set_fill_color(*CODE_BLOCK_BG_RGB)
                pdf.set_font("Courier", size=code_size)
                usable_w = w._page_width - w._page_margin_left - w._page_margin_right
                pdf.multi_cell(
                    w=usable_w,
                    h=line_h,
                    text=safe_text(text),
                    fill=True,
                    align=align,
                )
            self._apply_space_after(pdf, pf)
            reset_font(pdf)
            return

        # Block quote
        if style_name and "Quote" in style_name:
            line_h = self.line_height_mm(fs, pf)
            text = self._plain_text_from_runs(runs)
            with w._tag(pdf, "/BlockQuote"):
                pdf.set_font(DEFAULT_FONT_NAME, style="I", size=fs)
                pdf.set_text_color(*QUOTE_TEXT_RGB)
                indent = (
                    pf.left_indent * PT_TO_MM if pf.left_indent > 0 else DEFAULT_QUOTE_INDENT_MM
                )
                pdf.cell(w=indent)
                pdf.multi_cell(w=0, h=line_h, text=safe_text(text), align=align)
            self._apply_space_after(pdf, pf)
            reset_font(pdf)
            return

        # List item
        if list_format and list_format.is_list_item:
            level = list_format.list_level_number
            # Marker at ``npos`` (left_indent + first_line_indent), text at ``tpos`` (left_indent).
            text_indent_mm = (
                pf.left_indent * PT_TO_MM
                if pf.left_indent > 0
                else LIST_INDENT_PER_LEVEL_MM * (level + 1)
            )
            marker_indent_mm = max(0.0, text_indent_mm + pf.first_line_indent * PT_TO_MM)
            label = list_label.label_string if list_label and list_label.label_string else ""
            if not label:
                label = self._compute_list_label(list_format) or "-"
            with w._tag(pdf, "/LI"):
                run_size = get_dominant_font_size(runs)
                effective_fs = run_size if run_size > 0 else fs
                line_h = self.line_height_mm(effective_fs, pf)
                if marker_indent_mm > 0:
                    pdf.cell(w=marker_indent_mm)
                pdf.set_font(DEFAULT_FONT_NAME, size=effective_fs)
                pdf.write(h=line_h, text=safe_text(f"{label}"))
                pad = text_indent_mm - marker_indent_mm
                if pad > 0:
                    pdf.write(h=line_h, text=" ")
                    pdf.set_x(pdf.l_margin + text_indent_mm)
                w._run_renderer.render_formatted_runs(
                    pdf,
                    runs,
                    newline=True,
                    pf=pf,
                )
            self._apply_space_after(pdf, pf)
            return

        # Normal paragraph — render with inline formatting
        text = self._plain_text_from_runs(runs)
        if not text.strip():
            run_size = get_dominant_font_size(runs) or fs
            pdf.ln(self.line_height_mm(run_size, pf))
            return

        # Apply first-line indent.
        effective_indent = pf.left_indent + pf.first_line_indent
        if effective_indent > 0:
            pdf.cell(w=effective_indent * PT_TO_MM)

        run_size = get_dominant_font_size(runs) or fs
        line_h = self.line_height_mm(run_size, pf)

        with w._tag(pdf, "/P"):
            w._run_renderer.render_formatted_runs(
                pdf,
                runs,
                align=align,
                line_h_override=line_h,
                is_toc=is_toc_style(style_name),
                tab_stops=pf.tab_stops if pf.tab_stops else None,
                default_tab_stop=w._default_tab_stop,
                pf=pf,
            )
        self._apply_space_after(pdf, pf)

    # ------------------------------------------------------------------
    # Content-sequence path (mixed image+text paragraphs)
    # ------------------------------------------------------------------

    def _render_content_sequence(self, pdf: FPDF, para: ldm.Paragraph) -> None:
        """Render a paragraph using content_sequence to preserve XML element order."""
        w = self._writer
        pf = para.paragraph_format
        align = FPDF_ALIGN.get(pf.alignment, "L")

        self._apply_space_before(pdf, pf)

        pending_runs: list[ldm.Run] = []

        def _flush_styled() -> None:
            if not pending_runs:
                return
            runs_snapshot = list(pending_runs)
            pending_runs.clear()
            plain = safe_text("".join(r.text or "" for r in runs_snapshot))
            if not plain:
                return
            self._render_styled_block(
                pdf, runs_snapshot, pf, para.list_format, para.list_label, align
            )

        for item in para._children:
            if isinstance(item, ldm.ShapeNode):
                if item._is_positioned or item.wrap_type == WrapType.NONE:
                    continue  # drawn by positioned/floating pass
                if item.is_inline is False:
                    continue  # anchored/floating; drawn separately
                if item.has_image and item.image_data is not None:
                    _flush_styled()
                    w._shape_renderer.render_shape(pdf, item)
                elif item.text_box:
                    _flush_styled()
                    w._shape_renderer.render_shape(pdf, item)
            elif isinstance(item, ldm.Run):
                pending_runs.append(item)

        _flush_styled()

    # ------------------------------------------------------------------
    # List label helpers
    # ------------------------------------------------------------------

    def _compute_list_label(
        self, list_format: ldm.ListFormat
    ) -> Optional[str]:
        """Generate the marker for an ordered/bulleted list item."""
        w = self._writer
        doc = getattr(w, "_doc", None)
        if doc is None:
            return None
        list_id = list_format.list_id
        level = list_format.list_level_number
        lst = next((dl for dl in doc.lists if dl.list_id == list_id), None)
        if lst is None or not lst.levels:
            return None
        if level >= len(lst.levels):
            level = len(lst.levels) - 1
        lvl = lst.levels[level]
        if lvl.number_style in (23, 255):
            return "•"

        key = (list_id, level)
        counters = w._list_counters
        if key not in counters:
            counters[key] = max(lvl.start_at, 1)
        else:
            counters[key] += 1
        # Advancing a level resets all deeper-level counters.
        for k in list(counters.keys()):
            if k[0] == list_id and k[1] > level:
                counters.pop(k, None)

        fmt = lvl.number_format or "%1."
        return _expand_number_format(fmt, lst, counters, level, lvl.number_style)

    # ------------------------------------------------------------------
    # Spacing helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _apply_space_before(pdf: FPDF, pf: ldm.ParagraphFormat) -> None:
        """Add vertical space before a paragraph based on space_before."""
        if pf.space_before > 0:
            pdf.ln(pf.space_before * PT_TO_MM)

    @staticmethod
    def _apply_space_after(pdf: FPDF, pf: ldm.ParagraphFormat) -> None:
        """Add vertical space after a paragraph based on space_after."""
        if pf.space_after > 0:
            pdf.ln(pf.space_after * PT_TO_MM)

    # ------------------------------------------------------------------
    # Text helpers (thin wrappers for para-level use)
    # ------------------------------------------------------------------

    @staticmethod
    def _plain_text_from_runs(runs: list[ldm.Run]) -> str:
        """Return plain text from a list of runs, stripping Markdown links."""

        return "".join(
            apply_caps(chunk, run.font)
            for run in runs
            for chunk, _ in extract_link_segments(run.text or "")
        )

