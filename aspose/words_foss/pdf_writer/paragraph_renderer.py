"""Paragraph rendering for the PDF writer.

Contains the unified styled-block rendering logic used by both the
legacy path and the content-sequence (mixed image+text) path,
eliminating the previous code duplication.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Optional

from fpdf import FPDF

from aspose.words_foss import light_document_model as ldm
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
)
from aspose.words_foss.pdf_writer.font import reset_font
from aspose.words_foss.pdf_writer.page_bands import register_bookmarks
from aspose.words_foss.pdf_writer.text import (
    get_dominant_color,
    get_dominant_font_size,
    is_pure_page_break,
    is_toc_style,
    safe_text,
)

if TYPE_CHECKING:
    from aspose.words_foss.pdf_writer.renderer import LdmPdfWriter


class ParagraphRenderer:
    """Renders LDM paragraphs into PDF."""

    def __init__(self, writer: LdmPdfWriter) -> None:
        self._writer = writer

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
        from aspose.words_foss.model.enums import LineSpacingRule
        from aspose.words_foss.pdf_writer.constants import WORD_LINE_SPACING_DIVISOR

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

        register_bookmarks(pdf, para, self._writer._anchor_links)
        self._render_paragraph_body(pdf, para)

        if any("\f" in (run.text or "") for run in para.runs):
            pdf.add_page()

    # ------------------------------------------------------------------
    # Body dispatcher
    # ------------------------------------------------------------------

    def _render_paragraph_body(self, pdf: FPDF, para: ldm.Paragraph) -> None:
        """Render paragraph body: floating images, then styled content."""
        w = self._writer

        # Draw floating (wrapNone) images at their anchor coordinates
        w._shape_renderer.render_floating_images(pdf, para)

        # Use content_sequence only when the paragraph actually mixes
        # *inline* images with text runs.
        has_mixed = any(
            isinstance(i, ldm.ShapeNode)
            and i.has_image
            and not i._is_positioned
            and i.wrap_type != WrapType.NONE
            for i in para.content_sequence
        )
        if has_mixed:
            self._render_content_sequence(pdf, para)
            return

        # Legacy path: no inline images — emit any standalone shapes first.
        for item in para.inline_extras:
            if isinstance(item, ldm.ShapeNode):
                if item._is_positioned or item.wrap_type == WrapType.NONE:
                    continue
                if item.has_image and item.image_data is not None:
                    w._shape_renderer.render_shape(pdf, item)
                elif item.text_box:
                    w._shape_renderer.render_shape(pdf, item)

        pf = para.paragraph_format
        align = FPDF_ALIGN.get(pf.alignment, "L")

        # Ensure each paragraph starts at the left margin.
        pdf.set_x(w._page_margin_left)

        self._render_styled_block(pdf, para.runs, pf, para.list_format, align)

    # ------------------------------------------------------------------
    # Unified styled-block rendering (DRY fix)
    # ------------------------------------------------------------------

    def _render_styled_block(
        self,
        pdf: FPDF,
        runs: list[ldm.Run],
        pf: ldm.ParagraphFormat,
        list_format: Optional[ldm.ListFormat],
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
            if w.options.export_document_structure and text.strip():
                pdf.start_section(text.strip(), level=level - 1)
            with w._tag(pdf, f"/H{level}", title=text.strip()):
                heading_color = get_dominant_color(runs)
                if heading_color:
                    pdf.set_text_color(*heading_color)
                pdf.set_font(DEFAULT_FONT_NAME, style="B", size=size)
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
            indent = (
                pf.left_indent * PT_TO_MM
                if pf.left_indent > 0
                else LIST_INDENT_PER_LEVEL_MM * (level + 1)
            )
            label = list_format.list_label or "-"
            with w._tag(pdf, "/LI"):
                pdf.cell(w=indent)
                run_size = get_dominant_font_size(runs)
                effective_fs = run_size if run_size > 0 else fs
                line_h = self.line_height_mm(effective_fs, pf)
                label_text = safe_text(f"{label} ")
                pdf.set_font(DEFAULT_FONT_NAME, size=effective_fs)
                pdf.write(h=line_h, text=label_text)
                w._run_renderer.render_formatted_runs(
                    pdf,
                    runs,
                    newline=True,
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
            self._render_styled_block(pdf, runs_snapshot, pf, para.list_format, align)

        for item in para.content_sequence:
            if isinstance(item, ldm.ShapeNode):
                if item._is_positioned or item.wrap_type == WrapType.NONE:
                    continue  # drawn by positioned/floating pass
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
        from aspose.words_foss.pdf_writer.text import apply_caps, extract_link_segments

        return "".join(
            apply_caps(chunk, run.font)
            for run in runs
            for chunk, _ in extract_link_segments(run.text or "")
        )
