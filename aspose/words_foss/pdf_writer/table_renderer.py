"""Table rendering for the PDF writer."""

from __future__ import annotations

from typing import TYPE_CHECKING, Optional

from fpdf import FPDF

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.model.enums import LineStyle, ParagraphAlignment
from aspose.words_foss.pdf_writer.color import parse_color
from aspose.words_foss.pdf_writer.constants import (
    A4_HEIGHT_MM,
    CHAR_WIDTH_ESTIMATE_FACTOR,
    DEFAULT_CELL_LINE_H_FACTOR,
    DEFAULT_CELL_PAD_LEFT_MM,
    DEFAULT_CELL_PAD_TOP_MM,
    DEFAULT_FONT_SIZE_PT,
    MIN_ROW_HEIGHT_FACTOR,
    POST_TABLE_SPACING_MM,
    PT_TO_MM,
)
from aspose.words_foss.pdf_writer.font import apply_run_font, reset_font
from aspose.words_foss.pdf_writer.text import cell_text, safe_text

if TYPE_CHECKING:
    from aspose.words_foss.pdf_writer.renderer import LdmPdfWriter


class TableRenderer:
    """Renders LDM tables into PDF."""

    def __init__(self, writer: LdmPdfWriter) -> None:
        self._writer = writer

    def render_table(self, pdf: FPDF, table: ldm.Table) -> None:
        """Render a complete table."""
        if not table.rows:
            return

        w = self._writer

        # If any cell has images, render as flowing paragraphs rather than a fixed grid
        has_cell_images = any(
            isinstance(item, ldm.ShapeNode) and item.has_image
            for row in table.rows
            for cell in row.cells
            for para in cell.paragraphs
            for item in para.inline_extras
        )
        if has_cell_images:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        w._paragraph_renderer.render_paragraph(pdf, para)
            pdf.ln(POST_TABLE_SPACING_MM)
            return

        num_cols = max(len(row.cells) for row in table.rows)
        if num_cols == 0:
            return

        usable_w = w._page_width - w._page_margin_left - w._page_margin_right
        col_widths = self._compute_col_widths(table, num_cols, usable_w)

        # Table padding
        table_pad_top = table.top_padding * PT_TO_MM if table.top_padding else 0
        table_pad_bottom = table.bottom_padding * PT_TO_MM if table.bottom_padding else 0

        if table_pad_top:
            pdf.ln(table_pad_top)

        # Table horizontal position
        table_x = pdf.get_x()
        if table.left_indent > 0:
            table_x += table.left_indent * PT_TO_MM
        total_table_w = sum(col_widths)
        free_space = usable_w - total_table_w
        if free_space > 0:
            if table.alignment == ParagraphAlignment.CENTER:
                table_x += free_space / 2
            elif table.alignment == ParagraphAlignment.RIGHT:
                table_x += free_space

        for row in table.rows:
            row_height = self._compute_row_height(pdf, row, col_widths, num_cols)
            row_y = pdf.get_y()

            # Check for page break
            if (
                row_y + row_height
                > getattr(w, "_page_height", A4_HEIGHT_MM) - w._page_margin_bottom
            ):
                pdf.add_page()
                row_y = pdf.get_y()

            for i, cell in enumerate(row.cells):
                if i >= num_cols:
                    break
                cw = col_widths[i]
                cf = cell.cell_format

                cell_x = table_x + sum(col_widths[:i])

                # Draw cell background
                bg_color = parse_color(cf.shading.background_color)
                if bg_color:
                    pdf.set_fill_color(*bg_color)
                    pdf.rect(cell_x, row_y, cw, row_height, "F")

                # Draw cell border only when the cell actually requests one.
                if self._cell_has_borders(cell):
                    pdf.rect(cell_x, row_y, cw, row_height)

                # Cell padding
                pad_left = (
                    cf.left_padding * PT_TO_MM if cf.left_padding else DEFAULT_CELL_PAD_LEFT_MM
                )
                pad_top = cf.top_padding * PT_TO_MM if cf.top_padding else DEFAULT_CELL_PAD_TOP_MM

                # Render cell text
                text = safe_text(cell_text(cell))
                pdf.set_xy(cell_x + pad_left, row_y + pad_top)
                with w._tag(pdf, "/TD"):
                    first_font = self._get_cell_first_font(cell)
                    if first_font:
                        apply_run_font(pdf, first_font)
                    else:
                        reset_font(pdf)
                    pdf.multi_cell(
                        w=cw - 2 * pad_left,
                        h=DEFAULT_FONT_SIZE_PT * DEFAULT_CELL_LINE_H_FACTOR,
                        text=text,
                    )
                reset_font(pdf)

            # Move to next row
            pdf.set_xy(table_x, row_y + row_height)

        if table_pad_bottom:
            pdf.ln(table_pad_bottom)
        else:
            pdf.ln(POST_TABLE_SPACING_MM)

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _compute_col_widths(table: ldm.Table, num_cols: int, usable_w: float) -> list[float]:
        """Compute column widths from cell formats or distribute evenly."""
        widths = [0.0] * num_cols
        has_explicit = False
        for row in table.rows:
            for i, cell in enumerate(row.cells):
                if i >= num_cols:
                    break
                w = cell.cell_format.width
                if w > 0:
                    w_mm = w * PT_TO_MM
                    widths[i] = max(widths[i], w_mm)
                    has_explicit = True

        if not has_explicit:
            return [usable_w / num_cols] * num_cols

        total_explicit = sum(w for w in widths if w > 0)
        zero_count = sum(1 for w in widths if w == 0)
        if zero_count > 0 and total_explicit < usable_w:
            fill = (usable_w - total_explicit) / zero_count
            widths = [w if w > 0 else fill for w in widths]

        total = sum(widths)
        if total > usable_w and total > 0:
            scale = usable_w / total
            widths = [w * scale for w in widths]

        return widths

    def _compute_row_height(
        self, pdf: FPDF, row: ldm.Row, col_widths: list[float], num_cols: int
    ) -> float:
        """Estimate row height based on text content."""
        min_h = DEFAULT_FONT_SIZE_PT * MIN_ROW_HEIGHT_FACTOR
        max_h = min_h
        for i, cell in enumerate(row.cells):
            if i >= num_cols:
                break
            text = cell_text(cell)
            cw = col_widths[i]
            if cw > 2:
                char_w = DEFAULT_FONT_SIZE_PT * CHAR_WIDTH_ESTIMATE_FACTOR
                chars_per_line = max(1, int(cw / char_w))
                num_lines = 0
                for line in text.split("\n") or [""]:
                    num_lines += max(1, (len(line) // chars_per_line) + 1)
                h = num_lines * DEFAULT_FONT_SIZE_PT * DEFAULT_CELL_LINE_H_FACTOR + 1.0
                max_h = max(max_h, h)
        if row.row_format.height > 0:
            explicit_h = row.row_format.height * PT_TO_MM
            max_h = max(max_h, explicit_h)
        return max_h

    @staticmethod
    def _get_cell_first_font(cell: ldm.Cell) -> Optional[ldm.Font]:
        """Get the font from the first non-empty run in a cell."""
        for para in cell.paragraphs:
            for run in para.runs:
                if run.text:
                    return run.font
        return None

    @staticmethod
    def _cell_has_borders(cell: ldm.Cell) -> bool:
        """Return True when the cell requests any visible borders."""
        for border in cell.cell_format.borders or ():
            if border.line_style != LineStyle.NONE or border.line_width > 0:
                return True
        return False
