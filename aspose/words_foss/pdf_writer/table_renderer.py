"""Table rendering for the PDF writer."""


from typing import Optional

from fpdf import FPDF
from fpdf.errors import FPDFException

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss._visible_runs import visible_runs
from aspose.words_foss.model.enums import LineStyle, ParagraphAlignment
from aspose.words_foss.pdf_writer.color import parse_color
from aspose.words_foss.pdf_writer.constants import (
    A4_HEIGHT_MM,
    CHAR_WIDTH_ESTIMATE_FACTOR,
    DEFAULT_CELL_LINE_H_FACTOR,
    DEFAULT_CELL_PAD_LEFT_MM,
    DEFAULT_CELL_PAD_TOP_MM,
    DEFAULT_FONT_SIZE_PT,
    MIN_LINE_WIDTH_MM,
    MIN_ROW_HEIGHT_FACTOR,
    POST_TABLE_SPACING_MM,
    PT_TO_MM,
)
from aspose.words_foss.pdf_writer.font import apply_run_font, reset_font
from aspose.words_foss.pdf_writer._context import PDFWriterContext
from aspose.words_foss.pdf_writer.text import cell_text, safe_text

# Border slot indices in the LDM `borders` list, matching the canonical
# BorderType layout shared with the DOCX writer:
#   0=Bottom, 1=Left, 2=Right, 3=Top, 4=insideH, 5=insideV
_B_BOTTOM, _B_LEFT, _B_RIGHT, _B_TOP = 0, 1, 2, 3
_B_INSIDE_H, _B_INSIDE_V = 4, 5


# CellFormat.orientation values that require rotated text rendering.
# 1=btLr (bottom-to-top), 2=tbRl (top-to-bottom), 3=lrTbV, 4=tbRlV, 5=tbLrV
_VERTICAL_ORIENTATIONS = {1, 2, 3, 4, 5}

# Maps orientation enum to clockwise rotation degrees for fpdf2.
_ORIENTATION_TO_ANGLE: dict[int, float] = {
    1: 90.0,    # btLr: text reads bottom-to-top
    2: -90.0,   # tbRl: text reads top-to-bottom
    3: 90.0,    # lrTbV
    4: -90.0,   # tbRlV
    5: 90.0,    # tbLrV
}


class TableRenderer:
    """Renders LDM tables into PDF."""

    def __init__(self, writer: PDFWriterContext) -> None:
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
            for item in para._children
        )
        if has_cell_images:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        w._paragraph_renderer.render_paragraph(pdf, para)
            pdf.ln(POST_TABLE_SPACING_MM)
            return

        # Tables with an explicit ``tblpXSpec`` (right/center/inside/
        # outside) are pinned to a side of the column area; we render
        # at the float coordinates and leave the cursor alone so the
        # inline flow continues under the table. ``tblpXSpec`` left
        # (or absent) is equivalent to a normal inline table that just
        # happens to have an offset.
        attrs = table._tblp_pr_attrs
        if (
            attrs
            and table.text_wrapping == 1
            and attrs.get("tblpXSpec") in ("right", "center", "inside", "outside")
        ):
            self._render_floating_table(pdf, table)
            return

        num_cols = max(len(row.cells) for row in table.rows)
        if num_cols == 0:
            return
        num_rows = len(table.rows)

        usable_w = w._page_width - w._page_margin_left - w._page_margin_right
        col_widths = self._compute_col_widths(table, num_cols, usable_w)

        # Borders cascade: cell explicit > table-level (carried per-row by
        # the reader as row_format.borders, since the LDM has no Table
        # field) > table style. The table-level set can disable insideH /
        # insideV the style enables, so it must win over the style.
        table_level_borders = self._table_level_borders(table)
        style_borders = self._inherited_table_borders(table)

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

        for row_idx, row in enumerate(table.rows):
            row_height = self._compute_row_height(pdf, row, col_widths, num_cols)
            row_y = pdf.get_y()

            # Check for page break (skip during header/footer rendering
            # to avoid infinite recursion with the footer callback).
            in_hf = getattr(pdf, "_in_header_render", False) or getattr(
                pdf, "_in_footer_render", False
            )
            if (
                not in_hf
                and row_y + row_height
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
                bg_color = parse_color(cf.shading.background_pattern_color)
                if bg_color:
                    pdf.set_fill_color(*bg_color)
                    pdf.rect(cell_x, row_y, cw, row_height, "F")

                self._draw_cell_borders(
                    pdf, cell, table_level_borders, style_borders,
                    row_idx, i, num_rows, num_cols,
                    cell_x, row_y, cw, row_height,
                )

                # Cell padding
                pad_left = (
                    cf.left_padding * PT_TO_MM if cf.left_padding else DEFAULT_CELL_PAD_LEFT_MM
                )
                pad_top = cf.top_padding * PT_TO_MM if cf.top_padding else DEFAULT_CELL_PAD_TOP_MM

                # Render cell text
                text = safe_text(cell_text(cell))
                line_h = DEFAULT_FONT_SIZE_PT * DEFAULT_CELL_LINE_H_FACTOR

                with w._tag(pdf, "/TD"):
                    first_font = self._get_cell_first_font(cell)
                    if first_font:
                        apply_run_font(pdf, first_font)
                    else:
                        reset_font(pdf)

                    if cf.orientation in _VERTICAL_ORIENTATIONS:
                        self._render_rotated_cell(
                            pdf, text, cell_x, row_y, cw, row_height,
                            pad_left, pad_top, line_h, cf.orientation,
                        )
                    elif not cf.wrap_text:
                        with pdf.rect_clip(cell_x, row_y, cw, row_height):
                            pdf.set_xy(cell_x + pad_left, row_y + pad_top)
                            pdf.cell(
                                w=cw - 2 * pad_left,
                                h=line_h,
                                text=text,
                            )
                    else:
                        usable_w = cw - 2 * pad_left
                        pdf.set_xy(cell_x + pad_left, row_y + pad_top)
                        # multi_cell raises if the usable width is smaller
                        # than one glyph; clip the text in that case.
                        if usable_w > 0:
                            try:
                                pdf.multi_cell(w=usable_w, h=line_h, text=text)
                            except FPDFException:
                                with pdf.rect_clip(cell_x, row_y, cw, row_height):
                                    pdf.set_xy(cell_x + pad_left, row_y + pad_top)
                                    pdf.cell(w=usable_w, h=line_h, text=text)
                reset_font(pdf)

            # Move to next row
            pdf.set_xy(table_x, row_y + row_height)

        if table_pad_bottom:
            pdf.ln(table_pad_bottom)
        else:
            pdf.ln(POST_TABLE_SPACING_MM)

    def _render_floating_table(self, pdf: FPDF, table: ldm.Table) -> None:
        """Draw a `w:tblpPr` floating table without disturbing the cursor.

        Honours ``tblpXSpec`` (``left``/``center``/``right``); ``tblpY``
        is added to the live cursor Y. The text cursor is restored so
        the column flow continues underneath instead of being pushed
        below the table.
        """
        w = self._writer
        saved_x, saved_y = pdf.get_x(), pdf.get_y()

        attrs = table._tblp_pr_attrs
        spec = attrs.get("tblpXSpec", "left")
        # Page-level margins (not the current column's): floating tables
        # position themselves against the page or section margin, not
        # the column edge.
        page_left = getattr(w, "_page_full_margin_left", w._page_margin_left)
        page_right = getattr(w, "_page_full_margin_right", w._page_margin_right)
        page_usable_w = w._page_width - page_left - page_right
        num_cols = max(len(row.cells) for row in table.rows)
        col_widths = self._compute_col_widths(table, num_cols, page_usable_w)
        total_w = sum(col_widths)

        if spec == "right":
            float_x = w._page_width - page_right - total_w
        elif spec == "center":
            float_x = page_left + (page_usable_w - total_w) / 2
        else:
            float_x = saved_x

        tblpY = attrs.get("tblpY")
        y_offset_mm = 0.0
        if tblpY:
            try:
                y_offset_mm = int(tblpY) / 20.0 * PT_TO_MM  # twips → pt → mm
            except ValueError:
                pass
        # When the table pins to a side column, take the anchor Y from
        # that column's live cursor — not the writer cursor, which may
        # have hopped to a different column during section reflow.
        col_y_map = getattr(w, "_active_col_y", None)
        anchor_y = saved_y
        if spec == "right" and col_y_map:
            anchor_y = col_y_map.get(col_y_map and (len(col_y_map) - 1) or 0, saved_y)
            # Pick the rightmost known column; multi-col docs almost
            # always have ncols ≤ 2 so the last key is the right side.
            anchor_y = col_y_map.get(max(col_y_map), saved_y)
        float_y = anchor_y + y_offset_mm

        prev_left = w._page_margin_left
        prev_right = w._page_margin_right
        pdf.set_left_margin(float_x)
        pdf.set_right_margin(max(0.0, w._page_width - float_x - total_w))
        w._page_margin_left = float_x
        w._page_margin_right = max(0.0, w._page_width - float_x - total_w)
        pdf.set_xy(float_x, float_y)

        attrs_holder = table._tblp_pr_attrs
        table._tblp_pr_attrs = {}  # avoid recursion
        try:
            self.render_table(pdf, table)
            table_bottom = pdf.get_y()
        finally:
            table._tblp_pr_attrs = attrs_holder
            pdf.set_left_margin(prev_left)
            pdf.set_right_margin(prev_right)
            w._page_margin_left = prev_left
            w._page_margin_right = prev_right
            # If the inline cursor is already in the same column the
            # table just landed in, skip past it so the next paragraph
            # doesn't render on top of the table. (Cursor was always
            # in another column when reflow had moved us elsewhere.)
            if pdf.l_margin == w._page_margin_left and saved_y < table_bottom:
                pdf.set_xy(saved_x, table_bottom)
            else:
                pdf.set_xy(saved_x, saved_y)
        # Push the float's target column cursor past the table so the
        # next paragraphs in that column don't render on top of it.
        if col_y_map is not None and spec == "right":
            tgt_col = max(col_y_map)
            col_y_map[tgt_col] = max(col_y_map.get(tgt_col, 0.0), table_bottom)

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
            cf = cell.cell_format
            pad_top = cf.top_padding * PT_TO_MM if cf.top_padding else DEFAULT_CELL_PAD_TOP_MM
            pad_left = cf.left_padding * PT_TO_MM if cf.left_padding else DEFAULT_CELL_PAD_LEFT_MM
            line_h = DEFAULT_FONT_SIZE_PT * DEFAULT_CELL_LINE_H_FACTOR
            if cf.orientation in _VERTICAL_ORIENTATIONS:
                text_w = self._measure_text_width(pdf, text, cell)
                h = text_w + 2 * pad_top
                max_h = max(max_h, h)
            elif not cf.wrap_text:
                h = line_h + 2 * pad_top
                max_h = max(max_h, h)
            elif cw > 2 * pad_left + 1:
                inner_w = cw - 2 * pad_left
                num_lines = self._estimate_wrap_lines(pdf, text, cell, inner_w)
                h = num_lines * line_h + 2 * pad_top
                max_h = max(max_h, h)
        if row.row_format.height > 0:
            explicit_h = row.row_format.height * PT_TO_MM
            max_h = max(max_h, explicit_h)
        return max_h

    @staticmethod
    def _render_rotated_cell(
        pdf: FPDF,
        text: str,
        cell_x: float,
        row_y: float,
        cw: float,
        rh: float,
        pad_left: float,
        pad_top: float,
        line_h: float,
        orientation: int,
    ) -> None:
        """Render text rotated 90/-90 degrees inside a cell."""
        angle = _ORIENTATION_TO_ANGLE.get(orientation, 90.0)
        cx = cell_x + cw / 2
        cy = row_y + rh / 2
        text = safe_text(text)
        with pdf.rotation(angle, cx, cy):
            tw = pdf.get_string_width(text)
            pdf.set_xy(cx - tw / 2, cy - line_h / 2)
            pdf.cell(w=tw, h=line_h, text=text)

    @staticmethod
    def _get_cell_first_font(cell: ldm.Cell) -> Optional[ldm.Font]:
        """Get the font from the first non-empty run in a cell."""
        for para in cell.paragraphs:
            for run in visible_runs(para):
                if run.text:
                    return run.font
        return None

    def _measure_text_width(
        self, pdf: FPDF, text: str, cell: ldm.Cell
    ) -> float:
        """Width of *text* in mm under the cell's first font."""
        if not text:
            return 0.0
        # Strip characters the core PDF font can't render so fpdf2's
        # encoder doesn't raise mid-measurement.
        text = safe_text(text)
        font = self._get_cell_first_font(cell)
        prev_family = pdf.font_family
        prev_style = pdf.font_style
        prev_size = pdf.font_size_pt
        try:
            if font:
                apply_run_font(pdf, font)
            else:
                reset_font(pdf)
            return float(pdf.get_string_width(text))
        finally:
            pdf.set_font(prev_family, prev_style, prev_size)

    def _estimate_wrap_lines(
        self, pdf: FPDF, text: str, cell: ldm.Cell, inner_w: float
    ) -> int:
        """Number of lines *text* takes when wrapped to *inner_w* (mm)."""
        if not text:
            return 1
        lines = 0
        for source_line in text.split("\n"):
            if not source_line:
                lines += 1
                continue
            line_w = self._measure_text_width(pdf, source_line, cell)
            if line_w <= inner_w:
                lines += 1
            else:
                lines += max(1, int(line_w / inner_w + 0.999))
        return max(1, lines)

    @staticmethod
    def _cell_has_borders(cell: ldm.Cell) -> bool:
        """Return True when the cell requests any visible borders."""
        for border in cell.cell_format.borders or ():
            if border.line_style != LineStyle.NONE or border.line_width > 0:
                return True
        return False

    @staticmethod
    def _border_is_visible(border: Optional[ldm.Border]) -> bool:
        if border is None:
            return False
        return border.line_style != LineStyle.NONE or border.line_width > 0

    def _inherited_table_borders(
        self, table: ldm.Table
    ) -> Optional[list[ldm.Border]]:
        """Return borders from the table's style chain, if any."""
        doc = getattr(self._writer, "_doc", None)
        if doc is None or not table.style_name:
            return None
        seen: set[str] = set()
        name = table.style_name
        while name and name not in seen:
            seen.add(name)
            style = doc.find_style(name)
            if style is None:
                return None
            fmt = style.table_style_format
            if fmt and fmt.borders:
                return list(fmt.borders)
            name = style.base_style_name
        return None

    @staticmethod
    def _table_level_borders(table: ldm.Table) -> Optional[list[ldm.Border]]:
        """Return the table-level w:tblBorders (carried per-row by the reader)."""
        for row in table.rows:
            if row.row_format.borders:
                return list(row.row_format.borders)
        return None

    def _resolve_cell_side(
        self,
        cell_borders: list[ldm.Border],
        table_level_borders: Optional[list[ldm.Border]],
        style_borders: Optional[list[ldm.Border]],
        outer_slot: int,
        inside_slot: int,
        is_outer: bool,
    ) -> Optional[ldm.Border]:
        """Resolve a single cell side: explicit cell > table-level > style."""
        cb = cell_borders[outer_slot] if outer_slot < len(cell_borders) else None
        if self._border_is_visible(cb):
            return cb
        slot = outer_slot if is_outer else inside_slot
        for source in (table_level_borders, style_borders):
            if source is None:
                continue
            if slot >= len(source):
                continue
            # An explicitly-set ``none`` entry in this layer is a deliberate
            # override that should hide the side rather than letting the
            # next layer leak through.
            return source[slot] if self._border_is_visible(source[slot]) else None
        return None

    def _draw_cell_borders(
        self,
        pdf: FPDF,
        cell: ldm.Cell,
        table_level_borders: Optional[list[ldm.Border]],
        style_borders: Optional[list[ldm.Border]],
        row_idx: int,
        col_idx: int,
        num_rows: int,
        num_cols: int,
        x: float,
        y: float,
        w: float,
        h: float,
    ) -> None:
        """Draw cell sides using cell → table-level → style cascade."""
        cell_borders = cell.cell_format.borders or []

        # Interior edges fall back to insideH / insideV; outer edges use
        # the outer slot. The outer-slot border can still hide an interior
        # edge if a cell explicitly sets its own border.
        top = self._resolve_cell_side(
            cell_borders, table_level_borders, style_borders,
            _B_TOP, _B_INSIDE_H, row_idx == 0,
        )
        bottom = self._resolve_cell_side(
            cell_borders, table_level_borders, style_borders,
            _B_BOTTOM, _B_INSIDE_H, row_idx == num_rows - 1,
        )
        left = self._resolve_cell_side(
            cell_borders, table_level_borders, style_borders,
            _B_LEFT, _B_INSIDE_V, col_idx == 0,
        )
        right = self._resolve_cell_side(
            cell_borders, table_level_borders, style_borders,
            _B_RIGHT, _B_INSIDE_V, col_idx == num_cols - 1,
        )

        if not any((top, left, bottom, right)):
            return

        prev_lw = pdf.line_width

        def draw_line(b: ldm.Border, x1: float, y1: float, x2: float, y2: float) -> None:
            rgb = parse_color(b.color) or (0, 0, 0)
            pdf.set_draw_color(*rgb)
            lw_mm = max(b.line_width * PT_TO_MM, MIN_LINE_WIDTH_MM)
            pdf.set_line_width(lw_mm)
            pdf.line(x1, y1, x2, y2)

        if top:
            draw_line(top, x, y, x + w, y)
        if bottom:
            draw_line(bottom, x, y + h, x + w, y + h)
        if left:
            draw_line(left, x, y, x, y + h)
        if right:
            draw_line(right, x + w, y, x + w, y + h)

        pdf.set_draw_color(0, 0, 0)
        pdf.set_line_width(prev_lw)
