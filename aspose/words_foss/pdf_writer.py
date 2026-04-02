"""PDF writer operating on the light document model.

Converts an ldm.Document into a PDF file using fpdf2.
"""

from __future__ import annotations

from io import BytesIO
from pathlib import Path
from typing import Optional, Union

from fpdf import FPDF

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.saving import ColorMode, PdfSaveOptions

# Internal rendering defaults (not part of the Aspose API)
_DEFAULT_FONT = "Helvetica"
_DEFAULT_FONT_SIZE = 11.0
_PAGE_WIDTH = 210.0  # A4 mm
_PAGE_HEIGHT = 297.0  # A4 mm
_MARGIN = 20.0


class LdmPdfWriter:
    """Converts a ``light_document_model.Document`` to a PDF file."""

    def __init__(self, options: Optional[PdfSaveOptions] = None):
        self.options = options or PdfSaveOptions()

    def write(self, doc: ldm.Document, output_path: Union[str, Path]) -> None:
        """Convert *doc* to PDF and write to *output_path*."""
        output_path = Path(output_path)

        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=_MARGIN)
        pdf.add_page()
        pdf.set_margins(_MARGIN, _MARGIN, _MARGIN)
        pdf.set_font(_DEFAULT_FONT, size=_DEFAULT_FONT_SIZE)

        # Header images
        for para in doc.header_paragraphs:
            self._render_paragraph(pdf, para)

        # Body
        for section in doc.sections:
            for child in section.body.children:
                if isinstance(child, ldm.Paragraph):
                    self._render_paragraph(pdf, child)
                elif isinstance(child, ldm.Table):
                    self._render_table(pdf, child)

        # Footer images
        for para in doc.footer_paragraphs:
            self._render_paragraph(pdf, para)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        pdf.output(str(output_path))

    def _render_image_shape(self, pdf: FPDF, shape: ldm.ShapeNode) -> None:
        """Embed an image from a ShapeNode into the PDF, scaling to fit page width."""
        img = shape.image_data
        if img is None or not img.image_bytes:
            return

        PT_TO_MM = 0.352778
        w_pt = shape.width or 100.0
        h_pt = shape.height or 100.0
        w_mm = w_pt * PT_TO_MM
        h_mm = h_pt * PT_TO_MM

        usable_w = _PAGE_WIDTH - 2 * _MARGIN
        if w_mm > usable_w:
            scale = usable_w / w_mm
            w_mm = usable_w
            h_mm = h_mm * scale

        pdf.image(BytesIO(img.image_bytes), w=w_mm, h=h_mm)
        pdf.ln(2)

    def _render_paragraph(self, pdf: FPDF, para: ldm.Paragraph) -> None:
        # Use content_sequence only when the paragraph actually mixes images with
        # text runs.  Pure-text paragraphs from the DOCX reader also have a
        # non-empty content_sequence, so checking merely for non-emptiness would
        # route every paragraph (including code-blocks, quotes, and list items)
        # through the incomplete content_sequence path.
        has_mixed = any(
            isinstance(i, ldm.ShapeNode) and i.has_image
            for i in para.content_sequence
        )
        if has_mixed:
            self._render_content_sequence(pdf, para)
            return

        # Legacy path: no inline images — emit any standalone images first, then text.
        for item in para.inline_extras:
            if isinstance(item, ldm.ShapeNode) and item.has_image and item.image_data is not None:
                self._render_image_shape(pdf, item)

        pf = para.paragraph_format
        style_name = pf.style_name
        fs = _DEFAULT_FONT_SIZE

        # Heading
        if pf.is_heading:
            level = min(pf.outline_level + 1, 6)
            size = fs + (6 - level) * 2
            pdf.set_font(_DEFAULT_FONT, style="B", size=size)
            text = self._plain_text(para)
            pdf.multi_cell(w=0, h=size * 0.6, text=self._safe_text(text))
            pdf.ln(size * 0.3)
            pdf.set_font(_DEFAULT_FONT, size=fs)
            return

        # Code block
        is_code = bool(style_name and ("Code" in style_name or "code" in style_name))
        if is_code:
            pdf.set_font("Courier", size=fs - 1)
            text = self._plain_text(para)
            pdf.multi_cell(w=0, h=fs * 0.5, text=self._safe_text(text))
            pdf.ln(2)
            pdf.set_font(_DEFAULT_FONT, size=fs)
            return

        # Block quote
        if style_name and "Quote" in style_name:
            pdf.set_font(_DEFAULT_FONT, style="I", size=fs)
            text = self._plain_text(para)
            pdf.cell(w=10)  # indent
            pdf.multi_cell(w=0, h=fs * 0.5, text=self._safe_text(text))
            pdf.ln(2)
            pdf.set_font(_DEFAULT_FONT, size=fs)
            return

        # List item
        if para.list_format and para.list_format.is_list_item:
            text = self._plain_text(para)
            level = para.list_format.list_level_number
            indent = 5 * (level + 1)
            label = para.list_format.list_label or "-"
            pdf.cell(w=indent)
            pdf.multi_cell(
                w=0, h=fs * 0.5, text=self._safe_text(f"{label} {text}")
            )
            pdf.ln(1)
            return

        # Normal paragraph — render with inline formatting
        text = self._plain_text(para)
        if not text.strip():
            pdf.ln(fs * 0.4)
            return

        self._render_formatted_runs(pdf, para.runs)
        pdf.ln(fs * 0.3)

    def _render_content_sequence(self, pdf: FPDF, para: ldm.Paragraph) -> None:
        """Render a paragraph using content_sequence to preserve XML element order.

        Images are emitted at their original positions; accumulated text runs are
        flushed at each image boundary and at the end with full paragraph-level
        styling (heading, code-block, quote, list-item, or plain).
        """
        fs = _DEFAULT_FONT_SIZE
        pf = para.paragraph_format
        style_name = pf.style_name
        is_code = bool(style_name and ("Code" in style_name or "code" in style_name))

        pending_runs: list[ldm.Run] = []

        def _flush_styled() -> None:
            if not pending_runs:
                return
            runs_snapshot = list(pending_runs)
            pending_runs.clear()
            plain = self._safe_text("".join(r.text or "" for r in runs_snapshot))
            if not plain:
                return
            if pf.is_heading:
                level = min(pf.outline_level + 1, 6)
                size = fs + (6 - level) * 2
                pdf.set_font(_DEFAULT_FONT, style="B", size=size)
                pdf.multi_cell(w=0, h=size * 0.6, text=plain)
                pdf.ln(size * 0.3)
                pdf.set_font(_DEFAULT_FONT, size=fs)
            elif is_code:
                pdf.set_font("Courier", size=fs - 1)
                pdf.multi_cell(w=0, h=fs * 0.5, text=plain)
                pdf.ln(2)
                pdf.set_font(_DEFAULT_FONT, size=fs)
            elif style_name and "Quote" in style_name:
                pdf.set_font(_DEFAULT_FONT, style="I", size=fs)
                pdf.cell(w=10)
                pdf.multi_cell(w=0, h=fs * 0.5, text=plain)
                pdf.ln(2)
                pdf.set_font(_DEFAULT_FONT, size=fs)
            elif para.list_format and para.list_format.is_list_item:
                level = para.list_format.list_level_number
                indent = 5 * (level + 1)
                label = para.list_format.list_label or "-"
                pdf.cell(w=indent)
                pdf.multi_cell(w=0, h=fs * 0.5, text=self._safe_text(f"{label} {plain}"))
                pdf.ln(1)
            else:
                self._render_formatted_runs(pdf, runs_snapshot)
                pdf.ln(fs * 0.3)

        for item in para.content_sequence:
            if isinstance(item, ldm.ShapeNode) and item.has_image and item.image_data is not None:
                _flush_styled()
                self._render_image_shape(pdf, item)
            elif isinstance(item, ldm.Run):
                pending_runs.append(item)

        _flush_styled()

    def _render_formatted_runs(self, pdf: FPDF, runs: list[ldm.Run]) -> None:
        """Render runs with inline bold/italic formatting."""
        fs = _DEFAULT_FONT_SIZE
        for run in runs:
            text = run.text or ""
            if not text:
                continue
            style = ""
            if run.font.bold:
                style += "B"
            if run.font.italic:
                style += "I"
            pdf.set_font(_DEFAULT_FONT, style=style, size=fs)
            pdf.write(h=fs * 0.5, text=self._safe_text(text))
        pdf.ln()

    def _render_table(self, pdf: FPDF, table: ldm.Table) -> None:
        if not table.rows:
            return

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
                        self._render_paragraph(pdf, para)
            pdf.ln(2)
            return

        num_cols = max(len(row.cells) for row in table.rows)
        if num_cols == 0:
            return

        page_width = _PAGE_WIDTH - _MARGIN * 2
        col_width = page_width / num_cols
        line_height = _DEFAULT_FONT_SIZE * 0.5

        for row in table.rows:
            for i, cell in enumerate(row.cells):
                text = self._safe_text(self._cell_text(cell))
                pdf.cell(
                    w=col_width,
                    h=line_height,
                    text=text,
                    border=1,
                )
            pdf.ln(line_height)

        pdf.ln(2)

    @staticmethod
    def _plain_text(para: ldm.Paragraph) -> str:
        return "".join(run.text or "" for run in para.runs)

    @staticmethod
    def _cell_text(cell: ldm.Cell) -> str:
        parts = []
        for para in cell.paragraphs:
            text = "".join(run.text or "" for run in para.runs)
            if text:
                parts.append(text)
        return " ".join(parts)

    @staticmethod
    def _safe_text(text: str) -> str:
        """Replace characters outside latin-1 range for built-in PDF fonts."""
        return text.encode("latin-1", errors="replace").decode("latin-1")
