"""PDF writer orchestrator operating on the light document model.

Converts an ldm.Document into a PDF file using fpdf2.  This is a thin
orchestrator that delegates to specialised sub-renderers for paragraphs,
tables, shapes, and formatted runs.
"""

from __future__ import annotations

from contextlib import contextmanager
from pathlib import Path
from typing import Iterator, Optional, Union

from fpdf import FPDF

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.pdf_writer.constants import (
    A4_HEIGHT_MM,
    A4_WIDTH_MM,
    COMPLIANCE_TO_VERSION,
    DEFAULT_FONT_NAME,
    DEFAULT_FONT_SIZE_PT,
    DEFAULT_MARGIN_MM,
    PT_TO_MM,
)
from aspose.words_foss.pdf_writer.page_bands import install_page_footer, install_page_header
from aspose.words_foss.pdf_writer.paragraph_renderer import ParagraphRenderer
from aspose.words_foss.pdf_writer.run_renderer import RunRenderer
from aspose.words_foss.pdf_writer.shape_renderer import ShapeRenderer
from aspose.words_foss.pdf_writer.table_renderer import TableRenderer
from aspose.words_foss.pdf_writer.text import (
    apply_caps,
    cell_text,
    extract_link_segments,
    is_pure_page_break,
    is_toc_style,
    plain_text,
    safe_text,
)
from aspose.words_foss.saving import PdfSaveOptions


class LdmPdfWriter:
    """Converts a ``light_document_model.Document`` to a PDF file."""

    def __init__(self, options: Optional[PdfSaveOptions] = None):
        self.options = options or PdfSaveOptions()
        # Page layout defaults — overridden by write() from section page_setup
        self._page_margin_left = DEFAULT_MARGIN_MM
        self._page_margin_right = DEFAULT_MARGIN_MM
        self._page_margin_bottom = DEFAULT_MARGIN_MM
        self._page_width = A4_WIDTH_MM
        self._page_height = A4_HEIGHT_MM
        # Internal-link registry: anchor name -> fpdf2 link ID.
        self._anchor_links: dict[str, int] = {}
        # Offset applied to pdf.page_no() when resolving PAGE fields.
        self._page_number_offset = 0

        # Sub-renderers (composed, not inherited)
        self._paragraph_renderer = ParagraphRenderer(self)
        self._run_renderer = RunRenderer(self)
        self._shape_renderer = ShapeRenderer(self)
        self._table_renderer = TableRenderer(self)

    def write(self, doc: ldm.Document, output_path: Union[str, Path]) -> None:
        """Convert *doc* to PDF and write to *output_path*."""
        output_path = Path(output_path)

        # Fresh anchor-link state per write
        self._anchor_links = {}

        # Reset page layout to defaults before each write
        self._page_margin_left = DEFAULT_MARGIN_MM
        self._page_margin_right = DEFAULT_MARGIN_MM
        self._page_margin_bottom = DEFAULT_MARGIN_MM
        self._page_width = A4_WIDTH_MM

        # Determine page size from the first section
        page_w_mm = A4_WIDTH_MM
        page_h_mm = A4_HEIGHT_MM
        if doc.sections:
            ps = doc.sections[0].page_setup
            if ps.page_width > 0:
                page_w_mm = ps.page_width * PT_TO_MM
            if ps.page_height > 0:
                page_h_mm = ps.page_height * PT_TO_MM

        pdf = FPDF(unit="mm", format=(page_w_mm, page_h_mm))
        # Apply PDF version from compliance setting
        version = COMPLIANCE_TO_VERSION.get(self.options.compliance)
        if version:
            pdf.pdf_version = version

        # Determine margins from the first section's page setup
        margin_top = DEFAULT_MARGIN_MM
        margin_bottom = DEFAULT_MARGIN_MM
        margin_left = DEFAULT_MARGIN_MM
        margin_right = DEFAULT_MARGIN_MM
        if doc.sections:
            ps = doc.sections[0].page_setup
            if ps.left_margin > 0:
                margin_left = ps.left_margin * PT_TO_MM
            if ps.right_margin > 0:
                margin_right = ps.right_margin * PT_TO_MM
            if ps.top_margin > 0:
                margin_top = ps.top_margin * PT_TO_MM
            if ps.bottom_margin > 0:
                margin_bottom = ps.bottom_margin * PT_TO_MM

        pdf.set_margins(margin_left, margin_top, margin_right)
        pdf.set_auto_page_break(auto=True, margin=margin_bottom)

        # Install per-page header / footer callbacks
        title_pg = bool(
            doc.sections and doc.sections[0].page_setup.different_first_page_header_footer
        )
        header_distance_mm = 0.0
        footer_distance_mm = 0.0
        if doc.sections:
            ps = doc.sections[0].page_setup
            if ps.header_distance > 0:
                header_distance_mm = ps.header_distance * PT_TO_MM
            if ps.footer_distance > 0:
                footer_distance_mm = ps.footer_distance * PT_TO_MM

        # Compute the page-number offset
        page_start = 1
        if doc.sections:
            ps = doc.sections[0].page_setup
            if ps.restart_page_numbering:
                page_start = ps.page_starting_number
        self._page_number_offset = page_start - 1

        # Update writer-level layout
        self._page_margin_left = margin_left
        self._page_margin_right = margin_right
        self._page_margin_bottom = margin_bottom
        self._page_width = page_w_mm
        self._page_height = page_h_mm

        install_page_header(pdf, doc, self, title_pg, header_distance_mm)
        install_page_footer(pdf, doc, self, title_pg, footer_distance_mm)

        pdf.add_page()
        pdf.set_font(DEFAULT_FONT_NAME, size=DEFAULT_FONT_SIZE_PT)

        # Render cover-page positioned shapes first
        self._shape_renderer.render_positioned_shapes(pdf, doc)

        # Body
        for section in doc.sections:
            for child in section.body.children:
                if isinstance(child, ldm.Paragraph):
                    self._paragraph_renderer.render_paragraph(pdf, child)
                elif isinstance(child, ldm.Table):
                    self._table_renderer.render_table(pdf, child)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        pdf.output(str(output_path))

    # ------------------------------------------------------------------
    # Tagged-PDF helpers
    # ------------------------------------------------------------------

    @contextmanager
    def _tag(self, pdf: FPDF, struct_type: str, title: Optional[str] = None) -> Iterator[None]:
        """Wrap content in a PDF structure element when tagging is on."""
        if not self.options.export_document_structure:
            yield
            return
        mcid = pdf.struct_builder.next_mcid_for_page(pdf.page)
        kwargs: dict = {"struct_type": struct_type, "mcid": mcid}
        if title:
            kwargs["title"] = title
        pdf._add_marked_content(**kwargs)
        pdf._out(f"/P <</MCID {mcid}>> BDC")
        yield
        pdf._out("EMC")

    # ------------------------------------------------------------------
    # Link helpers
    # ------------------------------------------------------------------

    def _link_target_for(self, pdf: FPDF, url: Optional[str]) -> Union[int, str]:
        """Return the fpdf2 ``link`` value for a Markdown hyperlink target."""
        if not url:
            return ""
        if url.startswith("#"):
            name = url[1:]
            if not name:
                return ""
            link_id = self._anchor_links.get(name)
            if link_id is None:
                link_id = pdf.add_link()
                self._anchor_links[name] = link_id
            return link_id
        return url

    # ------------------------------------------------------------------
    # Backward-compatible instance method aliases
    # ------------------------------------------------------------------

    def _compress_image_bytes(self, image_bytes: bytes) -> bytes:
        """Delegate to shape renderer (backward compat)."""
        return self._shape_renderer.compress_image_bytes(image_bytes)

    # ------------------------------------------------------------------
    # Backward-compatible static/class method aliases
    # ------------------------------------------------------------------
    # Tests and external code may call these directly on the class.

    _plain_text = staticmethod(plain_text)
    _cell_text = staticmethod(cell_text)
    _safe_text = staticmethod(safe_text)
    _extract_link_segments = staticmethod(extract_link_segments)
    _apply_caps = staticmethod(apply_caps)
    _is_pure_page_break = staticmethod(is_pure_page_break)
    _is_toc_style = staticmethod(is_toc_style)
    _cell_has_borders = staticmethod(TableRenderer._cell_has_borders)
    _line_height_mm = staticmethod(ParagraphRenderer.line_height_mm)
