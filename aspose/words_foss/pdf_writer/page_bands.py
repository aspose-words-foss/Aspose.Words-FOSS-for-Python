"""Header/footer callback installation and bookmark registration."""

from __future__ import annotations

from typing import TYPE_CHECKING

from fpdf import FPDF

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.pdf_writer.constants import MIN_HEADER_FOOTER_Y_MM

if TYPE_CHECKING:
    from aspose.words_foss.pdf_writer.renderer import LdmPdfWriter


def install_page_header(
    pdf: FPDF,
    doc: ldm.Document,
    writer: LdmPdfWriter,
    skip_first_page: bool,
    header_y_mm: float,
) -> None:
    """Install a per-page header callback on *pdf*."""
    if not doc.header_paragraphs:
        return

    def _header_callback() -> None:
        if skip_first_page and pdf.page_no() == 1:
            return
        if getattr(pdf, "_in_header_render", False):
            return
        pdf._in_header_render = True  # type: ignore[attr-defined]
        prev_auto = pdf.auto_page_break
        prev_bottom = pdf.b_margin
        body_start_y = pdf.t_margin
        try:
            pdf.set_auto_page_break(auto=False, margin=0)
            pdf.set_xy(writer._page_margin_left, max(header_y_mm, MIN_HEADER_FOOTER_Y_MM))
            for para in doc.header_paragraphs:
                para_start_y = pdf.get_y()
                writer._paragraph_renderer.render_paragraph(pdf, para)
                writer._shape_renderer.render_positioned_shapes_in(
                    pdf, [para], line_y_override=para_start_y
                )
            final_y = max(pdf.get_y(), body_start_y)
            pdf.set_xy(writer._page_margin_left, final_y)
        finally:
            pdf.set_auto_page_break(auto=prev_auto, margin=prev_bottom)
            pdf._in_header_render = False  # type: ignore[attr-defined]

    pdf.header = _header_callback  # type: ignore[method-assign]


def install_page_footer(
    pdf: FPDF,
    doc: ldm.Document,
    writer: LdmPdfWriter,
    skip_first_page: bool,
    footer_y_mm: float,
) -> None:
    """Install a per-page footer callback on *pdf*."""
    if not doc.footer_paragraphs:
        return

    def _footer_callback() -> None:
        if skip_first_page and pdf.page_no() == 1:
            return
        if getattr(pdf, "_in_footer_render", False):
            return
        pdf._in_footer_render = True  # type: ignore[attr-defined]
        prev_auto = pdf.auto_page_break
        prev_bottom = pdf.b_margin
        try:
            pdf.set_auto_page_break(auto=False, margin=0)
            band_top = max(
                writer._page_height - max(footer_y_mm, MIN_HEADER_FOOTER_Y_MM),
                writer._page_height - writer._page_margin_bottom,
            )
            pdf.set_xy(writer._page_margin_left, band_top)
            for para in doc.footer_paragraphs:
                writer._paragraph_renderer.render_paragraph(pdf, para)
            writer._shape_renderer.render_positioned_shapes_in(pdf, doc.footer_paragraphs)
        finally:
            pdf.set_auto_page_break(auto=prev_auto, margin=prev_bottom)
            pdf._in_footer_render = False  # type: ignore[attr-defined]

    pdf.footer = _footer_callback  # type: ignore[method-assign]


def register_bookmarks(
    pdf: FPDF,
    para: ldm.Paragraph,
    anchor_links: dict[str, int],
) -> None:
    """Point every :class:`BookmarkStart` in *para* at the current page."""
    for extra in para.inline_extras:
        if not isinstance(extra, ldm.BookmarkStart) or not extra.name:
            continue
        link_id = anchor_links.get(extra.name)
        if link_id is None:
            link_id = pdf.add_link()
            anchor_links[extra.name] = link_id
        pdf.set_link(link_id, y=pdf.get_y(), page=pdf.page_no())
