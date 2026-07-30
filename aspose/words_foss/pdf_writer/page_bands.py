"""Header/footer callback installation and bookmark registration."""


from typing import Union

from fpdf import FPDF

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.pdf_writer._context import PDFWriterContext
from aspose.words_foss.pdf_writer.constants import MIN_HEADER_FOOTER_Y_MM


def _collect_hf_children(
    doc: ldm.Document, *, header: bool
) -> list[Union[ldm.Paragraph, ldm.Table]]:
    """Collect ordered children for header or footer rendering.

    Prefers the per-section ``HeaderFooter.children`` list (which preserves
    interleaved paragraph/table order) and falls back to the flat
    ``doc.header_paragraphs`` / ``doc.footer_paragraphs`` for backwards
    compatibility with LDMs that only carry paragraphs.
    """
    hf_type = 0 if header else 1
    for sec in doc.sections:
        for hf in sec.headers_footers:
            if hf.header_footer_type == hf_type and hf.children:
                return list(hf.children)
    paras = doc.header_paragraphs if header else doc.footer_paragraphs
    return list(paras)


def _has_hf_content(doc: ldm.Document, *, header: bool) -> bool:
    """Return True when the document has header or footer content."""
    if header and doc.header_paragraphs:
        return True
    if not header and doc.footer_paragraphs:
        return True
    hf_type = 0 if header else 1
    for sec in doc.sections:
        for hf in sec.headers_footers:
            if hf.header_footer_type == hf_type and hf.children:
                return True
    return False


def _hf_content_height(
    writer: PDFWriterContext, children: list[Union[ldm.Paragraph, ldm.Table]]
) -> float:
    """Height in mm that *children* occupy once rendered in a band."""
    usable_w = writer._page_width - writer._page_margin_left - writer._page_margin_right
    return sum(writer._estimate_child_height(child, usable_w) for child in children)


def _render_hf_children(
    pdf: FPDF,
    writer: PDFWriterContext,
    children: list[Union[ldm.Paragraph, ldm.Table]],
    *,
    render_positioned: bool = False,
) -> None:
    """Render an ordered list of paragraphs and tables in a header/footer band."""
    paras_for_shapes: list[ldm.Paragraph] = []
    for child in children:
        if isinstance(child, ldm.Paragraph):
            para_start_y = pdf.get_y()
            writer._paragraph_renderer.render_paragraph(pdf, child)
            if render_positioned:
                writer._shape_renderer.render_positioned_shapes_in(
                    pdf, [child], line_y_override=para_start_y
                )
            paras_for_shapes.append(child)
        elif isinstance(child, ldm.Table):
            writer._table_renderer.render_table(pdf, child)
    if not render_positioned and paras_for_shapes:
        writer._shape_renderer.render_positioned_shapes_in(pdf, paras_for_shapes)


def install_page_header(
    pdf: FPDF,
    doc: ldm.Document,
    writer: PDFWriterContext,
    skip_first_page: bool,
    header_y_mm: float,
) -> None:
    """Install a per-page header callback on *pdf*."""
    if not _has_hf_content(doc, header=True):
        return

    children = _collect_hf_children(doc, header=True)

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
            _render_hf_children(pdf, writer, children, render_positioned=True)
            final_y = max(pdf.get_y(), body_start_y)
            pdf.set_xy(writer._page_margin_left, final_y)
        finally:
            pdf.set_auto_page_break(auto=prev_auto, margin=prev_bottom)
            pdf._in_header_render = False  # type: ignore[attr-defined]

    pdf.header = _header_callback  # type: ignore[method-assign]


def install_page_footer(
    pdf: FPDF,
    doc: ldm.Document,
    writer: PDFWriterContext,
    skip_first_page: bool,
    footer_y_mm: float,
) -> None:
    """Install a per-page footer callback on *pdf*."""
    if not _has_hf_content(doc, header=False):
        return

    children = _collect_hf_children(doc, header=False)
    content_h_mm = _hf_content_height(writer, children)

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
            band_bottom = max(
                writer._page_height - max(footer_y_mm, MIN_HEADER_FOOTER_Y_MM),
                writer._page_height - writer._page_margin_bottom,
            )
            band_top = max(band_bottom - content_h_mm, MIN_HEADER_FOOTER_Y_MM)
            pdf.set_xy(writer._page_margin_left, band_top)
            _render_hf_children(pdf, writer, children)
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
    for extra in para._children:
        if not isinstance(extra, ldm.BookmarkStart) or not extra.name:
            continue
        link_id = anchor_links.get(extra.name)
        if link_id is None:
            link_id = pdf.add_link()
            anchor_links[extra.name] = link_id
        pdf.set_link(link_id, y=pdf.get_y(), page=pdf.page_no())
