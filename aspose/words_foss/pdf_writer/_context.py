"""Structural protocol for the PDF sub-renderers' ``writer`` parameter.

Lets the sub-renderer modules type their ``writer`` without importing
:class:`LdmPdfWriter` (which would form a runtime circular import).
``LdmPdfWriter`` satisfies it structurally.
"""

from typing import Any, Optional, Protocol, Union

from fpdf import FPDF

from aspose.words_foss.saving import PdfSaveOptions


class PDFWriterContext(Protocol):
    options: PdfSaveOptions

    _page_width: float
    _page_height: float
    _page_margin_left: float
    _page_margin_right: float
    _page_margin_bottom: float
    _page_number_offset: int

    _anchor_links: dict[str, int]
    _pre_rendered_shapes: set[int]

    _paragraph_renderer: Any
    _shape_renderer: Any
    _table_renderer: Any

    def _link_target_for(self, pdf: FPDF, url: Optional[str]) -> Union[int, str]: ...
