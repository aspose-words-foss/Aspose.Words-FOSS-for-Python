"""PDF writer module — backward-compatible public API.

Re-exports ``LdmPdfWriter`` and ``_parse_color`` so existing imports
of the form ``from aspose.words_foss.pdf_writer import LdmPdfWriter``
continue to work without changes.
"""

from aspose.words_foss.pdf_writer.color import parse_color as _parse_color
from aspose.words_foss.pdf_writer.writer import LdmPdfWriter

__all__ = ["LdmPdfWriter", "_parse_color"]
