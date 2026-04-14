"""Font application and reset helpers for the PDF writer."""

from __future__ import annotations

from fpdf import FPDF

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.pdf_writer.color import set_text_color
from aspose.words_foss.pdf_writer.constants import (
    CALIBRI_COMPAT_FAMILIES,
    CALIBRI_TO_HELVETICA_RATIO,
    DEFAULT_FONT_NAME,
    DEFAULT_FONT_SIZE_PT,
)


def apply_run_font(pdf: FPDF, font: ldm.Font, default_size: float = DEFAULT_FONT_SIZE_PT) -> str:
    """Apply font properties from an LDM Font and return the fpdf2 style string."""
    style = ""
    if font.bold:
        style += "B"
    if font.italic:
        style += "I"
    if font.underline:
        style += "U"

    size = font.size if font.size > 0 else default_size

    # Use Courier for monospace font names
    font_name = DEFAULT_FONT_NAME
    name_lower = (font.name or "").lower()
    if name_lower and any(kw in name_lower for kw in ("courier", "mono", "consolas", "menlo")):
        font_name = "Courier"
    elif any(fam in name_lower for fam in CALIBRI_COMPAT_FAMILIES):
        # Helvetica has a bigger x-height than Calibri; shrink the
        # nominal size so substituted runs match the original's
        # visual footprint (header vs. footer balance in particular).
        size *= CALIBRI_TO_HELVETICA_RATIO

    pdf.set_font(font_name, style=style, size=size)

    # Apply text color
    set_text_color(pdf, font.color)

    return style


def reset_font(pdf: FPDF) -> None:
    """Reset font to defaults."""
    pdf.set_font(DEFAULT_FONT_NAME, size=DEFAULT_FONT_SIZE_PT)
    pdf.set_text_color(0, 0, 0)
