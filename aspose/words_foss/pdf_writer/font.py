"""Font application and reset helpers for the PDF writer."""


from fpdf import FPDF

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.pdf_writer.color import set_text_color
from aspose.words_foss.pdf_writer.constants import (
    CORE_FAMILY_X_HEIGHT,
    DEFAULT_FONT_NAME,
    DEFAULT_FONT_SIZE_PT,
    FONT_X_HEIGHT_MAP,
    FPDF_FONT_FAMILY_MAP,
)


def _resolve_fpdf_family(font_name: str) -> str:
    """Pick the closest fpdf2 core family for an LDM font name."""
    name_lower = font_name.lower()
    if not name_lower:
        return DEFAULT_FONT_NAME
    for triggers, family in FPDF_FONT_FAMILY_MAP:
        if any(t in name_lower for t in triggers):
            return family
    return DEFAULT_FONT_NAME


def _source_xheight(name_lower: str) -> int | None:
    for triggers, xheight in FONT_X_HEIGHT_MAP:
        if any(t in name_lower for t in triggers):
            return xheight
    return None


def _size_ratio(name_lower: str, target_family: str) -> float:
    """Scale factor to apply when rendering ``name_lower`` as ``target_family``.

    Matches the source font's x-height against the target family's, so the
    rendered glyph height stays visually close to the requested font.
    """
    src = _source_xheight(name_lower)
    tgt = CORE_FAMILY_X_HEIGHT.get(target_family)
    if src is None or not tgt:
        return 1.0
    return src / tgt


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

    name_lower = (font.name or "").lower()
    font_name = _resolve_fpdf_family(name_lower)
    size *= _size_ratio(name_lower, font_name)

    pdf.set_font(font_name, style=style, size=size)

    # Apply text color
    set_text_color(pdf, font.color)

    return style


def reset_font(pdf: FPDF) -> None:
    """Reset font to defaults."""
    pdf.set_font(DEFAULT_FONT_NAME, size=DEFAULT_FONT_SIZE_PT)
    pdf.set_text_color(0, 0, 0)
