"""Color parsing and setting helpers for the PDF writer."""


from typing import Optional, Tuple

from fpdf import FPDF

from aspose.words_foss.pdf_writer.constants import COLOR_RE, HEX_COLOR_RE


def parse_color(color_str: str) -> Optional[Tuple[int, int, int]]:
    """Parse a color string into (R, G, B) tuple, or None if unparseable/empty."""
    if not color_str or "Empty" in color_str:
        return None
    m = COLOR_RE.search(color_str)
    if m:
        return (int(m.group(2)), int(m.group(3)), int(m.group(4)))
    m = HEX_COLOR_RE.match(color_str.strip())
    if m:
        h = m.group(1)
        return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    return None


def set_text_color(pdf: FPDF, color_str: str) -> None:
    """Set text color from an LDM color string."""
    rgb = parse_color(color_str)
    if rgb:
        pdf.set_text_color(*rgb)
    else:
        pdf.set_text_color(0, 0, 0)


def set_fill_color(pdf: FPDF, color_str: str) -> bool:
    """Set fill color from an LDM color string. Returns True if a color was set."""
    rgb = parse_color(color_str)
    if rgb:
        pdf.set_fill_color(*rgb)
        return True
    return False
