"""Constants for the PDF writer module.

Centralises all magic numbers, regex patterns, and lookup dicts used
across the pdf_writer sub-modules.
"""


import re

from aspose.words_foss.model.enums import ParagraphAlignment
from aspose.words_foss.saving import PdfCompliance

# ---------------------------------------------------------------------------
# Page layout
# ---------------------------------------------------------------------------

A4_WIDTH_MM = 210.0
A4_HEIGHT_MM = 297.0
DEFAULT_MARGIN_MM = 20.0

# ---------------------------------------------------------------------------
# Font / text sizing
# ---------------------------------------------------------------------------

DEFAULT_FONT_NAME = "Helvetica"
DEFAULT_FONT_SIZE_PT = 11.0

# Word font name → fpdf2 core family (case-insensitive substring match).
FPDF_FONT_FAMILY_MAP: tuple[tuple[tuple[str, ...], str], ...] = (
    (("courier", "consolas", "menlo", "monaco", "monospace", "lucida console",
      "andale mono", "dejavu sans mono", "liberation mono"), "Courier"),
    (("times", "serif", "georgia", "cambria", "garamond", "palatino",
      "bookman", "minion", "century", "constantia", "didot",
      "liberation serif", "dejavu serif", "freeserif", "noto serif"), "Times"),
    (("symbol",), "Symbol"),
    (("wingdings", "webdings", "dingbat", "zapf"), "ZapfDingbats"),
    (("helvetica", "arial", "sans", "calibri", "carlito", "verdana",
      "tahoma", "geneva", "trebuchet", "lucida sans", "segoe",
      "open sans", "roboto", "liberation sans", "dejavu sans",
      "noto sans", "freesans"), "Helvetica"),
)

# Points to mm conversion factor
PT_TO_MM = 0.352778

# x-height per 1000-em. Used to rescale a run's point size when the requested
# font is substituted with a different fpdf2 core family, so the rendered
# x-height stays visually close to the source. Most specific triggers first —
# first match wins.
#
# Sources:
#   * PostScript core 14 — XHeight from Adobe AFM files
#     (https://github.com/tecnickcom/tc-font-core14-afms).
#   * Common system fonts — measured ratios at
#     https://jkorpela.fi/x-height.html (value * 1000).
# Only fonts with a published value are listed; for unknown names the lookup
# falls through to ratio = 1.0 (no scaling).
FONT_X_HEIGHT_MAP: tuple[tuple[tuple[str, ...], int], ...] = (
    # Serif
    (("times new roman",), 448),
    (("georgia",), 481),
    (("cambria",), 466),
    (("times",), 450),
    # Monospace
    (("courier new",), 423),
    (("courier",), 426),
    # Sans-serif
    (("arial",), 519),
    (("calibri",), 466),
    (("verdana",), 545),
    (("tahoma",), 545),
    (("trebuchet",), 523),
    (("comic sans",), 532),
    (("helvetica",), 523),
)

# x-height of fpdf2's three text core families (PostScript AFM XHeight).
CORE_FAMILY_X_HEIGHT: dict[str, int] = {
    "Helvetica": 523,
    "Times": 450,
    "Courier": 426,
}

# Line-height multiplier for natural leading (size_pt * PT_TO_MM * this).
LINE_HEIGHT_FACTOR = 1.4

# Word stores line-spacing multiples in twentieths of a line.
WORD_LINE_SPACING_DIVISOR = 12.0

# ---------------------------------------------------------------------------
# Rendering — shapes / images
# ---------------------------------------------------------------------------

DEFAULT_SHAPE_DIM_PT = 100.0
TEXTBOX_INNER_PAD_MM = 3.0
POST_IMAGE_SPACING_MM = 2.0
POST_TABLE_SPACING_MM = 2.0

# ---------------------------------------------------------------------------
# Rendering — lines / borders
# ---------------------------------------------------------------------------

DEFAULT_LINE_WIDTH_MM = 0.2
MIN_LINE_WIDTH_MM = 0.1

# ---------------------------------------------------------------------------
# Rendering — strikethrough
# ---------------------------------------------------------------------------

STRIKETHROUGH_Y_RATIO = 0.55

# ---------------------------------------------------------------------------
# Rendering — highlight
# ---------------------------------------------------------------------------

HIGHLIGHT_Y_OFFSET_RATIO = 0.15
HIGHLIGHT_HEIGHT_RATIO = 0.75

# ---------------------------------------------------------------------------
# Rendering — code blocks / quotes / hyperlinks
# ---------------------------------------------------------------------------

CODE_BLOCK_BG_RGB = (245, 245, 245)
QUOTE_TEXT_RGB = (100, 100, 100)
HYPERLINK_TEXT_RGB = (0, 0, 238)

# ---------------------------------------------------------------------------
# Rendering — lists / indentation
# ---------------------------------------------------------------------------

DEFAULT_QUOTE_INDENT_MM = 10.0
LIST_INDENT_PER_LEVEL_MM = 5.0

# ---------------------------------------------------------------------------
# Rendering — tables
# ---------------------------------------------------------------------------

# mm per pt of font size. ~0.42 ≈ font-height (PT_TO_MM ≈ 0.353) × 1.2
# line spacing, matching Word's tight default.
DEFAULT_CELL_LINE_H_FACTOR = 0.42
MIN_ROW_HEIGHT_FACTOR = 0.5
CHAR_WIDTH_ESTIMATE_FACTOR = 0.25
DEFAULT_CELL_PAD_LEFT_MM = 1.0
DEFAULT_CELL_PAD_TOP_MM = 0.5

# ---------------------------------------------------------------------------
# Header / footer
# ---------------------------------------------------------------------------

MIN_HEADER_FOOTER_Y_MM = 5.0

# ---------------------------------------------------------------------------
# Gap detection (footer two-column layout)
# ---------------------------------------------------------------------------

GAP_MIN_SPACES = 10

# ---------------------------------------------------------------------------
# Regex patterns
# ---------------------------------------------------------------------------

# Parse color strings like "Color [A=255, R=0, G=125, B=164]"
COLOR_RE = re.compile(
    r"Color\s*\[\s*A\s*=\s*(\d+)\s*,\s*R\s*=\s*(\d+)\s*,\s*G\s*=\s*(\d+)\s*,\s*B\s*=\s*(\d+)\s*\]"
)

# Hex color strings like "#007DA4" or "007DA4"
HEX_COLOR_RE = re.compile(r"^#?([0-9A-Fa-f]{6})$")

# Inline Markdown link syntax: "[display](url)"
INLINE_LINK_RE = re.compile(r"\[([^\]]+)\]\(([^)]+)\)")

# ---------------------------------------------------------------------------
# Lookup dicts
# ---------------------------------------------------------------------------

# Map PdfCompliance values to PDF version strings understood by fpdf2.
COMPLIANCE_TO_VERSION: dict[str, str] = {
    PdfCompliance.PDF17: "1.7",
    PdfCompliance.PDF20: "2.0",
    PdfCompliance.PDF_A1A: "1.4",
    PdfCompliance.PDF_A1B: "1.4",
    PdfCompliance.PDF_A2A: "1.7",
    PdfCompliance.PDF_A2U: "1.7",
    PdfCompliance.PDF_A4: "2.0",
    PdfCompliance.PDF_UA1: "1.7",
}

# fpdf2 alignment mapping (keyed by ParagraphAlignment values)
FPDF_ALIGN = {
    ParagraphAlignment.LEFT: "L",
    ParagraphAlignment.CENTER: "C",
    ParagraphAlignment.RIGHT: "R",
    ParagraphAlignment.JUSTIFY: "J",
}
