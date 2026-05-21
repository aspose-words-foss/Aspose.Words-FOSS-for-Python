"""
Constants for DOCX XML parsing.

Extracts all namespace URIs, mapping dictionaries, and magic numbers
from the monolithic reader module into a single, importable location.
"""

from aspose.words_foss.model.enums import (
    CellVerticalAlignment as _CVA,
    LineSpacingRule as _LSRule,
    LineStyle as _LS,
    ParagraphAlignment as _PA,
    SectionStart as _SS,
    StyleType as _ST,
    Underline as _UL,
)

# =============================================================================
# XML NAMESPACES
# =============================================================================

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
WP_NS = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
PIC_NS = "{http://schemas.openxmlformats.org/drawingml/2006/picture}"
MC_NS = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
WPS_NS = "{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}"
WPG_NS = "{http://schemas.microsoft.com/office/word/2010/wordprocessingGroup}"

_PKG_RELS_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"

# =============================================================================
# UNIT CONVERSION
# =============================================================================

# 914400 EMU = 1 inch = 25.4 mm
_EMU_PER_MM = 914400.0 / 25.4

# 12700 EMU = 1 point
_EMU_PER_PT = 12700.0

# 20 twips = 1 point
_TWIPS_PER_PT = 20.0

# Half-point to point divisor
_HALF_PT_DIVISOR = 2.0

# Border size divisor (eighths of a point)
_BORDER_SIZE_DIVISOR = 8.0

# Points per inch
_POINTS_PER_INCH = 72.0

# Millimeters per inch
_MM_PER_INCH = 25.4

# Percentage divisor (1/50th of a percent for OOXML pct widths)
_PCT_DIVISOR = 50.0

# DrawingML modifier scale (1000ths of a percent)
_DML_MOD_SCALE = 100000.0

# Max color channel value
_MAX_COLOR_CHANNEL = 255

# Default tab stop in points
_DEFAULT_TAB_STOP_PT = 36.0

# Default page dimensions (mm) — A4 fallback
_DEFAULT_PAGE_HEIGHT_MM = 297.0

# Default margins (mm)
_DEFAULT_LEFT_MARGIN_MM = 19.05
_DEFAULT_TOP_MARGIN_MM = 25.4

# Paper size thresholds (points)
_LETTER_WIDTH_PT = 612.0
_LETTER_HEIGHT_PT = 792.0
_LEGAL_HEIGHT_PT = 1008.0
_A4_WIDTH_PT = 595.28
_A4_HEIGHT_PT = 841.89
_A3_WIDTH_PT = 841.89
_A3_HEIGHT_PT = 1190.55
_PAPER_SIZE_TOLERANCE_PT = 2.0
_PAPER_SIZE_TOLERANCE_A_PT = 3.0

# Paper size enum values
_PAPER_LETTER = 1
_PAPER_LEGAL = 5
_PAPER_A4 = 9
_PAPER_A3 = 8
_PAPER_CUSTOM = 0

# Outline level threshold for heading detection
_OUTLINE_LEVEL_BODY = 9

# =============================================================================
# SENTINEL VALUES
# =============================================================================

COLOR_EMPTY = "Color [Empty]"

# Null-bracketed sentinel for the Word ``PAGE`` field.  The reader emits a
# run containing this token in place of the field's cached literal result
# (e.g. "5") so the PDF writer can substitute it with the live page number
# at render time.  ``\x00`` cannot appear in legitimate run text and is
# therefore collision-free.
PAGE_FIELD_SENTINEL = "\x00PAGE\x00"

# =============================================================================
# MAPPING CONSTANTS
# =============================================================================

# DrawingML ``wps:bodyPr/@anchor`` → LDM ``vertical_alignment`` int.
_BODY_ANCHOR_MAP: dict[str, int] = {
    "t": _CVA.TOP,
    "ctr": _CVA.CENTER,
    "b": _CVA.BOTTOM,
}

_HIGHLIGHT_COLOR_MAP: dict[str, str] = {
    "yellow": "Color [A=255, R=255, G=255, B=0]",
    # Word distinguishes ``green`` (bright #00FF00) from ``darkGreen``
    # (#008000); a previous version of this table flattened both to
    # the dark variant which lost the highlight token on round-trip.
    "green": "Color [A=255, R=0, G=255, B=0]",
    "cyan": "Color [A=255, R=0, G=255, B=255]",
    "magenta": "Color [A=255, R=255, G=0, B=255]",
    "blue": "Color [A=255, R=0, G=0, B=255]",
    "red": "Color [A=255, R=255, G=0, B=0]",
    "darkBlue": "Color [A=255, R=0, G=0, B=128]",
    "darkCyan": "Color [A=255, R=0, G=128, B=128]",
    "darkGreen": "Color [A=255, R=0, G=128, B=0]",
    "darkMagenta": "Color [A=255, R=128, G=0, B=128]",
    "darkRed": "Color [A=255, R=128, G=0, B=0]",
    "darkYellow": "Color [A=255, R=128, G=128, B=0]",
    "darkGray": "Color [A=255, R=128, G=128, B=128]",
    "lightGray": "Color [A=255, R=192, G=192, B=192]",
    "black": "Color [A=255, R=0, G=0, B=0]",
    "white": "Color [A=255, R=255, G=255, B=255]",
}

_EXT_TO_CONTENT_TYPE: dict[str, str] = {
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "gif": "image/gif",
    "bmp": "image/bmp",
    "tiff": "image/tiff",
    "tif": "image/tiff",
    "webp": "image/webp",
}

_ALIGNMENT_MAP = {
    "left": _PA.LEFT,
    "start": _PA.LEFT,
    "center": _PA.CENTER,
    "right": _PA.RIGHT,
    "end": _PA.RIGHT,
    "both": _PA.JUSTIFY,
    "distribute": _PA.DISTRIBUTED,
}

_UNDERLINE_MAP = {
    "none": _UL.NONE,
    "single": _UL.SINGLE,
    "words": _UL.WORDS,
    "double": _UL.DOUBLE,
    "dotted": _UL.DOTTED,
    "thick": _UL.THICK,
    "dash": _UL.DASH,
    "dotDash": _UL.DOT_DASH,
    "dotDotDash": _UL.DOT_DOT_DASH,
    "wavy": _UL.WAVY,
}

_BORDER_STYLE_MAP = {
    "none": _LS.NONE,
    "nil": _LS.NONE,
    "single": _LS.SINGLE,
    "thick": _LS.THICK,
    "double": _LS.DOUBLE,
    "hairline": _LS.HAIRLINE,
    "dotted": _LS.DOT,
    "dashed": _LS.DASH_LARGE_GAP,
    "dotDash": _LS.DOT_DASH,
    "dotDotDash": _LS.DOT_DOT_DASH,
    "triple": _LS.TRIPLE,
    "thinThickSmallGap": _LS.THIN_THICK_SMALL_GAP,
    "thickThinSmallGap": _LS.THICK_THIN_SMALL_GAP,
    "thinThickThinSmallGap": _LS.THIN_THICK_THIN_SMALL_GAP,
    "thinThickMediumGap": _LS.THIN_THICK_MEDIUM_GAP,
    "thickThinMediumGap": _LS.THICK_THIN_MEDIUM_GAP,
    "thinThickThinMediumGap": _LS.THIN_THICK_THIN_MEDIUM_GAP,
    "thinThickLargeGap": _LS.THIN_THICK_LARGE_GAP,
    "thickThinLargeGap": _LS.THICK_THIN_LARGE_GAP,
    "thinThickThinLargeGap": _LS.THIN_THICK_THIN_LARGE_GAP,
    "wave": _LS.WAVE,
    "doubleWave": _LS.DOUBLE_WAVE,
    "dashSmallGap": _LS.DASH_SMALL_GAP,
    "dashDotStroked": _LS.DASH_DOT_STROKER,
    "threeDEmboss": _LS.EMBOSS_3D,
    "threeDEngrave": _LS.ENGRAVE_3D,
    "outset": _LS.OUTSET,
    "inset": _LS.INSET,
}

_STYLE_TYPE_MAP = {
    "paragraph": _ST.PARAGRAPH,
    "character": _ST.CHARACTER,
    "table": _ST.TABLE,
    "numbering": _ST.LIST,
}

_SECTION_START_MAP = {
    "continuous": _SS.CONTINUOUS,
    "newColumn": _SS.NEW_COLUMN,
    "newPage": _SS.NEW_PAGE,
    "evenPage": _SS.EVEN_PAGE,
    "oddPage": _SS.ODD_PAGE,
}

_LINE_RULE_MAP = {
    "atLeast": _LSRule.AT_LEAST,
    "exact": _LSRule.EXACTLY,
    "exactly": _LSRule.EXACTLY,
    "auto": _LSRule.MULTIPLE,
}

_TAB_ALIGNMENT_MAP = {
    "left": 0,
    "center": 1,
    "right": 2,
    "decimal": 3,
    "bar": 4,
    "list": 5,
    "clear": 6,
    "num": 5,
}

_TAB_LEADER_MAP = {
    "none": 0,
    "dot": 1,
    "hyphen": 2,
    "underscore": 3,
    "heavy": 4,
    "middleDot": 5,
}

_NUMBER_STYLE_MAP = {
    "decimal": 0,
    "upperRoman": 1,
    "lowerRoman": 2,
    "upperLetter": 3,
    "lowerLetter": 4,
    "ordinal": 5,
    "bullet": 23,
    "none": 255,
}

# Built-in Word style names: OOXML w:name → canonical display name.
# Built-in styles are resolved by their styleIdentifier to canonical
# English names, which may differ in casing or wording from the OOXML w:name.
# Built-in Word style names that we map to non-obvious canonical
# forms.  Entries whose canonical name can be derived by _canonicalize_style_name
# (title-case / acronym rules) are intentionally omitted.
_BUILTIN_STYLE_NAME_MAP: dict[str, str] = {
    "Normal Table": "Table Normal",
    "annotation reference": "Comment Reference",
    "annotation text": "Comment Text",
    "annotation subject": "Comment Subject",
    "Outline List 1": "1 / a / i",
    "Outline List 2": "1 / 1.1 / 1.1.1",
    "Outline List 3": "Article / Section",
    "macro": "Macro Text",
}

# Prefixes that are acronyms in canonical style names.
_ACRONYM_PREFIXES = ("toc", "toa")

# Words that stay lowercase in title-cased style names.
_LOWERCASE_WORDS = {"of", "and", "or", "the", "a", "an", "in", "on", "for", "to"}
