"""Constants for DOCX writing.

Mirrors the namespace URIs and unit conversion factors used by the
reader so that round-trips share a single source of truth.
"""


# Namespace URIs (no Clark-bracket form — we emit qualified element
# names like ``w:p`` so the prefix → URI mapping lives at the document
# level via ``xmlns`` attributes).
W_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_URI = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_RELS_URI = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_URI = "http://schemas.openxmlformats.org/package/2006/content-types"
# Markup-compatibility namespace — required on the document root
# whenever the body uses post-2007 OOXML extensions (``wps``, ``wpg``,
# ``w14``, etc.) so MS Word's strict loader honours ``mc:Ignorable``
# and accepts unknown-prefix elements gracefully instead of rejecting
# the package.
MC_URI = "http://schemas.openxmlformats.org/markup-compatibility/2006"
# Legacy VML namespaces — used inside ``<mc:Fallback><w:pict>...`` for
# wps:wsp shapes so MS Word's strict loader has a renderable
# alternative when it doesn't process the DrawingML Choice.
V_URI = "urn:schemas-microsoft-com:vml"
O_URI = "urn:schemas-microsoft-com:office:office"

# Hyperlink relationship type — used both inside the package
# relationship file and on the ``w:hyperlink`` element.
REL_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
REL_OFFICE_DOCUMENT = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
)
REL_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
REL_NUMBERING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
REL_HEADER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
REL_FOOTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"

# Content types
CT_DOCUMENT = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
CT_STYLES = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
CT_NUMBERING = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
CT_HEADER = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
CT_FOOTER = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
CT_SETTINGS = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
CT_FONT_TABLE = "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"
CT_CORE_PROPS = "application/vnd.openxmlformats-package.core-properties+xml"
CT_EXT_PROPS = "application/vnd.openxmlformats-officedocument.extended-properties+xml"
CT_RELS = "application/vnd.openxmlformats-package.relationships+xml"

REL_SETTINGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
REL_FONT_TABLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
REL_CORE_PROPS = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
REL_EXT_PROPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"

# Unit conversions
TWIPS_PER_PT = 20.0
HALF_PT_PER_PT = 2.0
EIGHTHS_PT_PER_PT = 8.0  # border sizes are in 1/8 of a point


def pt_to_twips(pt: float) -> int:
    """Points → twips (1/20 of a point)."""
    return int(round(pt * TWIPS_PER_PT))


def pt_to_half_pt(pt: float) -> int:
    """Points → half-points (used for font size and border widths)."""
    return int(round(pt * HALF_PT_PER_PT))


def pt_to_eighths(pt: float) -> int:
    """Points → eighths of a point (border ``w:sz``)."""
    return max(1, int(round(pt * EIGHTHS_PT_PER_PT)))


# OOXML alignment value mapping.  Indices match
# ``aspose.words_foss.model.enums.ParagraphAlignment``.
ALIGNMENT_VAL = {
    0: "left",
    1: "center",
    2: "right",
    3: "both",
    4: "distribute",
}

# Highlight color reverse-map (LDM color string → OOXML highlight name).
# Only the eight colors Word treats as highlights are emitted; anything
# else falls through to a shading-based emulation handled at the rPr level.
HIGHLIGHT_NAME_BY_COLOR: dict[str, str] = {
    "Color [A=255, R=255, G=255, B=0]": "yellow",
    "Color [A=255, R=0, G=255, B=0]": "green",
    "Color [A=255, R=0, G=255, B=255]": "cyan",
    "Color [A=255, R=255, G=0, B=255]": "magenta",
    "Color [A=255, R=0, G=0, B=255]": "blue",
    "Color [A=255, R=255, G=0, B=0]": "red",
    "Color [A=255, R=0, G=0, B=128]": "darkBlue",
    "Color [A=255, R=0, G=128, B=128]": "darkCyan",
    "Color [A=255, R=0, G=128, B=0]": "darkGreen",
    "Color [A=255, R=128, G=0, B=128]": "darkMagenta",
    "Color [A=255, R=128, G=0, B=0]": "darkRed",
    "Color [A=255, R=128, G=128, B=0]": "darkYellow",
    "Color [A=255, R=128, G=128, B=128]": "darkGray",
    "Color [A=255, R=192, G=192, B=192]": "lightGray",
    "Color [A=255, R=0, G=0, B=0]": "black",
    "Color [A=255, R=255, G=255, B=255]": "white",
}

# Underline integer (see ``model.enums.Underline``) → OOXML w:val token.
UNDERLINE_VAL = {
    0: None,  # NONE — emit no element
    1: "single",
    2: "words",
    3: "double",
    4: "dotted",
    5: "thick",
    6: "dash",
    7: "dotDash",
    8: "dotDotDash",
    9: "wavy",
}

# Border line style int → OOXML token.  Indices match
# ``aspose.words_foss.model.enums.LineStyle`` so ``writer → reader``
# resolves to the same enum value the source carried.
LINE_STYLE_VAL = {
    0: "none",
    1: "single",
    2: "thick",
    3: "double",
    5: "hairline",
    6: "dotted",
    7: "dashed",
    8: "dotDash",
    9: "dotDotDash",
    10: "triple",
    11: "thinThickSmallGap",
    12: "thickThinSmallGap",
    13: "thinThickThinSmallGap",
    14: "thinThickMediumGap",
    15: "thickThinMediumGap",
    16: "thinThickThinMediumGap",
    17: "thinThickLargeGap",
    18: "thickThinLargeGap",
    19: "thinThickThinLargeGap",
    20: "wave",
    21: "doubleWave",
    22: "dashSmallGap",
    23: "dashDotStroked",
    24: "threeDEmboss",
    25: "threeDEngrave",
    26: "outset",
    27: "inset",
}

# Tab stop alignment int → OOXML w:val token.
TAB_ALIGNMENT_VAL = {
    0: "left",
    1: "center",
    2: "right",
    3: "decimal",
    4: "bar",
    5: "num",
    6: "clear",
}

# Tab stop leader int → OOXML w:leader token.
TAB_LEADER_VAL = {
    0: "none",
    1: "dot",
    2: "hyphen",
    3: "underscore",
    4: "heavy",
    5: "middleDot",
}
