"""
Constants for Word 97-2003 Binary Format (.doc) parsing.

FIB field indices, SPRM codes, Escher record types, and color tables.
"""

# =============================================================================
# FIB FIELD INDICES in FibRgFcLcb97
# =============================================================================

IDX_STSHF = 1  # Style sheet
IDX_PLCFSED = 6  # Section descriptor table (PlcfSed)
IDX_PLCFHDD = 11  # Header/footer text positions (PlcfHdd)
IDX_PLCFBTECHPX = 12  # Character property bin table
IDX_PLCFBTEPAPX = 13  # Paragraph property bin table
IDX_STTBFFFN = 15  # Font table (SttbfFfn)
IDX_CLX = 33  # Complex (piece table)
IDX_PLCSPAMOM = 40  # OfficeArt shape anchors — main document
IDX_PLCSPAHDR = 41  # OfficeArt shape anchors — headers
IDX_DGGINFO = 50  # OfficeArt Drawing Group + Drawing Containers
IDX_PLCFTXBXTXT = 62  # Textbox text positions + FTXBXS shape mapping
IDX_PLFLST = 73  # List table
IDX_PLFLFO = 74  # List format overrides

# =============================================================================
# Section SPRM codes (from SEPX)
# =============================================================================

SPRM_SFTITLEPAGE = 0x300A  # Different first page header/footer
SPRM_SFPGNRESTART = 0x3011  # Restart page numbering
SPRM_SPGNSTART = 0x5001  # Starting page number (short)
SPRM_SXAPAGE = 0xB01F  # Page width (twips)
SPRM_SYAPAGE = 0xB020  # Page height (twips)
SPRM_SDXALEFT = 0xB021  # Left margin (twips)
SPRM_SDXARIGHT = 0xB022  # Right margin (twips)
SPRM_SDYATOP = 0x9023  # Top margin (twips, signed)
SPRM_SDYABOTTOM = 0x9024  # Bottom margin (twips, signed)
SPRM_SDZAGUTTER = 0xB025  # Gutter (binding margin) — twips
SPRM_SDYAHDRTOP = 0xB017  # Header distance (twips)
SPRM_SDYAHDRBOTTOM = 0xB018  # Footer distance (twips)
SPRM_SFEVENLYSPCOLS = 0x3005  # Evenly spaced columns flag
SPRM_SCCOLUMNS = 0x500B  # Number of columns minus 1
SPRM_SDXACOLUMNS = 0x900C  # Spacing between columns (twips)
SPRM_SFLBETWEEN = 0x3019  # Line between columns flag
SPRM_SBKC = 0x3009  # Section break type (0=continuous, 1=newColumn, 2=newPage, etc.)

# =============================================================================
# Paragraph SPRM codes
# =============================================================================

SPRM_PILFO = 0x460B  # Paragraph list format override index
SPRM_PILVL = 0x260A  # Paragraph indent level
SPRM_PJC80 = 0x2403  # Paragraph alignment (legacy, physical)
SPRM_PJC = 0x2461  # Paragraph alignment (logical)
SPRM_PDXALEFT80 = 0x840F  # Left indent (legacy) — twips
SPRM_PDXALEFT = 0x845E  # Left indent — twips
SPRM_PDXARIGHT80 = 0x840E  # Right indent (legacy) — twips
SPRM_PDXARIGHT = 0x845D  # Right indent — twips
SPRM_PDXALEFT1_80 = 0x8411  # First line indent (legacy) — twips
SPRM_PDXALEFT1 = 0x8460  # First line indent — twips
SPRM_PDYABEFORE = 0xA413  # Space before — twips
SPRM_PDYAAFTER = 0xA414  # Space after — twips
SPRM_PFDYABEFOREAUTO = 0x245B  # Space before auto
SPRM_PFDYAAFTERAUTO = 0x245C  # Space after auto
SPRM_PDYALINE = 0x6412  # Line spacing (LSPD)
SPRM_PFKEEP = 0x2405  # Keep paragraph together (keepLines / keep_together)
SPRM_PFKEEPFOLLOW = 0x2406  # Keep with next
SPRM_PFPAGEBREAKBEFORE = 0x2407  # Page break before
SPRM_PFINTABLE = 0x2416  # Paragraph belongs to a table (cell content)
SPRM_PFTTP = 0x2417  # Paragraph is a table-trailing (row-end) marker
SPRM_PFNOLINENUMB = 0x240C  # Suppress line numbers
SPRM_PFNOAUTOHYPH = 0x242A  # Suppress automatic hyphenation
SPRM_PFWIDOWCONTROL = 0x2431  # Widow / orphan control
SPRM_PFAUTOSPACEDE = 0x2437  # Add space between far-east and alphabetic text
SPRM_PFAUTOSPACEDN = 0x2438  # Add space between far-east and digits
SPRM_PFADJUSTRIGHT = 0x2448  # Auto-adjust right indent for far-east lists
SPRM_PFCONTEXTUALSPACING = 0x246D  # contextualSpacing (no space between same-style paras)
SPRM_POUTLVL = 0x2640  # Outline level
SPRM_PWALIGNFONT = 0x4439  # Vertical text alignment within a line (baseline)
SPRM_PCHGTABSPAPX = 0xC615  # Tab stop changes (add/delete)
SPRM_PCHGTABS = 0xC60D  # Tab stop changes (legacy)
SPRM_PCNF = 0xC666  # Conditional-style mask (table banded rows/cols)
# Indirection — full grpprl too big for an FKP slot, stored in the Data stream.
SPRM_PHUGEPAPX = 0x6646
# Paragraph borders (BRC80, 4-byte operand)
SPRM_PBRCTOP80 = 0x6424
SPRM_PBRCLEFT80 = 0x6425
SPRM_PBRCBOTTOM80 = 0x6426
SPRM_PBRCRIGHT80 = 0x6427
# Paragraph shading (Shd80, 2-byte operand: ipat:6 | icoFore:5 | icoBack:5)
SPRM_PSHD80 = 0x442D

# =============================================================================
# Character SPRM codes
# =============================================================================

SPRM_CFBOLD = 0x0835  # Character bold
SPRM_CFITALIC = 0x0836  # Character italic
SPRM_CFSTRIKE = 0x0837  # Character strikethrough
SPRM_CFOUTLINE = 0x0838  # Outline character effect
SPRM_CFSHADOW = 0x0839  # Shadow character effect
SPRM_CKUL = 0x2A3E  # Character underline type
SPRM_CHPS = 0x4A43  # Font size (half-points)
SPRM_CHPSKERN = 0x484B  # Minimum size to apply kerning (half-points)
SPRM_CRGFTC0 = 0x4A4F  # Font index (ASCII/Latin)
SPRM_CRGFTC2 = 0x4A51  # Font index (non-FE, non-ASCII)
SPRM_CICO = 0x2A42  # Indexed color (Ico)
SPRM_CCV = 0x6870  # COLORREF (R, G, B, flags)
SPRM_CHIGHLIGHT = 0x2A0C  # Highlight color (Ico)
SPRM_CISS = 0x2A48  # Superscript/subscript
SPRM_CFCAPS = 0x083B  # All caps
SPRM_CFSMALLCAPS = 0x083A  # Small caps
SPRM_CFVANISH = 0x083C  # Hidden text
SPRM_CFIMPRINT = 0x0854  # Imprint (engrave) character effect
SPRM_CFEMBOSS = 0x0858  # Emboss character effect
SPRM_CISTD = 0x4A30  # Character style index
SPRM_CFNOPROOF = 0x083D  # No proofing (suppress spell-check)
SPRM_CFBOLDBI = 0x085C  # Bold for complex script (BiDi) text
SPRM_CFITALICBI = 0x085D  # Italic for complex script (BiDi) text
SPRM_CFSPEC = 0x0855  # fSpec — special character (picture, symbol, etc.)
SPRM_CPICLOCATION = 0x6A03  # PicLocation — offset into Data stream for inline picture
SPRM_CRGFTC1 = 0x4A50  # Font index (Far East text)
SPRM_CRGLID0 = 0x486D  # Language ID (ASCII/Latin)
SPRM_CRGLID1 = 0x486E  # Language ID (Complex Script/BiDi)
SPRM_CRGLID2 = 0x486F  # Language ID (Far East)

# =============================================================================
# Ico color table — MS-DOC indexed colors
# =============================================================================

ICO_COLORS: dict[int, str] = {
    0x00: "",  # auto
    0x01: "Color [A=255, R=0, G=0, B=0]",  # black
    0x02: "Color [A=255, R=0, G=0, B=255]",  # blue
    0x03: "Color [A=255, R=0, G=255, B=255]",  # cyan
    0x04: "Color [A=255, R=0, G=255, B=0]",  # green
    0x05: "Color [A=255, R=255, G=0, B=255]",  # magenta
    0x06: "Color [A=255, R=255, G=0, B=0]",  # red
    0x07: "Color [A=255, R=255, G=255, B=0]",  # yellow
    0x08: "Color [A=255, R=255, G=255, B=255]",  # white
    0x09: "Color [A=255, R=0, G=0, B=128]",  # darkBlue
    0x0A: "Color [A=255, R=0, G=128, B=128]",  # darkCyan
    0x0B: "Color [A=255, R=0, G=128, B=0]",  # darkGreen
    0x0C: "Color [A=255, R=128, G=0, B=128]",  # darkMagenta
    0x0D: "Color [A=255, R=128, G=0, B=0]",  # darkRed
    0x0E: "Color [A=255, R=128, G=128, B=0]",  # darkYellow
    0x0F: "Color [A=255, R=128, G=128, B=128]",  # darkGray
    0x10: "Color [A=255, R=192, G=192, B=192]",  # lightGray
}

# =============================================================================
# Escher / OfficeArt record types
# =============================================================================

ESCHER_BSE = 0xF007
ESCHER_SP = 0xF00A
ESCHER_FOPT = 0xF00B
ESCHER_BLIP_PNG = 0xF01E
ESCHER_BLIP_JPEG = 0xF01D
ESCHER_BLIP_JPEG2 = 0xF021
# Vector / unsupported blip types — the raster writers can't decode them.
ESCHER_BLIP_EMF = 0xF01A
ESCHER_BLIP_WMF = 0xF01B
ESCHER_BLIP_PICT = 0xF01C
ESCHER_BLIP_DIB = 0xF01F
ESCHER_BLIP_TIFF = 0xF020

_RASTER_BLIP_TYPES: frozenset[int] = frozenset(
    (ESCHER_BLIP_PNG, ESCHER_BLIP_JPEG, ESCHER_BLIP_JPEG2)
)

# Magic-byte prefixes for raster cross-check (blip type can lie).
_RASTER_BLIP_MAGIC: tuple[bytes, ...] = (
    b"\x89PNG\r\n\x1a\n",
    b"\xff\xd8\xff",
    b"GIF87a",
    b"GIF89a",
    b"BM",
    b"II*\x00",
    b"MM\x00*",
)


def is_raster_blip(blip_type: int, image_bytes: bytes) -> bool:
    """True when *blip_type* and *image_bytes* describe a raster the writers can decode."""
    if not image_bytes:
        return False
    if blip_type in _RASTER_BLIP_TYPES:
        return True
    return any(image_bytes.startswith(prefix) for prefix in _RASTER_BLIP_MAGIC)

# =============================================================================
# Table SPRM codes
# =============================================================================

SPRM_TDEFTABLE = 0xD608  # Table definition (cell widths/boundaries)
SPRM_TTPC = 0x360D  # Table positioning control
SPRM_TTBLPY = 0x940F  # Table vertical position offset (tblpY, twips)
SPRM_TDXAFROMTEXT = 0x9410  # Table distance from text (leftFromText, twips)
SPRM_TDXAFROMTEXTRIGHT = 0x941E  # Table distance from text (rightFromText, twips)
SPRM_TTABLEWIDTH = 0xF614  # Table width
SPRM_TTABLEBORDERS = 0xD613  # Table borders (BRC, 8 bytes per side)
SPRM_TTABLEBORDERS80 = 0xD605  # Legacy table borders (BRC80, 4 bytes per side)
SPRM_TPROPREV = 0x6036  # Compound table positioning (contains tblpY at bytes[2:4])
SPRM_TCELLPADDINGDEFAULT = 0xD634  # Default cell margins for the table (tblCellMar)
SPRM_TCELLPADDING = 0xD632  # Per-cell-range margins override (tcMar)

# =============================================================================
# Built-in style STI values (sti → canonical English name)
# =============================================================================

STI_NAMES: dict[int, str] = {
    0: "Normal",
    65: "Default Paragraph Font",
    105: "Normal Table",
    107: "No List",
    34: "List Paragraph",
    # Note: Heading 1-9 (sti 1-9) and TOC 1-9 (sti 19-27) are handled separately in styles.py
}

# FOpt property ID for blip reference
FOPT_PID_PIB = 0x0104  # Blip index (1-based into BSE array)
