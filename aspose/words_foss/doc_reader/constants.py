"""
Constants for Word 97-2003 Binary Format (.doc) parsing.

FIB field indices, SPRM codes, Escher record types, and color tables.
"""

from __future__ import annotations

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
SPRM_SDYAHDRTOP = 0xB017  # Header distance (twips)
SPRM_SDYAHDRBOTTOM = 0xB018  # Footer distance (twips)

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
SPRM_PFKEEPFOLLOW = 0x2406  # Keep with next
SPRM_PFPAGEBREAKBEFORE = 0x2407  # Page break before
SPRM_POUTLVL = 0x2640  # Outline level

# =============================================================================
# Character SPRM codes
# =============================================================================

SPRM_CFBOLD = 0x0835  # Character bold
SPRM_CFITALIC = 0x0836  # Character italic
SPRM_CFSTRIKE = 0x0837  # Character strikethrough
SPRM_CKUL = 0x2A3E  # Character underline type
SPRM_CHPS = 0x4A43  # Font size (half-points)
SPRM_CRGFTC0 = 0x4A4F  # Font index (ASCII/Latin)
SPRM_CRGFTC2 = 0x4A51  # Font index (non-FE, non-ASCII)
SPRM_CICO = 0x2A42  # Indexed color (Ico)
SPRM_CCV = 0x6870  # COLORREF (R, G, B, flags)
SPRM_CHIGHLIGHT = 0x2A0C  # Highlight color (Ico)
SPRM_CISS = 0x2A48  # Superscript/subscript
SPRM_CFCAPS = 0x083B  # All caps
SPRM_CFSMALLCAPS = 0x083A  # Small caps
SPRM_CFVANISH = 0x083C  # Hidden text
SPRM_CISTD = 0x4A30  # Character style index

# =============================================================================
# Ico color table — MS-DOC indexed colors
# =============================================================================

ICO_COLORS: dict[int, str] = {
    0x00: "",  # auto
    0x01: "#000000",  # black
    0x02: "#0000FF",  # blue
    0x03: "#00FFFF",  # cyan
    0x04: "#00FF00",  # green
    0x05: "#FF00FF",  # magenta
    0x06: "#FF0000",  # red
    0x07: "#FFFF00",  # yellow
    0x08: "#FFFFFF",  # white
    0x09: "#000080",  # darkBlue
    0x0A: "#008080",  # darkCyan
    0x0B: "#008000",  # darkGreen
    0x0C: "#800080",  # darkMagenta
    0x0D: "#800000",  # darkRed
    0x0E: "#808000",  # darkYellow
    0x0F: "#808080",  # darkGray
    0x10: "#C0C0C0",  # lightGray
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

# FOpt property ID for blip reference
FOPT_PID_PIB = 0x0104  # Blip index (1-based into BSE array)
