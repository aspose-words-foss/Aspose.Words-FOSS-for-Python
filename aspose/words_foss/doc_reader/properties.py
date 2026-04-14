"""
Paragraph and character property parsing (PAPX/CHPX) for DOC files.

Includes SPRM parsing, FKP page parsing, and SEPX section property parsing.
"""

from __future__ import annotations

import struct

from aspose.words_foss.doc_reader.constants import (
    ICO_COLORS,
    SPRM_CFBOLD,
    SPRM_CFCAPS,
    SPRM_CFITALIC,
    SPRM_CFSMALLCAPS,
    SPRM_CFSTRIKE,
    SPRM_CFVANISH,
    SPRM_CHIGHLIGHT,
    SPRM_CHPS,
    SPRM_CICO,
    SPRM_CISS,
    SPRM_CISTD,
    SPRM_CCV,
    SPRM_CKUL,
    SPRM_CRGFTC0,
    SPRM_PDXALEFT,
    SPRM_PDXALEFT1,
    SPRM_PDXALEFT1_80,
    SPRM_PDXALEFT80,
    SPRM_PDXARIGHT,
    SPRM_PDXARIGHT80,
    SPRM_PDYAAFTER,
    SPRM_PDYABEFORE,
    SPRM_PDYALINE,
    SPRM_PFDYAAFTERAUTO,
    SPRM_PFDYABEFOREAUTO,
    SPRM_PFKEEPFOLLOW,
    SPRM_PFPAGEBREAKBEFORE,
    SPRM_PILFO,
    SPRM_PILVL,
    SPRM_PJC,
    SPRM_PJC80,
    SPRM_POUTLVL,
    SPRM_SDXALEFT,
    SPRM_SDXARIGHT,
    SPRM_SDYABOTTOM,
    SPRM_SDYAHDRBOTTOM,
    SPRM_SDYAHDRTOP,
    SPRM_SDYATOP,
    SPRM_SFTITLEPAGE,
    SPRM_SFPGNRESTART,
    SPRM_SPGNSTART,
    SPRM_SXAPAGE,
    SPRM_SYAPAGE,
)
from aspose.words_foss.model.enums import LineSpacingRule

# =============================================================================
# Paragraph Properties
# =============================================================================


class ParaProps:
    """Properties extracted from PAPX for a single paragraph."""

    def __init__(self):
        self.istd: int = 0
        self.ilfo: int = 0  # List format override index (1-based, 0 = no list)
        self.ilvl: int = 0  # List indent level
        self.alignment: int = 0  # 0=Left, 1=Center, 2=Right, 3=Justify
        self.left_indent: float = 0.0  # points
        self.right_indent: float = 0.0  # points
        self.first_line_indent: float = 0.0  # points
        self.space_before: float = 0.0  # points
        self.space_after: float = 0.0  # points
        self.space_before_auto: bool = False
        self.space_after_auto: bool = False
        self.line_spacing: float = 12.0  # default: single spacing (12pt)
        self.line_spacing_rule: int = LineSpacingRule.MULTIPLE
        self.keep_with_next: bool = False
        self.page_break_before: bool = False
        self.outline_level: int = 9  # 9 = body text
        self._set_fields: set[str] = set()  # which fields were directly set


# =============================================================================
# Character Properties
# =============================================================================


class CharProps:
    """Properties extracted from CHPX for a text run."""

    def __init__(self):
        self.bold: bool = False
        self.italic: bool = False
        self.underline: int = 0  # underline type (0=none, 1=single, etc.)
        self.strikethrough: bool = False
        self.font_index: int = -1  # index into font table (-1 = default)
        self.font_size: float = 0.0  # points (0 = default/inherit)
        self.color: str = ""  # '#RRGGBB' or '' for auto
        self.highlight_color: str = ""  # '#RRGGBB' or '' for none
        self.superscript: bool = False
        self.subscript: bool = False
        self.all_caps: bool = False
        self.small_caps: bool = False
        self.hidden: bool = False
        self.style_index: int = -1  # character style index
        self._set_fields: set[str] = set()  # which fields were directly set
        self._toggle_fields: set[str] = set()  # fields set via 0x81 toggle SPRM


# =============================================================================
# SPRM Parsing
# =============================================================================


def parse_sprms(grpprl: bytes) -> dict[int, bytes]:
    """Parse a grpprl (group of property modifiers) into sprm -> operand map."""
    result: dict[int, bytes] = {}
    off = 0
    sizes = {0: 1, 1: 1, 2: 2, 3: 4, 4: 2, 5: 2, 7: 3}

    while off + 2 <= len(grpprl):
        sprm = struct.unpack_from("<H", grpprl, off)[0]
        spra = (sprm >> 13) & 0x07
        op_size = sizes.get(spra, 0)

        if spra == 6:  # Variable length
            if off + 2 < len(grpprl):
                op_size = grpprl[off + 2]
                if off + 3 + op_size <= len(grpprl):
                    result[sprm] = grpprl[off + 3 : off + 3 + op_size]
                off += 3 + op_size
                continue
            else:
                break

        if op_size > 0 and off + 2 + op_size <= len(grpprl):
            result[sprm] = grpprl[off + 2 : off + 2 + op_size]

        off += 2 + max(op_size, 0)
        if op_size == 0:
            break  # Unknown size, stop parsing

    return result


# =============================================================================
# PAPX FKP Parsing
# =============================================================================


def parse_papx_fkp(
    wd: bytes, pn: int, text_byte_offset: int, text_is_compressed: bool
) -> list[tuple[int, int, ParaProps]]:
    """Parse a PAPX FKP page and return (char_start, char_end, props) tuples."""
    page_offset = pn * 512
    if page_offset + 512 > len(wd):
        return []

    page = wd[page_offset : page_offset + 512]
    cpara = page[511]
    if cpara == 0:
        return []

    results: list[tuple[int, int, ParaProps]] = []

    for j in range(cpara):
        fc_j = struct.unpack_from("<I", page, j * 4)[0]
        fc_next = struct.unpack_from("<I", page, (j + 1) * 4)[0]

        # Convert FC to character positions
        if text_is_compressed:
            char_start = fc_j - text_byte_offset
            char_end = fc_next - text_byte_offset
        else:
            char_start = (fc_j - text_byte_offset) // 2
            char_end = (fc_next - text_byte_offset) // 2

        props = ParaProps()

        bx_offset = (cpara + 1) * 4 + j * 13
        if bx_offset >= 511:
            results.append((char_start, char_end, props))
            continue

        bpapx = page[bx_offset]
        if bpapx == 0:
            results.append((char_start, char_end, props))
            continue

        papx_pos = bpapx * 2
        if papx_pos >= 512:
            results.append((char_start, char_end, props))
            continue

        cb = page[papx_pos]
        if cb == 0:
            if papx_pos + 1 < 512:
                cb = page[papx_pos + 1]
                total = cb * 2
                if papx_pos + 2 + total <= 512 and total >= 2:
                    props.istd = struct.unpack_from("<H", page, papx_pos + 2)[0]
                    grpprl = page[papx_pos + 4 : papx_pos + 2 + total]
                    sprms = parse_sprms(grpprl)
                    apply_para_sprms(props, sprms)
        else:
            total = cb * 2
            if papx_pos + 1 + total <= 512 and total >= 2:
                props.istd = struct.unpack_from("<H", page, papx_pos + 1)[0]
                grpprl = page[papx_pos + 3 : papx_pos + 1 + total]
                sprms = parse_sprms(grpprl)
                apply_para_sprms(props, sprms)

        results.append((char_start, char_end, props))

    return results


def apply_para_sprms(props: ParaProps, sprms: dict[int, bytes]) -> None:
    """Apply paragraph sprms to ParaProps."""
    if SPRM_PILVL in sprms:
        props.ilvl = sprms[SPRM_PILVL][0]
    if SPRM_PILFO in sprms:
        props.ilfo = struct.unpack_from("<H", sprms[SPRM_PILFO])[0]

    # Alignment: prefer logical (0x2461) over legacy (0x2403)
    if SPRM_PJC in sprms:
        props.alignment = sprms[SPRM_PJC][0]
    elif SPRM_PJC80 in sprms:
        props.alignment = sprms[SPRM_PJC80][0]

    # Left indent (twips → points, /20)
    if SPRM_PDXALEFT in sprms:
        props.left_indent = struct.unpack_from("<h", sprms[SPRM_PDXALEFT])[0] / 20.0
    elif SPRM_PDXALEFT80 in sprms:
        props.left_indent = struct.unpack_from("<h", sprms[SPRM_PDXALEFT80])[0] / 20.0

    # Right indent
    if SPRM_PDXARIGHT in sprms:
        props.right_indent = struct.unpack_from("<h", sprms[SPRM_PDXARIGHT])[0] / 20.0
    elif SPRM_PDXARIGHT80 in sprms:
        props.right_indent = struct.unpack_from("<h", sprms[SPRM_PDXARIGHT80])[0] / 20.0

    # First line indent
    if SPRM_PDXALEFT1 in sprms:
        props.first_line_indent = struct.unpack_from("<h", sprms[SPRM_PDXALEFT1])[0] / 20.0
    elif SPRM_PDXALEFT1_80 in sprms:
        props.first_line_indent = struct.unpack_from("<h", sprms[SPRM_PDXALEFT1_80])[0] / 20.0

    # Space before/after (twips → points)
    if SPRM_PDYABEFORE in sprms:
        props.space_before = struct.unpack_from("<H", sprms[SPRM_PDYABEFORE])[0] / 20.0
    if SPRM_PDYAAFTER in sprms:
        props.space_after = struct.unpack_from("<H", sprms[SPRM_PDYAAFTER])[0] / 20.0

    # Space before/after auto
    if SPRM_PFDYABEFOREAUTO in sprms:
        props.space_before_auto = bool(sprms[SPRM_PFDYABEFOREAUTO][0])
    if SPRM_PFDYAAFTERAUTO in sprms:
        props.space_after_auto = bool(sprms[SPRM_PFDYAAFTERAUTO][0])

    # Line spacing (LSPD: dyaLine int16 + fMultLinespace int16)
    if SPRM_PDYALINE in sprms and len(sprms[SPRM_PDYALINE]) >= 4:
        dya_line = struct.unpack_from("<h", sprms[SPRM_PDYALINE], 0)[0]
        f_mult = struct.unpack_from("<h", sprms[SPRM_PDYALINE], 2)[0]
        if f_mult == 1:
            # Multiple: value / 240 gives the multiplier, LDM stores raw
            props.line_spacing_rule = LineSpacingRule.MULTIPLE
            props.line_spacing = dya_line / 20.0  # twips to points
        elif dya_line >= 0:
            props.line_spacing_rule = LineSpacingRule.AT_LEAST
            props.line_spacing = dya_line / 20.0
        else:
            props.line_spacing_rule = LineSpacingRule.EXACTLY
            props.line_spacing = abs(dya_line) / 20.0

    # Keep with next
    if SPRM_PFKEEPFOLLOW in sprms:
        props.keep_with_next = bool(sprms[SPRM_PFKEEPFOLLOW][0])

    # Page break before
    if SPRM_PFPAGEBREAKBEFORE in sprms:
        props.page_break_before = bool(sprms[SPRM_PFPAGEBREAKBEFORE][0])

    # Outline level
    if SPRM_POUTLVL in sprms:
        props.outline_level = sprms[SPRM_POUTLVL][0]

    # Track which fields were directly set
    props._set_fields = get_para_sprm_fields(sprms)


def get_para_sprm_fields(sprms: dict[int, bytes]) -> set[str]:
    """Return which paragraph fields were set by the given SPRMs."""
    fields: set[str] = set()
    mapping = {
        SPRM_PJC: "alignment",
        SPRM_PJC80: "alignment",
        SPRM_PDXALEFT: "left_indent",
        SPRM_PDXALEFT80: "left_indent",
        SPRM_PDXARIGHT: "right_indent",
        SPRM_PDXARIGHT80: "right_indent",
        SPRM_PDXALEFT1: "first_line_indent",
        SPRM_PDXALEFT1_80: "first_line_indent",
        SPRM_PDYABEFORE: "space_before",
        SPRM_PDYAAFTER: "space_after",
        SPRM_PFDYABEFOREAUTO: "space_before_auto",
        SPRM_PFDYAAFTERAUTO: "space_after_auto",
        SPRM_PDYALINE: "line_spacing",
        SPRM_PFKEEPFOLLOW: "keep_with_next",
        SPRM_PFPAGEBREAKBEFORE: "page_break_before",
        SPRM_POUTLVL: "outline_level",
    }
    for sprm_code in sprms:
        if sprm_code in mapping:
            fields.add(mapping[sprm_code])
    return fields


# =============================================================================
# CHPX FKP Parsing
# =============================================================================


def parse_chpx_fkp(
    wd: bytes, pn: int, text_byte_offset: int, text_is_compressed: bool
) -> list[tuple[int, int, CharProps]]:
    """Parse a CHPX FKP page and return (char_start, char_end, props) tuples."""
    page_offset = pn * 512
    if page_offset + 512 > len(wd):
        return []

    page = wd[page_offset : page_offset + 512]
    crun = page[511]
    if crun == 0:
        return []

    results: list[tuple[int, int, CharProps]] = []

    for j in range(crun):
        fc_j = struct.unpack_from("<I", page, j * 4)[0]
        fc_next = struct.unpack_from("<I", page, (j + 1) * 4)[0]

        if text_is_compressed:
            char_start = fc_j - text_byte_offset
            char_end = fc_next - text_byte_offset
        else:
            char_start = (fc_j - text_byte_offset) // 2
            char_end = (fc_next - text_byte_offset) // 2

        props = CharProps()

        # CHPX BX entry is just 1 byte (offset) per run
        bx_pos = (crun + 1) * 4 + j
        if bx_pos >= 511:
            results.append((char_start, char_end, props))
            continue

        chpx_offset_word = page[bx_pos]
        if chpx_offset_word == 0:
            results.append((char_start, char_end, props))
            continue

        chpx_pos = chpx_offset_word * 2
        if chpx_pos >= 512:
            results.append((char_start, char_end, props))
            continue

        cb = page[chpx_pos]
        if cb > 0 and chpx_pos + 1 + cb <= 512:
            grpprl = page[chpx_pos + 1 : chpx_pos + 1 + cb]
            sprms = parse_sprms(grpprl)
            apply_char_sprms(props, sprms)

        results.append((char_start, char_end, props))

    return results


def apply_toggle(current: bool, operand: int) -> bool:
    """Apply a ToggleOperand to a boolean property.

    0x00 = OFF, 0x01 = ON, 0x80 = inherit, 0x81 = toggle.
    For style parsing (no prior value), 0x81 effectively means ON.
    """
    if operand == 0x00:
        return False
    if operand == 0x01:
        return True
    if operand == 0x81:
        return not current
    return current  # 0x80 = inherit


def apply_char_sprms(props: CharProps, sprms: dict[int, bytes]) -> None:
    """Apply character sprms to CharProps."""
    if SPRM_CFBOLD in sprms:
        operand = sprms[SPRM_CFBOLD][0]
        props.bold = apply_toggle(props.bold, operand)
        if operand == 0x81:
            props._toggle_fields.add("bold")
    if SPRM_CFITALIC in sprms:
        operand = sprms[SPRM_CFITALIC][0]
        props.italic = apply_toggle(props.italic, operand)
        if operand == 0x81:
            props._toggle_fields.add("italic")
    if SPRM_CFSTRIKE in sprms:
        operand = sprms[SPRM_CFSTRIKE][0]
        props.strikethrough = apply_toggle(props.strikethrough, operand)
        if operand == 0x81:
            props._toggle_fields.add("strikethrough")
    if SPRM_CKUL in sprms:
        props.underline = sprms[SPRM_CKUL][0]

    # Font index (ASCII/Latin text)
    if SPRM_CRGFTC0 in sprms:
        props.font_index = struct.unpack_from("<H", sprms[SPRM_CRGFTC0])[0]

    # Font size (half-points → points)
    if SPRM_CHPS in sprms:
        props.font_size = struct.unpack_from("<H", sprms[SPRM_CHPS])[0] / 2.0

    # Color: prefer COLORREF (0x6870) over Ico (0x2A42)
    if SPRM_CCV in sprms and len(sprms[SPRM_CCV]) >= 4:
        r, g, b = sprms[SPRM_CCV][0], sprms[SPRM_CCV][1], sprms[SPRM_CCV][2]
        props.color = f"Color [A=255, R={r}, G={g}, B={b}]"
    elif SPRM_CICO in sprms:
        ico = sprms[SPRM_CICO][0]
        props.color = ICO_COLORS.get(ico, "")

    # Highlight color
    if SPRM_CHIGHLIGHT in sprms:
        ico = sprms[SPRM_CHIGHLIGHT][0]
        props.highlight_color = ICO_COLORS.get(ico, "")

    # Superscript/subscript (sprmCIss: 0=normal, 1=super, 2=sub)
    if SPRM_CISS in sprms:
        iss = sprms[SPRM_CISS][0]
        props.superscript = iss == 1
        props.subscript = iss == 2

    # All caps
    if SPRM_CFCAPS in sprms:
        props.all_caps = apply_toggle(props.all_caps, sprms[SPRM_CFCAPS][0])

    # Small caps
    if SPRM_CFSMALLCAPS in sprms:
        props.small_caps = apply_toggle(props.small_caps, sprms[SPRM_CFSMALLCAPS][0])

    # Hidden text
    if SPRM_CFVANISH in sprms:
        props.hidden = apply_toggle(props.hidden, sprms[SPRM_CFVANISH][0])

    # Character style index
    if SPRM_CISTD in sprms:
        props.style_index = struct.unpack_from("<H", sprms[SPRM_CISTD])[0]

    # Track which fields were directly set
    props._set_fields = get_char_sprm_fields(sprms)


def get_char_sprm_fields(sprms: dict[int, bytes]) -> set[str]:
    """Return which character fields were set by the given SPRMs."""
    fields: set[str] = set()
    mapping = {
        SPRM_CFBOLD: "bold",
        SPRM_CFITALIC: "italic",
        SPRM_CFSTRIKE: "strikethrough",
        SPRM_CKUL: "underline",
        SPRM_CHPS: "font_size",
        SPRM_CRGFTC0: "font_index",
        SPRM_CICO: "color",
        SPRM_CCV: "color",
        SPRM_CHIGHLIGHT: "highlight_color",
        SPRM_CISS: "superscript",
        SPRM_CFCAPS: "all_caps",
        SPRM_CFSMALLCAPS: "small_caps",
        SPRM_CFVANISH: "hidden",
        SPRM_CISTD: "style_index",
    }
    for sprm_code in sprms:
        if sprm_code in mapping:
            fields.add(mapping[sprm_code])
    return fields


# =============================================================================
# SEPX Section Property Parsing
# =============================================================================


def parse_sepx_sprms(sepx: bytes) -> dict:
    """Parse SPRM entries in a SEPX record into section properties."""
    props: dict = {}
    pos = 0
    while pos + 2 <= len(sepx):
        sprm = struct.unpack_from("<H", sepx, pos)[0]
        spra = (sprm >> 13) & 0x07

        if spra in (0, 1):
            operand_size = 1
        elif spra == 2:
            operand_size = 2
        elif spra == 3:
            operand_size = 4
        elif spra in (4, 5):
            operand_size = 2
        elif spra == 7:
            operand_size = 3
        elif spra == 6:
            if pos + 2 < len(sepx):
                operand_size = sepx[pos + 2] + 1
            else:
                break
        else:
            operand_size = 0

        if pos + 2 + operand_size > len(sepx):
            break

        operand = sepx[pos + 2 : pos + 2 + operand_size]

        if sprm == SPRM_SXAPAGE and operand_size == 2:
            props["page_width"] = struct.unpack_from("<H", operand, 0)[0] / 20.0
        elif sprm == SPRM_SYAPAGE and operand_size == 2:
            props["page_height"] = struct.unpack_from("<H", operand, 0)[0] / 20.0
        elif sprm == SPRM_SDXALEFT and operand_size == 2:
            props["left_margin"] = struct.unpack_from("<H", operand, 0)[0] / 20.0
        elif sprm == SPRM_SDXARIGHT and operand_size == 2:
            props["right_margin"] = struct.unpack_from("<H", operand, 0)[0] / 20.0
        elif sprm == SPRM_SDYATOP and operand_size == 2:
            props["top_margin"] = struct.unpack_from("<h", operand, 0)[0] / 20.0
        elif sprm == SPRM_SDYABOTTOM and operand_size == 2:
            props["bottom_margin"] = struct.unpack_from("<h", operand, 0)[0] / 20.0
        elif sprm == SPRM_SDYAHDRTOP and operand_size == 2:
            props["header_distance"] = struct.unpack_from("<H", operand, 0)[0] / 20.0
        elif sprm == SPRM_SDYAHDRBOTTOM and operand_size == 2:
            props["footer_distance"] = struct.unpack_from("<H", operand, 0)[0] / 20.0
        elif sprm == SPRM_SFTITLEPAGE and operand_size == 1:
            props["different_first_page"] = bool(operand[0])
        elif sprm == SPRM_SFPGNRESTART and operand_size == 1:
            props["restart_page_numbering"] = bool(operand[0])
        elif sprm == SPRM_SPGNSTART and operand_size == 2:
            props["page_starting_number"] = struct.unpack_from("<H", operand, 0)[0]

        pos += 2 + operand_size

    return props
