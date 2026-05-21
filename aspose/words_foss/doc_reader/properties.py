"""
Paragraph and character property parsing (PAPX/CHPX) for DOC files.

Includes SPRM parsing, FKP page parsing, and SEPX section property parsing.
"""

import struct
from typing import Optional

from aspose.words_foss.doc_reader.constants import (
    ICO_COLORS,
    SPRM_CFBOLD,
    SPRM_CFCAPS,
    SPRM_CFEMBOSS,
    SPRM_CFIMPRINT,
    SPRM_CFITALIC,
    SPRM_CFOUTLINE,
    SPRM_CFSHADOW,
    SPRM_CFSMALLCAPS,
    SPRM_CFSPEC,
    SPRM_CFSTRIKE,
    SPRM_CFVANISH,
    SPRM_CHIGHLIGHT,
    SPRM_CHPS,
    SPRM_CHPSKERN,
    SPRM_CICO,
    SPRM_CISS,
    SPRM_CISTD,
    SPRM_CCV,
    SPRM_CKUL,
    SPRM_CPICLOCATION,
    SPRM_CRGFTC0,
    SPRM_PBRCBOTTOM80,
    SPRM_PBRCLEFT80,
    SPRM_PBRCRIGHT80,
    SPRM_PBRCTOP80,
    SPRM_PCHGTABSPAPX,
    SPRM_PCHGTABS,
    SPRM_PCNF,
    SPRM_PDXALEFT,
    SPRM_PDXALEFT1,
    SPRM_PDXALEFT1_80,
    SPRM_PDXALEFT80,
    SPRM_PDXARIGHT,
    SPRM_PDXARIGHT80,
    SPRM_PDYAAFTER,
    SPRM_PDYABEFORE,
    SPRM_PDYALINE,
    SPRM_PFADJUSTRIGHT,
    SPRM_PFAUTOSPACEDE,
    SPRM_PFAUTOSPACEDN,
    SPRM_PFCONTEXTUALSPACING,
    SPRM_PFDYAAFTERAUTO,
    SPRM_PFDYABEFOREAUTO,
    SPRM_PFINTABLE,
    SPRM_PFKEEP,
    SPRM_PFKEEPFOLLOW,
    SPRM_PFNOAUTOHYPH,
    SPRM_PFNOLINENUMB,
    SPRM_PFPAGEBREAKBEFORE,
    SPRM_PFTTP,
    SPRM_PFWIDOWCONTROL,
    SPRM_PHUGEPAPX,
    SPRM_PILFO,
    SPRM_PILVL,
    SPRM_PJC,
    SPRM_PJC80,
    SPRM_POUTLVL,
    SPRM_PSHD80,
    SPRM_PWALIGNFONT,
    SPRM_SBKC,
    SPRM_SDXALEFT,
    SPRM_SDXARIGHT,
    SPRM_SDYABOTTOM,
    SPRM_SCCOLUMNS,
    SPRM_SDXACOLUMNS,
    SPRM_SDYAHDRBOTTOM,
    SPRM_SDYAHDRTOP,
    SPRM_SDYATOP,
    SPRM_SDZAGUTTER,
    SPRM_SFEVENLYSPCOLS,
    SPRM_SFLBETWEEN,
    SPRM_SFTITLEPAGE,
    SPRM_SFPGNRESTART,
    SPRM_SPGNSTART,
    SPRM_SXAPAGE,
    SPRM_SYAPAGE,
    SPRM_TDEFTABLE,
    SPRM_TTPC,
    SPRM_TTBLPY,
    SPRM_TDXAFROMTEXT,
    SPRM_TDXAFROMTEXTRIGHT,
    SPRM_TTABLEWIDTH,
    SPRM_TTABLEBORDERS,
    SPRM_TTABLEBORDERS80,
    SPRM_TPROPREV,
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
        self.keep_together: bool = False
        # CT_OnOff defaults: widowControl/snapToGrid/autoSpace*/adjustRight are ON.
        self.widow_control: bool = True
        self.suppress_auto_hyphens: bool = False
        self.suppress_line_numbers: bool = False
        self.snap_to_grid: bool = True
        self.add_space_between_far_east_and_alpha: bool = True
        self.add_space_between_far_east_and_digit: bool = True
        self.auto_adjust_right_indent: bool = True
        self.no_space_between_paragraphs_of_same_style: bool = False
        self.page_break_before: bool = False
        self.outline_level: int = 9  # 9 = body text
        self.tab_stops: list[tuple[float, int, int]] = []  # (position_pt, alignment, leader)
        # (brcType, width_pt, color_str, dpt_space) per side (top/left/bottom/right).
        self.borders: list[Optional[tuple[int, float, str, int]]] = [None, None, None, None]
        self.shading_back: str = ""
        # 0=Auto, 1=Top, 2=Center, 3=Baseline, 4=Bottom.
        self.baseline_alignment: int = 0
        # 12-bit cnfStyle mask, "1"/"0" characters as in OOXML.
        self.conditional_style: str = ""
        # sprmPFInTable / sprmPFTtp — cell content vs row-end marker.
        self.in_table: bool = False
        self.is_table_terminator: bool = False
        self._set_fields: set[str] = set()
        self._raw_grpprl: Optional[bytes] = None


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
        # Font effects (toggle SPRMs).
        self.emboss: bool = False
        self.engrave: bool = False  # CFImprint
        self.outline: bool = False
        self.shadow: bool = False
        # Minimum font size at which kerning kicks in; 0 = no kerning.
        self.kerning: float = 0.0
        self.style_index: int = -1  # character style index
        self.is_special: bool = False  # fSpec — inline picture, symbol, etc.
        self.pic_location: int = -1  # offset into Data stream for inline picture
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
            if sprm == SPRM_TDEFTABLE:
                # sprmTDefTable uses a 2-byte cb (MS-DOC spec exception)
                if off + 4 <= len(grpprl):
                    op_size = struct.unpack_from("<H", grpprl, off + 2)[0]
                    if off + 4 + op_size <= len(grpprl):
                        result[sprm] = grpprl[off + 4 : off + 4 + op_size]
                    off += 4 + op_size
                    continue
                else:
                    break
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


def _expand_huge_papx(grpprl: bytes, data_stream: bytes) -> bytes:
    """Dereference ``sprmPHugePapx`` — the real grpprl lives in the Data stream."""
    sprms = parse_sprms(grpprl)
    op = sprms.get(SPRM_PHUGEPAPX)
    if op is None or len(op) < 4 or not data_stream:
        return grpprl
    off = struct.unpack_from("<I", op)[0]
    if off + 2 > len(data_stream):
        return grpprl
    cb = struct.unpack_from("<H", data_stream, off)[0]
    if off + 2 + cb > len(data_stream):
        return grpprl
    return data_stream[off + 2 : off + 2 + cb]


def _fc_to_cp(fc: int, pieces: list[tuple[int, int, int, bool]], fallback_offset: int, fallback_compressed: bool) -> int:
    """Map a byte FC to a character CP through the piece table."""
    for cp_s, cp_e, fc_bs, f_compressed in pieces:
        bytes_per_char = 1 if f_compressed else 2
        byte_len = (cp_e - cp_s) * bytes_per_char
        if fc_bs <= fc <= fc_bs + byte_len:
            return cp_s + (fc - fc_bs) // bytes_per_char
    if fallback_compressed:
        return fc - fallback_offset
    return (fc - fallback_offset) // 2


def parse_papx_fkp(
    wd: bytes,
    pn: int,
    text_byte_offset: int,
    text_is_compressed: bool,
    data_stream: bytes = b"",
    pieces: list[tuple[int, int, int, bool]] = None,
) -> list[tuple[int, int, ParaProps]]:
    """Parse a PAPX FKP page and return (char_start, char_end, props) tuples."""
    page_offset = pn * 512
    if page_offset + 512 > len(wd):
        return []

    page = wd[page_offset : page_offset + 512]
    cpara = page[511]
    if cpara == 0:
        return []

    pieces = pieces or []
    results: list[tuple[int, int, ParaProps]] = []

    for j in range(cpara):
        fc_j = struct.unpack_from("<I", page, j * 4)[0]
        fc_next = struct.unpack_from("<I", page, (j + 1) * 4)[0]

        if pieces:
            char_start = _fc_to_cp(fc_j, pieces, text_byte_offset, text_is_compressed)
            char_end = _fc_to_cp(fc_next, pieces, text_byte_offset, text_is_compressed)
        elif text_is_compressed:
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
                    grpprl = _expand_huge_papx(
                        page[papx_pos + 4 : papx_pos + 2 + total], data_stream
                    )
                    sprms = parse_sprms(grpprl)
                    apply_para_sprms(props, sprms)
                    props._raw_grpprl = grpprl
        else:
            total = cb * 2
            if papx_pos + 1 + total <= 512 and total >= 2:
                props.istd = struct.unpack_from("<H", page, papx_pos + 1)[0]
                grpprl = _expand_huge_papx(
                    page[papx_pos + 3 : papx_pos + 1 + total], data_stream
                )
                sprms = parse_sprms(grpprl)
                apply_para_sprms(props, sprms)
                props._raw_grpprl = grpprl

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

    # Keep paragraph together (keepLines)
    if SPRM_PFKEEP in sprms:
        props.keep_together = bool(sprms[SPRM_PFKEEP][0])

    # Widow / orphan control
    if SPRM_PFWIDOWCONTROL in sprms:
        props.widow_control = bool(sprms[SPRM_PFWIDOWCONTROL][0])

    # Suppress line numbers
    if SPRM_PFNOLINENUMB in sprms:
        props.suppress_line_numbers = bool(sprms[SPRM_PFNOLINENUMB][0])

    # Suppress automatic hyphenation
    if SPRM_PFNOAUTOHYPH in sprms:
        props.suppress_auto_hyphens = bool(sprms[SPRM_PFNOAUTOHYPH][0])

    # Far-east / latin auto-spacing
    if SPRM_PFAUTOSPACEDE in sprms:
        props.add_space_between_far_east_and_alpha = bool(sprms[SPRM_PFAUTOSPACEDE][0])
    if SPRM_PFAUTOSPACEDN in sprms:
        props.add_space_between_far_east_and_digit = bool(sprms[SPRM_PFAUTOSPACEDN][0])

    # Auto-adjust right indent
    if SPRM_PFADJUSTRIGHT in sprms:
        props.auto_adjust_right_indent = bool(sprms[SPRM_PFADJUSTRIGHT][0])

    # Contextual spacing (no space between paragraphs of the same style)
    if SPRM_PFCONTEXTUALSPACING in sprms:
        props.no_space_between_paragraphs_of_same_style = bool(
            sprms[SPRM_PFCONTEXTUALSPACING][0]
        )

    # Page break before
    if SPRM_PFPAGEBREAKBEFORE in sprms:
        props.page_break_before = bool(sprms[SPRM_PFPAGEBREAKBEFORE][0])

    if SPRM_PFINTABLE in sprms:
        props.in_table = bool(sprms[SPRM_PFINTABLE][0])
    if SPRM_PFTTP in sprms:
        props.is_table_terminator = bool(sprms[SPRM_PFTTP][0])

    # Outline level
    if SPRM_POUTLVL in sprms:
        props.outline_level = sprms[SPRM_POUTLVL][0]

    # Tab stops (sprmPChgTabsPapx 0xC615 or sprmPChgTabs 0xC60D)
    tab_data = sprms.get(SPRM_PCHGTABSPAPX) or sprms.get(SPRM_PCHGTABS)
    if tab_data is not None:
        props.tab_stops = _parse_tab_sprm(tab_data)

    # Paragraph borders (BRC80 — 4 bytes per side)
    for slot, sprm_code in (
        (0, SPRM_PBRCTOP80),
        (1, SPRM_PBRCLEFT80),
        (2, SPRM_PBRCBOTTOM80),
        (3, SPRM_PBRCRIGHT80),
    ):
        if sprm_code in sprms and len(sprms[sprm_code]) >= 4:
            props.borders[slot] = _parse_brc80(sprms[sprm_code])

    # Paragraph shading (Shd80 — 2 bytes: ipat | icoFore | icoBack)
    if SPRM_PSHD80 in sprms and len(sprms[SPRM_PSHD80]) >= 2:
        shd_word = struct.unpack_from("<H", sprms[SPRM_PSHD80])[0]
        ico_back = (shd_word >> 11) & 0x1F
        if ico_back:
            props.shading_back = ICO_COLORS.get(ico_back, "")

    if SPRM_PWALIGNFONT in sprms and len(sprms[SPRM_PWALIGNFONT]) >= 2:
        props.baseline_alignment = struct.unpack_from("<H", sprms[SPRM_PWALIGNFONT])[0]

    if SPRM_PCNF in sprms and len(sprms[SPRM_PCNF]) >= 2:
        cnf_word = struct.unpack_from("<H", sprms[SPRM_PCNF])[0]
        props.conditional_style = "".join(
            "1" if (cnf_word >> i) & 1 else "0" for i in range(12)
        )

    props._set_fields = get_para_sprm_fields(sprms)


def _parse_brc80(data: bytes) -> Optional[tuple[int, float, str, int]]:
    """Parse a BRC80 (4-byte border) into (brcType, width_pt, color_str, dpt_space).

    Bytes: dptLineWidth(1) | brcType(1) | ico(1) | dptSpace:5 | fShadow:1 | fFrame:1 | _:1
    """
    dpt_line_width = data[0]
    brc_type = data[1]
    ico = data[2]
    dpt_space = data[3] & 0x1F
    if brc_type == 0 or dpt_line_width == 0xFF or brc_type == 0xFF:
        return None
    color_str = ICO_COLORS.get(ico, "")
    width_pt = dpt_line_width / 8.0
    return (brc_type, width_pt, color_str, dpt_space)


def _parse_tab_sprm(data: bytes) -> list[tuple[float, int, int]]:
    """Parse sprmPChgTabsPapx / sprmPChgTabs operand into tab stop tuples.

    The operand format is:
      - 1 byte: cTabs (number of delete positions to skip)
      - cTabs * 2 bytes: positions to delete (ignored here)
      - 1 byte: cAdds (number of tabs to add)
      - cAdds * 2 bytes: rgdxaTab positions (signed int16, twips)
      - cAdds * 1 byte: rgtbd descriptors (bits 0-2 = alignment, bits 3-5 = leader)
    """
    tabs: list[tuple[float, int, int]] = []
    if len(data) < 1:
        return tabs
    off = 0
    c_del = data[off]
    off += 1
    off += c_del * 2  # skip delete positions
    if off >= len(data):
        return tabs
    c_add = data[off]
    off += 1
    if off + c_add * 2 + c_add > len(data):
        return tabs
    positions: list[float] = []
    for i in range(c_add):
        pos_twips = struct.unpack_from("<h", data, off + i * 2)[0]
        positions.append(pos_twips / 20.0)
    off += c_add * 2
    for i in range(c_add):
        tbd = data[off + i]
        alignment = tbd & 0x07
        leader = (tbd >> 3) & 0x07
        tabs.append((positions[i], alignment, leader))
    return tabs


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
        SPRM_PFKEEP: "keep_together",
        SPRM_PFKEEPFOLLOW: "keep_with_next",
        SPRM_PFWIDOWCONTROL: "widow_control",
        SPRM_PFNOLINENUMB: "suppress_line_numbers",
        SPRM_PFNOAUTOHYPH: "suppress_auto_hyphens",
        SPRM_PFAUTOSPACEDE: "add_space_between_far_east_and_alpha",
        SPRM_PFAUTOSPACEDN: "add_space_between_far_east_and_digit",
        SPRM_PFADJUSTRIGHT: "auto_adjust_right_indent",
        SPRM_PFCONTEXTUALSPACING: "no_space_between_paragraphs_of_same_style",
        SPRM_PFPAGEBREAKBEFORE: "page_break_before",
        SPRM_POUTLVL: "outline_level",
        SPRM_PCHGTABSPAPX: "tab_stops",
        SPRM_PCHGTABS: "tab_stops",
        SPRM_PBRCTOP80: "borders",
        SPRM_PBRCLEFT80: "borders",
        SPRM_PBRCBOTTOM80: "borders",
        SPRM_PBRCRIGHT80: "borders",
        SPRM_PSHD80: "shading_back",
        SPRM_PWALIGNFONT: "baseline_alignment",
        SPRM_PCNF: "conditional_style",
    }
    for sprm_code in sprms:
        if sprm_code in mapping:
            fields.add(mapping[sprm_code])
    return fields


# =============================================================================
# CHPX FKP Parsing
# =============================================================================


def parse_chpx_fkp(
    wd: bytes,
    pn: int,
    text_byte_offset: int,
    text_is_compressed: bool,
    pieces: list[tuple[int, int, int, bool]] = None,
) -> list[tuple[int, int, CharProps]]:
    """Parse a CHPX FKP page and return (char_start, char_end, props) tuples."""
    page_offset = pn * 512
    if page_offset + 512 > len(wd):
        return []

    page = wd[page_offset : page_offset + 512]
    crun = page[511]
    if crun == 0:
        return []

    pieces = pieces or []
    results: list[tuple[int, int, CharProps]] = []

    for j in range(crun):
        fc_j = struct.unpack_from("<I", page, j * 4)[0]
        fc_next = struct.unpack_from("<I", page, (j + 1) * 4)[0]

        if pieces:
            char_start = _fc_to_cp(fc_j, pieces, text_byte_offset, text_is_compressed)
            char_end = _fc_to_cp(fc_next, pieces, text_byte_offset, text_is_compressed)
        elif text_is_compressed:
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

    # Font effects (toggle SPRMs)
    if SPRM_CFOUTLINE in sprms:
        operand = sprms[SPRM_CFOUTLINE][0]
        props.outline = apply_toggle(props.outline, operand)
        if operand == 0x81:
            props._toggle_fields.add("outline")
    if SPRM_CFSHADOW in sprms:
        operand = sprms[SPRM_CFSHADOW][0]
        props.shadow = apply_toggle(props.shadow, operand)
        if operand == 0x81:
            props._toggle_fields.add("shadow")
    if SPRM_CFEMBOSS in sprms:
        operand = sprms[SPRM_CFEMBOSS][0]
        props.emboss = apply_toggle(props.emboss, operand)
        if operand == 0x81:
            props._toggle_fields.add("emboss")
    if SPRM_CFIMPRINT in sprms:
        operand = sprms[SPRM_CFIMPRINT][0]
        props.engrave = apply_toggle(props.engrave, operand)
        if operand == 0x81:
            props._toggle_fields.add("engrave")

    # Kerning: minimum font size in points to apply kerning (half-points → points)
    if SPRM_CHPSKERN in sprms and len(sprms[SPRM_CHPSKERN]) >= 2:
        props.kerning = struct.unpack_from("<H", sprms[SPRM_CHPSKERN])[0] / 2.0

    # Character style index
    if SPRM_CISTD in sprms:
        props.style_index = struct.unpack_from("<H", sprms[SPRM_CISTD])[0]

    # Inline picture / special character
    if SPRM_CFSPEC in sprms:
        props.is_special = sprms[SPRM_CFSPEC][0] != 0
    if SPRM_CPICLOCATION in sprms:
        props.pic_location = struct.unpack_from("<i", sprms[SPRM_CPICLOCATION])[0]

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
        SPRM_CHPSKERN: "kerning",
        SPRM_CRGFTC0: "font_index",
        SPRM_CICO: "color",
        SPRM_CCV: "color",
        SPRM_CHIGHLIGHT: "highlight_color",
        SPRM_CISS: "superscript",
        SPRM_CFCAPS: "all_caps",
        SPRM_CFSMALLCAPS: "small_caps",
        SPRM_CFVANISH: "hidden",
        SPRM_CFOUTLINE: "outline",
        SPRM_CFSHADOW: "shadow",
        SPRM_CFEMBOSS: "emboss",
        SPRM_CFIMPRINT: "engrave",
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
        elif sprm == SPRM_SDZAGUTTER and operand_size == 2:
            # Gutter is unsigned: extra binding-edge margin in twips.
            props["gutter"] = struct.unpack_from("<H", operand, 0)[0] / 20.0
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
        elif sprm == SPRM_SCCOLUMNS and operand_size == 2:
            props["columns_count"] = struct.unpack_from("<H", operand, 0)[0] + 1
        elif sprm == SPRM_SDXACOLUMNS and operand_size == 2:
            props["columns_spacing"] = struct.unpack_from("<h", operand, 0)[0] / 20.0
        elif sprm == SPRM_SFEVENLYSPCOLS and operand_size == 1:
            props["columns_evenly_spaced"] = bool(operand[0])
        elif sprm == SPRM_SFLBETWEEN and operand_size == 1:
            props["columns_line_between"] = bool(operand[0])
        elif sprm == SPRM_SBKC and operand_size == 1:
            props["section_break_type"] = operand[0]

        pos += 2 + operand_size

    return props


# =============================================================================
# Table Property Parsing (from PAPX row-end paragraphs)
# =============================================================================


class TableRowProps:
    """Properties extracted from PAPX SPRMs on table row-end paragraphs."""

    def __init__(self):
        self.cell_widths: list[float] = []  # widths in points
        self.table_width: float = 0.0  # total table width in points
        self.tblp_y: int = 0  # vertical offset (tblpY, twips)
        self.dxa_from_text: int = 0  # leftFromText (twips)
        self.dxa_from_text_right: int = 0  # rightFromText (twips)
        self.tpc: int = 0  # table positioning control byte
        self.has_positioning: bool = False
        # 6 borders: (brcType, line_width_eighth_pt, color_str, dptSpace)
        # order: top, left, bottom, right, insideH, insideV
        self.borders: list[tuple[int, float, str, int]] = []
        self.horz_pos: int = -1  # 0=left, 1=center, 2=right from TPROPREV pcHorz


def _parse_brc_borders(data: bytes) -> list[tuple[int, float, str, int]]:
    """Parse 6 BRC records (8 bytes each) into (brcType, width_pt, color_str, space)."""
    borders: list[tuple[int, float, str, int]] = []
    for i in range(6):
        off = i * 8
        cv = struct.unpack_from("<I", data, off)[0]
        dpt_line_width = data[off + 4]
        brc_type = data[off + 5]
        dpt_space = data[off + 6]
        if (cv & 0xFF000000) == 0xFF000000:
            color_str = ""
        else:
            r, g, b = cv & 0xFF, (cv >> 8) & 0xFF, (cv >> 16) & 0xFF
            color_str = f"Color [A=255, R={r}, G={g}, B={b}]"
        width_pt = dpt_line_width / 8.0
        borders.append((brc_type, width_pt, color_str, dpt_space))
    return borders


def _parse_brc80_borders(data: bytes) -> list[tuple[int, float, str, int]]:
    """Parse 6 BRC80 records (4 bytes each) into (brcType, width_pt, color_str, space).

    Fallback for ``sprmTTableBorders80`` when the BRC8 SPRM (0xD613) is
    unreachable past ``sprmTDefTable``'s 2-byte cb.
    """
    borders: list[tuple[int, float, str, int]] = []
    for i in range(6):
        off = i * 4
        if off + 4 > len(data):
            break
        dpt_line_width = data[off]
        brc_type = data[off + 1]
        ico = data[off + 2]
        dpt_space = data[off + 3] & 0x1F
        color_str = ICO_COLORS.get(ico, "")
        width_pt = dpt_line_width / 8.0
        borders.append((brc_type, width_pt, color_str, dpt_space))
    return borders


def parse_table_row_sprms(grpprl: bytes) -> TableRowProps:
    """Parse table SPRMs from a row-end paragraph's grpprl.

    Walks the grpprl manually because ``sprmTDefTable`` (0xD608) uses
    a 2-byte cb instead of the standard 1-byte for spra=6.
    """
    props = TableRowProps()
    pos = 0
    sizes = {0: 1, 1: 1, 2: 2, 3: 4, 4: 2, 5: 2, 7: 3}
    has_explicit_tblp_y = False
    tblp_y_from_compound = 0

    while pos + 2 <= len(grpprl):
        sprm = struct.unpack_from("<H", grpprl, pos)[0]
        spra = (sprm >> 13) & 0x07

        if sprm == SPRM_TDEFTABLE:
            # sprmTDefTable uses 2-byte cb
            if pos + 4 > len(grpprl):
                break
            cb = struct.unpack_from("<H", grpprl, pos + 2)[0]
            operand_start = pos + 4
            if operand_start + cb > len(grpprl):
                pos += 4 + cb
                continue
            operand = grpprl[operand_start : operand_start + cb]
            if len(operand) >= 1:
                num_cells = operand[0]
                bnd_size = (num_cells + 1) * 2
                if len(operand) >= 1 + bnd_size:
                    boundaries = []
                    for i in range(num_cells + 1):
                        b = struct.unpack_from("<h", operand, 1 + i * 2)[0]
                        boundaries.append(b)
                    widths = []
                    for i in range(num_cells):
                        widths.append((boundaries[i + 1] - boundaries[i]) / 20.0)
                    props.cell_widths = widths
            pos = operand_start + cb
            continue

        if spra == 6:
            if pos + 2 < len(grpprl):
                op_size = grpprl[pos + 2]
                operand_6 = grpprl[pos + 3 : pos + 3 + op_size]
                if sprm == SPRM_TTABLEBORDERS and len(operand_6) >= 48:
                    props.borders = _parse_brc_borders(operand_6)
                elif sprm == SPRM_TTABLEBORDERS80 and len(operand_6) >= 24:
                    if not props.borders:
                        props.borders = _parse_brc80_borders(operand_6)
                pos += 3 + op_size
            else:
                break
            continue

        op_size = sizes.get(spra, 0)
        if op_size == 0:
            break
        if pos + 2 + op_size > len(grpprl):
            break
        operand = grpprl[pos + 2 : pos + 2 + op_size]

        if sprm == SPRM_TTPC and op_size == 1:
            props.tpc = operand[0]
            if operand[0] & 0x20:
                props.has_positioning = True
        elif sprm == SPRM_TTBLPY and op_size == 2:
            props.tblp_y = struct.unpack_from("<h", operand)[0]
            props.has_positioning = True
            has_explicit_tblp_y = True
        elif sprm == SPRM_TPROPREV and op_size == 4:
            tblp_y_from_compound = struct.unpack_from("<h", operand, 2)[0]
            props.has_positioning = True
            props.horz_pos = operand[0] & 0x03
        elif sprm == SPRM_TDXAFROMTEXT and op_size == 2:
            props.dxa_from_text = struct.unpack_from("<H", operand)[0]
            props.has_positioning = True
        elif sprm == SPRM_TDXAFROMTEXTRIGHT and op_size == 2:
            props.dxa_from_text_right = struct.unpack_from("<H", operand)[0]
        elif sprm == SPRM_TTABLEWIDTH and op_size == 3:
            fts = operand[0]
            w = struct.unpack_from("<H", operand, 1)[0]
            if fts == 3:  # dxa (twips → points)
                props.table_width = w / 20.0

        pos += 2 + op_size

    if not has_explicit_tblp_y and tblp_y_from_compound:
        props.tblp_y = tblp_y_from_compound

    return props
