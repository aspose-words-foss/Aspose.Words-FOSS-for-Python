"""
Style sheet (STSH) parser for DOC files.
"""

from __future__ import annotations

import struct

from aspose.words_foss.doc_reader.properties import (
    CharProps,
    ParaProps,
    apply_char_sprms,
    apply_para_sprms,
    get_char_sprm_fields,
    get_para_sprm_fields,
    parse_sprms,
)


class StyleData:
    """Properties parsed from a style definition (STSH UPX)."""

    def __init__(self):
        self.name: str = ""
        self.sti: int = 0x0FFF  # standard style identifier (built-in type)
        self.stk: int = 0  # 1=para, 2=char, 3=table, 4=list
        self.istd_base: int = 0x0FFF  # base style index (0xFFF = none)
        # Paragraph format properties (from UPX1)
        self.para_props: ParaProps = ParaProps()
        self.para_props_set: set[str] = set()  # which fields were explicitly set
        # Character format properties (from UPX2 or UPX1 for char styles)
        self.char_props: CharProps = CharProps()
        self.char_props_set: set[str] = set()  # which fields were explicitly set


def parse_stsh(table: bytes, fc: int, lcb: int) -> dict[int, str]:
    """Parse the STSH (style sheet) and return a map of istd -> style name."""
    if lcb == 0:
        return {}

    stsh = table[fc : fc + lcb]
    cb_stshi = struct.unpack_from("<H", stsh, 0)[0]
    stshi = stsh[2 : 2 + cb_stshi]
    cstd = struct.unpack_from("<H", stshi, 0)[0]
    cb_std_base = struct.unpack_from("<H", stshi, 2)[0]

    styles: dict[int, str] = {}
    offset = 2 + cb_stshi

    for i in range(cstd):
        if offset + 2 > len(stsh):
            break
        cb_std = struct.unpack_from("<H", stsh, offset)[0]
        offset += 2
        if cb_std == 0:
            continue

        std_data = stsh[offset : offset + cb_std]
        offset += cb_std

        if len(std_data) >= cb_std_base + 2:
            name_len = struct.unpack_from("<H", std_data, cb_std_base)[0]
            name_bytes = std_data[cb_std_base + 2 : cb_std_base + 2 + name_len * 2]
            name = name_bytes.decode("utf-16-le", errors="replace")
            styles[i] = name

    return styles


def parse_stsh_full(
    table: bytes, fc: int, lcb: int
) -> tuple[dict[int, str], dict[int, StyleData], int]:
    """Parse the STSH and return (name_map, style_data_map, default_font_index)."""
    if lcb == 0:
        return {}, {}, 0

    stsh = table[fc : fc + lcb]
    cb_stshi = struct.unpack_from("<H", stsh, 0)[0]
    stshi = stsh[2 : 2 + cb_stshi]
    cstd = struct.unpack_from("<H", stshi, 0)[0]
    cb_std_base = struct.unpack_from("<H", stshi, 2)[0]

    # Default font index from STSHI header (offset 12 = rgftcStandardChpStsh[0])
    default_ftc = 0
    if cb_stshi >= 14:
        default_ftc = struct.unpack_from("<H", stshi, 12)[0]

    names: dict[int, str] = {}
    style_data: dict[int, StyleData] = {}

    offset = 2 + cb_stshi

    for i in range(cstd):
        if offset + 2 > len(stsh):
            break
        cb_std = struct.unpack_from("<H", stsh, offset)[0]
        offset += 2
        if cb_std == 0:
            continue

        std_data = stsh[offset : offset + cb_std]
        offset += cb_std

        if len(std_data) < cb_std_base:
            continue

        sd = StyleData()

        # StdfBase: word at offset 0 has sti (bits 0-11)
        w0 = struct.unpack_from("<H", std_data, 0)[0]
        sd.sti = w0 & 0x0FFF

        # StdfBase: word at offset 2 has stk(4 bits) + istdBase(12 bits)
        w1 = struct.unpack_from("<H", std_data, 2)[0]
        sd.stk = w1 & 0x0F
        sd.istd_base = (w1 >> 4) & 0x0FFF

        w2 = struct.unpack_from("<H", std_data, 4)[0]
        cupx = w2 & 0x0F

        # Parse name
        if len(std_data) >= cb_std_base + 2:
            name_len = struct.unpack_from("<H", std_data, cb_std_base)[0]
            name_bytes = std_data[cb_std_base + 2 : cb_std_base + 2 + name_len * 2]
            raw_name = name_bytes.decode("utf-16-le", errors="replace")
            # Use primary name only (before comma for alternate names)
            sd.name = raw_name.split(",")[0].strip()
            # Normalize built-in style names via sti so downstream
            # code can detect them regardless of document locale.
            if 1 <= sd.sti <= 9:
                sd.name = f"Heading {sd.sti}"
            elif 19 <= sd.sti <= 27:
                sd.name = f"TOC {sd.sti - 18}"
            names[i] = sd.name
        else:
            name_len = 0

        # Parse UPX blocks
        upx_start = cb_std_base + 2 + name_len * 2 + 2
        if upx_start % 2 != 0:
            upx_start += 1

        if upx_start < len(std_data):
            upx_remaining = std_data[upx_start:]

            if sd.stk == 1 and cupx >= 2:
                # Paragraph style: UPX1 (para) + UPX2 (char)
                if len(upx_remaining) >= 2:
                    cb_upx1 = struct.unpack_from("<H", upx_remaining, 0)[0]
                    upx1 = upx_remaining[2 : 2 + cb_upx1]
                    if len(upx1) >= 2:
                        grpprl = upx1[2:]  # skip istd
                        sprms = parse_sprms(grpprl)
                        apply_para_sprms(sd.para_props, sprms)
                        sd.para_props_set = get_para_sprm_fields(sprms)

                    upx2_off = 2 + cb_upx1
                    if upx2_off % 2:
                        upx2_off += 1
                    if upx2_off + 2 <= len(upx_remaining):
                        cb_upx2 = struct.unpack_from("<H", upx_remaining, upx2_off)[0]
                        upx2 = upx_remaining[upx2_off + 2 : upx2_off + 2 + cb_upx2]
                        sprms2 = parse_sprms(upx2)
                        apply_char_sprms(sd.char_props, sprms2)
                        sd.char_props_set = get_char_sprm_fields(sprms2)

            elif sd.stk == 2 and cupx >= 1:
                # Character style: UPX1 (char)
                if len(upx_remaining) >= 2:
                    cb_upx = struct.unpack_from("<H", upx_remaining, 0)[0]
                    upx = upx_remaining[2 : 2 + cb_upx]
                    sprms = parse_sprms(upx)
                    apply_char_sprms(sd.char_props, sprms)
                    sd.char_props_set = get_char_sprm_fields(sprms)

        style_data[i] = sd

    return names, style_data, default_ftc
