"""
Font table (SttbfFfn) parser for DOC files.
"""

from __future__ import annotations

import struct


def parse_font_table(table: bytes, fc: int, lcb: int) -> dict[int, str]:
    """Parse the SttbfFfn font table and return index -> font name mapping."""
    if lcb == 0:
        return {}

    ffn_data = table[fc : fc + lcb]
    if len(ffn_data) < 4:
        return {}

    # SttbfFfn: count (uint16) + cbExtra (uint16) + FFN entries
    count = struct.unpack_from("<H", ffn_data, 0)[0]
    # cbExtra = struct.unpack_from("<H", ffn_data, 2)[0]  # unused

    fonts: dict[int, str] = {}
    pos = 4
    for i in range(count):
        if pos >= lcb:
            break
        cb_ffn_m1 = ffn_data[pos]
        entry_len = cb_ffn_m1 + 1
        if pos + entry_len > lcb:
            break
        entry = ffn_data[pos : pos + entry_len]
        # Font name starts at byte 40 within FFN entry (after fixed fields)
        # FFN: cbFfnM1(1) + flags(1) + wWeight(2) + chs(1) + ixchSzAlt(1)
        #      + PANOSE(10) + FONTSIGNATURE(24) + xszFfn(rest)
        name_offset = 1 + 1 + 2 + 1 + 1 + 10 + 24  # = 40
        if name_offset < len(entry):
            name_bytes = entry[name_offset:]
            try:
                name = name_bytes.decode("utf-16-le").rstrip("\x00")
            except UnicodeDecodeError:
                name = name_bytes.split(b"\x00")[0].decode("cp1252", errors="replace")
            fonts[i] = name
        pos += entry_len

    return fonts
