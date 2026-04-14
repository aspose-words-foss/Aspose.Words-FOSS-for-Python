"""
List definition (PlfLst) and list format override (PlfLfo) parsers for DOC files.
"""

from __future__ import annotations

import struct


class ListDef:
    """Parsed list definition."""

    def __init__(self):
        self.lsid: int = 0
        self.is_hybrid: bool = False  # hybrid = bullet list in Aspose output


def parse_list_defs(table: bytes, fc_lst: int, lcb_lst: int) -> dict[int, ListDef]:
    """Parse PlfLst to get list definitions keyed by lsid."""
    if lcb_lst == 0:
        return {}

    lst_data = table[fc_lst : fc_lst + lcb_lst]
    if len(lst_data) < 2:
        return {}

    c_lst = struct.unpack_from("<H", lst_data, 0)[0]
    result: dict[int, ListDef] = {}

    for i in range(c_lst):
        offset = 2 + i * 28
        if offset + 28 > len(lst_data):
            break

        ld = ListDef()
        ld.lsid = struct.unpack_from("<I", lst_data, offset)[0]
        flags = lst_data[offset + 26]
        ld.is_hybrid = bool(flags & 0x10)
        result[ld.lsid] = ld

    return result


def parse_lfo_map(table: bytes, fc_lfo: int, lcb_lfo: int) -> dict[int, int]:
    """Parse PlfLfo to map LFO index (1-based) -> lsid."""
    if lcb_lfo == 0:
        return {}

    lfo_data = table[fc_lfo : fc_lfo + lcb_lfo]
    if len(lfo_data) < 4:
        return {}

    lfo_mac = struct.unpack_from("<I", lfo_data, 0)[0]
    result: dict[int, int] = {}

    for i in range(lfo_mac):
        offset = 4 + i * 16
        if offset + 16 > len(lfo_data):
            break
        lsid = struct.unpack_from("<I", lfo_data, offset)[0]
        result[i + 1] = lsid  # LFO indices are 1-based in PAPX sprms

    return result
