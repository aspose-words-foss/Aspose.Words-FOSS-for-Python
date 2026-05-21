"""
List definition (PlfLst) and list format override (PlfLfo) parsers for DOC files.

The PlfLst structure has two parts:

  1. A 2-byte count followed by an array of LSTF records (28 bytes each).
     ``LSTF.fSimpleList`` (bit 0 at offset 26) tells us whether the list
     has 1 active level or 9; ``LSTF.fHybrid`` (bit 4) is Word's
     "mixed-style" multilevel flag.

  2. Right after the LSTF array, the actual LVL records are packed.
     One LVL per active level per LSTF — so a hybrid / multi-level
     list contributes 9 LVL blocks, a simple list contributes 1.

Each LVL is variable-length:

  * 28-byte LVLF header.
  * ``cbGrpprlPapx`` bytes of paragraph SPRMs (left / first-line indent,
    tab stops).
  * ``cbGrpprlChpx`` bytes of character SPRMs (glyph font + size + bold).
  * A 2-byte ``cch`` followed by ``cch * 2`` bytes of UTF-16 LE bullet /
    number-format text (e.g. ``\\uf0b7`` for the Symbol bullet).

The reader walks both parts and returns a ``ListDef`` per LSID carrying
fully parsed ``ListLevel`` entries — npos / tpos resolved from the
PAPX, glyph font + size from the CHPX, and the literal bullet / format
text from the xst.
"""

import struct
from dataclasses import dataclass, field
from typing import Optional

# Number format codes (LVLF.nfc) that we map onto the LDM NumberStyle
# enum.  These mirror the MS-DOC values: 0 = decimal arabic, 1 = upper
# roman, 2 = lower roman, 3 = upper letter, 4 = lower letter,
# 23 = bullet, 255 = NoNumber (suppress the marker entirely).
_NFC_DECIMAL = 0
_NFC_UPPER_ROMAN = 1
_NFC_LOWER_ROMAN = 2
_NFC_UPPER_LETTER = 3
_NFC_LOWER_LETTER = 4
_NFC_BULLET = 23
_NFC_NO_NUMBER = 255

# CHPX SPRMs we care about for the bullet/number glyph rPr.
_SPRM_CFBOLD = 0x0835
_SPRM_CFITALIC = 0x0836
_SPRM_CHPS = 0x4A43  # font size, half-points
_SPRM_CRGFTC0 = 0x4A4F  # font index into FFN table
_SPRM_CRGFTC1 = 0x4A50
_SPRM_CRGFTC2 = 0x4A51

# PAPX SPRMs we care about for indent + first-line indent.
_SPRM_PDXALEFT_80 = 0x840F
_SPRM_PDXALEFT = 0x845E
_SPRM_PDXALEFT1_80 = 0x8411
_SPRM_PDXALEFT1 = 0x8460

_TWIPS_PER_PT = 20.0
_HALF_POINTS_PER_PT = 2.0


@dataclass
class ListLevelData:
    """One parsed LVL record (measurements in points)."""

    nfc: int = _NFC_BULLET
    start_at: int = 1
    alignment: int = 0  # 0/1/2 = L/C/R
    follow: int = 0  # 0=tab, 1=space, 2=nothing
    lvl_text: str = ""
    left_indent_pt: float = 0.0
    first_line_indent_pt: float = 0.0
    font_index: int = -1
    font_size_pt: float = 0.0
    bold: bool = False
    italic: bool = False


@dataclass
class ListDef:
    """Parsed list definition with full level information."""

    lsid: int = 0
    is_hybrid: bool = False
    is_simple: bool = False  # True → only one level active, else 9.
    levels: list[ListLevelData] = field(default_factory=list)


def parse_list_defs(table: bytes, fc_lst: int, lcb_lst: int) -> dict[int, ListDef]:
    """Parse PlfLst into ``{lsid: ListDef}``.

    Layout is ``[cLst:u16] [LSTF * cLst] [LVL ...]``; ``lcb_lst`` only
    covers the LSTF prefix, so the LVL walker keeps consuming the full
    table stream until every active level has been materialised.
    """
    if lcb_lst == 0:
        return {}

    if fc_lst + 2 > len(table):
        return {}

    c_lst = struct.unpack_from("<H", table, fc_lst)[0]
    if c_lst == 0:
        return {}

    defs: list[ListDef] = []
    for i in range(c_lst):
        offset = fc_lst + 2 + i * 28
        if offset + 28 > len(table):
            break
        ld = ListDef()
        ld.lsid = struct.unpack_from("<I", table, offset)[0]
        flags = table[offset + 26]
        ld.is_simple = bool(flags & 0x01)
        ld.is_hybrid = bool(flags & 0x10)
        defs.append(ld)

    pos = fc_lst + 2 + c_lst * 28
    for ld in defs:
        n_levels = 1 if ld.is_simple else 9
        for _ in range(n_levels):
            level, advance = _parse_lvl(table, pos)
            if level is None:
                break
            ld.levels.append(level)
            pos += advance
    return {ld.lsid: ld for ld in defs}


def _parse_lvl(buf: bytes, pos: int) -> tuple[Optional[ListLevelData], int]:
    """Parse one LVL record starting at *pos*, return (level, bytes_consumed)."""
    if pos + 28 > len(buf):
        return None, 0
    lvlf = buf[pos : pos + 28]

    level = ListLevelData()
    level.start_at = struct.unpack_from("<i", lvlf, 0)[0]
    level.nfc = lvlf[4]
    flags = lvlf[5]
    level.alignment = flags & 0x03  # lower 2 bits = jc
    level.follow = lvlf[15]
    cb_grpprl_chpx = lvlf[24]
    cb_grpprl_papx = lvlf[25]

    cursor = pos + 28

    papx_end = cursor + cb_grpprl_papx
    if papx_end > len(buf):
        return None, 0
    _apply_lvl_papx(buf[cursor:papx_end], level)
    cursor = papx_end

    chpx_end = cursor + cb_grpprl_chpx
    if chpx_end > len(buf):
        return None, 0
    _apply_lvl_chpx(buf[cursor:chpx_end], level)
    cursor = chpx_end

    if cursor + 2 > len(buf):
        return None, 0
    cch = struct.unpack_from("<H", buf, cursor)[0]
    cursor += 2
    xst_end = cursor + cch * 2
    if xst_end > len(buf):
        return None, 0
    raw_text = buf[cursor:xst_end].decode("utf-16-le", errors="replace")
    # U+0000..U+0008 in xst are level-N placeholders → OOXML %N.
    level.lvl_text = "".join(
        f"%{ord(ch) + 1}" if ord(ch) < 9 else ch for ch in raw_text
    )
    cursor = xst_end

    return level, cursor - pos


def _apply_lvl_papx(grpprl: bytes, level: ListLevelData) -> None:
    """Apply paragraph SPRMs from the level's grpprlPapx to *level*."""
    for sprm, op_size, operand in _walk_sprms(grpprl):
        if sprm in (_SPRM_PDXALEFT, _SPRM_PDXALEFT_80) and op_size == 2:
            level.left_indent_pt = struct.unpack_from("<h", operand)[0] / _TWIPS_PER_PT
        elif sprm in (_SPRM_PDXALEFT1, _SPRM_PDXALEFT1_80) and op_size == 2:
            level.first_line_indent_pt = (
                struct.unpack_from("<h", operand)[0] / _TWIPS_PER_PT
            )


def _apply_lvl_chpx(grpprl: bytes, level: ListLevelData) -> None:
    """Apply character SPRMs from the level's grpprlChpx to *level*."""
    for sprm, op_size, operand in _walk_sprms(grpprl):
        if sprm == _SPRM_CHPS and op_size == 2:
            level.font_size_pt = struct.unpack_from("<H", operand)[0] / _HALF_POINTS_PER_PT
        elif sprm in (_SPRM_CRGFTC0, _SPRM_CRGFTC1, _SPRM_CRGFTC2) and op_size == 2:
            # First ascii-script font wins; the FE/BiDi slots are only
            # consulted when the latin slot is unset.
            if level.font_index < 0:
                level.font_index = struct.unpack_from("<H", operand)[0]
        elif sprm == _SPRM_CFBOLD and op_size == 1 and operand[0] in (0x01, 0x81):
            level.bold = True
        elif sprm == _SPRM_CFITALIC and op_size == 1 and operand[0] in (0x01, 0x81):
            level.italic = True


def _walk_sprms(grpprl: bytes):
    """Yield ``(sprm_code, operand_size, operand_bytes)`` tuples."""
    pos = 0
    sizes = {0: 1, 1: 1, 2: 2, 3: 4, 4: 2, 5: 2, 7: 3}
    while pos + 2 <= len(grpprl):
        sprm = struct.unpack_from("<H", grpprl, pos)[0]
        spra = (sprm >> 13) & 0x07
        if spra == 6:
            if pos + 2 >= len(grpprl):
                return
            op_size = grpprl[pos + 2]
            if pos + 3 + op_size > len(grpprl):
                return
            yield sprm, op_size, grpprl[pos + 3 : pos + 3 + op_size]
            pos += 3 + op_size
            continue
        op_size = sizes.get(spra, 0)
        if op_size == 0 or pos + 2 + op_size > len(grpprl):
            return
        yield sprm, op_size, grpprl[pos + 2 : pos + 2 + op_size]
        pos += 2 + op_size


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
