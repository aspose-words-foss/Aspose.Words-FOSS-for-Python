"""
Document Reader for Word 97-2003 Binary Format (.doc) using olefile.

Parses .doc files and produces the same data structures as DocumentReader (reader.py),
enabling transparent support for both .doc and .docx formats.
"""

from __future__ import annotations

import re
import struct
from io import BytesIO
from pathlib import Path
from typing import Optional, Union, BinaryIO, Iterator, TYPE_CHECKING

import olefile

if TYPE_CHECKING:
    from aspose.words_foss import light_document_model as ldm

from aspose.words_foss.reader import (
    ParagraphData,
    TableData,
    RowData,
    CellData,
    RunData,
    NumberingInfo,
    NumberingLevel,
)

# =============================================================================
# FIB FIELD INDICES in FibRgFcLcb97
# =============================================================================

_IDX_STSHF = 1  # Style sheet
_IDX_PLCFBTECHPX = 12  # Character property bin table
_IDX_PLCFBTEPAPX = 13  # Paragraph property bin table
_IDX_CLX = 33  # Complex (piece table)
_IDX_PLFLST = 73  # List table
_IDX_PLFLFO = 74  # List format overrides

# SPRM codes
_SPRM_PILFO = 0x460B  # Paragraph list format override index
_SPRM_PILVL = 0x260A  # Paragraph indent level
_SPRM_CFBOLD = 0x0835  # Character bold
_SPRM_CFITALIC = 0x0836  # Character italic
_SPRM_CFSTRIKE = 0x0837  # Character strikethrough
_SPRM_CKUL = 0x2A3E  # Character underline type


# =============================================================================
# FIB PARSER
# =============================================================================


class _FibData:
    """Parsed FIB (File Information Block) data."""

    def __init__(self):
        self.f_which_tbl_stm: bool = False
        self.ccp_text: int = 0
        self.fc_start: int = 0  # Start of FibRgFcLcb in WordDocument stream
        self.cb_rg_fc_lcb: int = 0


def _parse_fib(wd: bytes) -> _FibData:
    """Parse the FIB from the WordDocument stream."""
    fib = _FibData()

    # Check magic number
    magic = struct.unpack_from("<H", wd, 0)[0]
    if magic != 0xA5EC:
        raise ValueError(f"Not a valid Word document (magic=0x{magic:04X})")

    # Flags at offset 10
    flags = struct.unpack_from("<H", wd, 10)[0]
    fib.f_which_tbl_stm = bool(flags & 0x0200)

    # FibRgW97
    csw = struct.unpack_from("<H", wd, 32)[0]
    rgw_end = 34 + csw * 2

    # FibRgLw97
    cslw = struct.unpack_from("<H", wd, rgw_end)[0]
    lw_start = rgw_end + 2
    fib.ccp_text = struct.unpack_from("<I", wd, lw_start + 12)[0]
    lw_end = lw_start + cslw * 4

    # FibRgFcLcb
    fib.cb_rg_fc_lcb = struct.unpack_from("<H", wd, lw_end)[0]
    fib.fc_start = lw_end + 2

    return fib


def _get_fc_lcb(wd: bytes, fib: _FibData, idx: int) -> tuple[int, int]:
    """Get an (fc, lcb) pair from FibRgFcLcb at the given index."""
    if idx >= fib.cb_rg_fc_lcb:
        return 0, 0
    offset = fib.fc_start + idx * 8
    fc = struct.unpack_from("<I", wd, offset)[0]
    lcb = struct.unpack_from("<I", wd, offset + 4)[0]
    return fc, lcb


# =============================================================================
# TEXT EXTRACTION
# =============================================================================


def _extract_text_via_piece_table(wd: bytes, table: bytes, fib: _FibData) -> str:
    """Extract document text using the CLX piece table."""
    fc_clx, lcb_clx = _get_fc_lcb(wd, fib, _IDX_CLX)

    if lcb_clx == 0:
        # No CLX: try reading text from fixed offset 0x800
        text_offset = 0x800
        if text_offset + fib.ccp_text * 2 <= len(wd):
            text_bytes = wd[text_offset : text_offset + fib.ccp_text * 2]
            return text_bytes.decode("utf-16-le", errors="replace")
        return ""

    clx = table[fc_clx : fc_clx + lcb_clx]
    pos = 0
    text_parts: list[str] = []

    while pos < len(clx):
        clxt = clx[pos]
        if clxt == 0x02:  # Piece table
            pos += 1
            pcdt_size = struct.unpack_from("<I", clx, pos)[0]
            pos += 4
            pt = clx[pos : pos + pcdt_size]

            n = (pcdt_size - 4) // 12
            cps = [struct.unpack_from("<I", pt, i * 4)[0] for i in range(n + 1)]

            for i in range(n):
                pcd_start = (n + 1) * 4 + i * 8
                fc_val = struct.unpack_from("<I", pt, pcd_start + 2)[0]
                f_compressed = bool(fc_val & 0x40000000)
                fc_real = fc_val & 0x3FFFFFFF
                char_count = cps[i + 1] - cps[i]

                if f_compressed:
                    text_bytes = wd[fc_real // 2 : fc_real // 2 + char_count]
                    text_parts.append(text_bytes.decode("cp1252", errors="replace"))
                else:
                    text_bytes = wd[fc_real : fc_real + char_count * 2]
                    text_parts.append(text_bytes.decode("utf-16-le", errors="replace"))
            break
        elif clxt == 0x01:  # Grpprl
            pos += 1
            cb = struct.unpack_from("<H", clx, pos)[0]
            pos += 2 + cb
        else:
            pos += 1

    return "".join(text_parts)


# =============================================================================
# STYLE SHEET PARSER
# =============================================================================


def _parse_stsh(table: bytes, fc: int, lcb: int) -> dict[int, str]:
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


# =============================================================================
# PARAGRAPH PROPERTIES (PAPX)
# =============================================================================


class _ParaProps:
    """Properties extracted from PAPX for a single paragraph."""

    def __init__(self):
        self.istd: int = 0
        self.ilfo: int = 0  # List format override index (1-based, 0 = no list)
        self.ilvl: int = 0  # List indent level


def _parse_sprms(grpprl: bytes) -> dict[int, bytes]:
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


def _parse_papx_fkp(wd: bytes, pn: int, text_start_fc: int) -> list[tuple[int, int, _ParaProps]]:
    """Parse a PAPX FKP page and return (char_start, char_end, props) tuples."""
    page_offset = pn * 512
    if page_offset + 512 > len(wd):
        return []

    page = wd[page_offset : page_offset + 512]
    cpara = page[511]
    if cpara == 0:
        return []

    results: list[tuple[int, int, _ParaProps]] = []

    for j in range(cpara):
        fc_j = struct.unpack_from("<I", page, j * 4)[0]
        fc_next = struct.unpack_from("<I", page, (j + 1) * 4)[0]

        # Convert FC to character positions
        char_start = (fc_j - text_start_fc) // 2
        char_end = (fc_next - text_start_fc) // 2

        props = _ParaProps()

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
                    sprms = _parse_sprms(grpprl)
                    _apply_para_sprms(props, sprms)
        else:
            total = cb * 2
            if papx_pos + 1 + total <= 512 and total >= 2:
                props.istd = struct.unpack_from("<H", page, papx_pos + 1)[0]
                grpprl = page[papx_pos + 3 : papx_pos + 1 + total]
                sprms = _parse_sprms(grpprl)
                _apply_para_sprms(props, sprms)

        results.append((char_start, char_end, props))

    return results


def _apply_para_sprms(props: _ParaProps, sprms: dict[int, bytes]) -> None:
    """Apply paragraph sprms to ParaProps."""
    if _SPRM_PILVL in sprms:
        props.ilvl = sprms[_SPRM_PILVL][0]
    if _SPRM_PILFO in sprms:
        props.ilfo = struct.unpack_from("<H", sprms[_SPRM_PILFO])[0]


# =============================================================================
# CHARACTER PROPERTIES (CHPX)
# =============================================================================


class _CharProps:
    """Properties extracted from CHPX for a text run."""

    def __init__(self):
        self.bold: bool = False
        self.italic: bool = False
        self.underline: bool = False
        self.strikethrough: bool = False


def _parse_chpx_fkp(wd: bytes, pn: int, text_start_fc: int) -> list[tuple[int, int, _CharProps]]:
    """Parse a CHPX FKP page and return (char_start, char_end, props) tuples."""
    page_offset = pn * 512
    if page_offset + 512 > len(wd):
        return []

    page = wd[page_offset : page_offset + 512]
    crun = page[511]
    if crun == 0:
        return []

    results: list[tuple[int, int, _CharProps]] = []

    for j in range(crun):
        fc_j = struct.unpack_from("<I", page, j * 4)[0]
        fc_next = struct.unpack_from("<I", page, (j + 1) * 4)[0]

        char_start = (fc_j - text_start_fc) // 2
        char_end = (fc_next - text_start_fc) // 2

        props = _CharProps()

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
            sprms = _parse_sprms(grpprl)
            _apply_char_sprms(props, sprms)

        results.append((char_start, char_end, props))

    return results


def _apply_char_sprms(props: _CharProps, sprms: dict[int, bytes]) -> None:
    """Apply character sprms to CharProps."""
    if _SPRM_CFBOLD in sprms and sprms[_SPRM_CFBOLD][0] == 1:
        props.bold = True
    if _SPRM_CFITALIC in sprms and sprms[_SPRM_CFITALIC][0] == 1:
        props.italic = True
    if _SPRM_CFSTRIKE in sprms and sprms[_SPRM_CFSTRIKE][0] == 1:
        props.strikethrough = True
    if _SPRM_CKUL in sprms and sprms[_SPRM_CKUL][0] != 0:
        props.underline = True


# =============================================================================
# LIST FORMAT PARSER
# =============================================================================


class _ListDef:
    """Parsed list definition."""

    def __init__(self):
        self.lsid: int = 0
        self.is_hybrid: bool = False  # hybrid = bullet list in Aspose output


def _parse_list_defs(table: bytes, fc_lst: int, lcb_lst: int) -> dict[int, _ListDef]:
    """Parse PlfLst to get list definitions keyed by lsid."""
    if lcb_lst == 0:
        return {}

    lst_data = table[fc_lst : fc_lst + lcb_lst]
    if len(lst_data) < 2:
        return {}

    c_lst = struct.unpack_from("<H", lst_data, 0)[0]
    result: dict[int, _ListDef] = {}

    for i in range(c_lst):
        offset = 2 + i * 28
        if offset + 28 > len(lst_data):
            break

        ld = _ListDef()
        ld.lsid = struct.unpack_from("<I", lst_data, offset)[0]
        flags = lst_data[offset + 26]
        ld.is_hybrid = bool(flags & 0x10)
        result[ld.lsid] = ld

    return result


def _parse_lfo_map(table: bytes, fc_lfo: int, lcb_lfo: int) -> dict[int, int]:
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


# =============================================================================
# TABLE TEXT PARSER
# =============================================================================


def _parse_table_text(text: str) -> TableData:
    """Parse Word table text (using \\x07 cell separators) into TableData."""
    data = TableData()
    rows_text = text.split("\x07\r")

    for row_text in rows_text:
        row_text = row_text.strip("\x07").strip("\r")
        if not row_text:
            continue

        row = RowData()
        cell_texts = row_text.split("\x07")
        # Last element after split is often empty or the row-end marker
        for ct in cell_texts:
            cell = CellData()
            para = ParagraphData()
            para.text = ct
            run = RunData()
            run.text = ct
            para.runs = [run]
            cell.paragraphs = [para]
            row.cells.append(cell)
        data.rows.append(row)

    return data


# =============================================================================
# HYPERLINK PARSER
# =============================================================================


def _parse_hyperlink(text: str) -> tuple[str, str]:
    """Parse a Word field code hyperlink.

    Field codes: \\x13 HYPERLINK "url" \\x01\\x14display_text\\x15
    Returns (display_text, url).
    """
    url = ""
    display = ""

    # Extract URL from between quotes after HYPERLINK
    begin_idx = text.find("\x13")
    sep_idx = text.find("\x14")
    end_idx = text.find("\x15")

    if begin_idx >= 0 and sep_idx >= 0:
        field_code = text[begin_idx + 1 : sep_idx]
        # Remove the \x01 separator if present
        field_code = field_code.replace("\x01", "").strip()

        # Extract URL from HYPERLINK "url"
        if "HYPERLINK" in field_code:
            parts = field_code.split('"')
            if len(parts) >= 2:
                url = parts[1]

    if sep_idx >= 0 and end_idx >= 0:
        display = text[sep_idx + 1 : end_idx]
    elif sep_idx >= 0:
        display = text[sep_idx + 1 :]

    return display, url


# =============================================================================
# DOC DOCUMENT READER
# =============================================================================


class DocFileReader:
    """
    Reads Word 97-2003 (.doc) files and produces the same data structures
    as DocumentReader (for .docx), enabling transparent format support.

    Uses olefile for OLE2 container access and custom parsing for
    the Word Binary File Format structures.
    """

    def __init__(self):
        self._text: str = ""
        self._para_props: list[tuple[int, int, _ParaProps]] = []
        self._char_props: list[tuple[int, int, _CharProps]] = []
        self._styles: dict[int, str] = {}
        self._list_defs: dict[int, _ListDef] = {}
        self._lfo_map: dict[int, int] = {}
        self._text_start_fc: int = 0x800 * 2  # Default for UTF-16LE at 0x800

    def load_file(self, filepath: Union[str, Path]) -> None:
        """Load .doc from file path."""
        ole = olefile.OleFileIO(str(filepath))
        try:
            self._load_from_ole(ole)
        finally:
            ole.close()

    def load_stream(self, stream: BinaryIO) -> None:
        """Load .doc from stream."""
        ole = olefile.OleFileIO(stream)
        try:
            self._load_from_ole(ole)
        finally:
            ole.close()

    def load_bytes(self, data: bytes) -> None:
        """Load .doc from bytes."""
        ole = olefile.OleFileIO(BytesIO(data))
        try:
            self._load_from_ole(ole)
        finally:
            ole.close()

    def _load_from_ole(self, ole: olefile.OleFileIO) -> None:
        """Parse the OLE2 container and extract document data."""
        wd = ole.openstream("WordDocument").read()
        fib = _parse_fib(wd)

        table_name = "1Table" if fib.f_which_tbl_stm else "0Table"
        table = ole.openstream(table_name).read()

        # Extract text
        self._text = _extract_text_via_piece_table(wd, table, fib)

        # Determine text start FC (file character offset)
        fc_clx, lcb_clx = _get_fc_lcb(wd, fib, _IDX_CLX)
        if lcb_clx > 0:
            clx = table[fc_clx : fc_clx + lcb_clx]
            self._text_start_fc = self._get_first_fc_from_clx(clx)
        else:
            self._text_start_fc = 0x800 * 2  # UTF-16LE offset

        # Parse styles
        fc_stsh, lcb_stsh = _get_fc_lcb(wd, fib, _IDX_STSHF)
        self._styles = _parse_stsh(table, fc_stsh, lcb_stsh)

        # Parse paragraph properties
        self._para_props = self._collect_papx(wd, table, fib)

        # Parse character properties
        self._char_props = self._collect_chpx(wd, table, fib)

        # Parse list definitions
        fc_lst, lcb_lst = _get_fc_lcb(wd, fib, _IDX_PLFLST)
        self._list_defs = _parse_list_defs(table, fc_lst, lcb_lst)

        fc_lfo, lcb_lfo = _get_fc_lcb(wd, fib, _IDX_PLFLFO)
        self._lfo_map = _parse_lfo_map(table, fc_lfo, lcb_lfo)

    def _get_first_fc_from_clx(self, clx: bytes) -> int:
        """Get the FC of the first piece from CLX."""
        pos = 0
        while pos < len(clx):
            clxt = clx[pos]
            if clxt == 0x02:
                pos += 1
                pcdt_size = struct.unpack_from("<I", clx, pos)[0]
                pos += 4
                pt = clx[pos : pos + pcdt_size]
                n = (pcdt_size - 4) // 12
                if n > 0:
                    pcd_start = (n + 1) * 4
                    fc_val = struct.unpack_from("<I", pt, pcd_start + 2)[0]
                    fc_real = fc_val & 0x3FFFFFFF
                    return fc_real
                break
            elif clxt == 0x01:
                pos += 1
                cb = struct.unpack_from("<H", clx, pos)[0]
                pos += 2 + cb
            else:
                pos += 1
        return 0x800 * 2

    def _collect_papx(
        self, wd: bytes, table: bytes, fib: _FibData
    ) -> list[tuple[int, int, _ParaProps]]:
        """Collect all paragraph properties from PAPX FKP pages."""
        fc_papx, lcb_papx = _get_fc_lcb(wd, fib, _IDX_PLCFBTEPAPX)
        if lcb_papx == 0:
            return []

        papx_data = table[fc_papx : fc_papx + lcb_papx]
        n = (lcb_papx - 4) // 8
        if n <= 0:
            return []

        text_start = self._text_start_fc
        all_props: list[tuple[int, int, _ParaProps]] = []

        for i in range(n):
            bte_offset = (n + 1) * 4 + i * 4
            pn = struct.unpack_from("<I", papx_data, bte_offset)[0]
            all_props.extend(_parse_papx_fkp(wd, pn, text_start))

        return sorted(all_props, key=lambda x: x[0])

    def _collect_chpx(
        self, wd: bytes, table: bytes, fib: _FibData
    ) -> list[tuple[int, int, _CharProps]]:
        """Collect all character properties from CHPX FKP pages."""
        fc_chpx, lcb_chpx = _get_fc_lcb(wd, fib, _IDX_PLCFBTECHPX)
        if lcb_chpx == 0:
            return []

        chpx_data = table[fc_chpx : fc_chpx + lcb_chpx]
        n = (lcb_chpx - 4) // 8
        if n <= 0:
            return []

        text_start = self._text_start_fc
        all_props: list[tuple[int, int, _CharProps]] = []

        for i in range(n):
            bte_offset = (n + 1) * 4 + i * 4
            pn = struct.unpack_from("<I", chpx_data, bte_offset)[0]
            all_props.extend(_parse_chpx_fkp(wd, pn, text_start))

        return sorted(all_props, key=lambda x: x[0])

    def _get_para_props_at(self, char_pos: int) -> _ParaProps:
        """Find paragraph properties for a character position."""
        for start, end, props in self._para_props:
            if start <= char_pos < end:
                return props
        return _ParaProps()

    def _get_char_props_in_range(self, start: int, end: int) -> list[tuple[int, int, _CharProps]]:
        """Get all character property ranges overlapping [start, end)."""
        result = []
        for cs, ce, cp in self._char_props:
            if ce <= start:
                continue
            if cs >= end:
                break
            result.append((max(cs, start), min(ce, end), cp))
        return result

    def _is_bullet_list(self, ilfo: int) -> bool:
        """Check if the given LFO index refers to a bullet list."""
        lsid = self._lfo_map.get(ilfo)
        if lsid is not None:
            ld = self._list_defs.get(lsid)
            if ld is not None:
                return ld.is_hybrid
        return True  # Default to bullet if unknown

    def _iterate_body_elements(self) -> Iterator[Union[ParagraphData, TableData]]:
        """Iterate over document body elements in order."""
        text = self._text
        if not text:
            return

        # Split text into paragraphs by \r
        paragraphs: list[tuple[int, int]] = []
        start = 0
        for i, ch in enumerate(text):
            if ch == "\r":
                paragraphs.append((start, i))
                start = i + 1

        # Pre-process: assign group IDs to consecutive list items
        # so that items in the same logical list share a counter
        self._list_group_ids: dict[int, int] = {}
        self._assign_list_groups(paragraphs, text)

        # Process paragraphs, detecting tables and hyperlinks
        i = 0
        while i < len(paragraphs):
            p_start, p_end = paragraphs[i]
            para_text = text[p_start:p_end]

            # Check if this paragraph contains table data (\x07 = cell marker)
            if "\x07" in para_text:
                table_data, trailing_text = self._build_table_from_text(para_text)
                if table_data.rows:
                    yield table_data
                # If there's trailing text after the table, treat it as paragraphs
                if trailing_text.strip():
                    # Find where this trailing text starts in the full document
                    trail_start = p_start + len(para_text) - len(trailing_text)
                    yield self._build_paragraph(trailing_text, trail_start, p_end)
                i += 1
                continue

            # Skip empty trailing paragraphs
            if not para_text.strip():
                i += 1
                continue

            # Build paragraph data
            yield self._build_paragraph(para_text, p_start, p_end)
            i += 1

    def _assign_list_groups(self, paragraphs: list[tuple[int, int]], text: str) -> None:
        """Assign group IDs to consecutive list items of the same type."""
        current_group_id = 1000  # Start with a high number to avoid conflicts
        prev_was_list = False
        prev_list_type_bullet = False

        for p_start, p_end in paragraphs:
            para_text = text[p_start:p_end]
            if "\x07" in para_text or not para_text.strip():
                prev_was_list = False
                continue

            props = self._get_para_props_at(p_start)
            if props.ilfo > 0:
                is_bullet = self._is_bullet_list(props.ilfo)
                if not prev_was_list or is_bullet != prev_list_type_bullet:
                    current_group_id += 1
                    prev_list_type_bullet = is_bullet

                self._list_group_ids[p_start] = current_group_id
                prev_was_list = True
            else:
                prev_was_list = False

    def _build_paragraph(self, para_text: str, p_start: int, p_end: int) -> ParagraphData:
        """Build a ParagraphData from text and properties."""
        data = ParagraphData()

        # Check for hyperlink field codes
        if "\x13" in para_text and "\x14" in para_text:
            return self._build_hyperlink_paragraph(para_text, p_start, p_end)

        # Get paragraph properties
        props = self._get_para_props_at(p_start)
        style_name = self._styles.get(props.istd, "Normal")
        data.style_name = style_name

        # Check for list
        if props.ilfo > 0:
            data.is_list_item = True
            data.list_level = props.ilvl
            # Use group ID so consecutive items share a counter
            data.list_id = self._list_group_ids.get(p_start, props.ilfo)

        # Build runs with character formatting
        char_ranges = self._get_char_props_in_range(p_start, p_end)
        if char_ranges:
            for cs, ce, cp in char_ranges:
                run_text = self._text[cs:ce]
                if not run_text:
                    continue
                run = RunData(
                    text=run_text,
                    bold=cp.bold,
                    italic=cp.italic,
                    underline=cp.underline,
                    strikethrough=cp.strikethrough,
                )
                data.runs.append(run)
        else:
            run = RunData(text=para_text)
            data.runs.append(run)

        data.text = "".join(r.text for r in data.runs)
        return data

    def _build_hyperlink_paragraph(self, para_text: str, p_start: int, p_end: int) -> ParagraphData:
        """Build a paragraph containing a hyperlink."""
        data = ParagraphData()
        data.style_name = "Normal"

        # Split on field boundaries
        parts = []
        remaining = para_text
        while "\x13" in remaining:
            pre_idx = remaining.find("\x13")
            if pre_idx > 0:
                parts.append(("text", remaining[:pre_idx]))

            end_idx = remaining.find("\x15", pre_idx)
            if end_idx < 0:
                end_idx = len(remaining) - 1

            field_text = remaining[pre_idx : end_idx + 1]
            display, url = _parse_hyperlink(field_text)
            if display and url:
                parts.append(("link", display, url))
            elif display:
                parts.append(("text", display))

            remaining = remaining[end_idx + 1 :]

        if remaining:
            parts.append(("text", remaining))

        # Build runs
        for part in parts:
            if part[0] == "text":
                text = part[1]
                if text:
                    run = RunData(text=text)
                    data.runs.append(run)
            elif part[0] == "link":
                display_text = part[1]
                link_url = part[2]
                run = RunData(text=f"[{display_text}]({link_url})")
                data.runs.append(run)

        data.text = "".join(r.text for r in data.runs)
        return data

    def _build_table_from_text(self, table_text: str) -> tuple[TableData, str]:
        """Build a TableData from table text with \\x07 separators.

        In .doc format, table cells end with \\x07 and rows end with
        an additional \\x07. So a row with cells A and B looks like:
        A\\x07B\\x07\\x07

        Returns (TableData, trailing_text) where trailing_text is any
        text after the last table row.
        """
        data = TableData()
        trailing = ""

        # Split by \x07 to get all segments
        segments = table_text.split("\x07")

        # Group segments into rows: empty segments indicate row boundaries
        current_row: list[str] = []
        last_row_end_idx = -1

        for idx, seg in enumerate(segments):
            if seg == "":
                # Row boundary: save current row if it has cells
                if current_row:
                    row = RowData()
                    for ct in current_row:
                        cell = CellData()
                        para = ParagraphData()
                        para.text = ct
                        run = RunData(text=ct)
                        para.runs = [run]
                        cell.paragraphs = [para]
                        row.cells.append(cell)
                    data.rows.append(row)
                    current_row = []
                    last_row_end_idx = idx
            else:
                current_row.append(seg)

        # Any remaining non-empty text after the last row is trailing text
        if current_row:
            trailing = "\x07".join(current_row)

        return data, trailing

    def _get_list_format(self, num_id: int, level: int) -> tuple[str, int]:
        """Get list format and start number for a given list id and level.

        Compatible with DocumentReader._get_list_format().
        The num_id here is the group ID assigned by _assign_list_groups.
        """
        # Check if this group ID maps to a bullet list
        # Look up any paragraph that belongs to this group
        for p_start, gid in self._list_group_ids.items():
            if gid == num_id:
                props = self._get_para_props_at(p_start)
                if props.ilfo > 0:
                    if self._is_bullet_list(props.ilfo):
                        return "bullet", 1
                    else:
                        return "decimal", 1
                break
        return "bullet", 1

    def _get_numbering_info(self, num_id: int) -> Optional[NumberingInfo]:
        """Get numbering info (for compatibility with DocumentReader)."""
        is_bullet = self._is_bullet_list(num_id)
        fmt = "bullet" if is_bullet else "decimal"

        info = NumberingInfo(
            num_id=num_id,
            abstract_num_id=num_id,
            levels={0: NumberingLevel(format=fmt, start=1)},
        )
        return info

    # =========================================================================
    # LIGHT DOCUMENT MODEL BUILDER
    # =========================================================================

    def to_light_document(self) -> ldm.Document:
        """Build a light_document_model.Document from the loaded DOC file.

        Returns a populated Pydantic model. DOC binary format provides
        less metadata than DOCX XML, so some fields will use defaults.
        """
        from aspose.words_foss import light_document_model as ldm

        doc = ldm.Document()

        # Build styles from the parsed style names
        doc.styles = self._build_ldm_styles()

        # Build lists from parsed list definitions
        doc.lists = self._build_ldm_lists()

        # Build a single section containing all body content
        sec = ldm.Section()
        sec.body = ldm.Body(children=self._build_ldm_body_children())
        doc.sections = [sec]

        return doc

    def _build_ldm_styles(self) -> list[ldm.Style]:
        from aspose.words_foss import light_document_model as ldm

        styles: list[ldm.Style] = []
        for istd, name in self._styles.items():
            s = ldm.Style()
            s.name = name
            s.type = 1  # paragraph style (default for DOC styles)

            heading_match = re.search(r"[Hh]eading\s*(\d+)", name)
            if heading_match:
                s.is_heading = True
                level = int(heading_match.group(1))
                pf = ldm.ParagraphFormat()
                pf.style_name = name
                pf.is_heading = True
                pf.outline_level = level - 1
                s.paragraph_format = pf

            styles.append(s)
        return styles

    def _build_ldm_lists(self) -> list[ldm.DocList]:
        from aspose.words_foss import light_document_model as ldm

        lists: list[ldm.DocList] = []
        for lsid, list_def in self._list_defs.items():
            dl = ldm.DocList()
            dl.list_id = lsid
            dl.is_multi_level = False

            ll = ldm.ListLevel()
            if list_def.is_hybrid:
                ll.number_style = 23  # bullet
            else:
                ll.number_style = 0  # decimal
                ll.number_format = "%1."
            dl.levels = [ll]
            lists.append(dl)
        return lists

    def _build_ldm_body_children(
        self,
    ) -> list[ldm.Paragraph | ldm.Table | ldm.UnknownNode]:
        from aspose.words_foss import light_document_model as ldm

        children: list[ldm.Paragraph | ldm.Table | ldm.UnknownNode] = []
        text = self._text
        if not text:
            return children

        # Split text into paragraphs by \r
        paragraphs: list[tuple[int, int]] = []
        start = 0
        for i, ch in enumerate(text):
            if ch == "\r":
                paragraphs.append((start, i))
                start = i + 1

        for p_start, p_end in paragraphs:
            para_text = text[p_start:p_end]

            # Table detection
            if "\x07" in para_text:
                tbl, trailing = self._build_ldm_table_from_text(para_text)
                if tbl.rows:
                    children.append(tbl)
                if trailing.strip():
                    trail_start = p_start + len(para_text) - len(trailing)
                    if "\x13" in trailing and "\x14" in trailing:
                        para = self._build_ldm_hyperlink_paragraph(trailing, trail_start, p_end)
                    else:
                        para = self._build_ldm_paragraph(trailing, trail_start, p_end)
                    children.append(para)
                continue

            # Skip empty paragraphs
            if not para_text.strip():
                continue

            # Handle hyperlink field codes
            if "\x13" in para_text and "\x14" in para_text:
                para = self._build_ldm_hyperlink_paragraph(para_text, p_start, p_end)
            else:
                para = self._build_ldm_paragraph(para_text, p_start, p_end)
            children.append(para)

        return children

    def _build_ldm_paragraph(self, para_text: str, p_start: int, p_end: int) -> ldm.Paragraph:
        from aspose.words_foss import light_document_model as ldm

        para = ldm.Paragraph()
        props = self._get_para_props_at(p_start)
        style_name = self._styles.get(props.istd, "Normal")

        # Paragraph format
        pf = ldm.ParagraphFormat()
        pf.style_name = style_name

        heading_match = re.search(r"[Hh]eading\s*(\d+)", style_name)
        if heading_match:
            pf.is_heading = True
            pf.outline_level = int(heading_match.group(1)) - 1

        if props.ilfo > 0:
            pf.is_list_item = True
            lf = ldm.ListFormat()
            lf.is_list_item = True
            lf.list_level_number = props.ilvl
            lsid = self._lfo_map.get(props.ilfo)
            if lsid is not None:
                lf.list_id = lsid
            else:
                lf.list_id = props.ilfo
            para.list_format = lf

        para.paragraph_format = pf

        # Build runs with character formatting
        char_ranges = self._get_char_props_in_range(p_start, p_end)
        text_parts: list[str] = []

        if char_ranges:
            for cs, ce, cp in char_ranges:
                run_text = self._text[cs:ce]
                if not run_text:
                    continue
                run = ldm.Run()
                run.text = run_text
                font = ldm.Font()
                font.bold = cp.bold
                font.italic = cp.italic
                font.strike_through = cp.strikethrough
                if cp.underline:
                    font.underline = 1
                run.font = font
                para.runs.append(run)
                text_parts.append(run_text)
        else:
            run = ldm.Run()
            run.text = para_text
            para.runs.append(run)
            text_parts.append(para_text)

        para.text = "".join(text_parts)
        return para

    def _build_ldm_hyperlink_paragraph(
        self, para_text: str, p_start: int, p_end: int
    ) -> ldm.Paragraph:
        """Build an LDM paragraph containing hyperlink field codes."""
        from aspose.words_foss import light_document_model as ldm

        para = ldm.Paragraph()
        para.paragraph_format = ldm.ParagraphFormat(style_name="Normal")

        parts: list[tuple] = []
        remaining = para_text
        while "\x13" in remaining:
            pre_idx = remaining.find("\x13")
            if pre_idx > 0:
                parts.append(("text", remaining[:pre_idx]))

            end_idx = remaining.find("\x15", pre_idx)
            if end_idx < 0:
                end_idx = len(remaining) - 1

            field_text = remaining[pre_idx : end_idx + 1]
            display, url = _parse_hyperlink(field_text)
            if display and url:
                parts.append(("link", display, url))
            elif display:
                parts.append(("text", display))

            remaining = remaining[end_idx + 1 :]

        if remaining:
            parts.append(("text", remaining))

        text_pieces: list[str] = []
        for part in parts:
            if part[0] == "text":
                text = part[1]
                if text:
                    run = ldm.Run()
                    run.text = text
                    para.runs.append(run)
                    text_pieces.append(text)
            elif part[0] == "link":
                display_text = part[1]
                link_url = part[2]
                run = ldm.Run()
                run.text = f"[{display_text}]({link_url})"
                para.runs.append(run)
                text_pieces.append(run.text)

        para.text = "".join(text_pieces)
        return para

    def _build_ldm_table_from_text(self, table_text: str) -> tuple[ldm.Table, str]:
        """Build an LDM Table from table text with \\x07 separators.

        Returns ``(table, trailing_text)`` where *trailing_text* is any
        content after the last complete table row.
        """
        from aspose.words_foss import light_document_model as ldm

        tbl = ldm.Table()
        trailing = ""
        segments = table_text.split("\x07")
        current_row: list[str] = []

        for seg in segments:
            if seg == "":
                if current_row:
                    row = ldm.Row()
                    for ct in current_row:
                        cell = ldm.Cell()
                        para = ldm.Paragraph()
                        para.text = ct
                        run = ldm.Run()
                        run.text = ct
                        para.runs = [run]
                        cell.paragraphs = [para]
                        row.cells.append(cell)
                    tbl.rows.append(row)
                    current_row = []
            else:
                current_row.append(seg)

        if current_row:
            trailing = "\x07".join(current_row)

        return tbl, trailing
