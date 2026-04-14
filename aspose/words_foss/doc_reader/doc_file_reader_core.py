"""
Core DOC file reader — loading, parsing, property resolution, Markdown iteration.
"""

from __future__ import annotations

import re
import struct
from io import BytesIO
from pathlib import Path
from typing import Optional, Union, BinaryIO, Iterator

import olefile

from aspose.words_foss.doc_reader.constants import (
    IDX_CLX,
    IDX_DGGINFO,
    IDX_PLCFBTECHPX,
    IDX_PLCFBTEPAPX,
    IDX_PLCFSED,
    IDX_PLCFHDD,
    IDX_PLCFTXBXTXT,
    IDX_PLFLFO,
    IDX_PLFLST,
    IDX_PLCSPAHDR,
    IDX_PLCSPAMOM,
    IDX_STSHF,
    IDX_STTBFFFN,
)
from aspose.words_foss.doc_reader.fib import FibData, parse_fib, get_fc_lcb
from aspose.words_foss.doc_reader.fonts import parse_font_table
from aspose.words_foss.doc_reader.images import (
    BlipInfo,
    ChildAnchorInfo,
    GroupShapeInfo,
    ShapeAnchor,
    ShapeLineProps,
    parse_blip_store,
    parse_escher_child_positions,
    parse_plc_spa,
    parse_plcf_txbx_txt,
    parse_shape_blip_map,
    parse_shape_line_map,
    parse_shape_txid_map,
)
from aspose.words_foss.doc_reader.lists import ListDef, parse_list_defs, parse_lfo_map
from aspose.words_foss.doc_reader.properties import (
    CharProps,
    ParaProps,
    parse_chpx_fkp,
    parse_papx_fkp,
    parse_sepx_sprms,
)
from aspose.words_foss.doc_reader.styles import StyleData, parse_stsh_full
from aspose.words_foss.doc_reader.text import (
    extract_text_via_piece_table,
    parse_hyperlink,
)
from aspose.words_foss.docx_reader import (
    CellData,
    NumberingInfo,
    NumberingLevel,
    ParagraphData,
    RowData,
    RunData,
    TableData,
)


class DocFileReaderCore:
    """
    Core reader for Word 97-2003 (.doc) files.

    Loads the OLE2 container, parses binary structures, and provides
    property resolution and Markdown-oriented iteration.
    """

    def __init__(self):
        self._text: str = ""
        self._para_props: list[tuple[int, int, ParaProps]] = []
        self._char_props: list[tuple[int, int, CharProps]] = []
        self._styles: dict[int, str] = {}
        self._style_data: dict[int, StyleData] = {}
        self._fonts: dict[int, str] = {}
        self._default_font_index: int = 0
        self._list_defs: dict[int, ListDef] = {}
        self._lfo_map: dict[int, int] = {}
        self._text_byte_offset: int = 0x800
        self._text_is_compressed: bool = True
        self._ccp_text: int = 0
        self._ccp_ftn: int = 0
        self._ccp_hdd: int = 0
        self._ccp_txbx: int = 0
        self._section_props: list[dict] = []
        self._hdd_cps: list[int] = []
        # OfficeArt image data
        self._blips: list[BlipInfo] = []
        self._shape_blip_map: dict[int, int] = {}
        self._shape_line_map: dict[int, ShapeLineProps] = {}
        self._body_shape_anchors: list[ShapeAnchor] = []
        self._hdr_shape_anchors: list[ShapeAnchor] = []
        self._wd_bytes: bytes = b""
        # FTXBXS: shape ID → (cp_start, cp_end) in textbox text stream
        self._textbox_stories: dict[int, tuple[int, int]] = {}
        # Escher child shape positions and group coordinate systems
        self._child_anchors: dict[int, ChildAnchorInfo] = {}
        self._group_info: dict[int, GroupShapeInfo] = {}
        self._child_to_parent: dict[int, int] = {}  # child spid → group spid
        # FOpt txid: shape ID → textbox story index (0-based)
        self._shape_txid_map: dict[int, int] = {}

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
        fib = parse_fib(wd)

        table_name = "1Table" if fib.f_which_tbl_stm else "0Table"
        table = ole.openstream(table_name).read()

        # Extract text
        self._text = extract_text_via_piece_table(wd, table, fib)
        self._ccp_text = fib.ccp_text
        self._ccp_ftn = fib.ccp_ftn
        self._ccp_hdd = fib.ccp_hdd
        self._ccp_txbx = fib.ccp_txbx

        # Determine text byte offset and compression from piece table
        fc_clx, lcb_clx = get_fc_lcb(wd, fib, IDX_CLX)
        if lcb_clx > 0:
            clx = table[fc_clx : fc_clx + lcb_clx]
            self._text_byte_offset, self._text_is_compressed = self._get_text_offset_from_clx(clx)
        else:
            self._text_byte_offset = 0x800
            self._text_is_compressed = True

        # Parse styles (full parse with UPX property data)
        fc_stsh, lcb_stsh = get_fc_lcb(wd, fib, IDX_STSHF)
        self._styles, self._style_data, self._default_font_index = parse_stsh_full(
            table, fc_stsh, lcb_stsh
        )

        # Parse font table
        fc_ffn, lcb_ffn = get_fc_lcb(wd, fib, IDX_STTBFFFN)
        self._fonts = parse_font_table(table, fc_ffn, lcb_ffn)

        # Parse paragraph properties
        self._para_props = self._collect_papx(wd, table, fib)

        # Parse character properties
        self._char_props = self._collect_chpx(wd, table, fib)

        # Parse list definitions
        fc_lst, lcb_lst = get_fc_lcb(wd, fib, IDX_PLFLST)
        self._list_defs = parse_list_defs(table, fc_lst, lcb_lst)

        fc_lfo, lcb_lfo = get_fc_lcb(wd, fib, IDX_PLFLFO)
        self._lfo_map = parse_lfo_map(table, fc_lfo, lcb_lfo)

        # Parse section properties (page setup)
        fc_sed, lcb_sed = get_fc_lcb(wd, fib, IDX_PLCFSED)
        self._section_props = self._parse_plcf_sed(wd, table, fc_sed, lcb_sed)

        # Parse header/footer CP positions
        fc_hdd, lcb_hdd = get_fc_lcb(wd, fib, IDX_PLCFHDD)
        if lcb_hdd > 0:
            hdd_data = table[fc_hdd : fc_hdd + lcb_hdd]
            n_cp = lcb_hdd // 4
            self._hdd_cps = [struct.unpack_from("<I", hdd_data, i * 4)[0] for i in range(n_cp)]

        # Parse OfficeArt images and shapes
        self._wd_bytes = wd
        fc_dgg, lcb_dgg = get_fc_lcb(wd, fib, IDX_DGGINFO)
        if lcb_dgg > 0:
            self._blips = parse_blip_store(table, wd, fc_dgg, lcb_dgg)
            self._shape_blip_map = parse_shape_blip_map(table, fc_dgg, lcb_dgg)
            self._shape_line_map = parse_shape_line_map(table, fc_dgg, lcb_dgg)
            ca, gi, ctp = parse_escher_child_positions(table, fc_dgg, lcb_dgg)
            self._child_anchors = ca
            self._group_info = gi
            self._child_to_parent = ctp
            self._shape_txid_map = parse_shape_txid_map(table, fc_dgg, lcb_dgg)

        fc_spa, lcb_spa = get_fc_lcb(wd, fib, IDX_PLCSPAMOM)
        self._body_shape_anchors = parse_plc_spa(table, fc_spa, lcb_spa)

        fc_spa_h, lcb_spa_h = get_fc_lcb(wd, fib, IDX_PLCSPAHDR)
        self._hdr_shape_anchors = parse_plc_spa(table, fc_spa_h, lcb_spa_h)

        # Parse textbox story → shape mapping (FTXBXS)
        fc_txbx, lcb_txbx = get_fc_lcb(wd, fib, IDX_PLCFTXBXTXT)
        if lcb_txbx > 0:
            self._textbox_stories = parse_plcf_txbx_txt(table, fc_txbx, lcb_txbx)

    @staticmethod
    def _parse_plcf_sed(wd: bytes, table: bytes, fc_sed: int, lcb_sed: int) -> list[dict]:
        """Parse PlcfSed and SEPX records to extract section properties."""
        if lcb_sed == 0:
            return []

        sed_data = table[fc_sed : fc_sed + lcb_sed]
        n = (lcb_sed - 4) // 16
        if n <= 0:
            return []

        sections: list[dict] = []
        sed_start = (n + 1) * 4
        for i in range(n):
            off = sed_start + i * 12
            fc_sepx = struct.unpack_from("<I", sed_data, off + 2)[0]
            props: dict = {}

            if fc_sepx > 0 and fc_sepx != 0xFFFFFFFF and fc_sepx + 2 <= len(wd):
                cb = struct.unpack_from("<H", wd, fc_sepx)[0]
                if cb > 0 and fc_sepx + 2 + cb <= len(wd):
                    sepx = wd[fc_sepx + 2 : fc_sepx + 2 + cb]
                    props = parse_sepx_sprms(sepx)

            sections.append(props)

        return sections

    def _get_text_offset_from_clx(self, clx: bytes) -> tuple[int, bool]:
        """Get the byte offset and compression flag of text from CLX."""
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
                    f_compressed = bool(fc_val & 0x40000000)
                    fc_real = fc_val & 0x3FFFFFFF
                    if f_compressed:
                        return fc_real // 2, True
                    else:
                        return fc_real, False
                break
            elif clxt == 0x01:
                pos += 1
                cb = struct.unpack_from("<H", clx, pos)[0]
                pos += 2 + cb
            else:
                pos += 1
        return 0x800, True

    def _collect_papx(
        self, wd: bytes, table: bytes, fib: FibData
    ) -> list[tuple[int, int, ParaProps]]:
        """Collect all paragraph properties from PAPX FKP pages."""
        fc_papx, lcb_papx = get_fc_lcb(wd, fib, IDX_PLCFBTEPAPX)
        if lcb_papx == 0:
            return []

        papx_data = table[fc_papx : fc_papx + lcb_papx]
        n = (lcb_papx - 4) // 8
        if n <= 0:
            return []

        all_props: list[tuple[int, int, ParaProps]] = []

        for i in range(n):
            bte_offset = (n + 1) * 4 + i * 4
            pn = struct.unpack_from("<I", papx_data, bte_offset)[0]
            all_props.extend(
                parse_papx_fkp(wd, pn, self._text_byte_offset, self._text_is_compressed)
            )

        return sorted(all_props, key=lambda x: x[0])

    def _collect_chpx(
        self, wd: bytes, table: bytes, fib: FibData
    ) -> list[tuple[int, int, CharProps]]:
        """Collect all character properties from CHPX FKP pages."""
        fc_chpx, lcb_chpx = get_fc_lcb(wd, fib, IDX_PLCFBTECHPX)
        if lcb_chpx == 0:
            return []

        chpx_data = table[fc_chpx : fc_chpx + lcb_chpx]
        n = (lcb_chpx - 4) // 8
        if n <= 0:
            return []

        all_props: list[tuple[int, int, CharProps]] = []

        for i in range(n):
            bte_offset = (n + 1) * 4 + i * 4
            pn = struct.unpack_from("<I", chpx_data, bte_offset)[0]
            all_props.extend(
                parse_chpx_fkp(wd, pn, self._text_byte_offset, self._text_is_compressed)
            )

        return sorted(all_props, key=lambda x: x[0])

    def _get_para_props_at(self, char_pos: int) -> ParaProps:
        """Find paragraph properties for a character position."""
        for start, end, props in self._para_props:
            if start <= char_pos < end:
                return props
        return ParaProps()

    def _get_char_props_in_range(self, start: int, end: int) -> list[tuple[int, int, CharProps]]:
        """Get all character property ranges overlapping [start, end)."""
        result = []
        for cs, ce, cp in self._char_props:
            if ce <= start:
                continue
            if cs >= end:
                break
            result.append((max(cs, start), min(ce, end), cp))
        return result

    def _resolve_para_props(self, istd: int) -> ParaProps:
        """Resolve paragraph properties by walking the style chain."""
        result = ParaProps()
        chain = self._get_style_chain(istd)
        for sd in reversed(chain):
            for field in sd.para_props_set:
                setattr(result, field, getattr(sd.para_props, field))
                if field == "line_spacing":
                    result.line_spacing_rule = sd.para_props.line_spacing_rule
        return result

    def _resolve_char_props(self, istd: int) -> CharProps:
        """Resolve character properties by walking the style chain."""
        result = CharProps()
        chain = self._get_style_chain(istd)
        for sd in reversed(chain):
            for field in sd.char_props_set:
                setattr(result, field, getattr(sd.char_props, field))
                if field == "superscript":
                    result.subscript = sd.char_props.subscript
        return result

    def _get_style_chain(self, istd: int) -> list[StyleData]:
        """Walk the style base chain and return list from derived to base."""
        chain: list[StyleData] = []
        visited: set[int] = set()
        current = istd
        while current in self._style_data and current not in visited:
            visited.add(current)
            sd = self._style_data[current]
            chain.append(sd)
            if sd.istd_base == 0x0FFF:
                break
            current = sd.istd_base
        return chain

    def _merge_char_props(
        self, style_cp: CharProps, direct_cp: CharProps, direct_sprms: set[str]
    ) -> CharProps:
        """Merge direct character properties over style-resolved ones.

        Toggle SPRMs (operand 0x81) must be re-applied against the *style*
        value rather than using the CHPX result which was computed from
        defaults.  E.g. a style with bold=True and a toggle SPRM yields
        bold=False, not bold=True.
        """
        result = CharProps()
        for field in (
            "bold",
            "italic",
            "underline",
            "strikethrough",
            "font_index",
            "font_size",
            "color",
            "highlight_color",
            "superscript",
            "subscript",
            "all_caps",
            "small_caps",
            "hidden",
            "style_index",
        ):
            setattr(result, field, getattr(style_cp, field))
        for field in direct_sprms:
            if field in direct_cp._toggle_fields:
                # Toggle SPRM: flip the style's value instead of
                # blindly copying the CHPX value.
                style_val = getattr(style_cp, field)
                setattr(result, field, not style_val)
            else:
                setattr(result, field, getattr(direct_cp, field))
            if field == "superscript":
                result.subscript = direct_cp.subscript
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

        paragraphs: list[tuple[int, int]] = []
        start = 0
        for i, ch in enumerate(text):
            if ch == "\r":
                paragraphs.append((start, i))
                start = i + 1

        self._list_group_ids: dict[int, int] = {}
        self._assign_list_groups(paragraphs, text)

        i = 0
        while i < len(paragraphs):
            p_start, p_end = paragraphs[i]
            para_text = text[p_start:p_end]

            if "\x07" in para_text:
                table_data, trailing_text = self._build_table_from_text(para_text)
                if table_data.rows:
                    yield table_data
                if trailing_text.strip():
                    trail_start = p_start + len(para_text) - len(trailing_text)
                    yield self._build_paragraph(trailing_text, trail_start, p_end)
                i += 1
                continue

            if not para_text.strip():
                i += 1
                continue

            yield self._build_paragraph(para_text, p_start, p_end)
            i += 1

    def _assign_list_groups(self, paragraphs: list[tuple[int, int]], text: str) -> None:
        """Assign group IDs to consecutive list items of the same type."""
        current_group_id = 1000
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

        if "\x13" in para_text and "\x14" in para_text:
            return self._build_hyperlink_paragraph(para_text, p_start, p_end)

        props = self._get_para_props_at(p_start)
        style_name = self._styles.get(props.istd, "Normal")
        data.style_name = style_name

        if props.ilfo > 0:
            data.is_list_item = True
            data.list_level = props.ilvl
            data.list_id = self._list_group_ids.get(p_start, props.ilfo)

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
                    underline=cp.underline != 0,
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
            display, url = parse_hyperlink(field_text)
            if display and url:
                parts.append(("link", display, url))
            elif display:
                parts.append(("text", display))

            remaining = remaining[end_idx + 1 :]

        if remaining:
            parts.append(("text", remaining))

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
        """Build a TableData from table text with \\x07 separators."""
        data = TableData()
        trailing = ""
        segments = table_text.split("\x07")
        current_row: list[str] = []
        last_row_end_idx = -1

        for idx, seg in enumerate(segments):
            if seg == "":
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

        if current_row:
            trailing = "\x07".join(current_row)

        return data, trailing

    def _get_list_format(self, num_id: int, level: int) -> tuple[str, int]:
        """Get list format and start number for a given list id and level."""
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
