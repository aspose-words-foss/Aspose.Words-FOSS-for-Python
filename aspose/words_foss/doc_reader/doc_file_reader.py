"""
LDM builder for DOC files — extends DocFileReaderCore with to_light_document().
"""

import struct
from typing import Optional

# -- unit conversion constants ------------------------------------------------
_PT_PER_INCH = 72.0
_MM_PER_INCH = 25.4
_PT_TO_MM = _MM_PER_INCH / _PT_PER_INCH  # 1 pt ≈ 0.3528 mm

# -- Word default page dimensions (points) ------------------------------------
_DEFAULT_PAGE_WIDTH_PT = 612.0  # US Letter width (8.5 in)
_DEFAULT_PAGE_HEIGHT_PT = 792.0  # US Letter height (11 in)
_DEFAULT_MARGIN_PT = 72.0  # 1 inch
_DEFAULT_HEADER_FOOTER_DIST_PT = 36.0  # 0.5 inch

# Paper size detection: (width_pt, height_pt, paper_size_id)
_LETTER_WIDTH_PT = 612.0
_LETTER_HEIGHT_PT = 792.0
_A4_WIDTH_PT = 595.28
_A4_HEIGHT_PT = 841.89

# -- default font size --------------------------------------------------------
_DEFAULT_FONT_SIZE_PT = 10.0

# -- list-paragraph defaults --------------------------------------------------
# Word's default bullet/number list uses 36pt left-indent per level and a
# 18pt hanging first-line indent — the same defaults Aspose.Words returns
# from a paragraph's ``paragraph_format`` when the PAPX has no explicit
# indent SPRMs but the paragraph is part of a list.
_LIST_DEFAULT_LEVEL_INDENT_PT = 36.0
_LIST_DEFAULT_HANGING_PT = 18.0


from aspose.words_foss.doc_reader.constants import (
    ESCHER_BLIP_JPEG,
    ESCHER_BLIP_JPEG2,
    is_raster_blip,
)
from aspose.words_foss.doc_reader.doc_file_reader_core import DocFileReaderCore
from aspose.words_foss.doc_reader.images import ShapeAnchor
from aspose.words_foss.doc_reader.properties import CharProps, ParaProps
from aspose.words_foss.doc_reader.table_builder import DocTableBuilderMixin
from aspose.words_foss.doc_reader.text import clean_control_chars, evaluate_fields
from aspose.words_foss.model.style_identifiers import (
    IDENTIFIER_TO_STYLE_ID,
    resolve_style_identifier,
)
from aspose.words_foss.model.wrap_type import WrapType
from aspose.words_foss.model.enums import NumberStyle, ParagraphAlignment, SectionStart, StyleType
from aspose.words_foss import light_document_model as ldm


class DocFileReader(DocTableBuilderMixin, DocFileReaderCore):
    """Full DOC reader with LDM (Light Document Model) building capability."""

    def to_light_document(self) -> ldm.Document:
        doc = ldm.Document()
        doc.styles = self._build_ldm_styles()
        doc.lists = self._build_ldm_lists()

        body_children = self._build_ldm_body_children()

        if self._ccp_txbx > 0:
            self._inject_textbox_content(body_children)

        if len(self._section_props) > 1 and len(self._section_cps) > 1:
            doc.sections = self._split_into_sections(body_children)
        else:
            sec = ldm.Section()
            sec.page_setup = self._build_ldm_page_setup(0)
            sec.body = ldm.Body(children=body_children)
            doc.sections = [sec]

        hdr, ftr = self._build_ldm_headers_footers()
        doc.header_paragraphs = hdr
        doc.footer_paragraphs = ftr
        return doc

    def _split_into_sections(self, body_children: list) -> list[ldm.Section]:
        """Split body children into sections based on section boundary CPs."""
        sections: list[ldm.Section] = []
        n_sections = len(self._section_props)

        # Build CP→paragraph index mapping from the child paragraphs' CPs.
        child_cp_starts: list[int] = []
        for child in body_children:
            cp = getattr(child, "_cell_cp", getattr(child, "_cp_start", -1))
            child_cp_starts.append(cp)

        sec_boundaries = self._section_cps[1:]  # skip CP[0]=0

        child_idx = 0
        for sec_i in range(n_sections):
            sec = ldm.Section()
            sec.page_setup = self._build_ldm_page_setup(sec_i)
            sec_children: list = []

            if sec_i < len(sec_boundaries):
                boundary_cp = sec_boundaries[sec_i]
            else:
                boundary_cp = self._ccp_text + 1

            while child_idx < len(body_children):
                cp = child_cp_starts[child_idx]
                if cp >= 0 and cp >= boundary_cp:
                    break
                sec_children.append(body_children[child_idx])
                child_idx += 1

            sec.body = ldm.Body(children=sec_children)
            sections.append(sec)

        # Any remaining children go to the last section
        if child_idx < len(body_children) and sections:
            sections[-1].body.children.extend(body_children[child_idx:])

        return sections

    # -- page setup -----------------------------------------------------------

    def _build_ldm_page_setup(self, sec_index: int = 0) -> ldm.PageSetup:
        ps = ldm.PageSetup()
        props = self._section_props[sec_index] if sec_index < len(self._section_props) else {}

        ps.page_width = props.get("page_width", _DEFAULT_PAGE_WIDTH_PT)
        ps.page_height = props.get("page_height", _DEFAULT_PAGE_HEIGHT_PT)
        ps.top_margin = props.get("top_margin", _DEFAULT_MARGIN_PT)
        ps.bottom_margin = props.get("bottom_margin", _DEFAULT_MARGIN_PT)
        ps.left_margin = props.get("left_margin", _DEFAULT_MARGIN_PT)
        ps.right_margin = props.get("right_margin", _DEFAULT_MARGIN_PT)
        ps.header_distance = props.get("header_distance", _DEFAULT_HEADER_FOOTER_DIST_PT)
        ps.footer_distance = props.get("footer_distance", _DEFAULT_HEADER_FOOTER_DIST_PT)
        ps.gutter = props.get("gutter", 0.0)
        ps.different_first_page_header_footer = props.get("different_first_page", False)
        if props.get("restart_page_numbering"):
            ps.restart_page_numbering = True
            default_start = 0 if ps.different_first_page_header_footer else 1
            ps.page_starting_number = props.get("page_starting_number", default_start)

        col_count = props.get("columns_count", 1)
        if col_count > 1:
            tc = ldm.TextColumns()
            tc.count = col_count
            tc.spacing = props.get("columns_spacing", 36.0)
            tc.evenly_spaced = props.get("columns_evenly_spaced", True)
            tc.line_between = props.get("columns_line_between", False)
            ps.text_columns = tc

        # Section break type (sprmSBkc): 0=continuous, 1=newColumn,
        # 2=newPage, 3=evenPage, 4=oddPage
        sbt = props.get("section_break_type", -1)
        if sbt == 0:
            ps.section_start = SectionStart.CONTINUOUS
        elif sbt == 1:
            ps.section_start = SectionStart.NEW_COLUMN
        elif sbt == 2:
            ps.section_start = SectionStart.NEW_PAGE
        elif sbt == 3:
            ps.section_start = SectionStart.EVEN_PAGE
        elif sbt == 4:
            ps.section_start = SectionStart.ODD_PAGE

        w, h = min(ps.page_width, ps.page_height), max(ps.page_width, ps.page_height)
        if abs(w - _LETTER_WIDTH_PT) < 2 and abs(h - _LETTER_HEIGHT_PT) < 2:
            ps.paper_size = 1
        elif abs(w - _A4_WIDTH_PT) < 3 and abs(h - _A4_HEIGHT_PT) < 3:
            ps.paper_size = 9
        return ps

    # -- styles & lists -------------------------------------------------------

    def _build_ldm_styles(self) -> list[ldm.Style]:
        styles: list[ldm.Style] = []
        for istd, name in self._styles.items():
            s = ldm.Style()
            s.name = name
            s.type = StyleType.PARAGRAPH
            sd = self._style_data.get(istd)
            # ``sti`` matches the OOXML StyleIdentifier enum 1:1; 0x0FFF
            # / 0x0FFE are sentinels with no identifier, fall back to a
            # name-based lookup.
            sid = -1
            if sd is not None and sd.sti not in (0x0FFF, 0x0FFE):
                if sd.sti in IDENTIFIER_TO_STYLE_ID:
                    sid = sd.sti
            if sid < 0:
                sid = resolve_style_identifier(name, name)
            if sid >= 0:
                s.style_identifier = sid
                s.built_in = True
            if sd and 1 <= sd.sti <= 9:
                s.is_heading = True
                pf = ldm.ParagraphFormat()
                pf.style_name = name
                pf.is_heading = True
                pf.outline_level = sd.sti - 1
                if sid >= 0:
                    pf.style_identifier = sid
                s.paragraph_format = pf
            # Resolve the effective rPr font for paragraph-style slots
            # (stk=0/1) — character styles already expose their CHP.
            if sd is not None and sd.stk in (0, 1):
                resolved_cp = self._resolve_char_props(istd)
                font = self._build_ldm_font(resolved_cp)
                if (font.name or font.size or font.bold or font.italic
                        or font.underline or font.color):
                    s.font = font
            styles.append(s)
        return styles

    # LVLF.nfc → LDM NumberStyle (matches the OOXML w:numFmt enum).
    _NFC_TO_NUMBER_STYLE = {
        0: NumberStyle.ARABIC,
        1: NumberStyle.UPPERCASE_ROMAN,
        2: NumberStyle.LOWERCASE_ROMAN,
        3: NumberStyle.UPPERCASE_LETTER,
        4: NumberStyle.LOWERCASE_LETTER,
        5: NumberStyle.ORDINAL,
        22: NumberStyle.LEADING_ZERO,
        23: NumberStyle.BULLET,
        255: NumberStyle.NONE,
    }

    # LVLF.ixchFollow → OOXML <w:suff>: 0=tab, 1=space, 2=nothing.
    _FOLLOW_TO_TRAILING = {0: 0, 1: 1, 2: 2}

    def _build_ldm_lists(self) -> list[ldm.DocList]:
        lists: list[ldm.DocList] = []
        for lsid, list_def in self._list_defs.items():
            dl = ldm.DocList()
            dl.list_id = lsid
            dl.is_multi_level = not list_def.is_simple and len(list_def.levels) > 1
            if list_def.levels:
                dl.levels = [self._level_data_to_ldm(lvl) for lvl in list_def.levels]
            else:
                ll = ldm.ListLevel()
                ll.number_style = (
                    NumberStyle.BULLET if list_def.is_hybrid else NumberStyle.ARABIC
                )
                if ll.number_style == NumberStyle.ARABIC:
                    ll.number_format = "%1."
                dl.levels = [ll]
            lists.append(dl)
        return lists

    def _level_data_to_ldm(self, lvl) -> ldm.ListLevel:
        """Convert a binary ``ListLevelData`` into an LDM ``ListLevel``."""
        ll = ldm.ListLevel()
        ll.number_style = self._NFC_TO_NUMBER_STYLE.get(lvl.nfc, NumberStyle.ARABIC)
        ll.number_format = lvl.lvl_text
        ll.start_at = max(1, lvl.start_at)
        ll.alignment = lvl.alignment
        ll.text_position = lvl.left_indent_pt
        ll.number_position = lvl.left_indent_pt + lvl.first_line_indent_pt
        ll.trailing_character = self._FOLLOW_TO_TRAILING.get(lvl.follow, 0)

        if lvl.font_index >= 0 or lvl.font_size_pt > 0 or lvl.bold or lvl.italic:
            f = ldm.Font()
            if lvl.font_index >= 0 and lvl.font_index in self._fonts:
                f.name = self._fonts[lvl.font_index]
            if lvl.font_size_pt > 0:
                f.size = lvl.font_size_pt
            f.bold = lvl.bold
            f.italic = lvl.italic
            ll.font = f
        return ll

    def _resolve_list_level(self, ilfo: int, ilvl: int):
        """Look up the binary ``ListLevelData`` for an (ilfo, ilvl) pair.

        Returns ``None`` if the LFO doesn't resolve to a known list,
        the list has no parsed levels, or the requested level is out
        of range — the caller then falls back to Word's defaults.
        """
        lsid = self._lfo_map.get(ilfo)
        if lsid is None:
            return None
        list_def = self._list_defs.get(lsid)
        if list_def is None or not list_def.levels:
            return None
        # Simple (single-level) lists collapse every ilvl onto the
        # single LVL block.
        if list_def.is_simple:
            return list_def.levels[0]
        if 0 <= ilvl < len(list_def.levels):
            return list_def.levels[ilvl]
        return None

    # -- coordinate helpers ----------------------------------------------------

    def _spa_to_page_mm(self, anchor: ShapeAnchor) -> tuple[float, float]:
        """Convert SPA anchor left/top to absolute page position in mm.

        ``bx`` / ``by`` = 1 means the coordinate is already page-absolute;
        anything else is treated as margin-relative.
        """
        props = self._section_props[0] if self._section_props else {}
        ml = props.get("left_margin", _DEFAULT_MARGIN_PT)
        mt = props.get("top_margin", _DEFAULT_MARGIN_PT)
        x_pt = anchor.left if anchor.bx == 1 else ml + anchor.left
        y_pt = anchor.top if anchor.by == 1 else mt + anchor.top
        return x_pt * _PT_TO_MM, y_pt * _PT_TO_MM

    # -- image shapes ---------------------------------------------------------

    # SPA wr → LDM WrapType mapping.
    _SPA_WR_MAP = {
        0: WrapType.TIGHT,  # around both sides
        1: WrapType.TOP_BOTTOM,  # no wrapping (text above/below)
        2: WrapType.SQUARE,  # left side only
        3: WrapType.NONE,  # closest side
        4: WrapType.SQUARE,  # right side only
        5: WrapType.THROUGH,  # through
    }

    def _build_image_shape(
        self, anchor: ShapeAnchor, positioned: bool = False
    ) -> Optional[ldm.ShapeNode]:
        pib = self._shape_blip_map.get(anchor.spid)
        if not pib or pib < 1 or pib > len(self._blips):
            return None
        blip = self._blips[pib - 1]
        img_bytes = self._wd_bytes[blip.img_offset : blip.img_offset + blip.img_size]
        # BLIPSTORE can hold WMF/EMF/PICT/placeholder bytes the writers
        # can't decode — drop anything that isn't a real raster.
        if not is_raster_blip(blip.blip_type, img_bytes):
            return None

        content_type = "image/png"
        if blip.blip_type in (ESCHER_BLIP_JPEG, ESCHER_BLIP_JPEG2):
            content_type = "image/jpeg"

        shape = ldm.ShapeNode()
        shape.has_image = True
        shape.width = anchor.width
        shape.height = anchor.height
        shape.left = anchor.left
        shape.top = anchor.top
        shape.image_data = ldm.ImageData(content_type=content_type, image_bytes=img_bytes)
        if positioned:
            shape._is_positioned = True
            shape.is_inline = False
            shape.wrap_type = WrapType.NONE
        else:
            shape.is_inline = False
            shape.wrap_type = self._SPA_WR_MAP.get(anchor.wr, WrapType.SQUARE)
        return shape

    # -- inline pictures (fSpec + PicLocation in Data stream) -----------------

    def _parse_inline_picture(self, pic_location: int) -> Optional[ldm.ShapeNode]:
        data = self._data_stream_bytes
        if not data or pic_location < 0:
            return None
        offset = pic_location
        if offset + 68 > len(data):
            return None

        lcb = struct.unpack_from("<I", data, offset)[0]
        cb_header = struct.unpack_from("<H", data, offset + 4)[0]
        if cb_header < 68 or offset + lcb > len(data):
            return None

        pic_data_start = offset + cb_header
        pic_data_end = offset + lcb

        # Prefer the blip embedded directly in the PICF data — it carries
        # the actual inline image.  Fall back to the BSE blip table only
        # when no embedded blip exists (some DOC files store a pib
        # reference without an inline copy).
        img_bytes = self._find_blip_in_picf(data, pic_data_start, pic_data_end)
        if not img_bytes:
            pib = self._find_pib_in_picf(data, pic_data_start, pic_data_end)
            if pib is not None and 1 <= pib <= len(self._blips):
                blip = self._blips[pib - 1]
                img_bytes = self._wd_bytes[blip.img_offset:blip.img_offset + blip.img_size]

        # ``blip_type=0`` here means "unknown / signature only" — let
        # the magic-byte sniff inside ``is_raster_blip`` decide.  PICF
        # blocks for WMF / EMF inline pictures land here too, and the
        # writers can't draw them.
        if not is_raster_blip(0, img_bytes):
            return None

        content_type = "image/png"
        if img_bytes[:2] == b"\xff\xd8":
            content_type = "image/jpeg"

        mfp_mm = struct.unpack_from("<h", data, offset + 6)[0]
        dxa_goal = struct.unpack_from("<H", data, offset + 14)[0]
        dya_goal = struct.unpack_from("<H", data, offset + 16)[0]
        mx = struct.unpack_from("<H", data, offset + 18)[0]
        my = struct.unpack_from("<H", data, offset + 20)[0]

        if mfp_mm == 100:
            dim_w = struct.unpack_from("<H", data, offset + 28)[0]
            dim_h = struct.unpack_from("<H", data, offset + 30)[0]
            if dim_w > 0 and dim_h > 0:
                width_pt = dim_w / 20.0
                height_pt = dim_h / 20.0
            else:
                width_pt = self._get_image_width_pt(img_bytes)
                height_pt = self._get_image_height_pt(img_bytes)
        elif dxa_goal > 0 and mx > 0:
            width_pt = dxa_goal * mx / 1000.0 / 20.0
            height_pt = (dya_goal * my / 1000.0 / 20.0) if (dya_goal > 0 and my > 0) else self._get_image_height_pt(img_bytes)
        else:
            width_pt = self._get_image_width_pt(img_bytes)
            height_pt = self._get_image_height_pt(img_bytes)

        shape = ldm.ShapeNode()
        shape.has_image = True
        shape.is_inline = True
        shape.width = width_pt
        shape.height = height_pt
        shape.image_data = ldm.ImageData(content_type=content_type, image_bytes=img_bytes)
        return shape

    @staticmethod
    def _find_pib_in_picf(data: bytes, start: int, end: int) -> Optional[int]:
        pos = start
        while pos + 8 <= end:
            ver_inst = struct.unpack_from("<H", data, pos)[0]
            fbt = struct.unpack_from("<H", data, pos + 2)[0]
            rec_len = struct.unpack_from("<I", data, pos + 4)[0]
            ver = ver_inst & 0xF

            if ver == 0xF:
                result = DocFileReader._find_pib_in_picf(data, pos + 8, pos + 8 + rec_len)
                if result is not None:
                    return result
                pos += 8 + rec_len
                continue

            if fbt == 0xF00B:
                num_props = (ver_inst >> 4) & 0xFFF
                prop_offset = pos + 8
                for _ in range(num_props):
                    if prop_offset + 6 > end:
                        break
                    pid_flags = struct.unpack_from("<H", data, prop_offset)[0]
                    pid = pid_flags & 0x3FFF
                    val = struct.unpack_from("<I", data, prop_offset + 2)[0]
                    if pid == 0x0104:
                        return val
                    prop_offset += 6

            pos += 8 + rec_len
        return None

    @staticmethod
    def _find_blip_in_picf(data: bytes, start: int, end: int) -> Optional[bytes]:
        pos = start
        while pos + 8 <= end:
            ver_inst = struct.unpack_from("<H", data, pos)[0]
            fbt = struct.unpack_from("<H", data, pos + 2)[0]
            rec_len = struct.unpack_from("<I", data, pos + 4)[0]
            ver = ver_inst & 0xF

            if ver == 0xF:
                result = DocFileReader._find_blip_in_picf(data, pos + 8, pos + 8 + rec_len)
                if result is not None:
                    return result
                pos += 8 + rec_len
                continue

            if fbt == 0xF007:
                bse_body_start = pos + 8 + 36
                if bse_body_start + 8 <= end:
                    return DocFileReader._find_blip_in_picf(data, bse_body_start, pos + 8 + rec_len)
                pos += 8 + rec_len
                continue

            if 0xF01D <= fbt <= 0xF021:
                inst = (ver_inst >> 4) & 0xFFF
                uid_size = 32 if (inst & 1) else 16
                tag_size = 1
                img_start = pos + 8 + uid_size + tag_size
                img_end = pos + 8 + rec_len
                if img_start < img_end:
                    return data[img_start:img_end]

            pos += 8 + rec_len
        return None

    @staticmethod
    def _get_image_width_pt(img_bytes: bytes) -> float:
        if len(img_bytes) > 24 and img_bytes[:4] == b"\x89PNG":
            w_px = struct.unpack_from(">I", img_bytes, 16)[0]
            return w_px * 72.0 / 96.0
        if len(img_bytes) > 2 and img_bytes[:2] == b"\xff\xd8":
            return DocFileReader._jpeg_dimension(img_bytes, width=True)
        return 200.0

    @staticmethod
    def _get_image_height_pt(img_bytes: bytes) -> float:
        if len(img_bytes) > 24 and img_bytes[:4] == b"\x89PNG":
            h_px = struct.unpack_from(">I", img_bytes, 20)[0]
            return h_px * 72.0 / 96.0
        if len(img_bytes) > 2 and img_bytes[:2] == b"\xff\xd8":
            return DocFileReader._jpeg_dimension(img_bytes, width=False)
        return 150.0

    @staticmethod
    def _jpeg_dimension(img_bytes: bytes, width: bool) -> float:
        pos = 2
        while pos + 4 < len(img_bytes):
            if img_bytes[pos] != 0xFF:
                break
            marker = img_bytes[pos + 1]
            if marker in (0xC0, 0xC1, 0xC2):
                if pos + 9 <= len(img_bytes):
                    h = struct.unpack_from(">H", img_bytes, pos + 5)[0]
                    w = struct.unpack_from(">H", img_bytes, pos + 7)[0]
                    val = w if width else h
                    return val * 72.0 / 96.0
                break
            length = struct.unpack_from(">H", img_bytes, pos + 2)[0]
            if length < 2:
                break
            pos += 2 + length
        return 200.0 if width else 150.0

    # -- body children --------------------------------------------------------

    def _build_ldm_body_children(self) -> list:
        children: list = []
        text = self._text
        if not text:
            return children

        body_end = self._ccp_text if self._ccp_text > 0 else len(text)

        self._cp_to_image_anchor: dict[int, ShapeAnchor] = {}
        for anchor in self._body_shape_anchors:
            if anchor.spid in self._shape_blip_map:
                self._cp_to_image_anchor[anchor.cp] = anchor

        paragraphs: list[tuple[int, int]] = []
        start = 0
        for i, ch in enumerate(text[:body_end]):
            if ch == "\r":
                paragraphs.append((start, i))
                start = i + 1

        # Pre-scan for multi-paragraph table cell regions.
        self._cell_para_starts = self._pre_scan_cell_cps(paragraphs, text)

        def _make_page_break() -> ldm.Paragraph:
            pb = ldm.Paragraph()
            run = ldm.Run()
            run.text = "\f"
            pb._children = [run]
            pb.text = "\f"
            return pb

        # Cover-page mode: only when the document has textbox content,
        # indicating a structured cover page with positioned shapes.
        first_page = self._ccp_txbx > 0

        # Build set of section boundary CPs for quick lookup.
        sec_boundary_set = set(self._section_cps[1:]) if len(self._section_cps) > 1 else set()

        for p_start, p_end in paragraphs:
            para_text = text[p_start:p_end]
            has_page_break = "\x0c" in para_text
            para_text_clean = para_text.replace("\x0c", "")

            if not clean_control_chars(para_text_clean).strip():
                # Empty paragraphs may still carry image anchors
                # or inline pictures (fSpec 0x01 characters).
                img_shapes = []
                for pos in range(p_start, p_end):
                    anchor = self._cp_to_image_anchor.get(pos)
                    if anchor:
                        shape = self._build_image_shape(anchor, positioned=first_page)
                        if shape:
                            shape.left, shape.top = self._spa_to_page_mm(anchor)
                            if first_page:
                                shape.width = anchor.width * _PT_TO_MM
                                shape.height = anchor.height * _PT_TO_MM
                            img_shapes.append(shape)
                    elif text[pos] == "\x01":
                        char_ranges = self._get_char_props_in_range(pos, pos + 1)
                        for _cs, _ce, cp in char_ranges:
                            if cp.is_special and cp.pic_location >= 0:
                                shape = self._parse_inline_picture(cp.pic_location)
                                if shape:
                                    img_shapes.append(shape)

                if has_page_break:
                    pb = _make_page_break()
                    pb._cp_start = p_start  # type: ignore[attr-defined]
                    pb._children.extend(img_shapes)
                    children.append(pb)
                    first_page = False
                elif img_shapes:
                    holder = ldm.Paragraph()
                    holder._cp_start = p_start  # type: ignore[attr-defined]
                    holder._children.extend(img_shapes)
                    children.append(holder)
                else:
                    raw = text[p_start:p_end]
                    is_structural = (
                        "\x07" in raw
                        or p_start in self._cell_para_starts
                        or p_end in self._section_cps
                    )
                    if not is_structural:
                        empty_para = self._build_empty_paragraph(p_start)
                        children.append(empty_para)
                continue

            # When a page break starts a paragraph that crosses a
            # section boundary, emit the break first so the content
            # paragraph receives a _cp_start inside the next section.
            pb_at_start = has_page_break and para_text.startswith("\x0c")
            content_cp = p_start
            if pb_at_start:
                at_sec_boundary = (p_start + 1) in sec_boundary_set
                if at_sec_boundary:
                    sep = self._build_empty_paragraph(p_start)
                    children.append(sep)
                else:
                    pb = _make_page_break()
                    pb._cp_start = p_start  # type: ignore[attr-defined]
                    children.append(pb)
                first_page = False
                content_cp = p_start + 1

            if "\x07" in para_text_clean:
                tbl, trailing = self._build_ldm_table_from_text(
                    para_text, p_start, p_end
                )
                tbl._cp_start = content_cp  # type: ignore[attr-defined]
                if tbl.rows:
                    self._absorb_preceding_into_table(tbl, children)
                    children.append(tbl)
                    # DOCX requires a paragraph after a table; only emit a
                    # stub when no trailing text already serves as one.
                    if not trailing.strip():
                        post_tbl = self._build_empty_paragraph(p_end - 1)
                        children.append(post_tbl)
                if trailing.strip():
                    trail_start = p_start + len(para_text) - len(trailing)
                    if "\x13" in trailing and "\x14" in trailing:
                        para = self._build_ldm_hyperlink_paragraph(trailing, trail_start, p_end)
                    else:
                        para = self._build_ldm_paragraph(trailing, trail_start, p_end)
                    para._cp_start = trail_start  # type: ignore[attr-defined]
                    children.append(para)
                if has_page_break and not pb_at_start:
                    pb = _make_page_break()
                    pb._cp_start = p_start  # type: ignore[attr-defined]
                    children.append(pb)
                    first_page = False
                continue

            if "\x13" in para_text_clean and "\x14" in para_text_clean:
                para = self._build_ldm_hyperlink_paragraph(para_text_clean, p_start, p_end)
            else:
                para = self._build_ldm_paragraph(para_text_clean, p_start, p_end)

            # Tag with CP so _absorb_preceding_into_table and section split can work.
            para._cell_cp = content_cp  # type: ignore[attr-defined]
            para._cp_start = content_cp  # type: ignore[attr-defined]

            for pos in range(p_start, p_end):
                anchor = self._cp_to_image_anchor.get(pos)
                if anchor:
                    # Cover-page images (before the first page break)
                    # must be placed at their absolute SPA coordinates.
                    shape = self._build_image_shape(anchor, positioned=first_page)
                    if shape:
                        # Convert SPA coords to absolute page mm.
                        shape.left, shape.top = self._spa_to_page_mm(anchor)
                        if first_page:
                            shape.width = anchor.width * _PT_TO_MM
                            shape.height = anchor.height * _PT_TO_MM
                        para._children.append(shape)
                elif text[pos] == "\x01":
                    char_ranges = self._get_char_props_in_range(pos, pos + 1)
                    for _cs, _ce, cp in char_ranges:
                        if cp.is_special and cp.pic_location >= 0:
                            shape = self._parse_inline_picture(cp.pic_location)
                            if shape:
                                para._children.append(shape)

            children.append(para)
            if has_page_break and not pb_at_start:
                pb = _make_page_break()
                pb._cp_start = p_start  # type: ignore[attr-defined]
                children.append(pb)
                first_page = False

        return children

    # -- paragraph format -----------------------------------------------------

    # Paragraph fields that follow the simple "use direct, else style" rule.
    _PARA_FORMAT_FIELDS = (
        "alignment",
        "left_indent",
        "right_indent",
        "first_line_indent",
        "space_before",
        "space_after",
        "space_before_auto",
        "space_after_auto",
        "keep_with_next",
        "keep_together",
        "widow_control",
        "suppress_auto_hyphens",
        "suppress_line_numbers",
        "add_space_between_far_east_and_alpha",
        "add_space_between_far_east_and_digit",
        "auto_adjust_right_indent",
        "no_space_between_paragraphs_of_same_style",
        "page_break_before",
        "baseline_alignment",
        "conditional_style",
    )

    def _resolve_ldm_paragraph_format(self, props: ParaProps) -> ldm.ParagraphFormat:
        style_name = self._styles.get(props.istd, "Normal")
        style_pp = self._resolve_para_props(props.istd)
        pf = ldm.ParagraphFormat()
        pf.style_name = style_name

        for field in self._PARA_FORMAT_FIELDS:
            if field in props._set_fields:
                setattr(pf, field, getattr(props, field))
            else:
                setattr(pf, field, getattr(style_pp, field))

        if "line_spacing" in props._set_fields:
            pf.line_spacing = props.line_spacing
            pf.line_spacing_rule = props.line_spacing_rule
        else:
            pf.line_spacing = style_pp.line_spacing
            pf.line_spacing_rule = style_pp.line_spacing_rule

        if "outline_level" in props._set_fields:
            pf.outline_level = props.outline_level
        else:
            pf.outline_level = style_pp.outline_level

        sd = self._style_data.get(props.istd)
        if sd and 1 <= sd.sti <= 9:
            pf.is_heading = True
            pf.outline_level = sd.sti - 1
        elif pf.outline_level < 9:
            pf.is_heading = True

        if props.ilfo > 0:
            pf.is_list_item = True
            # Bullet/number paragraphs inherit indent from the list
            # level's npos / tpos unless the PAPX overrides them.
            li_from_papx = "left_indent" in props._set_fields
            fli_from_papx = "first_line_indent" in props._set_fields
            level = self._resolve_list_level(props.ilfo, props.ilvl)
            if level is not None and (level.left_indent_pt or level.first_line_indent_pt):
                if not li_from_papx:
                    pf.left_indent = level.left_indent_pt
                if not fli_from_papx:
                    pf.first_line_indent = level.first_line_indent_pt
            else:
                if not li_from_papx:
                    pf.left_indent = (props.ilvl + 1) * _LIST_DEFAULT_LEVEL_INDENT_PT
                if not fli_from_papx:
                    pf.first_line_indent = -_LIST_DEFAULT_HANGING_PT

        raw_tabs = props.tab_stops if props.tab_stops else style_pp.tab_stops
        if raw_tabs:
            tc = ldm.TabStopCollection()
            for pos_pt, align, leader in raw_tabs:
                tc.tab_stops.append(ldm.TabStop(
                    position=pos_pt, alignment=align, leader=leader,
                ))
            pf.tab_stops = tc

        # Paragraph borders — merge direct over style on a per-side basis.
        merged_borders = list(style_pp.borders)
        for i, b in enumerate(props.borders):
            if b is not None:
                merged_borders[i] = b
        if any(b is not None for b in merged_borders):
            pf.borders = [
                ldm.Border(
                    line_style=b[0],
                    line_width=b[1],
                    color=b[2],
                ) if b is not None else ldm.Border()
                for b in merged_borders
            ]

        shading_back = (
            props.shading_back
            if "shading_back" in props._set_fields
            else style_pp.shading_back
        )
        if shading_back:
            pf.shading = ldm.Shading(background_pattern_color=shading_back)

        return pf

    # -- paragraph & run builders ---------------------------------------------

    def _build_empty_paragraph(self, cp: int) -> ldm.Paragraph:
        """Build an empty paragraph whose format is resolved through the style chain at *cp*."""
        para = ldm.Paragraph()
        para._cp_start = cp  # type: ignore[attr-defined]
        props = self._get_para_props_at(cp)
        para.paragraph_format = self._resolve_ldm_paragraph_format(props)
        para.text = ""
        return para

    def _build_ldm_paragraph(self, para_text: str, p_start: int, p_end: int) -> ldm.Paragraph:
        para = ldm.Paragraph()
        props = self._get_para_props_at(p_start)
        para.paragraph_format = self._resolve_ldm_paragraph_format(props)

        if props.ilfo > 0:
            lf = ldm.ListFormat()
            lf.is_list_item = True
            lf.list_level_number = props.ilvl
            lsid = self._lfo_map.get(props.ilfo)
            lf.list_id = lsid if lsid is not None else props.ilfo
            para.list_format = lf

        style_cp = self._resolve_char_props(props.istd)
        char_ranges = self._get_char_props_in_range(p_start, p_end)
        text_parts: list[str] = []

        if char_ranges:
            for cs, ce, cp in char_ranges:
                run_text = clean_control_chars(self._text[cs:ce]).replace("\x0c", "")
                if not run_text:
                    continue
                merged = self._merge_char_props(style_cp, cp, cp._set_fields)
                run = ldm.Run()
                run.text = run_text
                run.font = self._build_ldm_font(merged)
                para._children.append(run)
                text_parts.append(run_text)
        else:
            run = ldm.Run()
            run.text = clean_control_chars(para_text)
            run.font = self._build_ldm_font(style_cp)
            para._children.append(run)
            text_parts.append(run.text)

        para.text = "".join(text_parts)
        return para

    def _build_ldm_font(self, cp: CharProps) -> ldm.Font:
        font = ldm.Font()
        font.bold = cp.bold
        font.italic = cp.italic
        font.strike_through = cp.strikethrough
        font.underline = cp.underline
        font.superscript = cp.superscript
        font.subscript = cp.subscript
        font.all_caps = cp.all_caps
        font.small_caps = cp.small_caps
        font.hidden = cp.hidden
        font.emboss = cp.emboss
        font.engrave = cp.engrave
        font.outline = cp.outline
        font.shadow = cp.shadow
        font.kerning = cp.kerning

        if cp.font_index >= 0 and cp.font_index in self._fonts:
            font.name = self._fonts[cp.font_index]
        elif self._default_font_index in self._fonts:
            font.name = self._fonts[self._default_font_index]
        font.size = cp.font_size if cp.font_size > 0 else _DEFAULT_FONT_SIZE_PT
        font.color = cp.color if cp.color else "Color [Empty]"
        font.highlight_color = cp.highlight_color if cp.highlight_color else "Color [Empty]"

        if cp.style_index >= 0 and cp.style_index in self._styles:
            font.style_name = self._styles[cp.style_index]
            sid = resolve_style_identifier(font.style_name, font.style_name)
            if sid >= 0:
                font.style_identifier = sid
        else:
            font.style_name = "Default Paragraph Font"
        return font

    # -- hyperlink paragraph --------------------------------------------------

    def _build_ldm_hyperlink_paragraph(
        self, para_text: str, p_start: int, p_end: int
    ) -> ldm.Paragraph:
        para = ldm.Paragraph()
        props = self._get_para_props_at(p_start)
        para.paragraph_format = self._resolve_ldm_paragraph_format(props)

        style_cp = self._resolve_char_props(props.istd)
        char_ranges = self._get_char_props_in_range(p_start, p_end)
        if char_ranges:
            first_cp = char_ranges[0][2]
            style_cp = self._merge_char_props(style_cp, first_cp, first_cp._set_fields)

        clean_text = clean_control_chars(evaluate_fields(para_text))
        if clean_text:
            run = ldm.Run()
            run.text = clean_text
            run.font = self._build_ldm_font(style_cp)
            para._children.append(run)
        para.text = clean_text
        return para

    # -- textbox injection (cover page) ---------------------------------------

    def _inject_textbox_content(self, body_children: list) -> None:
        """Inject textbox and fill-shape content using structural data.

        Uses FTXBXS (PlcftxbxTxt) to map shape IDs to textbox text
        ranges, and Escher ChildAnchor/Spgr records for child shape
        positions.  No font-size or position-ratio heuristics.
        """
        txbx_start = self._ccp_text + self._ccp_ftn + self._ccp_hdd
        txbx_end = txbx_start + self._ccp_txbx
        if txbx_end > len(self._text):
            return

        # SPA lookup by shape ID
        spa_by_spid = {a.spid: a for a in self._body_shape_anchors}

        # Page dimensions
        sec = self._section_props[0] if self._section_props else {}
        page_w_mm = sec.get("page_width", _DEFAULT_PAGE_WIDTH_PT) * _PT_TO_MM

        # Find page-break CP for cover page boundary
        pb_cp = -1
        for i, ch in enumerate(self._text[: self._ccp_text]):
            if ch == "\x0c":
                pb_cp = i
                break

        # Find attachment paragraphs
        first_para = None
        pb_para = None
        for child in body_children:
            if child.type != "Paragraph":
                continue
            if first_para is None:
                first_para = child
            if any("\f" in (r.text or "") for r in child.runs):
                pb_para = child
                break

        attach_para = first_para
        if attach_para is None:
            return

        # Build a unified shape→(cp_start, cp_end) mapping from
        # PlcftxbxTxt if available, otherwise from FOpt txid.
        story_map = dict(self._textbox_stories)
        if not story_map and self._shape_txid_map:
            story_map = self._build_stories_from_txid(txbx_start, txbx_end)

        # --- Phase 1: Fill shapes using Escher ChildAnchor positions ---
        # Emitted first so textbox content renders on top.
        fill_shapes = [
            (spid, props)
            for spid, props in self._shape_line_map.items()
            if props.fill_color_rgb and spid not in spa_by_spid
        ]
        handled_spids: set[int] = set()
        if fill_shapes:
            handled_spids = self._inject_fill_shapes_from_escher(
                fill_shapes,
                spa_by_spid,
                txbx_start,
                page_w_mm,
                attach_para,
                story_map,
            )

        # --- Phase 2: Textbox shapes (skip those handled as fill shapes) ---
        for shape_id, (cp_start, cp_end) in story_map.items():
            if shape_id in handled_spids:
                continue
            # Position: use SPA if the shape has an entry, otherwise
            # use ChildAnchor + group transform for group children.
            anchor = spa_by_spid.get(shape_id)

            if anchor:
                # Only process cover-page shapes
                if pb_cp >= 0 and anchor.cp > pb_cp:
                    continue
                # Skip full-page background shapes
                w_mm = anchor.width * _PT_TO_MM
                h_mm = anchor.height * _PT_TO_MM
                if w_mm > page_w_mm - 5 and h_mm > page_w_mm:
                    continue
                left_mm, top_mm = self._spa_to_page_mm(anchor)
            else:
                # Child shape — compute position from ChildAnchor
                child_anchor = self._child_anchors.get(shape_id)
                parent_spid = self._child_to_parent.get(shape_id)
                if not child_anchor or not parent_spid:
                    continue
                group = self._group_info.get(parent_spid)
                parent_spa = spa_by_spid.get(parent_spid)
                if not group or not parent_spa:
                    continue
                grp_w = group.coord_right - group.coord_left
                grp_h = group.coord_bottom - group.coord_top
                if grp_w <= 0 or grp_h <= 0:
                    continue
                abs_l, abs_t = self._spa_to_page_mm(parent_spa)
                pw = parent_spa.width * _PT_TO_MM
                ph = parent_spa.height * _PT_TO_MM
                left_mm = abs_l + (child_anchor.left - group.coord_left) / grp_w * pw
                top_mm = abs_t + (child_anchor.top - group.coord_top) / grp_h * ph
                w_mm = (child_anchor.right - child_anchor.left) / grp_w * pw
                h_mm = (child_anchor.bottom - child_anchor.top) / grp_h * ph

            abs_start = txbx_start + cp_start
            abs_end = txbx_start + cp_end
            if abs_end > len(self._text):
                continue

            tb_paras = self._build_textbox_paragraphs(abs_start, abs_end)
            if not tb_paras:
                continue

            shape = ldm.ShapeNode()
            shape.has_image = False
            shape.is_inline = False
            shape._is_positioned = True
            shape.wrap_type = WrapType.SQUARE
            shape.left = left_mm
            shape.top = top_mm
            shape.width = w_mm
            shape.height = h_mm
            shape.text_box = {"paragraphs": tb_paras}

            target = pb_para or attach_para
            target._children.append(shape)

    def _build_stories_from_txid(
        self, txbx_start: int, txbx_end: int
    ) -> dict[int, tuple[int, int]]:
        """Build shape→(cp_start, cp_end) mapping using FOpt txid.

        When PlcftxbxTxt is absent, split the textbox text stream into
        stories by finding content blocks separated by empty paragraphs.
        Each content block maps to a story index from ``_shape_txid_map``.
        """
        text = self._text[txbx_start:txbx_end]
        if not text:
            return {}

        # Split into paragraphs and find content block boundaries.
        # A content block is a maximal run of paragraphs containing
        # at least one non-empty paragraph, separated from other blocks
        # by one or more empty paragraphs.
        paragraphs: list[tuple[int, int, str]] = []  # (cp_start, cp_end, text)
        pos = 0
        for line in text.split("\r"):
            cp_end = pos + len(line)
            paragraphs.append((pos, cp_end, clean_control_chars(line).strip()))
            pos = cp_end + 1  # +1 for the \r

        # Identify content blocks
        blocks: list[tuple[int, int]] = []  # (cp_start, cp_end) for each block
        block_start = -1
        block_end = -1
        for cp_start, cp_end, txt in paragraphs:
            if txt:
                if block_start < 0:
                    block_start = cp_start
                block_end = cp_end
            else:
                if block_start >= 0:
                    blocks.append((block_start, block_end))
                    block_start = -1
                    block_end = -1
        if block_start >= 0:
            blocks.append((block_start, block_end))

        # Map blocks to shapes via txid story indices
        # Sort shapes by story index to align with block order
        sorted_shapes = sorted(self._shape_txid_map.items(), key=lambda x: x[1])

        result: dict[int, tuple[int, int]] = {}
        for i, (spid, _story_idx) in enumerate(sorted_shapes):
            if i < len(blocks):
                result[spid] = blocks[i]
        return result

    def _build_textbox_paragraphs(self, abs_start: int, abs_end: int) -> list:
        """Build LDM paragraphs from a textbox text range."""
        paras: list[ldm.Paragraph] = []
        line_pos = abs_start
        for line in self._text[abs_start:abs_end].split("\r"):
            cleaned = clean_control_chars(line).strip()
            if not cleaned:
                line_pos += len(line) + 1
                continue
            para = ldm.Paragraph()
            props = self._get_para_props_at(line_pos)
            para.paragraph_format = self._resolve_ldm_paragraph_format(props)
            style_cp = self._resolve_char_props(props.istd)
            chpx = self._get_char_props_in_range(line_pos, line_pos + len(line))
            if chpx:
                first_cp = chpx[0][2]
                style_cp = self._merge_char_props(style_cp, first_cp, first_cp._set_fields)
            run = ldm.Run()
            run.text = cleaned
            run.font = self._build_ldm_font(style_cp)
            para._children = [run]
            para.text = cleaned
            paras.append(para)
            line_pos += len(line) + 1
        return paras

    def _inject_fill_shapes_from_escher(
        self,
        fill_shapes: list,
        spa_by_spid: dict,
        txbx_start: int,
        page_w_mm: float,
        attach_para: object,
        story_map: dict[int, tuple[int, int]],
    ) -> set[int]:
        """Inject fill shapes using ChildAnchor + Spgr coordinate transform.

        Each fill shape is a child of a group shape.  Its page position
        is computed by mapping ChildAnchor coordinates through the parent
        group's Spgr coordinate system and SPA page extent.

        Returns set of shape IDs handled (so Phase 2 can skip them).
        """
        handled: set[int] = set()

        for spid, props in fill_shapes:
            r, g, b = props.fill_color_rgb
            color_str = f"Color [A=255, R={r}, G={g}, B={b}]"

            # Compute position from ChildAnchor + parent group transform
            child_anchor = self._child_anchors.get(spid)
            parent_spid = self._child_to_parent.get(spid)
            if not child_anchor or not parent_spid:
                continue

            group = self._group_info.get(parent_spid)
            parent_spa = spa_by_spid.get(parent_spid)
            if not group or not parent_spa:
                continue

            grp_w = group.coord_right - group.coord_left
            grp_h = group.coord_bottom - group.coord_top
            if grp_w <= 0 or grp_h <= 0:
                continue

            abs_left, abs_top = self._spa_to_page_mm(parent_spa)
            parent_w_mm = parent_spa.width * _PT_TO_MM
            parent_h_mm = parent_spa.height * _PT_TO_MM

            bar = ldm.ShapeNode()
            bar.is_inline = False
            bar._is_positioned = True
            bar.wrap_type = WrapType.NONE
            bar.shading = ldm.Shading(background_pattern_color=color_str)
            bar.left = abs_left + (child_anchor.left - group.coord_left) / grp_w * parent_w_mm
            bar.top = abs_top + (child_anchor.top - group.coord_top) / grp_h * parent_h_mm
            bar.width = (child_anchor.right - child_anchor.left) / grp_w * parent_w_mm
            bar.height = (child_anchor.bottom - child_anchor.top) / grp_h * parent_h_mm

            # Attach textbox text if this fill shape has a story entry
            if spid in story_map:
                cp_start, cp_end = story_map[spid]
                abs_s = txbx_start + cp_start
                abs_e = txbx_start + cp_end
                if abs_e <= len(self._text):
                    tb_paras = self._build_textbox_paragraphs(abs_s, abs_e)
                    if tb_paras:
                        bar.text_box = {"paragraphs": tb_paras}
                        bar.vertical_alignment = 2  # bottom-anchored
                handled.add(spid)

            attach_para._children.append(bar)

        return handled

    # -- headers & footers ----------------------------------------------------

    def _build_ldm_headers_footers(
        self,
    ) -> tuple[list[ldm.Paragraph], list[ldm.Paragraph]]:
        headers: list[ldm.Paragraph] = []
        footers: list[ldm.Paragraph] = []

        if not self._hdd_cps or self._ccp_hdd == 0:
            return headers, footers

        hdd_base = self._ccp_text + self._ccp_ftn
        text = self._text
        n_stories = len(self._hdd_cps) - 1

        def _extract_story(story_idx: int) -> str:
            if story_idx >= n_stories:
                return ""
            cp_start = self._hdd_cps[story_idx]
            cp_end = self._hdd_cps[story_idx + 1]
            if cp_start >= cp_end:
                return ""
            abs_start = hdd_base + cp_start
            abs_end = hdd_base + cp_end
            if abs_end > len(text):
                return ""
            return text[abs_start:abs_end]

        n_sections = (n_stories - 1) // 6 if n_stories > 1 else 0

        def _has_visible_text(raw: str) -> bool:
            return bool(clean_control_chars(raw).replace("\r", "").strip())

        hdr_cp_start = hdr_cp_end = ftr_cp_start = ftr_cp_end = -1
        for sec_i in range(n_sections):
            base = 1 + sec_i * 6
            for story_offset in (1, 0):
                idx = base + story_offset
                if _has_visible_text(_extract_story(idx)) and hdr_cp_start < 0:
                    hdr_cp_start = hdd_base + self._hdd_cps[idx]
                    hdr_cp_end = hdd_base + self._hdd_cps[idx + 1]
                    break
            for story_offset in (3, 2):
                idx = base + story_offset
                if _has_visible_text(_extract_story(idx)) and ftr_cp_start < 0:
                    ftr_cp_start = hdd_base + self._hdd_cps[idx]
                    ftr_cp_end = hdd_base + self._hdd_cps[idx + 1]
                    break

        for region_start, region_end, target in (
            (hdr_cp_start, hdr_cp_end, headers),
            (ftr_cp_start, ftr_cp_end, footers),
        ):
            if region_start < 0:
                continue
            region_text = text[region_start:region_end]
            if "\x13" in region_text:
                region_text = evaluate_fields(region_text)
            raw_lines = text[region_start:region_end].split("\r")
            eval_lines = region_text.split("\r")
            line_pos = region_start
            for raw_line, eval_line in zip(raw_lines, eval_lines):
                cleaned = clean_control_chars(eval_line).strip()
                if not cleaned:
                    line_pos += len(raw_line) + 1
                    continue
                para = ldm.Paragraph()
                props = self._get_para_props_at(line_pos)
                para.paragraph_format = self._resolve_ldm_paragraph_format(props)
                style_cp = self._resolve_char_props(props.istd)
                chpx = self._get_char_props_in_range(line_pos, line_pos + len(raw_line))
                if chpx:
                    first_cp = chpx[0][2]
                    style_cp = self._merge_char_props(style_cp, first_cp, first_cp._set_fields)
                run = ldm.Run()
                run.text = cleaned
                run.font = self._build_ldm_font(style_cp)
                para._children = [run]
                para.text = cleaned
                target.append(para)
                line_pos += len(raw_line) + 1

        if headers and self._hdr_shape_anchors:
            # SPA coords are relative to the text column / header band
            # — convert to absolute page positions (mm).
            props = self._section_props[0] if self._section_props else {}
            margin_left_pt = props.get("left_margin", 54.0)
            hdr_dist_pt = props.get("header_distance", 36.0)

            for anchor in self._hdr_shape_anchors:
                shape = self._build_image_shape(anchor, positioned=True)
                if shape:
                    max_h = 14.0
                    if anchor.height > max_h and anchor.width > 0:
                        scale = max_h / anchor.height
                        shape.width = anchor.width * scale
                        shape.height = max_h
                    # Convert to absolute page coords (mm)
                    shape.left = (margin_left_pt + anchor.left) * _PT_TO_MM
                    shape.top = (hdr_dist_pt + anchor.top) * _PT_TO_MM
                    headers[0]._children.append(shape)
                    continue

                # Non-image shape: check for line/border properties
                # (e.g. decorative bars under the header text).
                line_props = self._shape_line_map.get(anchor.spid)
                if line_props and line_props.line_color_rgb:
                    r, g, b = line_props.line_color_rgb
                    color_str = f"Color [A=255, R={r}, G={g}, B={b}]"
                    bar = ldm.ShapeNode()
                    bar.has_image = False
                    bar.is_inline = False
                    bar.width = anchor.width * _PT_TO_MM  # pt -> mm
                    bar.height = anchor.height * _PT_TO_MM
                    bar.left = (margin_left_pt + anchor.left) * _PT_TO_MM
                    bar.top = (hdr_dist_pt + anchor.top) * _PT_TO_MM
                    bar.wrap_type = WrapType.NONE
                    bar._is_positioned = True
                    bar.borders = [
                        ldm.Border(
                            line_style=1,
                            line_width=line_props.line_width_pt,
                            color=color_str,
                        )
                    ]
                    # Place in a new empty paragraph, matching the DOCX
                    # layout where the bar lives in its own paragraph.
                    bar_para = ldm.Paragraph()
                    bar_para.paragraph_format = ldm.ParagraphFormat(
                        alignment=ParagraphAlignment.RIGHT
                    )
                    bar_para._children.append(bar)
                    headers.append(bar_para)

        return headers, footers
