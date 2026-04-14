"""
LDM builder for DOC files — extends DocFileReaderCore with to_light_document().
"""

from __future__ import annotations

from typing import Optional, TYPE_CHECKING

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

if TYPE_CHECKING:
    from aspose.words_foss import light_document_model as ldm

from aspose.words_foss.doc_reader.constants import ESCHER_BLIP_JPEG, ESCHER_BLIP_JPEG2
from aspose.words_foss.doc_reader.doc_file_reader_core import DocFileReaderCore
from aspose.words_foss.doc_reader.images import ShapeAnchor
from aspose.words_foss.doc_reader.properties import CharProps, ParaProps
from aspose.words_foss.doc_reader.text import clean_control_chars, evaluate_fields
from aspose.words_foss.model.wrap_type import WrapType
from aspose.words_foss.model.enums import NumberStyle, ParagraphAlignment, StyleType


class DocFileReader(DocFileReaderCore):
    """Full DOC reader with LDM (Light Document Model) building capability."""

    def to_light_document(self) -> ldm.Document:
        from aspose.words_foss import light_document_model as ldm

        doc = ldm.Document()
        doc.styles = self._build_ldm_styles()
        doc.lists = self._build_ldm_lists()

        sec = ldm.Section()
        sec.page_setup = self._build_ldm_page_setup()
        body_children = self._build_ldm_body_children()

        if self._ccp_txbx > 0:
            self._inject_textbox_content(body_children)

        sec.body = ldm.Body(children=body_children)
        doc.sections = [sec]

        hdr, ftr = self._build_ldm_headers_footers()
        doc.header_paragraphs = hdr
        doc.footer_paragraphs = ftr
        return doc

    # -- page setup -----------------------------------------------------------

    def _build_ldm_page_setup(self) -> ldm.PageSetup:
        from aspose.words_foss import light_document_model as ldm

        ps = ldm.PageSetup()
        props = self._section_props[0] if self._section_props else {}

        ps.page_width = props.get("page_width", _DEFAULT_PAGE_WIDTH_PT)
        ps.page_height = props.get("page_height", _DEFAULT_PAGE_HEIGHT_PT)
        ps.top_margin = props.get("top_margin", _DEFAULT_MARGIN_PT)
        ps.bottom_margin = props.get("bottom_margin", _DEFAULT_MARGIN_PT)
        ps.left_margin = props.get("left_margin", _DEFAULT_MARGIN_PT)
        ps.right_margin = props.get("right_margin", _DEFAULT_MARGIN_PT)
        ps.header_distance = props.get("header_distance", _DEFAULT_HEADER_FOOTER_DIST_PT)
        ps.footer_distance = props.get("footer_distance", _DEFAULT_HEADER_FOOTER_DIST_PT)
        ps.different_first_page_header_footer = props.get("different_first_page", False)
        if props.get("restart_page_numbering"):
            ps.restart_page_numbering = True
            default_start = 0 if ps.different_first_page_header_footer else 1
            ps.page_starting_number = props.get("page_starting_number", default_start)

        w, h = min(ps.page_width, ps.page_height), max(ps.page_width, ps.page_height)
        if abs(w - _LETTER_WIDTH_PT) < 2 and abs(h - _LETTER_HEIGHT_PT) < 2:
            ps.paper_size = 1
        elif abs(w - _A4_WIDTH_PT) < 3 and abs(h - _A4_HEIGHT_PT) < 3:
            ps.paper_size = 9
        return ps

    # -- styles & lists -------------------------------------------------------

    def _build_ldm_styles(self) -> list[ldm.Style]:
        from aspose.words_foss import light_document_model as ldm

        styles: list[ldm.Style] = []
        for istd, name in self._styles.items():
            s = ldm.Style()
            s.name = name
            s.type = StyleType.PARAGRAPH
            sd = self._style_data.get(istd)
            if sd and 1 <= sd.sti <= 9:
                s.is_heading = True
                pf = ldm.ParagraphFormat()
                pf.style_name = name
                pf.is_heading = True
                pf.outline_level = sd.sti - 1
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
                ll.number_style = NumberStyle.BULLET
            else:
                ll.number_style = NumberStyle.ARABIC
                ll.number_format = "%1."
            dl.levels = [ll]
            lists.append(dl)
        return lists

    # -- coordinate helpers ----------------------------------------------------

    def _spa_to_page_mm(self, anchor: ShapeAnchor) -> tuple[float, float]:
        """Convert SPA anchor left/top to absolute page position in mm.

        Uses bx/by flags to determine the coordinate reference:
        bx: 0=margin-relative, 1=page-absolute, 2=char-relative (≈margin)
        by: 0=margin-relative, 1=page-absolute, 2=paragraph-relative (≈margin)
        """
        props = self._section_props[0] if self._section_props else {}
        ml = props.get("left_margin", _DEFAULT_MARGIN_PT)
        mt = props.get("top_margin", _DEFAULT_MARGIN_PT)
        # bx=1 → coordinate is already page-absolute
        x_pt = anchor.left if anchor.bx == 1 else ml + anchor.left
        # by=1 → coordinate is already page-absolute
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
        from aspose.words_foss import light_document_model as ldm

        pib = self._shape_blip_map.get(anchor.spid)
        if not pib or pib < 1 or pib > len(self._blips):
            return None
        blip = self._blips[pib - 1]
        img_bytes = self._wd_bytes[blip.img_offset : blip.img_offset + blip.img_size]
        if not img_bytes:
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
            # SPA-anchored images are floating, not inline.
            shape.is_inline = False
            shape.wrap_type = self._SPA_WR_MAP.get(anchor.wr, WrapType.SQUARE)
        return shape

    # -- body children --------------------------------------------------------

    def _build_ldm_body_children(self) -> list:
        from aspose.words_foss import light_document_model as ldm

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
            pb.runs = [run]
            pb.text = "\f"
            return pb

        first_page = True

        for p_start, p_end in paragraphs:
            para_text = text[p_start:p_end]
            has_page_break = "\x0c" in para_text
            para_text_clean = para_text.replace("\x0c", "")

            if not clean_control_chars(para_text_clean).strip():
                # Empty paragraphs may still carry image anchors.
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

                if has_page_break:
                    pb = _make_page_break()
                    pb.inline_extras.extend(img_shapes)
                    children.append(pb)
                    first_page = False
                elif img_shapes:
                    holder = ldm.Paragraph()
                    holder.inline_extras.extend(img_shapes)
                    children.append(holder)
                continue

            if "\x07" in para_text_clean:
                tbl, trailing = self._build_ldm_table_from_text(para_text_clean)
                if tbl.rows:
                    # Multi-paragraph cells: preceding body paragraphs
                    # that were emitted as standalone children may
                    # actually belong inside this table's first cell.
                    # Collect them back up to a structural boundary.
                    self._absorb_preceding_into_table(tbl, children)
                    children.append(tbl)
                if trailing.strip():
                    trail_start = p_start + len(para_text) - len(trailing)
                    if "\x13" in trailing and "\x14" in trailing:
                        para = self._build_ldm_hyperlink_paragraph(trailing, trail_start, p_end)
                    else:
                        para = self._build_ldm_paragraph(trailing, trail_start, p_end)
                    children.append(para)
                if has_page_break:
                    children.append(_make_page_break())
                    first_page = False
                continue

            if "\x13" in para_text_clean and "\x14" in para_text_clean:
                para = self._build_ldm_hyperlink_paragraph(para_text_clean, p_start, p_end)
            else:
                para = self._build_ldm_paragraph(para_text_clean, p_start, p_end)

            # Tag with CP so _absorb_preceding_into_table can check.
            para._cell_cp = p_start  # type: ignore[attr-defined]

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
                        para.inline_extras.append(shape)

            children.append(para)
            if has_page_break:
                children.append(_make_page_break())
                first_page = False

        return children

    # -- paragraph format -----------------------------------------------------

    def _resolve_ldm_paragraph_format(self, props: ParaProps) -> ldm.ParagraphFormat:
        from aspose.words_foss import light_document_model as ldm

        style_name = self._styles.get(props.istd, "Normal")
        style_pp = self._resolve_para_props(props.istd)
        pf = ldm.ParagraphFormat()
        pf.style_name = style_name

        _FIELDS = (
            "alignment",
            "left_indent",
            "right_indent",
            "first_line_indent",
            "space_before",
            "space_after",
            "space_before_auto",
            "space_after_auto",
            "keep_with_next",
            "page_break_before",
        )
        for field in _FIELDS:
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
        return pf

    # -- paragraph & run builders ---------------------------------------------

    def _build_ldm_paragraph(self, para_text: str, p_start: int, p_end: int) -> ldm.Paragraph:
        from aspose.words_foss import light_document_model as ldm

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
                run_text = clean_control_chars(self._text[cs:ce])
                if not run_text:
                    continue
                merged = self._merge_char_props(style_cp, cp, cp._set_fields)
                run = ldm.Run()
                run.text = run_text
                run.font = self._build_ldm_font(merged)
                para.runs.append(run)
                text_parts.append(run_text)
        else:
            run = ldm.Run()
            run.text = clean_control_chars(para_text)
            run.font = self._build_ldm_font(style_cp)
            para.runs.append(run)
            text_parts.append(run.text)

        para.text = "".join(text_parts)
        return para

    def _build_ldm_font(self, cp: CharProps) -> ldm.Font:
        from aspose.words_foss import light_document_model as ldm

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

        if cp.font_index >= 0 and cp.font_index in self._fonts:
            font.name = self._fonts[cp.font_index]
        elif self._default_font_index in self._fonts:
            font.name = self._fonts[self._default_font_index]
        font.size = cp.font_size if cp.font_size > 0 else _DEFAULT_FONT_SIZE_PT
        font.color = cp.color if cp.color else "Color [Empty]"
        font.highlight_color = cp.highlight_color if cp.highlight_color else "Color [Empty]"

        if cp.style_index >= 0 and cp.style_index in self._styles:
            font.style_name = self._styles[cp.style_index]
        else:
            font.style_name = "Default Paragraph Font"
        return font

    # -- hyperlink paragraph --------------------------------------------------

    def _build_ldm_hyperlink_paragraph(
        self, para_text: str, p_start: int, p_end: int
    ) -> ldm.Paragraph:
        from aspose.words_foss import light_document_model as ldm

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
            para.runs.append(run)
        para.text = clean_text
        return para

    # -- table builder --------------------------------------------------------

    def _build_ldm_table_from_text(self, table_text: str) -> tuple[ldm.Table, str]:
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
                        if "\x13" in ct:
                            ct = evaluate_fields(ct)
                        ct = clean_control_chars(ct)
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

    def _pre_scan_cell_cps(self, paragraphs: list[tuple[int, int]], text: str) -> set[int]:
        """Pre-scan body paragraphs to find CP starts of table cell content.

        Walks backwards from each ``\\x07`` paragraph, collecting
        preceding paragraphs that have direct PAPX formatting (non-empty
        ``_set_fields``).  Paragraphs without direct formatting are body
        text, not cell content.
        """
        cell_starts: set[int] = set()
        for idx, (p_start, p_end) in enumerate(paragraphs):
            if "\x07" not in text[p_start:p_end]:
                continue
            # The paragraph with \x07 is cell content.
            cell_starts.add(p_start)
            # Scan backwards for preceding cell paragraphs.
            for prev_idx in range(idx - 1, -1, -1):
                prev_start, prev_end = paragraphs[prev_idx]
                prev_text = text[prev_start:prev_end]
                if "\x0c" in prev_text:
                    break  # page break
                props = self._get_para_props_at(prev_start)
                if not props._set_fields:
                    break  # no direct formatting → not cell content
                cell_starts.add(prev_start)
        return cell_starts

    def _absorb_preceding_into_table(self, tbl: object, children: list) -> None:
        """Move preceding body paragraphs into the table's first cell.

        In DOC files, multi-paragraph table cells store only the final
        ``\\x07`` marker in the last paragraph.  Earlier paragraphs in
        the same cell appear as plain body children.  This method walks
        backwards through *children* and absorbs paragraphs that were
        tagged as cell content during the pre-scan (stored in
        ``_cell_para_starts``).
        """
        from aspose.words_foss import light_document_model as ldm

        if not tbl.rows or not tbl.rows[0].cells:
            return
        first_cell = tbl.rows[0].cells[0]

        absorb = 0
        for i in range(len(children) - 1, -1, -1):
            child = children[i]
            if not isinstance(child, ldm.Paragraph):
                break
            if not hasattr(child, "_cell_cp") or child._cell_cp not in self._cell_para_starts:
                break
            absorb += 1

        if absorb == 0:
            return

        merged = children[-absorb:]
        del children[-absorb:]
        first_cell.paragraphs = merged + first_cell.paragraphs

    # -- textbox injection (cover page) ---------------------------------------

    def _inject_textbox_content(self, body_children: list) -> None:
        """Inject textbox and fill-shape content using structural data.

        Uses FTXBXS (PlcftxbxTxt) to map shape IDs to textbox text
        ranges, and Escher ChildAnchor/Spgr records for child shape
        positions.  No font-size or position-ratio heuristics.
        """
        from aspose.words_foss import light_document_model as ldm

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
            target.inline_extras.append(shape)

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
        from aspose.words_foss import light_document_model as ldm

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
            para.runs = [run]
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
        from aspose.words_foss import light_document_model as ldm

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
            bar.shading = ldm.Shading(background_color=color_str)
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

            attach_para.inline_extras.append(bar)

        return handled

    # -- headers & footers ----------------------------------------------------

    def _build_ldm_headers_footers(
        self,
    ) -> tuple[list[ldm.Paragraph], list[ldm.Paragraph]]:
        from aspose.words_foss import light_document_model as ldm

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
                para.runs = [run]
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
                    headers[0].inline_extras.append(shape)
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
                    bar_para.inline_extras.append(bar)
                    headers.append(bar_para)

        return headers, footers
