"""
LDM table builder mixin for DOC files.

Extracts table-related logic from DocFileReader into a reusable mixin.
Methods here rely on attributes and helpers provided by DocFileReaderCore
and DocFileReader (e.g. ``_text``, ``_para_props``, ``_get_para_props_at``,
``_resolve_char_props``, ``_get_char_props_in_range``, ``_merge_char_props``,
``_build_ldm_font``, ``_cell_para_starts``).
"""


from aspose.words_foss.doc_reader.properties import parse_table_row_sprms
from aspose.words_foss.doc_reader.text import clean_control_chars, evaluate_fields
from aspose.words_foss import light_document_model as ldm


_HORZ_SPEC = {1: "center", 2: "right"}


class DocTableBuilderMixin:
    """Mixin that adds table-building helpers to the DOC reader."""

    def _build_ldm_table_from_text(
        self, table_text: str, p_start: int = 0, p_end: int = 0
    ) -> tuple[ldm.Table, str]:
        """Build a table from a body-builder text segment.

        Walks PAPX paragraphs in ``[p_start, p_end]``, grouping cells
        (PFInTable=1) into rows on PFTtp markers.  Falls back to
        plain ``\\x07`` segmentation when no in-table flag is set.
        """
        in_table_paras: list[tuple[int, int, object]] = []
        for s, e, props in self._para_props:
            if e <= p_start:
                continue
            if s >= p_end + 1:
                break
            if props.in_table:
                in_table_paras.append((s, e, props))

        if in_table_paras:
            return self._build_ldm_table_from_papx(in_table_paras, p_start, p_end)
        return self._build_ldm_table_from_text_legacy(table_text, p_start, p_end)

    def _build_ldm_table_from_papx(
        self,
        in_table_paras: list[tuple[int, int, object]],
        p_start: int,
        p_end: int,
    ) -> tuple[ldm.Table, str]:
        """Group in-table PAPX paragraphs into rows/cells.

        ``\\x07`` closes a cell; PFTtp closes a row.  A cell can hold
        multiple PAPX paragraphs (each ``\\r``-terminated) with only
        the last one ending in ``\\x07``.
        """
        tbl = ldm.Table()
        current_row: list[list[tuple[str, int, int]]] = []
        current_cell: list[tuple[str, int, int]] = []

        table_props = self._get_para_props_at(p_start)
        table_style_cp = self._resolve_char_props(table_props.istd)

        last_consumed_end = p_start
        for s, e, props in in_table_paras:
            last_consumed_end = e
            raw = self._text[s:e]
            if props.is_table_terminator:
                tail = raw.rstrip("\x07")
                if tail:
                    current_cell.append((tail, s, e))
                if current_cell:
                    current_row.append(current_cell)
                    current_cell = []
                row = self._row_from_cells(current_row, table_style_cp)
                if row is not None:
                    tbl.rows.append(row)
                current_row = []
                continue

            # Interior ``\x07`` splits one PAPX into multiple cells
            # (merged-cell layouts pack separators inside a single para).
            if "\x07" in raw:
                seg_start = s
                segs = raw.split("\x07")
                for seg_idx, seg in enumerate(segs):
                    seg_end = seg_start + len(seg)
                    is_last = seg_idx == len(segs) - 1
                    if not is_last:
                        current_cell.append((seg, seg_start, seg_end + 1))
                        current_row.append(current_cell)
                        current_cell = []
                        seg_start = seg_end + 1
                    elif seg:
                        current_cell.append((seg.rstrip("\r"), seg_start, seg_end))
                        seg_start = seg_end
            else:
                current_cell.append((raw.rstrip("\r"), s, e))

        if current_cell:
            current_row.append(current_cell)
        if current_row:
            row = self._row_from_cells(current_row, table_style_cp)
            if row is not None:
                tbl.rows.append(row)

        trailing = ""
        if last_consumed_end < p_end:
            trailing = self._text[last_consumed_end:p_end]

        self._apply_table_row_props(tbl, p_start, p_end)
        return tbl, trailing

    def _row_from_cells(
        self,
        cells_data: list,
        table_style_cp,
    ):
        """Build an ``ldm.Row`` from collected (text, cp_start, cp_end) cell segments."""
        if not cells_data:
            return None
        row = ldm.Row()
        for paras_data in cells_data:
            cell = ldm.Cell()
            ldm_paras = []
            if paras_data and isinstance(paras_data, tuple):
                paras_data = [paras_data]
            for cell_text, cell_cp, cell_end in paras_data:
                ldm_paras.append(self._cell_paragraph(
                    cell_text, cell_cp, cell_end, table_style_cp,
                ))
            cell.paragraphs = ldm_paras or [ldm.Paragraph()]
            row.cells.append(cell)
        return row

    def _cell_paragraph(
        self,
        cell_text: str,
        cell_cp: int,
        cell_end: int,
        table_style_cp,
    ) -> "ldm.Paragraph":
        """Build the LDM paragraph that lives inside a table cell."""
        para = ldm.Paragraph()
        ct = cell_text
        if "\x13" in ct:
            ct = evaluate_fields(ct)
        # Strip ``\x07`` / ``\x00`` (not legal XML), normalise ``\r`` to ``\n``.
        ct_clean = (
            clean_control_chars(ct)
            .replace("\x0c", "")
            .replace("\x07", "")
            .replace("\x00", "")
            .replace("\r", "\n")
        )
        para.text = ct_clean

        char_ranges = self._get_char_props_in_range(cell_cp, cell_end)
        if char_ranges and ct_clean:
            first_font_cp = None
            for cs, ce, cp in char_ranges:
                run_text = clean_control_chars(
                    self._text[cs:ce]
                ).replace("\x0c", "").replace("\x07", "")
                if not run_text:
                    continue
                merged = self._merge_char_props(
                    table_style_cp, cp, cp._set_fields
                )
                if first_font_cp is None:
                    first_font_cp = merged
                run = ldm.Run()
                run.text = run_text
                run.font = self._build_ldm_font(merged)
                para._children.append(run)
            built = "".join(r.text for r in para.runs)
            if built != ct_clean:
                para._children.clear()
                run = ldm.Run()
                run.text = ct_clean
                font_cp = first_font_cp or table_style_cp
                run.font = self._build_ldm_font(font_cp)
                para._children.append(run)
        elif ct_clean:
            run = ldm.Run()
            run.text = ct_clean
            run.font = self._build_ldm_font(table_style_cp)
            para._children.append(run)
        return para

    def _build_ldm_table_from_text_legacy(
        self, table_text: str, p_start: int = 0, p_end: int = 0
    ) -> tuple[ldm.Table, str]:
        """Plain ``\\x07``-segmentation fallback when PFInTable is missing."""
        tbl = ldm.Table()
        trailing = ""
        segments = table_text.split("\x07")
        current_row: list[tuple[str, int]] = []

        seg_positions: list[int] = []
        pos = p_start
        for seg in segments:
            seg_positions.append(pos)
            pos += len(seg) + 1

        table_props = self._get_para_props_at(p_start)
        table_style_cp = self._resolve_char_props(table_props.istd)

        for seg_idx, seg in enumerate(segments):
            seg_cp = seg_positions[seg_idx] if seg_idx < len(seg_positions) else p_start
            if seg == "":
                if current_row:
                    cells_for_row = [
                        [(ct, cell_cp, cell_cp + len(ct) + 1)]
                        for ct, cell_cp in current_row
                    ]
                    row = self._row_from_cells(cells_for_row, table_style_cp)
                    if row is not None:
                        tbl.rows.append(row)
                    current_row = []
            else:
                current_row.append((seg, seg_cp))

        if current_row:
            trailing = "\x07".join(ct for ct, _ in current_row)

        self._apply_table_row_props(tbl, p_start, p_end)
        return tbl, trailing

    def _apply_table_row_props(
        self, tbl: ldm.Table, p_start: int, p_end: int
    ) -> None:
        """Parse table SPRMs from row-end paragraphs and apply to LDM table."""
        if not tbl.rows:
            return

        # Find PAPX entries within the table CP range that have stored raw
        # grpprl containing table SPRMs (cell widths, positioning, etc.).
        row_sprms_list: list[object] = []
        for start, end, props in self._para_props:
            if start < p_start or start > p_end:
                continue
            if props._raw_grpprl is None:
                continue
            tp = parse_table_row_sprms(props._raw_grpprl)
            if tp.cell_widths or tp.has_positioning or tp.table_width > 0 or tp.borders:
                row_sprms_list.append(tp)

        if not row_sprms_list:
            return

        first_tp = row_sprms_list[0]

        if first_tp.table_width > 0:
            tbl.preferred_width = f"{first_tp.table_width}pt"
        elif first_tp.has_positioning and first_tp.cell_widths:
            total_width = sum(first_tp.cell_widths)
            if total_width > 0:
                tbl.preferred_width = f"{total_width}pt"

        if first_tp.has_positioning:
            tbl.text_wrapping = 1
            attrs: dict[str, str] = {}
            if first_tp.dxa_from_text:
                attrs["leftFromText"] = str(first_tp.dxa_from_text)
            if first_tp.dxa_from_text_right:
                attrs["rightFromText"] = str(first_tp.dxa_from_text_right)
            if first_tp.tblp_y:
                attrs["tblpY"] = str(first_tp.tblp_y)
            attrs["vertAnchor"] = "text"
            attrs["horzAnchor"] = "margin"
            if first_tp.horz_pos in _HORZ_SPEC:
                attrs["tblpXSpec"] = _HORZ_SPEC[first_tp.horz_pos]
            tbl._tblp_pr_attrs = attrs

        for row_idx, row in enumerate(tbl.rows):
            tp = row_sprms_list[row_idx] if row_idx < len(row_sprms_list) else first_tp
            if tp.borders:
                ldm_borders = []
                for brc_type, width_pt, color_str, _space in tp.borders:
                    ldm_borders.append(ldm.Border(
                        line_style=brc_type,
                        line_width=width_pt,
                        color=color_str,
                    ))
                row.row_format.borders = ldm_borders
            if tp.cell_widths:
                for cell_idx, cell in enumerate(row.cells):
                    if cell_idx < len(tp.cell_widths):
                        w_pt = tp.cell_widths[cell_idx]
                        cell.cell_format.width = w_pt

    def _pre_scan_cell_cps(self, paragraphs: list[tuple[int, int]], text: str) -> set[int]:
        """Find CPs that belong to multi-paragraph table cells.

        For each ``\\x07`` paragraph, walks backwards over paragraphs
        with direct PAPX formatting — those preceding paragraphs are
        body content of the same cell.
        """
        cell_starts: set[int] = set()
        for idx, (p_start, p_end) in enumerate(paragraphs):
            if "\x07" not in text[p_start:p_end]:
                continue
            cell_starts.add(p_start)
            for prev_idx in range(idx - 1, -1, -1):
                prev_start, prev_end = paragraphs[prev_idx]
                prev_text = text[prev_start:prev_end]
                if "\x0c" in prev_text:
                    break
                props = self._get_para_props_at(prev_start)
                if not props._set_fields:
                    break
                cell_starts.add(prev_start)
        return cell_starts

    def _absorb_preceding_into_table(self, tbl: object, children: list) -> None:
        """Pull cell-content paragraphs already emitted into the first cell.

        Multi-paragraph DOC cells only mark the final paragraph with
        ``\\x07``; earlier paragraphs surface as plain body children.
        """
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
