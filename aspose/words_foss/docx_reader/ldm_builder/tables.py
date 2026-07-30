"""Build :class:`ldm.Table`, :class:`ldm.Row`, :class:`ldm.Cell`.

The cell builder defers paragraph and nested-table construction back
to a caller-supplied :class:`ParagraphBuilder`, which lets the
paragraph and table builders depend on each other without a cyclic
import.
"""

from typing import Callable, Optional
from xml.etree import ElementTree as ET

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_reader.constants import (
    W_NS,
    _ALIGNMENT_MAP,
    _PCT_DIVISOR,
    _TWIPS_PER_PT,
)
from aspose.words_foss.docx_reader.utils import _empty_borders
from aspose.words_foss.model.enums import CellVerticalAlignment as _CVA

from ._helpers import apply_padding_sides, build_borders, build_shading
from ._context import ReaderContext


# Paragraph builder callback signature: (p_elem, image_rels?) -> ldm.Paragraph.
ParagraphBuilderFn = Callable[..., ldm.Paragraph]


_VERT_ALIGN_MAP = {
    "top": _CVA.TOP,
    "center": _CVA.CENTER,
    "bottom": _CVA.BOTTOM,
}

_TEXT_DIRECTION_MAP: dict[str, int] = {
    "btLr": 1,
    "tbRl": 2,
    "lrTbV": 3,
    "tbRlV": 4,
    "tbLrV": 5,
}

_HEIGHT_RULE_EXACT = 1
_HEIGHT_RULE_AUTO = 2
_VMERGE_RESTART = 1
_VMERGE_CONTINUE = 2
_GRID_SPAN_MERGE = 1
_FLOATING_TEXT_WRAPPING = 1

_FLOATING_ATTRS: tuple[str, ...] = (
    "leftFromText", "rightFromText",
    "topFromText", "bottomFromText",
    "vertAnchor", "horzAnchor",
    "tblpX", "tblpXSpec",
    "tblpY", "tblpYSpec",
)


class TableBuilder:
    """Build :class:`ldm.Table` from a ``<w:tbl>`` element."""

    def __init__(
        self,
        ctx: ReaderContext,
        paragraph_builder: ParagraphBuilderFn,
    ):
        self._ctx = ctx
        self._paragraph_builder = paragraph_builder
        self._row_builder = RowBuilder(ctx, paragraph_builder, self)

    def build(self, tbl_elem: ET.Element) -> ldm.Table:
        """Translate one ``<w:tbl>`` into :class:`ldm.Table` (rows + cells inclusive)."""
        tbl = ldm.Table()
        table_borders: list[ldm.Border] = []
        tblPr = tbl_elem.find(f"{W_NS}tblPr")
        if tblPr is not None:
            table_borders = self._apply_table_properties(tbl, tblPr)

        default_paddings = (
            tbl.left_padding, tbl.right_padding, tbl.top_padding, tbl.bottom_padding,
        )
        self._build_rows(tbl_elem, tbl, table_borders, default_paddings)
        return tbl

    def _apply_table_properties(
        self,
        tbl: ldm.Table,
        tblPr: ET.Element,
    ) -> list[ldm.Border]:
        tblStyle = tblPr.find(f"{W_NS}tblStyle")
        if tblStyle is not None:
            # Resolve the styleId to the style's display name (same as
            # ``<w:pStyle>`` handling) so the LDM holds a stable
            # identifier the writer's ``style_id_map`` can map back —
            # otherwise a ``<w:tblStyle w:val="-11"/>`` reference
            # against a renamed-on-write styleId becomes a dangling
            # ref that MS Word rejects on load.
            raw_id = tblStyle.get(f"{W_NS}val", "")
            tbl.style_name = self._ctx._resolve_style_name(raw_id) if raw_id else ""
        jc = tblPr.find(f"{W_NS}jc")
        if jc is not None:
            tbl.alignment = _ALIGNMENT_MAP.get(jc.get(f"{W_NS}val", "left"), 0)
        tblW = tblPr.find(f"{W_NS}tblW")
        if tblW is not None:
            tbl.preferred_width = self._parse_preferred_width(tblW)
        tblInd = tblPr.find(f"{W_NS}tblInd")
        if tblInd is not None:
            tbl.left_indent = int(tblInd.get(f"{W_NS}w", "0")) / _TWIPS_PER_PT
        borders: list[ldm.Border] = []
        tblBorders_elem = tblPr.find(f"{W_NS}tblBorders")
        if tblBorders_elem is not None:
            borders = build_borders(tblBorders_elem)
        tblCellMar = tblPr.find(f"{W_NS}tblCellMar")
        if tblCellMar is not None:
            apply_padding_sides(tbl, tblCellMar)
        tblpPr = tblPr.find(f"{W_NS}tblpPr")
        if tblpPr is not None:
            self._apply_floating(tbl, tblpPr)
        caption = tblPr.find(f"{W_NS}tblCaption")
        if caption is not None:
            tbl.title = caption.get(f"{W_NS}val", "")
        desc = tblPr.find(f"{W_NS}tblDescription")
        if desc is not None:
            tbl.description = desc.get(f"{W_NS}val", "")
        return borders

    @staticmethod
    def _parse_preferred_width(tblW: ET.Element) -> ldm.PreferredWidth:
        w_type = tblW.get(f"{W_NS}type", "")
        w_val = tblW.get(f"{W_NS}w", "0")
        if w_type == "pct":
            return ldm.PreferredWidth.from_percent(int(w_val) / _PCT_DIVISOR)
        if w_type == "dxa":
            return ldm.PreferredWidth.from_points(int(w_val) / _TWIPS_PER_PT)
        return ldm.PreferredWidth.auto()

    @staticmethod
    def _apply_floating(tbl: ldm.Table, tblpPr: ET.Element) -> None:
        tbl.text_wrapping = _FLOATING_TEXT_WRAPPING
        attrs: dict[str, str] = {}
        for attr_name in _FLOATING_ATTRS:
            value = tblpPr.get(f"{W_NS}{attr_name}", "")
            if value:
                attrs[attr_name] = value
        tbl._tblp_pr_attrs = attrs

    def _build_rows(
        self,
        tbl_elem: ET.Element,
        tbl: ldm.Table,
        table_borders: list[ldm.Border],
        default_paddings: tuple[float, float, float, float],
    ) -> None:
        ctx = self._ctx
        previous_style = getattr(ctx, "_current_table_style_id", "")
        ctx._current_table_style_id = tbl.style_name
        try:
            for tr_elem in tbl_elem.findall(f"{W_NS}tr"):
                row = self._row_builder.build(tr_elem, default_paddings)
                if table_borders:
                    row.row_format.borders = table_borders
                tbl.rows.append(row)
        finally:
            ctx._current_table_style_id = previous_style


class RowBuilder:
    """Build :class:`ldm.Row` from a ``<w:tr>`` element."""

    def __init__(
        self,
        ctx: ReaderContext,
        paragraph_builder: ParagraphBuilderFn,
        table_builder: TableBuilder,
    ):
        self._cell_builder = CellBuilder(ctx, paragraph_builder, table_builder)

    def build(
        self,
        tr_elem: ET.Element,
        default_paddings: Optional[tuple[float, float, float, float]] = None,
    ) -> ldm.Row:
        """Translate one ``<w:tr>``; *default_paddings* fall through to each cell."""
        row = ldm.Row()
        trPr = tr_elem.find(f"{W_NS}trPr")
        if trPr is not None:
            row.row_format = self._build_row_format(trPr)
        # Per-row override (CT_TblPrExBase); only width carried in the LDM.
        tblPrEx = tr_elem.find(f"{W_NS}tblPrEx")
        if tblPrEx is not None:
            tblW = tblPrEx.find(f"{W_NS}tblW")
            if tblW is not None:
                row.preferred_width = TableBuilder._parse_preferred_width(tblW)
        for tc_elem in tr_elem.findall(f"{W_NS}tc"):
            row.cells.append(self._cell_builder.build(tc_elem, default_paddings))
        return row

    @staticmethod
    def _build_row_format(trPr: ET.Element) -> ldm.RowFormat:
        rf = ldm.RowFormat()
        trHeight = trPr.find(f"{W_NS}trHeight")
        if trHeight is not None:
            rf.height = int(trHeight.get(f"{W_NS}val", "0")) / _TWIPS_PER_PT
            rule = trHeight.get(f"{W_NS}hRule", "")
            if rule == "exact":
                rf.height_rule = _HEIGHT_RULE_EXACT
            elif rule == "auto":
                rf.height_rule = _HEIGHT_RULE_AUTO
            else:
                rf.height_rule = 0
        if trPr.find(f"{W_NS}tblHeader") is not None:
            rf.heading_format = True
        if trPr.find(f"{W_NS}cantSplit") is not None:
            rf.allow_break_across_pages = False
        cnf = trPr.find(f"{W_NS}cnfStyle")
        if cnf is not None:
            rf.conditional_style = ldm.ConditionalStyleMask.from_val(cnf.get(f"{W_NS}val", ""))
        return rf


class CellBuilder:
    """Build :class:`ldm.Cell` from a ``<w:tc>`` element."""

    def __init__(
        self,
        ctx: ReaderContext,
        paragraph_builder: ParagraphBuilderFn,
        table_builder: TableBuilder,
    ):
        self._ctx = ctx
        self._paragraph_builder = paragraph_builder
        self._table_builder = table_builder

    def build(
        self,
        tc_elem: ET.Element,
        default_paddings: Optional[tuple[float, float, float, float]] = None,
    ) -> ldm.Cell:
        """Translate one ``<w:tc>``; *default_paddings* inherit from the parent table."""
        cell = ldm.Cell()
        tcPr = tc_elem.find(f"{W_NS}tcPr")
        if tcPr is not None:
            cell.cell_format = self._build_cell_format(tcPr, default_paddings)

        for p_elem in tc_elem.findall(f"{W_NS}p"):
            cell.paragraphs.append(self._paragraph_builder(p_elem))
        for nested_tbl in tc_elem.findall(f"{W_NS}tbl"):
            cell.tables.append(self._table_builder.build(nested_tbl))
        return cell

    def _build_cell_format(
        self,
        tcPr: ET.Element,
        default_paddings: Optional[tuple[float, float, float, float]],
    ) -> ldm.CellFormat:
        cf = ldm.CellFormat()
        if default_paddings is not None:
            (
                cf.left_padding,
                cf.right_padding,
                cf.top_padding,
                cf.bottom_padding,
            ) = default_paddings
        self._apply_width(tcPr, cf)
        self._apply_vertical_align(tcPr, cf)
        self._apply_merge(tcPr, cf)
        self._apply_grid_span(tcPr, cf)
        tcMar = tcPr.find(f"{W_NS}tcMar")
        if tcMar is not None:
            apply_padding_sides(cf, tcMar)
        shd = tcPr.find(f"{W_NS}shd")
        if shd is not None:
            cf.shading = build_shading(shd)
        tcBorders = tcPr.find(f"{W_NS}tcBorders")
        cf.borders = build_borders(tcBorders) if tcBorders is not None else _empty_borders()
        self._apply_text_direction(tcPr, cf)
        self._apply_no_wrap(tcPr, cf)
        cnf = tcPr.find(f"{W_NS}cnfStyle")
        if cnf is not None:
            cf.conditional_style = ldm.ConditionalStyleMask.from_val(cnf.get(f"{W_NS}val", ""))
        return cf

    @staticmethod
    def _apply_width(tcPr: ET.Element, cf: ldm.CellFormat) -> None:
        tcW = tcPr.find(f"{W_NS}tcW")
        if tcW is None:
            return
        w_type = tcW.get(f"{W_NS}type", "")
        w_val = tcW.get(f"{W_NS}w", "0")
        if w_type == "dxa":
            cf.width = int(w_val) / _TWIPS_PER_PT
            cf.preferred_width = ldm.PreferredWidth.from_points(cf.width)
        elif w_type == "pct":
            cf.preferred_width = ldm.PreferredWidth.from_percent(int(w_val) / _PCT_DIVISOR)
        elif w_type == "auto":
            cf.preferred_width = ldm.PreferredWidth.auto()

    @staticmethod
    def _apply_vertical_align(tcPr: ET.Element, cf: ldm.CellFormat) -> None:
        vAlign = tcPr.find(f"{W_NS}vAlign")
        if vAlign is not None:
            cf.vertical_alignment = _VERT_ALIGN_MAP.get(
                vAlign.get(f"{W_NS}val", "top"), _CVA.TOP
            )

    @staticmethod
    def _apply_merge(tcPr: ET.Element, cf: ldm.CellFormat) -> None:
        vMerge = tcPr.find(f"{W_NS}vMerge")
        if vMerge is None:
            return
        cf.vertical_merge = (
            _VMERGE_RESTART
            if vMerge.get(f"{W_NS}val", "continue") == "restart"
            else _VMERGE_CONTINUE
        )

    @staticmethod
    def _apply_grid_span(tcPr: ET.Element, cf: ldm.CellFormat) -> None:
        gridSpan = tcPr.find(f"{W_NS}gridSpan")
        if gridSpan is None:
            return
        span = int(gridSpan.get(f"{W_NS}val", "1"))
        if span > 1:
            cf.horizontal_merge = _GRID_SPAN_MERGE

    @staticmethod
    def _apply_text_direction(tcPr: ET.Element, cf: ldm.CellFormat) -> None:
        textDirection = tcPr.find(f"{W_NS}textDirection")
        if textDirection is not None:
            cf.orientation = _TEXT_DIRECTION_MAP.get(
                textDirection.get(f"{W_NS}val", ""), 0
            )

    @staticmethod
    def _apply_no_wrap(tcPr: ET.Element, cf: ldm.CellFormat) -> None:
        noWrap = tcPr.find(f"{W_NS}noWrap")
        if noWrap is None:
            return
        val = noWrap.get(f"{W_NS}val")
        if val is None or val not in ("false", "0"):
            cf.wrap_text = False
