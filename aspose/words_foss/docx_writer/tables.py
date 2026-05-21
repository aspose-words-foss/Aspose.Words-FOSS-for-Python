"""Table (``<w:tbl>``) rendering — rows, cells, borders, shading."""


from typing import Mapping, Optional

from aspose.words_foss import light_document_model as ldm

from aspose.words_foss.docx_writer.bookmarks import BookmarkState
from aspose.words_foss.docx_writer.drawing import ImageRenderState
from aspose.words_foss.docx_writer.constants import (
    ALIGNMENT_VAL,
    LINE_STYLE_VAL,
    pt_to_eighths,
    pt_to_twips,
)
from aspose.words_foss.docx_writer.paragraphs import render_paragraph
from aspose.words_foss.docx_writer.runs import color_to_hex
from aspose.words_foss.docx_writer.xml_utils import el

_EMPTY_NUM_ID_MAP: Mapping[int, int] = {}

# Border slot indices matching :class:`BorderType` (0..7):
# 0=Bottom, 1=Left, 2=Right, 3=Top, 4=Horizontal (insideH / between),
# 5=Vertical (insideV), 6=DiagonalDown (tl2br), 7=DiagonalUp (tr2bl).
_SLOT_BOTTOM, _SLOT_LEFT, _SLOT_RIGHT, _SLOT_TOP = 0, 1, 2, 3
_SLOT_HORIZONTAL, _SLOT_VERTICAL = 4, 5
_SLOT_DIAG_DOWN, _SLOT_DIAG_UP = 6, 7

# Side-name → slot-index, in OOXML CT_TblBorders / CT_TcBorders schema
# order (top, left, bottom, right, insideH, insideV [, tl2br, tr2bl]).
_TBL_BORDER_SIDES: tuple[tuple[str, int], ...] = (
    ("top", _SLOT_TOP),
    ("left", _SLOT_LEFT),
    ("bottom", _SLOT_BOTTOM),
    ("right", _SLOT_RIGHT),
    ("insideH", _SLOT_HORIZONTAL),
    ("insideV", _SLOT_VERTICAL),
)
_TC_BORDER_SIDES: tuple[tuple[str, int], ...] = _TBL_BORDER_SIDES + (
    ("tl2br", _SLOT_DIAG_DOWN),
    ("tr2bl", _SLOT_DIAG_UP),
)
# Legacy alias retained for downstream tests that index it directly.
_BORDER_SIDES = tuple(s for s, _ in _TBL_BORDER_SIDES)
_VERT_ALIGN_VAL = {0: "top", 1: "center", 2: "bottom"}
_VERT_MERGE_VAL = {0: None, 1: "restart", 2: "continue"}


def _border_element(name: str, border: ldm.Border) -> str:
    val = LINE_STYLE_VAL.get(border.line_style, "none")
    # ``border.line_width`` is preserved as-is — the previous code's
    # ``or 0.5`` fallback inflated empty borders into a visible 0.5pt
    # line on round-trip.  ``pt_to_eighths`` already clamps to ``>= 1``
    # for the legitimate non-zero case (Word's smallest supported width).
    sz = pt_to_eighths(border.line_width) if border.line_width else 0
    color = color_to_hex(border.color) or "auto"
    return el(
        f"w:{name}",
        {"w:val": val, "w:sz": sz, "w:space": 0, "w:color": color},
    )


def _is_empty_border(border: ldm.Border) -> bool:
    return border.line_style == 0 and border.line_width == 0.0


def _by_slot(
    borders: list[ldm.Border],
    pairs: tuple[tuple[str, int], ...],
) -> list[str]:
    """Emit ``<w:NAME>`` elements aligned to the BorderType slot list.

    Each pair is ``(side-name, slot-index)`` from
    :data:`BORDER_SLOTS`.  Slots beyond the supplied list — or those
    carrying an empty Border — are omitted so the resulting XML stays
    minimal and the reader round-trips back to the same trimmed
    ``_empty_borders`` shape.
    """
    if not borders:
        return []
    out: list[str] = []
    for side, slot in pairs:
        if slot >= len(borders):
            continue
        border = borders[slot]
        if _is_empty_border(border):
            continue
        out.append(_border_element(side, border))
    return out


def _per_side_borders(borders: list[ldm.Border], sides: tuple[str, ...]) -> list[str]:
    """Legacy positional emission used by callers that hand us a tuple of names.

    Kept for backwards compatibility with code that still pre-flattens
    its border list.  New code should use :func:`_by_slot` to align
    with the canonical :data:`BORDER_SLOTS` BorderType layout.
    """
    if not borders:
        return []
    relevant = [borders[i] if i < len(borders) else borders[0] for i, _ in enumerate(sides)]
    if all(_is_empty_border(b) for b in relevant):
        return []
    out: list[str] = []
    for i, side in enumerate(sides):
        border = borders[i] if i < len(borders) else borders[0]
        out.append(_border_element(side, border))
    return out


def _table_borders(borders: list[ldm.Border]) -> str:
    sides = _by_slot(borders, _TBL_BORDER_SIDES)
    return el("w:tblBorders", None, sides) if sides else ""


def _cell_borders(borders: list[ldm.Border]) -> str:
    sides = _by_slot(borders, _TC_BORDER_SIDES)
    return el("w:tcBorders", None, sides) if sides else ""


def _preferred_width_attrs(value: str) -> dict[str, object]:
    """Parse an LDM ``preferred_width`` string into OOXML ``w:tblW`` /
    ``w:tcW`` attributes.

    Accepts every form the reader and the test fixtures can produce,
    plus the bare ``PreferredWidthType`` enum names so consumers can
    construct a value-less ``"Auto"`` / ``"Percent"`` / ``"Points"``
    placeholder:

    * ``""`` / ``"Auto"``         → ``w:type="auto"`` ()
    * ``"<N>%"``                  → ``w:type="pct"``  (fifty-units-per-percent)
    * ``"<N>pt"`` or bare ``"<N>"`` → ``w:type="dxa"`` (points converted to twips per the OOXML schema)
    * ``"Percent"`` / ``"Points"`` → corresponding type, ``w=0``
      (degenerate but legal OOXML — matches the
      ``PreferredWidthType`` enum names verbatim)
    * Anything else → ``w:type="auto"`` fallback so a malformed
      value never breaks the document.
    """
    txt = (value or "").strip()
    if not txt or txt == "Auto":
        return {"w:w": 0, "w:type": "auto"}
    if txt == "Percent":
        return {"w:w": 0, "w:type": "pct"}
    if txt == "Points":
        return {"w:w": 0, "w:type": "dxa"}
    if txt.endswith("%"):
        try:
            pct = float(txt[:-1]) * 50.0  # _PCT_DIVISOR inverse
        except ValueError:
            return {"w:w": 0, "w:type": "pct"}
        return {"w:w": int(round(pct)), "w:type": "pct"}
    if txt.endswith("pt"):
        try:
            pt = float(txt[:-2])
        except ValueError:
            return {"w:w": 0, "w:type": "dxa"}
        return {"w:w": pt_to_twips(pt), "w:type": "dxa"}
    # Bare numeric — ``PreferredWidthType.Points``: interpret
    # as points and convert to twips.
    try:
        pt = float(txt)
    except ValueError:
        return {"w:w": 0, "w:type": "auto"}
    return {"w:w": pt_to_twips(pt), "w:type": "dxa"}


def _is_auto_preferred_width(value: str) -> bool:
    """``True`` when *value* names the ``PreferredWidthType.Auto`` case."""
    return not value or value.strip() in ("", "Auto")


def _render_tblW(value: str) -> str:
    """Render ``<w:tblW>`` from an LDM ``preferred_width`` string."""
    return el("w:tblW", _preferred_width_attrs(value))


def _render_tcMar(fmt: ldm.CellFormat) -> str:
    """Build a ``<w:tcMar>`` for the cell's per-side padding.

    OOXML schema (CT_TcMar) accepts ``top`` / ``left`` / ``bottom`` /
    ``right`` children, each carrying ``w:w`` (twips) + ``w:type="dxa"``.
    Only sides with a non-zero LDM value are emitted so cells without
    custom padding don't get a noisy element.
    """
    sides: list[tuple[str, float]] = [
        ("w:top", fmt.top_padding),
        ("w:left", fmt.left_padding),
        ("w:bottom", fmt.bottom_padding),
        ("w:right", fmt.right_padding),
    ]
    children = [
        el(side, {"w:w": pt_to_twips(value), "w:type": "dxa"}) for side, value in sides if value
    ]
    if not children:
        return ""
    return el("w:tcMar", None, children)


def _tcPr(cell: ldm.Cell) -> str:
    """Build a ``<w:tcPr>`` honouring the OOXML schema order.

    ``CT_TcPrBase`` requires:
    cnfStyle → tcW → gridSpan → hMerge → vMerge → tcBorders → shd →
    noWrap → tcMar → textDirection → tcFitText → vAlign → hideMark →
    headers.
    """
    fmt = cell.cell_format
    children: list[str] = []

    # ``<w:cnfStyle>`` is the *first* child of CT_TcPrBase — emitting it
    # last produced a schema-invalid order that Word flags as document
    # corruption (other parsers accept it silently).
    if fmt.conditional_style:
        children.append(el("w:cnfStyle", {"w:val": fmt.conditional_style}))

    # Always emit <w:tcW/> — Word / Aspose fixtures carry it on every
    # cell, falling back to ``w:type="auto"`` when no explicit width.
    if fmt.width > 0:
        children.append(
            el("w:tcW", {"w:w": pt_to_twips(fmt.width), "w:type": "dxa"})
        )
    else:
        children.append(el("w:tcW", _preferred_width_attrs(fmt.preferred_width)))

    if fmt.horizontal_merge == 1:
        children.append(el("w:hMerge", {"w:val": "restart"}))
    elif fmt.horizontal_merge == 2:
        children.append(el("w:hMerge", {"w:val": "continue"}))

    vert_merge = _VERT_MERGE_VAL.get(fmt.vertical_merge)
    if vert_merge is not None:
        children.append(el("w:vMerge", {"w:val": vert_merge}))

    borders = _cell_borders(fmt.borders)
    if borders:
        children.append(borders)

    shading_hex = color_to_hex(fmt.shading.background_pattern_color)
    if shading_hex:
        children.append(
            el(
                "w:shd",
                {"w:val": "clear", "w:color": "auto", "w:fill": shading_hex},
            )
        )
    elif fmt.shading.background_pattern_color:
        # Non-hex sentinel ("auto") — re-emit verbatim so the reader
        # recovers the same string instead of dropping the field to "".
        children.append(
            el(
                "w:shd",
                {
                    "w:val": "clear",
                    "w:color": "auto",
                    "w:fill": fmt.shading.background_pattern_color,
                },
            )
        )

    if not fmt.wrap_text:
        children.append(el("w:noWrap"))

    tcMar = _render_tcMar(fmt)
    if tcMar:
        children.append(tcMar)

    if fmt.orientation != 0:
        td_val = {1: "btLr", 2: "tbRl", 3: "lrTbV", 4: "tbRlV", 5: "tbLrV"}.get(
            fmt.orientation, ""
        )
        if td_val:
            children.append(el("w:textDirection", {"w:val": td_val}))

    valign = _VERT_ALIGN_VAL.get(fmt.vertical_alignment)
    if valign and valign != "top":
        children.append(el("w:vAlign", {"w:val": valign}))

    if not children:
        return ""
    return el("w:tcPr", None, children)


def _render_cell(
    cell: ldm.Cell,
    rels: dict,
    *,
    num_id_map: Mapping[int, int] = _EMPTY_NUM_ID_MAP,
    image_state: Optional[ImageRenderState] = None,
    bookmark_state: Optional[BookmarkState] = None,
    style_pf_map: Optional[Mapping[str, ldm.ParagraphFormat]] = None,
    style_id_map: Optional[Mapping[str, str]] = None,
    style_font_map: Optional[Mapping[str, ldm.Font]] = None,
) -> str:
    """Render a single LDM cell.

    OOXML allows tables to nest inside cells (``w:tc/w:tbl``), and the
    reader populates ``cell.tables``.  We emit those alongside the
    cell's paragraphs; per the spec the cell must still end with a
    paragraph, so a synthetic empty ``<w:p>`` is appended when the
    cell holds *only* nested tables.
    """
    children: list[str] = []
    tcPr = _tcPr(cell)
    if tcPr:
        children.append(tcPr)

    has_content = bool(cell.paragraphs or cell.tables)
    for p in cell.paragraphs:
        children.append(
            render_paragraph(
                p,
                rels,
                num_id_map=num_id_map,
                image_state=image_state,
                bookmark_state=bookmark_state,
                style_pf_map=style_pf_map,
                style_id_map=style_id_map,
                style_font_map=style_font_map,
            )
        )
    for t in cell.tables:
        children.append(
            render_table(
                t,
                rels,
                num_id_map=num_id_map,
                image_state=image_state,
                bookmark_state=bookmark_state,
                style_pf_map=style_pf_map,
                style_id_map=style_id_map,
                style_font_map=style_font_map,
            )
        )

    # OOXML requires every w:tc to end with a w:p.  Append a placeholder
    # paragraph after a trailing nested table, or as the only child if
    # the cell is empty.
    needs_trailing_p = (not has_content) or bool(cell.tables)
    if needs_trailing_p:
        children.append(el("w:p"))
    return el("w:tc", None, children)


def _trPr(row: ldm.Row) -> str:
    """Build a ``<w:trPr>`` honouring the OOXML schema order.

    ``CT_TrPrBase`` requires:
    cnfStyle → divId → gridBefore → gridAfter → wBefore → wAfter →
    cantSplit → trHeight → tblHeader → tblCellSpacing → jc → hidden.
    """
    fmt = row.row_format
    children: list[str] = []
    # ``<w:cnfStyle>`` is the *first* child of CT_TrPrBase — emitting it
    # last produced a schema-invalid order that Word flags as document
    # corruption (other parsers accept it silently).
    if fmt.conditional_style:
        children.append(el("w:cnfStyle", {"w:val": fmt.conditional_style}))
    if not fmt.allow_break_across_pages:
        children.append(el("w:cantSplit"))
    if fmt.height > 0:
        rule = {0: "atLeast", 1: "exact"}.get(fmt.height_rule, "atLeast")
        children.append(el("w:trHeight", {"w:val": pt_to_twips(fmt.height), "w:hRule": rule}))
    if fmt.heading_format:
        children.append(el("w:tblHeader"))
    if not children:
        return ""
    return el("w:trPr", None, children)


def _render_row(
    row: ldm.Row,
    rels: dict,
    *,
    num_id_map: Mapping[int, int] = _EMPTY_NUM_ID_MAP,
    image_state: Optional[ImageRenderState] = None,
    bookmark_state: Optional[BookmarkState] = None,
    style_pf_map: Optional[Mapping[str, ldm.ParagraphFormat]] = None,
    style_id_map: Optional[Mapping[str, str]] = None,
    style_font_map: Optional[Mapping[str, ldm.Font]] = None,
) -> str:
    children: list[str] = []
    # CT_Row order: tblPrEx (0), trPr (1), tc (2+).
    if not _is_auto_preferred_width(row.row_format.preferred_width):
        children.append(
            el("w:tblPrEx", None, _render_tblW(row.row_format.preferred_width))
        )
    trPr = _trPr(row)
    if trPr:
        children.append(trPr)
    for c in row.cells:
        children.append(
            _render_cell(
                c,
                rels,
                num_id_map=num_id_map,
                image_state=image_state,
                bookmark_state=bookmark_state,
                style_pf_map=style_pf_map,
                style_id_map=style_id_map,
                style_font_map=style_font_map,
            )
        )
    return el("w:tr", None, children)


def _tblPr(
    table: ldm.Table,
    *,
    style_id_map: Optional[Mapping[str, str]] = None,
) -> str:
    """Build a ``<w:tblPr>`` honouring the OOXML schema order.

    ``CT_TblPrBase`` requires:
    tblStyle → tblpPr → tblOverlap → bidiVisual → tblStyleRowBandSize →
    tblStyleColBandSize → tblW → jc → tblCellSpacing → tblInd →
    tblBorders → shd → tblLayout → tblCellMar → tblLook → tblCaption →
    tblDescription.

    ``style_id_map`` resolves ``table.style_name`` to the on-disk
    ``w:styleId`` the writer emitted for that style — without it a
    document carrying a table style whose source ``styleId`` was a
    short token like ``"-11"`` would end up with a ``<w:tblStyle
    w:val="-11"/>`` reference against a styles.xml that now declares
    the same style under a sanitized ``"ListTable1LightAccent1"`` id,
    leaving MS Word with a dangling reference it refuses to open.
    """
    children: list[str] = []
    if table.style_name:
        if style_id_map is not None:
            from aspose.words_foss.docx_writer.styles_part import _sanitize_style_id
            tbl_style_id = style_id_map.get(
                table.style_name,
                _sanitize_style_id(table.style_name.replace(" ", "")),
            )
        else:
            tbl_style_id = table.style_name
        children.append(el("w:tblStyle", {"w:val": tbl_style_id}))
    if table.text_wrapping:
        attrs = table._tblp_pr_attrs
        if attrs:
            children.append(el("w:tblpPr", {f"w:{k}": v for k, v in attrs.items()}))
    # Always emit <w:tblW/>; "Auto" round-trips as w:type="auto".
    children.append(_render_tblW(table.preferred_width))
    if table.alignment:
        val = ALIGNMENT_VAL.get(table.alignment)
        if val:
            children.append(el("w:jc", {"w:val": val}))
    if table.left_indent:
        children.append(
            el(
                "w:tblInd",
                {"w:w": pt_to_twips(table.left_indent), "w:type": "dxa"},
            )
        )

    if table.rows:
        tbl_borders = _table_borders(table.rows[0].row_format.borders)
        if tbl_borders:
            children.append(tbl_borders)

    margins: list[str] = []
    margin_map = {
        "top": table.top_padding,
        "left": table.left_padding,
        "bottom": table.bottom_padding,
        "right": table.right_padding,
    }
    for side, val in margin_map.items():
        if val > 0:
            margins.append(
                el(
                    f"w:{side}",
                    {"w:w": pt_to_twips(val), "w:type": "dxa"},
                )
            )
    if margins:
        children.append(el("w:tblCellMar", None, margins))

    if not children:
        return ""
    return el("w:tblPr", None, children)


def _tblGrid(table: ldm.Table) -> str:
    """Emit a ``<w:tblGrid>`` with per-column widths from row 0."""
    if not table.rows:
        return el("w:tblGrid")
    cells = table.rows[0].cells
    cols: list[str] = []
    for cell in cells:
        width = cell.cell_format.width
        attrs = {"w:w": pt_to_twips(width) if width > 0 else 2400}
        cols.append(el("w:gridCol", attrs))
    return el("w:tblGrid", None, cols)


def render_table(
    table: ldm.Table,
    rels: dict,
    *,
    num_id_map: Mapping[int, int] = _EMPTY_NUM_ID_MAP,
    image_state: Optional[ImageRenderState] = None,
    bookmark_state: Optional[BookmarkState] = None,
    style_pf_map: Optional[Mapping[str, ldm.ParagraphFormat]] = None,
    style_id_map: Optional[Mapping[str, str]] = None,
    style_font_map: Optional[Mapping[str, ldm.Font]] = None,
) -> str:
    """Render a single LDM table to ``<w:tbl>...</w:tbl>``.

    The optional state args (``image_state`` / ``bookmark_state`` /
    ``style_pf_map`` / ``style_id_map`` / ``style_font_map``) thread
    through to every paragraph emitted from a cell so cells share the
    document-wide relationship counters and style maps.
    """
    children: list[str] = []
    tblPr = _tblPr(table, style_id_map=style_id_map)
    if tblPr:
        children.append(tblPr)
    children.append(_tblGrid(table))
    for r in table.rows:
        children.append(
            _render_row(
                r,
                rels,
                num_id_map=num_id_map,
                image_state=image_state,
                bookmark_state=bookmark_state,
                style_pf_map=style_pf_map,
                style_id_map=style_id_map,
                style_font_map=style_font_map,
            )
        )
    return el("w:tbl", None, children)
