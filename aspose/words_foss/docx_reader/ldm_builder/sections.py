"""Build :class:`ldm.Section` and :class:`ldm.PageSetup`.

The section splitter walks body children in document order,
emitting one :class:`ldm.Section` each time a ``<w:sectPr>`` is
encountered — either nested in a paragraph's ``pPr`` or as the
trailing standalone body break.
"""

from typing import Callable, Optional
from xml.etree import ElementTree as ET

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_reader.constants import (
    W_NS,
    _A3_HEIGHT_PT,
    _A3_WIDTH_PT,
    _A4_HEIGHT_PT,
    _A4_WIDTH_PT,
    _LEGAL_HEIGHT_PT,
    _LETTER_HEIGHT_PT,
    _LETTER_WIDTH_PT,
    _PAPER_A3,
    _PAPER_A4,
    _PAPER_CUSTOM,
    _PAPER_LEGAL,
    _PAPER_LETTER,
    _PAPER_SIZE_TOLERANCE_A_PT,
    _PAPER_SIZE_TOLERANCE_PT,
    _SECTION_START_MAP,
    _TWIPS_PER_PT,
)
from aspose.words_foss.model.enums import Orientation as _Orient

from .paragraphs import ParagraphBuilder
from ._context import ReaderContext


_DEFAULT_SECTION_START = 2

_BodyChild = ldm.Paragraph | ldm.Table | ldm.UnknownNode
TableBuilderFn = Callable[[ET.Element], ldm.Table]


_PAPER_TABLE: tuple[tuple[float, float, float, int], ...] = (
    (_LETTER_WIDTH_PT, _LETTER_HEIGHT_PT, _PAPER_SIZE_TOLERANCE_PT, _PAPER_LETTER),
    (_LETTER_WIDTH_PT, _LEGAL_HEIGHT_PT, _PAPER_SIZE_TOLERANCE_PT, _PAPER_LEGAL),
    (_A4_WIDTH_PT, _A4_HEIGHT_PT, _PAPER_SIZE_TOLERANCE_A_PT, _PAPER_A4),
    (_A3_WIDTH_PT, _A3_HEIGHT_PT, _PAPER_SIZE_TOLERANCE_A_PT, _PAPER_A3),
)


def detect_paper_size(width: float, height: float) -> int:
    """Detect a standard paper-size constant from raw point dimensions."""
    w, h = min(width, height), max(width, height)
    for std_w, std_h, tol, paper_id in _PAPER_TABLE:
        if abs(w - std_w) < tol and abs(h - std_h) < tol:
            return paper_id
    return _PAPER_CUSTOM


class PageSetupBuilder:
    """Build :class:`ldm.PageSetup` from a ``<w:sectPr>`` element."""

    _MARGIN_FIELDS: tuple[tuple[str, str], ...] = (
        ("top", "top_margin"),
        ("bottom", "bottom_margin"),
        ("left", "left_margin"),
        ("right", "right_margin"),
        ("header", "header_distance"),
        ("footer", "footer_distance"),
        ("gutter", "gutter"),
    )

    def build(self, sect_pr: ET.Element) -> ldm.PageSetup:
        """Translate one ``<w:sectPr>`` into :class:`ldm.PageSetup`."""
        ps = ldm.PageSetup()
        self._apply_size(sect_pr, ps)
        self._apply_margins(sect_pr, ps)
        self._apply_section_type(sect_pr, ps)
        self._apply_title_page(sect_pr, ps)
        self._apply_page_numbering(sect_pr, ps)
        self._apply_columns(sect_pr, ps)
        return ps

    @staticmethod
    def _apply_size(sect_pr: ET.Element, ps: ldm.PageSetup) -> None:
        pg_sz = sect_pr.find(f"{W_NS}pgSz")
        if pg_sz is None:
            return
        w = pg_sz.get(f"{W_NS}w", "")
        h = pg_sz.get(f"{W_NS}h", "")
        if w:
            ps.page_width = int(w) / _TWIPS_PER_PT
        if h:
            ps.page_height = int(h) / _TWIPS_PER_PT
        orient = pg_sz.get(f"{W_NS}orient", "")
        ps.orientation = _Orient.LANDSCAPE if orient == "landscape" else _Orient.PORTRAIT
        ps.paper_size = detect_paper_size(ps.page_width, ps.page_height)

    @classmethod
    def _apply_margins(cls, sect_pr: ET.Element, ps: ldm.PageSetup) -> None:
        pg_mar = sect_pr.find(f"{W_NS}pgMar")
        if pg_mar is None:
            return
        for attr, field_name in cls._MARGIN_FIELDS:
            val = pg_mar.get(f"{W_NS}{attr}", "")
            if val:
                setattr(ps, field_name, int(val) / _TWIPS_PER_PT)

    @staticmethod
    def _apply_section_type(sect_pr: ET.Element, ps: ldm.PageSetup) -> None:
        type_elem = sect_pr.find(f"{W_NS}type")
        if type_elem is not None:
            ps.section_start = _SECTION_START_MAP.get(
                type_elem.get(f"{W_NS}val", ""), _DEFAULT_SECTION_START
            )

    @staticmethod
    def _apply_title_page(sect_pr: ET.Element, ps: ldm.PageSetup) -> None:
        if sect_pr.find(f"{W_NS}titlePg") is not None:
            ps.different_first_page_header_footer = True

    @staticmethod
    def _apply_page_numbering(sect_pr: ET.Element, ps: ldm.PageSetup) -> None:
        pgnum = sect_pr.find(f"{W_NS}pgNumType")
        if pgnum is None:
            return
        start = pgnum.get(f"{W_NS}start")
        if start is not None:
            ps.page_starting_number = int(start)
            ps.restart_page_numbering = True

    @staticmethod
    def _apply_columns(sect_pr: ET.Element, ps: ldm.PageSetup) -> None:
        cols = sect_pr.find(f"{W_NS}cols")
        if cols is None:
            return
        tc = ldm.TextColumns()
        num = cols.get(f"{W_NS}num")
        if num:
            tc.count = int(num)
        space = cols.get(f"{W_NS}space")
        if space:
            tc.spacing = int(space) / _TWIPS_PER_PT
        eq = cols.get(f"{W_NS}equalWidth")
        if eq is not None:
            tc.evenly_spaced = eq not in ("0", "false")
        sep = cols.get(f"{W_NS}sep")
        if sep is not None:
            tc.line_between = sep in ("1", "true")
        for col_el in cols.findall(f"{W_NS}col"):
            col = ldm.TextColumn()
            w = col_el.get(f"{W_NS}w")
            if w:
                col.width = int(w) / _TWIPS_PER_PT
            sp = col_el.get(f"{W_NS}space")
            if sp:
                col.space_after = int(sp) / _TWIPS_PER_PT
            tc.columns.append(col)
        ps.text_columns = tc


class SectionBuilder:
    """Split the body into :class:`ldm.Section`\\ s at sectPr boundaries."""

    def __init__(
        self,
        ctx: ReaderContext,
        paragraph_builder: ParagraphBuilder,
        table_builder: TableBuilderFn,
        page_setup_builder: PageSetupBuilder,
    ):
        self._ctx = ctx
        self._paragraphs = paragraph_builder
        self._build_table = table_builder
        self._page_setup = page_setup_builder

    def build_all(self) -> list[ldm.Section]:
        """Walk the body and split into Sections at every ``<w:sectPr>``."""
        ctx = self._ctx
        if ctx._document_xml is None:
            return []
        body = ctx._document_xml.find(f"{W_NS}body")
        if body is None:
            return []

        self._cache_first_section_page_setup(body)
        ctx._first_body_page_active = True  # type: ignore[attr-defined]

        sections: list[ldm.Section] = []
        current: list[_BodyChild] = []
        for element in ctx._resolve_body_children(body):
            tag = element.tag
            if tag == f"{W_NS}p":
                self._consume_paragraph(element, current, sections)
            elif tag == f"{W_NS}tbl":
                current.append(self._build_table(element))
            elif tag == f"{W_NS}sectPr":
                sections.append(self._make_section(element, current))
                current = []

        if current:
            trailing = ldm.Section()
            trailing.body = ldm.Body(children=current)
            sections.append(trailing)
        return sections

    def _cache_first_section_page_setup(self, body: ET.Element) -> None:
        ctx = self._ctx
        ctx._current_page_setup = None
        first = next(iter(body.iter(f"{W_NS}sectPr")), None)
        if first is not None:
            ctx._current_page_setup = self._page_setup.build(first)

    def _consume_paragraph(
        self,
        element: ET.Element,
        current: list[_BodyChild],
        sections: list[ldm.Section],
    ) -> None:
        para = self._paragraphs.build(element)
        current.append(para)
        self._update_first_page_state(para)
        sect_pr = self._paragraph_section_break(element)
        if sect_pr is not None:
            sections.append(self._make_section(sect_pr, current))
            current.clear()

    def _update_first_page_state(self, para: ldm.Paragraph) -> None:
        if not getattr(self._ctx, "_first_body_page_active", False):
            return
        if ParagraphBuilder.paragraph_ends_first_page(para):
            self._ctx._first_body_page_active = False  # type: ignore[attr-defined]

    @staticmethod
    def _paragraph_section_break(p_elem: ET.Element) -> Optional[ET.Element]:
        pPr = p_elem.find(f"{W_NS}pPr")
        if pPr is None:
            return None
        return pPr.find(f"{W_NS}sectPr")

    def _make_section(
        self,
        sect_pr: ET.Element,
        children: list[_BodyChild],
    ) -> ldm.Section:
        sec = ldm.Section()
        sec.page_setup = self._page_setup.build(sect_pr)
        sec.body = ldm.Body(children=list(children))
        return sec
