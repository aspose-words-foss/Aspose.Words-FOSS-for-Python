"""Builders for the document's style and list definitions.

:class:`StyleBuilder` translates ``<w:styles>`` into a list of
``ldm.Style`` objects, applying the cascade resolvers for paragraph
styles so the resulting formats are fully merged.

:class:`ListBuilder` translates ``<w:numbering>`` (abstractNum + num
elements with their lvlOverride children) into a list of
``ldm.DocList`` objects.  Both abstract levels and override-inline
levels share one private builder method so the two ListLevel paths
stay in sync.
"""

import re
from typing import Optional
from xml.etree import ElementTree as ET

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_reader.constants import (
    W_NS,
    _ALIGNMENT_MAP,
    _NUMBER_STYLE_MAP,
    _STYLE_TYPE_MAP,
    _TWIPS_PER_PT,
)
from aspose.words_foss.docx_reader.field_mappings import LIST_TRAILING_CHARACTER_MAP
from aspose.words_foss.docx_reader.utils import _canonicalize_style_name
from aspose.words_foss.model.style_identifiers import resolve_style_identifier

from ._context import ReaderContext
from ._helpers import apply_padding_sides, build_borders, parse_int
from .cascading import (
    FontBuilder,
    FontResolver,
    ParagraphFormatBuilder,
    ParagraphFormatResolver,
)


_STYLE_TYPE_PARAGRAPH = 1
_STYLE_TYPE_TABLE = 3
_BUILTIN_HEADING_MAX = 9


class StyleBuilder:
    """Translate ``<w:styles>`` into a list of :class:`ldm.Style`."""

    def __init__(
        self,
        ctx: ReaderContext,
        fonts: FontBuilder,
        font_resolver: FontResolver,
        pfs: ParagraphFormatBuilder,
        pf_resolver: ParagraphFormatResolver,
    ):
        self._ctx = ctx
        self._fonts = fonts
        self._font_resolver = font_resolver
        self._pfs = pfs
        self._pf_resolver = pf_resolver

    def build_all(self) -> list[ldm.Style]:
        """Return one :class:`ldm.Style` per ``<w:style>`` in ``styles.xml``."""
        styles_xml = self._ctx._styles_xml
        if styles_xml is None:
            # Synthesize the default Normal that the writer always emits,
            # so docs lacking styles.xml round-trip symmetrically.
            synth = ET.fromstring(
                f'<w:style xmlns:w="{W_NS.strip("{}")}" '
                'w:type="paragraph" w:default="1" w:styleId="Normal">'
                '<w:name w:val="Normal"/></w:style>'
            )
            return [self._build_one(synth)]
        return [self._build_one(elem) for elem in styles_xml.findall(f"{W_NS}style")]

    def _build_one(self, style_elem: ET.Element) -> ldm.Style:
        s = ldm.Style()
        is_custom = style_elem.get(f"{W_NS}customStyle", "") == "1"
        xml_style_id = style_elem.get(f"{W_NS}styleId", "")

        self._apply_name(style_elem, s, is_custom)
        s.type = _STYLE_TYPE_MAP.get(style_elem.get(f"{W_NS}type", ""), 0)
        self._apply_identifier(s, xml_style_id, is_custom)
        self._apply_priority(style_elem, s)
        heading_match = self._apply_heading_flag(s)
        self._apply_relations(style_elem, s)
        self._apply_format(style_elem, s, xml_style_id, heading_match)
        self._apply_table_style_format(style_elem, s)
        return s

    @staticmethod
    def _apply_name(style_elem: ET.Element, s: ldm.Style, is_custom: bool) -> None:
        name_elem = style_elem.find(f"{W_NS}name")
        if name_elem is None:
            # Real-world DOCX files occasionally declare a style with
            # a ``<w:styleId>`` but no ``<w:name>``.  Fall back to the
            # styleId so downstream code (writer's style_id_map,
            # tblStyle lookup, etc.) has a stable handle that survives
            # a round-trip — otherwise the writer would emit the style
            # under a synthesised ``_StyleN`` id while body references
            # like ``<w:tblStyle w:val="a"/>`` still pointed at the
            # original ``"a"`` token, leaving a dangling reference.
            s.name = style_elem.get(f"{W_NS}styleId", "")
            return
        raw = name_elem.get(f"{W_NS}val", "")
        s.name = raw if is_custom else _canonicalize_style_name(raw)

    @staticmethod
    def _apply_identifier(s: ldm.Style, xml_style_id: str, is_custom: bool) -> None:
        sid_int = resolve_style_identifier(xml_style_id, s.name)
        resolved = sid_int >= 0
        s.style_identifier = sid_int if resolved else 0
        s.built_in = resolved and not is_custom

    @staticmethod
    def _apply_priority(style_elem: ET.Element, s: ldm.Style) -> None:
        ui_pri = style_elem.find(f"{W_NS}uiPriority")
        if ui_pri is not None:
            s.priority = parse_int(ui_pri.get(f"{W_NS}val", "99"), 99)

    @staticmethod
    def _apply_heading_flag(s: ldm.Style) -> Optional[re.Match[str]]:
        if 1 <= s.style_identifier <= _BUILTIN_HEADING_MAX:
            s.is_heading = True
            return None
        match = re.search(r"[Hh]eading\s*(\d+)", s.name)
        s.is_heading = match is not None
        return match

    def _apply_relations(self, style_elem: ET.Element, s: ldm.Style) -> None:
        based_on = style_elem.find(f"{W_NS}basedOn")
        if based_on is not None:
            s.base_style_name = self._ctx._resolve_style_name(
                based_on.get(f"{W_NS}val", "")
            )
        next_style = style_elem.find(f"{W_NS}next")
        if next_style is not None:
            s.next_paragraph_style_name = self._ctx._resolve_style_name(
                next_style.get(f"{W_NS}val", "")
            )
        else:
            # Word's implicit default: self-reference.
            s.next_paragraph_style_name = s.name

    def _apply_format(
        self,
        style_elem: ET.Element,
        s: ldm.Style,
        style_id: str,
        heading_match: Optional[re.Match[str]],
    ) -> None:
        pPr = style_elem.find(f"{W_NS}pPr")
        rPr = style_elem.find(f"{W_NS}rPr")
        if s.type == _STYLE_TYPE_PARAGRAPH:
            s.paragraph_format = self._pf_resolver.resolve(pPr, style_id)
            s.paragraph_format.style_name = s.name
            s.paragraph_format.style_identifier = s.style_identifier
            self._apply_heading_outline(s.paragraph_format, s, heading_match)
            s.font = self._font_resolver.resolve(rPr, style_id)
            return
        if pPr is not None:
            s.paragraph_format = self._pfs.build(pPr)
            self._apply_heading_outline(s.paragraph_format, s, heading_match)
        if rPr is not None:
            s.font = self._fonts.build(rPr)

    @staticmethod
    def _apply_heading_outline(
        pf: ldm.ParagraphFormat,
        s: ldm.Style,
        heading_match: Optional[re.Match[str]],
    ) -> None:
        if not s.is_heading:
            return
        pf.is_heading = True
        if 1 <= s.style_identifier <= _BUILTIN_HEADING_MAX:
            pf.outline_level = s.style_identifier - 1
        elif heading_match:
            pf.outline_level = int(heading_match.group(1)) - 1

    @staticmethod
    def _apply_table_style_format(style_elem: ET.Element, s: ldm.Style) -> None:
        if s.type != _STYLE_TYPE_TABLE:
            return
        tblPr = style_elem.find(f"{W_NS}tblPr")
        if tblPr is None:
            return
        tsf = ldm.TableStyleFormat()
        tblBorders = tblPr.find(f"{W_NS}tblBorders")
        if tblBorders is not None:
            tsf.borders = build_borders(tblBorders)
        tblCellMar = tblPr.find(f"{W_NS}tblCellMar")
        if tblCellMar is not None:
            apply_padding_sides(tsf, tblCellMar)
        has_borders = any(
            b.line_style != 0 or b.line_width != 0.0 for b in tsf.borders
        )
        has_padding = any(
            (tsf.left_padding, tsf.right_padding, tsf.top_padding, tsf.bottom_padding)
        )
        if not (has_borders or has_padding):
            return
        if not has_borders:
            tsf.borders = []
        s.table_style_format = tsf


class ListBuilder:
    """Translate ``<w:numbering>`` into a list of :class:`ldm.DocList`."""

    def __init__(self, ctx: ReaderContext, fonts: FontBuilder):
        self._ctx = ctx
        self._fonts = fonts

    def build_all(self) -> list[ldm.DocList]:
        """Return one :class:`ldm.DocList` per ``<w:num>`` in ``numbering.xml``."""
        numbering_xml = self._ctx._numbering_xml
        if numbering_xml is None:
            return []
        abstract_defs, abstract_multi = self._parse_abstract_definitions(numbering_xml)
        return self._parse_concrete_numbers(numbering_xml, abstract_defs, abstract_multi)

    def _parse_abstract_definitions(
        self, numbering_xml: ET.Element
    ) -> tuple[dict[int, list[ldm.ListLevel]], dict[int, bool]]:
        abstract_defs: dict[int, list[ldm.ListLevel]] = {}
        abstract_multi: dict[int, bool] = {}
        for abstract in numbering_xml.findall(f".//{W_NS}abstractNum"):
            abs_id = parse_int(abstract.get(f"{W_NS}abstractNumId", "0"))
            levels = [self._build_level(lvl) for lvl in abstract.findall(f"{W_NS}lvl")]
            abstract_defs[abs_id] = levels
            # ``<w:multiLevelType w:val="...">`` is a child element, not
            # an attribute on ``<w:abstractNum>`` — use ``find()`` and
            # read ``w:val`` off the matched element.  The previous
            # ``abstract.get(...)`` form always returned ``""`` so every
            # multilevel / hybridMultilevel list was misclassified as
            # singleLevel, which then made the writer emit
            # ``multiLevelType="singleLevel"`` alongside nine
            # ``<w:lvl>`` rows — an Aspose-rejected schema violation.
            #
            # When ``<w:multiLevelType>`` is absent (``mlt_val == ""``)
            # the OOXML spec leaves the type implementation-defined;
            # Word treats it as multilevel whenever more than one
            # ``<w:lvl>`` row is present, and we follow suit so the
            # writer doesn't downgrade these to singleLevel and trip
            # the same schema violation on the next round-trip.
            mlt_elem = abstract.find(f"{W_NS}multiLevelType")
            mlt_val = mlt_elem.get(f"{W_NS}val", "") if mlt_elem is not None else ""
            abstract_multi[abs_id] = (
                mlt_val in ("multilevel", "hybridMultilevel")
                or (mlt_val == "" and len(levels) > 1)
            )
        return abstract_defs, abstract_multi

    def _build_level(self, lvl: ET.Element) -> ldm.ListLevel:
        ll = ldm.ListLevel()
        self._apply_level_format(lvl, ll)
        self._apply_level_indent(lvl, ll)
        self._apply_level_misc(lvl, ll)
        return ll

    @staticmethod
    def _apply_level_format(lvl: ET.Element, ll: ldm.ListLevel) -> None:
        num_fmt = lvl.find(f"{W_NS}numFmt")
        token = num_fmt.get(f"{W_NS}val", "bullet") if num_fmt is not None else "bullet"
        ll.number_style = _NUMBER_STYLE_MAP.get(token, 0)
        lvl_text = lvl.find(f"{W_NS}lvlText")
        ll.number_format = lvl_text.get(f"{W_NS}val", "") if lvl_text is not None else ""
        start = lvl.find(f"{W_NS}start")
        if start is not None:
            ll.start_at = parse_int(start.get(f"{W_NS}val", "1"), 1)
        lvl_jc = lvl.find(f"{W_NS}lvlJc")
        if lvl_jc is not None:
            ll.alignment = _ALIGNMENT_MAP.get(lvl_jc.get(f"{W_NS}val", "left"), 0)

    @staticmethod
    def _apply_level_indent(lvl: ET.Element, ll: ldm.ListLevel) -> None:
        pPr = lvl.find(f"{W_NS}pPr")
        if pPr is None:
            return
        ind = pPr.find(f"{W_NS}ind")
        if ind is None:
            return
        left = ind.get(f"{W_NS}left", "")
        hanging = ind.get(f"{W_NS}hanging", "")
        first_line = ind.get(f"{W_NS}firstLine", "")
        if left:
            ll.number_position = int(left) / _TWIPS_PER_PT
        if left and hanging:
            ll.text_position = int(left) / _TWIPS_PER_PT
            ll.number_position = (int(left) - int(hanging)) / _TWIPS_PER_PT
        elif first_line:
            ll.text_position = int(left) / _TWIPS_PER_PT if left else 0.0

    def _apply_level_misc(self, lvl: ET.Element, ll: ldm.ListLevel) -> None:
        lvl_rPr = lvl.find(f"{W_NS}rPr")
        if lvl_rPr is not None:
            ll.font = self._fonts.build(lvl_rPr)
        lvl_restart = lvl.find(f"{W_NS}lvlRestart")
        if lvl_restart is not None:
            ll.restart_after_level = parse_int(lvl_restart.get(f"{W_NS}val", "-1"), -1)
        suff = lvl.find(f"{W_NS}suff")
        if suff is not None:
            ll.trailing_character = LIST_TRAILING_CHARACTER_MAP.get(
                suff.get(f"{W_NS}val", "tab"), 0
            )

    def _parse_concrete_numbers(
        self,
        numbering_xml: ET.Element,
        abstract_defs: dict[int, list[ldm.ListLevel]],
        abstract_multi: dict[int, bool],
    ) -> list[ldm.DocList]:
        lists: list[ldm.DocList] = []
        for num in numbering_xml.findall(f".//{W_NS}num"):
            num_id = parse_int(num.get(f"{W_NS}numId", "0"))
            abs_id_elem = num.find(f"{W_NS}abstractNumId")
            if abs_id_elem is None:
                continue
            abs_id = parse_int(abs_id_elem.get(f"{W_NS}val", "0"))
            dl = ldm.DocList()
            dl.list_id = num_id
            dl.levels = list(abstract_defs.get(abs_id, []))
            dl.is_multi_level = abstract_multi.get(abs_id, False)
            for ov_elem in num.findall(f"{W_NS}lvlOverride"):
                dl.overrides.append(self._build_override(ov_elem))
            lists.append(dl)
        return lists

    def _build_override(self, ov_elem: ET.Element) -> ldm.ListLevelOverride:
        ov = ldm.ListLevelOverride()
        ov.ilvl = parse_int(ov_elem.get(f"{W_NS}ilvl", "0"))
        start_ov = ov_elem.find(f"{W_NS}startOverride")
        if start_ov is not None:
            ov.is_start_at = True
            ov.start_at_raw = parse_int(start_ov.get(f"{W_NS}val", "1"), 1)
        inner_lvl = ov_elem.find(f"{W_NS}lvl")
        if inner_lvl is not None:
            ov.is_formatting = True
            ov.list_level = self._build_level(inner_lvl)
        return ov
