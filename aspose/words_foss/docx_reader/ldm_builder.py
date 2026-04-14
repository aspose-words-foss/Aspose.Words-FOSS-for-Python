"""
Light Document Model builder for DOCX documents.

Provides the LDM construction methods as a mixin for DocumentReader.
Handles building styles, lists, sections, paragraphs, runs, tables,
fonts, borders, and shading from parsed DOCX XML.
"""

from __future__ import annotations

import re
from typing import Optional, Iterator, TYPE_CHECKING
from xml.etree import ElementTree as ET

from aspose.words_foss.docx_reader.constants import (
    R_NS,
    W_NS,
    _ALIGNMENT_MAP,
    _BORDER_STYLE_MAP,
    _BORDER_SIZE_DIVISOR,
    _DEFAULT_TAB_STOP_PT,
    _HIGHLIGHT_COLOR_MAP,
    _LINE_RULE_MAP,
    _NUMBER_STYLE_MAP,
    _OUTLINE_LEVEL_BODY,
    _PAPER_A3,
    _PAPER_A4,
    _PAPER_CUSTOM,
    _PAPER_LEGAL,
    _PAPER_LETTER,
    _PAPER_SIZE_TOLERANCE_A_PT,
    _PAPER_SIZE_TOLERANCE_PT,
    _A3_HEIGHT_PT,
    _A3_WIDTH_PT,
    _A4_HEIGHT_PT,
    _A4_WIDTH_PT,
    _HALF_PT_DIVISOR,
    _LEGAL_HEIGHT_PT,
    _LETTER_HEIGHT_PT,
    _LETTER_WIDTH_PT,
    _PCT_DIVISOR,
    _SECTION_START_MAP,
    _STYLE_TYPE_MAP,
    _TWIPS_PER_PT,
    _UNDERLINE_MAP,
    COLOR_EMPTY,
    PAGE_FIELD_SENTINEL,
)
from aspose.words_foss.docx_reader.utils import (
    _apply_theme_color_modifiers,
    _canonicalize_style_name,
    _collect_run_text,
    _empty_borders,
    _hex_to_aspose_color,
)
from aspose.words_foss.model.enums import Orientation as _Orient

if TYPE_CHECKING:
    from aspose.words_foss import light_document_model as ldm


class LdmBuilderMixin:
    """Mixin providing Light Document Model construction methods."""

    # These attributes are defined on DocumentReader but referenced here.
    _document_xml: Optional[ET.Element]
    _numbering_xml: Optional[ET.Element]
    _styles_xml: Optional[ET.Element]
    _settings_xml: Optional[ET.Element]
    _numbering_cache: dict[int, "NumberingInfo"]
    _rels: dict[str, str]
    _style_id_to_name: dict[str, str]
    _doc_image_rels: dict[str, str]
    _header_data: list[tuple[ET.Element, dict[str, str]]]
    _footer_data: list[tuple[ET.Element, dict[str, str]]]
    _theme_fonts: dict[str, str]
    _theme_colors: dict[str, str]
    _doc_default_rPr: Optional[ET.Element]
    _doc_default_pPr: Optional[ET.Element]
    _current_page_setup: "Optional[ldm.PageSetup]"
    _style_elem_cache: dict[str, ET.Element]
    _name_to_style_id: dict[str, str]

    # =========================================================================
    # LIGHT DOCUMENT MODEL BUILDER — entry point
    # =========================================================================

    def _get_default_tab_stop(self) -> float:
        """Read default tab stop from settings.xml (twips → points)."""
        if self._settings_xml is not None:
            elem = self._settings_xml.find(f"{W_NS}defaultTabStop")
            if elem is not None:
                val = elem.get(f"{W_NS}val", "")
                if val:
                    try:
                        return int(val) / _TWIPS_PER_PT
                    except ValueError:
                        pass
        return _DEFAULT_TAB_STOP_PT

    def _get_page_color(self) -> str:
        """Read document background color from document.xml."""
        if self._document_xml is not None:
            bg = self._document_xml.find(f"{W_NS}background")
            if bg is not None:
                color = bg.get(f"{W_NS}color", "")
                if color:
                    return _hex_to_aspose_color(color)
        return COLOR_EMPTY

    def _build_resolved_font(self, rPr: Optional[ET.Element], style_id: str = "") -> "ldm.Font":
        """Build a Font resolving through style chain and docDefaults.

        Resolution order: docDefaults → style chain → direct rPr.
        """
        from aspose.words_foss import light_document_model as ldm

        # Start with docDefaults
        base_font = ldm.Font(color=COLOR_EMPTY, highlight_color=COLOR_EMPTY)
        if self._doc_default_rPr is not None:
            base_font = self._build_ldm_font(self._doc_default_rPr)
            if not base_font.color:
                base_font.color = COLOR_EMPTY
            if not base_font.highlight_color:
                base_font.highlight_color = COLOR_EMPTY

        # Resolve through style chain
        if style_id:
            chain = self._get_style_chain(style_id)
            for sid in chain:
                style_elem = self._find_style_elem(sid)
                if style_elem is not None:
                    srPr = style_elem.find(f"{W_NS}rPr")
                    if srPr is not None:
                        self._merge_font(base_font, self._build_ldm_font(srPr))

        # Apply direct formatting
        if rPr is not None:
            self._merge_font(base_font, self._build_ldm_font(rPr))

        # Set style_name for the character style if not already set
        if not base_font.style_name:
            base_font.style_name = "Default Paragraph Font"

        # Ensure color fields are never empty
        if not base_font.color:
            base_font.color = COLOR_EMPTY
        if not base_font.highlight_color:
            base_font.highlight_color = COLOR_EMPTY

        return base_font

    def _build_resolved_paragraph_format(
        self, pPr: Optional[ET.Element], style_id: str = ""
    ) -> "ldm.ParagraphFormat":
        """Build a ParagraphFormat resolving through style chain and docDefaults."""
        from aspose.words_foss import light_document_model as ldm

        # Start with docDefaults
        base_pf = ldm.ParagraphFormat(borders=_empty_borders())
        if self._doc_default_pPr is not None:
            base_pf = self._build_ldm_paragraph_format(self._doc_default_pPr)

        # Resolve through style chain
        if style_id:
            chain = self._get_style_chain(style_id)
            for sid in chain:
                style_elem = self._find_style_elem(sid)
                if style_elem is not None:
                    spPr = style_elem.find(f"{W_NS}pPr")
                    if spPr is not None:
                        self._merge_pf(base_pf, self._build_ldm_paragraph_format(spPr))

        # Apply direct formatting
        if pPr is not None:
            self._merge_pf(base_pf, self._build_ldm_paragraph_format(pPr))

        # Set style_name to the paragraph style if not already resolved
        if style_id and not base_pf.style_name:
            base_pf.style_name = self._resolve_style_name(style_id)

        # Ensure borders always present
        if not base_pf.borders:
            base_pf.borders = _empty_borders()

        return base_pf

    def _get_style_chain(self, style_id_or_name: str) -> list[str]:
        """Get the style inheritance chain (from base to most derived).

        Returns a list of style IDs from the root ancestor to the given style.
        """
        if self._styles_xml is None:
            return []
        # First resolve to style ID
        style_id = self._name_to_style_id.get(style_id_or_name, style_id_or_name)

        chain: list[str] = []
        visited: set[str] = set()
        current = style_id
        while current and current not in visited:
            visited.add(current)
            chain.append(current)
            elem = self._find_style_elem(current)
            if elem is None:
                break
            based_on = elem.find(f"{W_NS}basedOn")
            if based_on is not None:
                current = based_on.get(f"{W_NS}val", "")
            else:
                break
        chain.reverse()  # base → derived
        return chain

    def _resolve_numPr_from_style(self, style_id: str) -> Optional[ET.Element]:
        """Walk the style chain to find a numPr element."""
        chain = self._get_style_chain(style_id)
        # Walk from most derived to base
        for sid in reversed(chain):
            elem = self._find_style_elem(sid)
            if elem is not None:
                pPr = elem.find(f"{W_NS}pPr")
                if pPr is not None:
                    numPr = pPr.find(f"{W_NS}numPr")
                    if numPr is not None:
                        return numPr
        return None

    def _find_style_elem(self, style_id: str) -> Optional[ET.Element]:
        """Find a <w:style> element by its styleId (O(1) cached lookup)."""
        return self._style_elem_cache.get(style_id)

    @staticmethod
    def _merge_font(base: "ldm.Font", override: "ldm.Font") -> None:
        """Merge override font properties into base.

        Uses Pydantic's ``model_fields_set`` to decide which properties were
        explicitly set in the override.  This correctly handles falsy values
        such as ``bold=False`` overriding ``bold=True``.
        """
        _set: set[str] = override.model_fields_set
        if "name" in _set:
            base.name = override.name
        if "size" in _set:
            base.size = override.size
        if "bold" in _set:
            base.bold = override.bold
        if "italic" in _set:
            base.italic = override.italic
        if "underline" in _set:
            base.underline = override.underline
        if "color" in _set and override.color != COLOR_EMPTY:
            base.color = override.color
        if "strike_through" in _set:
            base.strike_through = override.strike_through
        if "superscript" in _set:
            base.superscript = override.superscript
        if "subscript" in _set:
            base.subscript = override.subscript
        if "highlight_color" in _set and override.highlight_color != COLOR_EMPTY:
            base.highlight_color = override.highlight_color
        if "all_caps" in _set:
            base.all_caps = override.all_caps
        if "small_caps" in _set:
            base.small_caps = override.small_caps
        if "hidden" in _set:
            base.hidden = override.hidden
        if "style_name" in _set:
            base.style_name = override.style_name

    @staticmethod
    def _merge_pf(base: "ldm.ParagraphFormat", override: "ldm.ParagraphFormat") -> None:
        """Merge override paragraph format into base.

        Uses Pydantic's ``model_fields_set`` to decide which properties were
        explicitly set.  This correctly handles zero-valued overrides
        (e.g. ``space_after=0.0``) and boolean unsets
        (e.g. ``keep_with_next=False``).
        """
        _set: set[str] = override.model_fields_set
        if "style_name" in _set:
            base.style_name = override.style_name
        if "alignment" in _set:
            base.alignment = override.alignment
        if "left_indent" in _set:
            base.left_indent = override.left_indent
        if "right_indent" in _set:
            base.right_indent = override.right_indent
        if "first_line_indent" in _set:
            base.first_line_indent = override.first_line_indent
        if "space_before" in _set:
            base.space_before = override.space_before
        if "space_after" in _set:
            base.space_after = override.space_after
        if "space_before_auto" in _set:
            base.space_before_auto = override.space_before_auto
        if "space_after_auto" in _set:
            base.space_after_auto = override.space_after_auto
        if "line_spacing" in _set:
            base.line_spacing = override.line_spacing
        if "line_spacing_rule" in _set:
            base.line_spacing_rule = override.line_spacing_rule
        if "keep_with_next" in _set:
            base.keep_with_next = override.keep_with_next
        if "page_break_before" in _set:
            base.page_break_before = override.page_break_before
        if "outline_level" in _set:
            base.outline_level = override.outline_level
        if "is_heading" in _set:
            base.is_heading = override.is_heading
        if "is_list_item" in _set:
            base.is_list_item = override.is_list_item

    def to_light_document(self) -> "ldm.Document":
        """Build a light_document_model.Document from the loaded DOCX.

        Returns a fully populated Pydantic model representing the document
        structure including styles, lists, sections, paragraphs, runs,
        tables, and page setup.
        """
        from aspose.words_foss import light_document_model as ldm

        doc = ldm.Document()
        doc.page_color = self._get_page_color()
        doc.default_tab_stop = self._get_default_tab_stop()

        # Build reverse name→ID map for style chain resolution.
        # Store both raw and canonicalized names so lookups work
        # regardless of which form a basedOn reference uses.
        self._name_to_style_id: dict[str, str] = {}
        if self._styles_xml is not None:
            for style_elem in self._styles_xml.findall(f"{W_NS}style"):
                sid = style_elem.get(f"{W_NS}styleId", "")
                name_elem = style_elem.find(f"{W_NS}name")
                if name_elem is not None and sid:
                    raw_name = name_elem.get(f"{W_NS}val", "")
                    self._name_to_style_id[raw_name] = sid
                    is_custom = style_elem.get(f"{W_NS}customStyle", "") == "1"
                    if not is_custom:
                        canonical = _canonicalize_style_name(raw_name)
                        if canonical != raw_name:
                            self._name_to_style_id[canonical] = sid

        # Parse styles
        if self._styles_xml is not None:
            doc.styles = self._build_ldm_styles()

        # Parse lists
        if self._numbering_xml is not None:
            doc.lists = self._build_ldm_lists()

        # Parse sections (body + page setup)
        doc.sections = self._build_ldm_sections()

        # Parse header paragraphs — only direct <w:p> children of the root,
        # not paragraphs nested inside table cells within the header.
        # ``_anchor_y_base_mode`` tells :meth:`_anchor_page_origin_mm` to
        # resolve ``paragraph``-relative offsets against the header /
        # footer band instead of the body-top margin so anchored logos
        # land inside the header band on every page.
        hdr_paras: list[ldm.Paragraph] = []
        self._anchor_y_base_mode = "header"  # type: ignore[attr-defined]
        try:
            for hdr_xml, hdr_rels in self._header_data:
                for p_elem in hdr_xml:
                    if p_elem.tag == f"{W_NS}p":
                        hdr_paras.append(self._build_ldm_paragraph(p_elem, hdr_rels))
        finally:
            self._anchor_y_base_mode = "body"  # type: ignore[attr-defined]
        doc.header_paragraphs = hdr_paras

        # Parse footer paragraphs — same rule: top-level paragraphs only.
        ftr_paras: list[ldm.Paragraph] = []
        self._anchor_y_base_mode = "footer"  # type: ignore[attr-defined]
        try:
            for ftr_xml, ftr_rels in self._footer_data:
                for p_elem in ftr_xml:
                    if p_elem.tag == f"{W_NS}p":
                        ftr_paras.append(self._build_ldm_paragraph(p_elem, ftr_rels))
        finally:
            self._anchor_y_base_mode = "body"  # type: ignore[attr-defined]
        doc.footer_paragraphs = ftr_paras

        return doc

    # -- Styles ---------------------------------------------------------------

    def _build_ldm_styles(self) -> "list[ldm.Style]":
        from aspose.words_foss import light_document_model as ldm

        styles: list[ldm.Style] = []
        if self._styles_xml is None:
            return styles

        for style_elem in self._styles_xml.findall(f"{W_NS}style"):
            s = ldm.Style()
            # Name (resolve built-in names to canonical form)
            is_custom = style_elem.get(f"{W_NS}customStyle", "") == "1"
            name_elem = style_elem.find(f"{W_NS}name")
            if name_elem is not None:
                raw_name = name_elem.get(f"{W_NS}val", "")
                s.name = raw_name if is_custom else _canonicalize_style_name(raw_name)

            # Type
            st = style_elem.get(f"{W_NS}type", "")
            s.type = _STYLE_TYPE_MAP.get(st, 0)

            # Is heading
            heading_match = re.search(r"[Hh]eading\s*(\d+)", s.name)
            s.is_heading = heading_match is not None

            # Base style (resolve ID to display name)
            based_on = style_elem.find(f"{W_NS}basedOn")
            if based_on is not None:
                base_id = based_on.get(f"{W_NS}val", "")
                s.base_style_name = self._resolve_style_name(base_id)

            # Next style (resolve ID to display name; default to self)
            next_style = style_elem.find(f"{W_NS}next")
            if next_style is not None:
                next_id = next_style.get(f"{W_NS}val", "")
                s.next_paragraph_style_name = self._resolve_style_name(next_id)
            else:
                # Aspose.Words defaults to self-referencing the style's own name
                s.next_paragraph_style_name = s.name

            # Paragraph format (resolved through basedOn chain + docDefaults)
            style_id = style_elem.get(f"{W_NS}styleId", "")
            if s.type == 1:  # paragraph style
                s.paragraph_format = self._build_resolved_paragraph_format(
                    style_elem.find(f"{W_NS}pPr"), style_id
                )
                s.paragraph_format.style_name = s.name
                if s.is_heading and heading_match:
                    s.paragraph_format.is_heading = True
                    s.paragraph_format.outline_level = int(heading_match.group(1)) - 1
                # Resolved font for paragraph style
                s.font = self._build_resolved_font(style_elem.find(f"{W_NS}rPr"), style_id)
            else:
                pPr = style_elem.find(f"{W_NS}pPr")
                if pPr is not None:
                    s.paragraph_format = self._build_ldm_paragraph_format(pPr)
                    if s.is_heading and heading_match:
                        s.paragraph_format.is_heading = True
                        s.paragraph_format.outline_level = int(heading_match.group(1)) - 1
                rPr = style_elem.find(f"{W_NS}rPr")
                if rPr is not None:
                    s.font = self._build_ldm_font(rPr)

            styles.append(s)

        return styles

    # -- Lists ----------------------------------------------------------------

    def _build_ldm_lists(self) -> "list[ldm.DocList]":
        from aspose.words_foss import light_document_model as ldm

        lists: list[ldm.DocList] = []
        if self._numbering_xml is None:
            return lists

        # Parse abstract numbering definitions
        abstract_defs: dict[int, list[ldm.ListLevel]] = {}
        abstract_multi: dict[int, bool] = {}

        for abstract in self._numbering_xml.findall(f".//{W_NS}abstractNum"):
            abs_id = int(abstract.get(f"{W_NS}abstractNumId", "0"))
            levels: list[ldm.ListLevel] = []
            multi = abstract.get(f"{W_NS}multiLevelType", "")

            for lvl in abstract.findall(f"{W_NS}lvl"):
                ll = ldm.ListLevel()

                num_fmt_elem = lvl.find(f"{W_NS}numFmt")
                num_fmt = (
                    num_fmt_elem.get(f"{W_NS}val", "bullet")
                    if num_fmt_elem is not None
                    else "bullet"
                )
                ll.number_style = _NUMBER_STYLE_MAP.get(num_fmt, 0)

                lvl_text_elem = lvl.find(f"{W_NS}lvlText")
                ll.number_format = (
                    lvl_text_elem.get(f"{W_NS}val", "") if lvl_text_elem is not None else ""
                )

                start_elem = lvl.find(f"{W_NS}start")
                if start_elem is not None:
                    ll.start_at = int(start_elem.get(f"{W_NS}val", "1"))

                # Alignment
                lvl_jc = lvl.find(f"{W_NS}lvlJc")
                if lvl_jc is not None:
                    ll.alignment = _ALIGNMENT_MAP.get(lvl_jc.get(f"{W_NS}val", "left"), 0)

                # Indentation
                pPr = lvl.find(f"{W_NS}pPr")
                if pPr is not None:
                    ind = pPr.find(f"{W_NS}ind")
                    if ind is not None:
                        left = ind.get(f"{W_NS}left", "")
                        if left:
                            ll.number_position = int(left) / _TWIPS_PER_PT
                        hanging = ind.get(f"{W_NS}hanging", "")
                        first_line = ind.get(f"{W_NS}firstLine", "")
                        if left and hanging:
                            ll.text_position = int(left) / _TWIPS_PER_PT
                            ll.number_position = (int(left) - int(hanging)) / _TWIPS_PER_PT
                        elif first_line:
                            ll.text_position = int(left) / _TWIPS_PER_PT if left else 0.0

                levels.append(ll)

            abstract_defs[abs_id] = levels
            abstract_multi[abs_id] = multi in (
                "multilevel",
                "hybridMultilevel",
            )

        # Map num -> abstract
        for num in self._numbering_xml.findall(f".//{W_NS}num"):
            num_id = int(num.get(f"{W_NS}numId", "0"))
            abs_id_elem = num.find(f"{W_NS}abstractNumId")
            if abs_id_elem is None:
                continue
            abs_id = int(abs_id_elem.get(f"{W_NS}val", "0"))

            dl = ldm.DocList()
            dl.list_id = num_id
            dl.levels = list(abstract_defs.get(abs_id, []))
            dl.is_multi_level = abstract_multi.get(abs_id, False)
            lists.append(dl)

        return lists

    # -- Sections -------------------------------------------------------------

    def _build_ldm_sections(self) -> "list[ldm.Section]":
        from aspose.words_foss import light_document_model as ldm

        sections: list[ldm.Section] = []
        if self._document_xml is None:
            return sections

        body = self._document_xml.find(f"{W_NS}body")
        if body is None:
            return sections

        # Pre-scan for the first sectPr so positioned-shape extraction can
        # resolve ``relativeFrom="margin"``/``"column"``/``"paragraph"``
        # offsets into absolute page coordinates while paragraphs are still
        # being built.
        self._current_page_setup = None
        first_sect_pr = next(iter(body.iter(f"{W_NS}sectPr")), None)
        if first_sect_pr is not None:
            self._current_page_setup = self._build_ldm_page_setup(first_sect_pr)

        self._first_body_page_active = True  # type: ignore[attr-defined]

        # Collect body children, splitting at sectPr boundaries
        # In DOCX, sections are delimited by w:sectPr inside w:pPr (for
        # non-final sections) and by w:sectPr directly in w:body (final
        # section).
        current_children: list[ldm.Paragraph | ldm.Table | ldm.UnknownNode] = []
        section_props: list[Optional[ET.Element]] = []

        for element in self._resolve_body_children(body):
            if element.tag == f"{W_NS}p":
                para = self._build_ldm_paragraph(element)
                current_children.append(para)
                # Any form-feed in a run (``<w:br w:type="page"/>``)
                # or an explicit ``pageBreakBefore`` on the next
                # paragraph marks the end of the cover page.
                ends_page = self._paragraph_ends_first_page(  # type: ignore[attr-defined]
                    para, element
                )
                if self._first_body_page_active and ends_page:  # type: ignore[attr-defined]
                    self._first_body_page_active = False  # type: ignore[attr-defined]
                # Check for section break in paragraph properties
                pPr = element.find(f"{W_NS}pPr")
                if pPr is not None:
                    sect_pr = pPr.find(f"{W_NS}sectPr")
                    if sect_pr is not None:
                        section_props.append(sect_pr)
                        sec = ldm.Section()
                        sec.page_setup = self._build_ldm_page_setup(sect_pr)
                        sec.body = ldm.Body(children=current_children)
                        sections.append(sec)
                        current_children = []
            elif element.tag == f"{W_NS}tbl":
                tbl = self._build_ldm_table(element)
                current_children.append(tbl)
            elif element.tag == f"{W_NS}sectPr":
                # Final section properties
                sec = ldm.Section()
                sec.page_setup = self._build_ldm_page_setup(element)
                sec.body = ldm.Body(children=current_children)
                sections.append(sec)
                current_children = []

        # If there are remaining children without a sectPr, add a default section
        if current_children:
            sec = ldm.Section()
            sec.body = ldm.Body(children=current_children)
            sections.append(sec)

        return sections

    @staticmethod
    def _paragraph_ends_first_page(para: "ldm.Paragraph", p_elem: ET.Element) -> bool:
        """Return True when *para* closes Word's first body page.

        Matches an explicit ``<w:br w:type="page"/>`` inside any run
        (the form-feed character the reader emits for those) and the
        paragraph-level ``<w:pageBreakBefore/>`` flag.  Used by
        :meth:`_build_ldm_sections` to clear
        ``_first_body_page_active`` so cover-page-only absolute anchor
        promotion doesn't leak into later-page shapes.
        """
        for run in para.runs:
            if "\f" in (run.text or ""):
                return True
        return para.paragraph_format.page_break_before

    def _build_ldm_page_setup(self, sect_pr: ET.Element) -> "ldm.PageSetup":
        from aspose.words_foss import light_document_model as ldm

        ps = ldm.PageSetup()

        pg_sz = sect_pr.find(f"{W_NS}pgSz")
        if pg_sz is not None:
            w = pg_sz.get(f"{W_NS}w", "")
            h = pg_sz.get(f"{W_NS}h", "")
            if w:
                ps.page_width = int(w) / _TWIPS_PER_PT
            if h:
                ps.page_height = int(h) / _TWIPS_PER_PT
            orient = pg_sz.get(f"{W_NS}orient", "")
            ps.orientation = _Orient.LANDSCAPE if orient == "landscape" else _Orient.PORTRAIT
            # Derive paper_size from dimensions
            ps.paper_size = self._detect_paper_size(ps.page_width, ps.page_height)

        pg_mar = sect_pr.find(f"{W_NS}pgMar")
        if pg_mar is not None:
            for attr, field_name in [
                ("top", "top_margin"),
                ("bottom", "bottom_margin"),
                ("left", "left_margin"),
                ("right", "right_margin"),
                ("header", "header_distance"),
                ("footer", "footer_distance"),
            ]:
                val = pg_mar.get(f"{W_NS}{attr}", "")
                if val:
                    setattr(ps, field_name, int(val) / _TWIPS_PER_PT)

        type_elem = sect_pr.find(f"{W_NS}type")
        if type_elem is not None:
            ps.section_start = _SECTION_START_MAP.get(type_elem.get(f"{W_NS}val", ""), 2)

        title_pg = sect_pr.find(f"{W_NS}titlePg")
        if title_pg is not None:
            ps.different_first_page_header_footer = True

        # Page numbering: w:pgNumType w:start="N" w:fmt="..."
        pgnum = sect_pr.find(f"{W_NS}pgNumType")
        if pgnum is not None:
            start = pgnum.get(f"{W_NS}start")
            if start is not None:
                ps.page_starting_number = int(start)
                ps.restart_page_numbering = True

        return ps

    @staticmethod
    def _detect_paper_size(width: float, height: float) -> int:
        """Detect standard paper size from dimensions in points."""
        w, h = min(width, height), max(width, height)
        # Letter: 612 x 792
        if (
            abs(w - _LETTER_WIDTH_PT) < _PAPER_SIZE_TOLERANCE_PT
            and abs(h - _LETTER_HEIGHT_PT) < _PAPER_SIZE_TOLERANCE_PT
        ):
            return _PAPER_LETTER
        # Legal: 612 x 1008
        if (
            abs(w - _LETTER_WIDTH_PT) < _PAPER_SIZE_TOLERANCE_PT
            and abs(h - _LEGAL_HEIGHT_PT) < _PAPER_SIZE_TOLERANCE_PT
        ):
            return _PAPER_LEGAL
        # A4: 595.28 x 841.89
        if (
            abs(w - _A4_WIDTH_PT) < _PAPER_SIZE_TOLERANCE_A_PT
            and abs(h - _A4_HEIGHT_PT) < _PAPER_SIZE_TOLERANCE_A_PT
        ):
            return _PAPER_A4
        # A3: 841.89 x 1190.55
        if (
            abs(w - _A3_WIDTH_PT) < _PAPER_SIZE_TOLERANCE_A_PT
            and abs(h - _A3_HEIGHT_PT) < _PAPER_SIZE_TOLERANCE_A_PT
        ):
            return _PAPER_A3
        return _PAPER_CUSTOM

    # -- Paragraph ------------------------------------------------------------

    def _build_ldm_paragraph(
        self,
        p_elem: ET.Element,
        image_rels: Optional[dict[str, str]] = None,
    ) -> "ldm.Paragraph":
        from aspose.words_foss import light_document_model as ldm

        # Default to document-level image relationships
        if image_rels is None:
            image_rels = self._doc_image_rels

        para = ldm.Paragraph()

        # Determine paragraph style ID for inheritance resolution
        pPr = p_elem.find(f"{W_NS}pPr")
        para_style_id = ""
        if pPr is not None:
            pStyle = pPr.find(f"{W_NS}pStyle")
            if pStyle is not None:
                para_style_id = pStyle.get(f"{W_NS}val", "")

        # Build resolved paragraph format (docDefaults → style chain → direct)
        para.paragraph_format = self._build_resolved_paragraph_format(pPr, para_style_id)
        pf = para.paragraph_format

        # Set style_name from paragraph style
        if para_style_id:
            pf.style_name = self._resolve_style_name(para_style_id)

        # If heading not detected from outlineLvl, infer from style name
        if not pf.is_heading and pf.style_name:
            heading_match = re.search(r"[Hh]eading\s*(\d+)", pf.style_name)
            if heading_match:
                pf.is_heading = True
                pf.outline_level = int(heading_match.group(1)) - 1

        # List format (numId=0 means explicitly no list)
        # Check direct numPr first, then resolve from style chain
        numPr = pPr.find(f"{W_NS}numPr") if pPr is not None else None
        if numPr is None and para_style_id:
            numPr = self._resolve_numPr_from_style(para_style_id)
        if numPr is not None:
            numId_elem = numPr.find(f"{W_NS}numId")
            num_id_val = int(numId_elem.get(f"{W_NS}val", "0")) if numId_elem is not None else 0
            if num_id_val > 0:
                lf = ldm.ListFormat()
                lf.is_list_item = True
                lf.list_id = num_id_val
                ilvl = numPr.find(f"{W_NS}ilvl")
                if ilvl is not None:
                    lf.list_level_number = int(ilvl.get(f"{W_NS}val", "0"))
                para.list_format = lf
                para.paragraph_format.is_list_item = True

        # Parse runs and hyperlinks
        text_parts: list[str] = []
        # Field-code tracking: ``{PAGE}`` instrText runs produce a single
        # ``PAGE_FIELD_SENTINEL`` run that the PDF writer substitutes with
        # the live page number.  Runs between ``separate`` and ``end`` carry
        # the cached field value and are suppressed for supported fields.
        field_depth = 0
        page_field_pending = False  # saw begin, awaiting separate
        suppress_cached = False  # inside supported field's cached-result range
        for child in self._resolve_paragraph_children(p_elem):
            if child.tag == f"{W_NS}bookmarkStart":
                # Preserve named bookmarks so the writer can register
                # them as PDF link targets and resolve ``#name`` anchor
                # references (TOC entries).  Unnamed / ``_GoBack``
                # bookmarks carry no semantic value and are skipped.
                name = child.get(f"{W_NS}name", "")
                if name and not name.startswith("_GoBack"):
                    para.inline_extras.append(ldm.BookmarkStart(name=name))
                continue
            if child.tag == f"{W_NS}bookmarkEnd":
                continue
            if child.tag == f"{W_NS}r":
                # Classify field-code runs (``w:fldChar`` / ``w:instrText``)
                # before emitting text so cached PAGE numbers are replaced
                # with a live sentinel instead of a hard-coded literal.
                fld_char = child.find(f"{W_NS}fldChar")
                instr = child.find(f"{W_NS}instrText")
                if fld_char is not None:
                    ft = fld_char.get(f"{W_NS}fldCharType", "")
                    if ft == "begin":
                        field_depth += 1
                        page_field_pending = False
                    elif ft == "separate" and page_field_pending:
                        # Emit the sentinel now; suppress the cached runs
                        # until the matching ``end`` marker.
                        sentinel = ldm.Run()
                        sentinel.text = PAGE_FIELD_SENTINEL
                        sentinel.font = self._build_resolved_font(
                            child.find(f"{W_NS}rPr"), para_style_id
                        )
                        para.runs.append(sentinel)
                        para.content_sequence.append(sentinel)
                        text_parts.append(sentinel.text)
                        suppress_cached = True
                        page_field_pending = False
                    elif ft == "end":
                        field_depth = max(0, field_depth - 1)
                        suppress_cached = False
                        page_field_pending = False
                    continue
                if instr is not None:
                    # Recognise PAGE field (ignoring switches like ``\* Arabic``).
                    if field_depth > 0 and re.match(r"\s*PAGE(\s|$)", instr.text or ""):
                        page_field_pending = True
                    continue
                if suppress_cached:
                    # Drop the cached literal (e.g. "5") — the sentinel
                    # already stands in for the live value.
                    continue
                # Extract shapes from any embedded <w:drawing> elements.
                # ``iter`` is used so drawings wrapped in
                # ``<mc:AlternateContent><mc:Choice>`` (the common pattern
                # for cover-page text boxes) are still discovered, while
                # skipping the legacy ``<mc:Fallback>`` branch whose
                # ``<w:pict>`` would otherwise duplicate them.
                for drawing in self._iter_effective_drawings(child):
                    # Cover-page anchored groups become a list of
                    # absolutely-positioned shapes (rectangles with fill,
                    # overlaid text boxes).  These are stored separately
                    # so the PDF writer renders them at their page
                    # coordinates instead of inlining them into the flow.
                    positioned = self._extract_positioned_shapes(drawing, image_rels)
                    if positioned:
                        for pshape in positioned:
                            para.inline_extras.append(pshape)
                        continue
                    shape = self._build_drawing_shape(drawing, image_rels)
                    if shape is not None:
                        para.inline_extras.append(shape)
                        para.content_sequence.append(shape)
                # Build run with resolved font (style chain + docDefaults)
                run = self._build_ldm_run_resolved(child, para_style_id)
                if run.text:
                    para.runs.append(run)
                    para.content_sequence.append(run)
                    text_parts.append(run.text)
            elif child.tag == f"{W_NS}hyperlink":
                r_id = child.get(f"{R_NS}id", "")
                url = self._rels.get(r_id, "")
                anchor = child.get(f"{W_NS}anchor", "")
                if anchor:
                    url = url + "#" + anchor if url else "#" + anchor
                # TOC entries wrap the whole line — title, tab, and
                # cached ``PAGEREF`` result — in a single hyperlink.  Split
                # at the first ``<w:tab/>`` so the title becomes the
                # clickable link and the tab + page number are emitted as
                # trailing plain runs.  The writer's tab-split path then
                # right-aligns the page number at the right margin
                # instead of welding "Executive Summary2" into one word.
                hyp_runs = list(child.findall(f"{W_NS}r"))
                split_idx = len(hyp_runs)
                for i, r_elem in enumerate(hyp_runs):
                    if r_elem.find(f"{W_NS}tab") is not None:
                        split_idx = i
                        break
                head_runs = hyp_runs[:split_idx]
                tail_runs = hyp_runs[split_idx:]

                # rPr from the first inner run so direct formatting is preserved
                first_link_rPr: Optional[ET.Element] = (
                    head_runs[0].find(f"{W_NS}rPr") if head_runs else None
                )
                link_text = "".join(_collect_run_text(r) for r in head_runs)
                if link_text and url:
                    run = ldm.Run()
                    run.text = f"[{link_text}]({url})"
                    run.font = self._build_resolved_font(first_link_rPr, para_style_id)
                    para.runs.append(run)
                    para.content_sequence.append(run)
                    text_parts.append(run.text)
                elif link_text:
                    run = ldm.Run()
                    run.text = link_text
                    run.font = self._build_resolved_font(first_link_rPr, para_style_id)
                    para.runs.append(run)
                    para.content_sequence.append(run)
                    text_parts.append(run.text)
                # Trailing runs (tab + cached PAGEREF result).  Field
                # markers (``fldChar``/``instrText``) have empty visible
                # text via :meth:`_collect_run_text`, so they are skipped
                # naturally by the ``if trun.text`` guard below.
                for r_elem in tail_runs:
                    trun = self._build_ldm_run_resolved(r_elem, para_style_id)
                    if trun.text:
                        para.runs.append(trun)
                        para.content_sequence.append(trun)
                        text_parts.append(trun.text)

        para.text = "".join(text_parts)
        return para

    # -- Paragraph Format -----------------------------------------------------

    def _build_ldm_paragraph_format(self, pPr: ET.Element) -> "ldm.ParagraphFormat":
        from aspose.words_foss import light_document_model as ldm

        pf = ldm.ParagraphFormat()

        # Style name (resolve ID to display name)
        pStyle = pPr.find(f"{W_NS}pStyle")
        if pStyle is not None:
            style_id = pStyle.get(f"{W_NS}val", "")
            pf.style_name = self._resolve_style_name(style_id)

        # Alignment
        jc = pPr.find(f"{W_NS}jc")
        if jc is not None:
            pf.alignment = _ALIGNMENT_MAP.get(jc.get(f"{W_NS}val", "left"), 0)

        # Indentation
        ind = pPr.find(f"{W_NS}ind")
        if ind is not None:
            left = ind.get(f"{W_NS}left", "")
            if left:
                pf.left_indent = int(left) / _TWIPS_PER_PT
            right = ind.get(f"{W_NS}right", "")
            if right:
                pf.right_indent = int(right) / _TWIPS_PER_PT
            first_line = ind.get(f"{W_NS}firstLine", "")
            hanging = ind.get(f"{W_NS}hanging", "")
            if first_line:
                pf.first_line_indent = int(first_line) / _TWIPS_PER_PT
            elif hanging:
                pf.first_line_indent = -int(hanging) / _TWIPS_PER_PT

        # Spacing
        spacing = pPr.find(f"{W_NS}spacing")
        if spacing is not None:
            before = spacing.get(f"{W_NS}before", "")
            if before:
                pf.space_before = int(before) / _TWIPS_PER_PT
            after = spacing.get(f"{W_NS}after", "")
            if after:
                pf.space_after = int(after) / _TWIPS_PER_PT
            line = spacing.get(f"{W_NS}line", "")
            if line:
                pf.line_spacing = int(line) / _TWIPS_PER_PT
            line_rule = spacing.get(f"{W_NS}lineRule", "")
            if line_rule:
                pf.line_spacing_rule = _LINE_RULE_MAP.get(line_rule, 0)
            if spacing.get(f"{W_NS}beforeAutospacing", "") in (
                "1",
                "true",
            ):
                pf.space_before_auto = True
            if spacing.get(f"{W_NS}afterAutospacing", "") in (
                "1",
                "true",
            ):
                pf.space_after_auto = True

        # Keep with next
        kwn = pPr.find(f"{W_NS}keepNext")
        if kwn is not None:
            val = kwn.get(f"{W_NS}val")
            pf.keep_with_next = val is None or val not in ("false", "0")

        # Page break before
        pbb = pPr.find(f"{W_NS}pageBreakBefore")
        if pbb is not None:
            val = pbb.get(f"{W_NS}val")
            pf.page_break_before = val is None or val not in ("false", "0")

        # Outline level
        outline = pPr.find(f"{W_NS}outlineLvl")
        if outline is not None:
            ol_val = outline.get(f"{W_NS}val", "9")
            pf.outline_level = int(ol_val)
            if pf.outline_level < _OUTLINE_LEVEL_BODY:
                pf.is_heading = True

        # Shading
        shd = pPr.find(f"{W_NS}shd")
        if shd is not None:
            pf.shading = self._build_ldm_shading(shd)

        # Borders (always 6 entries matching Aspose.Words output)
        pBdr = pPr.find(f"{W_NS}pBdr")
        if pBdr is not None:
            pf.borders = self._build_ldm_borders(pBdr)
        else:
            pf.borders = _empty_borders()

        return pf

    # -- Run / Font -----------------------------------------------------------

    def _build_ldm_run_resolved(self, r_elem: ET.Element, para_style_id: str = "") -> "ldm.Run":
        """Build a Run with fully resolved font (docDefaults → style → direct)."""
        from aspose.words_foss import light_document_model as ldm

        run = ldm.Run()
        run.text = _collect_run_text(r_elem)

        # Resolved font: docDefaults → paragraph style chain → character style → direct
        rPr = r_elem.find(f"{W_NS}rPr")
        run.font = self._build_resolved_font(rPr, para_style_id)

        return run

    def _build_ldm_font(self, rPr: ET.Element) -> "ldm.Font":
        from aspose.words_foss import light_document_model as ldm

        font = ldm.Font()

        # Font name (resolve theme references)
        rFonts = rPr.find(f"{W_NS}rFonts")
        if rFonts is not None:
            explicit_name = (
                rFonts.get(f"{W_NS}ascii", "")
                or rFonts.get(f"{W_NS}hAnsi", "")
                or rFonts.get(f"{W_NS}cs", "")
            )
            if explicit_name:
                font.name = explicit_name
            else:
                # Resolve theme font references
                theme = rFonts.get(f"{W_NS}asciiTheme", "") or rFonts.get(f"{W_NS}hAnsiTheme", "")
                if theme:
                    font.name = self._resolve_theme_font(theme)

        # Size (half-points → points)
        sz = rPr.find(f"{W_NS}sz")
        if sz is not None:
            val = sz.get(f"{W_NS}val", "")
            if val:
                font.size = int(val) / _HALF_PT_DIVISOR

        # Bold
        b = rPr.find(f"{W_NS}b")
        if b is not None:
            val = b.get(f"{W_NS}val")
            font.bold = val is None or val not in ("false", "0")

        # Italic
        i = rPr.find(f"{W_NS}i")
        if i is not None:
            val = i.get(f"{W_NS}val")
            font.italic = val is None or val not in ("false", "0")

        # Underline
        u = rPr.find(f"{W_NS}u")
        if u is not None:
            u_val = u.get(f"{W_NS}val", "none")
            font.underline = _UNDERLINE_MAP.get(u_val, 1 if u_val != "none" else 0)

        # Color — honour ``w:themeColor`` by resolving via the parsed
        # theme when the raw ``w:val`` is missing or set to ``auto``.
        # Modifiers are consumed directly on the ``w:color`` element:
        # ``w:themeShade`` darkens and ``w:themeTint`` lightens the base
        # theme color (values are 2-digit hex in 1-byte range).
        color = rPr.find(f"{W_NS}color")
        if color is not None:
            val = color.get(f"{W_NS}val", "")
            if val and val.lower() != "auto":
                font.color = _hex_to_aspose_color(val)
            else:
                theme_name = color.get(f"{W_NS}themeColor", "")
                if theme_name:
                    resolved = self._resolve_theme_color(theme_name)
                    if resolved:
                        resolved = _apply_theme_color_modifiers(
                            resolved,
                            tint=color.get(f"{W_NS}themeTint"),
                            shade=color.get(f"{W_NS}themeShade"),
                        )
                        font.color = _hex_to_aspose_color(resolved)

        # Strikethrough
        strike = rPr.find(f"{W_NS}strike")
        if strike is not None:
            val = strike.get(f"{W_NS}val")
            font.strike_through = val is None or val not in ("false", "0")

        # Superscript / subscript
        vert_align = rPr.find(f"{W_NS}vertAlign")
        if vert_align is not None:
            va_val = vert_align.get(f"{W_NS}val", "")
            font.superscript = va_val == "superscript"
            font.subscript = va_val == "subscript"

        # Highlight color
        highlight = rPr.find(f"{W_NS}highlight")
        if highlight is not None:
            hl_val = highlight.get(f"{W_NS}val", "")
            font.highlight_color = _HIGHLIGHT_COLOR_MAP.get(hl_val, COLOR_EMPTY)

        # Caps
        caps = rPr.find(f"{W_NS}caps")
        if caps is not None:
            val = caps.get(f"{W_NS}val")
            font.all_caps = val is None or val not in ("false", "0")

        small_caps = rPr.find(f"{W_NS}smallCaps")
        if small_caps is not None:
            val = small_caps.get(f"{W_NS}val")
            font.small_caps = val is None or val not in ("false", "0")

        # Hidden
        vanish = rPr.find(f"{W_NS}vanish")
        if vanish is not None:
            val = vanish.get(f"{W_NS}val")
            font.hidden = val is None or val not in ("false", "0")

        # Style name (character style - resolve ID to display name)
        rStyle = rPr.find(f"{W_NS}rStyle")
        if rStyle is not None:
            style_id = rStyle.get(f"{W_NS}val", "")
            font.style_name = self._resolve_style_name(style_id)

        # Shading
        shd = rPr.find(f"{W_NS}shd")
        if shd is not None:
            font.shading = self._build_ldm_shading(shd)

        return font

    # -- Table ----------------------------------------------------------------

    def _build_ldm_table(self, tbl_elem: ET.Element) -> "ldm.Table":
        from aspose.words_foss import light_document_model as ldm

        tbl = ldm.Table()

        # Table properties
        tblPr = tbl_elem.find(f"{W_NS}tblPr")
        if tblPr is not None:
            jc = tblPr.find(f"{W_NS}jc")
            if jc is not None:
                tbl.alignment = _ALIGNMENT_MAP.get(jc.get(f"{W_NS}val", "left"), 0)

            tblW = tblPr.find(f"{W_NS}tblW")
            if tblW is not None:
                w_type = tblW.get(f"{W_NS}type", "")
                w_val = tblW.get(f"{W_NS}w", "0")
                if w_type == "pct":
                    tbl.preferred_width = f"{int(w_val) / _PCT_DIVISOR}%"
                elif w_type == "dxa":
                    tbl.preferred_width = f"{int(w_val) / _TWIPS_PER_PT}pt"
                elif w_type == "auto":
                    tbl.preferred_width = "Auto"

            tblInd = tblPr.find(f"{W_NS}tblInd")
            if tblInd is not None:
                val = tblInd.get(f"{W_NS}w", "0")
                tbl.left_indent = int(val) / _TWIPS_PER_PT

            # Table cell margins (default paddings)
            tblCellMar = tblPr.find(f"{W_NS}tblCellMar")
            if tblCellMar is not None:
                for side, attr in [
                    ("left", "left_padding"),
                    ("right", "right_padding"),
                    ("top", "top_padding"),
                    ("bottom", "bottom_padding"),
                ]:
                    side_elem = tblCellMar.find(f"{W_NS}{side}")
                    if side_elem is not None:
                        val = side_elem.get(f"{W_NS}w", "0")
                        setattr(tbl, attr, int(val) / _TWIPS_PER_PT)

        # Collect default cell paddings from table for inheritance
        default_paddings = (
            tbl.left_padding,
            tbl.right_padding,
            tbl.top_padding,
            tbl.bottom_padding,
        )

        # Rows
        for tr_elem in tbl_elem.findall(f"{W_NS}tr"):
            row = self._build_ldm_row(tr_elem, default_paddings)
            tbl.rows.append(row)

        return tbl

    def _build_ldm_row(
        self,
        tr_elem: ET.Element,
        default_paddings: Optional[tuple[float, float, float, float]] = None,
    ) -> "ldm.Row":
        from aspose.words_foss import light_document_model as ldm

        row = ldm.Row()

        trPr = tr_elem.find(f"{W_NS}trPr")
        if trPr is not None:
            rf = ldm.RowFormat()
            trHeight = trPr.find(f"{W_NS}trHeight")
            if trHeight is not None:
                val = trHeight.get(f"{W_NS}val", "0")
                rf.height = int(val) / _TWIPS_PER_PT
                rule = trHeight.get(f"{W_NS}hRule", "")
                if rule == "exact":
                    rf.height_rule = 1
                elif rule == "auto":
                    rf.height_rule = 2
                else:
                    rf.height_rule = 0

            tblHeader = trPr.find(f"{W_NS}tblHeader")
            if tblHeader is not None:
                rf.heading_format = True

            cant_split = trPr.find(f"{W_NS}cantSplit")
            if cant_split is not None:
                rf.allow_break_across_pages = False

            row.row_format = rf

        for tc_elem in tr_elem.findall(f"{W_NS}tc"):
            cell = self._build_ldm_cell(tc_elem, default_paddings)
            row.cells.append(cell)

        return row

    def _build_ldm_cell(
        self,
        tc_elem: ET.Element,
        default_paddings: Optional[tuple[float, float, float, float]] = None,
    ) -> "ldm.Cell":
        from aspose.words_foss import light_document_model as ldm
        from aspose.words_foss.model.enums import CellVerticalAlignment as _CVA

        cell = ldm.Cell()

        tcPr = tc_elem.find(f"{W_NS}tcPr")
        if tcPr is not None:
            cf = ldm.CellFormat()

            # Inherit table-level default cell paddings
            if default_paddings is not None:
                cf.left_padding, cf.right_padding, cf.top_padding, cf.bottom_padding = (
                    default_paddings
                )

            tcW = tcPr.find(f"{W_NS}tcW")
            if tcW is not None:
                w_type = tcW.get(f"{W_NS}type", "")
                w_val = tcW.get(f"{W_NS}w", "0")
                if w_type == "dxa":
                    cf.width = int(w_val) / _TWIPS_PER_PT
                    # Raw twip string matches Aspose.Words reference output
                    cf.preferred_width = w_val
                elif w_type == "pct":
                    cf.preferred_width = f"{int(w_val) / _PCT_DIVISOR}%"
                elif w_type == "auto":
                    cf.preferred_width = "Auto"

            vAlign = tcPr.find(f"{W_NS}vAlign")
            if vAlign is not None:
                va = vAlign.get(f"{W_NS}val", "top")
                cf.vertical_alignment = {
                    "top": _CVA.TOP,
                    "center": _CVA.CENTER,
                    "bottom": _CVA.BOTTOM,
                }.get(va, _CVA.TOP)

            vMerge = tcPr.find(f"{W_NS}vMerge")
            if vMerge is not None:
                val = vMerge.get(f"{W_NS}val", "continue")
                cf.vertical_merge = 1 if val == "restart" else 2

            gridSpan = tcPr.find(f"{W_NS}gridSpan")
            if gridSpan is not None:
                span = int(gridSpan.get(f"{W_NS}val", "1"))
                if span > 1:
                    cf.horizontal_merge = 1

            # Cell margins
            tcMar = tcPr.find(f"{W_NS}tcMar")
            if tcMar is not None:
                for side, attr in [
                    ("left", "left_padding"),
                    ("right", "right_padding"),
                    ("top", "top_padding"),
                    ("bottom", "bottom_padding"),
                ]:
                    side_elem = tcMar.find(f"{W_NS}{side}")
                    if side_elem is not None:
                        val = side_elem.get(f"{W_NS}w", "0")
                        setattr(cf, attr, int(val) / _TWIPS_PER_PT)

            # Shading
            shd = tcPr.find(f"{W_NS}shd")
            if shd is not None:
                cf.shading = self._build_ldm_shading(shd)

            # Borders (always 6 entries)
            tcBorders = tcPr.find(f"{W_NS}tcBorders")
            if tcBorders is not None:
                cf.borders = self._build_ldm_borders(tcBorders)
            else:
                cf.borders = _empty_borders()

            cell.cell_format = cf

        # Paragraphs in cell
        for p_elem in tc_elem.findall(f"{W_NS}p"):
            cell.paragraphs.append(self._build_ldm_paragraph(p_elem))

        # Nested tables
        for tbl_elem in tc_elem.findall(f"{W_NS}tbl"):
            cell.tables.append(self._build_ldm_table(tbl_elem))

        return cell

    # -- Shared helpers -------------------------------------------------------

    def _build_ldm_shading(self, shd: ET.Element) -> "ldm.Shading":
        from aspose.words_foss import light_document_model as ldm

        s = ldm.Shading()
        fill = shd.get(f"{W_NS}fill", "")
        if fill:
            s.background_color = fill
        return s

    def _build_ldm_borders(self, bdr_elem: ET.Element) -> "list[ldm.Border]":
        from aspose.words_foss import light_document_model as ldm

        borders: list[ldm.Border] = []
        # Standard border sides in order: top, left, bottom, right, between, bar
        for side in ("top", "left", "bottom", "right", "between", "bar"):
            side_elem = bdr_elem.find(f"{W_NS}{side}")
            if side_elem is not None:
                b = ldm.Border()
                style_val = side_elem.get(f"{W_NS}val", "none")
                b.line_style = _BORDER_STYLE_MAP.get(style_val, 0)
                sz = side_elem.get(f"{W_NS}sz", "")
                if sz:
                    b.line_width = int(sz) / _BORDER_SIZE_DIVISOR
                color = side_elem.get(f"{W_NS}color", "")
                b.color = _hex_to_aspose_color(color)
                borders.append(b)
            else:
                borders.append(ldm.Border(color=COLOR_EMPTY))
        return borders

    # The following methods are provided by DocumentReader or ShapeParserMixin
    # at runtime via multiple inheritance.  They are declared here only for
    # static type checkers (mypy / pyright) and are guarded behind
    # TYPE_CHECKING so they never shadow the real implementations.
    if TYPE_CHECKING:

        def _resolve_style_name(self, style_id: str) -> str: ...

        def _resolve_theme_font(self, theme_name: str) -> str: ...

        def _resolve_theme_color(self, scheme_name: str) -> str: ...

        def _resolve_body_children(  # noqa: E501
            self, parent: ET.Element
        ) -> Iterator[ET.Element]: ...

        def _resolve_paragraph_children(
            self, p_elem: ET.Element
        ) -> Iterator[ET.Element]: ...

        def _iter_effective_drawings(
            self, run_elem: ET.Element
        ) -> Iterator[ET.Element]: ...

        def _extract_positioned_shapes(
            self,
            drawing_elem: ET.Element,
            image_rels: dict[str, str],
        ) -> "list[ldm.ShapeNode]": ...

        def _build_drawing_shape(
            self,
            drawing_elem: ET.Element,
            image_rels: dict[str, str],
        ) -> "Optional[ldm.ShapeNode]": ...
