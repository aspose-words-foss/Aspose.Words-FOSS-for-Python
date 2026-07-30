"""Facade mixin that wires the LDM builders into :class:`DocumentReader`.

The mixin owns one set of builder instances per document, instantiates
them lazily in :meth:`to_light_document`, and orchestrates the
top-level build in document order: styles → lists → sections →
header/footer parts.  No XML parsing happens here — every method
delegates to a specialised builder.
"""

from typing import Optional
from xml.etree import ElementTree as ET

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_reader.constants import (
    COLOR_EMPTY,
    W_NS,
    _DEFAULT_TAB_STOP_PT,
    _TWIPS_PER_PT,
)
from aspose.words_foss.docx_reader.utils import (
    _canonicalize_style_name,
    _hex_to_ldm_color,
)

from .cascading import (
    FontBuilder,
    FontResolver,
    ParagraphFormatBuilder,
    ParagraphFormatResolver,
    StyleChainResolver,
)
from .definitions import ListBuilder, StyleBuilder
from .paragraphs import ParagraphBuilder, RunBuilder
from .sections import PageSetupBuilder, SectionBuilder
from .tables import TableBuilder


class LdmBuilderMixin:
    """Mixin that adds :meth:`to_light_document` to :class:`DocumentReader`."""

    def to_light_document(self) -> ldm.Document:
        """Build a :class:`ldm.Document` from the loaded DOCX."""
        doc = ldm.Document()
        doc.page_color = self._get_page_color()
        doc.default_tab_stop = self._get_default_tab_stop()
        self._bookmark_id_to_name = {}

        self._build_name_to_style_id_map()
        self._ensure_builders_installed()

        # build_all() synthesizes a default Normal when styles.xml is
        # absent, matching the writer's emit-Normal-by-default behaviour.
        doc.styles = self._styles_builder.build_all()
        if self._numbering_xml is not None:
            doc.lists = self._lists_builder.build_all()
        doc.sections = self._sections_builder.build_all()

        if self._doc_default_rPr is not None:
            doc.doc_defaults_font = self._fonts_builder.build(self._doc_default_rPr)

        self._populate_headers_footers(doc)
        return doc

    # -- Reader-side state helpers ---------------------------------------------

    def _get_default_tab_stop(self) -> float:
        if self._settings_xml is None:
            return _DEFAULT_TAB_STOP_PT
        elem = self._settings_xml.find(f"{W_NS}defaultTabStop")
        if elem is None:
            return _DEFAULT_TAB_STOP_PT
        val = elem.get(f"{W_NS}val", "")
        if not val:
            return _DEFAULT_TAB_STOP_PT
        try:
            return int(val) / _TWIPS_PER_PT
        except ValueError:
            return _DEFAULT_TAB_STOP_PT

    def _get_page_color(self) -> str:
        if self._document_xml is None:
            return COLOR_EMPTY
        bg = self._document_xml.find(f"{W_NS}background")
        if bg is None:
            return COLOR_EMPTY
        color = bg.get(f"{W_NS}color", "")
        return _hex_to_ldm_color(color) if color else COLOR_EMPTY

    def _build_name_to_style_id_map(self) -> None:
        self._name_to_style_id = {}
        if self._styles_xml is None:
            return
        for style_elem in self._styles_xml.findall(f"{W_NS}style"):
            sid = style_elem.get(f"{W_NS}styleId", "")
            name_elem = style_elem.find(f"{W_NS}name")
            if name_elem is None or not sid:
                continue
            raw_name = name_elem.get(f"{W_NS}val", "")
            self._name_to_style_id[raw_name] = sid
            is_custom = style_elem.get(f"{W_NS}customStyle", "") == "1"
            if is_custom:
                continue
            canonical = _canonicalize_style_name(raw_name)
            if canonical != raw_name:
                self._name_to_style_id[canonical] = sid

    def _install_builders(self) -> None:
        """Instantiate per-document builders and wire their dependencies."""
        if not hasattr(self, "_bookmark_id_to_name"):
            self._bookmark_id_to_name = {}
        if not hasattr(self, "_current_table_style_id"):
            self._current_table_style_id = ""

        # The mixin is only ever used through DocumentReader, so `self`
        # carries the parser-side state the builders read from.
        reader = self
        chains = StyleChainResolver(reader)
        fonts = FontBuilder(reader)
        pfs = ParagraphFormatBuilder(reader, fonts)
        font_resolver = FontResolver(reader, fonts, chains)
        pf_resolver = ParagraphFormatResolver(reader, pfs, chains)
        run_builder = RunBuilder(font_resolver)
        paragraph_builder = ParagraphBuilder(
            reader, font_resolver, pf_resolver, chains, run_builder
        )
        table_builder = TableBuilder(reader, paragraph_builder.build)
        page_setup_builder = PageSetupBuilder()
        sections_builder = SectionBuilder(
            reader, paragraph_builder, table_builder.build, page_setup_builder
        )

        self._chains = chains
        self._fonts_builder = fonts
        self._pfs_builder = pfs
        self._font_resolver = font_resolver
        self._pf_resolver = pf_resolver
        self._paragraph_builder = paragraph_builder
        self._table_builder = table_builder
        self._page_setup_builder = page_setup_builder
        self._sections_builder = sections_builder
        self._styles_builder = StyleBuilder(
            reader, fonts, font_resolver, pfs, pf_resolver
        )
        self._lists_builder = ListBuilder(reader, fonts)

    def _build_ldm_paragraph(
        self,
        p_elem: ET.Element,
        image_rels: Optional[dict[str, str]] = None,
    ) -> ldm.Paragraph:
        """Compatibility shim for :class:`ShapeParserMixin`."""
        self._ensure_builders_installed()
        return self._paragraph_builder.build(p_elem, image_rels)

    def _ensure_builders_installed(self) -> None:
        if getattr(self, "_paragraph_builder", None) is None:
            self._install_builders()

    def _populate_headers_footers(self, doc: ldm.Document) -> None:
        """Build header/footer children and attach them to every section."""
        hdr_paras, hdr_children = self._build_part_children(
            self._header_data, anchor_mode="header"
        )
        ftr_paras, ftr_children = self._build_part_children(
            self._footer_data, anchor_mode="footer"
        )
        for sec in doc.sections:
            if hdr_children:
                sec.headers_footers.append(
                    ldm.HeaderFooter(header_footer_type=0, children=hdr_children)
                )
            if ftr_children:
                sec.headers_footers.append(
                    ldm.HeaderFooter(header_footer_type=1, children=ftr_children)
                )

    def _build_part_children(
        self,
        parts: list[tuple[ET.Element, dict[str, str]]],
        *,
        anchor_mode: str,
    ) -> tuple[list[ldm.Paragraph], list[ldm.BodyChild]]:
        paragraphs: list[ldm.Paragraph] = []
        children: list[ldm.BodyChild] = []
        self._anchor_y_base_mode = anchor_mode  # type: ignore[attr-defined]
        try:
            for part_xml, part_rels in parts:
                for elem in part_xml:
                    if elem.tag == f"{W_NS}p":
                        p = self._paragraph_builder.build(elem, part_rels)
                        paragraphs.append(p)
                        children.append(p)
                    elif elem.tag == f"{W_NS}tbl":
                        children.append(self._table_builder.build(elem))
        finally:
            self._anchor_y_base_mode = "body"  # type: ignore[attr-defined]
        return paragraphs, children

    # -- Type-only declarations of host-provided state -------------------------
    # The mixin reads these from the host :class:`DocumentReader`; the
    # annotations here are for IDEs / type checkers and are never executed
    # (they shadow nothing because the host always sets the real values).
    # Keep this block in sync with :class:`ReaderContext` in ``_context``.
    _document_xml: Optional[ET.Element]
    _numbering_xml: Optional[ET.Element]
    _styles_xml: Optional[ET.Element]
    _settings_xml: Optional[ET.Element]
    _rels: dict[str, str]
    _doc_image_rels: dict[str, str]
    _style_elem_cache: dict[str, ET.Element]
    _style_id_to_name: dict[str, str]
    _name_to_style_id: dict[str, str]
    _theme_fonts: dict[str, str]
    _theme_colors: dict[str, str]
    _doc_default_rPr: Optional[ET.Element]
    _doc_default_pPr: Optional[ET.Element]
    _header_data: list[tuple[ET.Element, dict[str, str]]]
    _footer_data: list[tuple[ET.Element, dict[str, str]]]
    _current_page_setup: Optional[ldm.PageSetup]
    _current_table_style_id: str
    _bookmark_id_to_name: dict[str, str]
