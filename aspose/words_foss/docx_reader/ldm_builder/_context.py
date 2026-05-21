"""Reader-context protocol shared by the LDM builders.

The builders read parser state (XML elements, style caches, theme
maps) and resolver methods that live on :class:`DocumentReader`.
:class:`ReaderContext` describes the subset they touch — declaring
it here as a structural :class:`typing.Protocol` lets the builder
modules type their ``ctx`` parameter without importing
``DocumentReader`` (which would form a circular import via
``document_reader.py``'s mixin inheritance).

``DocumentReader`` satisfies this protocol implicitly; no runtime
relationship is established.
"""

from typing import Iterator, Optional, Protocol
from xml.etree import ElementTree as ET

from aspose.words_foss import light_document_model as ldm


class ReaderContext(Protocol):
    # Parsed XML roots loaded from the DOCX package
    _document_xml: Optional[ET.Element]
    _numbering_xml: Optional[ET.Element]
    _styles_xml: Optional[ET.Element]
    _settings_xml: Optional[ET.Element]

    # Relationships and image targets
    _rels: dict[str, str]
    _doc_image_rels: dict[str, str]

    # Style lookup tables built from styles.xml
    _style_elem_cache: dict[str, ET.Element]
    _style_id_to_name: dict[str, str]
    _name_to_style_id: dict[str, str]

    # Theme tables built from theme1.xml
    _theme_fonts: dict[str, str]
    _theme_colors: dict[str, str]

    # Document-level defaults from styles.xml/docDefaults
    _doc_default_rPr: Optional[ET.Element]
    _doc_default_pPr: Optional[ET.Element]

    # Header/footer XML parts plus per-part image relationships
    _header_data: list[tuple[ET.Element, dict[str, str]]]
    _footer_data: list[tuple[ET.Element, dict[str, str]]]

    # Runtime state shared with ShapeParserMixin
    _current_page_setup: Optional[ldm.PageSetup]
    _current_table_style_id: str
    _bookmark_id_to_name: dict[str, str]

    # Resolvers implemented on DocumentReader / ShapeParserMixin
    def _resolve_style_name(self, style_id: str) -> str: ...
    def _resolve_theme_font(self, theme: str) -> str: ...
    def _resolve_theme_color(self, theme: str) -> str: ...
    def _resolve_body_children(self, parent: ET.Element) -> Iterator[ET.Element]: ...
    def _resolve_paragraph_children(self, p_elem: ET.Element) -> Iterator[ET.Element]: ...
    def _iter_effective_drawings(self, run_elem: ET.Element) -> Iterator[ET.Element]: ...
    def _extract_positioned_shapes(
        self, drawing_elem: ET.Element, image_rels: dict[str, str]
    ) -> list[ldm.ShapeNode]: ...
    def _build_drawing_shape(
        self, drawing_elem: ET.Element, image_rels: dict[str, str]
    ) -> Optional[ldm.ShapeNode]: ...
