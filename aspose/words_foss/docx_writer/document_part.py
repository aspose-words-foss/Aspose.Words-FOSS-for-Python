"""Renders the main ``word/document.xml`` part."""


from typing import Mapping, Optional

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_writer.bookmarks import BookmarkState
from aspose.words_foss.docx_writer.constants import (
    MC_URI,
    O_URI,
    R_URI,
    V_URI,
    W_URI,
    pt_to_twips,
)
from aspose.words_foss.docx_writer.drawing import A_URI, PIC_URI, WP_URI, WPS_URI, ImageRenderState
from aspose.words_foss.docx_writer.headers_footers import (
    FOOTER_REL_ID,
    HEADER_REL_ID,
    footer_reference,
    header_reference,
)
from aspose.words_foss.docx_writer.numbering_part import build_num_id_map
from aspose.words_foss.docx_writer.paragraphs import render_paragraph
from aspose.words_foss.docx_writer.styles_part import (
    build_style_font_map,
    build_style_id_map,
)
from aspose.words_foss.docx_writer.tables import render_table
from aspose.words_foss.docx_writer.xml_utils import XML_DECL, el

_SECTION_START_TOKEN = {
    0: "continuous",
    1: "newColumn",
    2: "nextPage",
    3: "evenPage",
    4: "oddPage",
}


def _sectPr(
    page_setup: ldm.PageSetup,
    *,
    header_ref: Optional[str] = None,
    footer_ref: Optional[str] = None,
) -> str:
    """Build ``<w:sectPr>`` honouring the ECMA-376 CT_SectPr child order:

    ``headerReference → footerReference → footnotePr → endnotePr →
    type → pgSz → pgMar → paperSrc → pgBorders → lnNumType → pgNumType →
    cols → ... → titlePg → ... → docGrid``.
    """
    children: list[str] = []
    # Header / footer references must come first.
    if header_ref:
        children.append(header_ref)
    if footer_ref:
        children.append(footer_ref)

    # ``<w:type>`` (section break kind) sits between footnote/endnote
    # properties and ``pgSz`` per CT_SectPr — emitting it after cols
    # produced a schema-invalid order that Aspose / Word would flag.
    sect_token = _SECTION_START_TOKEN.get(page_setup.section_start)
    if sect_token and page_setup.section_start != 2:
        # ``nextPage`` is the OOXML default — omit to keep the round-trip
        # silent for the most common case.
        children.append(el("w:type", {"w:val": sect_token}))

    pg_size_attrs: dict[str, object] = {}
    if page_setup.page_width > 0:
        pg_size_attrs["w:w"] = pt_to_twips(page_setup.page_width)
    if page_setup.page_height > 0:
        pg_size_attrs["w:h"] = pt_to_twips(page_setup.page_height)
    if page_setup.orientation == 1:
        pg_size_attrs["w:orient"] = "landscape"
    if pg_size_attrs:
        children.append(el("w:pgSz", pg_size_attrs))

    # ``CT_PageMar`` requires all seven attributes (top/right/bottom/left/
    # header/footer/gutter); Word refuses to open the file when any are
    # missing.  Emit zeros for fields the LDM doesn't carry — the reader
    # only consumes whichever attributes are present so the round-trip
    # is preserved.
    has_margin_data = any(
        v > 0 for v in (
            page_setup.top_margin,
            page_setup.right_margin,
            page_setup.bottom_margin,
            page_setup.left_margin,
            page_setup.header_distance,
            page_setup.footer_distance,
            page_setup.gutter,
        )
    )
    if has_margin_data:
        margin_attrs: dict[str, object] = {
            "w:top": pt_to_twips(page_setup.top_margin),
            "w:right": pt_to_twips(page_setup.right_margin),
            "w:bottom": pt_to_twips(page_setup.bottom_margin),
            "w:left": pt_to_twips(page_setup.left_margin),
            "w:header": pt_to_twips(page_setup.header_distance),
            "w:footer": pt_to_twips(page_setup.footer_distance),
            "w:gutter": pt_to_twips(page_setup.gutter),
        }
        children.append(el("w:pgMar", margin_attrs))

    # ``pgNumType`` precedes ``cols`` per CT_SectPr.  ``page_starting_number``
    # defaults to 0; only emit when the reader saw a real
    # ``<w:pgNumType w:start>`` (signalled by
    # ``restart_page_numbering=True``) so the round-trip doesn't invent
    # a fake "restart at 0" on every section.
    if page_setup.restart_page_numbering:
        children.append(el("w:pgNumType", {"w:start": page_setup.page_starting_number}))

    if page_setup.text_columns is not None:
        tc = page_setup.text_columns
        cols_attrs: dict[str, object] = {}
        if tc.count > 1:
            cols_attrs["w:num"] = tc.count
        if tc.spacing > 0:
            cols_attrs["w:space"] = pt_to_twips(tc.spacing)
        if tc.line_between:
            cols_attrs["w:sep"] = 1
        if not tc.evenly_spaced:
            cols_attrs["w:equalWidth"] = 0
        if tc.columns:
            col_children: list[str] = []
            for col in tc.columns:
                col_attrs: dict[str, object] = {}
                if col.width > 0:
                    col_attrs["w:w"] = pt_to_twips(col.width)
                if col.space_after > 0:
                    col_attrs["w:space"] = pt_to_twips(col.space_after)
                col_children.append(el("w:col", col_attrs))
            children.append(el("w:cols", cols_attrs, col_children))
        else:
            children.append(el("w:cols", cols_attrs))

    if page_setup.different_first_page_header_footer:
        children.append(el("w:titlePg"))

    if not children:
        return el("w:sectPr")
    return el("w:sectPr", None, children)


def _last_paragraph_index(children: list) -> Optional[int]:
    """Index of the last ``Paragraph`` in ``children``, or ``None``."""
    for i in range(len(children) - 1, -1, -1):
        if isinstance(children[i], ldm.Paragraph):
            return i
    return None


def _build_style_pf_map(doc: ldm.Document) -> dict[str, ldm.ParagraphFormat]:
    """Map canonical style id → resolved :class:`ParagraphFormat`.

    Keyed by ``style_name.replace(" ", "").lower()`` so paragraph references
    using either the display name (``Heading 1``) or style id (``Heading1``)
    in either case both find their style.  Built once per document and
    threaded through paragraph rendering for diff-based pPr emission.
    """
    pf_map: dict[str, ldm.ParagraphFormat] = {}
    for style in doc.styles:
        if style.type != 1 or style.paragraph_format is None:
            continue
        canonical = style.name.replace(" ", "").lower()
        if canonical:
            pf_map[canonical] = style.paragraph_format
    return pf_map


def _render_paragraph(
    para: ldm.Paragraph,
    rels: dict,
    *,
    embedded_sectPr: Optional[str],
    num_id_map: Mapping[int, int],
    image_state: ImageRenderState,
    bookmark_state: BookmarkState,
    style_pf_map: Mapping[str, ldm.ParagraphFormat],
    style_id_map: Mapping[str, str],
    style_font_map: Mapping[str, ldm.Font],
) -> str:
    return render_paragraph(
        para,
        rels,
        embedded_sectPr=embedded_sectPr,
        num_id_map=num_id_map,
        image_state=image_state,
        bookmark_state=bookmark_state,
        style_pf_map=style_pf_map,
        style_id_map=style_id_map,
        style_font_map=style_font_map,
    )


def _render_section(
    section: ldm.Section,
    rels: dict,
    *,
    is_last: bool,
    num_id_map: Mapping[int, int],
    image_state: ImageRenderState,
    bookmark_state: BookmarkState,
    style_pf_map: Mapping[str, ldm.ParagraphFormat],
    style_id_map: Mapping[str, str],
    style_font_map: Mapping[str, ldm.Font],
    has_header: bool,
    has_footer: bool,
) -> list[str]:
    """Render one section's body content.

    OOXML allows exactly one trailing ``<w:sectPr>`` directly under
    ``<w:body>`` (the final section's setup).  Every earlier section
    must embed its ``<w:sectPr>`` inside the last paragraph's
    ``<w:pPr>`` — so we hunt down that paragraph and pass the rendered
    ``sectPr`` through ``render_paragraph``.  When a section happens
    to end on a table (no trailing paragraph), we synthesise an empty
    paragraph to carry the section break, mirroring what Word itself
    does when the user inserts a section break after a table.
    """
    rendered: list[str] = []
    header_ref = header_reference(HEADER_REL_ID) if has_header else None
    footer_ref = footer_reference(FOOTER_REL_ID) if has_footer else None
    sectPr_xml = _sectPr(section.page_setup, header_ref=header_ref, footer_ref=footer_ref)
    children = list(section.body.children)
    image_state.set_section_margins(section.page_setup)

    if is_last:
        for child in children:
            if isinstance(child, ldm.Paragraph):
                rendered.append(
                    _render_paragraph(
                        child,
                        rels,
                        embedded_sectPr=None,
                        num_id_map=num_id_map,
                        image_state=image_state,
                        bookmark_state=bookmark_state,
                        style_pf_map=style_pf_map,
                        style_id_map=style_id_map,
                        style_font_map=style_font_map,
                    )
                )
            elif isinstance(child, ldm.Table):
                rendered.append(
                    render_table(
                        child,
                        rels,
                        num_id_map=num_id_map,
                        image_state=image_state,
                        bookmark_state=bookmark_state,
                        style_pf_map=style_pf_map,
                        style_id_map=style_id_map,
                        style_font_map=style_font_map,
                    )
                )
        rendered.append(sectPr_xml)
        return rendered

    last_para_idx = _last_paragraph_index(children)
    if last_para_idx is None:
        # Section ends on a table — render everything, then a
        # placeholder paragraph carrying the embedded sectPr.
        for child in children:
            if isinstance(child, ldm.Paragraph):
                rendered.append(
                    _render_paragraph(
                        child,
                        rels,
                        embedded_sectPr=None,
                        num_id_map=num_id_map,
                        image_state=image_state,
                        bookmark_state=bookmark_state,
                        style_pf_map=style_pf_map,
                        style_id_map=style_id_map,
                        style_font_map=style_font_map,
                    )
                )
            elif isinstance(child, ldm.Table):
                rendered.append(
                    render_table(
                        child,
                        rels,
                        num_id_map=num_id_map,
                        image_state=image_state,
                        bookmark_state=bookmark_state,
                        style_pf_map=style_pf_map,
                        style_id_map=style_id_map,
                        style_font_map=style_font_map,
                    )
                )
        rendered.append(
            _render_paragraph(
                ldm.Paragraph(),
                rels,
                embedded_sectPr=sectPr_xml,
                num_id_map=num_id_map,
                image_state=image_state,
                bookmark_state=bookmark_state,
                style_pf_map=style_pf_map,
                style_id_map=style_id_map,
                style_font_map=style_font_map,
            )
        )
        return rendered

    for i, child in enumerate(children):
        if isinstance(child, ldm.Paragraph):
            sectPr_arg = sectPr_xml if i == last_para_idx else None
            rendered.append(
                _render_paragraph(
                    child,
                    rels,
                    embedded_sectPr=sectPr_arg,
                    num_id_map=num_id_map,
                    image_state=image_state,
                    bookmark_state=bookmark_state,
                    style_pf_map=style_pf_map,
                    style_id_map=style_id_map,
                    style_font_map=style_font_map,
                )
            )
        elif isinstance(child, ldm.Table):
            rendered.append(
                render_table(
                    child,
                    rels,
                    num_id_map=num_id_map,
                    image_state=image_state,
                    bookmark_state=bookmark_state,
                    style_pf_map=style_pf_map,
                    style_id_map=style_id_map,
                    style_font_map=style_font_map,
                )
            )
    return rendered


def _render_body(
    doc: ldm.Document,
    rels: dict,
    num_id_map: Mapping[int, int],
    image_state: ImageRenderState,
    bookmark_state: BookmarkState,
    style_pf_map: Mapping[str, ldm.ParagraphFormat],
    style_id_map: Mapping[str, str],
    style_font_map: Mapping[str, ldm.Font],
    *,
    has_header: bool,
    has_footer: bool,
) -> str:
    """Walk all sections and render their body children."""
    rendered: list[str] = []
    if not doc.sections:
        # An empty document still needs at least one paragraph and
        # a sectPr, otherwise Word refuses to open it.
        rendered.append(el("w:p"))
        header_ref = header_reference(HEADER_REL_ID) if has_header else None
        footer_ref = footer_reference(FOOTER_REL_ID) if has_footer else None
        rendered.append(_sectPr(ldm.PageSetup(), header_ref=header_ref, footer_ref=footer_ref))
        return el("w:body", None, rendered)

    last_idx = len(doc.sections) - 1
    for i, section in enumerate(doc.sections):
        rendered.extend(
            _render_section(
                section,
                rels,
                is_last=(i == last_idx),
                num_id_map=num_id_map,
                image_state=image_state,
                bookmark_state=bookmark_state,
                style_pf_map=style_pf_map,
                style_id_map=style_id_map,
                style_font_map=style_font_map,
                has_header=has_header,
                has_footer=has_footer,
            )
        )
    # Word treats unclosed bookmarks as a corruption error; flush any
    # starts whose matching end never showed up at the document tail.
    for _, bm_id in bookmark_state.drain_open():
        rendered.append(el("w:bookmarkEnd", {"w:id": bm_id}))
    return el("w:body", None, rendered)


def render_document_xml(
    doc: ldm.Document,
    rels: dict,
    *,
    image_state: Optional[ImageRenderState] = None,
    bookmark_state: Optional[BookmarkState] = None,
    has_header: bool = False,
    has_footer: bool = False,
) -> str:
    """Build the entire ``word/document.xml`` payload as a string.

    ``rels`` is mutated by run/paragraph rendering — caller passes the
    same dict to the relationships writer afterwards.  ``image_state``
    accumulates inline shape relationships and image bytes; the writer's
    main path always passes a real :class:`ImageRenderState`, but tests
    that don't carry image data may omit it (a fresh one is created).

    The collision-free numId mapping is built once here and threaded
    through paragraphs and tables so every list reference resolves to
    the same id that ``numbering.xml`` declares.
    """
    if image_state is None:
        image_state = ImageRenderState()
    if bookmark_state is None:
        bookmark_state = BookmarkState()
    num_id_map = build_num_id_map(doc.lists)
    style_pf_map = _build_style_pf_map(doc)
    style_id_map = build_style_id_map(doc)
    style_font_map = build_style_font_map(doc)
    body = _render_body(
        doc,
        rels,
        num_id_map,
        image_state,
        bookmark_state,
        style_pf_map,
        style_id_map,
        style_font_map,
        has_header=has_header,
        has_footer=has_footer,
    )
    # <w:background> sits between document opening and body per CT_Document.
    background = ""
    if doc.page_color:
        from aspose.words_foss.docx_writer.runs import color_to_hex
        hex_val = color_to_hex(doc.page_color)
        if hex_val:
            background = el("w:background", {"w:color": hex_val})
    # Declare each prefix once on the root (Word's strict parser
    # rejects repeated xmlns on every drawing).  mc:Ignorable lists
    # post-2007 prefixes a strict ECMA-376 consumer may skip —
    # required for wps (text-box / cover-page shapes) or Word
    # rejects the package as corrupt.  Only declared prefixes are
    # allowed in mc:Ignorable.
    root = el(
        "w:document",
        {
            "xmlns:w": W_URI,
            "xmlns:r": R_URI,
            "xmlns:wp": WP_URI,
            "xmlns:a": A_URI,
            "xmlns:pic": PIC_URI,
            "xmlns:wps": WPS_URI,
            "xmlns:mc": MC_URI,
            # ``v`` / ``o`` are used inside the VML ``<mc:Fallback>``
            # branch of ``<mc:AlternateContent>`` for wps:wsp shapes.
            "xmlns:v": V_URI,
            "xmlns:o": O_URI,
            "mc:Ignorable": "wps",
        },
        background + body if background else body,
    )
    return XML_DECL + root
