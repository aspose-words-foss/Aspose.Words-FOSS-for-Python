"""Header / footer XML rendering.

The reader collects every paragraph from every ``word/header*.xml`` /
``word/footer*.xml`` part into a single flat list on
:class:`Document` (``header_paragraphs`` / ``footer_paragraphs``); it
does not preserve the per-section / per-type (default / firstPage /
even) split.  We mirror that on write by emitting one ``header1.xml``
and one ``footer1.xml`` containing the whole list, and reference the
single pair from every section's ``<w:sectPr>`` so a Word user still
sees the header on every page.

Headers/footers in the test fixtures carry no inline shapes, bookmarks
or hyperlinks, so this first pass renders only ``<w:p>`` content.  When
that limitation matters, ``ImageRenderState`` / ``BookmarkState`` /
``rels`` would need to be threaded in alongside their own ``.rels``
sidecar — see ``package.py`` for the structure that's already in place
for the document part.
"""


from typing import Mapping, Optional

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_writer.bookmarks import BookmarkState
from aspose.words_foss.docx_writer.constants import MC_URI, O_URI, R_URI, V_URI, W_URI
from aspose.words_foss.docx_writer.drawing import A_URI, PIC_URI, WP_URI, WPS_URI, ImageRenderState
from aspose.words_foss.docx_writer.paragraphs import render_paragraph
from aspose.words_foss.docx_writer.tables import render_table
from aspose.words_foss.docx_writer.xml_utils import XML_DECL, el


def _render_part(
    children: list[ldm.Paragraph | ldm.Table],
    rels: dict,
    *,
    root_tag: str,
    style_pf_map: Mapping[str, ldm.ParagraphFormat],
    style_id_map: Mapping[str, str],
    style_font_map: Mapping[str, ldm.Font],
    image_state: "ImageRenderState | None" = None,
) -> str:
    """Render a single ``<w:hdr>`` or ``<w:ftr>`` XML payload.

    ``children`` is an ordered list of paragraphs and tables preserving
    the original document order.  Empty lists still emit the root element
    with a single empty ``<w:p/>`` — Word rejects header / footer parts
    that contain no paragraphs.

    ``image_state`` accumulates images referenced from inside the part;
    the caller passes a fresh instance per part so the resulting
    ``ImageEntry`` rows can be wired into the part's sidecar ``.rels``
    file and added to ``word/media/`` with non-colliding filenames.
    """
    rendered: list[str] = []
    if image_state is None:
        image_state = ImageRenderState()
    bookmark_state = BookmarkState()
    for child in children:
        if isinstance(child, ldm.Paragraph):
            rendered.append(
                render_paragraph(
                    child,
                    rels,
                    num_id_map={},
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
                    num_id_map={},
                    image_state=image_state,
                    bookmark_state=bookmark_state,
                    style_pf_map=style_pf_map,
                    style_id_map=style_id_map,
                    style_font_map=style_font_map,
                )
            )
    for _, bm_id in bookmark_state.drain_open():
        rendered.append(el("w:bookmarkEnd", {"w:id": bm_id}))
    if not rendered:
        rendered.append(el("w:p"))
    root = el(
        root_tag,
        {
            "xmlns:w": W_URI,
            "xmlns:r": R_URI,
            "xmlns:wp": WP_URI,
            "xmlns:a": A_URI,
            "xmlns:pic": PIC_URI,
            "xmlns:wps": WPS_URI,
            "xmlns:mc": MC_URI,
            "xmlns:v": V_URI,
            "xmlns:o": O_URI,
            # ``mc:Ignorable`` only lists prefixes whose namespace is
            # declared above; Word rejects the package otherwise.
            "mc:Ignorable": "wps",
        },
        rendered,
    )
    return XML_DECL + root


def render_header_xml(
    children: list[ldm.Paragraph | ldm.Table],
    rels: dict,
    style_pf_map: Mapping[str, ldm.ParagraphFormat],
    style_id_map: Mapping[str, str],
    style_font_map: Mapping[str, ldm.Font],
    image_state: "ImageRenderState | None" = None,
) -> str:
    return _render_part(
        children,
        rels,
        root_tag="w:hdr",
        style_pf_map=style_pf_map,
        style_id_map=style_id_map,
        style_font_map=style_font_map,
        image_state=image_state,
    )


def render_footer_xml(
    children: list[ldm.Paragraph | ldm.Table],
    rels: dict,
    style_pf_map: Mapping[str, ldm.ParagraphFormat],
    style_id_map: Mapping[str, str],
    style_font_map: Mapping[str, ldm.Font],
    image_state: "ImageRenderState | None" = None,
) -> str:
    return _render_part(
        children,
        rels,
        root_tag="w:ftr",
        style_pf_map=style_pf_map,
        style_id_map=style_id_map,
        style_font_map=style_font_map,
        image_state=image_state,
    )


def header_reference(rid: str) -> str:
    """``<w:headerReference w:type="default" r:id="..."/>`` for ``<w:sectPr>``."""
    return el("w:headerReference", {"w:type": "default", "r:id": rid})


def footer_reference(rid: str) -> str:
    """``<w:footerReference w:type="default" r:id="..."/>`` for ``<w:sectPr>``."""
    return el("w:footerReference", {"w:type": "default", "r:id": rid})


HEADER_REL_ID = "rIdHeader1"
FOOTER_REL_ID = "rIdFooter1"

# Filenames inside the ``.docx`` zip — kept in sync with the ``Target``
# attribute in ``word/_rels/document.xml.rels``.
HEADER_PART_PATH = "word/header1.xml"
FOOTER_PART_PATH = "word/footer1.xml"
HEADER_TARGET = "header1.xml"
FOOTER_TARGET = "footer1.xml"


def needs_header(doc: ldm.Document) -> bool:
    if doc.header_paragraphs:
        return True
    for sec in doc.sections:
        for hf in sec.headers_footers:
            if hf.header_footer_type == 0 and hf.children:
                return True
    return False


def needs_footer(doc: ldm.Document) -> bool:
    if doc.footer_paragraphs:
        return True
    for sec in doc.sections:
        for hf in sec.headers_footers:
            if hf.header_footer_type == 1 and hf.children:
                return True
    return False


# ----------------------------------------------------------------------
# Filtering helper: the writer threads ``unused_kw`` through ``render_paragraph``
# but they aren't used by the headers/footers writer; this stub keeps the
# import surface symmetrical with ``document_part``.
# ----------------------------------------------------------------------


def _unused() -> Optional[str]:  # pragma: no cover - documentation glue
    return None
