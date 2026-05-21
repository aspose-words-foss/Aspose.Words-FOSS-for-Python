"""Generates ``word/settings.xml``.

Always emitted: the Aspose blank baseline carries a ``<w:compat>``
block that Word2010+ expects on every package.  LDM-tracked values
(currently :attr:`Document.default_tab_stop`) are added when they
differ from Word's compiled-in default.
"""


from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_writer.blank_template import settings_compat
from aspose.words_foss.docx_writer.constants import W_URI, pt_to_twips
from aspose.words_foss.docx_writer.xml_utils import XML_DECL, el


_DEFAULT_TAB_STOP_PT = 36.0


def needs_settings(doc: ldm.Document) -> bool:
    """Always True — the compat block belongs in every package."""
    return True


def render_settings_xml(doc: ldm.Document) -> str:
    """Build ``word/settings.xml`` payload."""
    children: list[str] = []
    if abs(doc.default_tab_stop - _DEFAULT_TAB_STOP_PT) > 0.01:
        children.append(
            el("w:defaultTabStop", {"w:val": pt_to_twips(doc.default_tab_stop)})
        )
    children.append(settings_compat())
    root = el("w:settings", {"xmlns:w": W_URI}, children)
    return XML_DECL + root
