"""Generates the optional but Word-recommended package parts.

MS Word's stricter loader (especially Word for Mac, Word 2016+
on Windows, and the various Word Online instances) rejects DOCX
files that omit ``docProps/core.xml``, ``docProps/app.xml``, or
``word/fontTable.xml``, even though ECMA-376 marks all three as
"Recommended" rather than "Required".  ``LibreOffice`` and
``Aspose.Words`` happily open packages that lack them.

We emit minimal stubs:

* ``docProps/core.xml`` carries a single ``<dc:title>`` placeholder
  and an empty creator field.
* ``docProps/app.xml`` declares the producing application.
* ``word/fontTable.xml`` is an empty ``<w:fonts>`` element — Word
  uses it only as a font-substitution table and tolerates an
  empty list.

Each part is paired with the ``<Override>`` entry in
``[Content_Types].xml`` and a relationship row in either
``_rels/.rels`` (for the docProps parts) or
``word/_rels/document.xml.rels`` (for the fontTable).
"""


from aspose.words_foss.docx_writer.xml_utils import XML_DECL, el


CORE_PROPS_URI = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC_URI = "http://purl.org/dc/elements/1.1/"
DCTERMS_URI = "http://purl.org/dc/terms/"
DCMITYPE_URI = "http://purl.org/dc/dcmitype/"
XSI_URI = "http://www.w3.org/2001/XMLSchema-instance"

EXT_PROPS_URI = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
DOCSEC_VT_URI = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"


def render_core_xml() -> str:
    """``docProps/core.xml`` — minimal Dublin Core metadata."""
    body = (
        el("dc:title", None, "")
        + el("dc:subject", None, "")
        + el("dc:creator", None, "")
        + el("cp:keywords", None, "")
        + el("dc:description", None, "")
        + el("cp:lastModifiedBy", None, "")
        + el("cp:revision", None, "1")
    )
    root = el(
        "cp:coreProperties",
        {
            "xmlns:cp": CORE_PROPS_URI,
            "xmlns:dc": DC_URI,
            "xmlns:dcterms": DCTERMS_URI,
            "xmlns:dcmitype": DCMITYPE_URI,
            "xmlns:xsi": XSI_URI,
        },
        body,
    )
    return XML_DECL + root


def render_app_xml() -> str:
    """``docProps/app.xml`` — minimal extended properties.

    ``<Application>`` is the only field Word actually inspects; the
    rest are commonly present and harmless to emit empty.
    """
    body = (
        el("Application", None, "aspose-words-foss")
        + el("DocSecurity", None, "0")
        + el("ScaleCrop", None, "false")
        + el("LinksUpToDate", None, "false")
        + el("SharedDoc", None, "false")
        + el("HyperlinksChanged", None, "false")
    )
    root = el(
        "Properties",
        {
            "xmlns": EXT_PROPS_URI,
            "xmlns:vt": DOCSEC_VT_URI,
        },
        body,
    )
    return XML_DECL + root


def render_font_table_xml() -> str:
    """``word/fontTable.xml`` — Word reads this for font substitution.

    Carries the same Times New Roman / Symbol / Arial entries
    (with full panose1 / charset / sig metadata) that Aspose's blank
    template ships, so Word doesn't fall back to its bare-bones
    built-in font registry on first open.
    """
    from aspose.words_foss.docx_writer.blank_template import font_definitions
    root = el(
        "w:fonts",
        {"xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"},
        font_definitions(),
    )
    return XML_DECL + root
