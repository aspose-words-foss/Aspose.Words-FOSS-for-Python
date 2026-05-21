"""OPC zip-package writer for DOCX files.

Builds the four ancillary parts every Word document needs:

* ``[Content_Types].xml`` — content-type registry.
* ``_rels/.rels`` — package-level relationships.
* ``word/_rels/document.xml.rels`` — document-level relationships
  (hyperlinks, styles, numbering).
* The ZIP archive itself.

Keeping packaging in its own module means the renderer never touches
``zipfile`` directly, and adding new parts (headers, footers, images)
later is a localised change.
"""


import zipfile
from pathlib import Path
from typing import IO, Optional, Union

from aspose.words_foss.docx_writer.constants import (
    CT_CORE_PROPS,
    CT_DOCUMENT,
    CT_EXT_PROPS,
    CT_FONT_TABLE,
    CT_FOOTER,
    CT_HEADER,
    CT_NUMBERING,
    CT_RELS,
    CT_SETTINGS,
    CT_STYLES,
    CT_URI,
    PKG_RELS_URI,
    REL_CORE_PROPS,
    REL_EXT_PROPS,
    REL_FONT_TABLE,
    REL_FOOTER,
    REL_HEADER,
    REL_HYPERLINK,
    REL_NUMBERING,
    REL_OFFICE_DOCUMENT,
    REL_SETTINGS,
    REL_STYLES,
)
from aspose.words_foss.docx_writer.doc_props_part import (
    render_app_xml,
    render_core_xml,
    render_font_table_xml,
)
from aspose.words_foss.docx_writer.drawing import REL_IMAGE, ImageEntry
from aspose.words_foss.docx_writer.headers_footers import (
    FOOTER_PART_PATH,
    FOOTER_REL_ID,
    FOOTER_TARGET,
    HEADER_PART_PATH,
    HEADER_REL_ID,
    HEADER_TARGET,
)
from aspose.words_foss.docx_writer.xml_utils import XML_DECL, el, indent_xml
from aspose.words_foss.docx_writer.drawing import _ext_to_content_type


def _content_types(
    has_numbering: bool,
    has_header: bool,
    has_footer: bool,
    has_settings: bool,
    image_extensions: set[str],
) -> str:
    """Build ``[Content_Types].xml``.

    Images live under ``word/media/imageN.<ext>``; we register one
    ``<Default>`` per distinct extension (``png``, ``jpeg`` …) so Word
    resolves their content type via the file suffix instead of needing
    a per-image ``<Override>``.
    """

    children = [
        el("Default", {"Extension": "rels", "ContentType": CT_RELS}),
        el(
            "Default",
            {"Extension": "xml", "ContentType": "application/xml"},
        ),
    ]
    for ext in sorted(image_extensions):
        children.append(
            el(
                "Default",
                {"Extension": ext, "ContentType": _ext_to_content_type(ext)},
            )
        )
    children.extend(
        [
            el(
                "Override",
                {"PartName": "/word/document.xml", "ContentType": CT_DOCUMENT},
            ),
            el(
                "Override",
                {"PartName": "/word/styles.xml", "ContentType": CT_STYLES},
            ),
        ]
    )
    if has_numbering:
        children.append(
            el(
                "Override",
                {"PartName": "/word/numbering.xml", "ContentType": CT_NUMBERING},
            )
        )
    if has_header:
        children.append(
            el(
                "Override",
                {"PartName": f"/{HEADER_PART_PATH}", "ContentType": CT_HEADER},
            )
        )
    if has_footer:
        children.append(
            el(
                "Override",
                {"PartName": f"/{FOOTER_PART_PATH}", "ContentType": CT_FOOTER},
            )
        )
    if has_settings:
        children.append(
            el(
                "Override",
                {"PartName": "/word/settings.xml", "ContentType": CT_SETTINGS},
            )
        )
    # Standard ancillary parts — always emitted because MS Word's
    # stricter loader (Word for Mac, Word 2016+, Office Online)
    # rejects packages that omit them even though ECMA-376 marks
    # them "Recommended".
    children.append(
        el("Override", {"PartName": "/word/fontTable.xml", "ContentType": CT_FONT_TABLE})
    )
    children.append(
        el("Override", {"PartName": "/docProps/core.xml", "ContentType": CT_CORE_PROPS})
    )
    children.append(
        el("Override", {"PartName": "/docProps/app.xml", "ContentType": CT_EXT_PROPS})
    )
    return XML_DECL + el("Types", {"xmlns": CT_URI}, children)


def _pkg_rels() -> str:
    """Top-level ``_rels/.rels`` linking the package root to the office
    document plus the two ``docProps/*`` parts MS Word expects to find
    referenced from the package root."""
    children = [
        el(
            "Relationship",
            {
                "Id": "rId1",
                "Type": REL_OFFICE_DOCUMENT,
                "Target": "word/document.xml",
            },
        ),
        el(
            "Relationship",
            {
                "Id": "rIdCoreProps",
                "Type": REL_CORE_PROPS,
                "Target": "docProps/core.xml",
            },
        ),
        el(
            "Relationship",
            {
                "Id": "rIdExtProps",
                "Type": REL_EXT_PROPS,
                "Target": "docProps/app.xml",
            },
        ),
    ]
    return XML_DECL + el("Relationships", {"xmlns": PKG_RELS_URI}, children)


def _doc_rels(
    hyperlinks: dict[str, str],
    has_numbering: bool,
    has_header: bool,
    has_footer: bool,
    has_settings: bool,
    images: list[ImageEntry],
) -> str:
    """Build ``word/_rels/document.xml.rels`` from accumulated relationships."""
    children = [
        el(
            "Relationship",
            {
                "Id": "rIdStyles",
                "Type": REL_STYLES,
                "Target": "styles.xml",
            },
        ),
    ]
    if has_numbering:
        children.append(
            el(
                "Relationship",
                {
                    "Id": "rIdNumbering",
                    "Type": REL_NUMBERING,
                    "Target": "numbering.xml",
                },
            )
        )
    if has_header:
        children.append(
            el(
                "Relationship",
                {
                    "Id": HEADER_REL_ID,
                    "Type": REL_HEADER,
                    "Target": HEADER_TARGET,
                },
            )
        )
    if has_footer:
        children.append(
            el(
                "Relationship",
                {
                    "Id": FOOTER_REL_ID,
                    "Type": REL_FOOTER,
                    "Target": FOOTER_TARGET,
                },
            )
        )
    if has_settings:
        children.append(
            el(
                "Relationship",
                {
                    "Id": "rIdSettings",
                    "Type": REL_SETTINGS,
                    "Target": "settings.xml",
                },
            )
        )
    # MS Word looks up the font table via this relationship and
    # falls back to default substitutions when it's missing —
    # several Office for Mac builds *require* the entry to be
    # present even if the fontTable.xml itself is empty.
    children.append(
        el(
            "Relationship",
            {
                "Id": "rIdFontTable",
                "Type": REL_FONT_TABLE,
                "Target": "fontTable.xml",
            },
        )
    )
    for image in images:
        # Targets in ``word/_rels/document.xml.rels`` are relative to
        # ``word/`` — strip the leading ``word/`` from media_path.
        target = image.media_path.removeprefix("word/")
        children.append(
            el(
                "Relationship",
                {
                    "Id": image.rel_id,
                    "Type": REL_IMAGE,
                    "Target": target,
                },
            )
        )
    for url, rid in hyperlinks.items():
        children.append(
            el(
                "Relationship",
                {
                    "Id": rid,
                    "Type": REL_HYPERLINK,
                    "Target": url,
                    "TargetMode": "External",
                },
            )
        )
    return XML_DECL + el("Relationships", {"xmlns": PKG_RELS_URI}, children)


def _hf_part_rels(hyperlinks: dict[str, str], images: list[ImageEntry]) -> str:
    """Build a sidecar ``.rels`` file for a header / footer part.

    Headers and footers can carry external hyperlinks and embedded
    images just like the document body; each part needs its own
    ``<Target>`` table because ``r:id`` references are resolved
    against the part's own rels file, not the document's.  Image
    targets are written relative to the part's location
    (``word/header1.xml`` resolves ``Target="media/image.png"`` as
    ``word/media/image.png``).
    """
    children: list[str] = []
    for url, rid in hyperlinks.items():
        children.append(
            el(
                "Relationship",
                {
                    "Id": rid,
                    "Type": REL_HYPERLINK,
                    "Target": url,
                    "TargetMode": "External",
                },
            )
        )
    for image in images:
        target = image.media_path.removeprefix("word/")
        children.append(
            el(
                "Relationship",
                {
                    "Id": image.rel_id,
                    "Type": REL_IMAGE,
                    "Target": target,
                },
            )
        )
    return XML_DECL + el("Relationships", {"xmlns": PKG_RELS_URI}, children)


def _build_parts(
    *,
    document_xml: str,
    styles_xml: str,
    numbering_xml: Optional[str],
    settings_xml: Optional[str],
    header_xml: Optional[str],
    header_hyperlinks: dict[str, str],
    header_images: list[ImageEntry],
    footer_xml: Optional[str],
    footer_hyperlinks: dict[str, str],
    footer_images: list[ImageEntry],
    hyperlinks: dict[str, str],
    images: list[ImageEntry],
) -> tuple[list[tuple[str, str]], list[tuple[str, bytes]]]:
    """Assemble ordered (text_parts, binary_parts) lists for the zip."""
    has_numbering = bool(numbering_xml)
    has_header = bool(header_xml)
    has_footer = bool(footer_xml)
    has_settings = bool(settings_xml)
    all_images = images + header_images + footer_images
    image_extensions = {image.media_path.rsplit(".", 1)[1].lower() for image in all_images}
    text_parts: list[tuple[str, str]] = [
        (
            "[Content_Types].xml",
            _content_types(has_numbering, has_header, has_footer, has_settings, image_extensions),
        ),
        ("_rels/.rels", _pkg_rels()),
        ("word/document.xml", document_xml),
        (
            "word/_rels/document.xml.rels",
            _doc_rels(hyperlinks, has_numbering, has_header, has_footer, has_settings, images),
        ),
        ("word/styles.xml", styles_xml),
    ]
    if has_numbering and numbering_xml:
        text_parts.append(("word/numbering.xml", numbering_xml))
    if has_settings and settings_xml:
        text_parts.append(("word/settings.xml", settings_xml))
    # Ancillary parts (always emitted) — see ``doc_props_part`` for why.
    text_parts.append(("word/fontTable.xml", render_font_table_xml()))
    text_parts.append(("docProps/core.xml", render_core_xml()))
    text_parts.append(("docProps/app.xml", render_app_xml()))
    if has_header and header_xml:
        text_parts.append((HEADER_PART_PATH, header_xml))
        if header_hyperlinks or header_images:
            text_parts.append(
                (
                    f"word/_rels/{HEADER_TARGET}.rels",
                    _hf_part_rels(header_hyperlinks, header_images),
                )
            )
    if has_footer and footer_xml:
        text_parts.append((FOOTER_PART_PATH, footer_xml))
        if footer_hyperlinks or footer_images:
            text_parts.append(
                (
                    f"word/_rels/{FOOTER_TARGET}.rels",
                    _hf_part_rels(footer_hyperlinks, footer_images),
                )
            )
    binary_parts: list[tuple[str, bytes]] = [
        (image.media_path, image.image_bytes) for image in all_images
    ]
    return text_parts, binary_parts


def write_docx_package(
    target: Union[str, Path, IO[bytes]],
    *,
    document_xml: str,
    styles_xml: str,
    numbering_xml: Optional[str],
    settings_xml: Optional[str] = None,
    hyperlinks: dict[str, str],
    images: Optional[list[ImageEntry]] = None,
    header_xml: Optional[str] = None,
    header_hyperlinks: Optional[dict[str, str]] = None,
    header_images: Optional[list[ImageEntry]] = None,
    footer_xml: Optional[str] = None,
    footer_hyperlinks: Optional[dict[str, str]] = None,
    footer_images: Optional[list[ImageEntry]] = None,
    compression: int = zipfile.ZIP_DEFLATED,
    compresslevel: Optional[int] = None,
    allow_zip64: bool = False,
    pretty_format: bool = False,
) -> None:
    """Write the assembled parts to ``target`` as a DOCX zip.

    ``target`` may be a path (``str`` / :class:`Path`) or any binary
    file-like object — :class:`zipfile.ZipFile` accepts both, so
    callers wanting in-memory output can pass an :class:`io.BytesIO`.
    Path targets get their parent directory auto-created; file-like
    targets are written to as-is.

    ``compresslevel`` is forwarded to :class:`zipfile.ZipFile` and
    controls the deflate compression effort (0–9).  ``None`` lets the
    standard library choose its default.

    ``allow_zip64`` enables ZIP64 extensions for archives exceeding the
    classic 4 GB / 65 535-entry ZIP limits.
    """
    text_parts, binary_parts = _build_parts(
        document_xml=document_xml,
        styles_xml=styles_xml,
        numbering_xml=numbering_xml,
        settings_xml=settings_xml,
        header_xml=header_xml,
        header_hyperlinks=header_hyperlinks or {},
        header_images=list(header_images or []),
        footer_xml=footer_xml,
        footer_hyperlinks=footer_hyperlinks or {},
        footer_images=list(footer_images or []),
        hyperlinks=hyperlinks,
        images=list(images or []),
    )

    if pretty_format:
        text_parts = [(name, indent_xml(content)) for name, content in text_parts]

    if isinstance(target, (str, Path)):
        path = Path(target)
        path.parent.mkdir(parents=True, exist_ok=True)
        target = path  # zipfile resolves Path internally

    zf_kwargs: dict[str, object] = {}
    if compresslevel is not None:
        zf_kwargs["compresslevel"] = compresslevel

    with zipfile.ZipFile(
        target, "w", compression, allowZip64=allow_zip64, **zf_kwargs
    ) as zf:
        for name, content in text_parts:
            zf.writestr(name, content)
        for name, payload in binary_parts:
            zf.writestr(name, payload)
