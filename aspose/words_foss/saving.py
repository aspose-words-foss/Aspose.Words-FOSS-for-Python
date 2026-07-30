"""
save options for markdown and PDF export.
Provides MarkdownSaveOptions, PdfSaveOptions, and related enums
matching the public save API.

Enum members are ``IntEnum``s valued as in ``aspose.words.saving``. The
lowercase string spellings earlier releases used are still accepted when
assigned to a save-option field.
"""

from enum import IntEnum
from typing import Any, Optional, Type, TypeVar

_E = TypeVar("_E", bound=IntEnum)


def coerce_enum(enum_cls: Type[_E], value: Any) -> Any:
    """Normalise *value* to a member of *enum_cls*, or return it unchanged."""
    if isinstance(value, enum_cls):
        return value
    if isinstance(value, str):
        member = enum_cls.__members__.get(value.upper())
        return member if member is not None else value
    if isinstance(value, int):
        try:
            return enum_cls(value)
        except ValueError:
            return value
    return value


class _EnumField:
    """Save-option attribute coercing legacy string spellings on assign."""

    def __init__(self, enum_cls: Type[_E], default: _E) -> None:
        self._enum_cls = enum_cls
        self._default = default
        self._slot = ""

    def __set_name__(self, owner: type, name: str) -> None:
        self._slot = f"_{name}"

    def __get__(self, obj: Any, objtype: Optional[type] = None) -> Any:
        if obj is None:
            return self
        return getattr(obj, self._slot, self._default)

    def __set__(self, obj: Any, value: Any) -> None:
        setattr(obj, self._slot, coerce_enum(self._enum_cls, value))


class TableContentAlignment(IntEnum):
    """Table content alignment options."""

    AUTO = 0
    LEFT = 1
    CENTER = 2
    RIGHT = 3


class MarkdownListExportMode(IntEnum):
    """List export mode options."""

    MARKDOWN_SYNTAX = 0
    PLAIN_TEXT = 1


class MarkdownLinkExportMode(IntEnum):
    """Link export mode options."""

    AUTO = 0
    INLINE = 1
    REFERENCE = 2


class MarkdownExportAsHtml(IntEnum):
    """Controls which elements are exported as raw HTML."""

    NONE = 0
    TABLES = 1
    NON_COMPATIBLE_TABLES = 2


class MarkdownEmptyParagraphExportMode(IntEnum):
    """Controls how empty paragraphs are exported. Note NONE is 2, not 0."""

    EMPTY_LINE = 0
    MARKDOWN_HARD_LINE_BREAK = 1
    NONE = 2


class PdfCompliance(IntEnum):
    """PDF standards compliance level.

    Only the PDF version implied by each level is applied; none of the
    archival validation the standards require is performed.
    """

    PDF17 = 0
    PDF20 = 1
    PDF_A1A = 2
    PDF_A1B = 3
    PDF_A2A = 4
    PDF_A2U = 5
    PDF_A3A = 6
    PDF_A3U = 7
    PDF_A4 = 8
    PDF_A4F = 9
    PDF_A4_UA_2 = 10
    PDF_UA1 = 11
    PDF_UA2 = 12


class PdfTextCompression(IntEnum):
    """Text compression in PDF."""

    NONE = 0
    FLATE = 1


class PdfImageCompression(IntEnum):
    """Image compression in PDF."""

    AUTO = 0
    JPEG = 1


class PdfPageMode(IntEnum):
    """PDF page display mode."""

    USE_NONE = 0
    USE_OUTLINES = 1
    USE_THUMBS = 2
    FULL_SCREEN = 3
    USE_OC = 4
    USE_ATTACHMENTS = 5


class PdfZoomBehavior(IntEnum):
    """Mirrors public API."""

    NONE = 0
    ZOOM_FACTOR = 1
    FIT_PAGE = 2
    FIT_WIDTH = 3
    FIT_HEIGHT = 4
    FIT_BOX = 5


class PdfFontEmbeddingMode(IntEnum):
    """Font embedding mode in PDF. Note EMBED_NONE is 2, not 1."""

    EMBED_ALL = 0
    EMBED_NONSTANDARD = 1
    EMBED_NONE = 2


class ColorMode(IntEnum):
    """Color rendering mode."""

    NORMAL = 0
    GRAYSCALE = 1


class OutlineOptions:
    """Controls how outlines (bookmarks panel) are generated in the PDF.
    Note: the fpdf2 backend always fills gaps between non-contiguous
    outline levels (e.g. H1 followed by H3 inserts an empty H2 entry)
    because the library requires a contiguous hierarchy, so gap-filling
    stays active regardless of ``create_missing_outline_levels``.
    """

    def __init__(self):
        # Values >= 6 behave alike: heading depth is clamped to 6.
        self.headings_outline_levels: int = 0
        # Currently unused. Will be added in a future release.
        self.expanded_outline_levels: int = 0
        self.default_bookmarks_outline_level: int = 0
        # Bookmarks whose name starts with "_" are skipped even if listed here.
        self.bookmarks_outline_levels: dict[str, int] = {}
        # Only affects tables rendered via the image-cell fallback path.
        self.create_outlines_for_headings_in_tables: bool = False
        # Currently unused: gap-filling is always active in the fpdf2 backend.
        self.create_missing_outline_levels: bool = False


class PdfSaveOptions:
    """
    Options for saving documents as PDF.

    Usage:
        from aspose.words_foss.saving import PdfSaveOptions, PdfCompliance

        opts = PdfSaveOptions()
        opts.compliance = PdfCompliance.PDF17
        opts.jpeg_quality = 75
        opts.export_document_structure = True
        doc.save("output.pdf", opts)
    """

    compliance = _EnumField(PdfCompliance, PdfCompliance.PDF17)
    image_compression = _EnumField(PdfImageCompression, PdfImageCompression.AUTO)
    text_compression = _EnumField(PdfTextCompression, PdfTextCompression.FLATE)
    font_embedding_mode = _EnumField(PdfFontEmbeddingMode, PdfFontEmbeddingMode.EMBED_ALL)
    page_mode = _EnumField(PdfPageMode, PdfPageMode.USE_OUTLINES)
    color_mode = _EnumField(ColorMode, ColorMode.NORMAL)
    zoom_behavior = _EnumField(PdfZoomBehavior, PdfZoomBehavior.NONE)

    def __init__(self):
        # PDF standard compliance
        # Only sets the PDF version header; PDF/A and PDF/UA conformance is not implemented.
        self.compliance = PdfCompliance.PDF17

        # Document structure
        self.export_document_structure: bool = False

        # Image options
        self.image_compression = PdfImageCompression.AUTO
        self.jpeg_quality: int = 100

        # Text compression
        # Currently unused. Will be added in a future release.
        self.text_compression = PdfTextCompression.FLATE

        # Font embedding
        # Currently unused. Will be added in a future release.
        self.embed_full_fonts: bool = False
        # Currently unused. Will be added in a future release.
        self.use_core_fonts: bool = False
        # Currently unused. Will be added in a future release.
        self.font_embedding_mode = PdfFontEmbeddingMode.EMBED_ALL

        # Page display mode
        # Currently unused. Will be added in a future release.
        self.page_mode = PdfPageMode.USE_OUTLINES

        # Color
        # Currently unused. Will be added in a future release.
        self.color_mode = ColorMode.NORMAL

        # Bookmarks and outlines
        # Ignored when OutlineOptions defines explicit bookmark levels.
        self.export_bookmarks_outline: bool = True
        self.outline_options: OutlineOptions = OutlineOptions()

        # Form fields
        # Currently unused. Will be added in a future release.
        self.preserve_form_fields: bool = False

        # Memory
        # Currently unused. Will be added in a future release.
        self.memory_optimization: bool = False

        # Zoom
        # 0 means "unspecified"; only read when zoom_behavior is ZOOM_FACTOR.
        self.zoom_factor: int = 0
        self.zoom_behavior = PdfZoomBehavior.NONE

        # Viewer preferences
        self.display_doc_title: bool = False


class OoxmlCompliance(IntEnum):
    """OOXML standards compliance level."""

    ECMA376_2006 = 0
    ISO29500_2008_TRANSITIONAL = 1
    ISO29500_2008_STRICT = 2


class CompressionLevel(IntEnum):
    """Compression level for OOXML files.

    DOCX and DOTX files are internally a ZIP-archive; this property
    controls the compression level of the archive.

    Note: FlatOpc files are not ZIP-archives, so this property does
    not affect FlatOpc files.
    """

    NORMAL = 0
    """Normal compression level. Default compression level used by Aspose.Words."""

    MAXIMUM = 1
    """Maximum compression level."""

    FAST = 2
    """Fast compression level."""

    SUPER_FAST = 3
    """Super Fast compression level."""


class Zip64Mode(IntEnum):
    """Controls when to use ZIP64 format extensions for OOXML files.
    OOXML files are ZIP archives subject to a 4 GB / 65 535-entry limit.
    ZIP64 extensions raise those limits to 2^64.
    """

    NEVER = 0
    IF_NECESSARY = 1
    ALWAYS = 2


class OoxmlSaveOptions:
    """Options for saving a document as DOCX (Office Open XML).

    * ``save_format``        — only ``"docx"`` is currently supported.
    * ``compression_level``  — :class:`CompressionLevel` integer enum
      (``NORMAL=0`` → ``compresslevel=6``, ``MAXIMUM=1`` → ``9``,
      ``FAST=2`` → ``3``, ``SUPER_FAST=3`` → ``1``).
    * ``compliance``         — ``ECMA376_2006`` and
      ``ISO29500_2008_TRANSITIONAL`` are treated identically (the
      document we emit conforms to both).  ``ISO29500_2008_STRICT``
      requires a different namespace and forbids transitional
      constructs, which we do not implement — the writer raises
      :class:`NotImplementedError` rather than silently producing
      non-strict output.
    * ``zip_64_mode``        — controls whether ZIP64 extensions are
      used in the output archive.  ``NEVER`` disables them (raises
      :class:`zipfile.LargeFileError` if the archive exceeds 4 GB),
      ``IF_NECESSARY`` enables them only when required (Python default),
      ``ALWAYS`` currently behaves like ``IF_NECESSARY`` — forcing ZIP64
      records unconditionally is not implemented.

    Usage::

        opts = OoxmlSaveOptions()
        opts.compression_level = CompressionLevel.MAXIMUM
        opts.zip_64_mode = Zip64Mode.IF_NECESSARY
        doc.save("output.docx", opts)
    """

    compliance = _EnumField(OoxmlCompliance, OoxmlCompliance.ECMA376_2006)
    compression_level = _EnumField(CompressionLevel, CompressionLevel.NORMAL)
    zip_64_mode = _EnumField(Zip64Mode, Zip64Mode.NEVER)

    def __init__(self, save_format: "int | str | None" = None):
        from aspose.words_foss.document import SaveFormat, _coerce_save_format

        # Currently unused. Will be added in a future release.
        self.save_format = (
            SaveFormat.DOCX if save_format is None
            else (_coerce_save_format(save_format) or save_format)
        )
        # ECMA376_2006 and ISO29500_2008_TRANSITIONAL are identical; STRICT raises an error.
        self.compliance = OoxmlCompliance.ECMA376_2006
        self.compression_level = CompressionLevel.NORMAL
        # ALWAYS behaves like IF_NECESSARY; ZIP64 records are never forced.
        self.zip_64_mode = Zip64Mode.NEVER
        self.pretty_format: bool = False


class MarkdownSaveOptions:
    """
    Options for saving documents as Markdown.

    Usage:
        opts = MarkdownSaveOptions()
        opts.table_content_alignment = TableContentAlignment.CENTER
        opts.list_export_mode = MarkdownListExportMode.PLAIN_TEXT
        doc.save("output.md", opts)
    """

    table_content_alignment = _EnumField(TableContentAlignment, TableContentAlignment.AUTO)
    list_export_mode = _EnumField(MarkdownListExportMode, MarkdownListExportMode.MARKDOWN_SYNTAX)
    link_export_mode = _EnumField(MarkdownLinkExportMode, MarkdownLinkExportMode.AUTO)
    export_as_html = _EnumField(MarkdownExportAsHtml, MarkdownExportAsHtml.NONE)
    empty_paragraph_export_mode = _EnumField(
        MarkdownEmptyParagraphExportMode, MarkdownEmptyParagraphExportMode.EMPTY_LINE
    )

    def __init__(self):
        self.table_content_alignment = TableContentAlignment.AUTO
        self.list_export_mode = MarkdownListExportMode.MARKDOWN_SYNTAX
        # Only takes effect when images_folder is set; otherwise base64 is always used.
        self.export_images_as_base64: bool = False
        self.images_folder: str = ""
        self.images_folder_alias: str = ""
        self.export_underline_formatting: bool = False
        # Only REFERENCE is implemented; AUTO and INLINE behave identically.
        self.link_export_mode = MarkdownLinkExportMode.AUTO
        # NON_COMPATIBLE_TABLES is not implemented and behaves like NONE.
        self.export_as_html = MarkdownExportAsHtml.NONE
        self.empty_paragraph_export_mode = MarkdownEmptyParagraphExportMode.EMPTY_LINE
        # Currently unused. Will be added in a future release.
        self.image_resolution: int = 96
        # Currently unused. Will be added in a future release.
        from aspose.words_foss.document import SaveFormat

        self.save_format = SaveFormat.MARKDOWN
        self.encoding: str = "utf-8"
        # Only applied between top-level blocks; inner line breaks are always "\n".
        self.paragraph_break: str = "\r\n"
