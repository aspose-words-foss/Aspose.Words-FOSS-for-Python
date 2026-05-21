"""
save options for markdown and PDF export.
Provides MarkdownSaveOptions, PdfSaveOptions, and related enums
matching the public save API.
"""


class TableContentAlignment:
    """Table content alignment options."""

    AUTO = "auto"
    LEFT = "left"
    CENTER = "center"
    RIGHT = "right"


class MarkdownListExportMode:
    """List export mode options."""

    MARKDOWN_SYNTAX = "markdown_syntax"
    PLAIN_TEXT = "plain_text"


class MarkdownLinkExportMode:
    """Link export mode options."""

    AUTO = "auto"
    INLINE = "inline"
    REFERENCE = "reference"


class MarkdownExportAsHtml:
    """Controls which elements are exported as raw HTML."""

    NONE = "none"
    TABLES = "tables"
    NON_COMPATIBLE_TABLES = "non_compatible_tables"


class MarkdownEmptyParagraphExportMode:
    """Controls how empty paragraphs are exported."""

    EMPTY_LINE = "empty_line"
    MARKDOWN_HARD_LINE_BREAK = "markdown_hard_line_break"
    NONE = "none"


class PdfCompliance:
    """PDF standards compliance level."""

    PDF17 = "pdf17"
    PDF20 = "pdf20"
    PDF_A1A = "pdf_a1a"
    PDF_A1B = "pdf_a1b"
    PDF_A2A = "pdf_a2a"
    PDF_A2U = "pdf_a2u"
    PDF_A4 = "pdf_a4"
    PDF_UA1 = "pdf_ua1"


class PdfTextCompression:
    """Text compression in PDF."""

    NONE = "none"
    FLATE = "flate"


class PdfImageCompression:
    """Image compression in PDF."""

    AUTO = "auto"
    JPEG = "jpeg"


class PdfPageMode:
    """PDF page display mode."""

    USE_NONE = "use_none"
    USE_OUTLINES = "use_outlines"
    USE_THUMBS = "use_thumbs"
    FULL_SCREEN = "full_screen"


class PdfZoomBehavior:
    """Mirrors public API."""

    NONE = "none"
    ZOOM_FACTOR = "zoom_factor"
    FIT_PAGE = "fit_page"
    FIT_WIDTH = "fit_width"
    FIT_HEIGHT = "fit_height"
    FIT_BOX = "fit_box"


class PdfFontEmbeddingMode:
    """Font embedding mode in PDF."""

    EMBED_ALL = "embed_all"
    EMBED_NONSTANDARD = "embed_nonstandard"
    EMBED_NONE = "embed_none"


class ColorMode:
    """Color rendering mode."""

    NORMAL = "normal"
    GRAYSCALE = "grayscale"


class OutlineOptions:
    """Controls how outlines (bookmarks panel) are generated in the PDF.
    Note: the fpdf2 backend always fills gaps between non-contiguous
    outline levels (e.g. H1 followed by H3 inserts an empty H2 entry)
    because the library requires a contiguous hierarchy.
    ``create_missing_outline_levels`` is
    but gap-filling is always active regardless of its value.
    """

    def __init__(self):
        self.headings_outline_levels: int = 0
        self.expanded_outline_levels: int = 0
        self.default_bookmarks_outline_level: int = 0
        self.bookmarks_outline_levels: dict[str, int] = {}
        self.create_outlines_for_headings_in_tables: bool = False
        # Always active in the fpdf2 backend.
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

    def __init__(self):
        # PDF standard compliance
        self.compliance: str = PdfCompliance.PDF17

        # Document structure
        self.export_document_structure: bool = False

        # Image options
        self.image_compression: str = PdfImageCompression.AUTO
        self.jpeg_quality: int = 100

        # Text compression
        self.text_compression: str = PdfTextCompression.FLATE

        # Font embedding
        self.embed_full_fonts: bool = False
        self.use_core_fonts: bool = False
        self.font_embedding_mode: str = PdfFontEmbeddingMode.EMBED_ALL

        # Page display mode
        self.page_mode: str = PdfPageMode.USE_NONE

        # Color
        self.color_mode: str = ColorMode.NORMAL

        # Bookmarks and outlines
        self.export_bookmarks_outline: bool = True
        self.outline_options: OutlineOptions = OutlineOptions()

        # Form fields
        self.preserve_form_fields: bool = False

        # Memory
        self.memory_optimization: bool = False

        # Zoom
        self.zoom_factor: int = 100
        self.zoom_behavior: str = PdfZoomBehavior.NONE

        # Viewer preferences
        self.display_doc_title: bool = False


class OoxmlCompliance:
    """OOXML standards compliance level."""

    ECMA376_2006 = "ecma376_2006"
    ISO29500_2008_TRANSITIONAL = "iso29500_2008_transitional"
    ISO29500_2008_STRICT = "iso29500_2008_strict"


class CompressionLevel:
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


class Zip64Mode:
    """Controls when to use ZIP64 format extensions for OOXML files.
    OOXML files are ZIP archives subject to a 4 GB / 65 535-entry limit.
    ZIP64 extensions raise those limits to 2^64.
    """

    NEVER = "never"
    IF_NECESSARY = "if_necessary"
    ALWAYS = "always"


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
      ``ALWAYS`` enables them unconditionally.

    Usage::

        opts = OoxmlSaveOptions()
        opts.compression_level = CompressionLevel.MAXIMUM
        opts.zip_64_mode = Zip64Mode.IF_NECESSARY
        doc.save("output.docx", opts)
    """

    def __init__(self, save_format: str = "docx"):
        self.save_format: str = save_format
        self.compliance: str = OoxmlCompliance.ECMA376_2006
        self.compression_level: int = CompressionLevel.NORMAL
        self.zip_64_mode: str = Zip64Mode.NEVER
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

    def __init__(self):
        self.table_content_alignment: str = TableContentAlignment.AUTO
        self.list_export_mode: str = MarkdownListExportMode.MARKDOWN_SYNTAX
        self.export_images_as_base64: bool = False
        self.images_folder: str = ""
        self.images_folder_alias: str = ""
        self.export_underline_formatting: bool = False
        self.link_export_mode: str = MarkdownLinkExportMode.AUTO
        self.export_as_html: str = MarkdownExportAsHtml.NONE
        self.empty_paragraph_export_mode: str = MarkdownEmptyParagraphExportMode.EMPTY_LINE
        self.image_resolution: int = 96
        self.save_format: str = "markdown"
        self.encoding: str = "utf-8"
        self.paragraph_break: str = "\n"
