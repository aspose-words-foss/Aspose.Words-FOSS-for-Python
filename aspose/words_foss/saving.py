"""
Aspose.Words-compatible save options for markdown and PDF export.

Provides MarkdownSaveOptions, PdfSaveOptions, and related enums
that mirror the Aspose.Words saving API.
"""


class TableContentAlignment:
    """Table content alignment options (mirrors aspose.words.saving.TableContentAlignment)."""

    AUTO = "auto"
    LEFT = "left"
    CENTER = "center"
    RIGHT = "right"


class MarkdownListExportMode:
    """List export mode options (mirrors aspose.words.saving.MarkdownListExportMode)."""

    MARKDOWN_SYNTAX = "markdown_syntax"
    PLAIN_TEXT = "plain_text"


class MarkdownLinkExportMode:
    """Link export mode options (mirrors aspose.words.saving.MarkdownLinkExportMode)."""

    AUTO = "auto"
    INLINE = "inline"
    REFERENCE = "reference"


class MarkdownExportAsHtml:
    """Controls which elements are exported as raw HTML (mirrors aspose.words.saving.MarkdownExportAsHtml)."""

    NONE = "none"
    TABLES = "tables"
    NON_COMPATIBLE_TABLES = "non_compatible_tables"


class MarkdownEmptyParagraphExportMode:
    """Controls how empty paragraphs are exported (mirrors aspose.words.saving.MarkdownEmptyParagraphExportMode)."""

    EMPTY_LINE = "empty_line"
    MARKDOWN_HARD_LINE_BREAK = "markdown_hard_line_break"
    NONE = "none"


class PdfCompliance:
    """PDF standards compliance level (mirrors aspose.words.saving.PdfCompliance)."""

    PDF17 = "pdf17"
    PDF20 = "pdf20"
    PDF_A1A = "pdf_a1a"
    PDF_A1B = "pdf_a1b"
    PDF_A2A = "pdf_a2a"
    PDF_A2U = "pdf_a2u"
    PDF_A4 = "pdf_a4"
    PDF_UA1 = "pdf_ua1"


class PdfTextCompression:
    """Text compression in PDF (mirrors aspose.words.saving.PdfTextCompression)."""

    NONE = "none"
    FLATE = "flate"


class PdfImageCompression:
    """Image compression in PDF (mirrors aspose.words.saving.PdfImageCompression)."""

    AUTO = "auto"
    JPEG = "jpeg"


class PdfPageMode:
    """PDF page display mode (mirrors aspose.words.saving.PdfPageMode)."""

    USE_NONE = "use_none"
    USE_OUTLINES = "use_outlines"
    USE_THUMBS = "use_thumbs"
    FULL_SCREEN = "full_screen"


class PdfFontEmbeddingMode:
    """Font embedding mode in PDF (mirrors aspose.words.saving.PdfFontEmbeddingMode)."""

    EMBED_ALL = "embed_all"
    EMBED_NONSTANDARD = "embed_nonstandard"
    EMBED_NONE = "embed_none"


class ColorMode:
    """Color rendering mode (mirrors aspose.words.saving.ColorMode)."""

    NORMAL = "normal"
    GRAYSCALE = "grayscale"


class PdfSaveOptions:
    """
    Options for saving documents as PDF.

    Mirrors the aspose.words.saving.PdfSaveOptions API (subset).

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

        # Form fields
        self.preserve_form_fields: bool = False

        # Memory
        self.memory_optimization: bool = False

        # Zoom
        self.zoom_factor: int = 100


class MarkdownSaveOptions:
    """
    Options for saving documents as Markdown.

    Mirrors the aspose.words.saving.MarkdownSaveOptions API.

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
