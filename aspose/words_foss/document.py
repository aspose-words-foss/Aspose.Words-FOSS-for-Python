"""
Document class.
supporting .doc, .docx, .rtf, .txt, and .md input formats.

Usage:
    import aspose.words_foss as aw

    doc = aw.Document("input.docx")
    doc.save("output.md", aw.SaveFormat.MARKDOWN)

    # Or with options:
    opts = aw.saving.MarkdownSaveOptions()
    opts.table_content_alignment = aw.saving.TableContentAlignment.CENTER
    doc.save("output.md", opts)
"""

from pathlib import Path
from typing import Optional, Union, BinaryIO

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.models import ConversionOptions
from aspose.words_foss.reader_factory import create_reader
from aspose.words_foss.saving import (
    MarkdownSaveOptions,
    OoxmlSaveOptions,
    PdfSaveOptions,
)


class LoadFormat:
    """Document load format constants."""

    AUTO = "auto"
    DOC = "doc"
    DOCX = "docx"
    RTF = "rtf"
    TEXT = "text"
    MARKDOWN = "markdown"


class SaveFormat:
    """Document save format constants."""

    MARKDOWN = "markdown"
    DOC = "doc"
    DOCX = "docx"
    TEXT = "text"
    PDF = "pdf"


class Document:
    """
    Represents a Word document.

    Loads the file at construction time
    and populates the internal Light Document Model immediately.

    Usage:
        doc = Document("input.docx")
        doc.save("output.md", SaveFormat.MARKDOWN)
    """

    def __init__(
        self,
        filepath: Optional[Union[str, Path]] = None,
        *,
        stream: Optional[BinaryIO] = None,
        data: Optional[bytes] = None,
    ):
        """Initialize Document, loading from a file, stream, or bytes.

        At least one source must be provided. The document is parsed immediately
        and the Light Document Model is populated.

        Args:
            filepath: Path to a .doc, .docx, .rtf, .txt, or .md file to load.
            stream: A binary stream containing the document data (DOCX only).
            data: Raw bytes of the document (DOCX only).
        """
        self._document: Optional[ldm.Document] = None

        if filepath is not None:
            filepath = Path(filepath)
            suffix = filepath.suffix.lower()
            reader = create_reader(suffix)
            reader.load_file(filepath)
            self._document = reader.to_light_document()
        elif stream is not None:
            reader = create_reader(".docx")
            reader.load_stream(stream)
            self._document = reader.to_light_document()
        elif data is not None:
            reader = create_reader(".docx")
            reader.load_bytes(data)
            self._document = reader.to_light_document()

    @property
    def light_document_model(self) -> ldm.Document:
        """Access the internal Light Document Model.

        Returns:
            The LDM Document populated during construction.

        Raises:
            ValueError: If no document was loaded.
        """
        if self._document is None:
            raise ValueError("No document loaded. Provide a filepath to Document().")
        return self._document

    @property
    def sections(self) -> "list[ldm.Section]":
        """All document sections."""
        return self.light_document_model.sections

    @property
    def first_section(self) -> "Optional[ldm.Section]":
        """The first section of the document.

        Returns ``None`` if the document has no sections.
        """
        secs = self.light_document_model.sections
        return secs[0] if secs else None

    @property
    def last_section(self) -> "Optional[ldm.Section]":
        """The last section of the document.

        Returns ``None`` if the document has no sections.
        """
        secs = self.light_document_model.sections
        return secs[-1] if secs else None

    @property
    def styles(self) -> "list[ldm.Style]":
        """All document styles."""
        return self.light_document_model.styles

    @property
    def lists(self) -> "list[ldm.DocList]":
        """All document list definitions."""
        return self.light_document_model.lists

    @property
    def page_count(self) -> int:
        """Estimated page count."""
        return self.light_document_model.page_count

    def get_text(self) -> str:
        """Extract plain text from the loaded document.

        Returns the raw text content of the document with paragraphs
        separated by newlines. Tables are included with cell text.

        Returns:
            The plain text content of the document.
        """
        return self.light_document_model.text

    def save(
        self,
        output_path: Union[str, Path],
        save_format_or_options: Union[
            str, MarkdownSaveOptions, PdfSaveOptions, OoxmlSaveOptions, None
        ] = None,
    ) -> None:
        """Save the document to the specified format.

        Args:
            output_path: Path to save the output file.
            save_format_or_options: A SaveFormat constant, MarkdownSaveOptions,
                PdfSaveOptions, or OoxmlSaveOptions instance.
        """
        doc = self.light_document_model  # validates that a document is loaded
        output_path = Path(output_path)
        suffix = output_path.suffix.lower()

        if isinstance(save_format_or_options, MarkdownSaveOptions):
            self._save_as_markdown(output_path, doc, save_format_or_options)
        elif isinstance(save_format_or_options, PdfSaveOptions):
            self._save_as_pdf(output_path, doc, save_format_or_options)
        elif isinstance(save_format_or_options, OoxmlSaveOptions):
            self._save_as_docx(output_path, doc, save_format_or_options)
        elif save_format_or_options == SaveFormat.MARKDOWN or suffix == ".md":
            self._save_as_markdown(output_path, doc)
        elif save_format_or_options == SaveFormat.TEXT or suffix == ".txt":
            self._save_as_text(output_path, doc)
        elif save_format_or_options == SaveFormat.PDF or suffix == ".pdf":
            self._save_as_pdf(output_path, doc)
        elif save_format_or_options == SaveFormat.DOCX or suffix == ".docx":
            self._save_as_docx(output_path, doc)
        else:
            raise ValueError(
                f"Unsupported save format: {save_format_or_options}. "
                f"Supported formats: Markdown, Text, PDF, DOCX."
            )

    def _save_as_markdown(
        self,
        output_path: Path,
        doc: ldm.Document,
        options: Optional[MarkdownSaveOptions] = None,
    ) -> None:
        """Convert the loaded LDM to Markdown and save."""
        from aspose.words_foss.md_writer import LdmMarkdownWriter

        conversion_opts = ConversionOptions()
        encoding = "utf-8"
        if options is not None:
            if options.export_underline_formatting:
                conversion_opts.export_underline = True
            conversion_opts.table_content_alignment = options.table_content_alignment
            conversion_opts.list_export_mode = options.list_export_mode
            conversion_opts.link_export_mode = options.link_export_mode
            conversion_opts.export_as_html = options.export_as_html
            conversion_opts.empty_paragraph_export_mode = options.empty_paragraph_export_mode
            conversion_opts.export_images_as_base64 = options.export_images_as_base64
            conversion_opts.images_folder = options.images_folder
            conversion_opts.images_folder_alias = options.images_folder_alias
            conversion_opts.paragraph_break = options.paragraph_break
            encoding = options.encoding

        writer = LdmMarkdownWriter(conversion_opts)
        markdown = writer.write(doc, output_path=output_path)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        # ``newline=""`` disables Python's universal-newline translation
        # so a caller's choice of ``paragraph_break`` ("\r\n", "\n",
        # ...) round-trips verbatim instead of being doubled into
        # "\r\r\n" on Windows (where text mode would otherwise rewrite
        # every ``\n`` to ``\r\n``).
        output_path.write_text(markdown, encoding=encoding, newline="")

    def _save_as_text(self, output_path: Path, doc: ldm.Document) -> None:
        """Extract plain text from the LDM and save."""
        output_path.parent.mkdir(parents=True, exist_ok=True)
        # Preserve the LDM's exact line endings on Windows by disabling
        # universal-newline translation.
        output_path.write_text(doc.text, encoding="utf-8", newline="")

    def _save_as_pdf(
        self,
        output_path: Path,
        doc: ldm.Document,
        options: Optional[PdfSaveOptions] = None,
    ) -> None:
        """Convert the loaded LDM to PDF and save."""
        from aspose.words_foss.pdf_writer import LdmPdfWriter

        writer = LdmPdfWriter(options)
        writer.write(doc, output_path)

    def _save_as_docx(
        self,
        output_path: Path,
        doc: ldm.Document,
        options: Optional[OoxmlSaveOptions] = None,
    ) -> None:
        """Convert the loaded LDM to DOCX and save."""
        from aspose.words_foss.docx_writer import LdmDocxWriter

        writer = LdmDocxWriter(options)
        writer.write(doc, output_path)
