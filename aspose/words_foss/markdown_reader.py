"""
Markdown (.md) import entry point.

``MarkdownReader`` is the reader that ``reader_factory.create_reader``
wires up for the ``.md`` suffix. ``to_light_document()`` parses the loaded
source through the real Markdown block parser and ``MarkdownReaderContext``
(``md_import.parse_and_build``), turning headings, lists, emphasis, tables,
footnotes, and raw HTML into proper LDM nodes instead of literal text.
Loading (``load_file``/``load_stream``/``load_bytes``) and the legacy
``_iterate_body_elements``/``_get_list_format``/``_get_numbering_info``
trio still delegate to ``text_reader.MarkdownFileReader``, which retains
its original plain-text behaviour for that vestigial interface.
"""

from pathlib import Path
from typing import Optional, Union, BinaryIO, Iterator

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_reader import ParagraphData, TableData, NumberingInfo
from aspose.words_foss.md_import import parse_and_build
from aspose.words_foss.text_reader import MarkdownFileReader


class MarkdownReader:
    """
    Reads Markdown (.md) files, producing the same data structures as
    DocumentReader (for .docx) and DocFileReader (for .doc).

    ``to_light_document()`` runs the loaded source through the real
    Markdown parser; loading and the legacy body-element iteration still
    delegate to ``MarkdownFileReader``.
    """

    def __init__(self) -> None:
        self._delegate = MarkdownFileReader()
        self.preserve_empty_lines: bool = False
        self._base_dir: Optional[Path] = None

    @property
    def encoding(self) -> str:
        """Text encoding used to decode the Markdown source."""
        return self._delegate.encoding

    @encoding.setter
    def encoding(self, value: str) -> None:
        self._delegate.encoding = value

    def load_file(self, filepath: Union[str, Path]) -> None:
        """Load a .md file from *filepath*."""
        self._delegate.load_file(filepath)
        self._base_dir = Path(filepath).resolve().parent

    def load_stream(self, stream: BinaryIO) -> None:
        """Load from a binary stream."""
        self._delegate.load_stream(stream)

    def load_bytes(self, data: bytes) -> None:
        """Load from raw bytes."""
        self._delegate.load_bytes(data)

    def _iterate_body_elements(
        self,
    ) -> Iterator[Union[ParagraphData, TableData]]:
        """Iterate over document body elements in order."""
        return self._delegate._iterate_body_elements()

    def _get_list_format(self, num_id: int, level: int) -> tuple[str, int]:
        """Get list format and start number for a given list id and level."""
        return self._delegate._get_list_format(num_id, level)

    def _get_numbering_info(self, num_id: int) -> Optional[NumberingInfo]:
        """Get numbering info (for compatibility with DocumentReader)."""
        return self._delegate._get_numbering_info(num_id)

    def to_light_document(self) -> ldm.Document:
        """Build a light_document_model.Document from the loaded Markdown."""
        text = self._delegate.text
        if not text:
            return self._delegate.to_light_document()
        return parse_and_build(text, self.preserve_empty_lines, base_dir=self._base_dir)
