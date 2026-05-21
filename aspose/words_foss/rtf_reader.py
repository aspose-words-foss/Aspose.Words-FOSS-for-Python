"""
RTF File Reader - Reads RTF files saved in OLE2/DOC binary format.

Many RTF files produced by Microsoft Word are actually stored in the
Word 97-2003 binary format (OLE2). This reader delegates to DocFileReader
for parsing, providing transparent RTF support alongside DOC and DOCX.
"""

from pathlib import Path
from typing import Optional, Union, BinaryIO, Iterator

from aspose.words_foss.doc_reader import DocFileReader
from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_reader import (
    ParagraphData,
    TableData,
    NumberingInfo,
)



class RtfFileReader:
    """
    Reads RTF files (OLE2-format) and produces the same data structures
    as DocumentReader (for .docx) and DocFileReader (for .doc).

    Delegates to DocFileReader since many .rtf files from Word
    are stored in the OLE2 binary format.
    """

    def __init__(self):
        self._delegate = DocFileReader()

    def load_file(self, filepath: Union[str, Path]) -> None:
        """Load .rtf from file path."""
        self._delegate.load_file(filepath)

    def load_stream(self, stream: BinaryIO) -> None:
        """Load .rtf from stream."""
        self._delegate.load_stream(stream)

    def load_bytes(self, data: bytes) -> None:
        """Load .rtf from bytes."""
        self._delegate.load_bytes(data)

    def _iterate_body_elements(self) -> Iterator[Union[ParagraphData, TableData]]:
        """Iterate over document body elements in order."""
        return self._delegate._iterate_body_elements()

    def _get_list_format(self, num_id: int, level: int) -> tuple[str, int]:
        """Get list format and start number for a given list id and level."""
        return self._delegate._get_list_format(num_id, level)

    def _get_numbering_info(self, num_id: int) -> Optional[NumberingInfo]:
        """Get numbering info (for compatibility with DocumentReader)."""
        return self._delegate._get_numbering_info(num_id)

    def to_light_document(self) -> ldm.Document:
        """Build a light_document_model.Document from the loaded RTF file."""
        return self._delegate.to_light_document()
