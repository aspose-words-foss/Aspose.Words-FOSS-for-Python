"""Reader factory and protocol for document format readers.

Provides a common Protocol defining the reader interface and a factory
function that returns the appropriate reader for a given file extension.
"""

from __future__ import annotations

from pathlib import Path
from typing import BinaryIO, Protocol, Union, runtime_checkable

from aspose.words_foss import light_document_model as ldm


@runtime_checkable
class DocumentFormatReader(Protocol):
    """Protocol defining the interface all document readers must implement."""

    def load_file(self, filepath: Union[str, Path]) -> None: ...

    def load_stream(self, stream: BinaryIO) -> None: ...

    def load_bytes(self, data: bytes) -> None: ...

    def to_light_document(self) -> ldm.Document: ...


def create_reader(suffix: str) -> DocumentFormatReader:
    """Create the appropriate reader based on file extension.

    Args:
        suffix: Lowercase file extension including dot (e.g. ".docx").

    Returns:
        A reader instance implementing the DocumentFormatReader protocol.
    """
    if suffix == ".md":
        from aspose.words_foss.text_reader import MarkdownFileReader

        return MarkdownFileReader()
    elif suffix == ".txt":
        from aspose.words_foss.text_reader import TextFileReader

        return TextFileReader()
    elif suffix == ".doc":
        from aspose.words_foss.doc_reader import DocFileReader

        return DocFileReader()
    elif suffix == ".rtf":
        from aspose.words_foss.rtf_reader import RtfFileReader

        return RtfFileReader()
    else:
        from aspose.words_foss.docx_reader import DocumentReader

        return DocumentReader()
