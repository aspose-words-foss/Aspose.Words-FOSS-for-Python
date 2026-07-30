"""
Plain-text and Markdown file readers.

Provides TextFileReader (for .txt) and MarkdownFileReader (for .md) that
implement the same public interface as DocumentReader and DocFileReader,
producing a light_document_model.Document.

Both readers parse content line-by-line into Paragraph nodes, which then
flow through the standard Converter pipeline — identical to how DOCX
and DOC formats are handled via the Document class. Real Markdown
structure (headings, lists, quotes, tables, inline emphasis/links) is
built by markdown_reader.MarkdownReader instead, which reader_factory
wires up for the ".md" suffix; MarkdownFileReader here stays a plain
per-block literal-text reader, reused by MarkdownReader only for loading.
"""

from pathlib import Path
from typing import Optional, Union, BinaryIO, Iterator

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_reader import (
    ParagraphData,
    RunData,
    TableData,
    NumberingInfo,
)



class TextFileReader:
    """
    Reads plain-text (.txt) files and yields one Paragraph per line.

    The reader preserves the original content without any formatting
    interpretation — the converter will output each line as a plain
    Markdown paragraph, effectively echoing the input.
    """

    def __init__(self) -> None:
        self._text: Optional[str] = None
        self.encoding: str = "utf-8"

    def load_file(self, filepath: Union[str, Path]) -> None:
        """Load a .txt file from *filepath*."""
        self._text = Path(filepath).read_text(encoding=self.encoding)

    def load_stream(self, stream: BinaryIO) -> None:
        """Load from a binary stream."""
        self._text = stream.read().decode(self.encoding)

    def load_bytes(self, data: bytes) -> None:
        """Load from raw bytes."""
        self._text = data.decode(self.encoding)

    def _iterate_body_elements(
        self,
    ) -> Iterator[Union[ParagraphData, TableData]]:
        """Yield one ``ParagraphData`` for each line of text."""
        if self._text is None:
            return
        for line in self._text.splitlines():
            yield ParagraphData(
                text=line,
                runs=[RunData(text=line)],
            )

    def _get_list_format(self, num_id: int, level: int) -> tuple[str, int]:
        """Plain text has no list metadata — return bullet default."""
        return ("bullet", 1)

    def _get_numbering_info(self, num_id: int) -> Optional[NumberingInfo]:
        """No numbering information in plain text."""
        return None

    def to_light_document(self) -> ldm.Document:
        """Build a light_document_model.Document from the loaded text."""
        doc = ldm.Document()
        children: list[ldm.Paragraph | ldm.Table | ldm.UnknownNode] = []

        if self._text is not None:
            for line in self._text.splitlines():
                para = ldm.Paragraph()
                run = ldm.Run()
                run.text = line
                para._children = [run]
                children.append(para)

        sec = ldm.Section()
        sec.body = ldm.Body(children=children)
        doc.sections = [sec]
        return doc


class MarkdownFileReader:
    """
    Reads Markdown (.md) files and yields one Paragraph per line.

    Follows the same reader interface as DocumentReader and
    DocFileReader so that .md files flow through the standard
    Converter pipeline via the Document class. Loading is reused by
    MarkdownReader (see markdown_reader.py), which builds real Markdown
    structure from the loaded text instead of this class's literal
    to_light_document().
    """

    def __init__(self) -> None:
        self._text: Optional[str] = None
        self.encoding: str = "utf-8"

    def load_file(self, filepath: Union[str, Path]) -> None:
        """Load a .md file from *filepath*."""
        self.load_bytes(Path(filepath).read_bytes())

    def load_stream(self, stream: BinaryIO) -> None:
        """Load from a binary stream."""
        self.load_bytes(stream.read())

    def load_bytes(self, data: bytes) -> None:
        """Load from raw bytes."""
        text = data.decode(self.encoding)
        if text.startswith("\ufeff"):
            text = text[1:]
        self._text = text

    @property
    def text(self) -> Optional[str]:
        """Raw loaded Markdown source, or ``None`` before ``load_*`` is called."""
        return self._text

    def _iterate_body_elements(
        self,
    ) -> Iterator[Union[ParagraphData, TableData]]:
        """Yield one ``ParagraphData`` for each line."""
        if self._text is None:
            return
        for line in self._text.splitlines():
            yield ParagraphData(
                text=line,
                runs=[RunData(text=line)],
            )

    def _get_list_format(self, num_id: int, level: int) -> tuple[str, int]:
        """No binary list metadata — return bullet default."""
        return ("bullet", 1)

    def _get_numbering_info(self, num_id: int) -> Optional[NumberingInfo]:
        """No numbering information in Markdown files."""
        return None

    def to_light_document(self) -> ldm.Document:
        """Build a light_document_model.Document from the loaded Markdown.

        Lines are grouped into blank-line-separated blocks so that the
        writer can emit blank lines between blocks without breaking
        tight constructs (lists, code fences, etc.) that span multiple
        consecutive lines.
        """
        doc = ldm.Document()
        children: list[ldm.Paragraph | ldm.Table | ldm.UnknownNode] = []

        if self._text is not None:
            lines = self._text.splitlines()
            block: list[str] = []
            for line in lines:
                if line.strip() == "":
                    if block:
                        para = ldm.Paragraph()
                        block_text = "\n".join(block)
                        run = ldm.Run()
                        run.text = block_text
                        para._children = [run]
                        children.append(para)
                        block = []
                else:
                    block.append(line)
            if block:
                para = ldm.Paragraph()
                block_text = "\n".join(block)
                run = ldm.Run()
                run.text = block_text
                para._children = [run]
                children.append(para)

        sec = ldm.Section()
        sec.body = ldm.Body(children=children)
        doc.sections = [sec]
        return doc
