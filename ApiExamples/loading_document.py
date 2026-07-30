"""
API examples for loading documents from files and streams.

Mirrors the Aspose.Words for Python via .NET loading examples:
https://github.com/aspose-words/Aspose.Words-for-Python-via-.NET

Run:
    python ApiExamples/loading_document.py
    python -m pytest ApiExamples/ -v --rootdir=ApiExamples -c ApiExamples/pytest.ini
"""

import io
import sys
from pathlib import Path

_HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(_HERE))
sys.path.append(str(_HERE.parent))  # append: must not shadow an installed wheel

import aspose.words_foss as aw  # noqa: E402
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR  # noqa: E402


class LoadingDocument(DocsExamplesBase):
    """Demonstrates loading documents from files and streams."""

    def test_open_document_from_file(self):
        """Load a document from a file path string."""
        # ExStart:OpenDocument
        doc = aw.Document(MY_DIR + "test_full_article.docx")
        # ExEnd:OpenDocument
        assert len(doc.get_text()) > 0

    def test_open_document_from_file_doc_format(self):
        """Load a .doc file from a file path string."""
        doc = aw.Document(MY_DIR + "test_bold.doc")
        assert len(doc.get_text()) > 0

    def test_open_from_stream(self):
        """Load a document from a binary stream."""
        # ExStart:OpeningFromStream
        with io.FileIO(MY_DIR + "test_full_article.docx") as stream:
            doc = aw.Document(stream)
        # ExEnd:OpeningFromStream
        assert len(doc.get_text()) > 0

    def test_open_from_stream_doc_format(self):
        """Load a .doc document from a stream (auto-detected)."""
        with io.FileIO(MY_DIR + "test_bold.doc") as stream:
            doc = aw.Document(stream)
        assert len(doc.get_text()) > 0

    def test_open_from_stream_with_load_options(self):
        """Load a document from a stream with explicit LoadOptions."""
        opts = aw.loading.LoadOptions()
        opts.load_format = aw.LoadFormat.DOCX

        with io.FileIO(MY_DIR + "test_full_article.docx") as stream:
            doc = aw.Document(stream, opts)
        assert len(doc.get_text()) > 0

    def test_open_from_bytes_io(self):
        """Load from a BytesIO stream (in-memory)."""
        file_path = Path(MY_DIR + "test_full_article.docx")
        raw = file_path.read_bytes()

        with io.BytesIO(raw) as stream:
            doc = aw.Document(stream)
        assert len(doc.get_text()) > 0

    def test_open_and_save(self):
        """Load from stream, then save to file."""
        with io.FileIO(MY_DIR + "test_full_article.docx") as stream:
            doc = aw.Document(stream)

        doc.save(ARTIFACTS_DIR + "LoadingDocument.OpenAndSave.md", aw.SaveFormat.MARKDOWN)
        assert Path(ARTIFACTS_DIR + "LoadingDocument.OpenAndSave.md").exists()


if __name__ == "__main__":
    examples = LoadingDocument()

    print("=== Loading Document Examples ===\n")

    print("  Open from file...")
    examples.test_open_document_from_file()
    print("  Open from file (.doc)...")
    examples.test_open_document_from_file_doc_format()
    print("  Open from stream...")
    examples.test_open_from_stream()
    print("  Open from stream (.doc)...")
    examples.test_open_from_stream_doc_format()
    print("  Open from stream with LoadOptions...")
    examples.test_open_from_stream_with_load_options()
    print("  Open from BytesIO...")
    examples.test_open_from_bytes_io()
    print("  Open and save...")
    examples.test_open_and_save()

    print("\n  All examples passed.")
    print(f"  Output files are in: {ARTIFACTS_DIR}")
