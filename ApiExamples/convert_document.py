"""
API examples for document conversion between all supported formats.

Input:  DOCX, DOC, RTF, TXT, MD
Output: Markdown, PDF, TXT

Run:
    python ApiExamples/convert_document.py
    python -m pytest ApiExamples/ -v --rootdir=ApiExamples -c ApiExamples/pytest.ini
"""

import sys
from pathlib import Path

_HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(_HERE))
sys.path.append(str(_HERE.parent))  # append: must not shadow an installed wheel

import aspose.words_foss as aw  # noqa: E402
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR  # noqa: E402

# Each input format with a representative test file
INPUT_FILES = {
    "docx": "test_full_article.docx",
    "doc": "test_full_article.doc",
    "rtf": "test_full_article.rtf",
    "txt": "test_plain.txt",
    "md": "test_markdown.md",
}

ALL_OUTPUT_FORMATS = {
    "md": aw.SaveFormat.MARKDOWN,
    "pdf": aw.SaveFormat.PDF,
    "txt": aw.SaveFormat.TEXT,
}


class ConvertDocument(DocsExamplesBase):
    """Demonstrates every supported input -> output format pair."""

    def test_all_conversions(self):
        """Convert each input format to every output format."""
        for in_label, in_file in INPUT_FILES.items():
            for out_ext, out_fmt in ALL_OUTPUT_FORMATS.items():
                out_file = f"ConvertDocument.{in_label}_to_{out_ext}.{out_ext}"
                self.convert(in_file, out_file, save_format=out_fmt)

    def test_get_text_from_all_formats(self):
        """Extract plain text via get_text() from every input format."""
        for in_label, in_file in INPUT_FILES.items():
            doc = aw.Document(MY_DIR + in_file)
            text = doc.get_text()
            assert len(text) > 0, f"get_text() returned empty for {in_file}"


if __name__ == "__main__":
    examples = ConvertDocument()

    print("=== Document Conversion Examples ===\n")
    for in_label, in_file in INPUT_FILES.items():
        for out_ext in ALL_OUTPUT_FORMATS:
            print(f"  {in_label} -> {out_ext}")
    examples.test_all_conversions()
    print("  Done.\n")

    print("=== get_text() Examples ===")
    examples.test_get_text_from_all_formats()
    print("  Done.\n")

    print(f"Output files are in: {ARTIFACTS_DIR}")
