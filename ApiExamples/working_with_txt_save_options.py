"""
API examples for plain-text export.

Mirrors the Aspose.Words for Python DocsExamples:
  file_formats_and_conversions/save_options/working_with_txt_save_options.py

Run:
    python ApiExamples/working_with_txt_save_options.py
    python -m pytest ApiExamples/ -v --rootdir=ApiExamples -c ApiExamples/pytest.ini
"""

import sys
from pathlib import Path

_HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(_HERE))
sys.path.append(str(_HERE.parent))  # append: must not shadow an installed wheel

import aspose.words_foss as aw  # noqa: E402
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR  # noqa: E402


class WorkingWithTxtSaveOptions(DocsExamplesBase):

    def test_text_from_all_input_formats(self):
        """Save DOCX, DOC, RTF, MD as plain text."""
        inputs = {
            "docx": "test_full_article.docx",
            "doc": "test_full_article.doc",
            "rtf": "test_full_article.rtf",
            "md": "test_markdown.md",
        }
        for label, filename in inputs.items():
            self.convert(filename, f"TxtSaveOptions.from_{label}.txt", save_format=aw.SaveFormat.TEXT)

    def test_get_text(self):
        """Extract text via get_text() without saving to file."""
        # ExStart:GetText
        doc = aw.Document(MY_DIR + "test_full_article.docx")
        text = doc.get_text()
        assert len(text) > 0
        # ExEnd:GetText


if __name__ == "__main__":
    examples = WorkingWithTxtSaveOptions()

    print("=== TXT Save Options Examples ===\n")
    examples.test_text_from_all_input_formats()
    print("  text_from_all_input_formats: Done.")
    examples.test_get_text()
    print("  get_text: Done.")
    print(f"\nOutput files are in: {ARTIFACTS_DIR}")
