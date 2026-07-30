"""
API examples for MarkdownSaveOptions.

Demonstrates options that are actually applied during conversion:
export_underline_formatting, encoding, paragraph_break.

Mirrors the Aspose.Words for Python DocsExamples:
  file_formats_and_conversions/save_options/working_with_markdown_save_options.py

Run:
    python ApiExamples/working_with_markdown_save_options.py
    python -m pytest ApiExamples/ -v --rootdir=ApiExamples -c ApiExamples/pytest.ini
"""

import sys
from pathlib import Path

_HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(_HERE))
sys.path.append(str(_HERE.parent))  # append: must not shadow an installed wheel

import aspose.words_foss as aw  # noqa: E402
from docs_examples_base import DocsExamplesBase, ARTIFACTS_DIR  # noqa: E402


class WorkingWithMarkdownSaveOptions(DocsExamplesBase):

    def test_export_underline_formatting(self):
        # ExStart:ExportUnderlineFormatting
        save_options = aw.saving.MarkdownSaveOptions()
        save_options.export_underline_formatting = True

        self.convert(
            "test_underline_text.docx",
            "MarkdownSaveOptions.export_underline.md",
            save_options=save_options,
        )
        # ExEnd:ExportUnderlineFormatting

    def test_set_encoding(self):
        # ExStart:SetEncoding
        save_options = aw.saving.MarkdownSaveOptions()
        save_options.encoding = "utf-16"

        self.convert(
            "test_full_article.docx",
            "MarkdownSaveOptions.encoding_utf16.md",
            save_options=save_options,
        )
        # ExEnd:SetEncoding

        output = Path(ARTIFACTS_DIR) / "MarkdownSaveOptions.encoding_utf16.md"
        raw = output.read_bytes()
        assert raw[:2] == b"\xff\xfe", "UTF-16 LE BOM expected"
        assert "Introduction" in raw.decode("utf-16")

    def test_set_paragraph_break(self):
        # ExStart:SetParagraphBreak
        save_options = aw.saving.MarkdownSaveOptions()
        save_options.paragraph_break = "\r\n"

        self.convert(
            "test_full_article.docx",
            "MarkdownSaveOptions.paragraph_break_crlf.md",
            save_options=save_options,
        )
        # ExEnd:SetParagraphBreak

        output = Path(ARTIFACTS_DIR) / "MarkdownSaveOptions.paragraph_break_crlf.md"
        content = output.read_bytes()
        assert b"\r\n" in content, "CRLF line endings expected"

    def test_markdown_from_all_input_formats(self):
        """Save DOC, DOCX, RTF, TXT as Markdown with default options."""
        inputs = {
            "docx": "test_full_article.docx",
            "doc": "test_full_article.doc",
            "rtf": "test_full_article.rtf",
            "txt": "test_plain.txt",
        }
        save_options = aw.saving.MarkdownSaveOptions()
        for label, filename in inputs.items():
            self.convert(
                filename,
                f"MarkdownSaveOptions.from_{label}.md",
                save_options=save_options,
            )


if __name__ == "__main__":
    examples = WorkingWithMarkdownSaveOptions()

    print("=== Markdown Save Options Examples ===\n")
    examples.test_export_underline_formatting()
    print("  export_underline_formatting: Done.")
    examples.test_set_encoding()
    print("  set_encoding: Done.")
    examples.test_set_paragraph_break()
    print("  set_paragraph_break: Done.")
    examples.test_markdown_from_all_input_formats()
    print("  markdown_from_all_input_formats: Done.")
    print(f"\nOutput files are in: {ARTIFACTS_DIR}")
