"""
API examples for PDF export.

PdfSaveOptions fields (compliance, image_compression, font_embedding_mode, etc.)
are accepted for API compatibility but are not yet applied during generation.

Mirrors the Aspose.Words for Python DocsExamples:
  file_formats_and_conversions/save_options/working_with_pdf_save_options.py

Run:
    python ApiExamples/working_with_pdf_save_options.py
    python -m pytest ApiExamples/ -v --rootdir=ApiExamples -c ApiExamples/pytest.ini
"""

import sys
from pathlib import Path

_HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(_HERE))
sys.path.insert(0, str(_HERE.parent))

import aspose.words_foss as aw  # noqa: E402
from docs_examples_base import DocsExamplesBase, ARTIFACTS_DIR  # noqa: E402


class WorkingWithPdfSaveOptions(DocsExamplesBase):

    def test_pdf_from_all_input_formats(self):
        """Save DOCX, DOC, RTF, TXT as PDF."""
        inputs = {
            "docx": "test_full_article.docx",
            "doc": "test_full_article.doc",
            "rtf": "test_full_article.rtf",
            "txt": "test_plain.txt",
        }
        for label, filename in inputs.items():
            self.convert(filename, f"PdfSaveOptions.from_{label}.pdf", save_format=aw.SaveFormat.PDF)

    def test_save_with_pdf_save_options(self):
        """PdfSaveOptions is accepted for API compatibility."""
        # ExStart:SaveWithPdfSaveOptions
        save_options = aw.saving.PdfSaveOptions()
        self.convert(
            "test_full_article.docx",
            "PdfSaveOptions.with_options.pdf",
            save_options=save_options,
        )
        # ExEnd:SaveWithPdfSaveOptions


if __name__ == "__main__":
    examples = WorkingWithPdfSaveOptions()

    print("=== PDF Save Options Examples ===\n")
    examples.test_pdf_from_all_input_formats()
    print("  pdf_from_all_input_formats: Done.")
    examples.test_save_with_pdf_save_options()
    print("  save_with_pdf_save_options: Done.")
    print(f"\nOutput files are in: {ARTIFACTS_DIR}")
