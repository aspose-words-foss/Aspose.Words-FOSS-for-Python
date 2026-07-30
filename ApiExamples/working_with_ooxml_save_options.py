"""
API examples for OoxmlSaveOptions.

Demonstrates options applied during DOCX export:
pretty_format, compression_level, zip_64_mode.

Mirrors the Aspose.Words for Python DocsExamples:
  file_formats_and_conversions/save_options/working_with_ooxml_save_options.py

Run:
    python ApiExamples/working_with_ooxml_save_options.py
    python -m pytest ApiExamples/ -v --rootdir=ApiExamples -c ApiExamples/pytest.ini
"""

import sys
import zipfile
from pathlib import Path

_HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(_HERE))
sys.path.append(str(_HERE.parent))  # append: must not shadow an installed wheel

import aspose.words_foss as aw  # noqa: E402
from docs_examples_base import ARTIFACTS_DIR, DocsExamplesBase  # noqa: E402


class WorkingWithOoxmlSaveOptions(DocsExamplesBase):

    def test_set_pretty_format(self):
        # ExStart:SetPrettyFormat
        save_options = aw.saving.OoxmlSaveOptions()
        save_options.pretty_format = True

        self.convert(
            "test_full_article.docx",
            "OoxmlSaveOptions.pretty_format.docx",
            save_options=save_options,
        )
        # ExEnd:SetPrettyFormat

        output = Path(ARTIFACTS_DIR) / "OoxmlSaveOptions.pretty_format.docx"
        with zipfile.ZipFile(output) as zf:
            doc_xml = zf.read("word/document.xml").decode("utf-8")
        assert "\r\n\t" in doc_xml, "XML should be indented"
        assert len(doc_xml.splitlines()) > 1, "Pretty mode: XML should span many lines"

    def test_pretty_format_off(self):
        # ExStart:PrettyFormatOff
        save_options = aw.saving.OoxmlSaveOptions()
        save_options.pretty_format = False

        self.convert(
            "test_full_article.docx",
            "OoxmlSaveOptions.compact.docx",
            save_options=save_options,
        )
        # ExEnd:PrettyFormatOff

        output = Path(ARTIFACTS_DIR) / "OoxmlSaveOptions.compact.docx"
        with zipfile.ZipFile(output) as zf:
            doc_xml = zf.read("word/document.xml").decode("utf-8")
        assert len(doc_xml.splitlines()) == 1, "Compact mode: whole part on one line"

    def test_pretty_format_round_trip(self):
        """Pretty-formatted DOCX should produce the same text after reload."""
        # ExStart:PrettyFormatRoundTrip
        save_options = aw.saving.OoxmlSaveOptions()
        save_options.pretty_format = True

        self.convert(
            "test_full_article.docx",
            "OoxmlSaveOptions.pretty_roundtrip.docx",
            save_options=save_options,
        )

        original = aw.Document(
            str(Path(ARTIFACTS_DIR) / ".." / ".." / "tests" / "data" / "input" / "test_full_article.docx")
        )
        reloaded = aw.Document(str(Path(ARTIFACTS_DIR) / "OoxmlSaveOptions.pretty_roundtrip.docx"))
        # ExEnd:PrettyFormatRoundTrip

        assert original.get_text().strip() == reloaded.get_text().strip()

    def test_set_compression_level(self):
        # ExStart:SetCompressionLevel
        save_options = aw.saving.OoxmlSaveOptions()
        save_options.compression_level = aw.saving.CompressionLevel.MAXIMUM

        self.convert(
            "test_full_article.docx",
            "OoxmlSaveOptions.max_compression.docx",
            save_options=save_options,
        )
        # ExEnd:SetCompressionLevel


if __name__ == "__main__":
    examples = WorkingWithOoxmlSaveOptions()

    print("=== OOXML Save Options Examples ===\n")
    examples.test_set_pretty_format()
    print("  set_pretty_format: Done.")
    examples.test_pretty_format_off()
    print("  pretty_format_off: Done.")
    examples.test_pretty_format_round_trip()
    print("  pretty_format_round_trip: Done.")
    examples.test_set_compression_level()
    print("  set_compression_level: Done.")
    print(f"\nOutput files are in: {ARTIFACTS_DIR}")
