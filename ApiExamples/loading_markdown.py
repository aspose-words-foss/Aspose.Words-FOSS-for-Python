"""
API examples for reading Markdown documents from in-memory content.

Mirrors the Aspose.Words for Python via .NET loading examples:
https://github.com/aspose-words/Aspose.Words-for-Python-via-.NET

Run:
    python ApiExamples/loading_markdown.py
    python -m pytest ApiExamples/ -v --rootdir=ApiExamples -c ApiExamples/pytest.ini
"""

import base64
import io
import sys
from pathlib import Path

_HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(_HERE))
sys.path.append(str(_HERE.parent))  # append: must not shadow an installed wheel

import aspose.words_foss as aw  # noqa: E402
import aspose.words_foss.drawing as awd  # noqa: E402
from docs_examples_base import DocsExamplesBase, ARTIFACTS_DIR  # noqa: E402

MARKDOWN_TEXT = """# Release Notes

This build adds **bold**, *italic*, and `inline code` support, plus a [project link](https://example.com).

- Parses headings, emphasis, and lists
- Round-trips fenced code blocks unchanged

```python
def add(a, b):
    return a + b


def subtract(a, b):
    return a - b
```
"""

# A tiny (64x32) PNG, base64-encoded -- self-contained the same way an image
# pasted into a Markdown source as a "data:" URI would be, no external file needed.
LOGO_PNG_BASE64 = (
    "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAgCAIAAAAt/+nTAAAAV0lEQVR42u3XIQ7AIBAEQCB9GzVYPtTvYDH9"
    "Hb4pSR0lmbVnbnJnNuarh52TwuYBAAAAWJtjNrhr+duuZ+teCAAAAAAAAAAAAADgQ6F5bQ8uAAAAAPDIAKxD"
    "BmazG9D0AAAAAElFTkSuQmCC"
)

PRODUCT_URL = "https://products.aspose.com/words/"

MARKDOWN_WITH_IMAGE_AND_LINK = f"""# Aspose.Words for Python

Here is our logo, embedded straight from a base64 data URI so the
Markdown source is fully self-contained:

![company logo](data:image/png;base64,{LOGO_PNG_BASE64})

## Learn more

Visit the [product page]({PRODUCT_URL}) to see what Aspose.Words can do.

## Get started

Check out the [product page]({PRODUCT_URL}) again for pricing and downloads.
"""


class LoadingMarkdown(DocsExamplesBase):
    """Demonstrates loading Markdown straight from a string/bytes, no file needed."""

    def test_load_markdown_from_bytes(self):
        # ExStart:LoadMarkdownFromBytes
        doc = aw.Document(io.BytesIO(MARKDOWN_TEXT.encode("utf-8")), aw.loading.MarkdownLoadOptions())
        # ExEnd:LoadMarkdownFromBytes
        assert "Release Notes" in doc.get_text()

    def test_load_markdown_from_stream(self):
        # ExStart:LoadMarkdownFromStream
        with io.BytesIO(MARKDOWN_TEXT.encode("utf-8")) as stream:
            doc = aw.Document(stream, aw.loading.MarkdownLoadOptions())
        # ExEnd:LoadMarkdownFromStream
        assert "def add" in doc.get_text()

    def test_markdown_round_trip_preserves_source(self):
        # ExStart:MarkdownRoundTrip
        doc = aw.Document(io.BytesIO(MARKDOWN_TEXT.encode("utf-8")), aw.loading.MarkdownLoadOptions())
        doc.save(ARTIFACTS_DIR + "LoadingMarkdown.RoundTrip.md", aw.SaveFormat.MARKDOWN)
        # ExEnd:MarkdownRoundTrip

        result = Path(ARTIFACTS_DIR + "LoadingMarkdown.RoundTrip.md").read_text(
            encoding="utf-8-sig"
        ).replace("\r\n", "\n")

        assert "# Release Notes" in result
        assert "**bold**" in result
        assert "*italic*" in result
        assert "`inline code`" in result
        assert "[project link](https://example.com)" in result
        assert "- Parses headings, emphasis, and lists" in result
        assert "- Round-trips fenced code blocks unchanged" in result
        assert "def add(a, b):\n    return a + b" in result
        assert "def subtract(a, b):\n    return a - b" in result

    def test_preserve_empty_lines(self):
        # ExStart:PreserveEmptyLines
        source = "# Title\n\n\nParagraph after two blank lines.\n"

        load_options = aw.loading.MarkdownLoadOptions()
        load_options.preserve_empty_lines = True

        doc = aw.Document(io.BytesIO(source.encode("utf-8")), load_options)
        # ExEnd:PreserveEmptyLines

        paragraph_texts = [p.get_text().strip() for p in doc.sections[0].body.paragraphs]
        assert paragraph_texts[-4:] == ["Title", "", "", "Paragraph after two blank lines."]

    def test_save_markdown_with_base64_image_to_docx_and_pdf(self):
        """A Markdown source with several headings, a base64 image, and repeated links,
        saved to DOCX and PDF with the image embedded and the links clickable."""
        # ExStart:SaveMarkdownWithImageToDocxAndPdf
        doc = aw.Document(
            io.BytesIO(MARKDOWN_WITH_IMAGE_AND_LINK.encode("utf-8")),
            aw.loading.MarkdownLoadOptions(),
        )
        docx_path = ARTIFACTS_DIR + "LoadingMarkdown.WithImage.docx"
        pdf_path = ARTIFACTS_DIR + "LoadingMarkdown.WithImage.pdf"
        doc.save(docx_path, aw.SaveFormat.DOCX)
        doc.save(pdf_path, aw.SaveFormat.PDF)
        # ExEnd:SaveMarkdownWithImageToDocxAndPdf

        # The base64 image becomes a real embedded Shape, not literal "![alt](data:...)" text.
        expected_png = base64.b64decode(LOGO_PNG_BASE64)
        shapes = [n.as_shape() for n in doc.get_child_nodes(aw.NodeType.SHAPE, True)]
        ours = [sh for sh in shapes if sh.image_data.image_bytes == expected_png]
        assert len(ours) == 1
        assert ours[0].has_image
        assert ours[0].image_data.image_type == awd.ImageType.PNG

        # Re-read the saved .docx from disk to confirm the image survived the round trip,
        # not just the in-memory Document.
        reloaded = aw.Document(docx_path)
        saved_shapes = [
            n.as_shape()
            for n in reloaded.get_child_nodes(aw.NodeType.SHAPE, True)
            if n.as_shape().has_image
        ]
        assert [sh for sh in saved_shapes if sh.image_data.image_bytes == expected_png]

        # Both headings ("Learn more" / "Get started") kept their own heading style,
        # and both links became real, clickable "<w:hyperlink>" elements pointing at
        # products.aspose.com -- not literal "[text](url)" text.
        import zipfile

        with zipfile.ZipFile(docx_path) as zf:
            document_xml = zf.read("word/document.xml").decode("utf-8")
            rels_xml = zf.read("word/_rels/document.xml.rels").decode("utf-8")
        assert document_xml.count("<w:hyperlink ") >= 2
        assert f"[product page]({PRODUCT_URL})" not in document_xml  # never literal markdown syntax
        assert rels_xml.count(f'Target="{PRODUCT_URL}"') >= 1

        # Same check on the PDF: two real clickable link annotations, not styled text.
        import pdfplumber

        with pdfplumber.open(pdf_path) as pdf:
            links = [link for page in pdf.pages for link in page.hyperlinks]
        assert len([link for link in links if link["uri"] == PRODUCT_URL]) == 2


if __name__ == "__main__":
    examples = LoadingMarkdown()

    print("=== Loading Markdown Examples ===\n")

    print("  Load from bytes...")
    examples.test_load_markdown_from_bytes()
    print("  Load from stream...")
    examples.test_load_markdown_from_stream()
    print("  Round-trip preserves source...")
    examples.test_markdown_round_trip_preserves_source()
    print("  Preserve empty lines...")
    examples.test_preserve_empty_lines()
    print("  Save Markdown with base64 image to DOCX and PDF...")
    examples.test_save_markdown_with_base64_image_to_docx_and_pdf()

    print("\n  All examples passed.")
    print(f"  Output files are in: {ARTIFACTS_DIR}")
