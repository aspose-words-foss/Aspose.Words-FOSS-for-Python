"""
API examples for converting documents that contain images.

Uses DOCX test files with embedded images. Converts each to all three
output formats (Markdown, PDF, Text).

Run:
    python ApiExamples/working_with_images.py
    python -m pytest ApiExamples/ -v --rootdir=ApiExamples -c ApiExamples/pytest.ini
"""

import sys
from pathlib import Path

_HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(_HERE))
sys.path.insert(0, str(_HERE.parent))

from docs_examples_base import DocsExamplesBase, ARTIFACTS_DIR  # noqa: E402

# Image-containing test files
IMAGE_FILES = [
    "simple_inline.docx",
    "image_with_caption.docx",
    "image_in_table.docx",
    "consecutive_images.docx",
    "multi_image_ordered.docx",
    "image_in_header.docx",
    "image_in_footer.docx",
    "wide_image.docx",
]


class WorkingWithImages(DocsExamplesBase):
    """Convert image-containing documents to all output formats."""

    def test_image_documents_to_all_formats(self):
        for filename in IMAGE_FILES:
            stem = Path(filename).stem
            self.convert_to_all_formats(filename, f"Images.{stem}")


if __name__ == "__main__":
    examples = WorkingWithImages()

    print("=== Image Document Examples ===\n")
    for f in IMAGE_FILES:
        print(f"  {f} -> md, pdf, txt")
    examples.test_image_documents_to_all_formats()
    print("  Done.")
    print(f"\nOutput files are in: {ARTIFACTS_DIR}")
