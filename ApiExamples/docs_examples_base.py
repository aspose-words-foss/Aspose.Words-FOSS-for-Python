"""
Base module for API examples.

Provides constants, base class, and a shared conversion helper
mirroring the Aspose.Words DocsExamples structure.
Uses test data files from tests/data/input/ as input sources.
"""

import os
import sys
from pathlib import Path

# Root of the project
_PROJECT_ROOT = Path(__file__).resolve().parent.parent

# Ensure project root is on sys.path for standalone script execution
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

# Input data directory (reuse test fixtures)
MY_DIR = str(_PROJECT_ROOT / "tests" / "data" / "input") + os.sep

# Output directory for generated files
ARTIFACTS_DIR = str(_PROJECT_ROOT / "ApiExamples" / "output") + os.sep

# Ensure the output directory exists
os.makedirs(ARTIFACTS_DIR, exist_ok=True)

# Format extension -> (SaveFormat constant, human label)
OUTPUT_FORMATS = {
    ".md": ("markdown", "Markdown"),
    ".pdf": ("pdf", "PDF"),
    ".txt": ("text", "Text"),
}


class DocsExamplesBase:
    """Base class for API example tests, mirroring Aspose.Words DocsExamples."""

    @staticmethod
    def convert(input_file, output_file, save_format=None, save_options=None):
        """Load input_file and save to output_file.

        Args:
            input_file: Filename (relative to MY_DIR) or absolute path.
            output_file: Filename (relative to ARTIFACTS_DIR) or absolute path.
            save_format: A SaveFormat constant. Inferred from extension if None.
            save_options: A save options object (overrides save_format).
        """
        import aspose.words_foss as aw

        input_path = input_file if os.sep in input_file else MY_DIR + input_file
        output_path = output_file if os.sep in output_file else ARTIFACTS_DIR + output_file

        doc = aw.Document(input_path)
        doc.save(output_path, save_options or save_format)

    @staticmethod
    def convert_to_all_formats(input_file, output_prefix):
        """Convert input_file to Markdown, PDF, and Text."""
        import aspose.words_foss as aw

        fmt_map = {
            ".md": aw.SaveFormat.MARKDOWN,
            ".pdf": aw.SaveFormat.PDF,
            ".txt": aw.SaveFormat.TEXT,
        }
        doc = aw.Document(MY_DIR + input_file)
        for ext, fmt in fmt_map.items():
            doc.save(ARTIFACTS_DIR + output_prefix + ext, fmt)
