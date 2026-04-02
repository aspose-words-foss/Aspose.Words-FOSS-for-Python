# Aspose.Words FOSS

A lightweight, open-source Python library for converting DOCX, DOC, RTF, TXT, and MD files to Markdown, plain text, and PDF without requiring Microsoft Word.

A free, lightweight version of [Aspose.Words for Python via .NET](https://github.com/aspose-words/Aspose.Words-for-Python-via-.NET) with a compatible API (`Document`, `SaveFormat`, `SaveOptions`).

[![Python](https://img.shields.io/badge/python-3.10%20%7C%203.11%20%7C%203.12-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## Features

- **DOCX Support**: Pure Python reader using only the standard library (`zipfile`, `xml.etree`)
- **DOC Support**: Word 97-2003 binary format reader via `olefile`
- **RTF Support**: Rich Text Format reader via OLE2 delegation
- **Plain Text & Markdown Input**: Read `.txt` and `.md` files
- **Markdown Export**: Rich formatting — headings, bold/italic/strikethrough/underline, ordered and unordered lists (including nested), tables, block quotes, code blocks, and hyperlinks
- **PDF Export**: Generate PDF output via `fpdf2`
- **Plain Text Export**: Extract document text content

## Installation

```bash
pip install git+https://github.com/aspose-words-foss/Aspose.Words-FOSS-for-Python.git
```

## Quick Start

### Convert a document to Markdown

```python
import aspose.words_foss as aw

doc = aw.Document("input.docx")  # or .doc, .rtf, .txt, .md
doc.save("output.md", aw.SaveFormat.MARKDOWN)
```

### Export to PDF

```python
import aspose.words_foss as aw

doc = aw.Document("input.docx")
doc.save("output.pdf", aw.SaveFormat.PDF)
```

### Extract plain text

```python
import aspose.words_foss as aw

doc = aw.Document("input.docx")
text = doc.get_text()
```

### Save with options

```python
import aspose.words_foss as aw
from aspose.words_foss.saving import MarkdownSaveOptions, PdfSaveOptions

doc = aw.Document("input.docx")

md_opts = MarkdownSaveOptions()
md_opts.export_underline_formatting = True
doc.save("output.md", md_opts)

pdf_opts = PdfSaveOptions()
doc.save("output.pdf", pdf_opts)
```

## Requirements

- Python 3.10 or higher
- olefile >= 0.46
- fpdf2 >= 2.7.0
- pydantic >= 2.0.0

## API Examples

Runnable examples demonstrating the `aspose.words_foss` API:
ApiExamples folder

### Files

| File | What it shows |
|------|---------------|
| `convert_document.py` | Every input format (DOCX, DOC, RTF, TXT, MD) to every output format (Markdown, PDF, TXT) |
| `working_with_markdown_save_options.py` | `MarkdownSaveOptions` — only options that are actually applied |
| `working_with_pdf_save_options.py` | PDF export from all input formats |
| `working_with_txt_save_options.py` | Plain-text export and `get_text()` |
| `working_with_images.py` | Image-containing documents to all output formats |

### Running

```bash
# Individual scripts
python ApiExamples/convert_document.py

# All via pytest
python -m pytest ApiExamples/ -v --rootdir=ApiExamples -c ApiExamples/pytest.ini
```

### Input / Output

- **Input**: `tests/data/input/` (shared test fixtures)
- **Output**: `ApiExamples/output/` (git-ignored)

## License

This project is licensed under the MIT License - see the [LICENSE](License/license.txt) file for details.

## Support

- **Issues**: [GitHub Issues](https://github.com/aspose-words-foss/Aspose.Words-FOSS-for-Python/issues)