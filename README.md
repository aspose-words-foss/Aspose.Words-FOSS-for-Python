# Aspose.Words FOSS

A lightweight, open-source Python library for converting DOCX, DOC, RTF, TXT, and MD files to DOCX, Markdown, plain text, and PDF without requiring Microsoft Word.

A free, lightweight version of [Aspose.Words for Python via .NET](https://github.com/aspose-words/Aspose.Words-for-Python-via-.NET) with a compatible API (`Document`, `SaveFormat`, `SaveOptions`).

[![PyPI](https://img.shields.io/pypi/v/aspose-words-foss.svg)](https://pypi.org/project/aspose-words-foss/)
[![Python](https://img.shields.io/badge/python-3.10%20%7C%203.11%20%7C%203.12-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## Features

- **DOCX Read/Write**: Pure Python reader using only the standard library (`zipfile`, `xml.etree`)
- **DOC Support**: Word 97-2003 binary format reader via `olefile`
- **RTF Support**: Rich Text Format reader via OLE2 delegation
- **Markdown Import**: `.md` is parsed into the document model, not read as literal text — headings, bold/italic/strikethrough/inline code, ordered and nested lists, tables, block quotes, fenced code blocks, links, and base64-embedded images all become proper nodes, so Markdown converts to DOCX and PDF like any other input
- **Plain Text Input**: Read `.txt` files
- **File or Stream Input**: DOCX, DOC and RTF are auto-detected from magic bytes (anything else falls back to plain text); `LoadOptions.load_format` overrides the guess
- **Markdown Export**: Rich formatting — headings, bold/italic/strikethrough/underline, ordered and unordered lists (including nested), tables, block quotes, code blocks, and hyperlinks. Encoding and paragraph break sequence are configurable
- **PDF Export**: Generate PDF output via `fpdf2`. Applied `PdfSaveOptions` fields: `compliance`, `image_compression`, `jpeg_quality`, `outline_options`, `export_document_structure`, `export_bookmarks_outline`, `zoom_behavior`, `zoom_factor`, `display_doc_title`
- **Plain Text Export**: Extract document text content

## Installation

From [PyPI](https://pypi.org/project/aspose-words-foss/):

```bash
pip install aspose-words-foss
```

Nightly (latest from GitHub):

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

### Export to DOCX

```python
import aspose.words_foss as aw

doc = aw.Document("input.docx")  # or .doc, .rtf
doc.save("output.docx", aw.SaveFormat.DOCX)
```

### Extract plain text

```python
import aspose.words_foss as aw

doc = aw.Document("input.docx")
text = doc.get_text()
```

### Load from a stream

```python
import io
import aspose.words_foss as aw

with io.FileIO("input.docx") as stream:
    doc = aw.Document(stream)              # DOCX / DOC / RTF from magic bytes

opts = aw.LoadOptions()
opts.load_format = aw.LoadFormat.MARKDOWN  # needed for .md, which has no magic bytes
with io.FileIO("input.md") as stream:
    doc = aw.Document(stream, opts)
```

### Save with options

```python
import aspose.words_foss as aw
from aspose.words_foss.saving import (
    MarkdownSaveOptions,
    OoxmlSaveOptions,
    PdfSaveOptions,
    CompressionLevel,
)

doc = aw.Document("input.docx")

# Markdown: underline, encoding, paragraph break
md_opts = MarkdownSaveOptions()
md_opts.export_underline_formatting = True
md_opts.encoding = "utf-8-sig"        # write a UTF-8 BOM
md_opts.paragraph_break = "\r\n"      # CRLF between paragraphs
doc.save("output.md", md_opts)

# DOCX: compression level
ooxml_opts = OoxmlSaveOptions()
ooxml_opts.compression_level = CompressionLevel.MAXIMUM
doc.save("output.docx", ooxml_opts)

pdf_opts = PdfSaveOptions()
doc.save("output.pdf", pdf_opts)
```

## Requirements

- Python 3.10 or higher
- olefile >= 0.46
- fpdf2 >= 2.7.5
- pydantic >= 2.0.0

## API Examples

Runnable examples demonstrating the `aspose.words_foss` API live in the `ApiExamples/` folder.

The examples are written against the API both libraries share, so the same
sources run on `aspose-words` too — replacing `aspose.words_foss` with
`aspose.words` in the imports is the only edit needed.

### Files

| File | What it shows |
|------|---------------|
| `convert_document.py` | Every input format (DOCX, DOC, RTF, TXT, MD) to every output format (Markdown, PDF, TXT) |
| `loading_document.py` | Loading documents from a file path and from a binary stream, with `LoadOptions` |
| `loading_markdown.py` | Reading Markdown from in-memory content |
| `working_with_markdown_save_options.py` | `MarkdownSaveOptions` — `export_underline_formatting`, `encoding`, `paragraph_break` |
| `working_with_ooxml_save_options.py` | `OoxmlSaveOptions` for DOCX export — `pretty_format`, `compression_level`, `zip_64_mode` |
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