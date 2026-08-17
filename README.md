# Aspose.Words FOSS for Python

[![PyPI version](https://img.shields.io/pypi/v/aspose-words-foss.svg)](https://pypi.org/project/aspose-words-foss/) [![Python versions](https://img.shields.io/pypi/pyversions/aspose-words-foss.svg)](https://pypi.org/project/aspose-words-foss/) [![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE) [![Contributors](https://img.shields.io/github/contributors/aspose-words-foss/Aspose.Words-FOSS-for-Python.svg)](https://github.com/aspose-words-foss/Aspose.Words-FOSS-for-Python/graphs/contributors)

[![Aspose.Words FOSS for Python](https://products.aspose.org/media/words/python/banner-readme.png)](https://products.aspose.org/words/python/)

Aspose.Words FOSS for Python is a free, open-source, MIT-licensed Python library for working with
Word documents. It reads DOCX, DOC (Word 97-2003), RTF, Markdown, and plain-text files into a
shared Light Document Model, and exports that model to DOCX, Markdown, PDF, or plain text — all
without Microsoft Word or any COM/Office automation. It follows a compatible API shape (`Document`,
`SaveFormat`) to the commercial Aspose.Words for Python.

## Navigation

- [At a Glance](#at-a-glance)
- [Key Capabilities](#key-capabilities)
- [Installation](#installation)
- [Dependencies](#dependencies)
- [Quick Start](#quick-start)
- [Additional Examples](#additional-examples)
- [API Reference](#api-reference)
- [Documentation & Resources](#documentation--resources)
- [Scope and Limitations](#scope-and-limitations)
- [Development and Testing](#development-and-testing)
- [License](#license)

## At a Glance

```mermaid
flowchart TD
  subgraph StartingPoints["Starting Points"]
    direction TB
    i1["An existing DOCX document"]
    i2["An existing DOC (Word 97-2003) document"]
    i3["An existing RTF document"]
    i4["An existing Markdown file"]
    i5["An existing plain-text file"]
  end
  PRODUCT["Aspose.Words FOSS for Python"]
  subgraph Capabilities["Core Capabilities"]
    direction TB
    c1["Load and parse DOC, DOCX, RTF, Markdown, and plain-text documents"]
    c2["DOCX round-trip editing via the Light Document Model"]
    c3["Markdown export with configurable headings, lists, and tables"]
    c4["PDF export via a built-in paragraph, table, and shape renderer"]
  end
  subgraph Outputs["Outputs"]
    direction TB
    o1["DOCX documents"]
    o2["Markdown files"]
    o3["PDF documents"]
    o4["Plain text"]
  end
  StartingPoints --> PRODUCT --> Capabilities --> Outputs
```

## Key Capabilities

- Open any supported input with a single `Document(filepath)` constructor — the concrete reader is
  chosen automatically from the `.docx`/`.doc`/`.rtf`/`.md`/`.txt` file extension.
- Read DOCX with a pure-Python parser built on the standard library `zipfile`/`xml.etree` modules —
  no compiled dependencies.
- Read legacy Word 97-2003 `.doc` binary files via `olefile`, and RTF documents through the same
  OLE2 delegation path.
- Parse Markdown on import with `MarkdownReader`, not just read it as literal text — headings, lists,
  emphasis, tables, block quotes, fenced code blocks, links, and base64-embedded images all become
  proper document-model nodes, so a loaded `.md` file converts to DOCX or PDF like any other input.
- Round-trip DOCX editing: `DocumentReader` builds a shared Light Document Model from an existing
  `.docx`, and `LdmDocxWriter` writes it back out, preserving headers/footers, inline shapes and
  images, bookmarks, and custom paragraph styles.
- Convert image-containing documents — inline images, captioned images, images in tables, and
  images in headers/footers — to every output format: images embed as base64 data URIs in
  Markdown by default, render through the built-in `ShapeRenderer` in PDF, and round-trip
  losslessly through DOCX.
- Export to Markdown with `LdmMarkdownWriter`: headings, bold/italic/underline/strikethrough runs,
  ordered and nested unordered lists, tables, block quotes, code blocks, and hyperlinks, all tunable
  through `MarkdownSaveOptions`.
- Export to PDF with `LdmPdfWriter` — a built-in paragraph, table, and shape renderer (backed by
  `fpdf2`) that needs no external PDF engine or system fonts, configurable through `PdfSaveOptions`.
- Extract plain text from any loaded document with `Document.get_text()`, or save it directly with
  `SaveFormat.TEXT`.
- Configure DOCX packaging through `OoxmlSaveOptions`: compression level, ECMA-376/ISO 29500
  Transitional compliance, and pretty-printed XML.
- Inspect the parsed document model directly — sections, paragraphs, runs, tables, styles, and
  numbered/bulleted lists are all typed Pydantic models reachable from `Document.light_document_model`.

## Installation

Install from PyPI:

```bash
pip install aspose-words-foss>=26.5.0
```

Or install the latest nightly build directly from GitHub:

```bash
pip install git+https://github.com/aspose-words-foss/Aspose.Words-FOSS-for-Python.git
```

Requires Python 3.10 or later; runtime dependencies install automatically via pip — see
[Dependencies](#dependencies) below for the full required and native-requirement breakdown.

Verify the install:

```bash
python -c "import aspose_words_foss; print('aspose_words_foss OK')"
```

## Dependencies

### Required Package Dependencies

- `olefile` >=0.46 — reads legacy Word 97-2003 .doc binary files and RTF documents via OLE2 delegation.
- `fpdf2` >=2.7.5 — backs the built-in PDF renderer.
- `pydantic` >=2.0.0 — the typed model layer for the parsed document model.

### Native and System Requirements

- Requires Python 3.10 or later (tested through 3.12; `pyproject.toml` caps at `<3.13`).

### Development Dependencies

- `Pillow>=10.0.0` — used by the `dev` extra's example/test tooling for image-containing documents.
- `pytest>=9.0.2` — the test runner used to execute `ApiExamples/`.

None of these are required to install or use the published package; they apply only when
installing with `pip install -e ".[dev]"` for local development and testing.

## Quick Start

Load a document and export it to Markdown:

```python
import aspose.words_foss as aw

doc = aw.Document("report.docx")  # or .doc, .rtf, .txt, .md
doc.save("report.md", aw.SaveFormat.MARKDOWN)
```

Export the same document to PDF:

```python
import aspose.words_foss as aw

doc = aw.Document("report.docx")
doc.save("report.pdf", aw.SaveFormat.PDF)
```

More runnable scripts — covering every input/output format combination, `MarkdownSaveOptions`, PDF
export, plain-text export, and image-containing documents — are collected in
[`ApiExamples/`](ApiExamples/); see [Additional Examples](#additional-examples) below.

## Additional Examples

More real, runnable examples are collected below, matching the standalone scripts under
[`ApiExamples/`](ApiExamples/).

### Save With Markdown and PDF Options

```python
import aspose.words_foss as aw

doc = aw.Document("report.docx")

md_opts = aw.saving.MarkdownSaveOptions()
md_opts.export_underline_formatting = True
doc.save("report.md", md_opts)

pdf_opts = aw.saving.PdfSaveOptions()
doc.save("report.pdf", pdf_opts)
```

<details>
<summary>View Additional Examples</summary>

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

### Extract Plain Text

```python
import aspose.words_foss as aw

doc = aw.Document("report.docx")
text = doc.get_text()
print(text)
```

### Convert Every Input Format to PDF

```python
import aspose.words_foss as aw

for path in ("report.docx", "legacy.doc", "notes.rtf"):
    doc = aw.Document(path)
    doc.save(f"{path}.pdf", aw.SaveFormat.PDF)
```

### Load From a Stream

DOCX, DOC, and RTF are auto-detected from magic bytes; formats with no magic bytes (like Markdown)
need `LoadOptions.load_format` to disambiguate:

```python
import io
import aspose.words_foss as aw

with io.FileIO("report.docx") as stream:
    doc = aw.Document(stream)  # DOCX / DOC / RTF from magic bytes

opts = aw.LoadOptions()
opts.load_format = aw.LoadFormat.MARKDOWN  # needed for .md, which has no magic bytes
with io.FileIO("notes.md") as stream:
    doc = aw.Document(stream, opts)
```

### Preserve a UTF-8 BOM in Markdown Output

```python
import aspose.words_foss as aw

doc = aw.Document("report.docx")
md_opts = aw.saving.MarkdownSaveOptions()
md_opts.encoding = "utf-8-sig"
doc.save("report.md", md_opts)
```

### Pretty-Print DOCX XML

```python
import aspose.words_foss as aw

doc = aw.Document("report.docx")
ooxml_opts = aw.saving.OoxmlSaveOptions()
ooxml_opts.pretty_format = True
doc.save("report-pretty.docx", ooxml_opts)
```

### Set DOCX Compression Level

```python
import aspose.words_foss as aw

doc = aw.Document("report.docx")
ooxml_opts = aw.saving.OoxmlSaveOptions()
ooxml_opts.compression_level = aw.saving.CompressionLevel.MAXIMUM
doc.save("report-max-compression.docx", ooxml_opts)
```

### Convert Image-Containing Documents to Every Output Format

Inline images, captioned images, images inside tables, and images in headers/footers all convert
cleanly across every output format:

```python
import aspose.words_foss as aw

image_containing_files = [
    "simple_inline.docx",
    "image_with_caption.docx",
    "image_in_table.docx",
    "image_in_header.docx",
    "image_in_footer.docx",
]
for filename in image_containing_files:
    doc = aw.Document(filename)
    doc.save(f"{filename}.md", aw.SaveFormat.MARKDOWN)   # images embed as base64 by default
    doc.save(f"{filename}.pdf", aw.SaveFormat.PDF)        # rendered via ShapeRenderer
    doc.save(f"{filename}.txt", aw.SaveFormat.TEXT)
```

### Load Markdown From In-Memory Content

`MarkdownLoadOptions` (not the generic `LoadOptions`) loads Markdown directly from bytes or a
stream — no file path required — and adds a `preserve_empty_lines` option that keeps blank lines
which would otherwise collapse during parsing:

```python
import io
import aspose.words_foss as aw

markdown_bytes = b"# Release Notes\n\nThis build adds **bold** and *italic* support.\n"

load_opts = aw.MarkdownLoadOptions()
doc = aw.Document(io.BytesIO(markdown_bytes), load_opts)
print(doc.get_text())

load_opts.preserve_empty_lines = True
doc = aw.Document(
    io.BytesIO(b"# Title\n\n\nParagraph after two blank lines.\n"), load_opts
)
```

A Markdown source with a base64-embedded image and links round-trips to real content, not
literal Markdown syntax — the image becomes a genuine embedded shape and each link becomes a real
clickable hyperlink, in both DOCX and PDF output:

```python
import io
import aspose.words_foss as aw

markdown_with_image = (
    "# Report\n\n"
    "![logo](data:image/png;base64,iVBORw0KGgo...)\n\n"
    "See the [documentation](https://example.com/docs) for details.\n"
)

doc = aw.Document(io.BytesIO(markdown_with_image.encode("utf-8")), aw.MarkdownLoadOptions())
doc.save("report.docx", aw.SaveFormat.DOCX)  # image becomes a real embedded Shape
doc.save("report.pdf", aw.SaveFormat.PDF)    # link becomes a real clickable PDF annotation
```

</details>

## API Reference

`Document` is the primary entry point: construct it from a file path, then call `save()` with a
`SaveFormat` constant or a `MarkdownSaveOptions` / `PdfSaveOptions` / `OoxmlSaveOptions` instance —
the target format is otherwise inferred from the output file's extension. The library ships 146
public types across the DOC/DOCX/RTF/Markdown/Text readers, the DOCX/Markdown/PDF writers, and the
shared Light Document Model, summarized in the module-grouped table below.

<details>
<summary>View the Supported Public API Surface</summary>

### Core API

| Class | Description |
|---|---|
| `Body` | Class with 2 properties. |
| `BookmarkEnd` | Class with 2 properties. |
| `BookmarkStart` | Marks the beginning of a Word bookmark (`<w:bookmarkStart>`). |
| `Border` | Class with 3 properties. |
| `Cell` | Class with 4 properties. |
| `CellFormat` | Class with 14 properties. |
| `ColorMode` | Color rendering mode. |
| `CompressionLevel` | Compression level for OOXML files. |
| `ConversionOptions` | Options for controlling DOCX to Markdown conversion. |
| `DocList` | Class with 4 properties. |
| `Document` | Represents a Word document. |
| `Document-light_document_model` | Class with 2 methods and 13 properties. |
| `DocumentFormatReader` | Protocol defining the interface all document readers must implement. |
| `FieldEnd` | Class with 3 properties. |
| `FieldSeparator` | Class with 2 properties. |
| `FieldStart` | Class with 2 properties. |
| `Font` | Class with 23 properties. |
| `FrameFormat` | Floating text-frame definition. |
| `HeaderFooter` | Class with 5 properties. |
| `ImageData` | Class with 7 properties. |
| `LdmMarkdownWriter` | Converts a `light_document_model.Document` to a Markdown string. |
| `ListFormat` | Class with 3 properties. |
| `ListLabel` | Snapshot of a list-item's rendered bullet/number label. |
| `ListLevel` | Class with 9 properties. |
| `ListLevelOverride` | One `<w:lvlOverride>` inside a concrete `<w:num>`. |
| `LoadFormat` | Document load format constants. |
| `MarkdownEmptyParagraphExportMode` | Controls how empty paragraphs are exported. |
| `MarkdownExportAsHtml` | Controls which elements are exported as raw HTML. |
| `MarkdownFileReader` | Reads Markdown (.md) files and yields one Paragraph per line. |
| `MarkdownLinkExportMode` | Link export mode options. |
| `MarkdownListExportMode` | List export mode options. |
| `MarkdownSaveOptions` | Options for saving documents as Markdown. |
| `OoxmlCompliance` | OOXML standards compliance level. |
| `OoxmlSaveOptions` | Options for saving a document as DOCX (Office Open XML). |
| `OutlineOptions` | Controls how outlines (bookmarks panel) are generated in the PDF. |
| `PageSetup` | Class with 18 properties. |
| `Paragraph` | A paragraph whose children — `Run`, `BookmarkStart` / `End`, `FieldStart` / `Separator` / `End` and inline `ShapeNode` — sit in a single ordered collection in document order. |
| `ParagraphFormat` | Class with 35 properties. |
| `ParagraphInfo` | Information about a paragraph's style and context. |
| `PdfCompliance` | PDF standards compliance level. |
| `PdfFontEmbeddingMode` | Font embedding mode in PDF. |
| `PdfImageCompression` | Image compression in PDF. |
| `PdfPageMode` | PDF page display mode. |
| `PdfSaveOptions` | Options for saving documents as PDF. |
| `PdfTextCompression` | Text compression in PDF. |
| `PdfZoomBehavior` | Mirrors public API. |
| `Row` | Class with 3 properties. |
| `RowFormat` | Class with 7 properties. |
| `RtfFileReader` | Reads RTF files (OLE2-format) and produces the same data structures as DocumentReader (for .docx) and DocFileReader (for .doc). |
| `Run` | Class with 3 properties. |
| `RunFormatting` | Text run formatting properties. |
| `SaveFormat` | Document save format constants. |
| `Section` | Class with 4 properties. |
| `Shading` | Class with 2 properties. |
| `ShapeNode` | Class with 25 properties. |
| `Style` | Class with 11 properties. |
| `TabStop` | Class with 4 properties. |
| `TabStopCollection` | Class with 5 methods and 1 property. |
| `Table` | Class with 11 properties. |
| `Table-models` | Represents a table structure. |
| `TableCell` | Represents a table cell. |
| `TableContentAlignment` | Table content alignment options. |
| `TableRow` | Represents a table row. |
| `TableStyleFormat` | Table-level properties stored on table styles (`w:tblPr` inside `w:style`). |
| `TextColumn` | Class with 2 properties. |
| `TextColumns` | Class with 5 properties. |
| `TextFileReader` | Reads plain-text (.txt) files and yields one Paragraph per line. |
| `UnknownNode` | Class with 1 property. |
| `Zip64Mode` | Controls when to use ZIP64 format extensions for OOXML files. |

#### Enumerations

| Enumeration | Description |
|---|---|
| `CodeBlockStyle` | Code block style preference. |
| `HeadingStyle` | Heading export style preference. |
| `ListMarker` | Bullet list marker style. |

### Converters

| Class | Description |
|---|---|
| `ListHandler` | Handles parsing and conversion of lists. |
| `ParagraphConverter` | Handles conversion of paragraphs to Markdown. |
| `TableConverter` | Handles conversion of tables to Markdown. |

### DOC Reader

| Class | Description |
|---|---|
| `BlipInfo` | Parsed BSE (Blip Store Entry) with location of image data. |
| `CharProps` | Properties extracted from CHPX for a text run. |
| `ChildAnchorInfo` | Position of a child shape within its parent group's coordinate system. |
| `DocFileReader` | Full DOC reader with LDM (Light Document Model) building capability. |
| `DocFileReaderCore` | Core reader for Word 97-2003 (.doc) files. |
| `DocTableBuilderMixin` | Mixin that adds table-building helpers to the DOC reader. |
| `FibData` | Parsed FIB (File Information Block) data. |
| `GroupShapeInfo` | Coordinate system of a shape group (from Spgr record). |
| `ListDef` | Parsed list definition with full level information. |
| `ListLevelData` | One parsed LVL record (measurements in points). |
| `ParaProps` | Properties extracted from PAPX for a single paragraph. |
| `ShapeAnchor` | Parsed SPA (Shape Address) from PlcSpaMom / PlcSpaHdr. |
| `ShapeLineProps` | Escher line/fill properties for a shape (from FOpt records). |
| `StyleData` | Properties parsed from a style definition (STSH UPX). |
| `TableRowProps` | Properties extracted from PAPX SPRMs on table row-end paragraphs. |

### DOCX Reader

| Class | Description |
|---|---|
| `CellBuilder` | Build `ldm.Cell` from a `<w:tc>` element. |
| `CellData` | Table cell. |
| `DocumentReader` | Reads DOCX documents and produces abstracted data structures. |
| `FontBuilder` | Build a non-cascaded `ldm.Font` from one `<w:rPr>`. |
| `FontResolver` | Compose a fully cascaded `ldm.Font`. |
| `LdmBuilderMixin` | Mixin that adds `to_light_document` to `DocumentReader`. |
| `ListBuilder` | Translate `<w:numbering>` into a list of `ldm.DocList`. |
| `NumberingInfo` | Numbering definition. |
| `NumberingLevel` | List level definition. |
| `PageSetupBuilder` | Build `ldm.PageSetup` from a `<w:sectPr>` element. |
| `ParagraphBuilder` | Build `ldm.Paragraph` from a `<w:p>` element. |
| `ParagraphData` | Paragraph with style and content. |
| `ParagraphFormatBuilder` | Build a non-cascaded `ldm.ParagraphFormat` from one `<w:pPr>`. |
| `ParagraphFormatResolver` | Compose a fully cascaded `ldm.ParagraphFormat`. |
| `ReaderContext` | Class extending Protocol. |
| `RowBuilder` | Build `ldm.Row` from a `<w:tr>` element. |
| `RowData` | Table row. |
| `RunBuilder` | Build a single `ldm.Run` with a resolved font cascade. |
| `RunData` | Text run with formatting. |
| `SectionBuilder` | Split the body into `ldm.Section` objects at sectPr boundaries. |
| `ShapeParserMixin` | Mixin providing drawing/shape parsing methods for DocumentReader. |
| `StyleBuilder` | Translate `<w:styles>` into a list of `ldm.Style`. |
| `StyleChainResolver` | Walk the `<w:basedOn>` graph for a styleId. |
| `TableBuilder` | Build `ldm.Table` from a `<w:tbl>` element. |
| `TableData` | Table structure. |

### DOCX Writer

| Class | Description |
|---|---|
| `BookmarkState` | Hands out monotonically increasing bookmark ids and pairs starts/ends across paragraphs. |
| `DocxWriterLossyWarning` | Warns the caller that the writer is dropping known LDM constructs. |
| `ImageEntry` | One image to add to `word/media/` plus its relationship row. |
| `ImageRenderState` | Accumulator threaded through paragraph rendering for inline shapes. |
| `LdmDocxWriter` | Convert an `ldm.Document` into a DOCX file. |

### Model

| Class | Description |
|---|---|
| `CellMerge` | Specifies how a cell in a table is merged with other cells. |
| `CellVerticalAlignment` | Specifies vertical justification of text inside a table cell. |
| `HeightRule` | Specifies the rule for determining the height of an object. |
| `LineSpacingRule` | Specifies values for line spacing. |
| `LineStyle` | Specifies line style of a border. |
| `NumberStyle` | Specifies the number style for a list, footnotes, endnotes, page numbers. |
| `Orientation` | Specifies page orientation. |
| `ParagraphAlignment` | Specifies text alignment in a paragraph. |
| `SectionStart` | Specifies the type of break at the beginning of the section. |
| `StyleIdentifier` | Locale-independent built-in style identifier. |
| `StyleType` | Specifies type of the style. |
| `TabAlignment` | Tab stop alignment. |
| `TabLeader` | Tab stop leader character. |
| `Underline` | Specifies type of the underline applied to a font. |
| `WrapType` | Specifies how text is wrapped around a shape or picture. |

### Parsers

| Class | Description |
|---|---|
| `ListInfo` | Information about a list. |
| `ListLevelInfo` | Information about a list level. |
| `NumberingParser` | Parser for DOCX numbering definitions. |
| `ParsedStyle` | Parsed style information. |
| `StyleParser` | Parser for DOCX style names and properties. |

### PDF Writer

| Class | Description |
|---|---|
| `LdmPdfWriter` | Converts a `light_document_model.Document` to a PDF file. |
| `PDFWriterContext` | Protocol requiring a `PdfSaveOptions` options property. |
| `ParagraphRenderer` | Renders LDM paragraphs into PDF. |
| `RunRenderer` | Renders formatted runs (text segments with fonts, colors, links). |
| `ShapeRenderer` | Renders shapes, images, and positioned elements. |
| `TableRenderer` | Renders LDM tables into PDF. |

---

#### Detailed Member Reference

### Document Loading and Saving

- `Document(filepath=None, *, stream=None, data=None)` — loads a `.doc`, `.docx`, `.rtf`, `.txt`, or
  `.md` file (or DOCX bytes/stream) and populates the Light Document Model immediately at
  construction time.
  - `light_document_model: ldm.Document` — the underlying parsed model
  - `sections`, `first_section`, `last_section`, `styles`, `lists`, `page_count`
  - `get_text() -> str` — plain-text extraction
  - `save(output_path, save_format_or_options=None)` — accepts a `SaveFormat` constant or a
    `MarkdownSaveOptions` / `PdfSaveOptions` / `OoxmlSaveOptions` instance; otherwise the target
    format is inferred from the output file's extension
- `SaveFormat`: `MARKDOWN`, `DOC` (reserved for API compatibility, not implemented — see
  [Scope and Limitations](#scope-and-limitations)), `DOCX`, `TEXT`, `PDF`
- `LoadFormat`: `AUTO`, `DOC`, `DOCX`, `RTF`, `TEXT`, `MARKDOWN`

### Format Readers

- `DocumentReader` (DOCX) — combines `LdmBuilderMixin` and `ShapeParserMixin`; `load_file()` /
  `load_stream()` / `load_bytes()`, then `to_light_document()`
- `DocFileReader` (DOC, Word 97-2003 binary) — combines `DocTableBuilderMixin` and
  `DocFileReaderCore`, backed by `olefile`
- `RtfFileReader` — OLE2-delegated RTF reader with the same `to_light_document()` contract
- `MarkdownFileReader`, `TextFileReader` — yield one `Paragraph` per line
- `DocumentFormatReader` — the `Protocol` all four readers implement

### Markdown Export

- `LdmMarkdownWriter.write(doc, output_path) -> str`
- `MarkdownSaveOptions` — `heading_style` (`HeadingStyle.ATX` / `SETEXT`), `list_marker`
  (`ListMarker.DASH` / `ASTERISK` / `PLUS`), `code_block_style`, `table_content_alignment`,
  `link_export_mode`, `export_as_html`, `export_underline_formatting`, `export_images_as_base64`,
  `encoding`, `paragraph_break`

### PDF Export

- `LdmPdfWriter.write(doc, output_path) -> None`
- `PdfSaveOptions` — `compliance`, `export_document_structure`, `image_compression`,
  `jpeg_quality`, `zoom_factor`, `zoom_behavior`, `export_bookmarks_outline`, `outline_options`,
  `display_doc_title` (all applied by the writer); `text_compression`, `embed_full_fonts`,
  `use_core_fonts`, `font_embedding_mode`, `page_mode`, `color_mode`, `preserve_form_fields`, and
  `memory_optimization` (accepted but not yet consumed — see
  [Scope and Limitations](#scope-and-limitations))
- `ParagraphRenderer`, `RunRenderer`, `TableRenderer`, `ShapeRenderer` — the internal rendering
  pipeline `LdmPdfWriter` drives, built on `fpdf2`

### DOCX Round-Trip Writing

- `LdmDocxWriter.write(doc, output_path) -> None` / `write_to_bytes(doc) -> bytes`
- `OoxmlSaveOptions` — `compliance` (`OoxmlCompliance.ECMA376_2006` / `ISO29500_2008_TRANSITIONAL`;
  `ISO29500_2008_STRICT` is not implemented), `compression_level` (`CompressionLevel.NORMAL` /
  `MAXIMUM` / `FAST` / `SUPER_FAST`), `zip_64_mode`, `pretty_format`
- `DocxWriterLossyWarning` — a `UserWarning` subclass callers can opt into with
  `warnings.simplefilter("error", DocxWriterLossyWarning)`

### Document Object Model

- `Document` (LDM) — `all_paragraphs`, `tables`, `all_tables`, `text`, `styles`, `lists`,
  `sections`, `header_paragraphs`, `footer_paragraphs`, `find_style(name)`, `headings(max_level)`
- `Section` — `page_setup`, `body`, `headers_footers`
- `Paragraph` — `runs`, `paragraph_format`, `list_format`, `list_label`, `text`
- `Table` / `Row` / `Cell` — `rows`, `cells`, `cell_format`, `row_format`
- `Style` — `paragraph_format`, `font`, `base_style_name`, `built_in`
- `Font` / `ParagraphFormat` / `PageSetup` — cascaded formatting attributes resolved during DOCX
  read

### Enums

- `LineStyle`, `LineSpacingRule`, `HeightRule`, `CellMerge`, `CellVerticalAlignment`, `Orientation`,
  `ParagraphAlignment`, `SectionStart`, `StyleType`, `StyleIdentifier`, `TabAlignment`, `TabLeader`,
  `Underline`, `WrapType`

</details>

## Documentation & Resources

- **[Getting started guide](https://docs.aspose.org/words/python/)** — installation, walkthroughs, and feature guides for this library.
- **[How-to guides & FAQ](https://kb.aspose.org/words/python/)** — task-focused answers for common Word-processing questions.
- **[Full API reference](https://reference.aspose.org/words/python/)** — the complete, browsable reference for all 146 public types (the [API reference](#api-reference) section above covers the essentials).
- The `ApiExamples/` scripts are written against the API this library shares with the commercial
  [`aspose-words`](https://github.com/aspose-words/Aspose.Words-for-Python-via-.NET) package — the
  same sources run there too, by replacing `aspose.words_foss` with `aspose.words` in the imports.
- Found a bug or have a feature request? [Open an issue](https://github.com/aspose-words-foss/Aspose.Words-FOSS-for-Python/issues) on GitHub.

## Scope and Limitations

- `SaveFormat.DOC` is defined as a save-format constant for API compatibility with the commercial
  product, but `Document.save()` does not implement a DOC writer — `.doc` files can be read, not
  written. Saving with an unsupported target raises `ValueError` (`Markdown`, `Text`, `PDF`, and
  `DOCX` are the four save formats actually implemented).
- Saving DOCX with `OoxmlSaveOptions.compliance` set to `OoxmlCompliance.ISO29500_2008_STRICT` is
  not implemented — the writer raises `NotImplementedError` rather than silently emitting
  non-strict output. Use `ECMA376_2006` or `ISO29500_2008_TRANSITIONAL` instead, both of which this
  edition treats as producing the same, compliant output.
- Of `PdfSaveOptions`'s 17 fields, 8 — `text_compression`, `embed_full_fonts`, `use_core_fonts`,
  `font_embedding_mode`, `page_mode`, `color_mode`, `preserve_form_fields`, and
  `memory_optimization` — exist for API forward-compatibility with the commercial Aspose.Words
  API and are not yet consumed by the PDF writer.

For DOC writing, strict ISO 29500 compliance, and the additional load/save formats and page-layout
features beyond this edition's scope, see
[Aspose.Words for Python — Enterprise Edition](https://products.aspose.com/words/python-net/),
which adds the full commercial feature set on top of the same compatible document model.

## Development and Testing

Install the development dependencies and run the example test suite:

```bash
pip install -e ".[dev]"
python -m pytest ApiExamples/ -v --rootdir=ApiExamples -c ApiExamples/pytest.ini
```

Run an individual example script directly:

```bash
python ApiExamples/convert_document.py
```

Test fixtures live under `tests/data/input/` (shared input files for `ApiExamples/`); generated
output is written to `ApiExamples/output/` (git-ignored).

## License

This project is licensed under the [MIT License](LICENSE). The MIT License permits use, copying,
modification, distribution, sublicensing, and commercial use, provided its copyright and permission
notice are retained. The software is provided without warranty.
