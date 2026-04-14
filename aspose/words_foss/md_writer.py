"""Markdown writer operating on the light document model.

Converts an ldm.Document into a Markdown string.  All formatting
decisions (headings, lists, tables, inline emphasis) are driven
exclusively by the Pydantic model classes defined in
``light_document_model.py``.
"""

from __future__ import annotations

import base64
import html
import os
import re
from pathlib import Path
from typing import Optional

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.models import (
    ConversionOptions,
    HeadingStyle,
    RunFormatting,
)

# Alignment int -> string (matches light_document_model conventions)
_ALIGN_STR = {0: "left", 1: "center", 2: "right", 3: "left"}


class LdmMarkdownWriter:
    """Converts a ``light_document_model.Document`` to a Markdown string."""

    # Block-type tags used by _join_blocks to decide blank-line insertion.
    _LIST = "list"
    _BLOCK = "block"
    _BLANK = "blank"

    def __init__(self, options: Optional[ConversionOptions] = None):
        self.options = options or ConversionOptions()
        self._list_counters: dict[tuple[int, int], int] = {}
        self._doc: Optional[ldm.Document] = None
        self._output_path: Optional[Path] = None
        self._image_counter: int = 0
        self._reference_links: list[tuple[str, str]] = []  # (label, url)

    # ------------------------------------------------------------------
    # Public entry point
    # ------------------------------------------------------------------

    def write(
        self,
        doc: ldm.Document,
        output_path: Optional[Path] = None,
    ) -> str:
        """Convert *doc* to Markdown and return the result string."""
        self._list_counters.clear()
        self._reference_links.clear()
        self._image_counter = 0
        self._doc = doc
        self._output_path = output_path
        blocks: list[tuple[str, str]] = []  # (tag, markdown_text)

        # Header images
        for para in doc.header_paragraphs:
            for item in para.inline_extras:
                if isinstance(item, ldm.ShapeNode) and item.has_image and item.image_data:
                    blocks.append((self._BLOCK, self._render_image(item)))

        # Body
        for section in doc.sections:
            for child in section.body.children:
                if isinstance(child, ldm.Paragraph):
                    tag, md = self._convert_paragraph_tagged(child)
                    if md is not None:
                        blocks.append((tag, md))
                elif isinstance(child, ldm.Table):
                    md = self._convert_table(child)
                    blocks.append((self._BLOCK, md))
                # UnknownNode is silently skipped

        # Footer images
        for para in doc.footer_paragraphs:
            for item in para.inline_extras:
                if isinstance(item, ldm.ShapeNode) and item.has_image and item.image_data:
                    blocks.append((self._BLOCK, self._render_image(item)))

        result = self._join_blocks(blocks)

        # Append reference-style link definitions if any were collected
        if self._reference_links:
            defs = "\n".join(f"[{label}]: {url}" for label, url in self._reference_links)
            result = result.rstrip("\n") + "\n\n" + defs + "\n"

        return result

    # ------------------------------------------------------------------
    # Block joining
    # ------------------------------------------------------------------

    def _join_blocks(self, blocks: list[tuple[str, str]]) -> str:
        """Join block-level elements with blank lines between them.

        Per the CommonMark spec, consecutive non-blank lines that are not
        otherwise recognised as block-level constructs form a single
        paragraph (soft line breaks become spaces).  Every block element
        must therefore be separated by a blank line so that each one is
        rendered as its own block.

        The one exception is consecutive list items: they must stay on
        adjacent lines to form a *tight* list.
        """
        result: list[str] = []
        prev_tag: Optional[str] = None
        for tag, text in blocks:
            if not text.strip():
                # Explicit blank -- collapse consecutive blanks.
                if prev_tag != self._BLANK:
                    result.append("")
                    prev_tag = self._BLANK
                continue

            # Decide whether a blank line is needed before this block.
            need_blank = False
            if prev_tag is not None and prev_tag != self._BLANK:
                if tag == self._LIST and prev_tag == self._LIST:
                    # Consecutive list items -> tight list, no blank line.
                    need_blank = False
                else:
                    need_blank = True

            if need_blank:
                result.append("")
            result.append(text)
            prev_tag = tag

        output = "\n".join(result)
        return output.rstrip() + "\n" if output else ""

    # ------------------------------------------------------------------
    # Paragraph conversion
    # ------------------------------------------------------------------

    def _render_image(self, shape: ldm.ShapeNode) -> str:
        """Render a ShapeNode with image data as a Markdown inline image tag."""
        img = shape.image_data
        assert img is not None

        # If images_folder is set, save to file instead of base64
        if self.options.images_folder and not self.options.export_images_as_base64:
            self._image_counter += 1
            ext = self._guess_image_extension(img.content_type, img.source_filename)
            # Sanitize filename to prevent path traversal
            filename = Path(img.source_filename).name if img.source_filename else ""
            filename = filename or f"image{self._image_counter}{ext}"
            folder = Path(self.options.images_folder)
            folder.mkdir(parents=True, exist_ok=True)
            filepath = folder / filename
            filepath.write_bytes(img.image_bytes)

            # Use alias if set, otherwise compute relative path
            if self.options.images_folder_alias:
                url = f"{self.options.images_folder_alias.rstrip('/')}/{filename}"
            elif self._output_path is not None:
                url = os.path.relpath(filepath, self._output_path.parent)
            else:
                url = str(filepath)
            return f"![{img.source_filename}]({url})"

        # Default: inline base64 data URI
        b64 = base64.b64encode(img.image_bytes).decode("ascii")
        return f"![{img.source_filename}](data:{img.content_type};base64,{b64})"

    @staticmethod
    def _guess_image_extension(content_type: str, filename: str) -> str:
        """Determine file extension from content type or filename."""
        if filename and "." in filename:
            return ""  # filename already has extension
        mime_map = {
            "image/png": ".png",
            "image/jpeg": ".jpg",
            "image/gif": ".gif",
            "image/bmp": ".bmp",
            "image/svg+xml": ".svg",
            "image/tiff": ".tiff",
        }
        return mime_map.get(content_type, ".png")

    def _is_list_paragraph(self, para: ldm.Paragraph) -> bool:
        """Return True if this paragraph is a list item."""
        return bool(para.list_format and para.list_format.is_list_item)

    def _convert_paragraph_tagged(self, para: ldm.Paragraph) -> tuple[str, Optional[str]]:
        """Convert a paragraph and return (block_tag, markdown_text)."""
        is_list = self._is_list_paragraph(para)
        # In plain_text list mode, list items are regular blocks (no tight list)
        if is_list and self.options.list_export_mode == "plain_text":
            tag = self._BLOCK
        else:
            tag = self._LIST if is_list else self._BLOCK
        md = self._convert_paragraph(para)
        return tag, md

    def _convert_paragraph(self, para: ldm.Paragraph) -> Optional[str]:
        pf = para.paragraph_format
        style_name = pf.style_name

        # Horizontal rule takes precedence over everything else
        if self._is_horizontal_rule(para):
            return "---"

        # Handle empty paragraphs according to empty_paragraph_export_mode
        if self._is_empty_paragraph(para):
            mode = self.options.empty_paragraph_export_mode
            if mode == "none":
                return None
            elif mode == "markdown_hard_line_break":
                return "\\"
            # "empty_line" is the default — returns ""

        is_code_block = bool(style_name and ("Code" in style_name or "code" in style_name))

        # Use content_sequence only when the paragraph actually mixes images with
        # text runs.  Every DOCX-reader paragraph has a non-empty content_sequence,
        # so checking merely for non-emptiness would send image-free paragraphs
        # (code-blocks, quotes, list items) down this path unnecessarily.
        if any(isinstance(i, ldm.ShapeNode) for i in para.content_sequence):
            output_parts: list[str] = []
            pending_runs: list[ldm.Run] = []

            def _flush_runs() -> None:
                if not pending_runs:
                    return
                text = self._convert_runs(pending_runs, is_code_block, para)
                pending_runs.clear()
                if text:
                    formatted = self._format_text_part(text, pf, style_name, is_code_block, para)
                    if formatted is not None:
                        output_parts.append(formatted)

            for item in para.content_sequence:
                if (
                    isinstance(item, ldm.ShapeNode)
                    and item.has_image
                    and item.image_data is not None
                ):
                    _flush_runs()
                    output_parts.append(self._render_image(item))
                elif isinstance(item, ldm.Run):
                    pending_runs.append(item)
            _flush_runs()

            if not output_parts:
                return ""
            return "\n".join(output_parts)

        # Legacy path: no content_sequence (non-reader sources or text-only paragraphs).
        # Images from inline_extras come first, then the paragraph text.
        image_parts = [
            self._render_image(item)
            for item in para.inline_extras
            if isinstance(item, ldm.ShapeNode) and item.has_image and item.image_data is not None
        ]

        text = self._convert_runs(para.runs, is_code_block, para)
        text_part = (
            self._format_text_part(text, pf, style_name, is_code_block, para) if text else None
        )

        parts = image_parts[:]
        if text_part is not None:
            parts.append(text_part)

        if not parts:
            return ""
        return "\n".join(parts)

    def _format_text_part(
        self,
        text: str,
        pf: ldm.ParagraphFormat,
        style_name: str,
        is_code_block: bool,
        para: ldm.Paragraph,
    ) -> Optional[str]:
        """Apply paragraph-level formatting to an already-assembled text string."""
        if not text:
            return None
        if pf.is_heading:
            level = min(pf.outline_level + 1, 6)
            return self._format_heading(text, level, style_name)
        if is_code_block:
            lang = self._extract_code_language(style_name)
            return self._format_code_block(text, lang, style_name)
        if style_name and "Quote" in style_name:
            level = self._extract_quote_level(style_name)
            return self._format_quote(text, level)
        if para.list_format and para.list_format.is_list_item:
            return self._format_list_item(text, para)
        return text

    # ------------------------------------------------------------------
    # Horizontal rule detection
    # ------------------------------------------------------------------

    @staticmethod
    def _is_empty_paragraph(para: ldm.Paragraph) -> bool:
        """Return True if the paragraph has no visible content."""
        if para.runs and any(r.text and r.text.strip() for r in para.runs):
            return False
        if para.content_sequence and any(
            isinstance(i, ldm.ShapeNode) for i in para.content_sequence
        ):
            return False
        if para.inline_extras and any(
            isinstance(i, ldm.ShapeNode) and i.has_image for i in para.inline_extras
        ):
            return False
        text = para.text.strip() if para.text else ""
        return not text

    def _is_horizontal_rule(self, para: ldm.Paragraph) -> bool:
        pf = para.paragraph_format

        # Bottom-border based detection
        if len(pf.borders) >= 3:
            bottom = pf.borders[2]  # top, left, bottom, ...
            if bottom.line_style > 0 and bottom.line_width >= 1.5:
                return True

        # Text-pattern based detection
        text = para.text.strip()
        if re.match(r"^[-*_]{3,}$", text.replace(" ", "")):
            return True

        return False

    # ------------------------------------------------------------------
    # Run conversion
    # ------------------------------------------------------------------

    # Regex for pre-rendered inline links: [text](url)
    _INLINE_LINK_RE = re.compile(r"\[([^\]]+)\]\(([^)]+)\)")

    def _convert_runs(
        self,
        runs: list[ldm.Run],
        is_code_block: bool,
        para: Optional[ldm.Paragraph] = None,
    ) -> str:
        # When the paragraph style itself defines bold/italic, suppress those
        # markers on runs so Markdown doesn't double-apply emphasis.
        style_bold, style_italic = self._get_style_emphasis(para)

        parts: list[str] = []
        for run in runs:
            text = run.text or ""
            if not text:
                continue
            fmt = self._get_run_formatting(run)
            if style_bold and fmt.bold:
                fmt.bold = False
            if style_italic and fmt.italic:
                fmt.italic = False
            formatted = self._apply_formatting(text, fmt, is_code_block)
            parts.append(formatted)
        result = "".join(parts)

        # Transform links based on link_export_mode
        if self.options.link_export_mode == "reference":
            result = self._convert_links_to_reference(result)

        return result

    def _get_style_emphasis(self, para: Optional[ldm.Paragraph]) -> tuple[bool, bool]:
        """Return (bold, italic) defined by the paragraph's style font.

        Looks up the style in the document's styles list and checks its
        font properties.  This lets us suppress emphasis inherited from the
        style rather than applied as direct run formatting.
        """
        if para is None or self._doc is None:
            return False, False
        style_name = para.paragraph_format.style_name
        if not style_name:
            return False, False
        style = self._doc.find_style(style_name)
        if style is None or style.font is None:
            return False, False
        return style.font.bold, style.font.italic

    def _convert_links_to_reference(self, text: str) -> str:
        """Convert inline links [text](url) to reference-style [text][N]."""

        def _replace(m: re.Match) -> str:
            label = str(len(self._reference_links) + 1)
            self._reference_links.append((label, m.group(2)))
            return f"[{m.group(1)}][{label}]"

        return self._INLINE_LINK_RE.sub(_replace, text)

    def _get_run_formatting(self, run: ldm.Run) -> RunFormatting:
        f = run.font
        fmt = RunFormatting()
        fmt.bold = f.bold
        fmt.italic = f.italic
        fmt.underline = f.underline > 0
        fmt.strikethrough = f.strike_through
        fmt.code = bool(f.style_name and "code" in f.style_name.lower())
        return fmt

    def _apply_formatting(self, text: str, fmt: RunFormatting, is_code_block: bool) -> str:
        if not text:
            return text

        if self.options.escape_special_chars and not fmt.code and not is_code_block:
            text = self._escape_markdown(text)

        if fmt.code:
            return f"`{text}`"

        result = text
        if fmt.bold and fmt.italic:
            result = f"***{result}***"
        elif fmt.bold:
            result = f"**{result}**"
        elif fmt.italic:
            result = f"*{result}*"

        if fmt.strikethrough and self.options.export_strikethrough:
            result = f"~~{result}~~"

        if fmt.underline and self.options.export_underline:
            result = f"++{result}++"

        return result

    @staticmethod
    def _escape_markdown(text: str) -> str:
        return text

    # ------------------------------------------------------------------
    # Heading
    # ------------------------------------------------------------------

    def _format_heading(self, text: str, level: int, style_name: str) -> str:
        level = min(max(level, 1), 6)
        use_setext = self.options.heading_style == HeadingStyle.SETEXT or "Setext" in style_name
        if use_setext and level <= 2:
            underline = "=" if level == 1 else "-"
            return f"{text}\n{underline * len(text)}"
        return f"{'#' * level} {text}"

    # ------------------------------------------------------------------
    # Code block
    # ------------------------------------------------------------------

    def _extract_code_language(self, style_name: str) -> str:
        if "." in style_name:
            return style_name.split(".")[-1]
        return ""

    def _format_code_block(self, text: str, language: str, style_name: str = "") -> str:
        is_indented = (
            self.options.code_block_style.value == "indented" or "indented" in style_name.lower()
        )
        if not is_indented:
            lang_spec = language.lower() if language else ""
            return f"```{lang_spec}\n{text}\n```"
        else:
            lines = text.split("\n")
            return "\n".join(f"    {line}" for line in lines)

    # ------------------------------------------------------------------
    # Block quote
    # ------------------------------------------------------------------

    def _extract_quote_level(self, style_name: str) -> int:
        match = re.search(r"Quote(\d+)?", style_name)
        if match and match.group(1):
            return int(match.group(1))
        return 1

    def _format_quote(self, text: str, level: int) -> str:
        prefix = "> " * level
        lines = text.split("\n")
        return "\n".join(f"{prefix}{line}" for line in lines)

    # ------------------------------------------------------------------
    # List item
    # ------------------------------------------------------------------

    def _format_list_item(self, text: str, para: ldm.Paragraph) -> str:
        lf = para.list_format
        assert lf is not None
        level = lf.list_level_number
        list_id = lf.list_id

        # Plain text list mode: render as indented text without markers
        if self.options.list_export_mode == "plain_text":
            indent = "    " * level
            return f"{indent}{text}"

        list_type, marker = self._get_list_type(list_id, level)
        indent = "  " * level
        return f"{indent}{marker} {text}"

    def _get_list_type(self, list_id: int, level: int) -> tuple[str, str]:
        """Determine list type and marker from the document's list definitions."""
        if self._doc is not None and list_id > 0:
            for dl in self._doc.lists:
                if dl.list_id == list_id:
                    if level < len(dl.levels):
                        ll = dl.levels[level]
                    elif dl.levels:
                        ll = dl.levels[0]
                    else:
                        return ("bullet", self.options.list_marker.value)

                    # number_style: 0=decimal, 1-4=letter/roman, 23=bullet
                    if ll.number_style not in (23, 255):
                        return self._get_ordered_marker(list_id, level, ll.start_at)
                    else:
                        return ("bullet", self.options.list_marker.value)

        return ("bullet", self.options.list_marker.value)

    def _get_ordered_marker(self, list_id: int, level: int, start: int) -> tuple[str, str]:
        counter_key = (list_id, level)
        if counter_key not in self._list_counters:
            self._list_counters[counter_key] = start
        else:
            self._list_counters[counter_key] += 1
        num = self._list_counters[counter_key]
        return ("ordered", f"{num}.")

    # ------------------------------------------------------------------
    # Table conversion
    # ------------------------------------------------------------------

    def _convert_table(self, table: ldm.Table) -> str:
        if not table.rows:
            return ""

        # export_as_html: render table as raw HTML
        if self.options.export_as_html == "tables":
            return self._convert_table_as_html(table)

        # Determine number of columns
        num_cols = max(len(row.cells) for row in table.rows)

        # Extract cell texts and alignments
        cell_data: list[list[tuple[str, str]]] = []
        for row in table.rows:
            row_cells: list[tuple[str, str]] = []
            for cell in row.cells:
                text = self._extract_cell_text(cell)
                align = self._resolve_cell_alignment(cell)
                row_cells.append((text, align))
            cell_data.append(row_cells)

        # Column widths
        col_widths = [3] * num_cols
        for row_cells in cell_data:
            for j, (text, _) in enumerate(row_cells):
                if j < num_cols:
                    col_widths[j] = max(col_widths[j], len(text))

        # Render rows
        lines: list[str] = []
        for i, row_cells in enumerate(cell_data):
            cells: list[str] = []
            for j in range(num_cols):
                if j < len(row_cells):
                    text = row_cells[j][0]
                else:
                    text = ""
                cells.append(text.ljust(col_widths[j]))
            lines.append("| " + " | ".join(cells) + " |")

            # Header separator after the first row
            if i == 0:
                seps: list[str] = []
                for j in range(num_cols):
                    width = col_widths[j]
                    align = row_cells[j][1] if j < len(row_cells) else "left"
                    if align == "center":
                        sep = ":" + "-" * (width - 2) + ":"
                    elif align == "right":
                        sep = "-" * (width - 1) + ":"
                    else:
                        sep = "-" * width
                    seps.append(sep)
                lines.append("| " + " | ".join(seps) + " |")

        return "\n".join(lines)

    def _extract_cell_text(self, cell: ldm.Cell) -> str:
        parts: list[str] = []
        for para in cell.paragraphs:
            if para.content_sequence:
                # Preserve XML element order using content_sequence
                para_parts: list[str] = []
                for item in para.content_sequence:
                    if isinstance(item, ldm.ShapeNode) and item.has_image and item.image_data:
                        para_parts.append(self._render_image(item))
                    elif isinstance(item, ldm.Run):
                        text = item.text or ""
                        if text:
                            fmt = self._get_run_formatting(item)
                            para_parts.append(self._apply_formatting(text, fmt, False))
                if para_parts:
                    parts.append("".join(para_parts))
            else:
                # Legacy path: images first, then text runs
                for item in para.inline_extras:
                    if isinstance(item, ldm.ShapeNode) and item.has_image and item.image_data:
                        parts.append(self._render_image(item))
                run_parts: list[str] = []
                for run in para.runs:
                    text = run.text or ""
                    if text:
                        fmt = self._get_run_formatting(run)
                        text = self._apply_formatting(text, fmt, False)
                        run_parts.append(text)
                if run_parts:
                    parts.append("".join(run_parts))

        result = " ".join(parts)
        result = result.replace("|", "\\|")
        result = result.replace("\n", " ").replace("\r", " ")
        return result.strip()

    @staticmethod
    def _cell_alignment(cell: ldm.Cell) -> str:
        if cell.paragraphs:
            a = cell.paragraphs[0].paragraph_format.alignment
            return _ALIGN_STR.get(a, "left")
        return "left"

    def _resolve_cell_alignment(self, cell: ldm.Cell) -> str:
        """Return cell alignment, respecting table_content_alignment override."""
        override = self.options.table_content_alignment
        if override and override != "auto":
            return override
        return self._cell_alignment(cell)

    def _convert_table_as_html(self, table: ldm.Table) -> str:
        """Render a table as raw HTML."""
        lines: list[str] = ["<table>"]
        for i, row in enumerate(table.rows):
            lines.append("<tr>")
            tag = "th" if i == 0 else "td"
            for cell in row.cells:
                text = html.escape(self._extract_cell_text(cell))
                align = self._resolve_cell_alignment(cell)
                if align != "left":
                    lines.append(f'<{tag} style="text-align: {align}">{text}</{tag}>')
                else:
                    lines.append(f"<{tag}>{text}</{tag}>")
            lines.append("</tr>")
        lines.append("</table>")
        return "\n".join(lines)
