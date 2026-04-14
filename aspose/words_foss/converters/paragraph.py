"""Paragraph conversion utilities."""

import re
from typing import Optional

from aspose.words_foss.models import ParagraphInfo, RunFormatting, ConversionOptions
from aspose.words_foss.docx_reader import ParagraphData, RunData


class ParagraphConverter:
    """Handles conversion of paragraphs to Markdown.

    This class works with ParagraphData from the reader module,
    not with python-docx objects directly.
    """

    def __init__(self, options: ConversionOptions):
        self.options = options

    def get_paragraph_info(self, para: ParagraphData) -> ParagraphInfo:
        """Extract comprehensive information from a paragraph."""
        info = ParagraphInfo()
        style_name = para.style_name

        if style_name:
            info.style_name = style_name
            info.heading_level = self._detect_heading_level(style_name)
            info.is_quote, info.quote_level = self._detect_quote(style_name)
            info.is_code_block, info.code_language = self._detect_code_block(style_name)

        info.is_list_item = para.is_list_item
        info.list_level = para.list_level
        info.alignment = para.alignment

        return info

    def _detect_heading_level(self, style_name: str) -> int:
        """Detect heading level from style name."""
        if style_name.startswith("Heading"):
            match = re.search(r"Heading\s*(\d+)", style_name)
            if match:
                return min(int(match.group(1)), 6)
        elif "SetextHeading1" in style_name:
            return 1
        elif "SetextHeading2" in style_name:
            return 2
        return 0

    def _detect_quote(self, style_name: str) -> tuple[bool, int]:
        """Detect if paragraph is a quote and its nesting level."""
        if "Quote" not in style_name:
            return False, 0
        match = re.search(r"Quote(\d+)?", style_name)
        if match:
            level = int(match.group(1)) if match.group(1) else 1
            return True, level
        return True, 1

    def _detect_code_block(self, style_name: str) -> tuple[bool, str]:
        """Detect if paragraph is a code block and its language."""
        style_lower = style_name.lower()
        if "code" not in style_lower:
            return False, ""

        if "fencedcode" in style_lower or "fenced" in style_lower:
            if "." in style_name:
                lang = style_name.split(".")[-1]
                return True, lang
            return True, ""
        elif "indentedcode" in style_lower or "indented" in style_lower:
            return True, ""
        elif "inlinecode" in style_lower:
            return False, ""  # Inline code is handled at run level
        elif "code" in style_lower:
            if "." in style_name:
                lang = style_name.split(".")[-1]
                return True, lang
            return True, ""
        return False, ""

    def get_run_formatting(self, run: RunData) -> RunFormatting:
        """Extract formatting information from a run."""
        fmt = RunFormatting()
        fmt.bold = run.bold
        fmt.italic = run.italic
        fmt.underline = run.underline
        fmt.strikethrough = run.strikethrough
        fmt.code = run.is_code_style
        return fmt

    def format_text(self, text: str, fmt: RunFormatting, is_code_block: bool = False) -> str:
        """Apply Markdown formatting to text."""
        if not text:
            return text

        if fmt.code and not is_code_block:
            backticks = self._get_backtick_count(text)
            return f"{backticks}{text}{backticks}"

        if is_code_block:
            return text

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

    def _get_backtick_count(self, text: str) -> str:
        """Determine how many backticks to use for inline code."""
        if "`" not in text:
            return "`"
        max_consecutive = 0
        current = 0
        for char in text:
            if char == "`":
                current += 1
                max_consecutive = max(max_consecutive, current)
            else:
                current = 0
        return "`" * (max_consecutive + 1)
