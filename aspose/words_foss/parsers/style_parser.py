"""Style parsing utilities for DOCX documents."""

import re
from typing import Optional
from dataclasses import dataclass


@dataclass
class ParsedStyle:
    """Parsed style information."""

    name: str
    base_style: Optional[str] = None
    heading_level: int = 0
    is_quote: bool = False
    quote_level: int = 0
    is_code: bool = False
    code_type: str = ""  # "fenced", "indented", "inline"
    code_language: str = ""
    is_list_paragraph: bool = False


class StyleParser:
    """Parser for DOCX style names and properties.

    Heading detection uses ``style_identifier`` first (locale-independent),
    then falls back to the regex-based name matching for custom styles.
    """

    HEADING_PATTERN = re.compile(r"Heading\s*(\d+)", re.IGNORECASE)
    SETEXT_HEADING_PATTERN = re.compile(r"SetextHeading(\d+)", re.IGNORECASE)
    QUOTE_PATTERN = re.compile(r"Quote(\d*)", re.IGNORECASE)
    CODE_PATTERN = re.compile(r"(Fenced|Indented|Inline)?Code\.?(\w+)?", re.IGNORECASE)

    def parse(self, style_name: str, *, style_identifier: int = 0) -> ParsedStyle:
        """Parse a style name into structured information.

        ``style_identifier`` is the locale-independent built-in ID.  When
        it falls in the heading range (1-9), the regex fallback is skipped.
        """
        if not style_name:
            return ParsedStyle(name="Normal")

        result = ParsedStyle(name=style_name)

        result.heading_level = self._parse_heading(style_name, style_identifier)
        result.is_quote, result.quote_level = self._parse_quote(style_name)
        result.is_code, result.code_type, result.code_language = self._parse_code(style_name)
        result.is_list_paragraph = self._is_list_style(style_name)

        return result

    def _parse_heading(self, style_name: str, style_identifier: int = 0) -> int:
        """Extract heading level — prefer style_identifier over name regex."""
        if 1 <= style_identifier <= 9:
            return min(style_identifier, 6)

        setext_match = self.SETEXT_HEADING_PATTERN.search(style_name)
        if setext_match:
            return min(int(setext_match.group(1)), 6)

        heading_match = self.HEADING_PATTERN.search(style_name)
        if heading_match:
            return min(int(heading_match.group(1)), 6)

        return 0

    def _parse_quote(self, style_name: str) -> tuple[bool, int]:
        """Extract quote information from style name."""
        match = self.QUOTE_PATTERN.search(style_name)
        if match:
            level_str = match.group(1)
            level = int(level_str) if level_str else 1
            return True, level
        return False, 0

    def _parse_code(self, style_name: str) -> tuple[bool, str, str]:
        """Extract code block information from style name."""
        style_lower = style_name.lower()

        if "code" not in style_lower:
            return False, "", ""

        code_type = ""
        language = ""

        if "fencedcode" in style_lower or "fenced" in style_lower:
            code_type = "fenced"
        elif "indentedcode" in style_lower:
            code_type = "indented"
        elif "inlinecode" in style_lower:
            code_type = "inline"
        else:
            code_type = "fenced"  # Default

        if "." in style_name:
            parts = style_name.split(".")
            if len(parts) > 1:
                language = parts[-1]

        return True, code_type, language

    def _is_list_style(self, style_name: str) -> bool:
        """Check if style represents a list paragraph."""
        style_lower = style_name.lower()
        return "list" in style_lower and "paragraph" in style_lower

    def get_style_chain(self, style_names: list[str]) -> list[ParsedStyle]:
        """Parse a chain of style names (for inherited styles)."""
        return [self.parse(name) for name in style_names]

    def is_setext_heading(self, style_name: str) -> bool:
        """Check if style is a Setext-style heading."""
        return bool(self.SETEXT_HEADING_PATTERN.search(style_name))

    def extract_all_styles(self, style_chain: str) -> list[str]:
        """Extract individual style names from a comma-separated chain."""
        if not style_chain:
            return []
        return [s.strip() for s in style_chain.split(",") if s.strip()]
