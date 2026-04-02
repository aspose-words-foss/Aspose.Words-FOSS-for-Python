"""List handling utilities for DOCX to Markdown conversion."""

from typing import Optional

from aspose.words_foss.models import ConversionOptions, ListMarker
from aspose.words_foss.reader import ParagraphData, DocumentReader


class ListHandler:
    """Handles parsing and conversion of lists.

    This class works with ParagraphData from the reader module,
    not with python-docx objects directly.
    """

    def __init__(self, options: ConversionOptions, reader: Optional[DocumentReader] = None):
        self.options = options
        self._reader = reader
        self._counters: dict[tuple[int, int], int] = {}
        self._current_list_id: Optional[int] = None

    def set_reader(self, reader: DocumentReader) -> None:
        """Set the document reader for numbering lookups."""
        self._reader = reader

    def reset(self):
        """Reset list state for new document."""
        self._counters.clear()
        self._current_list_id = None

    def get_list_info(self, para: ParagraphData) -> tuple[bool, int, str, str]:
        """Get list information for a paragraph.

        Returns:
            Tuple of (is_list_item, level, list_type, marker)
        """
        if not para.is_list_item:
            return False, 0, "", ""

        level = para.list_level
        list_id = para.list_id

        if list_id is not None:
            list_type, marker = self._determine_list_type(list_id, level)
        else:
            list_type, marker = "bullet", self.options.list_marker.value

        return True, level, list_type, marker

    def _determine_list_type(self, list_id: int, level: int) -> tuple[str, str]:
        """Determine whether list is bullet or ordered and get appropriate marker."""
        if self._reader is None:
            return "bullet", self.options.list_marker.value

        num_format, start = self._reader._get_list_format(list_id, level)

        if num_format in ("decimal", "lowerLetter", "upperLetter", "lowerRoman", "upperRoman"):
            return self._get_ordered_marker(list_id, level, start)
        else:
            return self._get_bullet_marker(level)

    def _get_ordered_marker(self, list_id: int, level: int, start: int) -> tuple[str, str]:
        """Get marker for ordered list item."""
        counter_key = (list_id, level)

        if counter_key not in self._counters:
            self._counters[counter_key] = start
        else:
            self._counters[counter_key] += 1

        number = self._counters[counter_key]
        return "ordered", f"{number}."

    def _get_bullet_marker(self, level: int) -> tuple[str, str]:
        """Get marker for bullet list item based on nesting level."""
        markers = [
            self.options.list_marker.value,
            "o" if self.options.list_marker != ListMarker.ASTERISK else "+",
            "+" if self.options.list_marker != ListMarker.PLUS else "-",
        ]
        marker = markers[level % len(markers)] if level < len(markers) else markers[-1]
        return "bullet", marker

    def format_list_item(self, text: str, level: int, marker: str) -> str:
        """Format a list item with proper indentation."""
        indent = "  " * level
        return f"{indent}{marker} {text}"

    def break_list(self):
        """Signal that the current list has been broken (e.g., by a regular paragraph)."""
        self._counters.clear()
