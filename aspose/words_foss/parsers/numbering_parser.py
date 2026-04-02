"""Numbering/list parsing utilities for DOCX documents."""

from typing import Optional
from dataclasses import dataclass
from xml.etree import ElementTree as ET

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


@dataclass
class ListLevelInfo:
    """Information about a list level."""

    level: int
    num_format: str  # decimal, bullet, lowerLetter, etc.
    start_value: int
    delimiter: str  # "." or ")"
    text_pattern: str  # e.g., "%1." or "%1)"


@dataclass
class ListInfo:
    """Information about a list."""

    list_id: int
    abstract_num_id: int
    levels: dict[int, ListLevelInfo]


class NumberingParser:
    """Parser for DOCX numbering definitions."""

    def __init__(self):
        self._lists: dict[int, ListInfo] = {}
        self._abstract_nums: dict[int, dict[int, ListLevelInfo]] = {}

    def parse_numbering_part(self, numbering_element: ET.Element) -> None:
        """Parse the numbering.xml content."""
        self._parse_abstract_nums(numbering_element)
        self._parse_nums(numbering_element)

    def _parse_abstract_nums(self, root: ET.Element) -> None:
        """Parse abstract numbering definitions."""
        for abstract_num in root.findall(f".//{W_NS}abstractNum"):
            abs_id = int(abstract_num.get(f"{W_NS}abstractNumId", "0"))
            levels = {}

            for lvl in abstract_num.findall(f"{W_NS}lvl"):
                level_info = self._parse_level(lvl)
                levels[level_info.level] = level_info

            self._abstract_nums[abs_id] = levels

    def _parse_level(self, lvl: ET.Element) -> ListLevelInfo:
        """Parse a single list level definition."""
        level = int(lvl.get(f"{W_NS}ilvl", "0"))

        num_fmt = lvl.find(f"{W_NS}numFmt")
        num_format = num_fmt.get(f"{W_NS}val", "bullet") if num_fmt is not None else "bullet"

        start = lvl.find(f"{W_NS}start")
        start_value = int(start.get(f"{W_NS}val", "1")) if start is not None else 1

        lvl_text = lvl.find(f"{W_NS}lvlText")
        text_pattern = lvl_text.get(f"{W_NS}val", "") if lvl_text is not None else ""
        delimiter = ")" if ")" in text_pattern else "."

        return ListLevelInfo(
            level=level,
            num_format=num_format,
            start_value=start_value,
            delimiter=delimiter,
            text_pattern=text_pattern,
        )

    def _parse_nums(self, root: ET.Element) -> None:
        """Parse numbering instances."""
        for num in root.findall(f".//{W_NS}num"):
            num_id = int(num.get(f"{W_NS}numId", "0"))

            abstract_num_id_elem = num.find(f"{W_NS}abstractNumId")
            abstract_num_id = (
                int(abstract_num_id_elem.get(f"{W_NS}val", "0"))
                if abstract_num_id_elem is not None
                else 0
            )

            levels = dict(self._abstract_nums.get(abstract_num_id, {}))

            for lvl_override in num.findall(f"{W_NS}lvlOverride"):
                override_level = int(lvl_override.get(f"{W_NS}ilvl", "0"))

                start_override = lvl_override.find(f"{W_NS}startOverride")
                if start_override is not None:
                    new_start = int(start_override.get(f"{W_NS}val", "1"))
                    if override_level in levels:
                        levels[override_level] = ListLevelInfo(
                            level=levels[override_level].level,
                            num_format=levels[override_level].num_format,
                            start_value=new_start,
                            delimiter=levels[override_level].delimiter,
                            text_pattern=levels[override_level].text_pattern,
                        )

                lvl = lvl_override.find(f"{W_NS}lvl")
                if lvl is not None:
                    level_info = self._parse_level(lvl)
                    levels[override_level] = level_info

            self._lists[num_id] = ListInfo(
                list_id=num_id, abstract_num_id=abstract_num_id, levels=levels
            )

    def get_list_info(self, list_id: int) -> Optional[ListInfo]:
        """Get information about a specific list."""
        return self._lists.get(list_id)

    def get_level_info(self, list_id: int, level: int) -> Optional[ListLevelInfo]:
        """Get information about a specific list level."""
        list_info = self._lists.get(list_id)
        if list_info:
            return list_info.levels.get(level)
        return None

    def is_ordered_list(self, list_id: int, level: int) -> bool:
        """Check if a list level is ordered (numbered)."""
        level_info = self.get_level_info(list_id, level)
        if level_info:
            return level_info.num_format in (
                "decimal",
                "lowerLetter",
                "upperLetter",
                "lowerRoman",
                "upperRoman",
                "ordinal",
            )
        return False

    def get_start_value(self, list_id: int, level: int) -> int:
        """Get the starting value for a list level."""
        level_info = self.get_level_info(list_id, level)
        return level_info.start_value if level_info else 1

    def get_delimiter(self, list_id: int, level: int) -> str:
        """Get the delimiter for a list level."""
        level_info = self.get_level_info(list_id, level)
        return level_info.delimiter if level_info else "."
