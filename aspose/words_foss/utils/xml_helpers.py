"""XML helper utilities for parsing DOCX files."""

from typing import Iterator, Optional
from xml.etree import ElementTree as ET

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def get_element_text(element: ET.Element, default: str = "") -> str:
    """Get text content from an XML element."""
    if element is None:
        return default
    return element.text or default


def find_all_elements(root: ET.Element, tag: str) -> Iterator[ET.Element]:
    """Find all elements with the given tag."""
    for elem in root.iter(tag):
        yield elem


def get_attribute(element: ET.Element, attr: str, default: str = "") -> str:
    """Get an attribute value from an element."""
    if element is None:
        return default
    return element.get(f"{W_NS}{attr}", default)


def has_element(parent: ET.Element, tag: str) -> bool:
    """Check if parent has a child element with the given tag."""
    return parent.find(f"{W_NS}{tag}") is not None


def get_bool_val(element: Optional[ET.Element]) -> bool:
    """Get boolean value from an element with optional val attribute."""
    if element is None:
        return False
    val = element.get(f"{W_NS}val")
    if val is None:
        return True  # Presence without val means true
    return val.lower() not in ("false", "0", "off")
