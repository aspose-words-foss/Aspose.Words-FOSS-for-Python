"""
Utility functions for DOCX parsing.

Pure functions for color conversion, content-type detection,
style name canonicalization, and XML element text collection.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from aspose.words_foss.docx_reader.constants import (
    COLOR_EMPTY,
    W_NS,
    _ACRONYM_PREFIXES,
    _BUILTIN_STYLE_NAME_MAP,
    _EXT_TO_CONTENT_TYPE,
    _LOWERCASE_WORDS,
    _MAX_COLOR_CHANNEL,
)

if TYPE_CHECKING:
    from xml.etree import ElementTree as ET
    from aspose.words_foss import light_document_model as ldm


def _hex_to_aspose_color(hex_color: str) -> str:
    """Convert a hex color string to Aspose.Words Color format.

    Examples:
        "" or "auto"  -> "Color [Empty]"
        "000000"       -> "Color [A=255, R=0, G=0, B=0]"
        "FF0000"       -> "Color [A=255, R=255, G=0, B=0]"
    """
    if not hex_color or hex_color.lower() == "auto":
        return COLOR_EMPTY
    hex_color = hex_color.lstrip("#")
    if len(hex_color) != 6:
        return COLOR_EMPTY
    try:
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
    except ValueError:
        return COLOR_EMPTY
    return f"Color [A=255, R={r}, G={g}, B={b}]"


def _empty_borders() -> "list[ldm.Border]":
    """Return 6 default Border objects with Color [Empty]."""
    from aspose.words_foss import light_document_model as ldm

    return [ldm.Border(color=COLOR_EMPTY) for _ in range(6)]


def _ext_to_content_type(filename: str) -> str:
    """Map a filename extension to its MIME content type."""
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    return _EXT_TO_CONTENT_TYPE.get(ext, f"image/{ext}")


def _canonicalize_style_name(raw_name: str) -> str:
    """Resolve a built-in OOXML style name to its Aspose.Words canonical form.

    Rules applied in order:
    1. Explicit overrides (word renames, outline lists, etc.)
    2. Acronym prefixes: "toc 1" → "TOC 1", "toa heading" → "TOA Heading"
    3. Title-case with preserved lowercase words: "table of figures" → "Table of Figures"
    4. If the name is already mixed/upper case, return as-is (custom style).
    """
    # 1. Explicit override
    canonical = _BUILTIN_STYLE_NAME_MAP.get(raw_name)
    if canonical is not None:
        return canonical

    # Only transform names that are fully lowercase (built-in convention).
    # Mixed-case names (e.g. "MyCustomStyle") are custom and returned as-is.
    if raw_name != raw_name.lower():
        return raw_name

    parts = raw_name.split()
    if not parts:
        return raw_name

    # 2. Acronym prefix
    if parts[0] in _ACRONYM_PREFIXES:
        rest = [w.title() for w in parts[1:]]
        return " ".join([parts[0].upper()] + rest)

    # 3. Title-case with lowercase words preserved (first word always capitalised)
    result = [parts[0].title()]
    for w in parts[1:]:
        result.append(w if w in _LOWERCASE_WORDS else w.title())
    return " ".join(result)


def _collect_run_text(r_elem: "ET.Element") -> str:
    """Collect text from a <w:r> element, handling special characters.

    Processes <w:t>, <w:br>, <w:tab>, and <w:cr> children.
    """
    parts: list[str] = []
    for child in r_elem:
        if child.tag == f"{W_NS}t":
            parts.append(child.text or "")
        elif child.tag == f"{W_NS}br":
            br_type = child.get(f"{W_NS}type", "")
            if br_type == "page":
                parts.append("\f")
            elif br_type == "column":
                parts.append("\v")
            else:
                parts.append("\n")
        elif child.tag == f"{W_NS}tab":
            parts.append("\t")
        elif child.tag == f"{W_NS}cr":
            parts.append("\r")
    return "".join(parts)


def _apply_theme_color_modifiers(
    base_hex: str,
    *,
    tint: "str | None" = None,
    shade: "str | None" = None,
) -> str:
    """Apply ``w:themeTint`` / ``w:themeShade`` modifiers to a base RGB.

    Both are hex strings in the 00–FF range.  ``themeTint`` mixes
    toward white, ``themeShade`` mixes toward black.  At most one
    should be present on a given ``w:color`` element; if both are
    given, tint wins (matches Word's behaviour).
    """
    try:
        r = int(base_hex[0:2], 16)
        g = int(base_hex[2:4], 16)
        b = int(base_hex[4:6], 16)
    except (ValueError, IndexError):
        return base_hex

    if tint:
        try:
            t = int(tint, 16) / float(_MAX_COLOR_CHANNEL)
            # w:themeTint=0 -> fully white, 255 -> unchanged.
            r = int(r + (_MAX_COLOR_CHANNEL - r) * (1.0 - t))
            g = int(g + (_MAX_COLOR_CHANNEL - g) * (1.0 - t))
            b = int(b + (_MAX_COLOR_CHANNEL - b) * (1.0 - t))
        except ValueError:
            pass
    elif shade:
        try:
            s = int(shade, 16) / float(_MAX_COLOR_CHANNEL)
            # w:themeShade=0 -> fully black, 255 -> unchanged.
            r = int(r * s)
            g = int(g * s)
            b = int(b * s)
        except ValueError:
            pass

    r, g, b = (max(0, min(_MAX_COLOR_CHANNEL, c)) for c in (r, g, b))
    return f"{r:02X}{g:02X}{b:02X}"
