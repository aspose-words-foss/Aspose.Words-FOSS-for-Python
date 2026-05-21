"""
Utility functions for DOCX parsing.

Pure functions for color conversion, content-type detection,
style name canonicalization, and XML element text collection.
"""

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_reader.constants import (
    COLOR_EMPTY,
    W_NS,
    _ACRONYM_PREFIXES,
    _BUILTIN_STYLE_NAME_MAP,
    _EXT_TO_CONTENT_TYPE,
    _LOWERCASE_WORDS,
    _MAX_COLOR_CHANNEL,
)

from xml.etree import ElementTree as ET


def _hex_to_ldm_color(hex_color: str) -> str:
    """Convert a hex color string to LDM Color format.

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


#: Canonical 8-slot border order matching
#: :class:`BorderType` (0..7):
#: 0=Bottom, 1=Left, 2=Right, 3=Top, 4=Horizontal, 5=Vertical,
#: 6=DiagonalDown, 7=DiagonalUp.  Slots that don't apply to a given
#: context (e.g. diagonals on a paragraph) stay ``Color [Empty]``.
#: ``Horizontal`` carries both OOXML ``<w:insideH>`` (table/cell) and
#: ``<w:between>`` (paragraph) — both unify into one
#: ``BorderType.Horizontal``.  OOXML ``<w:bar>`` has no counterpart on
#: ``BorderType`` and is dropped on read.
BORDER_SLOTS: tuple[str, ...] = (
    "bottom",        # 0  BorderType.Bottom
    "left",          # 1  BorderType.Left
    "right",         # 2  BorderType.Right
    "top",           # 3  BorderType.Top
    "horizontal",    # 4  BorderType.Horizontal  (insideH / between)
    "vertical",      # 5  BorderType.Vertical    (insideV)
    "diagonal_down", # 6  BorderType.DiagonalDown (tl2br)
    "diagonal_up",   # 7  BorderType.DiagonalUp   (tr2bl)
)

#: OOXML side-tag → :data:`BORDER_SLOTS` index.  Handles the bidi
#: aliases ``start`` / ``end`` (synonyms for ``left`` / ``right`` in
#: right-to-left documents) and unifies the paragraph-only
#: ``<w:between>`` with the table-only ``<w:insideH>`` into the
#: ``Horizontal`` slot.  ``<w:bar>`` has no counterpart on
#: ``BorderType`` and is dropped on read.
OOXML_BORDER_TO_SLOT: dict[str, int] = {
    "bottom": 0,
    "left": 1, "start": 1,
    "right": 2, "end": 2,
    "top": 3,
    "insideH": 4, "between": 4,
    "insideV": 5,
    "tl2br": 6,
    "tr2bl": 7,
}

#: Slot baseline for the writer-shape border list: the four outer
#: borders plus Horizontal/Vertical are always present in the
#: serialized form; diagonals are grown on demand.
BORDER_MIN_SLOTS: int = 6


def _empty_borders() -> "list[ldm.Border]":
    """Return :data:`BORDER_MIN_SLOTS` empty ``Border``\\ s.

    The 8 ``BorderType`` slots collapse to 6 when the diagonals
    are unset; diagonals are grown by
    :meth:`LdmBuilderMixin._build_ldm_borders` whenever a
    ``<w:tl2br>`` / ``<w:tr2bl>`` is present in OOXML.
    """
    return [ldm.Border(color=COLOR_EMPTY) for _ in range(BORDER_MIN_SLOTS)]


def _is_empty_border(border: "ldm.Border") -> bool:
    """A border is empty when it carries no visible attributes.

    Accepts both ``Color [Empty]`` (the reader's canonical placeholder)
    and the bare ``Border()`` Pydantic default (``color=""``).
    """
    return (
        not border.line_style
        and not border.line_width
        and (not border.color or border.color == COLOR_EMPTY)
    )


def _trim_trailing_empty_borders(borders: "list[ldm.Border]") -> "list[ldm.Border]":
    """Trim trailing empty entries down to :data:`BORDER_MIN_SLOTS`.

    Preserves indexing for slots that *are* set: stops shrinking the
    instant a non-empty trailing slot is encountered, so a list with
    ``DiagonalUp`` set but ``DiagonalDown`` empty retains both
    (length 8) rather than collapsing the diagonals out of order.
    """
    while len(borders) > BORDER_MIN_SLOTS and _is_empty_border(borders[-1]):
        borders.pop()
    return borders


def parse_onoff(elem: ET.Element | None, *, default: bool = False) -> bool:
    """Parse a Word CT_OnOff element (``<w:b/>``, ``<w:keepLines/>`` …).

    Returns ``default`` when *elem* is ``None``.  When *elem* is present
    without ``@w:val``, returns ``True`` (Word's "tag present = on"
    convention).  Anything except ``"0"`` / ``"false"`` is also ``True``.
    """
    if elem is None:
        return default
    val = elem.get(f"{W_NS}val")
    return val is None or val not in ("false", "0")


def apply_onoff_attrs(
    target: object,
    parent: ET.Element | None,
    mapping: tuple[tuple[str, str, bool], ...],
) -> None:
    """Apply a batch of CT_OnOff children to attributes of *target*.

    *mapping* is an iterable of ``(child_local_name, attribute_name,
    default)`` triples.  For every triple the function looks up
    ``<w:{child_local_name}>`` under *parent*; when present, the
    matching attribute on *target* is set to the result of
    :func:`parse_onoff` (using *default* as the fallback).
    """
    if parent is None:
        return
    for tag, attr, default in mapping:
        elem = parent.find(f"{W_NS}{tag}")
        if elem is not None:
            setattr(target, attr, parse_onoff(elem, default=default))


def _ext_to_content_type(filename: str) -> str:
    """Map a filename extension to its MIME content type."""
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    return _EXT_TO_CONTENT_TYPE.get(ext, f"image/{ext}")


def _canonicalize_style_name(raw_name: str) -> str:
    """Resolve a built-in OOXML style name to its canonical form.

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


def _collect_run_text(r_elem: ET.Element) -> str:
    """Collect text from a <w:r> element, handling special characters.

    Processes <w:t>, <w:instrText>, <w:br>, <w:tab>, and <w:cr> children.
    ``<w:instrText>`` carries field-code text in the Aspose model (regular
    runs that happen to sit between ``begin`` and ``separate``); reading
    its text the same way as ``<w:t>`` lets the writer round-trip mixed
    sequences of breaks and instruction text inside one ``<w:r>``.
    """
    parts: list[str] = []
    for child in r_elem:
        if child.tag == f"{W_NS}t":
            parts.append(child.text or "")
        elif child.tag == f"{W_NS}instrText":
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
    tint: str | None = None,
    shade: str | None = None,
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
