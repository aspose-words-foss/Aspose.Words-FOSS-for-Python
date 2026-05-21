"""DRY helpers and atomic single-element builders for OOXML→LDM."""

from typing import Optional
from xml.etree import ElementTree as ET

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_reader.constants import (
    COLOR_EMPTY,
    W_NS,
    _BORDER_SIZE_DIVISOR,
    _BORDER_STYLE_MAP,
    _TWIPS_PER_PT,
)
from aspose.words_foss.docx_reader.field_mappings import (
    FRAME_ENUM_ATTRS,
    FRAME_TWIP_ATTRS,
)
from aspose.words_foss.docx_reader.utils import (
    OOXML_BORDER_TO_SLOT,
    _hex_to_ldm_color,
    _trim_trailing_empty_borders,
)


# -- XML traversal helpers ----------------------------------------------------

def find_val(parent: ET.Element, tag: str, default: str = "") -> str:
    """Return ``w:val`` of ``<w:{tag}/>`` under *parent*, or *default*."""
    elem = parent.find(f"{W_NS}{tag}")
    return elem.get(f"{W_NS}val", default) if elem is not None else default


def twip_to_pt(value: str) -> Optional[float]:
    """Twips string → points; ``None`` on empty/invalid."""
    if not value:
        return None
    try:
        return int(value) / _TWIPS_PER_PT
    except ValueError:
        return None


def read_twip(elem: ET.Element, attr: str) -> Optional[float]:
    return twip_to_pt(elem.get(f"{W_NS}{attr}", ""))


def parse_int(value: str, default: int = 0) -> int:
    try:
        return int(value)
    except (ValueError, TypeError):
        return default


# Named-unit → points factors (subset Aspose's NrxXmlUtil supports).
_UNIT_TO_PT: dict[str, float] = {
    "pt": 1.0,
    "hp": 0.5,
    "halfpt": 0.5,
    "twip": 1.0 / 20.0,
    "twips": 1.0 / 20.0,
    "dxa": 1.0 / 20.0,
    "in": 72.0,
    "inch": 72.0,
    "cm": 28.346456692913385,    # 72 / 2.54
    "mm": 2.834645669291338,
    "pi": 12.0,
    "pc": 12.0,
    "emu": 1.0 / 12700.0,
}


def parse_universal_measure(
    value: str,
    target_unit: str,
    default: float = 0.0,
) -> float:
    """Parse OOXML ``ST_UniversalMeasure`` (or plain number) to ``target_unit``.

    Suffix-less input is treated as already in ``target_unit`` (mirrors
    Aspose's ``ReadValAs*``).  Returns ``default`` on empty / malformed.
    """
    if not value:
        return default
    last_digit = -1
    for i in range(len(value) - 1, -1, -1):
        if value[i].isdigit():
            last_digit = i
            break
    if last_digit < 0:
        return default
    numeric_part = value[: last_digit + 1]
    unit_part = value[last_digit + 1 :].strip().lower()
    try:
        n = float(numeric_part)
    except ValueError:
        return default
    if not unit_part:
        return n
    src_factor = _UNIT_TO_PT.get(unit_part)
    if src_factor is None:
        return n
    target_factor = _UNIT_TO_PT.get(target_unit.lower(), 1.0)
    return n * src_factor / target_factor


def is_truthy_onoff(raw: str) -> bool:
    return raw in ("1", "true")


_PADDING_SIDES: tuple[tuple[str, str], ...] = (
    ("left", "left_padding"),
    ("right", "right_padding"),
    ("top", "top_padding"),
    ("bottom", "bottom_padding"),
)


def apply_padding_sides(target: object, parent: ET.Element) -> None:
    """Copy ``<w:{side} w:w="..."/>`` twips onto ``{side}_padding`` fields."""
    for side, attr in _PADDING_SIDES:
        side_elem = parent.find(f"{W_NS}{side}")
        if side_elem is None:
            continue
        value = read_twip(side_elem, "w")
        if value is not None:
            setattr(target, attr, value)


# -- Stateless single-element builders ----------------------------------------

_BORDER_FULL_SLOTS = 8


def build_borders(bdr_elem: ET.Element) -> list[ldm.Border]:
    """Per-side ``ldm.Border`` list, trimmed to the 6-slot writer baseline."""
    borders: list[ldm.Border] = [
        ldm.Border(color=COLOR_EMPTY) for _ in range(_BORDER_FULL_SLOTS)
    ]
    for side_elem in bdr_elem:
        tag = side_elem.tag
        if not tag.startswith(W_NS):
            continue
        slot = OOXML_BORDER_TO_SLOT.get(tag[len(W_NS):])
        if slot is None:
            continue
        borders[slot] = _build_one_border(side_elem)
    return _trim_trailing_empty_borders(borders)


def _build_one_border(side_elem: ET.Element) -> ldm.Border:
    b = ldm.Border()
    b.line_style = _BORDER_STYLE_MAP.get(side_elem.get(f"{W_NS}val", "none"), 0)
    sz = side_elem.get(f"{W_NS}sz", "")
    if sz:
        b.line_width = int(sz) / _BORDER_SIZE_DIVISOR
    b.color = _hex_to_ldm_color(side_elem.get(f"{W_NS}color", ""))
    return b


def build_shading(shd: ET.Element) -> ldm.Shading:
    """``w:fill`` → background, ``w:color`` → foreground (auto skipped)."""
    s = ldm.Shading()
    fill = shd.get(f"{W_NS}fill", "")
    if fill:
        s.background_pattern_color = fill
    color = shd.get(f"{W_NS}color", "")
    if color and color != "auto":
        s.foreground_pattern_color = color
    return s


def build_frame(frame_pr: ET.Element) -> ldm.FrameFormat:
    """``<w:framePr/>`` → :class:`ldm.FrameFormat`."""
    ff = ldm.FrameFormat()
    for attr, target in FRAME_TWIP_ATTRS:
        raw = frame_pr.get(f"{W_NS}{attr}")
        if raw is None:
            continue
        try:
            setattr(ff, target, int(raw) / _TWIPS_PER_PT)
        except ValueError:
            continue
    for attr, target, token_map in FRAME_ENUM_ATTRS:
        raw = frame_pr.get(f"{W_NS}{attr}")
        if raw is None:
            continue
        mapped = token_map.get(raw)
        if mapped is not None:
            setattr(ff, target, mapped)
    anchor_lock = frame_pr.get(f"{W_NS}anchorLock")
    ff.lock_anchor = anchor_lock not in (None, "0", "false")
    return ff
