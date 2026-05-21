"""
DOCX Reader — Pure Python DOCX parser (no python-docx dependency).

This package parses DOCX files using only the standard library.
It is a refactored version of the monolithic ``reader.py`` module,
split for maintainability following DRY/KISS/SOLID principles.

All public names are re-exported here for backward compatibility.
"""

from aspose.words_foss.docx_reader.constants import (
    A_NS,
    COLOR_EMPTY,
    MC_NS,
    PAGE_FIELD_SENTINEL,
    PIC_NS,
    R_NS,
    W_NS,
    WP_NS,
    WPG_NS,
    WPS_NS,
    _ALIGNMENT_MAP,
    _BODY_ANCHOR_MAP,
    _BORDER_STYLE_MAP,
    _BUILTIN_STYLE_NAME_MAP,
    _EMU_PER_MM,
    _EXT_TO_CONTENT_TYPE,
    _HIGHLIGHT_COLOR_MAP,
    _LINE_RULE_MAP,
    _NUMBER_STYLE_MAP,
    _PKG_RELS_NS,
    _SECTION_START_MAP,
    _STYLE_TYPE_MAP,
    _UNDERLINE_MAP,
)
from aspose.words_foss.docx_reader.data_classes import (
    CellData,
    NumberingInfo,
    NumberingLevel,
    ParagraphData,
    RowData,
    RunData,
    TableData,
)
from aspose.words_foss.docx_reader.utils import (
    _canonicalize_style_name,
    _collect_run_text,
    _empty_borders,
    _ext_to_content_type,
    _hex_to_ldm_color,
)
from aspose.words_foss.docx_reader.document_reader import DocumentReader

__all__ = [
    # Core class
    "DocumentReader",
    # Data classes
    "CellData",
    "NumberingInfo",
    "NumberingLevel",
    "ParagraphData",
    "RowData",
    "RunData",
    "TableData",
    # Constants
    "A_NS",
    "COLOR_EMPTY",
    "MC_NS",
    "PAGE_FIELD_SENTINEL",
    "PIC_NS",
    "R_NS",
    "W_NS",
    "WP_NS",
    "WPG_NS",
    "WPS_NS",
    # Utility functions
    "_canonicalize_style_name",
    "_collect_run_text",
    "_empty_borders",
    "_ext_to_content_type",
    "_hex_to_ldm_color",
    # Mapping constants (re-exported for backward compat)
    "_ALIGNMENT_MAP",
    "_BODY_ANCHOR_MAP",
    "_BORDER_STYLE_MAP",
    "_BUILTIN_STYLE_NAME_MAP",
    "_EMU_PER_MM",
    "_EXT_TO_CONTENT_TYPE",
    "_HIGHLIGHT_COLOR_MAP",
    "_LINE_RULE_MAP",
    "_NUMBER_STYLE_MAP",
    "_PKG_RELS_NS",
    "_SECTION_START_MAP",
    "_STYLE_TYPE_MAP",
    "_UNDERLINE_MAP",
]
