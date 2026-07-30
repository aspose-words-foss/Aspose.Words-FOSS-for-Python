"""
Field-mapping tables shared between the LDM builder and writers.

Pure data: tuples and dicts that pair OOXML attributes / child names with
their Pydantic counterparts on :class:`ParagraphFormat`, :class:`Font`,
:class:`FrameFormat`, and the style-chain merge code.  Extracted out of
``ldm_builder.py`` so the builder mixin stays focused on traversal
logic, and additions are localised to a single small module.
"""

# ---------------------------------------------------------------------------
# CT_OnOff toggles
# ---------------------------------------------------------------------------

#: ``<w:pPr>`` CT_OnOff children mapped to :class:`ParagraphFormat` attrs.
#:
#: Each entry is ``(child_local_name, attribute_name, default_when_missing)``.
#: Defaults match Word's behaviour: most flags are ``False`` when absent,
#: while ``widowControl``, ``snapToGrid``, ``autoSpaceDE/DN`` and
#: ``adjustRightInd`` default to ``True``.
PARA_ONOFF_FLAGS: tuple[tuple[str, str, bool], ...] = (
    ("keepNext", "keep_with_next", False),
    ("keepLines", "keep_together", False),
    ("widowControl", "widow_control", True),
    ("suppressAutoHyphens", "suppress_auto_hyphens", False),
    ("suppressLineNumbers", "suppress_line_numbers", False),
    ("snapToGrid", "snap_to_grid", True),
    ("autoSpaceDE", "add_space_between_far_east_and_alpha", True),
    ("autoSpaceDN", "add_space_between_far_east_and_digit", True),
    ("adjustRightInd", "auto_adjust_right_indent", True),
    ("pageBreakBefore", "page_break_before", False),
    ("contextualSpacing", "no_space_between_paragraphs_of_same_style", False),
)

#: ``<w:rPr>`` CT_OnOff children mapped to :class:`Font` attrs.  All default
#: to ``False`` (Word's behaviour when the element is absent).
RUN_ONOFF_FLAGS: tuple[tuple[str, str, bool], ...] = (
    ("b", "bold", False),
    ("bCs", "bold_bi", False),
    ("i", "italic", False),
    ("iCs", "italic_bi", False),
    ("strike", "strike_through", False),
    ("caps", "all_caps", False),
    ("smallCaps", "small_caps", False),
    ("vanish", "hidden", False),
    ("emboss", "emboss", False),
    ("imprint", "engrave", False),
    ("outline", "outline", False),
    ("shadow", "shadow", False),
    ("noProof", "no_proofing", False),
)

# ---------------------------------------------------------------------------
# Enum value mappings
# ---------------------------------------------------------------------------

#: Mapping from ``<w:textAlignment w:val>`` → BaselineAlignment int.
BASELINE_ALIGN_MAP: dict[str, int] = {
    "auto": 0, "top": 1, "center": 2, "baseline": 3, "bottom": 4,
}

# ---------------------------------------------------------------------------
# Frame (``<w:framePr/>``) attribute → :class:`FrameFormat` field
# ---------------------------------------------------------------------------

#: ``<w:framePr/>`` twip-valued attributes → :class:`FrameFormat` floats.
#: Field names match the ``FrameFormat`` API.
FRAME_TWIP_ATTRS: tuple[tuple[str, str], ...] = (
    ("w", "width"),
    ("h", "height"),
    ("x", "horizontal_position"),
    ("y", "vertical_position"),
    ("hSpace", "horizontal_distance_from_text"),
    ("vSpace", "vertical_distance_from_text"),
)

#: ``HeightRule`` enum: OOXML ``w:framePr/@hRule`` token → int.
FRAME_HEIGHT_RULE_MAP: dict[str, int] = {
    "atLeast": 0,
    "exact": 1,
    "auto": 2,
}

#: ``HorizontalAlignment`` enum: OOXML ``w:framePr/@xAlign`` token → int.
FRAME_HALIGN_MAP: dict[str, int] = {
    "left": 1,
    "center": 2,
    "right": 3,
    "inside": 4,
    "outside": 5,
}

#: ``VerticalAlignment`` enum: OOXML ``w:framePr/@yAlign`` token → int.
FRAME_VALIGN_MAP: dict[str, int] = {
    "inline": -1,
    "top": 1,
    "center": 2,
    "bottom": 3,
    "inside": 4,
    "outside": 5,
}

#: ``RelativeHorizontalPosition`` enum: OOXML ``w:framePr/@hAnchor`` → int.
#: OOXML only emits three values; the others originate from RTF / DOC.
FRAME_HANCHOR_MAP: dict[str, int] = {
    "margin": 0,
    "page": 1,
    "text": 3,
}

#: ``RelativeVerticalPosition`` enum: OOXML ``w:framePr/@vAnchor`` → int.
FRAME_VANCHOR_MAP: dict[str, int] = {
    "margin": 0,
    "page": 1,
    "text": 2,
}

#: ``WrapType`` enum: OOXML ``w:framePr/@wrap`` token → int.
FRAME_WRAP_MAP: dict[str, int] = {
    "notBeside": 1,
    "around": 2,
    "auto": 2,            # synonym in older OOXML
    "none": 3,
    "tight": 4,
    "through": 5,
}

#: ``<w:framePr/>`` int-enum attributes → ``(FrameFormat field, OOXML→int map)``.
#: ``w:anchorLock`` (boolean) is handled separately by the caller.
FRAME_ENUM_ATTRS: tuple[tuple[str, str, dict[str, int]], ...] = (
    ("hRule", "height_rule", FRAME_HEIGHT_RULE_MAP),
    ("xAlign", "horizontal_alignment", FRAME_HALIGN_MAP),
    ("yAlign", "vertical_alignment", FRAME_VALIGN_MAP),
    ("hAnchor", "relative_horizontal_position", FRAME_HANCHOR_MAP),
    ("vAnchor", "relative_vertical_position", FRAME_VANCHOR_MAP),
    ("wrap", "wrap_type", FRAME_WRAP_MAP),
)

#: ``DropCapPosition`` enum: OOXML token → int value.
DROP_CAP_POSITION_MAP: dict[str, int] = {
    "none": 0,
    "drop": 1,
    "margin": 2,
}

#: ``EmphasisMark`` enum: OOXML ``w:em w:val`` token → int value.
EMPHASIS_MARK_MAP: dict[str, int] = {
    "none": 0,
    "dot": 1,
    "comma": 2,
    "circle": 3,
    "underDot": 4,
}

#: ``TextEffect`` enum: OOXML ``w:effect w:val`` token → int value.
TEXT_EFFECT_MAP: dict[str, int] = {
    "none": 0,
    "lights": 1,
    "blinkBackground": 2,
    "sparkle": 3,
    "antsBlack": 4,
    "antsRed": 5,
    "shimmer": 6,
}

#: ``ListTrailingCharacter`` enum: OOXML ``w:suff/@val`` → int.
LIST_TRAILING_CHARACTER_MAP: dict[str, int] = {
    "tab": 0,
    "space": 1,
    "nothing": 2,
}

# ---------------------------------------------------------------------------
# Style-chain merge field lists
# ---------------------------------------------------------------------------

#: Fields propagated by ``LdmBuilderMixin._merge_font``.  The two
#: colour fields ignore an empty-sentinel override so a derived style
#: doesn't wipe an explicit base colour (handled in the merge code).
MERGE_FONT_FIELDS: tuple[str, ...] = (
    "name", "size", "bold", "bold_bi", "italic", "italic_bi",
    "underline", "color",
    "strike_through", "superscript", "subscript", "highlight_color",
    "all_caps", "small_caps", "hidden", "no_proofing",
    "style_name", "style_identifier",
    "shading",
    "emboss", "engrave", "outline", "shadow",
    "text_effect", "emphasis_mark",
    "kerning",
    "name_bi", "name_far_east", "name_ascii",
    "locale_id", "locale_id_bi", "locale_id_far_east",
)

#: Fields propagated by ``LdmBuilderMixin._merge_pf``.  ``tab_stops``
#: is handled separately so the clear-position merge logic survives.
MERGE_PF_FIELDS: tuple[str, ...] = (
    "style_name", "style_identifier", "alignment",
    "left_indent", "right_indent", "first_line_indent",
    "space_before", "space_after", "space_before_auto", "space_after_auto",
    "line_spacing", "line_spacing_rule",
    "keep_with_next", "page_break_before",
    "no_space_between_paragraphs_of_same_style",
    "outline_level", "is_heading", "is_list_item",
    "shading", "borders",
    "paragraph_break_font",
    "keep_together", "widow_control",
    "suppress_auto_hyphens", "suppress_line_numbers",
    "snap_to_grid",
    "add_space_between_far_east_and_alpha",
    "add_space_between_far_east_and_digit",
    "auto_adjust_right_indent",
    "baseline_alignment", "conditional_style",
    "frame_format", "lines_to_drop", "drop_cap_position",
)
