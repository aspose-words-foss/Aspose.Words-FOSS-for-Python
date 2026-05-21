"""Paragraph (``<w:p>``) and paragraph-properties (``<w:pPr>``) rendering."""


from typing import Mapping, Optional

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_reader.constants import PAGE_FIELD_SENTINEL
from aspose.words_foss.docx_writer.constants import (
    ALIGNMENT_VAL,
    TAB_ALIGNMENT_VAL,
    TAB_LEADER_VAL,
    pt_to_twips,
)
from aspose.words_foss.docx_writer.numbering_part import remap_num_id
from aspose.words_foss.docx_writer.runs import color_to_hex, render_rPr, render_run
from aspose.words_foss.docx_writer.xml_utils import el

_EMPTY_NUM_ID_MAP: Mapping[int, int] = {}


# Heading style id used in ``<w:pStyle>`` — Word's built-in convention is
# ``Heading1`` … ``Heading9`` (no space).
def _heading_style_id(level: int) -> str:
    return f"Heading{level}"


_LINE_RULE_TOKEN = {0: "atLeast", 1: "exact", 2: "auto"}


#: ``FrameFormat`` numeric (point) attributes → ``<w:framePr/>`` twip attrs.
_FRAME_TWIP_OUT: tuple[tuple[str, str], ...] = (
    ("width", "w"),
    ("height", "h"),
    ("horizontal_position", "x"),
    ("vertical_position", "y"),
    ("horizontal_distance_from_text", "hSpace"),
    ("vertical_distance_from_text", "vSpace"),
)

# ``HeightRule`` int → OOXML token.  0=AtLeast is default,
# omitted to keep the round-trip empty when not set.
_FRAME_HEIGHT_RULE_TOKEN: dict[int, str] = {1: "exact", 2: "auto"}
_FRAME_HALIGN_TOKEN: dict[int, str] = {
    1: "left", 2: "center", 3: "right", 4: "inside", 5: "outside",
}
_FRAME_VALIGN_TOKEN: dict[int, str] = {
    -1: "inline", 1: "top", 2: "center", 3: "bottom", 4: "inside", 5: "outside",
}
_FRAME_HANCHOR_TOKEN: dict[int, str] = {0: "margin", 1: "page", 3: "text"}
_FRAME_VANCHOR_TOKEN: dict[int, str] = {0: "margin", 1: "page", 2: "text"}
_FRAME_WRAP_TOKEN: dict[int, str] = {
    1: "notBeside", 2: "around", 3: "none", 4: "tight", 5: "through",
}
_FRAME_ENUM_OUT: tuple[tuple[str, str, dict[int, str]], ...] = (
    ("height_rule", "hRule", _FRAME_HEIGHT_RULE_TOKEN),
    ("horizontal_alignment", "xAlign", _FRAME_HALIGN_TOKEN),
    ("vertical_alignment", "yAlign", _FRAME_VALIGN_TOKEN),
    ("relative_horizontal_position", "hAnchor", _FRAME_HANCHOR_TOKEN),
    ("relative_vertical_position", "vAnchor", _FRAME_VANCHOR_TOKEN),
    ("wrap_type", "wrap", _FRAME_WRAP_TOKEN),
)
_DROP_CAP_POSITION_TOKEN: dict[int, str] = {
    0: "none",
    1: "drop",    
    2: "margin",
}


def _frame_pr_xml(
    frame: ldm.FrameFormat,
    *,
    lines_to_drop: int = 0,
    drop_cap_position: int = 0,
) -> str:
    """Render ``<w:framePr/>`` from a typed :class:`ldm.FrameFormat`.

    The two drop-cap attributes (``w:lines`` / ``w:dropCap``) live on
    ``ParagraphFormat``, so they are passed alongside the
    frame instead of being read from it.
    """
    attrs: dict[str, object] = {}
    for src, target in _FRAME_TWIP_OUT:
        val = getattr(frame, src)
        if val:
            attrs[f"w:{target}"] = pt_to_twips(val)
    for src, target, token_map in _FRAME_ENUM_OUT:
        val = getattr(frame, src)
        token = token_map.get(val)
        if token is not None:
            attrs[f"w:{target}"] = token
    if lines_to_drop:
        attrs["w:lines"] = lines_to_drop
    if drop_cap_position:
        attrs["w:dropCap"] = _DROP_CAP_POSITION_TOKEN.get(drop_cap_position, "none")
    if frame.lock_anchor:
        attrs["w:anchorLock"] = "1"
    return el("w:framePr", attrs)


#: Paragraph CT_OnOff toggles emitted *before* ``<w:framePr>`` per the
#: CT_PPrBase schema sequence (keepNext → keepLines → pageBreakBefore →
#: framePr → widowControl → ...).
_PRE_FRAME_ONOFF_TAGS: tuple[tuple[str, str], ...] = (
    ("w:keepNext", "keep_with_next"),
    ("w:keepLines", "keep_together"),
    ("w:pageBreakBefore", "page_break_before"),
)

#: Toggles emitted between ``<w:framePr>`` (position 5) and ``<w:shd>``
#: (position 10) in the CT_PPrBase sequence.  ``<w:widowControl>`` (6)
#: precedes ``<w:numPr>``; ``<w:suppressLineNumbers>`` (8) follows it.
#: ``render_pPr`` later splits ``<w:widowControl>`` into the pre-numPr
#: block so the final order matches: framePr → widowControl → numPr →
#: suppressLineNumbers → pBdr → shd.
_PRE_SHADING_ONOFF_TAGS: tuple[tuple[str, str], ...] = (
    ("w:widowControl", "widow_control"),
    ("w:suppressLineNumbers", "suppress_line_numbers"),
)

#: Toggles emitted between <w:tabs> (pos 11) and <w:spacing> (pos 22)
#: per CT_PPrBase.  <w:contextualSpacing> belongs after <w:ind> (pos 24)
#: — emitting it here would violate the schema and break Word.
_POST_TABS_ONOFF_TAGS: tuple[tuple[str, str], ...] = (
    ("w:suppressAutoHyphens", "suppress_auto_hyphens"),
    ("w:autoSpaceDE", "add_space_between_far_east_and_alpha"),
    ("w:autoSpaceDN", "add_space_between_far_east_and_digit"),
    ("w:adjustRightInd", "auto_adjust_right_indent"),
    ("w:snapToGrid", "snap_to_grid"),
)


#: BaselineAlignment int → OOXML ``<w:textAlignment w:val>`` token.
_BASELINE_ALIGN_VAL: dict[int, str] = {
    0: "auto", 1: "top", 2: "center", 3: "baseline", 4: "bottom",
}


def _onoff_xml(tag: str, value: bool) -> str:
    """Emit a CT_OnOff element matching Word's ``tag present = on``
    convention: ``<w:tag/>`` for true, ``<w:tag w:val="0"/>`` for false.
    """
    return el(tag) if value else el(tag, {"w:val": "0"})


def _normalize_tab_list(raw: object) -> list[tuple[float, int, int, bool]]:
    """Return a uniform list of ``(position, alignment, leader, is_clear)``
    from either a TabStopCollection, a list of TabStop objects, or a list
    of ``(position, alignment, leader)`` tuples (DOC reader legacy)."""
    if not raw:
        return []
    items = raw.tab_stops if hasattr(raw, 'tab_stops') else raw
    if not items:
        return []
    first = items[0]
    if isinstance(first, tuple):
        return [(p, a, l, False) for p, a, l in items]
    return [(t.position, t.alignment, t.leader, getattr(t, 'is_clear', False)) for t in items]


def _spacing_attrs(pf: ldm.ParagraphFormat, base: ldm.ParagraphFormat) -> dict[str, object]:
    """Build the ``<w:spacing>`` attribute dict.

    Emits only the attributes whose value differs from ``base`` (the
    paragraph's resolved style chain).  Zero values DO get emitted as
    explicit overrides when the style chain would otherwise inherit a
    non-zero value — that's the only way the LDM's ``space_after=0`` on a
    paragraph using a style with ``space_after=8`` survives a round-trip.
    """
    attrs: dict[str, object] = {}
    if pf.space_before != base.space_before:
        attrs["w:before"] = pt_to_twips(pf.space_before)
    if pf.space_after != base.space_after:
        attrs["w:after"] = pt_to_twips(pf.space_after)
    if pf.space_before_auto != base.space_before_auto:
        attrs["w:beforeAutospacing"] = "1" if pf.space_before_auto else "0"
    if pf.space_after_auto != base.space_after_auto:
        attrs["w:afterAutospacing"] = "1" if pf.space_after_auto else "0"
    if pf.line_spacing != base.line_spacing or pf.line_spacing_rule != base.line_spacing_rule:
        attrs["w:line"] = pt_to_twips(pf.line_spacing)
        attrs["w:lineRule"] = _LINE_RULE_TOKEN.get(pf.line_spacing_rule, "auto")
    return attrs


def _ind_attrs(pf: ldm.ParagraphFormat, base: ldm.ParagraphFormat) -> dict[str, object]:
    """Build the ``<w:ind>`` attribute dict — handles hanging indents.

    Like :func:`_spacing_attrs`, only emits attributes that differ from the
    style-chain base so explicit zero overrides survive the round-trip.
    """
    attrs: dict[str, object] = {}
    if pf.left_indent != base.left_indent:
        attrs["w:left"] = pt_to_twips(pf.left_indent)
    if pf.right_indent != base.right_indent:
        attrs["w:right"] = pt_to_twips(pf.right_indent)
    if pf.first_line_indent != base.first_line_indent:
        if pf.first_line_indent >= 0:
            attrs["w:firstLine"] = pt_to_twips(pf.first_line_indent)
        else:
            attrs["w:hanging"] = pt_to_twips(-pf.first_line_indent)
    return attrs


def _resolve_style_id(
    pf: ldm.ParagraphFormat,
    style_id_map: Optional[Mapping[str, str]] = None,
) -> Optional[str]:
    """Pick the style id to emit in ``<w:pStyle>``.

    Headings always win over the raw style name so ``Heading 1`` →
    ``Heading1`` regardless of how the reader stored it.  Other style
    names route through ``style_id_map`` (built by ``styles_part``) so
    paragraph references stay in lockstep with the styleIds emitted
    into ``styles.xml`` even when names collide after ``replace(" ", "")``.
    """
    if pf.is_heading and 0 <= pf.outline_level < 9:
        if pf.style_name and style_id_map is not None and pf.style_name in style_id_map:
            return style_id_map[pf.style_name]
        return _heading_style_id(pf.outline_level + 1)
    if pf.style_name:
        # Even ``"Normal"`` is emitted when explicitly carried — the
        # reader only assigns ``style_name="Normal"`` when the paragraph
        # actually had a ``<w:pStyle w:val="Normal"/>`` reference, so
        # suppressing it on write would erase that distinction on the
        # next read.
        from aspose.words_foss.docx_writer.styles_part import _sanitize_style_id
        fallback = _sanitize_style_id(pf.style_name.replace(" ", ""))
        if style_id_map is not None:
            return style_id_map.get(pf.style_name, fallback)
        return fallback
    return None


_DEFAULT_PF: Optional[ldm.ParagraphFormat] = None


def _default_pf() -> ldm.ParagraphFormat:
    """Lazy-cached zero-filled ``ParagraphFormat`` used as the no-base default."""
    global _DEFAULT_PF
    if _DEFAULT_PF is None:
        _DEFAULT_PF = ldm.ParagraphFormat()
    return _DEFAULT_PF


def pf_to_pPr_children(
    pf: ldm.ParagraphFormat,
    *,
    base: Optional[ldm.ParagraphFormat] = None,
) -> list[str]:
    """Return the schema-ordered ``<w:pPr>`` children for a ``ParagraphFormat``.

    Shared by paragraph rendering and by ``styles_part._custom_style``.  Each
    field is emitted only if it differs from the ``base`` paragraph format
    (the style chain's resolved values for paragraphs, or the basedOn style's
    resolved values when emitting a custom style).  ``base=None`` falls back
    to a zero-filled :class:`ParagraphFormat` so callers without style-chain
    context still get a correct (though redundant) emission.

    Excludes ``pStyle``, ``numPr`` and ``sectPr`` — those are paragraph-only
    concerns.  ``<w:outlineLvl>`` is included via the diff path above so
    both heading styles and explicit non-heading overrides round-trip.
    """
    if base is None:
        base = _default_pf()

    children: list[str] = []

    # CT_OnOff diff against the style base — disabling a toggle emits
    # <w:tag w:val="0"/>, not absence.  Pre-frame toggles precede
    # <w:framePr>, post-frame ones follow it per CT_PPrBase order.
    for tag, attr in _PRE_FRAME_ONOFF_TAGS:
        new_val = getattr(pf, attr)
        if new_val != getattr(base, attr):
            children.append(_onoff_xml(tag, new_val))

    # Floating frame — emit when the LDM has a populated FrameFormat
    # OR the paragraph carries a drop-cap (``lines_to_drop`` /
    # ``drop_cap_position``).  CT_PPrBase places ``<w:framePr>`` at
    # position 5, between ``<w:pageBreakBefore>`` and ``<w:widowControl>``.
    frame_changed = (
        pf.frame != base.frame
        or pf.lines_to_drop != base.lines_to_drop
        or pf.drop_cap_position != base.drop_cap_position
    )
    has_frame_data = (
        pf.frame is not None
        or pf.lines_to_drop
        or pf.drop_cap_position
    )
    if frame_changed and has_frame_data:
        frame_for_xml = pf.frame or ldm.FrameFormat()
        children.append(
            _frame_pr_xml(
                frame_for_xml,
                lines_to_drop=pf.lines_to_drop,
                drop_cap_position=pf.drop_cap_position,
            )
        )

    for tag, attr in _PRE_SHADING_ONOFF_TAGS:
        new_val = getattr(pf, attr)
        if new_val != getattr(base, attr):
            children.append(_onoff_xml(tag, new_val))

    shading_hex = color_to_hex(pf.shading.background_pattern_color)
    base_shading_hex = color_to_hex(base.shading.background_pattern_color)
    if shading_hex != base_shading_hex:
        if shading_hex:
            children.append(
                el(
                    "w:shd",
                    {"w:val": "clear", "w:color": "auto", "w:fill": shading_hex},
                )
            )
        else:
            children.append(el("w:shd", {"w:val": "nil"}))

    # Tab stops — emit ``<w:tabs>`` with delta against the base style.
    # For tabs present here but not in base (or with different properties): emit normally.
    # For base tabs NOT present here: emit ``w:val="clear"`` so the reader
    # doesn't re-inherit them on round-trip.
    _cur_tabs = _normalize_tab_list(pf.tab_stops)
    _base_tabs = _normalize_tab_list(base.tab_stops)
    if _cur_tabs or _base_tabs:
        base_tab_set: dict[int, tuple[int, int]] = {}
        for t_pos, t_aln, t_ldr, _ in _base_tabs:
            base_tab_set[pt_to_twips(t_pos)] = (t_aln, t_ldr)
        cur_tab_set: dict[int, tuple[int, int]] = {}
        for t_pos, t_aln, t_ldr, t_clr in _cur_tabs:
            if not t_clr:
                cur_tab_set[pt_to_twips(t_pos)] = (t_aln, t_ldr)
        tab_entries: list[tuple[int, str]] = []
        for t_pos, t_aln, t_ldr, t_clr in _cur_tabs:
            key = pt_to_twips(t_pos)
            base_entry = base_tab_set.get(key)
            if t_clr or base_entry is None or base_entry != (t_aln, t_ldr):
                attrs: dict[str, object] = {
                    "w:val": "clear" if t_clr
                    else TAB_ALIGNMENT_VAL.get(t_aln, "left"),
                    "w:pos": key,
                }
                leader_val = TAB_LEADER_VAL.get(t_ldr, "none")
                if leader_val != "none":
                    attrs["w:leader"] = leader_val
                tab_entries.append((key, el("w:tab", attrs)))
        for pos_twips in base_tab_set:
            if pos_twips not in cur_tab_set:
                tab_entries.append((pos_twips, el("w:tab", {"w:val": "clear", "w:pos": pos_twips})))
        if tab_entries:
            tab_entries.sort(key=lambda e: e[0])
            children.append(el("w:tabs", None, [x for _, x in tab_entries]))

    for tag, attr in _POST_TABS_ONOFF_TAGS:
        new_val = getattr(pf, attr)
        if new_val != getattr(base, attr):
            children.append(_onoff_xml(tag, new_val))

    spacing = _spacing_attrs(pf, base)
    if spacing:
        children.append(el("w:spacing", spacing))

    ind = _ind_attrs(pf, base)
    if ind:
        children.append(el("w:ind", ind))

    # ``<w:contextualSpacing>`` (CT_PPrBase position 24) sits after
    # ``<w:ind>`` per the schema sequence.  Word rejects documents that
    # place it earlier (e.g. inside the early OnOff block) when the
    # paragraph style has any other CT_PPrBase children after the
    # toggle — LibreOffice and pandoc happen to accept the wrong order.
    if (
        pf.no_space_between_paragraphs_of_same_style
        != base.no_space_between_paragraphs_of_same_style
    ):
        children.append(
            _onoff_xml(
                "w:contextualSpacing",
                pf.no_space_between_paragraphs_of_same_style,
            )
        )

    if pf.alignment != base.alignment:
        val = ALIGNMENT_VAL.get(pf.alignment, "left")
        children.append(el("w:jc", {"w:val": val}))

    if pf.baseline_alignment != base.baseline_alignment:
        children.append(
            el("w:textAlignment",
               {"w:val": _BASELINE_ALIGN_VAL.get(pf.baseline_alignment, "auto")})
        )

    # ``outline_level`` defaults to 9 (== "no outline"), so a diff against
    # the basedOn's resolved value covers all three useful cases:
    #   * heading style sets 0..8 against ``Normal``'s 9 → emit
    #   * "TOC Heading" sets 9 explicitly against ``Heading1``'s 0 → emit
    #   * paragraph inherits the style's outline → no emit
    if pf.outline_level != base.outline_level:
        children.append(el("w:outlineLvl", {"w:val": pf.outline_level}))

    # ``<w:cnfStyle>`` is CT_PPrBase position 33 — it must come *after*
    # ``<w:outlineLvl>``.  Emitting before ``<w:jc>`` produced a
    # schema-invalid order that Word flags as document corruption.
    if pf.conditional_style != base.conditional_style and pf.conditional_style:
        children.append(el("w:cnfStyle", {"w:val": pf.conditional_style}))

    return children


def render_pPr(
    para: ldm.Paragraph,
    *,
    list_format: Optional[ldm.ListFormat] = None,
    embedded_sectPr: Optional[str] = None,
    num_id_map: Mapping[int, int] = _EMPTY_NUM_ID_MAP,
    style_pf_map: Optional[Mapping[str, ldm.ParagraphFormat]] = None,
    style_id_map: Optional[Mapping[str, str]] = None,
) -> str:
    """Build a ``<w:pPr>`` for the paragraph.  Empty string if no props.

    Children are emitted in the order required by the OOXML schema
    (``CT_PPr``):
    pStyle → keepNext → keepLines → pageBreakBefore → widowControl →
    numPr → pBdr → shd → spacing → ind → jc → outlineLvl → sectPr.

    ``embedded_sectPr`` is an already-rendered ``<w:sectPr>`` element
    that should be placed inside this paragraph's ``pPr`` to mark a
    section break — see :mod:`document_part` for why intermediate
    sections embed their setup here instead of trailing it after the
    paragraph.

    ``style_pf_map`` maps canonical style ids (lowercase, no spaces) to
    each style's resolved :class:`ParagraphFormat`.  When provided, the
    paragraph's own format is emitted as a *diff* against its style's
    resolved values so explicit zero overrides (``space_after=0``
    against a style with ``space_after=8``) survive the round-trip.
    """
    pf = para.paragraph_format
    children: list[str] = []

    style_id = _resolve_style_id(pf, style_id_map)
    if style_id:
        children.append(el("w:pStyle", {"w:val": style_id}))

    base_pf = _resolve_base_pf(pf, style_id, style_pf_map)
    pf_children = pf_to_pPr_children(pf, base=base_pf)

    # CT_PPr schema places ``<w:numPr>`` at position 7, so everything at
    # positions 2–6 (keepNext, keepLines, pageBreakBefore, framePr,
    # widowControl) must precede it.  Splitting on the numPr-anchor tags
    # keeps the diff-based emission self-contained while still honouring
    # the schema order.
    pre_num_tags = (
        "<w:keepNext",
        "<w:keepLines",
        "<w:pageBreakBefore",
        "<w:framePr",
        "<w:widowControl",
    )
    pre_num: list[str] = []
    rest: list[str] = []
    for child in pf_children:
        (pre_num if child.startswith(pre_num_tags) else rest).append(child)
    children.extend(pre_num)

    if list_format and list_format.is_list_item:
        num_id = remap_num_id(list_format.list_id, num_id_map)
        num_pr = [
            el("w:ilvl", {"w:val": list_format.list_level_number}),
            el("w:numId", {"w:val": num_id}),
        ]
        children.append(el("w:numPr", None, num_pr))

    children.extend(rest)

    # ``pPr/rPr`` (paragraph-mark formatting) lives between the rest of
    # the format children and the trailing ``sectPr`` per CT_PPr.  Diff
    # against the basedOn pf's mark font so a child paragraph that
    # inherits the parent style's italic-bullet doesn't grow a redundant
    # rPr on every round-trip.
    mark_rPr = _render_mark_rPr(pf.paragraph_mark_font, base_pf.paragraph_mark_font)
    if mark_rPr:
        children.append(mark_rPr)

    if embedded_sectPr:
        children.append(embedded_sectPr)

    if not children:
        return ""
    return el("w:pPr", None, children)


def _render_mark_rPr(mark: Optional[ldm.Font], base_mark: Optional[ldm.Font]) -> str:
    """Emit ``<w:rPr>`` inside ``<w:pPr>`` for the paragraph-mark font.

    Returns the empty string when ``mark`` matches the basedOn baseline
    or both sides are ``None``.  Always emits an empty ``<w:rPr/>`` when
    the LDM explicitly carries a default-valued mark font — the reader
    populates ``paragraph_mark_font`` from the *element's presence*, not
    its content.
    """
    if mark is None and base_mark is None:
        return ""
    if mark is None:
        # Parent had a mark font; child clears it.  No widely-supported
        # OOXML "delete this rPr" toggle, so we drop the field rather
        # than emit a synthetic reset that would also override every
        # individual property the parent set.
        return ""
    rPr = render_rPr(mark, base=base_mark)
    return rPr or el("w:rPr")


def _resolve_base_pf(
    pf: ldm.ParagraphFormat,
    style_id: Optional[str],
    style_pf_map: Optional[Mapping[str, ldm.ParagraphFormat]],
) -> ldm.ParagraphFormat:
    """Look up the resolved ``ParagraphFormat`` of the paragraph's style.

    Falls back to ``Normal``'s pf, then to a zero-filled default.  Returning
    the right base is what makes diff-based emission work: a paragraph and
    its style share resolved values (no override) → emit nothing; the
    paragraph differs → emit the differing attributes only.

    ``style_pf_map`` is keyed by ``style.name.replace(" ", "").lower()``,
    so we have to look up under the LDM-side ``pf.style_name`` (the
    Word display name) rather than the on-disk sanitised ``style_id``
    — for names like ``"*TextBullet1"`` the sanitised id is
    ``"_TextBullet1"`` and would miss the ``"*textbullet1"`` key.
    """
    if style_pf_map is None:
        return _default_pf()
    if pf.style_name:
        canonical = pf.style_name.replace(" ", "").lower()
        base = style_pf_map.get(canonical)
        if base is not None:
            return base
    if style_id:
        canonical = style_id.replace(" ", "").lower()
        base = style_pf_map.get(canonical)
        if base is not None:
            return base
    base = style_pf_map.get("normal")
    if base is not None:
        return base
    return _default_pf()


def _identify_page_sentinel_extras(para: ldm.Paragraph) -> set[int]:
    """Return ``id()`` of field markers suppressed by a PAGE sentinel run.

    The sentinel re-emits the full begin/instr/separate/end chain via
    :func:`_render_page_field_run`, so the matching field triple must
    be skipped on emit or Word sees the chain twice.  We walk
    :attr:`Paragraph._children` in source order, push each open field
    onto a stack, and flag a triple when the run between separate/end
    is the PAGE sentinel.
    """
    skip_ids: set[int] = set()
    stack: list[list] = []  # [start_node, sep_node_or_None, is_page]

    for node in para._children:
        if isinstance(node, ldm.FieldStart):
            stack.append([node, None, False])
        elif isinstance(node, ldm.FieldSeparator):
            if stack:
                stack[-1][1] = node
        elif isinstance(node, ldm.FieldEnd):
            if stack:
                start_node, sep_node, is_page = stack.pop()
                if is_page:
                    skip_ids.add(id(start_node))
                    if sep_node is not None:
                        skip_ids.add(id(sep_node))
                    skip_ids.add(id(node))
        elif isinstance(node, ldm.Run):
            if node.text == PAGE_FIELD_SENTINEL and stack:
                stack[-1][2] = True

    return skip_ids


def render_paragraph(
    para: ldm.Paragraph,
    rels: dict,
    *,
    embedded_sectPr: Optional[str] = None,
    num_id_map: Mapping[int, int] = _EMPTY_NUM_ID_MAP,
    image_state: Optional["ImageRenderState"] = None,
    bookmark_state: Optional["BookmarkState"] = None,
    style_pf_map: Optional[Mapping[str, ldm.ParagraphFormat]] = None,
    style_id_map: Optional[Mapping[str, str]] = None,
    style_font_map: Optional[Mapping[str, ldm.Font]] = None,
) -> str:
    """Render a single LDM paragraph to ``<w:p>...</w:p>``.

    ``embedded_sectPr`` is forwarded to :func:`render_pPr` so callers
    that need to mark a section break at this paragraph (every section
    except the last) can do so without reaching into pPr internals.

    ``num_id_map`` is the document-level LDM ``list_id`` → writer
    ``numId`` mapping (see :func:`numbering_part.build_num_id_map`).
    The default empty map preserves single-paragraph rendering for
    unit tests; the writer's main path always passes a real map so
    list paragraphs and ``numbering.xml`` agree on ids.

    ``image_state`` accumulates inline shape relationships and image
    bytes; passing ``None`` disables image emission (used by isolated
    unit tests that don't carry shape data).  ``bookmark_state`` does
    the same for ``BookmarkStart`` markers in ``inline_extras``.
    """
    children: list[str] = []
    pPr = render_pPr(
        para,
        list_format=para.list_format,
        embedded_sectPr=embedded_sectPr,
        num_id_map=num_id_map,
        style_pf_map=style_pf_map,
        style_id_map=style_id_map,
    )
    if pPr:
        children.append(pPr)

    page_sentinel_skip_ids = _identify_page_sentinel_extras(para)
    # Stack of separator-seen flags per open field.  False = still in
    # instr-text region; FieldSeparator flips top to True, FieldEnd
    # pops.  Using a stack (not a single depth int) lets fields without
    # a separator correctly release instr mode on FieldEnd.
    field_separator_seen: list[bool] = []

    def field_code_depth() -> int:
        return sum(1 for seen in field_separator_seen if not seen)

    def render_extra(extra) -> Optional[str]:
        if isinstance(extra, ldm.BookmarkStart):
            if bookmark_state is not None and extra.name:
                bm_id = bookmark_state.open(extra.name)
                return el("w:bookmarkStart", {"w:id": bm_id, "w:name": extra.name})
        elif isinstance(extra, ldm.BookmarkEnd):
            if bookmark_state is not None and extra.name:
                bm_id = bookmark_state.close(extra.name)
                if bm_id is not None:
                    return el("w:bookmarkEnd", {"w:id": bm_id})
        elif isinstance(extra, ldm.ShapeNode):
            if image_state is not None:
                return image_state.render_inline_shape(
                    extra,
                    rels,
                    num_id_map=num_id_map,
                    bookmark_state=bookmark_state,
                    style_pf_map=style_pf_map,
                    style_id_map=style_id_map,
                    style_font_map=style_font_map,
                )
        elif isinstance(extra, ldm.FieldStart):
            if id(extra) in page_sentinel_skip_ids:
                return None
            field_separator_seen.append(False)
            return el("w:r", None, el("w:fldChar", {"w:fldCharType": "begin"}))
        elif isinstance(extra, ldm.FieldSeparator):
            if id(extra) in page_sentinel_skip_ids:
                return None
            if field_separator_seen:
                field_separator_seen[-1] = True
            return el("w:r", None, el("w:fldChar", {"w:fldCharType": "separate"}))
        elif isinstance(extra, ldm.FieldEnd):
            if id(extra) in page_sentinel_skip_ids:
                return None
            if field_separator_seen:
                field_separator_seen.pop()
            return el("w:r", None, el("w:fldChar", {"w:fldCharType": "end"}))
        return None

    base_font: Optional[ldm.Font] = None
    if style_font_map is not None and para.paragraph_format.style_name:
        canonical = para.paragraph_format.style_name.replace(" ", "").lower()
        base_font = style_font_map.get(canonical)

    # Walk ``children`` in source order — runs, bookmark / field
    # markers and inline shapes already sit in the right places, so
    # interleaving is implicit and no positional index is needed.
    for node in para._children:
        if isinstance(node, ldm.Run):
            rendered_run = render_run(
                node, rels, base_font=base_font, instr=field_code_depth() > 0
            )
            if rendered_run:
                children.append(rendered_run)
        else:
            rendered_extra = render_extra(node)
            if rendered_extra:
                children.append(rendered_extra)

    return el("w:p", None, children)


# Forward references — actual classes live in :mod:`drawing` and
# :mod:`bookmarks` to keep this module focused on paragraph layout.
from aspose.words_foss.docx_writer.drawing import ImageRenderState  # noqa: E402
from aspose.words_foss.docx_writer.bookmarks import BookmarkState  # noqa: E402
