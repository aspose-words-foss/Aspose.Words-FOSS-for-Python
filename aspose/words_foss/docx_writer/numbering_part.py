"""Generates ``word/numbering.xml`` from LDM lists.

The reader stores numbering as a flat list of ``DocList`` entries,
each with up to nine ``ListLevel`` rows.  Word requires both an
``abstractNum`` (the definition) and a ``num`` (the instance pointing
at it).  Their IDs share the same value here for simplicity — Word
allows that.
"""


from typing import Iterable, Mapping

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_writer.constants import W_URI, pt_to_twips
from aspose.words_foss.docx_writer.xml_utils import XML_DECL, el
from aspose.words_foss.docx_writer.runs import render_rPr
from aspose.words_foss.model.list_limits import MAX_LIST_LEVELS

_NUMBER_FORMAT_VAL = {
    0: "decimal",
    1: "upperRoman",
    2: "lowerRoman",
    3: "upperLetter",
    4: "lowerLetter",
    5: "ordinal",
    23: "bullet",
    255: "none",
}

# Inverse of the reader's ``_ALIGNMENT_MAP`` for the four values the
# numbering schema accepts (``CT_Lvl/lvlJc``).
_ALIGNMENT_TOKEN = {0: "left", 1: "center", 2: "right", 3: "both"}

#: ``ListTrailingCharacter`` int → OOXML ``w:suff/@val`` token.
#: 0=Tab is the OOXML default, so it's emitted only for the two
#: non-default values.
_LIST_TRAILING_CHARACTER_TOKEN: dict[int, str] = {1: "space", 2: "nothing"}


# An immutable empty mapping is used as the safe default when callers
# don't pass a per-document map — preserves test behaviour where a
# paragraph is rendered in isolation.
_EMPTY_NUM_ID_MAP: Mapping[int, int] = {}


def build_num_id_map(doc_lists: Iterable[ldm.DocList]) -> dict[int, int]:
    """Build a collision-free LDM ``list_id`` → writer ``numId`` map.

    Word reserves ``numId=0`` as the "no list" sentinel, and the writer
    must guarantee every emitted ``<w:abstractNum>`` / ``<w:num>`` has a
    distinct id.  The naive ``0 → 1`` rewrite collides with a document
    that already carries a ``DocList(list_id=1)``; instead we just
    renumber every distinct source id sequentially from 1.

    Two ``DocList`` entries that share the same source id share the
    same writer numId — the LDM treats them as the same logical list.
    """
    out: dict[int, int] = {}
    next_id = 1
    for dl in doc_lists:
        if dl.list_id in out:
            continue
        out[dl.list_id] = next_id
        next_id += 1
    return out


def remap_num_id(
    list_id: int,
    num_id_map: Mapping[int, int] = _EMPTY_NUM_ID_MAP,
) -> int:
    """Return the on-disk numId for an LDM ``list_id``.

    With a map built by :func:`build_num_id_map`, the result is the
    pre-allocated, collision-free id.  Without one, any ``0`` is
    bumped to ``1`` (preserving the previous single-list behaviour);
    other ids pass through unchanged.  The fallback path is used by
    isolated unit tests; the writer always passes a real map.
    """
    if list_id in num_id_map:
        return num_id_map[list_id]
    return 1 if list_id == 0 else list_id


def _level_xml(level_idx: int, level: ldm.ListLevel) -> str:

    fmt_val = _NUMBER_FORMAT_VAL.get(level.number_style, "decimal")
    # Always emit ``<w:lvlText w:val>`` from the LDM as-is, even when
    # empty \u2014 the reader stores exactly what it found, so folding "" into
    # a synthetic default would inflate ``ll.number_format`` on the next
    # read.
    text = level.number_format
    # ECMA-376 CT_Lvl child order:
    # start \u2192 numFmt \u2192 lvlRestart \u2192 pStyle \u2192 isLgl \u2192 suff \u2192 lvlText \u2192
    # lvlPicBulletId \u2192 legacy \u2192 lvlJc \u2192 pPr \u2192 rPr
    children = [
        el("w:start", {"w:val": level.start_at}),
        el("w:numFmt", {"w:val": fmt_val}),
    ]
    if level.restart_after_level != -1:
        children.append(el("w:lvlRestart", {"w:val": level.restart_after_level}))
    if level.trailing_character != 0:
        children.append(
            el("w:suff",
               {"w:val": _LIST_TRAILING_CHARACTER_TOKEN.get(level.trailing_character, "tab")})
        )
    children.extend([
        el("w:lvlText", {"w:val": text}),
        el("w:lvlJc", {"w:val": _ALIGNMENT_TOKEN.get(level.alignment, "left")}),
    ])
    ind_attrs = _ind_attrs_for_level(level)
    if ind_attrs:
        children.append(el("w:pPr", None, el("w:ind", ind_attrs)))
    # ``<w:lvl><w:rPr>`` carries the bullet glyph's font / italic / size.
    # Without it Word falls back to the paragraph mark's font, which
    # changes the bullet's appearance and (for Symbol-fonted bullets)
    # its glyph entirely.
    if level.font is not None:
        rPr = render_rPr(level.font) or el("w:rPr")
        children.append(rPr)
    return el("w:lvl", {"w:ilvl": level_idx}, children)


def _lvl_override_xml(override: "ldm.ListLevelOverride") -> str:
    """Render one ``<w:lvlOverride>`` from an LDM
    :class:`ListLevelOverride`.

    Word's schema is empty-element friendly: an override carrying
    neither a ``<w:startOverride>`` nor a nested ``<w:lvl>`` is still a
    valid element, so we round-trip the placeholder form
    ``<w:lvlOverride w:ilvl="N"/>`` byte-for-byte.
    """
    children: list[str] = []
    if override.start_at is not None:
        children.append(el("w:startOverride", {"w:val": override.start_at}))
    if override.list_level is not None:
        children.append(_level_xml(override.ilvl, override.list_level))
    return el("w:lvlOverride", {"w:ilvl": override.ilvl}, children)


def _ind_attrs_for_level(level: ldm.ListLevel) -> dict[str, object]:
    """Build the ``<w:ind>`` attributes that round-trip ``ListLevel``.

    The reader supports three branches: ``left+hanging``,
    ``left+firstLine``, and ``left`` only.  We mirror them so the
    paragraph reload sees the same numbers it carried before the trip.

    The previous implementation clamped ``hanging`` to ``>= 18pt``, which
    silently moved the bullet on reload whenever the source had a smaller
    hanging — gone now.  When ``text_position == 0`` we emit only
    ``w:left`` so the reader's "left only" branch fires and stores the
    value into ``number_position`` rather than ``text_position``.
    """
    tp, np = level.text_position, level.number_position
    if not tp and not np:
        return {}
    if not tp:
        return {"w:left": pt_to_twips(np)}
    hanging = max(tp - np, 0.0)
    return {"w:left": pt_to_twips(tp), "w:hanging": pt_to_twips(hanging)}


def _abstract_num(doc_list: ldm.DocList, num_id_map: Mapping[int, int]) -> str:
    levels = (doc_list.list_levels or [ldm.ListLevel(number_style=23)])[:MAX_LIST_LEVELS]
    # ECMA-376 forbids more than one ``<w:lvl>`` under a
    # ``<w:multiLevelType w:val="singleLevel"/>`` abstractNum — Aspose
    # rejects the extras with "Import of element 'lvlText' is not
    # supported".  Auto-upgrade to ``hybridMultilevel`` whenever the
    # LDM carries more than one level so the schema and the lvl count
    # agree even when ``is_multi_level`` got mis-set.
    is_multi = doc_list.is_multi_level or len(levels) > 1
    multi_token = "hybridMultilevel" if is_multi else "singleLevel"
    children = [el("w:multiLevelType", {"w:val": multi_token})]
    children.extend(_level_xml(i, lvl) for i, lvl in enumerate(levels))
    return el(
        "w:abstractNum",
        {"w:abstractNumId": remap_num_id(doc_list.list_id, num_id_map)},
        children,
    )


def _num(doc_list: ldm.DocList, num_id_map: Mapping[int, int]) -> str:
    num_id = remap_num_id(doc_list.list_id, num_id_map)
    children: list[str] = [el("w:abstractNumId", {"w:val": num_id})]
    in_range = [ov for ov in doc_list.overrides if 0 <= ov.ilvl < MAX_LIST_LEVELS]
    for ov in in_range[:MAX_LIST_LEVELS]:
        children.append(_lvl_override_xml(ov))
    return el("w:num", {"w:numId": num_id}, children)


def render_numbering_xml(doc: ldm.Document) -> str:
    """Build ``word/numbering.xml`` payload, or empty string if no lists.

    The collision-free numId map is built once per document and used
    for both the ``abstractNum`` and ``num`` rows so the on-disk ids
    stay in sync with the paragraph-side references emitted via
    :mod:`paragraphs`.
    """
    if not doc.lists:
        return ""
    num_id_map = build_num_id_map(doc.lists)
    # Deduplicate by source list_id — two LDM ``DocList`` entries with
    # the same id refer to the same logical list and would otherwise
    # produce duplicate <w:abstractNum>/<w:num> rows under one numId.
    seen: set[int] = set()
    unique_lists: list[ldm.DocList] = []
    for dl in doc.lists:
        if dl.list_id in seen:
            continue
        seen.add(dl.list_id)
        unique_lists.append(dl)

    children: list[str] = []
    children.extend(_abstract_num(dl, num_id_map) for dl in unique_lists)
    children.extend(_num(dl, num_id_map) for dl in unique_lists)
    root = el("w:numbering", {"xmlns:w": W_URI}, children)
    return XML_DECL + root
