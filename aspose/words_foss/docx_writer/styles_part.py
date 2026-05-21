"""Generates ``word/styles.xml``.

Emits a minimal but Word-compatible styles part: doc defaults plus the
nine built-in heading styles so paragraphs that reference
``HeadingN`` resolve to the expected outline level.  Custom styles
captured by the LDM are appended after the built-ins.
"""


import re
from typing import Optional

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.docx_reader.utils import _canonicalize_style_name
from aspose.words_foss.docx_writer.constants import W_URI, pt_to_half_pt, pt_to_twips
from aspose.words_foss.docx_writer.paragraphs import pf_to_pPr_children
from aspose.words_foss.docx_writer.runs import color_to_hex, render_rPr
from aspose.words_foss.docx_writer.blank_template import latent_styles
from aspose.words_foss.docx_writer.tables import _table_borders
from aspose.words_foss.docx_writer.xml_utils import XML_DECL, el
from aspose.words_foss.model.style_identifiers import IDENTIFIER_TO_STYLE_ID
from aspose.words_foss.docx_writer.paragraphs import _resolve_style_id

# Heading-name regex (reader-side).  Matches trigger a synthesised
# outline_level + is_heading=True after chain resolve, so re-emitting
# <w:outlineLvl> here would leak the value into descendants on round-trip.
_HEADING_NAME_RE = re.compile(r"[Hh]eading\s*(\d+)")


def build_style_font_map(doc: ldm.Document) -> dict[str, ldm.Font]:
    """Map canonical style id → resolved :class:`Font` for paragraph styles.

    Only paragraph styles (``type == 1``) carry a chain-resolved font in
    the LDM (the reader uses ``_build_resolved_font`` for those), so other
    style types are skipped.  Used by ``_custom_style`` for diff-based
    rPr emission against the basedOn chain.
    """
    out: dict[str, ldm.Font] = {}
    for style in doc.styles:
        if style.type != 1 or style.font is None:
            continue
        canonical = style.name.replace(" ", "").lower()
        if canonical:
            out[canonical] = style.font
    return out


_NAMELESS_STYLE_PREFIX = "Style"


# Word's styleId resolver only matches [A-Za-z0-9_-] even though the
# schema allows any string — sanitise so <w:pStyle> lookups succeed
# (otherwise Word silently downgrades to Normal and rejects on strict load).
def _sanitize_style_id(raw: str) -> str:
    cleaned = "".join(c if c.isalnum() or c in ("-", "_") else "_" for c in raw)
    if not cleaned or not (cleaned[0].isalpha() or cleaned[0] == "_"):
        cleaned = "_" + cleaned
    return cleaned


def _render_style_mark_rPr(mark: Optional[ldm.Font], base_mark: Optional[ldm.Font]) -> str:
    """Style-level wrapper around :func:`render_rPr` for the paragraph mark.

    Mirrors :func:`paragraphs._render_mark_rPr` but lives here to avoid
    a circular import — ``paragraphs.py`` imports ``styles_part`` for
    ``build_style_id_map``, so the helper has to land on this side.
    """
    if mark is None:
        return ""
    rPr = render_rPr(mark, base=base_mark, for_style=True)
    return rPr or el("w:rPr")


def build_style_id_map(doc: ldm.Document) -> dict[str, str]:
    """Map every LDM style ``name`` → a unique on-disk ``w:styleId``.

    The reader collapses Word's distinct ``styleId``s into the display
    ``name`` it stored on each :class:`Style`, so two styles named e.g.
    ``"Body Text 2"`` and ``"BodyText 2"`` both lose their differentiator.
    The naive ``name.replace(" ", "")`` scheme then assigns them the same
    id, and our ``seen_ids`` dedup drops the second style on round-trip.

    Resolve collisions by appending ``0`` first (matches Word's own
    behaviour for the ``Test19075`` fixture), then ``_2``, ``_3``…  The
    map is built once per document and consulted by every site that
    needs a styleId — ``_custom_style``, ``basedOn`` / ``next``
    references, and paragraph-level ``<w:pStyle>``.
    """
    out: dict[str, str] = {}
    used: set[str] = set()
    nameless_idx = 0
    for style in doc.styles:
        name = style.name
        if not name:
            if style.built_in and style.style_identifier:
                canonical_id = IDENTIFIER_TO_STYLE_ID.get(style.style_identifier)
                if canonical_id and canonical_id not in used:
                    used.add(canonical_id)
                    out[f"<nameless:{id(style)}>"] = canonical_id
                    continue
            nameless_idx += 1
            candidate = f"_{_NAMELESS_STYLE_PREFIX}{nameless_idx}"
            while candidate in used:
                nameless_idx += 1
                candidate = f"_{_NAMELESS_STYLE_PREFIX}{nameless_idx}"
            used.add(candidate)
            out[f"<nameless:{id(style)}>"] = candidate
            continue
        # For built-in styles, use the canonical styleId from the
        # identifier lookup table (e.g. HEADING_1 → "Heading1").
        if style.built_in and style.style_identifier:
            canonical_id = IDENTIFIER_TO_STYLE_ID.get(style.style_identifier)
            if canonical_id and canonical_id not in used:
                used.add(canonical_id)
                out[name] = canonical_id
                continue
        candidate = _sanitize_style_id(name.replace(" ", ""))
        if candidate in used:
            base = candidate
            i = 0
            while candidate in used:
                i += 1
                candidate = f"{base}{'0' if i == 1 else f'_{i}'}"
        used.add(candidate)
        out[name] = candidate
    return out


def _style_id(style: ldm.Style, style_id_map: dict[str, str]) -> str:
    """Resolve a style's on-disk id, including the synthesised form for
    nameless source styles."""
    if style.name:
        return style_id_map.get(style.name, _sanitize_style_id(style.name.replace(" ", "")))
    return style_id_map.get(f"<nameless:{id(style)}>", "_StyleX")


def _is_custom_style_name(name: str) -> bool:
    """Detect whether ``name`` would be capitalized by the reader on read.

    The reader treats unknown lowercase names as built-ins and applies
    title-casing (``paragraph`` → ``Paragraph``).  We mark any style whose
    name does not survive that canonicalisation as ``w:customStyle="1"``
    so the reader leaves the original case alone on the next round-trip.
    """
    return _canonicalize_style_name(name) != name


def _heading_style(level: int) -> str:
    """Built-in ``HeadingN`` paragraph style.

    Outline level encodes the heading level so readers can recover it
    via ``pPr/outlineLvl`` even without a styles map.
    """
    style_id = f"Heading{level}"
    name = f"heading {level}"
    pPr = el(
        "w:pPr",
        None,
        [
            el("w:keepNext"),
            el("w:keepLines"),
            el("w:outlineLvl", {"w:val": level - 1}),
        ],
    )
    # Default size: 16pt → ~32 half-points for H1, decreasing by 2 each level
    # (approximation matching Word defaults; precise values come from theme).
    size = max(11, 18 - (level - 1) * 2)
    rPr = el(
        "w:rPr",
        None,
        [
            el("w:b"),
            el("w:sz", {"w:val": pt_to_half_pt(size)}),
            el("w:szCs", {"w:val": pt_to_half_pt(size)}),
        ],
    )
    return el(
        "w:style",
        {"w:type": "paragraph", "w:styleId": style_id},
        [
            el("w:name", {"w:val": name}),
            el("w:basedOn", {"w:val": "Normal"}),
            el("w:next", {"w:val": "Normal"}),
            pPr,
            rPr,
        ],
    )


def _normal_style() -> str:
    return el(
        "w:style",
        {"w:type": "paragraph", "w:default": "1", "w:styleId": "Normal"},
        el("w:name", {"w:val": "Normal"}),
    )


def _hyperlink_style() -> str:
    """Word's built-in ``Hyperlink`` character style: blue + underlined.

    Runs synthesised by the markdown-link rewrite carry
    ``rStyle="Hyperlink"`` so this row gives them the expected visual
    treatment without polluting their direct formatting.
    """
    rPr = el(
        "w:rPr",
        None,
        [
            el("w:color", {"w:val": "0563C1"}),
            el("w:u", {"w:val": "single"}),
        ],
    )
    return el(
        "w:style",
        {"w:type": "character", "w:styleId": "Hyperlink"},
        [
            el("w:name", {"w:val": "Hyperlink"}),
            el("w:basedOn", {"w:val": "DefaultParagraphFont"}),
            el("w:uiPriority", {"w:val": "99"}),
            el("w:unhideWhenUsed"),
            rPr,
        ],
    )


def _custom_style(
    style: ldm.Style,
    style_pf_map: dict[str, ldm.ParagraphFormat],
    docDefaults_pf: Optional[ldm.ParagraphFormat],
    style_id_map: dict[str, str],
    style_font_map: dict[str, ldm.Font],
) -> str:
    """Render an LDM ``Style`` to ``<w:style>``.

    The resolved :class:`ParagraphFormat` is emitted as a *diff* against
    the basedOn ancestor (or ``docDefaults`` for top-level styles like
    ``Normal``) so inherited values aren't re-baked into every child on
    the next read.
    """
    style_id = _style_id(style, style_id_map)
    type_token = {1: "paragraph", 2: "character", 3: "table", 4: "numbering"}.get(
        style.type, "paragraph"
    )
    style_attrs: dict[str, object] = {"w:type": type_token, "w:styleId": style_id}
    if style.built_in and style.type == 1:
        is_default = style.name.lower() == "normal"
        if is_default:
            style_attrs["w:default"] = "1"
    if not style.built_in and (
        _is_custom_style_name(style.name) or style.style_identifier != 0
    ):
        style_attrs["w:customStyle"] = "1"
    children: list[str] = [el("w:name", {"w:val": style.name})]
    if style.base_style_name:
        base_id = style_id_map.get(
            style.base_style_name,
            _sanitize_style_id(style.base_style_name.replace(" ", "")),
        )
        children.append(el("w:basedOn", {"w:val": base_id}))
    if style.next_paragraph_style_name:
        next_id = style_id_map.get(
            style.next_paragraph_style_name,
            _sanitize_style_id(style.next_paragraph_style_name.replace(" ", "")),
        )
        children.append(el("w:next", {"w:val": next_id}))
    if style.priority != 99:
        children.append(el("w:uiPriority", {"w:val": str(style.priority)}))
    if style.paragraph_format is not None:
        # Paragraph styles diff against the basedOn chain (resolved pf);
        # table/character/numbering styles diff against zero (raw pPr).
        base_pf: Optional[ldm.ParagraphFormat] = None
        if style.type == 1:
            if style.base_style_name:
                canonical = style.base_style_name.replace(" ", "").lower()
                base_pf = style_pf_map.get(canonical)
            if base_pf is None:
                base_pf = docDefaults_pf
        pf = style.paragraph_format
        # Synthesised outline_level (e.g. "Heading 0" → -1) means source
        # had no <w:outlineLvl>; suppress to avoid descendant is_heading flip.
        heading_match = _HEADING_NAME_RE.search(style.name)
        if heading_match:
            implied = int(heading_match.group(1)) - 1
            if implied < 0 and pf.outline_level == implied:
                if base_pf is None:
                    base_pf = ldm.ParagraphFormat(outline_level=implied)
                else:
                    base_pf = base_pf.model_copy(update={"outline_level": implied})
        pPr_children = pf_to_pPr_children(pf, base=base_pf)
        # Paragraph-mark rPr (``pPr/rPr``) is style-level formatting for
        # the bullet/number glyph and pilcrow.  Emit a diff against the
        # basedOn's mark font so child styles only carry their explicit
        # overrides instead of duplicating the inherited block.
        base_mark = base_pf.paragraph_mark_font if base_pf else None
        mark_rPr = _render_style_mark_rPr(pf.paragraph_mark_font, base_mark)
        if mark_rPr:
            pPr_children.append(mark_rPr)
        if pPr_children:
            children.append(el("w:pPr", None, pPr_children))
        elif style.type != 1:
            # The reader uses the *presence* of ``<w:pPr>`` to decide
            # whether to populate ``Style.paragraph_format`` for
            # non-paragraph styles (table / character / numbering).
            # Emit an empty placeholder so a style whose source pPr
            # carried only ``<w:numPr>`` (encoded outside the pf in our
            # LDM) survives the round-trip.
            children.append(el("w:pPr"))
    if style.font is not None:
        # Mirror the paragraph_format diff strategy: paragraph-style fonts
        # are reader-resolved through the basedOn chain, so we diff
        # against the basedOn's resolved font to recover only the explicit
        # overrides; other style types carry raw rPr fields, so the diff
        # base is the zero default.  Always emit ``<w:rPr/>`` even when
        # the diff is empty — the reader uses the element's *presence*
        # to decide whether to populate ``Style.font`` at all.
        base_font: Optional[ldm.Font] = None
        if style.type == 1 and style.base_style_name:
            canonical = style.base_style_name.replace(" ", "").lower()
            base_font = style_font_map.get(canonical)
        rPr = render_rPr(style.font, base=base_font, for_style=True) or el("w:rPr")
        children.append(rPr)
    if style.table_style_format is not None:
        tblPr = _render_style_tblPr(style.table_style_format)
        if tblPr:
            children.append(tblPr)
    return el("w:style", style_attrs, children)


def _docDefaults(normal_pf: Optional[ldm.ParagraphFormat]) -> str:
    """Document-level defaults.

    The reader treats ``Normal``'s resolved :class:`ParagraphFormat` as the
    union of ``docDefaults`` + ``Normal``'s own ``<w:pPr>`` (whichever side
    of the chain set each field).  The LDM doesn't separate the two halves,
    so on write we put the entire resolved Normal back into ``docDefaults``;
    that way paragraphs with no ``pStyle`` (and therefore no style chain to
    walk) still inherit the same spacing they had originally.  Normal's own
    ``<w:pPr>`` then ends up empty under the diff-vs-basedOn rule.
    """
    pPrDefault_body: object = ""
    if normal_pf is not None:
        pPr_children = pf_to_pPr_children(normal_pf, base=None)
        if pPr_children:
            pPrDefault_body = el("w:pPr", None, pPr_children)
    # ``rPrDefault`` left empty: hard-coding ``<w:rFonts w:ascii="Calibri"/>``
    # would force every run that originally had no font name to inherit
    # ``"Calibri"`` on the next read, inflating ``font.name`` LDM-wide.
    # Documents that legitimately want a default font set it per-run via
    # the styles chain instead.
    return el(
        "w:docDefaults",
        None,
        [
            el("w:rPrDefault", None, el("w:rPr")),
            el("w:pPrDefault", None, pPrDefault_body) if pPrDefault_body else el("w:pPrDefault"),
        ],
    )


def _render_style_tblPr(tsf: ldm.TableStyleFormat) -> str:
    """Render ``<w:tblPr>`` for a table style definition."""
    children: list[str] = []
    if tsf.borders:
        tbl_borders = _table_borders(tsf.borders)
        if tbl_borders:
            children.append(tbl_borders)
    margins: list[str] = []
    for side, val in [
        ("top", tsf.top_padding),
        ("left", tsf.left_padding),
        ("bottom", tsf.bottom_padding),
        ("right", tsf.right_padding),
    ]:
        if val > 0:
            margins.append(el(f"w:{side}", {"w:w": pt_to_twips(val), "w:type": "dxa"}))
    if margins:
        children.append(el("w:tblCellMar", None, margins))
    if not children:
        return ""
    return el("w:tblPr", None, children)


def _collect_referenced_style_ids(
    doc: ldm.Document, style_id_map: dict[str, str]
) -> set[str]:
    """Return the set of style IDs that any paragraph in ``doc`` references.

    A style ID is considered referenced when a paragraph's resolved
    ``style_id`` (heading short form or ``style_name.replace(" ", "")``)
    appears anywhere — body, headers, footers, or table cells.  Used by
    :func:`render_styles_xml` to skip fallback emissions for unused
    built-ins so the LDM round-trip doesn't grow phantom Heading rows.
    """

    ids: set[str] = set()

    def walk(paragraphs: list[ldm.Paragraph]) -> None:
        for p in paragraphs:
            sid = _resolve_style_id(p.paragraph_format, style_id_map=style_id_map)
            if sid:
                ids.add(sid)

    walk(doc.header_paragraphs)
    walk(doc.footer_paragraphs)
    for sec in doc.sections:
        for child in sec.body.children:
            if isinstance(child, ldm.Paragraph):
                walk([child])
            elif isinstance(child, ldm.Table):
                _collect_table_style_ids(child, ids, style_id_map)
    # Hyperlinks are emitted with rStyle="Hyperlink" by ``_render_hyperlink_run``,
    # so we need the style def whenever any run carries a markdown link.
    if _doc_has_hyperlink_runs(doc):
        ids.add("Hyperlink")
    return ids


def _collect_table_style_ids(
    table: ldm.Table, ids: set[str], style_id_map: dict[str, str]
) -> None:

    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                sid = _resolve_style_id(p.paragraph_format, style_id_map=style_id_map)
                if sid:
                    ids.add(sid)
            for nested in cell.tables:
                _collect_table_style_ids(nested, ids, style_id_map)


def _doc_has_hyperlink_runs(doc: ldm.Document) -> bool:
    """Return True if any run carries the ``Hyperlink`` character style.

    Mirrors the gating in ``runs._render_hyperlink_run`` so the writer
    only adds the ``Hyperlink`` style fallback when something will
    actually reference it via ``rStyle``.
    """

    def runs(paragraphs: list[ldm.Paragraph]):
        for p in paragraphs:
            yield from p.runs

    def walk_table(table: ldm.Table):
        for row in table.rows:
            for cell in row.cells:
                yield from runs(cell.paragraphs)
                for nested in cell.tables:
                    yield from walk_table(nested)

    def is_hyperlink_run(r: ldm.Run) -> bool:
        return (r.font.style_name or "").replace(" ", "").lower() == "hyperlink"

    for r in runs(doc.header_paragraphs):
        if is_hyperlink_run(r):
            return True
    for r in runs(doc.footer_paragraphs):
        if is_hyperlink_run(r):
            return True
    for sec in doc.sections:
        for child in sec.body.children:
            if isinstance(child, ldm.Paragraph):
                for r in child.runs:
                    if is_hyperlink_run(r):
                        return True
            elif isinstance(child, ldm.Table):
                for r in walk_table(child):
                    if is_hyperlink_run(r):
                        return True
    return False


def render_styles_xml(doc: ldm.Document) -> str:
    """Build ``word/styles.xml`` payload.

    LDM-defined styles take priority over the hard-coded built-ins, so a
    document that round-trips through the reader keeps any per-heading
    italic/colour overrides defined in its original ``styles.xml``.

    Built-in fallbacks (``Normal``, ``Heading1``…``Heading9``,
    ``Hyperlink``) are only emitted when a paragraph actually references
    the corresponding ID.  Without this gate every round-trip would
    inflate the LDM ``styles`` list with phantom Heading rows the source
    document never carried.
    """
    style_pf_map: dict[str, ldm.ParagraphFormat] = {}
    for style in doc.styles:
        if style.type == 1 and style.paragraph_format is not None:
            canonical = style.name.replace(" ", "").lower()
            if canonical:
                style_pf_map[canonical] = style.paragraph_format

    docDefaults_pf = style_pf_map.get("normal")
    style_id_map = build_style_id_map(doc)
    style_font_map = build_style_font_map(doc)
    referenced_ids = _collect_referenced_style_ids(doc, style_id_map)

    # Aspose blank's <w:latentStyles> table — mirrors what Word
    # injects into every fresh doc to seed UI priorities for ~267
    # built-in styles.  Round-trips opaquely.
    children: list[str] = [_docDefaults(docDefaults_pf), latent_styles()]
    seen_ids: set[str] = set()

    for style in doc.styles:
        sid = _style_id(style, style_id_map)
        if not sid or sid in seen_ids:
            continue
        seen_ids.add(sid)
        children.append(
            _custom_style(style, style_pf_map, docDefaults_pf, style_id_map, style_font_map)
        )

    # Word always wants a Normal style row to exist, even if nothing
    # references it explicitly — emit the empty placeholder when the LDM
    # didn't already supply one.
    if "Normal" not in seen_ids:
        children.append(_normal_style())
        seen_ids.add("Normal")
    for level in range(1, 10):
        sid = f"Heading{level}"
        if sid in seen_ids or sid not in referenced_ids:
            continue
        seen_ids.add(sid)
        children.append(_heading_style(level))
    if "Hyperlink" not in seen_ids and "Hyperlink" in referenced_ids:
        children.append(_hyperlink_style())
        seen_ids.add("Hyperlink")

    root = el("w:styles", {"xmlns:w": W_URI}, children)
    return XML_DECL + root


# Re-export so callers don't import from runs just for one helper.
__all__ = ["render_styles_xml", "color_to_hex"]
