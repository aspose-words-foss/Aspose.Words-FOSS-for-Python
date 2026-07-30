"""Inline image rendering — turns LDM ``Shape`` into ``<w:drawing>``.

The reader stores image bytes on each ``Shape.image_data``.  The
writer's job is to (a) emit a ``<w:drawing>`` element referencing a
relationship id, (b) accumulate the image bytes in :class:`ImageRenderState`
so the package builder can drop them into ``word/media/`` and register a
matching ``Relationship`` row.

Only inline (``wp:inline``) drawings are emitted.  Anchored / floating
shapes carry positioning data the LDM exposes lossily, so they're routed
through the inline path which Word still renders correctly even though the
exact wrap mode is dropped.
"""


from dataclasses import dataclass, field
from typing import Any, Mapping, Optional

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss._visible_runs import is_horizontal_rule_shape
from aspose.words_foss.docx_writer.xml_utils import el

from aspose.words_foss.docx_writer.bookmarks import BookmarkState

# Drawing-ML namespace URIs — declared once at the document root so element
# tags can use bare ``wp:`` / ``a:`` / ``pic:`` prefixes.
WP_URI = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_URI = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_URI = "http://schemas.openxmlformats.org/drawingml/2006/picture"
REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

EMU_PER_PT = 12700  # 1 pt = 12700 English Metric Units (OOXML drawing space)
EMU_PER_MM = 36000  # 1 mm = 36000 EMUs — matches the reader's anchor parsing
_DEFAULT_IMAGE_PT = 96.0  # ~1.33 inches when reader didn't carry size

# WrapType integer (see ``light_document_model.Shape.wrap_type``) →
# OOXML element to emit inside ``<wp:anchor>``.  ``INLINE`` (0) routes
# through ``<wp:inline>`` and never appears here.
_WRAP_ELEMENT = {
    1: "wp:wrapTopAndBottom",
    2: "wp:wrapSquare",
    3: "wp:wrapNone",
    4: "wp:wrapTight",
    5: "wp:wrapThrough",
}

_WRAP_TEXT_REQUIRED = {"wp:wrapSquare", "wp:wrapTight", "wp:wrapThrough"}


def pt_to_emu(pt: float) -> int:
    """Points → EMUs, clamped to a strictly positive value (Word rejects 0)."""
    if pt <= 0:
        pt = _DEFAULT_IMAGE_PT
    return max(1, int(round(pt * EMU_PER_PT)))


def mm_to_emu(mm: float) -> int:
    """Millimetres → EMUs (used for anchored shape positioning)."""
    return int(round(mm * EMU_PER_MM))


# Map content type → file extension used inside ``word/media/``.
_CT_TO_EXT = {
    "image/png": "png",
    "image/jpeg": "jpeg",
    "image/jpg": "jpg",
    "image/gif": "gif",
    "image/bmp": "bmp",
    "image/tiff": "tiff",
    "image/x-emf": "emf",
    "image/x-wmf": "wmf",
    "image/svg+xml": "svg",
}


def _guess_extension(image: ldm.ImageData) -> str:
    """Pick the on-disk extension for an image's media filename."""
    ext = _CT_TO_EXT.get(image.content_type.lower())
    if ext:
        return ext
    name = image.source_full_name or ""
    if "." in name:
        return name.rsplit(".", 1)[1].lower()
    return "bin"


@dataclass
class ImageEntry:
    """One image to add to ``word/media/`` plus its relationship row."""

    rel_id: str
    media_path: str  # ``word/media/imageN.png``
    content_type: str
    image_bytes: bytes


_PT_PER_INCH = 72.0
_MM_PER_INCH = 25.4


def pt_to_mm(pt: float) -> float:
    """Points → millimetres (anchor offsets are stored in mm in the LDM)."""
    return pt / _PT_PER_INCH * _MM_PER_INCH


@dataclass
class ImageRenderState:
    """Accumulator threaded through paragraph rendering for inline shapes.

    ``media_name_prefix`` keeps the media filenames distinct between
    parts (body / header / footer) so they don't overwrite each other
    inside ``word/media/``; the relationship id namespace is also
    scoped to the prefix so the part-local ``.rels`` file's references
    stay unique."""

    images: list[ImageEntry] = field(default_factory=list)
    # ``wp:docPr/@id`` must be unique across the *entire* package
    # (body + every header + every footer).  Each part-level
    # ``ImageRenderState`` is initialised with a non-overlapping
    # starting offset so its ``<wp:docPr id="N"/>`` values can't
    # collide with another part's.  Word rejects packages with
    # duplicate docPr ids.
    _doc_pr: int = 0
    _seen: dict[bytes, str] = field(default_factory=dict)
    # Section margins flow in via :class:`SectionContext` so anchored
    # shapes can convert their absolute LDM coordinates back into the
    # ``relativeFrom="column"`` offsets the reader expects.
    section_left_margin_mm: float = 25.4
    section_top_margin_mm: float = 25.4
    media_name_prefix: str = "image"
    rel_id_prefix: str = "rIdImg"
    doc_pr_offset: int = 0  # starting offset for docPr id allocation
    # Emit image-less Shapes (text-boxes, cover-page rectangles,
    # side bands) as ``<w:drawing><wps:wsp>`` inside an
    # ``<mc:AlternateContent>`` block with a paired ``<w:pict><v:rect>``
    # VML fallback.  Modern Word picks ``<mc:Choice Requires="wps">``,
    # strict/legacy Word picks the Fallback — both render the same
    # rectangle and text-box content.  Toggle to ``False`` to drop
    # these shapes silently (pre-branch behaviour).
    emit_wps_shapes: bool = True

    def _next_doc_pr(self) -> int:
        self._doc_pr += 1
        return self._doc_pr + self.doc_pr_offset

    def _intern_image(self, image: ldm.ImageData) -> str:
        """Reuse the rId for an identical byte-for-byte image.

        Word documents commonly embed the same logo many times — the reader
        produces multiple ``Shape``s pointing at distinct in-memory
        ``ImageData`` instances with the same bytes.  Keying on the byte
        payload keeps the resulting docx compact and matches the reader's
        de-dup behaviour on next read (one media file → one rId → many
        Shapes).
        """
        existing = self._seen.get(image.image_bytes)
        if existing is not None:
            return existing

        idx = len(self.images) + 1
        rel_id = f"{self.rel_id_prefix}{idx}"
        ext = _guess_extension(image)
        media_path = f"word/media/{self.media_name_prefix}{idx}.{ext}"
        content_type = image.content_type or _ext_to_content_type(ext)
        self.images.append(
            ImageEntry(
                rel_id=rel_id,
                media_path=media_path,
                content_type=content_type,
                image_bytes=image.image_bytes,
            )
        )
        self._seen[image.image_bytes] = rel_id
        return rel_id

    def render_inline_shape(
        self,
        shape: ldm.Shape,
        rels: Optional[dict] = None,
        *,
        num_id_map: Optional[Mapping[int, int]] = None,
        bookmark_state: Optional[BookmarkState] = None,
        style_pf_map: Optional[Mapping[str, ldm.ParagraphFormat]] = None,
        style_id_map: Optional[Mapping[str, str]] = None,
        style_font_map: Optional[Mapping[str, ldm.Font]] = None,
    ) -> Optional[str]:
        """Emit ``<w:r><w:drawing>…</w:drawing></w:r>`` for ``shape``.

        Image-bearing shapes go through ``<pic:pic>`` (the standard
        DrawingML picture path).  Shapes with no image bytes — text
        boxes, cover-page rectangles, side bands — go through
        ``<wps:wsp>`` so the reader recovers them via
        ``_extract_positioned_shapes`` on the next round-trip.

        ``rels`` and the remaining kwargs thread per-document /
        per-part rendering context into text-box paragraphs — without
        them hyperlinks, bookmarks, list numIds, custom style ids and
        nested shapes inside text-boxes would be dropped.

        Returns ``None`` only when the shape has nothing worth emitting
        (no image, no fill, no outline, no text-box content).
        """
        if is_horizontal_rule_shape(shape):
            return _render_hr_pict_run(shape, self._next_doc_pr())
        image = shape.image_data
        if image is not None and image.image_bytes:
            rel_id = self._intern_image(image)
            return _render_drawing_run(
                shape,
                rel_id,
                self._next_doc_pr(),
                section_left_margin_mm=self.section_left_margin_mm,
                section_top_margin_mm=self.section_top_margin_mm,
            )
        # Image-less shape (text-box / rectangle).  Default emits both
        # the modern ``<wps:wsp>`` and a VML ``<v:rect>`` fallback
        # inside ``<mc:AlternateContent>`` — see ``emit_wps_shapes``
        # docstring for the rationale.  ``False`` drops these shapes
        # entirely (cover-page bars / floating text-boxes are gone).
        if not self.emit_wps_shapes:
            return None
        has_fill = bool(shape.fill_color)
        has_outline = bool(shape.stroke and (shape.stroke.line_style != 0 or shape.stroke.line_width > 0))
        has_text_box = bool(shape.text_box and shape.text_box.get("paragraphs"))
        if not (has_fill or has_outline or has_text_box):
            return None
        return _render_wsp_drawing_run(
            shape,
            self._next_doc_pr(),
            section_left_margin_mm=self.section_left_margin_mm,
            section_top_margin_mm=self.section_top_margin_mm,
            rels=rels if rels is not None else {},
            num_id_map=num_id_map,
            image_state=self,
            bookmark_state=bookmark_state,
            style_pf_map=style_pf_map,
            style_id_map=style_id_map,
            style_font_map=style_font_map,
        )

    def set_section_margins(self, page_setup: "ldm.PageSetup") -> None:
        """Update the cached section margins.

        Called once per section enter so subsequent anchor renderings
        (``_render_anchor``) emit the right column-relative offsets.
        """
        self.section_left_margin_mm = pt_to_mm(page_setup.left_margin or 72.0)
        self.section_top_margin_mm = pt_to_mm(page_setup.top_margin or 72.0)


# ─────────────────────────────────────────────
# XML emission
# ─────────────────────────────────────────────


def _ext_to_content_type(ext: str) -> str:
    """Reverse of ``_CT_TO_EXT`` for filenames the reader couldn't classify."""
    for ct, e in _CT_TO_EXT.items():
        if e == ext:
            return ct
    return "application/octet-stream"


def _render_drawing_run(
    shape: ldm.Shape,
    rel_id: str,
    doc_pr_id: int,
    *,
    section_left_margin_mm: float = 25.4,
    section_top_margin_mm: float = 25.4,
) -> str:
    """Render the full ``<w:r><w:drawing>…</w:r>`` chain.

    Inline shapes (``is_inline=True`` or unset) are wrapped in ``<wp:inline>``;
    anchored shapes (``is_inline=False``) are wrapped in ``<wp:anchor>`` and
    carry their ``left`` / ``top`` / ``wrap_type`` so the next read sees a
    floating shape with the same position rather than collapsing it inline.

    The reader stores ``width``/``height`` in millimetres for anchored shapes
    that were "promoted" to absolute page coordinates (cover-page logos, etc.)
    and in points everywhere else — the runtime ``_is_positioned`` flag tells
    us which side we're on.
    """
    is_positioned = getattr(shape, "_is_positioned", False)
    # Positioned shapes have explicit mm dimensions from the reader's
    # anchor extraction; preserve zero-height ones as-is (they're
    # used as invisible layout markers).  Only inline shapes that
    # lack dimensions get the fallback default.
    if is_positioned:
        cx = mm_to_emu(shape.width if shape.width is not None else 0.0)
        cy = mm_to_emu(shape.height if shape.height is not None else 0.0)
    else:
        width = shape.width if shape.width and shape.width > 0 else _DEFAULT_IMAGE_PT
        height = shape.height if shape.height and shape.height > 0 else _DEFAULT_IMAGE_PT
        cx = pt_to_emu(width)
        cy = pt_to_emu(height)

    name = shape.name or (
        shape.image_data.source_full_name if shape.image_data else f"Picture {doc_pr_id}"
    )

    graphic = _render_graphic(rel_id, name, cx, cy, shape)
    cnv_gfp = el(
        "wp:cNvGraphicFramePr",
        None,
        el("a:graphicFrameLocks"),
    )
    is_inline = shape.is_inline is None or shape.is_inline
    if is_inline:
        wrapper = el(
            "wp:inline",
            {"distT": 0, "distB": 0, "distL": 0, "distR": 0},
            [
                el("wp:extent", {"cx": cx, "cy": cy}),
                el("wp:effectExtent", {"l": 0, "t": 0, "r": 0, "b": 0}),
                el("wp:docPr", {"id": doc_pr_id, "name": name}),
                cnv_gfp,
                graphic,
            ],
        )
    else:
        wrapper = _render_anchor(
            shape,
            doc_pr_id,
            cx,
            cy,
            name,
            graphic,
            section_left_margin_mm=section_left_margin_mm,
            section_top_margin_mm=section_top_margin_mm,
        )
    return el("w:r", None, el("w:drawing", None, wrapper))


def _render_graphic(
    rel_id: str,
    name: str,
    cx: int,
    cy: int,
    shape: ldm.Shape,
) -> str:
    """Build the shared ``<a:graphic>…<pic:pic/>…</a:graphic>`` subtree."""
    pic_nvPicPr = el(
        "pic:nvPicPr",
        None,
        [
            el("pic:cNvPr", {"id": "0", "name": name}),
            el("pic:cNvPicPr"),
        ],
    )
    blip_fill_children: list[str] = [el("a:blip", {"r:embed": rel_id})]
    img = shape.image_data
    if img and (img.crop_left or img.crop_top or img.crop_right or img.crop_bottom):
        attrs: dict[str, float] = {}
        if img.crop_left:
            attrs["l"] = img.crop_left
        if img.crop_top:
            attrs["t"] = img.crop_top
        if img.crop_right:
            attrs["r"] = img.crop_right
        if img.crop_bottom:
            attrs["b"] = img.crop_bottom
        blip_fill_children.append(el("a:srcRect", attrs))
    blip_fill_children.append(el("a:stretch", None, el("a:fillRect")))
    pic_blipFill = el("pic:blipFill", None, blip_fill_children)
    pic_spPr = el(
        "pic:spPr",
        None,
        [
            el(
                "a:xfrm",
                None,
                [
                    el("a:off", {"x": 0, "y": 0}),
                    el("a:ext", {"cx": cx, "cy": cy}),
                ],
            ),
            el("a:prstGeom", {"prst": "rect"}, el("a:avLst")),
        ],
    )
    # All namespace prefixes (``a:``, ``pic:``, ``wps:``) are declared
    # once on the document root; emitting them again here would
    # duplicate the binding and trip MS Word's strict OPC validator.
    pic_pic = el("pic:pic", None, [pic_nvPicPr, pic_blipFill, pic_spPr])
    return el(
        "a:graphic",
        None,
        el("a:graphicData", {"uri": PIC_URI}, pic_pic),
    )


def _anchor_lock_attrs() -> Optional[dict[str, object]]:
    # ``xmlns:a`` is declared once at the document root — emitting it
    # again on every ``<a:graphicFrameLocks>`` instance just duplicates
    # the binding and bloats the package.
    return None


# Reverse of ``shapes._H_REL_FROM`` / ``_V_REL_FROM``: LDM int →
# ``positionH/positionV @relativeFrom`` token Word expects.  Keep these
# next to the writer so round-trip stays symmetric with the reader.
_H_REL_FROM_TOKEN = {
    0: "margin", 1: "page", 2: "column", 3: "character",
    4: "leftMargin", 5: "rightMargin",
    6: "insideMargin", 7: "outsideMargin",
}
_V_REL_FROM_TOKEN = {
    0: "margin", 1: "page", 2: "paragraph", 3: "line",
    4: "topMargin", 5: "bottomMargin",
    6: "insideMargin", 7: "outsideMargin",
}

# CT_PosH / CT_PosV align tokens (mirror reader's _H_ALIGN / _V_ALIGN).
_H_ALIGN_TOKEN = {1: "left", 2: "center", 3: "right", 4: "inside", 5: "outside"}
_V_ALIGN_TOKEN = {1: "top", 2: "center", 3: "bottom", 4: "inside", 5: "outside"}


def _render_anchor(
    shape: ldm.Shape,
    doc_pr_id: int,
    cx: int,
    cy: int,
    name: str,
    graphic: str,
    *,
    section_left_margin_mm: float,
    section_top_margin_mm: float,
) -> str:
    """Render a ``<wp:anchor>`` carrying the shape's position + wrap type.

    For shapes that carry an explicit ``relative_horizontal_position`` /
    ``relative_vertical_position`` plus ``horizontal_position`` /
    ``vertical_position`` (filled in by
    :func:`shapes._apply_anchor_metadata`), emit those values verbatim so
    the LDM round-trip preserves the source's anchor frame.  Older code
    paths populate only ``left`` / ``top`` (absolute mm) — those fall
    through to the legacy ``"page"``-for-positioned /
    ``"column"`` + ``"paragraph"`` heuristic and subtract section margins.
    """
    # has_explicit_offset: LDM carries a non-default pt-offset that was
    # the source's exact <wp:posOffset>.  Distinct from "carries a
    # relative_from token" — a wpg child has relativeFrom on the group
    # but no per-child offset (its offset lives in _page_left_mm/_page_top_mm).
    has_explicit_offset = (
        shape.left != 0.0 or shape.top != 0.0
    )
    has_anchor_metadata = has_explicit_offset or (
        shape.relative_horizontal_position != 0
        or shape.relative_vertical_position != 0
    )
    # _is_positioned shapes (wpg-group children, cover-page rectangles)
    # carry absolute page coords in left/top — use those, not the
    # group-level offset, so each child lands at its individual mm
    # position rather than stacking at the group origin.
    is_positioned = getattr(shape, "_is_positioned", False) and not has_explicit_offset
    if is_positioned:
        # _is_positioned shapes carry absolute page coordinates;
        # subtract the matching margin baseline so the source's
        # relativeFrom token survives even without wpg reconstruction.
        relative_from_h = _H_REL_FROM_TOKEN.get(shape.relative_horizontal_position, "page")
        relative_from_v = _V_REL_FROM_TOKEN.get(shape.relative_vertical_position, "page")
        if relative_from_h in ("margin", "column", "character", "leftMargin"):
            pos_x_emu = mm_to_emu(shape._page_left_mm - section_left_margin_mm)
        else:
            pos_x_emu = mm_to_emu(shape._page_left_mm)
        if relative_from_v in ("margin", "paragraph", "line", "topMargin"):
            pos_y_emu = mm_to_emu(shape._page_top_mm - section_top_margin_mm)
        else:
            pos_y_emu = mm_to_emu(shape._page_top_mm)
    elif has_anchor_metadata:
        relative_from_h = _H_REL_FROM_TOKEN.get(shape.relative_horizontal_position, "margin")
        relative_from_v = _V_REL_FROM_TOKEN.get(shape.relative_vertical_position, "margin")
        # ``pt_to_emu`` clamps non-positive values to a default (because
        # ``wp:extent`` rejects zero); use a non-clamping ``round(...)``
        # for ``wp:posOffset`` since negative offsets are valid (a shape
        # anchored above the paragraph baseline carries ``posV<0``).
        pos_x_emu = int(round(shape.left * EMU_PER_PT))
        pos_y_emu = int(round(shape.top * EMU_PER_PT))
    else:
        relative_from_h = "column"
        relative_from_v = "paragraph"
        pos_x_emu = mm_to_emu(shape._page_left_mm - section_left_margin_mm)
        pos_y_emu = mm_to_emu(shape._page_top_mm - section_top_margin_mm)
    wrap_elem = _WRAP_ELEMENT.get(shape.wrap_type, "wp:wrapNone")
    wrap_attrs = {"wrapText": "bothSides"} if wrap_elem in _WRAP_TEXT_REQUIRED else None
    # CT_WrapTight / CT_WrapThrough require <wp:wrapPolygon> (minOccurs=1).
    wrap_children: list[str] | None = None
    if wrap_elem in ("wp:wrapTight", "wp:wrapThrough"):
        wrap_children = [
            el(
                "wp:wrapPolygon",
                {"edited": "0"},
                [
                    el("wp:start", {"x": 0, "y": 0}),
                    el("wp:lineTo", {"x": 0, "y": 21600}),
                    el("wp:lineTo", {"x": 21600, "y": 21600}),
                    el("wp:lineTo", {"x": 21600, "y": 0}),
                    el("wp:lineTo", {"x": 0, "y": 0}),
                ],
            )
        ]
    return el(
        "wp:anchor",
        {
            "distT": 0,
            "distB": 0,
            "distL": 0,
            "distR": 0,
            "simplePos": "0",
            "relativeHeight": "1",
            "behindDoc": "1" if shape.behind_text else "0",
            "locked": "1" if shape.anchor_locked else "0",
            "layoutInCell": "1" if shape.is_layout_in_cell else "0",
            "allowOverlap": "1" if shape.allow_overlap else "0",
        },
        [
            el("wp:simplePos", {"x": 0, "y": 0}),
            el(
                "wp:positionH",
                {"relativeFrom": relative_from_h},
                # CT_PosH: choice of posOffset or align — prefer align
                # when the LDM carries one (matches source semantics).
                el("wp:align", None, _H_ALIGN_TOKEN[shape.horizontal_alignment])
                if shape.horizontal_alignment in _H_ALIGN_TOKEN
                else el("wp:posOffset", None, str(pos_x_emu)),
            ),
            el(
                "wp:positionV",
                {"relativeFrom": relative_from_v},
                el("wp:align", None, _V_ALIGN_TOKEN[shape.vertical_alignment])
                if shape.vertical_alignment in _V_ALIGN_TOKEN
                else el("wp:posOffset", None, str(pos_y_emu)),
            ),
            el("wp:extent", {"cx": cx, "cy": cy}),
            el("wp:effectExtent", {"l": 0, "t": 0, "r": 0, "b": 0}),
            el(wrap_elem, wrap_attrs, wrap_children),
            el("wp:docPr", {"id": doc_pr_id, "name": name}),
            el(
                "wp:cNvGraphicFramePr",
                None,
                el("a:graphicFrameLocks", _anchor_lock_attrs()),
            ),
            graphic,
        ],
    )


# ----------------------------------------------------------------------
# Image-less wps:wsp shape rendering
# ----------------------------------------------------------------------

WPS_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
_WPS_GRAPHIC_DATA_URI = WPS_URI


def _color_to_srgb_hex(color: str) -> Optional[str]:
    """Convert the LDM colour-string format to a 6-digit hex.

    Returns ``None`` for empty / sentinel values so the caller can
    decide whether to emit no fill at all.
    """
    if not color or color in ("auto", "Color [Empty]"):
        return None
    if color.startswith("#") and len(color) in (7, 9):
        return color[1:7].upper()
    if color.startswith("Color [") and color.endswith("]"):
        try:
            parts = color[len("Color ["):-1].split(",")
            channels: dict[str, int] = {}
            for p in parts:
                k, _, v = p.partition("=")
                channels[k.strip()] = int(v.strip())
            r = channels.get("R", 0)
            g = channels.get("G", 0)
            b = channels.get("B", 0)
            return f"{r:02X}{g:02X}{b:02X}"
        except (ValueError, KeyError):
            return None
    if len(color) == 6:
        try:
            int(color, 16)
            return color.upper()
        except ValueError:
            return None
    return None


def _render_wsp_sp_pr(shape: ldm.Shape, cx: int, cy: int) -> str:
    """Build the ``<wps:spPr>`` block: xfrm + prstGeom + optional fill / outline."""
    children: list[str] = [
        el(
            "a:xfrm",
            None,
            [el("a:off", {"x": 0, "y": 0}), el("a:ext", {"cx": cx, "cy": cy})],
        ),
        el("a:prstGeom", {"prst": "rect"}, el("a:avLst")),
    ]
    fill_hex = _color_to_srgb_hex(shape.fill_color)
    if fill_hex:
        children.append(
            el("a:solidFill", None, el("a:srgbClr", {"val": fill_hex}))
        )
    border = shape.stroke
    if border and (border.line_style != 0 or border.line_width > 0):
        line_attrs: dict[str, object] = {}
        if border.line_width > 0:
            line_attrs["w"] = int(round(border.line_width * EMU_PER_PT))
        ln_children: list[str] = []
        line_hex = _color_to_srgb_hex(border.color)
        if line_hex:
            ln_children.append(
                el("a:solidFill", None, el("a:srgbClr", {"val": line_hex}))
            )
        children.append(el("a:ln", line_attrs or None, ln_children))
    return el("wps:spPr", None, children)


_VERT_ANCHOR_TOKEN = {0: "t", 1: "ctr", 2: "b"}

# OOXML defaults from ECMA-376 (in EMU):
#   lIns / rIns = 91440  (2.54 mm)
#   tIns / bIns = 45720  (1.27 mm)
# The reader applies the same defaults when ``wps:bodyPr`` is absent,
# so emitting them explicitly is only necessary when the LDM carries a
# non-default value.
_DEFAULT_BODY_INS_MM = (2.54, 1.27, 2.54, 1.27)  # (lIns, tIns, rIns, bIns)


def _render_wsp_body_pr(shape: ldm.Shape) -> str:
    """Build the ``<wps:bodyPr>`` element carrying vertical alignment + insets."""
    attrs: dict[str, object] = {}
    anchor = _VERT_ANCHOR_TOKEN.get(shape.text_box_anchor)
    if anchor and anchor != "t":
        attrs["anchor"] = anchor
    insets = shape.text_box.get("insets_mm") if shape.text_box else None
    if insets and tuple(insets) != _DEFAULT_BODY_INS_MM:
        l_mm, t_mm, r_mm, b_mm = insets
        attrs["lIns"] = int(round(l_mm * EMU_PER_MM))
        attrs["tIns"] = int(round(t_mm * EMU_PER_MM))
        attrs["rIns"] = int(round(r_mm * EMU_PER_MM))
        attrs["bIns"] = int(round(b_mm * EMU_PER_MM))
    return el("wps:bodyPr", attrs or None)


def _render_textbox_paragraphs(
    paragraphs: list[Any],
    rels: dict,
    *,
    num_id_map: Optional[Mapping[int, int]] = None,
    image_state: Optional[ImageRenderState] = None,
    bookmark_state: Optional[BookmarkState] = None,
    style_pf_map: Optional[Mapping[str, ldm.ParagraphFormat]] = None,
    style_id_map: Optional[Mapping[str, str]] = None,
    style_font_map: Optional[Mapping[str, ldm.Font]] = None,
) -> str:
    """Render ``shape.text_box.paragraphs`` to ``<w:txbxContent>``.

    Imported lazily to avoid a circular import with ``paragraphs``.
    Optional kwargs are forwarded so bookmarks / list numIds / custom
    style ids / nested shapes inside the text-box survive round-trip.
    """
    from aspose.words_foss.docx_writer.paragraphs import render_paragraph

    if not paragraphs:
        return ""
    body: list[str] = []
    render_kwargs: dict[str, Any] = {}
    if num_id_map is not None:
        render_kwargs["num_id_map"] = num_id_map
    if image_state is not None:
        render_kwargs["image_state"] = image_state
    if bookmark_state is not None:
        render_kwargs["bookmark_state"] = bookmark_state
    if style_pf_map is not None:
        render_kwargs["style_pf_map"] = style_pf_map
    if style_id_map is not None:
        render_kwargs["style_id_map"] = style_id_map
    if style_font_map is not None:
        render_kwargs["style_font_map"] = style_font_map
    for para in paragraphs:
        body.append(render_paragraph(para, rels, **render_kwargs))
    return el("w:txbxContent", None, body)


def _render_hr_pict_run(shape: ldm.Shape, pict_id: int) -> str:
    """Emit ``<w:r><w:pict><v:rect o:hr="t">`` — how Word stores a horizontal rule."""
    width_pt = shape.width if shape.width else 432.0
    height_pt = shape.height if shape.height else 1.5
    rect = el(
        "v:rect",
        {
            "id": f"_x0000_i{1024 + pict_id}",
            "style": f"width:{width_pt:.2f}pt;height:{height_pt:.2f}pt",
            "o:hrpct": "1000",
            "o:hrstd": "t",
            "o:hr": "t",
            "filled": "t",
            "fillcolor": "gray",
            "stroked": "f",
        },
        el("v:path", {"strokeok": "f"}),
    )
    return el("w:r", None, el("w:pict", None, rect))


def _render_wsp_drawing_run(
    shape: ldm.Shape,
    doc_pr_id: int,
    *,
    section_left_margin_mm: float = 25.4,
    section_top_margin_mm: float = 25.4,
    rels: Optional[dict] = None,
    num_id_map: Optional[Mapping[int, int]] = None,
    image_state: Optional[ImageRenderState] = None,
    bookmark_state: Optional[BookmarkState] = None,
    style_pf_map: Optional[Mapping[str, ldm.ParagraphFormat]] = None,
    style_id_map: Optional[Mapping[str, str]] = None,
    style_font_map: Optional[Mapping[str, ldm.Font]] = None,
) -> str:
    """Render an image-less Shape as ``<w:r><w:drawing><wp:anchor><wps:wsp>``.

    Used for cover-page rectangles, side bands, and stand-alone text
    boxes — anything where the reader recovered the shape via
    :func:`shapes._make_wsp_shape` rather than the picture path.  The
    emitted ``<a:graphicData uri>`` matches the reader's
    ``wordprocessingShape`` lookup so the next read sees the same
    fill / outline / text-box content.
    """
    is_positioned = getattr(shape, "_is_positioned", False)
    # Positioned shapes have explicit mm dimensions from the reader's
    # anchor extraction; preserve zero-height ones as-is (they're
    # used as invisible layout markers).  Only inline shapes that
    # lack dimensions get the fallback default.
    if is_positioned:
        cx = mm_to_emu(shape.width if shape.width is not None else 0.0)
        cy = mm_to_emu(shape.height if shape.height is not None else 0.0)
    else:
        width = shape.width if shape.width and shape.width > 0 else _DEFAULT_IMAGE_PT
        height = shape.height if shape.height and shape.height > 0 else _DEFAULT_IMAGE_PT
        cx = pt_to_emu(width)
        cy = pt_to_emu(height)

    name = shape.name or f"Shape {doc_pr_id}"

    # ``wps:cNvSpPr@txBox="1"`` flags this shape as a text box so MS
    # Word renders the ``<wps:txbx>`` content; without it strict Word
    # builds reject the wsp as missing a required hint.
    has_text_box = bool(shape.text_box and shape.text_box.get("paragraphs"))
    cnv_attrs = {"txBox": "1"} if has_text_box else None
    wsp_children: list[str] = [
        el("wps:cNvSpPr", cnv_attrs),
        _render_wsp_sp_pr(shape, cx, cy),
    ]
    if has_text_box:
        txbx_body = _render_textbox_paragraphs(
            shape.text_box["paragraphs"],
            rels if rels is not None else {},
            num_id_map=num_id_map,
            image_state=image_state,
            bookmark_state=bookmark_state,
            style_pf_map=style_pf_map,
            style_id_map=style_id_map,
            style_font_map=style_font_map,
        )
        if txbx_body:
            wsp_children.append(el("wps:txbx", None, txbx_body))
    wsp_children.append(_render_wsp_body_pr(shape))

    # Namespaces ``a``, ``wps`` are declared at the document root.
    graphic = el(
        "a:graphic",
        None,
        el(
            "a:graphicData",
            {"uri": _WPS_GRAPHIC_DATA_URI},
            el("wps:wsp", None, wsp_children),
        ),
    )

    is_inline = shape.is_inline is None or shape.is_inline
    if is_inline:
        wrapper = el(
            "wp:inline",
            {"distT": 0, "distB": 0, "distL": 0, "distR": 0},
            [
                el("wp:extent", {"cx": cx, "cy": cy}),
                el("wp:effectExtent", {"l": 0, "t": 0, "r": 0, "b": 0}),
                el("wp:docPr", {"id": doc_pr_id, "name": name}),
                el(
                    "wp:cNvGraphicFramePr",
                    None,
                    el("a:graphicFrameLocks", _anchor_lock_attrs()),
                ),
                graphic,
            ],
        )
    else:
        wrapper = _render_anchor(
            shape,
            doc_pr_id,
            cx,
            cy,
            name,
            graphic,
            section_left_margin_mm=section_left_margin_mm,
            section_top_margin_mm=section_top_margin_mm,
        )
    drawing = el("w:drawing", None, wrapper)
    # Wrap raw <wps:wsp> in <mc:AlternateContent> (Choice=wps + VML
    # Fallback) — strict OOXML rejects the bare post-2007 shape, and
    # an empty Fallback.  VML mirrors geometry / fill / text-box so
    # legacy renderers still draw something.
    vml_fallback = _render_vml_fallback(
        shape,
        doc_pr_id,
        section_left_margin_mm=section_left_margin_mm,
        section_top_margin_mm=section_top_margin_mm,
        rels=rels,
        num_id_map=num_id_map,
        image_state=image_state,
        bookmark_state=bookmark_state,
        style_pf_map=style_pf_map,
        style_id_map=style_id_map,
        style_font_map=style_font_map,
    )
    alt = el(
        "mc:AlternateContent",
        None,
        [
            el("mc:Choice", {"Requires": "wps"}, drawing),
            el("mc:Fallback", None, vml_fallback),
        ],
    )
    return el("w:r", None, alt)


def _render_vml_fallback(
    shape: ldm.Shape,
    doc_pr_id: int,
    *,
    section_left_margin_mm: float,
    section_top_margin_mm: float,
    rels: Optional[dict] = None,
    num_id_map: Optional[Mapping[int, int]] = None,
    image_state: Optional[ImageRenderState] = None,
    bookmark_state: Optional[BookmarkState] = None,
    style_pf_map: Optional[Mapping[str, ldm.ParagraphFormat]] = None,
    style_id_map: Optional[Mapping[str, str]] = None,
    style_font_map: Optional[Mapping[str, ldm.Font]] = None,
) -> str:
    """Emit a legacy VML ``<w:pict><v:rect>`` mirror of ``shape``.

    Used as the ``<mc:Fallback>`` alternative when the writer emits a
    ``<wps:wsp>`` inside ``<mc:Choice Requires="wps">``.  VML is the
    pre-2007 vector shape grammar — Word universally renders it, so a
    strict loader that drops the Choice still gets a visually
    equivalent rectangle plus text-box from this branch.

    All sizes are emitted in points (``pt`` suffix), which is VML's
    natural unit.  Positioned shapes (``_is_positioned=True``) get
    absolute ``mso-position-*-relative:page`` anchors mirroring the
    Choice's ``relativeFrom="page"`` offsets; everything else uses
    column/paragraph-relative anchors that match the Choice anchor
    metadata.
    """
    is_positioned = getattr(shape, "_is_positioned", False)
    # Mirror the DrawingML Choice's relativeFrom + offset semantics so
    # VML and DrawingML place the shape at the *same* location.  When
    # the Choice uses explicit pt offsets (has_explicit_offset), the
    # VML margin-* values come from those same pt offsets relative to
    # the LDM's stored relative_from token — without this the Fallback
    # branch positioned the shape at `shape.top` mm from page top,
    # which for paragraph-anchored shapes (cover page AIA / July
    # textboxes) lands the rect on the *next* page once the anchor's
    # paragraph straddles a page break.
    has_explicit_offset = (
        shape.left != 0.0 or shape.top != 0.0
    )
    if has_explicit_offset:
        relative_from_h = _H_REL_FROM_TOKEN.get(shape.relative_horizontal_position, "margin")
        relative_from_v = _V_REL_FROM_TOKEN.get(shape.relative_vertical_position, "margin")
        # width/height live in mm for positioned shapes; convert to pt
        # without substituting a default (preserve zero-height markers).
        w_mm = shape.width if shape.width is not None else 0.0
        h_mm = shape.height if shape.height is not None else 0.0
        width_pt = w_mm / _MM_PER_INCH * _PT_PER_INCH
        height_pt = h_mm / _MM_PER_INCH * _PT_PER_INCH
        left_pt = shape.left
        top_pt = shape.top
    else:
        relative_from_h = "page" if is_positioned else "column"
        relative_from_v = "page" if is_positioned else "paragraph"
        width_pt, height_pt, left_pt, top_pt = _vml_geometry_pt(
            shape,
            is_positioned,
            section_left_margin_mm=section_left_margin_mm,
            section_top_margin_mm=section_top_margin_mm,
        )

    # VML ``style`` attribute is a CSS-like declaration list.
    style_parts = [
        "position:absolute",
        f"margin-left:{left_pt:.2f}pt",
        f"margin-top:{top_pt:.2f}pt",
        f"width:{width_pt:.2f}pt",
        f"height:{height_pt:.2f}pt",
    ]
    if shape.behind_text:
        style_parts.append("z-index:-251658240")
    style_parts.append(f"mso-position-horizontal-relative:{relative_from_h}")
    style_parts.append(f"mso-position-vertical-relative:{relative_from_v}")
    style = ";".join(style_parts)

    rect_attrs: dict[str, object] = {
        "id": f"VRect{doc_pr_id}",
        "o:spid": f"_x0000_s{1024 + doc_pr_id}",
        "style": style,
    }
    fill_hex = _color_to_srgb_hex(shape.fill_color)
    if fill_hex:
        rect_attrs["fillcolor"] = f"#{fill_hex}"
    rect_attrs["stroked"] = "f"

    children: list[str] = []
    if shape.text_box and shape.text_box.get("paragraphs"):
        txbx_body = _render_textbox_paragraphs(
            shape.text_box["paragraphs"],
            rels if rels is not None else {},
            num_id_map=num_id_map,
            image_state=image_state,
            bookmark_state=bookmark_state,
            style_pf_map=style_pf_map,
            style_id_map=style_id_map,
            style_font_map=style_font_map,
        )
        if txbx_body:
            ins = _vml_inset_attr(shape)
            tb_attrs = {"inset": ins} if ins else None
            children.append(el("v:textbox", tb_attrs, txbx_body))
    rect = el("v:rect", rect_attrs, children)
    return el("w:pict", None, rect)


def _vml_geometry_pt(
    shape: ldm.Shape,
    is_positioned: bool,
    *,
    section_left_margin_mm: float,
    section_top_margin_mm: float,
) -> tuple[float, float, float, float]:
    """Return ``(width_pt, height_pt, margin_left_pt, margin_top_pt)``."""
    if is_positioned:
        # mm dimensions from the reader's anchor extraction; preserve
        # zero-height markers (source's invisible layout anchors) as-is.
        w_mm = shape.width if shape.width is not None else 0.0
        h_mm = shape.height if shape.height is not None else 0.0
        width_pt = w_mm / _MM_PER_INCH * _PT_PER_INCH
        height_pt = h_mm / _MM_PER_INCH * _PT_PER_INCH
        left_pt = shape._page_left_mm / _MM_PER_INCH * _PT_PER_INCH
        top_pt = shape._page_top_mm / _MM_PER_INCH * _PT_PER_INCH
    else:
        raw_w = shape.width if shape.width and shape.width > 0 else _DEFAULT_IMAGE_PT
        raw_h = shape.height if shape.height and shape.height > 0 else _DEFAULT_IMAGE_PT
        width_pt = raw_w
        height_pt = raw_h
        # Same baseline subtraction the DrawingML branch uses so the
        # absolute mm coordinates the LDM stores become the same
        # margin-relative offsets in VML.
        left_pt = (shape._page_left_mm - section_left_margin_mm) / _MM_PER_INCH * _PT_PER_INCH
        top_pt = (shape._page_top_mm - section_top_margin_mm) / _MM_PER_INCH * _PT_PER_INCH
    return width_pt, height_pt, left_pt, top_pt


def _vml_inset_attr(shape: ldm.Shape) -> Optional[str]:
    """Return a comma-separated VML ``inset`` attribute (``lIns,tIns,rIns,bIns``)
    in points, or ``None`` when the shape carries no insets."""
    if not shape.text_box:
        return None
    insets = shape.text_box.get("insets_mm")
    if not insets:
        return None
    l_mm, t_mm, r_mm, b_mm = insets

    def mm_to_pt(mm: float) -> float:
        return mm / _MM_PER_INCH * _PT_PER_INCH

    return ",".join(f"{mm_to_pt(v):.2f}pt" for v in (l_mm, t_mm, r_mm, b_mm))
