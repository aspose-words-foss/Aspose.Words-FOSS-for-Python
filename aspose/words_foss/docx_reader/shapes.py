"""
Drawing and shape parsing for DOCX documents.

Handles DrawingML shapes, anchored positions, text boxes, and
group shape hierarchies.  These are mixed into DocumentReader
via multiple inheritance.
"""

from __future__ import annotations

from typing import Optional, Iterator, TYPE_CHECKING
from xml.etree import ElementTree as ET

from aspose.words_foss.docx_reader.constants import (
    A_NS,
    MC_NS,
    PIC_NS,
    R_NS,
    W_NS,
    WP_NS,
    WPG_NS,
    WPS_NS,
    _BODY_ANCHOR_MAP,
    _DEFAULT_LEFT_MARGIN_MM,
    _DEFAULT_PAGE_HEIGHT_MM,
    _DEFAULT_TOP_MARGIN_MM,
    _DML_MOD_SCALE,
    _EMU_PER_MM,
    _EMU_PER_PT,
    _MAX_COLOR_CHANNEL,
    _MM_PER_INCH,
    _POINTS_PER_INCH,
    COLOR_EMPTY,
)
from aspose.words_foss.docx_reader.utils import (
    _ext_to_content_type,
    _hex_to_aspose_color,
)

if TYPE_CHECKING:
    from aspose.words_foss import light_document_model as ldm


class ShapeParserMixin:
    """Mixin providing drawing/shape parsing methods for DocumentReader."""

    # These attributes are defined on DocumentReader but referenced here.
    _media: dict[str, bytes]
    _doc_image_rels: dict[str, str]
    _current_page_setup: "Optional[ldm.PageSetup]"
    _theme_colors: dict[str, str]

    def _build_drawing_shape(
        self, drawing_elem: ET.Element, image_rels: dict[str, str]
    ) -> "Optional[ldm.ShapeNode]":
        """Parse a <w:drawing> element and return a ShapeNode if it contains an image."""
        from aspose.words_foss import light_document_model as ldm

        inline = drawing_elem.find(f"{WP_NS}inline")
        anchor = drawing_elem.find(f"{WP_NS}anchor")
        container = inline if inline is not None else anchor
        if container is None:
            return None

        is_inline = inline is not None

        # Dimensions from wp:extent (EMU → points)
        extent = container.find(f"{WP_NS}extent")
        width_pt: Optional[float] = None
        height_pt: Optional[float] = None
        if extent is not None:
            cx = extent.get("cx")
            cy = extent.get("cy")
            if cx:
                width_pt = int(cx) / _EMU_PER_PT
            if cy:
                height_pt = int(cy) / _EMU_PER_PT

        # Find <a:blip r:embed="rIdN"/>
        blip = container.find(f".//{A_NS}blip")
        textbox_paragraphs = self._harvest_textbox_paragraphs(drawing_elem)

        if blip is None:
            # No picture — only a text box.  Still worth emitting a shape so
            # the writer can render the text (cover-page titles, side
            # captions, etc.).
            if not textbox_paragraphs:
                return None
            shape = ldm.ShapeNode()
            shape.has_image = False
            shape.is_inline = is_inline
            shape.width = width_pt
            shape.height = height_pt
            shape.text_box = {"paragraphs": textbox_paragraphs}
            return shape

        r_id = blip.get(f"{R_NS}embed")
        if not r_id:
            return None

        media_path = image_rels.get(r_id)
        if not media_path:
            return None

        image_bytes = self._media.get(media_path)
        if image_bytes is None:
            return None

        # Prefer the original name from pic:cNvPr[@name] if available
        cNvPr = container.find(f".//{PIC_NS}cNvPr")
        if cNvPr is not None and cNvPr.get("name"):
            filename = cNvPr.get("name", "")
        else:
            filename = media_path.rsplit("/", 1)[-1]
        content_type = _ext_to_content_type(filename)

        shape = ldm.ShapeNode()
        shape.has_image = True
        shape.is_inline = is_inline
        shape.width = width_pt
        shape.height = height_pt
        shape.image_data = ldm.ImageData(
            source_filename=filename,
            content_type=content_type,
            image_bytes=image_bytes,
        )
        if textbox_paragraphs:
            shape.text_box = {"paragraphs": textbox_paragraphs}

        anchor_mode = getattr(self, "_anchor_y_base_mode", "body")
        promote = anchor_mode != "body"
        if (
            not promote
            and anchor_mode == "body"
            and not is_inline
            and anchor is not None
            and getattr(self, "_first_body_page_active", False)
        ):
            posH = anchor.find(f"{WP_NS}positionH")
            if posH is not None and posH.get("relativeFrom", "") == "page":
                promote = True
        if not is_inline and anchor is not None:
            left_mm, top_mm = self._anchor_page_origin_mm(anchor)
            shape.left = left_mm
            shape.top = top_mm
            # Parse wrap type from the anchor element.
            shape.wrap_type = self._parse_wrap_type(anchor)
            if promote:
                if width_pt:
                    shape.width = width_pt / _POINTS_PER_INCH * _MM_PER_INCH  # pt → mm
                if height_pt:
                    shape.height = height_pt / _POINTS_PER_INCH * _MM_PER_INCH
                shape._is_positioned = True
        return shape

    @staticmethod
    def _parse_wrap_type(anchor: ET.Element) -> int:
        """Return the Aspose-compatible WrapType integer from a ``<wp:anchor>``.

        Matches Aspose.Words WrapType enum:
        0=Inline, 1=TopBottom, 2=Square, 3=None, 4=Tight, 5=Through.
        """
        from aspose.words_foss.model.wrap_type import WrapType

        _WRAP_MAP = {
            f"{WP_NS}wrapNone": WrapType.NONE,
            f"{WP_NS}wrapTopAndBottom": WrapType.TOP_BOTTOM,
            f"{WP_NS}wrapSquare": WrapType.SQUARE,
            f"{WP_NS}wrapTight": WrapType.TIGHT,
            f"{WP_NS}wrapThrough": WrapType.THROUGH,
        }
        for child in anchor:
            if child.tag in _WRAP_MAP:
                return _WRAP_MAP[child.tag]
        return WrapType.INLINE

    # ------------------------------------------------------------------
    # Positioned-shape extraction (wp:anchor → wpg:wgp → wps:wsp)
    # ------------------------------------------------------------------

    def _extract_positioned_shapes(
        self, drawing_elem: ET.Element, image_rels: dict[str, str]
    ) -> "list[ldm.ShapeNode]":
        """Return positioned ShapeNodes for an anchored shape drawing.

        Word cover pages and other layouts express their visual elements
        (filled rectangles, captions, big title text) as ``<wp:anchor>``
        blocks.  Two structural shapes are handled here:

        * ``wp:anchor > wpg:wgp`` — a group containing multiple shapes.
          Each ``<wps:wsp>`` child carries its own ``a:xfrm`` and
          optional ``a:solidFill``; children are flattened to absolute
          page coordinates.
        * ``wp:anchor > wps:wsp`` — a single shape without a group
          wrapper.  We position it directly from the anchor's
          ``<wp:positionH/V>`` and size from ``<wp:extent>``.

        Returns an empty list for simple picture anchors (those stay on
        the inline path via :meth:`_build_drawing_shape`).
        """
        from aspose.words_foss import light_document_model as ldm

        anchor = drawing_elem.find(f"{WP_NS}anchor")
        if anchor is None:
            return []

        page_x_mm, page_y_mm = self._anchor_page_origin_mm(anchor)

        group = anchor.find(f".//{WPG_NS}wgp")
        if group is not None:
            shapes: list[ldm.ShapeNode] = []
            self._walk_group(
                group,
                parent_off_mm=(page_x_mm, page_y_mm),
                parent_scale=(1.0, 1.0),
                parent_ch_off_emu=(0, 0),
                image_rels=image_rels,
                out=shapes,
            )
            return shapes

        # Single wps:wsp directly under wp:anchor (no group wrapper).
        # The blip path (image anchor) still goes through the inline
        # renderer — positioned shapes are reserved for fillable shapes
        # and overlaid text boxes.
        wsp = anchor.find(f".//{WPS_NS}wsp")
        if wsp is None:
            return []
        if anchor.find(f".//{A_NS}blip") is not None:
            return []

        extent = anchor.find(f"{WP_NS}extent")
        if extent is None:
            return []
        try:
            width_mm = int(extent.get("cx", "0")) / _EMU_PER_MM
            height_mm = int(extent.get("cy", "0")) / _EMU_PER_MM
        except ValueError:
            return []

        shape = self._make_wsp_shape(
            wsp,
            left=page_x_mm,
            top=page_y_mm,
            width=width_mm,
            height=height_mm,
        )
        return [shape] if shape is not None else []

    def _anchor_page_origin_mm(self, anchor: ET.Element) -> tuple[float, float]:
        """Return the (x, y) page position of a ``<wp:anchor>`` in mm.

        Resolves ``<wp:positionH>`` / ``<wp:positionV>`` ``posOffset``
        values, converting ``margin``- and ``column``-relative offsets by
        adding the section's left / top margin.  ``paragraph``-relative
        vertical offsets are approximated as top-margin-relative, which is
        accurate enough for the cover-page use case where all anchors live
        in the first paragraph on page 1.  When the parser is currently
        walking a header (``_anchor_y_base_mode == "header"``), the
        baseline flips to the header distance so anchored logos land
        inside the header band rather than at the top of the body.
        """
        ps = self._current_page_setup
        left_margin_mm = (
            (ps.left_margin / _POINTS_PER_INCH * _MM_PER_INCH) if ps else _DEFAULT_LEFT_MARGIN_MM
        )
        top_margin_mm = (
            (ps.top_margin / _POINTS_PER_INCH * _MM_PER_INCH) if ps else _DEFAULT_TOP_MARGIN_MM
        )
        header_distance_mm = (
            (ps.header_distance / _POINTS_PER_INCH * _MM_PER_INCH)
            if ps and ps.header_distance
            else top_margin_mm / 2.0
        )
        footer_distance_mm = (
            (ps.footer_distance / _POINTS_PER_INCH * _MM_PER_INCH)
            if ps and ps.footer_distance
            else top_margin_mm / 2.0
        )
        page_height_mm = (
            (ps.page_height / _POINTS_PER_INCH * _MM_PER_INCH)
            if ps and ps.page_height
            else _DEFAULT_PAGE_HEIGHT_MM
        )

        x_emu = 0
        y_emu = 0
        x_rel = "page"
        y_rel = "page"

        posH = anchor.find(f"{WP_NS}positionH")
        if posH is not None:
            x_rel = posH.get("relativeFrom", "page")
            off = posH.find(f"{WP_NS}posOffset")
            if off is not None and off.text:
                x_emu = int(off.text)
        posV = anchor.find(f"{WP_NS}positionV")
        if posV is not None:
            y_rel = posV.get("relativeFrom", "page")
            off = posV.find(f"{WP_NS}posOffset")
            if off is not None and off.text:
                y_emu = int(off.text)

        x_mm = x_emu / _EMU_PER_MM
        y_mm = y_emu / _EMU_PER_MM
        if x_rel in ("margin", "column", "leftMargin", "character"):
            x_mm += left_margin_mm
        if y_rel in ("margin", "paragraph", "line", "topMargin"):
            y_base = top_margin_mm
            mode = getattr(self, "_anchor_y_base_mode", "body")
            if mode == "header":
                y_base = header_distance_mm
            elif mode == "footer":
                y_base = page_height_mm - footer_distance_mm
            y_mm += y_base
        return x_mm, y_mm

    def _walk_group(
        self,
        group: ET.Element,
        *,
        parent_off_mm: tuple[float, float],
        parent_scale: tuple[float, float],
        parent_ch_off_emu: tuple[int, int],
        image_rels: dict[str, str],
        out: "list[ldm.ShapeNode]",
    ) -> None:
        """Recursively flatten a ``<wpg:wgp>`` group into positioned shapes.

        Every ``<wps:wsp>`` child contributes a ShapeNode with absolute
        page coordinates (mm); nested ``<wpg:wgp>`` children are walked
        recursively with an updated transform.
        """
        grp_xfrm = group.find(f"{WPG_NS}grpSpPr/{A_NS}xfrm")
        ch_off_emu = (0, 0)
        scale = parent_scale
        if grp_xfrm is not None:
            ch_off = grp_xfrm.find(f"{A_NS}chOff")
            ch_ext = grp_xfrm.find(f"{A_NS}chExt")
            ext = grp_xfrm.find(f"{A_NS}ext")
            if ch_off is not None:
                ch_off_emu = (
                    int(ch_off.get("x", "0")),
                    int(ch_off.get("y", "0")),
                )
            if ch_ext is not None and ext is not None:
                try:
                    sx = int(ext.get("cx", "0")) / max(1, int(ch_ext.get("cx", "1")))
                    sy = int(ext.get("cy", "0")) / max(1, int(ch_ext.get("cy", "1")))
                    scale = (parent_scale[0] * sx, parent_scale[1] * sy)
                except (ValueError, ZeroDivisionError):
                    pass

        for child in group:
            if child.tag == f"{WPS_NS}wsp":
                self._emit_wsp_shape(
                    child,
                    parent_off_mm=parent_off_mm,
                    parent_scale=scale,
                    parent_ch_off_emu=ch_off_emu,
                    image_rels=image_rels,
                    out=out,
                )
            elif child.tag == f"{WPG_NS}grpSp" or child.tag == f"{WPG_NS}wgp":
                # Nested group: compute its origin as a transformed shape
                # and walk its children in the new coordinate system.
                sub_xfrm = child.find(f"{WPG_NS}grpSpPr/{A_NS}xfrm") or child.find(f"{A_NS}xfrm")
                sub_off_mm = parent_off_mm
                if sub_xfrm is not None:
                    off = sub_xfrm.find(f"{A_NS}off")
                    if off is not None:
                        cx = int(off.get("x", "0"))
                        cy = int(off.get("y", "0"))
                        sub_off_mm = (
                            parent_off_mm[0]
                            + (cx - parent_ch_off_emu[0]) * parent_scale[0] / _EMU_PER_MM,
                            parent_off_mm[1]
                            + (cy - parent_ch_off_emu[1]) * parent_scale[1] / _EMU_PER_MM,
                        )
                self._walk_group(
                    child,
                    parent_off_mm=sub_off_mm,
                    parent_scale=scale,
                    parent_ch_off_emu=(0, 0),
                    image_rels=image_rels,
                    out=out,
                )

    def _emit_wsp_shape(
        self,
        wsp: ET.Element,
        *,
        parent_off_mm: tuple[float, float],
        parent_scale: tuple[float, float],
        parent_ch_off_emu: tuple[int, int],
        image_rels: dict[str, str],
        out: "list[ldm.ShapeNode]",
    ) -> None:
        """Emit one positioned ShapeNode from a grouped ``<wps:wsp>``.

        The shape's a:xfrm provides child-coordinate ``off`` / ``ext``;
        these are projected into absolute page-mm using the accumulated
        transform from the enclosing ``wpg:wgp`` chain.
        """
        spPr = wsp.find(f"{WPS_NS}spPr")
        if spPr is None:
            return
        xfrm = spPr.find(f"{A_NS}xfrm")
        if xfrm is None:
            return
        off = xfrm.find(f"{A_NS}off")
        ext = xfrm.find(f"{A_NS}ext")
        if off is None or ext is None:
            return
        try:
            cx = int(off.get("x", "0"))
            cy = int(off.get("y", "0"))
            cw = int(ext.get("cx", "0"))
            ch = int(ext.get("cy", "0"))
        except ValueError:
            return

        sx, sy = parent_scale
        left_mm = parent_off_mm[0] + (cx - parent_ch_off_emu[0]) * sx / _EMU_PER_MM
        top_mm = parent_off_mm[1] + (cy - parent_ch_off_emu[1]) * sy / _EMU_PER_MM
        width_mm = cw * sx / _EMU_PER_MM
        height_mm = ch * sy / _EMU_PER_MM

        shape = self._make_wsp_shape(
            wsp,
            left=left_mm,
            top=top_mm,
            width=width_mm,
            height=height_mm,
        )
        if shape is not None:
            out.append(shape)

    def _make_wsp_shape(
        self,
        wsp: ET.Element,
        *,
        left: float,
        top: float,
        width: float,
        height: float,
    ) -> "Optional[ldm.ShapeNode]":
        """Build a positioned ShapeNode from a ``<wps:wsp>`` at known coords.

        Populates the standard Aspose.Words ShapeNode fields:
          * ``left`` / ``top`` / ``width`` / ``height`` — absolute page
            position and size in mm;
          * ``shading.background_color`` — solid fill, honouring both
            literal ``a:srgbClr`` and theme ``a:schemeClr`` references;
          * single ``borders`` entry — outline color + width from
            ``a:ln`` (same color resolver);
          * ``vertical_alignment`` — 0/1/2 from the text-box anchor
            attribute (``t`` / ``ctr`` / ``b`` in ``wps:bodyPr``);
          * ``text_box.paragraphs`` — harvested textbox content.

        The runtime-only ``_is_positioned`` flag is set so the PDF
        writer can pick the shape up for absolute rendering.  Returns
        ``None`` for empty shapes that carry nothing worth rendering.
        """
        from aspose.words_foss import light_document_model as ldm

        fill_hex = ""
        line_hex = ""
        line_width_pt = 0.0
        spPr = wsp.find(f"{WPS_NS}spPr")
        if spPr is not None:
            fill_hex = self._resolve_drawingml_fill(spPr.find(f"{A_NS}solidFill"))
            ln = spPr.find(f"{A_NS}ln")
            if ln is not None and ln.find(f"{A_NS}noFill") is None:
                # ``a:noFill`` explicitly suppresses the outline and
                # wins over any accompanying ``a:solidFill`` sibling.
                line_hex = self._resolve_drawingml_fill(ln.find(f"{A_NS}solidFill"))
                w_emu = ln.get("w")
                if w_emu:
                    try:
                        line_width_pt = int(w_emu) / _EMU_PER_PT
                    except ValueError:
                        line_width_pt = 0.0

        # DrawingML ``wps:bodyPr/@anchor`` values: t=top, ctr=center,
        # b=bottom.  Map onto the 0/1/2 convention used by the rest of
        # the LDM (CellFormat.vertical_alignment).
        vertical_alignment = 0
        body_pr = wsp.find(f"{WPS_NS}bodyPr")
        if body_pr is not None:
            vertical_alignment = _BODY_ANCHOR_MAP.get(body_pr.get("anchor", "t"), 0)

        textbox_paragraphs: list[ldm.Paragraph] = []
        txbx = wsp.find(f"{WPS_NS}txbx")
        if txbx is not None:
            for container in txbx.iter(f"{W_NS}txbxContent"):
                for element in self._resolve_body_children(container):
                    if element.tag == f"{W_NS}p":
                        textbox_paragraphs.append(self._build_ldm_paragraph(element))

        if not fill_hex and not line_hex and not textbox_paragraphs:
            return None

        shape = ldm.ShapeNode()
        shape.has_image = False
        shape.is_inline = False
        shape.left = left
        shape.top = top
        shape.width = width
        shape.height = height
        if fill_hex:
            shape.shading = ldm.Shading(background_color=_hex_to_aspose_color(fill_hex))
        if line_hex or line_width_pt > 0:
            shape.borders = [
                ldm.Border(
                    line_style=1 if line_hex and line_width_pt > 0 else 0,
                    line_width=line_width_pt,
                    color=_hex_to_aspose_color(line_hex) if line_hex else COLOR_EMPTY,
                )
            ]
        shape.vertical_alignment = vertical_alignment
        if textbox_paragraphs:
            shape.text_box = {"paragraphs": textbox_paragraphs}
        shape._is_positioned = True
        return shape

    def _resolve_drawingml_fill(self, fill_elem: "Optional[ET.Element]") -> str:
        """Resolve a DrawingML fill element to an ``"RRGGBB"`` hex string.

        Handles the two color primitives found in ``spPr`` fills and
        ``ln`` outlines:

        * ``<a:srgbClr val="RRGGBB"/>`` — literal RGB.
        * ``<a:schemeClr val="accent1"/>`` — resolved via the theme's
          ``clrScheme`` populated in :meth:`_parse_theme`.

        Both primitives may carry lightness / shading modifiers
        (``a:lumMod``, ``a:lumOff``, ``a:shade``, ``a:tint``).  We apply
        them in-order so a shape painted with e.g. ``accent1 + shade 75%``
        lands on the correct darker color instead of the base accent.
        Returns an empty string when the fill is missing or uses an
        unsupported primitive (``gradFill``, ``pattFill``, ``blipFill``).
        """
        if fill_elem is None:
            return ""
        base_hex = ""
        color_child = None
        srgb = fill_elem.find(f"{A_NS}srgbClr")
        scheme = fill_elem.find(f"{A_NS}schemeClr")
        if srgb is not None:
            base_hex = (srgb.get("val") or "").upper()
            color_child = srgb
        elif scheme is not None:
            base_hex = self._resolve_theme_color(scheme.get("val") or "")
            color_child = scheme
        if not base_hex or color_child is None:
            return ""

        try:
            r = int(base_hex[0:2], 16)
            g = int(base_hex[2:4], 16)
            b = int(base_hex[4:6], 16)
        except (ValueError, IndexError):
            return base_hex

        # Apply modifiers in document order.  Each modifier's ``val`` is
        # expressed in 1000ths of a percent (50000 == 50%).
        for mod in color_child:
            tag = mod.tag.split("}", 1)[-1]
            try:
                v = int(mod.get("val", "0")) / _DML_MOD_SCALE
            except ValueError:
                continue
            if tag == "lumMod":
                r, g, b = (int(r * v), int(g * v), int(b * v))
            elif tag == "lumOff":
                # ``lumOff`` brightens toward white by ``val`` of the
                # remaining range.
                r = int(r + (_MAX_COLOR_CHANNEL - r) * v)
                g = int(g + (_MAX_COLOR_CHANNEL - g) * v)
                b = int(b + (_MAX_COLOR_CHANNEL - b) * v)
            elif tag == "shade":
                r, g, b = (int(r * v), int(g * v), int(b * v))
            elif tag == "tint":
                r = int(r + (_MAX_COLOR_CHANNEL - r) * (1.0 - v))
                g = int(g + (_MAX_COLOR_CHANNEL - g) * (1.0 - v))
                b = int(b + (_MAX_COLOR_CHANNEL - b) * (1.0 - v))
        r, g, b = (max(0, min(_MAX_COLOR_CHANNEL, c)) for c in (r, g, b))
        return f"{r:02X}{g:02X}{b:02X}"

    @staticmethod
    def _iter_effective_drawings(run_elem: ET.Element) -> Iterator[ET.Element]:
        """Yield every ``<w:drawing>`` inside *run_elem* at any depth.

        ``<mc:AlternateContent>`` branches are preferred as follows:
        ``<mc:Choice>`` (DrawingML) is descended into, and
        ``<mc:Fallback>`` (VML) is skipped — otherwise cover-page text
        boxes would be processed twice.
        """
        stack: list[ET.Element] = list(run_elem)
        while stack:
            node = stack.pop()
            if node.tag == f"{W_NS}drawing":
                yield node
                continue
            if node.tag == f"{MC_NS}AlternateContent":
                choice = node.find(f"{MC_NS}Choice")
                if choice is not None:
                    stack.extend(choice)
                continue
            stack.extend(node)

    def _harvest_textbox_paragraphs(self, drawing_or_pict: ET.Element) -> list["ldm.Paragraph"]:
        """Collect paragraphs from Word text boxes inside a drawing/picture.

        DrawingML (``wps:txbx``) and legacy VML (``v:textbox``) both wrap
        their contents in ``w:txbxContent``; we flatten every occurrence
        into a list of fully-resolved LDM Paragraphs so the PDF writer can
        render them inline. Without this pass cover-page titles and footer bands are
        silently dropped because the reader only extracts pictures from
        ``<w:drawing>``.
        """
        from aspose.words_foss import light_document_model as ldm  # noqa: F401

        paragraphs: list[ldm.Paragraph] = []
        # ``<w:txbxContent>`` may live at any depth inside the drawing/pict
        # and its paragraphs may themselves be wrapped in content controls
        # (``<w:sdt>``), so we reuse the body resolver to unwrap them.
        for container in drawing_or_pict.iter(f"{W_NS}txbxContent"):
            for element in self._resolve_body_children(container):
                if element.tag == f"{W_NS}p":
                    paragraphs.append(self._build_ldm_paragraph(element))
        return paragraphs

    # The following methods are provided by DocumentReader or LdmBuilderMixin
    # at runtime via multiple inheritance.  They are declared here only for
    # static type checkers and are guarded behind TYPE_CHECKING so they
    # never shadow the real implementations.
    if TYPE_CHECKING:

        def _resolve_body_children(
            self, parent: ET.Element
        ) -> Iterator[ET.Element]: ...

        def _build_ldm_paragraph(
            self,
            p_elem: ET.Element,
            image_rels: "Optional[dict[str, str]]" = None,
        ) -> "ldm.Paragraph": ...

        def _resolve_theme_color(self, scheme_name: str) -> str: ...
