"""Shape and image rendering for the PDF writer."""


from io import BytesIO
from typing import Optional

from fpdf import FPDF

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss._visible_runs import is_horizontal_rule_shape
from aspose.words_foss.model.wrap_type import WrapType
from aspose.words_foss.model.enums import CellVerticalAlignment
from aspose.words_foss.pdf_writer.color import parse_color
from aspose.words_foss.pdf_writer.constants import (
    DEFAULT_FONT_SIZE_PT,
    DEFAULT_LINE_WIDTH_MM,
    DEFAULT_SHAPE_DIM_PT,
    HORIZONTAL_RULE_RGB,
    LINE_HEIGHT_FACTOR,
    MIN_LINE_WIDTH_MM,
    POST_IMAGE_SPACING_MM,
    PT_TO_MM,
    TEXTBOX_INNER_PAD_MM,
)
from aspose.words_foss.pdf_writer._context import PDFWriterContext
from aspose.words_foss.saving import PdfImageCompression

try:
    from PIL import Image as _PILImage
except ImportError:  # pragma: no cover
    _PILImage = None  # type: ignore[assignment,misc]



class ShapeRenderer:
    """Renders shapes, images, and positioned elements."""

    def __init__(self, writer: PDFWriterContext) -> None:
        self._writer = writer

    def compress_image_bytes(self, image_bytes: bytes) -> bytes:
        """Apply image_compression and jpeg_quality options to raw image bytes."""
        options = self._writer.options
        compression = options.image_compression
        quality = options.jpeg_quality

        if compression == PdfImageCompression.AUTO and quality >= 100:
            return image_bytes

        if _PILImage is None:
            return image_bytes

        try:
            pil_img = _PILImage.open(BytesIO(image_bytes))
        except Exception:
            return image_bytes

        if compression == PdfImageCompression.JPEG or (
            compression == PdfImageCompression.AUTO and quality < 100
        ):
            if pil_img.mode in ("RGBA", "P", "LA"):
                pil_img = pil_img.convert("RGB")
            buf = BytesIO()
            pil_img.save(buf, format="JPEG", quality=max(1, min(quality, 100)))
            return buf.getvalue()

        return image_bytes

    def render_shape(self, pdf: FPDF, shape: ldm.Shape) -> None:
        """Render a Shape: image first (if any), then text-box paragraphs."""
        if is_horizontal_rule_shape(shape):
            self.render_horizontal_rule(pdf, shape)
            return
        if shape.has_image and shape.image_data is not None:
            self.render_image_shape(pdf, shape)

        text_box = shape.text_box or {}
        for p in text_box.get("paragraphs", []) or []:
            self._writer._paragraph_renderer.render_paragraph(pdf, p)

    def render_horizontal_rule(self, pdf: FPDF, shape: ldm.Shape) -> None:
        """Draw the rule the shape stands for across the content width."""
        w = self._writer
        y = pdf.get_y() + DEFAULT_FONT_SIZE_PT * PT_TO_MM / 2
        pdf.set_line_width((shape.height or 1.5) * PT_TO_MM)
        pdf.set_draw_color(*HORIZONTAL_RULE_RGB)
        pdf.line(w._page_margin_left, y, w._page_width - w._page_margin_right, y)
        pdf.set_y(y)

    def render_image_shape(self, pdf: FPDF, shape: ldm.Shape) -> None:
        """Embed an image from a Shape into the PDF, scaling to fit page width."""
        img = shape.image_data
        if img is None or not img.image_bytes:
            return

        image_bytes = self.compress_image_bytes(img.image_bytes)

        w = self._writer
        w_pt = shape.width or DEFAULT_SHAPE_DIM_PT
        h_pt = shape.height or DEFAULT_SHAPE_DIM_PT
        w_mm = w_pt * PT_TO_MM
        h_mm = h_pt * PT_TO_MM

        usable_w = w._page_width - w._page_margin_left - w._page_margin_right
        if w_mm > usable_w:
            scale = usable_w / w_mm
            w_mm = usable_w
            h_mm = h_mm * scale

        if shape._page_left_mm > 0:
            pdf.image(BytesIO(image_bytes), x=shape._page_left_mm, w=w_mm, h=h_mm)
        else:
            pdf.image(BytesIO(image_bytes), w=w_mm, h=h_mm)
        pdf.ln(POST_IMAGE_SPACING_MM)

    def render_floating_images(self, pdf: FPDF, para: ldm.Paragraph) -> None:
        """Draw ``wrapNone`` images at their anchor coordinates."""
        for extra in para._children:
            if not isinstance(extra, ldm.Shape):
                continue
            if not extra.has_image or extra.image_data is None:
                continue
            if extra._is_positioned or extra.wrap_type != WrapType.NONE:
                continue
            img_bytes = extra.image_data.image_bytes
            if not img_bytes:
                continue
            saved_x, saved_y = pdf.get_x(), pdf.get_y()
            w_pt = extra.width or DEFAULT_SHAPE_DIM_PT
            h_pt = extra.height or DEFAULT_SHAPE_DIM_PT
            w_mm = w_pt * PT_TO_MM
            h_mm = h_pt * PT_TO_MM
            img_bytes = self.compress_image_bytes(img_bytes)
            x = extra._page_left_mm if extra._page_left_mm > 0 else pdf.get_x()
            para_offset = extra._page_top_mm - (pdf.t_margin if extra._page_top_mm > pdf.t_margin else 0)
            y = saved_y + max(para_offset, 0)
            pdf.image(BytesIO(img_bytes), x=x, y=y, w=w_mm, h=h_mm)
            pdf.set_xy(saved_x, saved_y)

    def render_anchored_wrapped_shapes(self, pdf: FPDF, para: ldm.Paragraph) -> None:
        """Draw anchored (non-inline) shapes that wrap text around them."""
        writer = self._writer
        for extra in para._children:
            if not isinstance(extra, ldm.Shape):
                continue
            if extra._is_positioned:
                continue
            if extra.is_inline is not False or extra.wrap_type == WrapType.NONE:
                continue
            if not extra.has_image or extra.image_data is None:
                continue
            if id(extra) in writer._pre_rendered_shapes:
                continue  # balancer drew this one out-of-band
            img_bytes = extra.image_data.image_bytes
            if not img_bytes:
                continue
            saved_x, saved_y = pdf.get_x(), pdf.get_y()
            w_pt = extra.width or DEFAULT_SHAPE_DIM_PT
            h_pt = extra.height or DEFAULT_SHAPE_DIM_PT
            w_mm = w_pt * PT_TO_MM
            h_mm = h_pt * PT_TO_MM

            # positionH/@relativeFrom="column" is relative to the live column's left edge.
            if extra.relative_horizontal_position == 2:
                x = saved_x + extra.left
            else:
                x = extra._page_left_mm if extra._page_left_mm > 0 else saved_x
            x = max(x, writer._page_margin_left)
            y = saved_y
            img_bytes = self.compress_image_bytes(img_bytes)
            pdf.image(BytesIO(img_bytes), x=x, y=y, w=w_mm, h=h_mm)
            pdf.set_xy(saved_x, saved_y + h_mm)

    def render_positioned_shapes(self, pdf: FPDF, doc: ldm.Document) -> None:
        """Draw every absolutely-positioned shape on the current (first) page."""
        if not doc.sections:
            return
        body_paragraphs = [
            child
            for section in doc.sections
            for child in section.body.children
            if isinstance(child, ldm.Paragraph)
        ]
        self.render_positioned_shapes_in(pdf, body_paragraphs)

    def render_positioned_shapes_in(
        self,
        pdf: FPDF,
        paragraphs: list[ldm.Paragraph],
        *,
        line_y_override: Optional[float] = None,
    ) -> None:
        """Draw positioned shapes carried by *paragraphs* without moving cursor.

        Skips shapes anchored ``relative_vertical_position == Paragraph``
        (= 2): those need the host paragraph's actual Y at render time,
        and are drawn by :meth:`ParagraphRenderer.render_paragraph`
        right after the paragraph body is laid out.
        """
        if not paragraphs:
            return
        saved_x, saved_y = pdf.get_x(), pdf.get_y()
        for para in paragraphs:
            for extra in para._children:
                if not isinstance(extra, ldm.Shape) or not extra._is_positioned:
                    continue
                if extra.relative_vertical_position == 2:
                    continue
                self.render_positioned_shape(pdf, extra, line_y_override=line_y_override)
        pdf.set_xy(saved_x, saved_y)

    def render_positioned_shape(
        self,
        pdf: FPDF,
        shape: ldm.Shape,
        *,
        line_y_override: Optional[float] = None,
        y_override: Optional[float] = None,
    ) -> None:
        """Draw a positioned shape at its absolute page coordinates.

        When *y_override* is supplied it replaces ``shape._page_top_mm`` — used by
        the paragraph renderer to position shapes whose
        ``relative_vertical_position`` is *Paragraph* (the anchor Y
        depends on where the host paragraph lands at render time).
        """
        x = shape._page_left_mm
        y = shape._page_top_mm if y_override is None else y_override
        w = shape.width or 0.0
        h = shape.height or 0.0

        # Word stores group-relative offsets in EMU; nested ``wpg:wgp``
        # transforms accumulate sub-point rounding (e.g. 9525 EMU ≈ 0.265
        # mm) that leaves a hairline white strip when a cover-page band is
        # supposed to flush against the page edge.  When a fill is nominally
        # page-spanning (within 0.5 mm on the trailing edge) snap BOTH the
        # leading and trailing edges to the page boundary so neither side
        # leaves a strip.
        page_w = self._writer._page_width
        page_h = self._writer._page_height
        edge_eps_mm = 0.5
        if 0.0 < x < edge_eps_mm and x + w >= page_w - edge_eps_mm:
            right = max(x + w, page_w)
            x = 0.0
            w = right
        if 0.0 < y < edge_eps_mm and y + h >= page_h - edge_eps_mm:
            bottom = max(y + h, page_h)
            y = 0.0
            h = bottom

        # Rectangle fill and/or border.
        fill_rgb = parse_color(shape.fill_color)
        border = shape.stroke
        line_rgb = parse_color(border.color) if border else None
        draw_border = bool(line_rgb) and bool(border) and border.line_width > 0

        # Anchored picture
        if shape.has_image and shape.image_data is not None:
            img_bytes = shape.image_data.image_bytes
            if img_bytes and w > 0 and h > 0:
                pdf.image(
                    BytesIO(self.compress_image_bytes(img_bytes)),
                    x=x,
                    y=y,
                    w=w,
                    h=h,
                )
            return

        # Degenerate connector lines
        is_line = draw_border and (w <= 0 or h <= 0) and (w > 0 or h > 0)
        if is_line:
            if line_y_override is not None and h == 0:
                y = line_y_override
            pdf.set_draw_color(*line_rgb)
            pdf.set_line_width(max(border.line_width * PT_TO_MM, MIN_LINE_WIDTH_MM))
            pdf.line(x, y, x + w, y + h)
            pdf.set_draw_color(0, 0, 0)
            pdf.set_line_width(DEFAULT_LINE_WIDTH_MM)
            return

        if (fill_rgb or draw_border) and w > 0 and h > 0:
            style = ""
            if fill_rgb:
                pdf.set_fill_color(*fill_rgb)
                style += "F"
            if draw_border:
                pdf.set_draw_color(*line_rgb)
                pdf.set_line_width(max(border.line_width * PT_TO_MM, MIN_LINE_WIDTH_MM))
                style = "D" + style  # "DF" if both, "D" if border only
            pdf.rect(x, y, w, h, style)
            pdf.set_fill_color(255, 255, 255)
            pdf.set_draw_color(0, 0, 0)
            pdf.set_line_width(DEFAULT_LINE_WIDTH_MM)

        # Overlaid text box
        text_box = shape.text_box or {}
        paragraphs = text_box.get("paragraphs", []) or []
        if paragraphs and w > 0:
            writer = self._writer
            saved_margin_left = writer._page_margin_left
            saved_margin_right = writer._page_margin_right
            saved_width = writer._page_width
            # Per-shape body insets from ``wps:bodyPr/@lIns…bIns`` if the
            # reader captured them; otherwise fall back to the uniform
            # constant so older shape data still renders.
            insets = text_box.get("insets_mm")
            if insets and len(insets) == 4:
                pad_l, pad_t, pad_r, pad_b = insets
            else:
                pad_l = pad_t = pad_r = pad_b = TEXTBOX_INNER_PAD_MM
            pdf.set_margins(
                x + pad_l, y + pad_t, max(0.0, writer._page_width - (x + w - pad_r))
            )
            writer._page_margin_left = x + pad_l
            writer._page_margin_right = max(0.0, writer._page_width - (x + w - pad_r))
            prev_auto = pdf.auto_page_break
            prev_bottom = pdf.b_margin
            pdf.set_auto_page_break(auto=False, margin=0)

            content_h = self.estimate_text_box_height(paragraphs)
            inner_h = max(0.0, h - pad_t - pad_b)
            v_offset = 0.0
            if shape.text_box_anchor == CellVerticalAlignment.CENTER and content_h < inner_h:
                v_offset = (inner_h - content_h) / 2.0
            elif shape.text_box_anchor == CellVerticalAlignment.BOTTOM and content_h < inner_h:
                v_offset = inner_h - content_h

            pdf.set_xy(x + pad_l, y + pad_t + v_offset)
            for p in paragraphs:
                writer._paragraph_renderer.render_paragraph(pdf, p)
            pdf.set_auto_page_break(auto=prev_auto, margin=prev_bottom)
            pdf.set_margins(saved_margin_left, writer._page_margin_bottom, saved_margin_right)
            writer._page_margin_left = saved_margin_left
            writer._page_margin_right = saved_margin_right
            writer._page_width = saved_width

    @staticmethod
    def estimate_text_box_height(paragraphs: list[ldm.Paragraph]) -> float:
        """Approximate the laid-out height of a list of paragraphs in mm."""
        if not paragraphs:
            return 0.0
        total = 0.0
        for p in paragraphs:
            size_pt = DEFAULT_FONT_SIZE_PT
            for run in p.runs:
                if run.font.size > 0:
                    size_pt = run.font.size
                    break
            line_h_mm = size_pt * PT_TO_MM * LINE_HEIGHT_FACTOR
            text = "".join(run.text or "" for run in p.runs)
            lines = max(1, text.count("\n") + 1)
            total += lines * line_h_mm
        return total
