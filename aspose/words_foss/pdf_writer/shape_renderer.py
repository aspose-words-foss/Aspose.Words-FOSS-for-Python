"""Shape and image rendering for the PDF writer."""

from __future__ import annotations

from io import BytesIO
from typing import TYPE_CHECKING, Optional

from fpdf import FPDF

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.model.wrap_type import WrapType
from aspose.words_foss.model.enums import CellVerticalAlignment
from aspose.words_foss.pdf_writer.color import parse_color
from aspose.words_foss.pdf_writer.constants import (
    DEFAULT_FONT_SIZE_PT,
    DEFAULT_LINE_WIDTH_MM,
    DEFAULT_SHAPE_DIM_PT,
    LINE_HEIGHT_FACTOR,
    MIN_LINE_WIDTH_MM,
    POST_IMAGE_SPACING_MM,
    PT_TO_MM,
    TEXTBOX_INNER_PAD_MM,
)
from aspose.words_foss.saving import PdfImageCompression

try:
    from PIL import Image as _PILImage
except ImportError:  # pragma: no cover
    _PILImage = None  # type: ignore[assignment,misc]

if TYPE_CHECKING:
    from aspose.words_foss.pdf_writer.renderer import LdmPdfWriter


class ShapeRenderer:
    """Renders shapes, images, and positioned elements."""

    def __init__(self, writer: LdmPdfWriter) -> None:
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

    def render_shape(self, pdf: FPDF, shape: ldm.ShapeNode) -> None:
        """Render a ShapeNode: image first (if any), then text-box paragraphs."""
        if shape.has_image and shape.image_data is not None:
            self.render_image_shape(pdf, shape)

        text_box = shape.text_box or {}
        for p in text_box.get("paragraphs", []) or []:
            self._writer._paragraph_renderer.render_paragraph(pdf, p)

    def render_image_shape(self, pdf: FPDF, shape: ldm.ShapeNode) -> None:
        """Embed an image from a ShapeNode into the PDF, scaling to fit page width."""
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

        if shape.left > 0:
            pdf.image(BytesIO(image_bytes), x=shape.left, w=w_mm, h=h_mm)
        else:
            pdf.image(BytesIO(image_bytes), w=w_mm, h=h_mm)
        pdf.ln(POST_IMAGE_SPACING_MM)

    def render_floating_images(self, pdf: FPDF, para: ldm.Paragraph) -> None:
        """Draw ``wrapNone`` images at their anchor coordinates."""
        for extra in para.inline_extras:
            if not isinstance(extra, ldm.ShapeNode):
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
            x = extra.left if extra.left > 0 else pdf.get_x()
            para_offset = extra.top - (pdf.t_margin if extra.top > pdf.t_margin else 0)
            y = saved_y + max(para_offset, 0)
            pdf.image(BytesIO(img_bytes), x=x, y=y, w=w_mm, h=h_mm)
            pdf.set_xy(saved_x, saved_y)

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
        """Draw positioned shapes carried by *paragraphs* without moving cursor."""
        if not paragraphs:
            return
        saved_x, saved_y = pdf.get_x(), pdf.get_y()
        for para in paragraphs:
            for extra in para.inline_extras:
                if isinstance(extra, ldm.ShapeNode) and extra._is_positioned:
                    self.render_positioned_shape(pdf, extra, line_y_override=line_y_override)
        pdf.set_xy(saved_x, saved_y)

    def render_positioned_shape(
        self,
        pdf: FPDF,
        shape: ldm.ShapeNode,
        *,
        line_y_override: Optional[float] = None,
    ) -> None:
        """Draw a positioned shape at its absolute page coordinates."""
        x = shape.left
        y = shape.top
        w = shape.width or 0.0
        h = shape.height or 0.0

        # Rectangle fill and/or border.
        fill_rgb = parse_color(shape.shading.background_color)
        border = shape.borders[0] if shape.borders else None
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
            pad = TEXTBOX_INNER_PAD_MM
            pdf.set_margins(x + pad, y + pad, max(0.0, writer._page_width - (x + w - pad)))
            writer._page_margin_left = x + pad
            writer._page_margin_right = max(0.0, writer._page_width - (x + w - pad))
            prev_auto = pdf.auto_page_break
            prev_bottom = pdf.b_margin
            pdf.set_auto_page_break(auto=False, margin=0)

            content_h = self.estimate_text_box_height(paragraphs)
            inner_h = max(0.0, h - 2 * pad)
            v_offset = 0.0
            if shape.vertical_alignment == CellVerticalAlignment.CENTER and content_h < inner_h:
                v_offset = (inner_h - content_h) / 2.0
            elif shape.vertical_alignment == CellVerticalAlignment.BOTTOM and content_h < inner_h:
                v_offset = inner_h - content_h

            pdf.set_xy(x + pad, y + pad + v_offset)
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
