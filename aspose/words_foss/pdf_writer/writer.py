"""PDF writer orchestrator operating on the light document model.

Converts an ldm.Document into a PDF file using fpdf2.  This is a thin
orchestrator that delegates to specialised sub-renderers for paragraphs,
tables, shapes, and formatted runs.
"""


import re
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator, Optional, Union

from fpdf import FPDF
from fpdf.prefs import ViewerPreferences

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss._visible_runs import visible_runs
from aspose.words_foss.pdf_writer.constants import (
    A4_HEIGHT_MM,
    A4_WIDTH_MM,
    COMPLIANCE_TO_VERSION,
    DEFAULT_FONT_NAME,
    DEFAULT_FONT_SIZE_PT,
    DEFAULT_MARGIN_MM,
    PT_TO_MM,
)
from aspose.words_foss.pdf_writer.page_bands import install_page_footer, install_page_header
from aspose.words_foss.pdf_writer.paragraph_renderer import ParagraphRenderer
from aspose.words_foss.pdf_writer.run_renderer import RunRenderer
from aspose.words_foss.pdf_writer.shape_renderer import ShapeRenderer
from aspose.words_foss.pdf_writer.table_renderer import TableRenderer
from aspose.words_foss.pdf_writer.text import (
    apply_caps,
    cell_text,
    extract_link_segments,
    is_pure_page_break,
    is_toc_style,
    plain_text,
    safe_text,
)
from aspose.words_foss.saving import PdfSaveOptions, PdfZoomBehavior

_FIT_HEIGHT = PdfZoomBehavior.FIT_HEIGHT
_FIT_BOX = PdfZoomBehavior.FIT_BOX

_ZOOM_BEHAVIOR_TO_DISPLAY_MODE = {
    PdfZoomBehavior.FIT_PAGE: "fullpage",
    PdfZoomBehavior.FIT_WIDTH: "fullwidth",
}

_OPEN_ACTION_RE = re.compile(rb"/OpenAction\s*\[[^\]]*\]")

_ZOOM_BEHAVIOR_TO_DEST = {
    _FIT_HEIGHT: b"/FitV null",
    _FIT_BOX: b"/FitB",
}


def _remove_open_action(pdf_bytes: bytes) -> bytes:
    match = _OPEN_ACTION_RE.search(pdf_bytes)
    if match is None:
        return pdf_bytes
    return pdf_bytes[: match.start()] + b" " * len(match.group()) + pdf_bytes[match.end() :]


def _replace_open_action(pdf_bytes: bytes, zoom_behavior: str) -> bytes:
    match = _OPEN_ACTION_RE.search(pdf_bytes)
    if match is None:
        return pdf_bytes
    page_ref_match = re.search(rb"\d+ \d+ R", match.group())
    if page_ref_match is None:
        return pdf_bytes
    page_ref = page_ref_match.group()
    dest = _ZOOM_BEHAVIOR_TO_DEST[zoom_behavior]
    new_action = b"/OpenAction [" + page_ref + b" " + dest + b"]"
    old = match.group()
    if len(new_action) <= len(old):
        padded = new_action + b" " * (len(old) - len(new_action))
    else:
        padded = new_action
    return pdf_bytes[: match.start()] + padded + pdf_bytes[match.end() :]


class _ColumnState:
    """Tracks multi-column layout position within a section."""

    __slots__ = (
        "ncols", "col_x", "col_w", "line_between", "current",
        "margin_top", "margin_left", "margin_right",
    )

    def __init__(
        self,
        ncols: int,
        col_x: list[float],
        col_w: list[float],
        line_between: bool,
        current: int,
        margin_top: float,
        margin_left: float,
        margin_right: float,
    ) -> None:
        self.ncols = ncols
        self.col_x = col_x
        self.col_w = col_w
        self.line_between = line_between
        self.current = current
        self.margin_top = margin_top
        self.margin_left = margin_left
        self.margin_right = margin_right


class LdmPdfWriter:
    """Converts a ``light_document_model.Document`` to a PDF file."""

    def __init__(self, options: Optional[PdfSaveOptions] = None):
        self.options = options or PdfSaveOptions()
        # Page layout defaults — overridden by write() from section page_setup
        self._page_margin_left = DEFAULT_MARGIN_MM
        self._page_margin_right = DEFAULT_MARGIN_MM
        self._page_margin_bottom = DEFAULT_MARGIN_MM
        self._page_width = A4_WIDTH_MM
        self._page_height = A4_HEIGHT_MM
        # Internal-link registry: anchor name -> fpdf2 link ID.
        self._anchor_links: dict[str, int] = {}
        # Offset applied to pdf.page_no() when resolving PAGE fields.
        self._page_number_offset = 0
        # Outline generation state
        self._in_table: bool = False
        self._last_outline_level: int = 0
        # Exposed to sub-renderers that need style lookups (e.g. table borders).
        self._doc: Optional[ldm.Document] = None
        self._list_counters: dict[tuple[int, int], int] = {}
        # Shapes already drawn out-of-band (currently anchored-wrapped
        # images placed at the prior column's bottom by the balancer);
        # the paragraph renderer skips re-drawing them.
        self._pre_rendered_shapes: set[int] = set()
        # Per-column last cursor Y in the active multi-column flow;
        # exposed so floating tables can pick the right column's
        # anchor Y instead of the live cursor.
        self._active_col_y: dict[int, float] = {}

        # Sub-renderers (composed, not inherited)
        self._paragraph_renderer = ParagraphRenderer(self)
        self._run_renderer = RunRenderer(self)
        self._shape_renderer = ShapeRenderer(self)
        self._table_renderer = TableRenderer(self)

    def write(self, doc: ldm.Document, output_path: Union[str, Path]) -> None:
        """Convert *doc* to PDF and write to *output_path*."""
        output_path = Path(output_path)

        # Fresh anchor-link state per write
        self._anchor_links = {}
        self._default_tab_stop = doc.default_tab_stop
        self._in_table = False
        self._last_outline_level = 0
        self._doc = doc
        self._list_counters = {}
        self._pre_rendered_shapes = set()

        # Reset page layout to defaults before each write
        self._page_margin_left = DEFAULT_MARGIN_MM
        self._page_margin_right = DEFAULT_MARGIN_MM
        self._page_margin_bottom = DEFAULT_MARGIN_MM
        self._page_width = A4_WIDTH_MM

        # Determine page size from the first section
        page_w_mm = A4_WIDTH_MM
        page_h_mm = A4_HEIGHT_MM
        if doc.sections:
            ps = doc.sections[0].page_setup
            if ps.page_width > 0:
                page_w_mm = ps.page_width * PT_TO_MM
            if ps.page_height > 0:
                page_h_mm = ps.page_height * PT_TO_MM

        pdf = FPDF(unit="mm", format=(page_w_mm, page_h_mm))
        # Apply PDF version from compliance setting
        version = COMPLIANCE_TO_VERSION.get(self.options.compliance)
        if version:
            pdf.pdf_version = version

        # Determine margins from the first section's page setup
        margin_top = DEFAULT_MARGIN_MM
        margin_bottom = DEFAULT_MARGIN_MM
        margin_left = DEFAULT_MARGIN_MM
        margin_right = DEFAULT_MARGIN_MM
        if doc.sections:
            ps = doc.sections[0].page_setup
            if ps.left_margin > 0:
                margin_left = ps.left_margin * PT_TO_MM
            if ps.right_margin > 0:
                margin_right = ps.right_margin * PT_TO_MM
            if ps.top_margin > 0:
                margin_top = ps.top_margin * PT_TO_MM
            if ps.bottom_margin > 0:
                margin_bottom = ps.bottom_margin * PT_TO_MM

        pdf.set_margins(margin_left, margin_top, margin_right)
        pdf.set_auto_page_break(auto=True, margin=margin_bottom)

        # Install per-page header / footer callbacks
        title_pg = bool(
            doc.sections and doc.sections[0].page_setup.different_first_page_header_footer
        )
        header_distance_mm = 0.0
        footer_distance_mm = 0.0
        if doc.sections:
            ps = doc.sections[0].page_setup
            if ps.header_distance > 0:
                header_distance_mm = ps.header_distance * PT_TO_MM
            if ps.footer_distance > 0:
                footer_distance_mm = ps.footer_distance * PT_TO_MM

        # Compute the page-number offset
        page_start = 1
        if doc.sections:
            ps = doc.sections[0].page_setup
            if ps.restart_page_numbering:
                page_start = ps.page_starting_number
        self._page_number_offset = page_start - 1

        # Update writer-level layout
        self._page_margin_left = margin_left
        self._page_margin_right = margin_right
        self._page_margin_bottom = margin_bottom
        self._page_width = page_w_mm
        self._page_height = page_h_mm
        # Full-page margins persist while ``_page_margin_left`` /
        # ``_page_margin_right`` shrink to the active column edges in
        # multi-column flows. Floating tables / shapes need the
        # original page margins to compute their position.
        self._page_full_margin_left = margin_left
        self._page_full_margin_right = margin_right

        install_page_header(pdf, doc, self, title_pg, header_distance_mm)
        install_page_footer(pdf, doc, self, title_pg, footer_distance_mm)

        pdf.add_page()
        pdf.set_font(DEFAULT_FONT_NAME, size=DEFAULT_FONT_SIZE_PT)

        # Render cover-page positioned shapes first
        self._shape_renderer.render_positioned_shapes(pdf, doc)

        # Body
        active_col_state: Optional[_ColumnState] = None
        prev_section: Optional[ldm.Section] = None
        # Per-column "cursor parked here last time" Y; balanced flows
        # bounce between columns and each side has to resume from its
        # own last position instead of the page top.
        col_y: dict[int, float] = {}
        for sec_idx, section in enumerate(doc.sections):
            section_col_state = self._init_column_state(
                section, margin_left, margin_right, margin_top,
            )
            section_start = section.page_setup.section_start

            # Continuous section break inherits the previous flow when column geometry is unchanged.
            continues_flow = (
                active_col_state is not None
                and section_start == 0  # Continuous
                and self._column_layouts_match(prev_section, section)
            )

            if continues_flow:
                col_state = active_col_state
            else:
                if active_col_state is not None:
                    self._restore_full_width(pdf, margin_left, margin_right)
                    self._remove_column_hook(pdf)
                    active_col_state = None
                col_state = section_col_state
                col_y = {}
                if col_state is not None:
                    self._install_column_hook(pdf, col_state, col_y)
                    self._apply_column(pdf, col_state, 0)
                    active_col_state = col_state

            self._active_col_y = col_y
            col_w_mm = (
                col_state.col_w[0] if col_state and col_state.col_w else 0.0
            )

            children = list(section.body.children)
            # Non-final sections carry their ``w:sectPr`` inside an empty
            # paragraph at the very end. Word doesn't render it; we did,
            # which left an extra blank line of space before the next
            # section's first floating element.
            is_last_section = sec_idx == len(doc.sections) - 1
            if not is_last_section and children and isinstance(children[-1], ldm.Paragraph):
                tail = children[-1]
                if not "".join(r.text or "" for r in tail.runs).strip():
                    children = children[:-1]

            # New continuous section: when it won't fit in what's left
            # of the current column, jump back to an earlier column
            # that still has room. Word's balance fills the left
            # column's leftovers before walking off the page.
            if continues_flow and col_state is not None and col_state.current > 0:
                col_y[col_state.current] = pdf.get_y()
                section_h = sum(
                    self._estimate_child_height(c, col_w_mm) for c in children
                )
                page_bottom = self._page_height - self._page_margin_bottom
                # Small fudge for line-h rounding / space-after that
                # might push a marginal-fit section past the bottom.
                safety_mm = DEFAULT_FONT_SIZE_PT * PT_TO_MM
                if pdf.get_y() + section_h + safety_mm > page_bottom:
                    for col_idx in range(col_state.current):
                        y_back = col_y.get(col_idx, col_state.margin_top)
                        if y_back < page_bottom - 1:
                            self._apply_column(pdf, col_state, col_idx)
                            pdf.set_y(y_back)
                            break

            balance_target_y: Optional[float] = None
            if col_state is not None:
                balance_target_y = self._compute_section_balance_target(
                    section, col_state, pdf.get_y(), children,
                )

            for child in children:
                if (
                    col_state is not None
                    and balance_target_y is not None
                    and col_state.current < col_state.ncols - 1
                ):
                    child_h = self._estimate_child_height(child, col_w_mm)
                    if pdf.get_y() + child_h / 2 > balance_target_y:
                        if isinstance(child, ldm.Paragraph):
                            self._pre_draw_anchored_wrapped(pdf, child)
                        col_y[col_state.current] = pdf.get_y()
                        self._force_next_column(pdf, col_state, col_y)

                if isinstance(child, ldm.Paragraph):
                    self._paragraph_renderer.render_paragraph(pdf, child)
                elif isinstance(child, ldm.Table):
                    self._in_table = True
                    self._table_renderer.render_table(pdf, child)
                    self._in_table = False

            if col_state is not None:
                col_y[col_state.current] = pdf.get_y()

            prev_section = section

        if active_col_state is not None:
            self._restore_full_width(pdf, margin_left, margin_right)
            self._remove_column_hook(pdf)

        self._apply_viewer_options(pdf)

        output_path.parent.mkdir(parents=True, exist_ok=True)

        if self.options.zoom_behavior == PdfZoomBehavior.NONE:
            pdf_bytes = pdf.output()
            pdf_bytes = _remove_open_action(pdf_bytes)
            output_path.write_bytes(pdf_bytes)
        elif self.options.zoom_behavior in (_FIT_HEIGHT, _FIT_BOX):
            pdf_bytes = pdf.output()
            pdf_bytes = _replace_open_action(pdf_bytes, self.options.zoom_behavior)
            output_path.write_bytes(pdf_bytes)
        else:
            pdf.output(str(output_path))

    # ------------------------------------------------------------------
    # Viewer-option helpers
    # ------------------------------------------------------------------

    def _apply_viewer_options(self, pdf: FPDF) -> None:
        opts = self.options
        if opts.display_doc_title:
            pdf.viewer_preferences = ViewerPreferences(display_doc_title=True)

        zb = opts.zoom_behavior
        display_mode = _ZOOM_BEHAVIOR_TO_DISPLAY_MODE.get(zb)
        if display_mode is not None:
            pdf.set_display_mode(display_mode)
        elif zb == PdfZoomBehavior.ZOOM_FACTOR:
            pdf.set_display_mode(opts.zoom_factor)

    # ------------------------------------------------------------------
    # Column balancing
    # ------------------------------------------------------------------

    def _compute_section_balance_target(
        self,
        section: ldm.Section,
        col_state: _ColumnState,
        start_y: float,
        children: list,
    ) -> Optional[float]:
        """Pre-compute the Y at which to force a balancing column break.

        Estimates inline-flow height of *children* (already filtered to
        skip section-marker paragraphs) and returns the Y that splits
        the height evenly across ``col_state.ncols`` columns.
        """
        col_w = col_state.col_w[0] if col_state.col_w else 0.0
        if col_w <= 0:
            return None
        total_h = sum(self._estimate_child_height(c, col_w) for c in children)
        if total_h <= 0:
            return None
        per_col = total_h / col_state.ncols
        page_bottom = self._page_height - self._page_margin_bottom
        return min(start_y + per_col, page_bottom)

    def _force_next_column(
        self,
        pdf: FPDF,
        state: _ColumnState,
        col_y: dict[int, float],
    ) -> None:
        next_col = state.current + 1
        if next_col >= state.ncols:
            return
        if state.line_between:
            mid_x = (
                state.col_x[state.current]
                + state.col_w[state.current]
                + state.col_x[next_col]
            ) / 2
            pdf.line(mid_x, state.margin_top, mid_x, self._page_height - self._page_margin_bottom)
        self._apply_column(pdf, state, next_col)
        pdf.set_y(col_y.get(next_col, state.margin_top))

    def _pre_draw_anchored_wrapped(self, pdf: FPDF, para: ldm.Paragraph) -> None:
        """Render a paragraph's anchored-wrapped images at the cursor.

        Used right before a balanced column break so each image stays
        in the prior column, like Word's column-balanced output. Their
        ids land in ``_pre_rendered_shapes`` so the paragraph renderer
        doesn't draw them a second time.
        """
        from aspose.words_foss.model.wrap_type import WrapType

        anchored = [
            item for item in para._children
            if isinstance(item, ldm.ShapeNode)
            and not item._is_positioned
            and item.is_inline is False
            and item.wrap_type != WrapType.NONE
            and item.has_image and item.image_data is not None
        ]
        if not anchored:
            return
        for item in anchored:
            self._pre_rendered_shapes.add(id(item))
        self._shape_renderer.render_anchored_wrapped_shapes(pdf, para)
        # The shape renderer skipped them because of the id() guard, so
        # draw them here directly.
        from io import BytesIO
        for item in anchored:
            img_bytes = item.image_data.image_bytes
            if not img_bytes:
                continue
            w_pt = item.width or 0.0
            h_pt = item.height or 0.0
            w_mm = w_pt * PT_TO_MM
            h_mm = h_pt * PT_TO_MM
            saved_x, saved_y = pdf.get_x(), pdf.get_y()
            if item.relative_horizontal_position == 2:
                x = saved_x + item.horizontal_position
            else:
                x = item.left if item.left > 0 else saved_x
            x = max(x, self._page_margin_left)
            img_bytes = self._shape_renderer.compress_image_bytes(img_bytes)
            pdf.image(BytesIO(img_bytes), x=x, y=saved_y, w=w_mm, h=h_mm)
            pdf.set_xy(saved_x, saved_y + h_mm)

    def _estimate_child_height(
        self, child: Union[ldm.Paragraph, ldm.Table], col_w_mm: float
    ) -> float:
        """Approximate inline-flow height the child consumes (mm)."""
        if isinstance(child, ldm.Paragraph):
            return self._estimate_paragraph_height(child, col_w_mm)
        if isinstance(child, ldm.Table):
            return self._estimate_table_height(child, col_w_mm)
        return 0.0

    def _estimate_paragraph_height(self, para: ldm.Paragraph, col_w_mm: float) -> float:
        from aspose.words_foss.pdf_writer.text import get_dominant_font_size
        from aspose.words_foss.model.wrap_type import WrapType

        pf = para.paragraph_format
        runs = visible_runs(para)
        fs = get_dominant_font_size(runs) or DEFAULT_FONT_SIZE_PT
        line_h = self._paragraph_renderer.line_height_mm(fs, pf)

        wrap_h = 0.0  # anchored, treated as floating (no inline cost)
        inline_shape_h = 0.0
        for item in para._children:
            if not isinstance(item, ldm.ShapeNode):
                continue
            if item._is_positioned or item.wrap_type == WrapType.NONE:
                continue
            if not item.has_image:
                continue
            h_mm = (item.height or 0.0) * PT_TO_MM
            if item.is_inline is False:
                wrap_h = max(wrap_h, h_mm)
            else:
                inline_shape_h = max(inline_shape_h, h_mm)

        text = "".join(r.text or "" for r in runs)
        if not text.strip() and inline_shape_h == 0 and wrap_h == 0:
            # Empty para: real render advances by just line_h.
            return line_h

        char_w_mm = fs * PT_TO_MM * 0.5
        if col_w_mm > char_w_mm:
            chars_per_line = max(1, int(col_w_mm / char_w_mm))
            lines = max(1, (len(text) + chars_per_line - 1) // chars_per_line)
        else:
            lines = 1
        text_h = lines * line_h if text.strip() else 0.0
        space = (pf.space_before + pf.space_after) * PT_TO_MM
        # Anchored-wrapped images reserve their height in our inline
        # flow (we can't do real wrap-around), so count them here.
        return wrap_h + text_h + inline_shape_h + space

    def _estimate_table_height(self, table: ldm.Table, col_w_mm: float) -> float:
        if not table.rows:
            return 0.0
        # Floating tables don't displace the inline cursor for paragraphs
        # in *other* columns, but they do consume vertical space in the
        # column they pin to — count their height so the balancer knows
        # the right column's room is smaller than it looks.
        line_h = DEFAULT_FONT_SIZE_PT * PT_TO_MM * 1.2
        pad_mm = (table.top_padding + table.bottom_padding) * PT_TO_MM
        return len(table.rows) * line_h + pad_mm

    # ------------------------------------------------------------------
    # Multi-column layout helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _column_layouts_match(
        prev: Optional[ldm.Section], curr: ldm.Section
    ) -> bool:
        """Return True when two sections share the same column geometry."""
        if prev is None:
            return False
        prev_tc = prev.page_setup.text_columns
        curr_tc = curr.page_setup.text_columns
        if prev_tc is None or curr_tc is None:
            return prev_tc is None and curr_tc is None
        if prev_tc.count != curr_tc.count:
            return False
        if abs(prev_tc.spacing - curr_tc.spacing) > 0.01:
            return False
        if prev_tc.evenly_spaced != curr_tc.evenly_spaced:
            return False
        return True

    @staticmethod
    def _init_column_state(
        section: ldm.Section,
        margin_left: float,
        margin_right: float,
        margin_top: float,
    ) -> Optional[_ColumnState]:
        """Build column geometry from a section's ``text_columns``."""
        tc = section.page_setup.text_columns
        if tc is None or tc.count <= 1:
            return None
        usable_w = section.page_setup.page_width * PT_TO_MM - margin_left - margin_right
        if usable_w <= 0:
            usable_w = A4_WIDTH_MM - margin_left - margin_right
        gutter_mm = tc.spacing * PT_TO_MM if tc.spacing > 0 else 10.0
        if tc.columns and not tc.evenly_spaced:
            col_x: list[float] = []
            col_w: list[float] = []
            x = margin_left
            for col in tc.columns:
                w = col.width * PT_TO_MM if col.width > 0 else usable_w / tc.count
                col_x.append(x)
                col_w.append(w)
                x += w + (col.space_after * PT_TO_MM if col.space_after > 0 else gutter_mm)
        else:
            total_gutter = gutter_mm * (tc.count - 1)
            cw = (usable_w - total_gutter) / tc.count
            col_x = [margin_left + i * (cw + gutter_mm) for i in range(tc.count)]
            col_w = [cw] * tc.count
        return _ColumnState(
            ncols=tc.count,
            col_x=col_x,
            col_w=col_w,
            line_between=tc.line_between,
            current=0,
            margin_top=margin_top,
            margin_left=margin_left,
            margin_right=margin_right,
        )

    def _apply_column(self, pdf: FPDF, state: _ColumnState, col_idx: int) -> None:
        """Set margins and cursor so rendering flows into column *col_idx*."""
        state.current = col_idx
        left = state.col_x[col_idx]
        right = self._page_width - left - state.col_w[col_idx]
        pdf.set_left_margin(left)
        pdf.set_right_margin(right)
        self._page_margin_left = left
        self._page_margin_right = right
        pdf.set_x(left)

    def _install_column_hook(
        self,
        pdf: FPDF,
        state: _ColumnState,
        col_y: dict[int, float],
    ) -> None:
        """Override ``accept_page_break`` so fpdf2 flows to the next column
        instead of adding a new page when content overflows."""
        writer = self
        # Tiny overshoots are inevitable because empty-paragraph leading is
        # estimated, not measured against the actual line metrics. Tolerate
        # up to one line height of slop so a paragraph that *almost* fits
        # the column doesn't jump to the next one and break the layout.
        overshoot_tolerance = DEFAULT_FONT_SIZE_PT * PT_TO_MM * 1.0

        @property  # type: ignore[misc]
        def _col_accept_page_break(self_pdf: FPDF) -> bool:
            col_y[state.current] = pdf.get_y()
            next_col = state.current + 1
            page_bottom = writer._page_height - writer._page_margin_bottom
            if (
                next_col < state.ncols
                and pdf.get_y() <= page_bottom + overshoot_tolerance
            ):
                # Within slop — let the line render in the current column
                # rather than ripping it across the column boundary.
                return False
            if next_col < state.ncols:
                if state.line_between:
                    mid_x = (
                        state.col_x[state.current]
                        + state.col_w[state.current]
                        + state.col_x[next_col]
                    ) / 2
                    pdf.line(mid_x, state.margin_top, mid_x, page_bottom)
                writer._apply_column(pdf, state, next_col)
                pdf.set_y(col_y.get(next_col, state.margin_top))
                return False  # don't let fpdf2 add a page
            # All columns exhausted — restore full-width margins so the
            # header/footer callbacks render at the correct position.
            writer._restore_full_width(pdf, state.margin_left, state.margin_right)
            return True  # let fpdf2 add the page

        pdf.__class__ = type(
            "_ColumnPDF", (FPDF,),
            {"accept_page_break": _col_accept_page_break},
        )
        # After fpdf2 adds a page, reset to column 0
        orig_add_page = pdf.add_page.__func__  # type: ignore[attr-defined]

        def _add_page_then_col0(*args, **kwargs):  # type: ignore[no-untyped-def]
            orig_add_page(pdf, *args, **kwargs)
            writer._apply_column(pdf, state, 0)

        pdf.add_page = _add_page_then_col0  # type: ignore[method-assign]

    @staticmethod
    def _remove_column_hook(pdf: FPDF) -> None:
        """Remove the column accept_page_break override."""
        pdf.__class__ = FPDF
        pdf.__dict__.pop("add_page", None)

    def _restore_full_width(
        self, pdf: FPDF, margin_left: float, margin_right: float
    ) -> None:
        """Restore page-wide margins after a multi-column section."""
        pdf.set_left_margin(margin_left)
        pdf.set_right_margin(margin_right)
        self._page_margin_left = margin_left
        self._page_margin_right = margin_right

    # ------------------------------------------------------------------
    # Tagged-PDF helpers
    # ------------------------------------------------------------------

    @contextmanager
    def _tag(self, pdf: FPDF, struct_type: str, title: Optional[str] = None) -> Iterator[None]:
        """Wrap content in a PDF structure element when tagging is on."""
        if not self.options.export_document_structure:
            yield
            return
        mcid = pdf.struct_builder.next_mcid_for_page(pdf.page)
        kwargs: dict = {"struct_type": struct_type, "mcid": mcid}
        if title:
            kwargs["title"] = title
        pdf._add_marked_content(**kwargs)
        pdf._out(f"/P <</MCID {mcid}>> BDC")
        yield
        pdf._out("EMC")

    # ------------------------------------------------------------------
    # Link helpers
    # ------------------------------------------------------------------

    def _link_target_for(self, pdf: FPDF, url: Optional[str]) -> Union[int, str]:
        """Return the fpdf2 ``link`` value for a Markdown hyperlink target."""
        if not url:
            return ""
        if url.startswith("#"):
            name = url[1:]
            if not name:
                return ""
            link_id = self._anchor_links.get(name)
            if link_id is None:
                link_id = pdf.add_link()
                self._anchor_links[name] = link_id
            return link_id
        return url

    # ------------------------------------------------------------------
    # Backward-compatible instance method aliases
    # ------------------------------------------------------------------

    def _compress_image_bytes(self, image_bytes: bytes) -> bytes:
        """Delegate to shape renderer (backward compat)."""
        return self._shape_renderer.compress_image_bytes(image_bytes)

    # ------------------------------------------------------------------
    # Backward-compatible static/class method aliases
    # ------------------------------------------------------------------

    _plain_text = staticmethod(plain_text)
    _cell_text = staticmethod(cell_text)
    _safe_text = staticmethod(safe_text)
    _extract_link_segments = staticmethod(extract_link_segments)
    _apply_caps = staticmethod(apply_caps)
    _is_pure_page_break = staticmethod(is_pure_page_break)
    _is_toc_style = staticmethod(is_toc_style)
    _cell_has_borders = staticmethod(TableRenderer._cell_has_borders)
    _line_height_mm = staticmethod(ParagraphRenderer.line_height_mm)
