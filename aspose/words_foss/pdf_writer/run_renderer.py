"""Formatted-run rendering for the PDF writer.

Handles left-aligned writes, aligned cell rows, highlight rectangles,
and strikethrough lines.
"""


from typing import Optional, Tuple, Union

from fpdf import FPDF

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.pdf_writer.color import parse_color
from aspose.words_foss.pdf_writer.constants import (
    DEFAULT_FONT_SIZE_PT,
    GAP_MIN_SPACES,
    HIGHLIGHT_HEIGHT_RATIO,
    HIGHLIGHT_Y_OFFSET_RATIO,
    HYPERLINK_TEXT_RGB,
    LINE_HEIGHT_FACTOR,
    PT_TO_MM,
    STRIKETHROUGH_Y_RATIO,
)
from aspose.words_foss.pdf_writer.font import apply_run_font
from aspose.words_foss.pdf_writer.text import (
    apply_caps,
    extract_link_segments,
    safe_text,
)
from aspose.words_foss.docx_reader import PAGE_FIELD_SENTINEL
from aspose.words_foss.pdf_writer._context import PDFWriterContext


class RunRenderer:
    """Renders formatted runs (text segments with fonts, colors, links)."""

    def __init__(self, writer: PDFWriterContext) -> None:
        self._writer = writer

    # ------------------------------------------------------------------
    # Public rendering methods
    # ------------------------------------------------------------------

    def render_formatted_runs(
        self,
        pdf: FPDF,
        runs: list[ldm.Run],
        *,
        align: str = "L",
        newline: bool = True,
        line_h_override: Optional[float] = None,
        is_toc: bool = False,
        tab_stops: Optional[ldm.TabStopCollection] = None,
        default_tab_stop: float = 36.0,
        pf: Optional[ldm.ParagraphFormat] = None,
    ) -> None:
        """Render runs with inline formatting: bold, italic, underline, colors, sizes."""
        fs = DEFAULT_FONT_SIZE_PT

        # A ``\t`` between runs marks a Word tab stop — TOC entries put
        # the title before the tab and the page number after.  Route
        # TOC entries through the aligned path so the column-split
        # detector can anchor the trailing runs to the right margin
        # (Word's default TOC right-tab stop).  Ordinary numbered
        # paragraphs like ``4.1<tab>Heading`` must stay on the
        # left-aligned path — the aligned renderer would otherwise
        # split them at the tab and collide the trailing text with
        # the number.
        has_tab = any("\t" in (r.text or "") for r in runs)
        # For non-left alignment, use multi_cell which supports align
        if align != "L" or (has_tab and is_toc):
            self.render_formatted_runs_aligned(
                pdf, runs, align=align, line_h_override=line_h_override, is_toc=is_toc, pf=pf
            )
            return

        for run in runs:
            text = run.text or ""
            if not text:
                continue
            font = run.font
            size = font.size if font.size > 0 else fs
            line_h = (
                line_h_override
                if line_h_override is not None
                else size * PT_TO_MM * LINE_HEIGHT_FACTOR
            )

            # Strip form-feed characters ("\f" = Word page-break marker).
            text = text.replace("\f", "")
            if "\t" in text:
                if tab_stops and tab_stops.tab_stops:
                    apply_run_font(pdf, font, default_size=fs)
                    pieces = text.split("\t")
                    for p_idx, piece in enumerate(pieces):
                        if p_idx > 0:
                            self._advance_to_tab_stop(
                                pdf, tab_stops, default_tab_stop, line_h,
                                upcoming_text=piece,
                            )
                        if piece:
                            resolved = self._resolve_run_text(pdf, piece)
                            if resolved:
                                resolved = apply_caps(resolved, font)
                                pdf.write(h=line_h, text=safe_text(resolved))
                    continue
                text = text.replace("\t", " ")
            text = self._resolve_run_text(pdf, text)
            if not text:
                continue
            text = apply_caps(text, font)
            for chunk, link in extract_link_segments(text):
                if not chunk:
                    continue
                apply_run_font(pdf, font, default_size=fs)
                highlight = parse_color(font.highlight_color)
                if link is not None:
                    # Typical hyperlink convention: blue text.
                    pdf.set_text_color(*HYPERLINK_TEXT_RGB)
                safe = safe_text(chunk)
                link_target = self._writer._link_target_for(pdf, link)
                if highlight:
                    # Use cell-based rendering so the fill rectangle
                    # and text are always perfectly aligned — even
                    # when the run wraps across lines.
                    self._write_with_highlight(pdf, safe, line_h, highlight, link=link_target)
                elif font.strike_through:
                    x_before = pdf.get_x()
                    y_before = pdf.get_y()
                    pdf.write(h=line_h, text=safe, link=link_target)
                    x_after = pdf.get_x()
                    strike_y = y_before + size * PT_TO_MM * STRIKETHROUGH_Y_RATIO
                    pdf.line(x_before, strike_y, x_after, strike_y)
                else:
                    pdf.write(h=line_h, text=safe, link=link_target)

        # Reset color
        pdf.set_text_color(0, 0, 0)
        if newline:
            pdf.ln()

    def _advance_to_tab_stop(
        self,
        pdf: FPDF,
        tab_stops: ldm.TabStopCollection,
        default_tab_stop: float,
        line_h: float,
        upcoming_text: str = "",
    ) -> None:
        """Advance cursor to the next tab stop position, drawing leader fill.

        Handles CENTER/RIGHT/DECIMAL alignment by offsetting the target
        position based on the width of *upcoming_text*.
        """
        w = self._writer
        current_x = pdf.get_x()
        margin_left = w._page_margin_left
        pos_from_margin = current_x - margin_left
        pos_pt = pos_from_margin / PT_TO_MM

        next_tab = tab_stops.after(pos_pt)
        if next_tab is not None:
            target_pt = next_tab.position
        else:
            step = default_tab_stop if default_tab_stop > 0 else 36.0
            target_pt = ((pos_pt // step) + 1) * step

        if next_tab and upcoming_text:
            alignment = next_tab.alignment
            text_w_mm = pdf.get_string_width(safe_text(upcoming_text))
            text_w_pt = text_w_mm / PT_TO_MM
            if alignment == 1:  # CENTER
                target_pt -= text_w_pt / 2
            elif alignment == 2:  # RIGHT
                target_pt -= text_w_pt
            elif alignment == 3:  # DECIMAL
                dot_idx = -1
                for ci, ch in enumerate(upcoming_text):
                    if ch in ".,":
                        dot_idx = ci
                        break
                if dot_idx >= 0:
                    before_w = pdf.get_string_width(safe_text(upcoming_text[:dot_idx])) / PT_TO_MM
                else:
                    before_w = text_w_pt
                target_pt -= before_w

        target_x = margin_left + target_pt * PT_TO_MM
        gap = target_x - current_x
        if gap <= 0:
            return

        if next_tab and next_tab.leader > 0:
            _LEADER_CHARS = {1: ".", 2: "-", 3: "_", 4: "_", 5: "·"}
            ch = _LEADER_CHARS.get(next_tab.leader, ".")
            char_w = pdf.get_string_width(ch)
            if char_w > 0:
                count = int(gap / char_w)
                if count > 0:
                    pdf.write(h=line_h, text=ch * count)

        pdf.set_x(target_x)

    def render_formatted_runs_aligned(
        self,
        pdf: FPDF,
        runs: list[ldm.Run],
        *,
        align: str = "C",
        line_h_override: Optional[float] = None,
        is_toc: bool = False,
        pf: Optional[ldm.ParagraphFormat] = None,
    ) -> None:
        """Render runs with per-run formatting inside an aligned line.

        Computes the total width of all runs, then positions the cursor
        according to *align* before emitting each run with its own font,
        color and style via ``pdf.cell()``.
        """
        fs = DEFAULT_FONT_SIZE_PT
        w = self._writer
        usable_w = w._page_width - w._page_margin_left - w._page_margin_right

        # Pre-compute chunk widths using each run's own font settings.
        segments: list[Tuple[ldm.Run, str, float, float, Optional[str]]] = []
        for run in runs:
            text = self._resolve_run_text(pdf, (run.text or "").replace("\f", ""))
            if not text:
                continue
            font = run.font
            size = font.size if font.size > 0 else fs
            apply_run_font(pdf, font, default_size=fs)
            if not is_toc and "\t" in text:
                text = text.replace("\t", " ")
            text = apply_caps(text, font)
            # Split runs on ``\t`` so each tab becomes its own segment
            pieces = text.split("\t")
            for idx, piece in enumerate(pieces):
                if idx > 0:
                    segments.append((run, "\t", 0.0, size, None))
                if not piece:
                    continue
                for chunk, link in extract_link_segments(piece):
                    if not chunk:
                        continue
                    safe_chunk = safe_text(chunk)
                    chunk_w = pdf.get_string_width(safe_chunk)
                    segments.append((run, safe_chunk, chunk_w, size, link))

        if not segments:
            return

        max_size = max(s for _, _, _, s, _ in segments)
        line_h = (
            line_h_override
            if line_h_override is not None
            else max_size * PT_TO_MM * LINE_HEIGHT_FACTOR
        )

        # Tab-based two-column layout (TOC entries): title on the left,
        # trailing content pinned to the right margin, same baseline.
        tab_idx = (
            next((i for i, (_, t, _, _, _) in enumerate(segments) if t == "\t"), None)
            if is_toc
            else None
        )
        if tab_idx is not None:
            left = [s for s in segments[:tab_idx] if s[1] != "\t"]
            right = [s for s in segments[tab_idx + 1 :] if s[1] != "\t"]
            self._render_segment_row(pdf, left, line_h, fs, at_x=w._page_margin_left)
            if right:
                right_w = sum(sw for _, _, sw, _, _ in right)
                right_x = max(
                    w._page_margin_left,
                    w._page_margin_left + usable_w - right_w,
                )
                pdf.set_y(pdf.get_y() - line_h)  # stay on the same baseline
                self._render_segment_row(pdf, right, line_h, fs, at_x=right_x)
            pdf.set_text_color(0, 0, 0)
            return

        # Recompute visible-only total after tab segments have been ruled out.
        visible = [s for s in segments if s[1] != "\t"]
        if not visible:
            return
        total_w = sum(sw for _, _, sw, _, _ in visible)

        # Word footers frequently encode a two-column "left / right" line
        # as a single right-aligned paragraph with a long run of spaces
        # between the two halves.  Detect the padding idiom and render
        # the halves at the margins instead.
        if align == "R":
            split_idx = self._find_gap_split(visible, usable_w)
            if split_idx is not None:
                left = visible[:split_idx]
                right = visible[split_idx + 1 :]
                self._render_segment_row(pdf, left, line_h, fs, at_x=w._page_margin_left)
                right_w = sum(sw for _, _, sw, _, _ in right)
                right_x = max(
                    w._page_margin_left,
                    w._page_margin_left + usable_w - right_w,
                )
                pdf.set_y(pdf.get_y() - line_h)  # stay on the same baseline
                self._render_segment_row(pdf, right, line_h, fs, at_x=right_x)
                pdf.set_text_color(0, 0, 0)
                return
        segments = visible

        # Position cursor for alignment
        x_start = w._page_margin_left
        if align == "C":
            x_start += (usable_w - total_w) / 2
        elif align == "R":
            x_start += usable_w - total_w
        # "J" and "L" start at the left margin (default)

        # Clamp to left margin when text is wider than the page
        x_start = max(w._page_margin_left, x_start)

        self._render_segment_row(pdf, segments, line_h, fs, at_x=x_start)
        pdf.set_text_color(0, 0, 0)

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _render_segment_row(
        self,
        pdf: FPDF,
        segments: list[Tuple[ldm.Run, str, float, float, Optional[str]]],
        line_h: float,
        fs: float,
        *,
        at_x: float,
    ) -> None:
        """Emit *segments* on a single line starting at *at_x*."""
        pdf.set_x(at_x)
        for run, safe, seg_w, size, link in segments:
            apply_run_font(pdf, run.font, default_size=fs)
            highlight = parse_color(run.font.highlight_color)
            if highlight:
                self._draw_highlight(pdf, safe, line_h, highlight)
            if link is not None:
                pdf.set_text_color(*HYPERLINK_TEXT_RGB)
            link_target = self._writer._link_target_for(pdf, link)
            if run.font.strike_through:
                x_before = pdf.get_x()
                y_before = pdf.get_y()
                pdf.cell(w=seg_w, h=line_h, text=safe, link=link_target)
                x_after = pdf.get_x()
                strike_y = y_before + size * PT_TO_MM * STRIKETHROUGH_Y_RATIO
                pdf.line(x_before, strike_y, x_after, strike_y)
            else:
                pdf.cell(w=seg_w, h=line_h, text=safe, link=link_target)
        pdf.ln(line_h)

    @staticmethod
    def _find_gap_split(
        segments: list[Tuple[ldm.Run, str, float, float, Optional[str]]],
        usable_w: float,
    ) -> Optional[int]:
        """Return the index of the whitespace segment that splits a right-
        aligned line into left / right halves, or ``None`` if no such gap
        exists.
        """
        total_w = sum(w for _, _, w, _, _ in segments)
        if total_w <= usable_w:
            return None
        best = None
        best_w = 0.0
        for i, (_, text, w, _, _) in enumerate(segments):
            if len(text) < GAP_MIN_SPACES or not text.isspace():
                continue
            if i == 0 or i == len(segments) - 1:
                continue
            if w > best_w:
                best = i
                best_w = w
        return best

    @staticmethod
    def _draw_highlight(
        pdf: FPDF,
        text: str,
        line_h: float,
        rgb: Tuple[int, int, int],
    ) -> None:
        """Paint a single highlight rectangle behind *text*."""
        if not text:
            return
        w = pdf.get_string_width(text)
        if w <= 0:
            return
        # ``pdf.cell`` renders text shifted right by ``c_margin`` (1 mm by
        # default).  Align the fill rectangle with that text origin so the
        # highlight tracks the glyphs instead of leading them.
        x = pdf.get_x() + pdf.c_margin
        y = pdf.get_y()
        pdf.set_fill_color(*rgb)
        pdf.rect(x, y + line_h * HIGHLIGHT_Y_OFFSET_RATIO, w, line_h * HIGHLIGHT_HEIGHT_RATIO, "F")

    def _write_with_highlight(
        self,
        pdf: FPDF,
        text: str,
        line_h: float,
        rgb: Tuple[int, int, int],
        *,
        link: Union[int, str] = "",
    ) -> None:
        """Write *text* with a filled highlight background.

        Paints the fill rectangle separately and writes the text on top so
        the rectangle stays aligned with the glyphs.  ``pdf.cell`` offsets
        text by ``c_margin`` (1 mm by default); using ``fill=True`` would
        anchor the fill at the cell's left edge instead, producing a
        half-character drift between text and highlight.
        """
        w = self._writer
        right_edge = w._page_width - w._page_margin_right
        pdf.set_fill_color(*rgb)

        while text:
            x = pdf.get_x()
            avail = right_edge - x
            text_w = pdf.get_string_width(text)

            if text_w <= avail:
                self._fill_text_rect(pdf, x, pdf.get_y(), text_w, line_h)
                pdf.cell(w=text_w, h=line_h, text=text, link=link)
                break

            # Find the last space-separated word that fits.
            words = text.split(" ")
            fit_count = 0
            for i in range(len(words)):
                test_str = " ".join(words[: i + 1])
                if pdf.get_string_width(test_str) > avail:
                    break
                fit_count = i + 1

            if fit_count == 0:
                fit_count = 1

            first = " ".join(words[:fit_count])
            rest = " ".join(words[fit_count:])

            first_w = pdf.get_string_width(first)
            self._fill_text_rect(pdf, x, pdf.get_y(), first_w, line_h)
            pdf.cell(w=first_w, h=line_h, text=first, link=link)

            # Advance to the next line at the left margin.
            pdf.ln(line_h)
            pdf.set_x(w._page_margin_left)

            text = rest

    @staticmethod
    def _fill_text_rect(pdf: FPDF, x: float, y: float, w: float, h: float) -> None:
        """Paint the highlight rectangle behind a cell of width *w*.

        Shifts by ``pdf.c_margin`` to match the text origin used by
        ``pdf.cell``.
        """
        if w <= 0:
            return
        pdf.rect(x + pdf.c_margin, y, w, h, "F")

    def _resolve_run_text(self, pdf: FPDF, text: str) -> str:
        """Resolve reader-emitted sentinels to their live page-local values."""
        if PAGE_FIELD_SENTINEL in text:
            logical = pdf.page_no() + self._writer._page_number_offset
            text = text.replace(PAGE_FIELD_SENTINEL, str(logical))
        return text
