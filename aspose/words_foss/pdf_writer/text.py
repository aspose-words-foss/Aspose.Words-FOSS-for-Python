"""Text utility functions for the PDF writer."""


from typing import Optional, Tuple

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss._visible_runs import visible_runs
from aspose.words_foss.pdf_writer.constants import INLINE_LINK_RE
from aspose.words_foss.pdf_writer.color import parse_color

# Built-in PDF fonts are latin-1 only; transliterate common typographic
# punctuation before encoding so it survives.
_LATIN1_FALLBACK = {
    "\u2018": "'",  # LEFT SINGLE QUOTATION MARK
    "\u2019": "'",  # RIGHT SINGLE QUOTATION MARK / apostrophe
    "\u201a": ",",  # SINGLE LOW-9 QUOTATION MARK
    "\u201b": "'",  # SINGLE HIGH-REVERSED-9 QUOTATION MARK
    "\u201c": '"',  # LEFT DOUBLE QUOTATION MARK
    "\u201d": '"',  # RIGHT DOUBLE QUOTATION MARK
    "\u201e": ",,",  # DOUBLE LOW-9 QUOTATION MARK
    "\u2013": "-",  # EN DASH
    "\u2014": "--",  # EM DASH
    "\u2015": "--",  # HORIZONTAL BAR
    "\u2026": "...",  # HORIZONTAL ELLIPSIS
    # Bullet glyphs absent from Latin-1 collapse to MIDDLE DOT.
    "\u2022": "\u00b7",  # BULLET
    "\u2023": "\u00b7",  # TRIANGULAR BULLET
    "\u2043": "\u00b7",  # HYPHEN BULLET
    "\uf0b7": "\u00b7",  # Symbol-font bullet (Word lvlText)
    "\uf0a7": "\u00b7",  # Wingdings diamond (sub-bullet)
    "\u2009": " ",  # THIN SPACE
    "\u200a": " ",  # HAIR SPACE
    "\u200b": "",  # ZERO WIDTH SPACE
    "\u00a0": " ",  # NO-BREAK SPACE -> regular space (also in latin-1 but normalise)
    "\u2011": "-",  # NON-BREAKING HYPHEN
    "\u2212": "-",  # MINUS SIGN
    "\u2122": "(TM)",  # TRADE MARK SIGN
}


def safe_text(text: str) -> str:
    """Transliterate so *text* survives the latin-1 built-in PDF fonts."""
    if not text:
        return text
    for src, dst in _LATIN1_FALLBACK.items():
        if src in text:
            text = text.replace(src, dst)
    return text.encode("latin-1", errors="replace").decode("latin-1")


def extract_link_segments(text: str) -> list[Tuple[str, Optional[str]]]:
    """Split ``text`` into ``(chunk, url)`` segments around Markdown links."""
    segments: list[Tuple[str, Optional[str]]] = []
    idx = 0
    for m in INLINE_LINK_RE.finditer(text):
        if m.start() > idx:
            segments.append((text[idx : m.start()], None))
        display, url = m.group(1), m.group(2)
        segments.append((display, url))
        idx = m.end()
    if idx < len(text):
        segments.append((text[idx:], None))
    return segments or [(text, None)]


def apply_caps(text: str, font: ldm.Font) -> str:
    """Uppercase *text* when the run requests ``w:caps`` / ``w:smallCaps``."""
    if text and (font.all_caps or font.small_caps):
        return text.upper()
    return text


def plain_text(para: ldm.Paragraph) -> str:
    """Return paragraph text with Markdown link syntax stripped and ``w:caps`` applied."""
    return "".join(
        apply_caps(chunk, run.font)
        for run in visible_runs(para)
        for chunk, _ in extract_link_segments(run.text or "")
    )


def cell_text(cell: ldm.Cell) -> str:
    """Flatten a cell's paragraphs to plain text.

    Paragraphs are joined with newlines so ``multi_cell`` renders each
    one on its own line (important for layout tables that use a single
    cell to hold multi-line blocks like address blocks).  Markdown link
    syntax is stripped so the ``[text](url)`` form from the readers does
    not leak into the cell.  Per-run ``w:caps`` is honoured via
    :func:`apply_caps`.
    """
    lines: list[str] = []
    for para in cell.paragraphs:
        pieces: list[str] = []
        for run in visible_runs(para):
            for chunk, _ in extract_link_segments(run.text or ""):
                pieces.append(apply_caps(chunk, run.font))
        line = "".join(pieces)
        if line:
            lines.append(line)
    return "\n".join(lines)


def is_pure_page_break(para: ldm.Paragraph) -> bool:
    """True when the paragraph contains only form-feeds and no images."""
    for item in para._children or ():
        if isinstance(item, ldm.ShapeNode) and item.has_image:
            return False
    visible = False
    saw_form_feed = False
    for run in visible_runs(para):
        text = run.text or ""
        for ch in text:
            if ch == "\f":
                saw_form_feed = True
            elif not ch.isspace():
                visible = True
    return saw_form_feed and not visible


def is_toc_style(style_name: str) -> bool:
    """True when *style_name* names a Word Table-of-Contents style."""
    if not style_name:
        return False
    lowered = style_name.lower()
    return lowered.startswith("toc") or "toc " in lowered or "table of contents" in lowered


def get_dominant_font_size(runs: list[ldm.Run]) -> float:
    """Get the font size from the first run that has a non-zero size."""
    for run in runs:
        if run.font.size > 0:
            return run.font.size
    return 0.0


def get_dominant_color(runs: list[ldm.Run]) -> Optional[Tuple[int, int, int]]:
    """Get the color from the first run that has a parseable color."""

    for run in runs:
        rgb = parse_color(run.font.color)
        if rgb:
            return rgb
    return None
