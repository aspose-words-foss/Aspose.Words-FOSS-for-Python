"""Text utility functions for the PDF writer."""

from __future__ import annotations

from typing import Optional, Tuple

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.pdf_writer.constants import INLINE_LINK_RE

# Map of common Unicode punctuation and symbols to latin-1 safe equivalents.
# The built-in PDF fonts (Helvetica/Times/Courier) only support the latin-1
# character set, so any code point outside that range becomes "?" unless we
# transliterate it here first.  This preserves typographic intent (curly
# quotes, dashes, ellipses) without requiring Unicode font embedding.
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
    "\u2022": "-",  # BULLET
    "\u2023": "-",  # TRIANGULAR BULLET
    "\u2043": "-",  # HYPHEN BULLET
    "\u2009": " ",  # THIN SPACE
    "\u200a": " ",  # HAIR SPACE
    "\u200b": "",  # ZERO WIDTH SPACE
    "\u00a0": " ",  # NO-BREAK SPACE -> regular space (also in latin-1 but normalise)
    "\u2011": "-",  # NON-BREAKING HYPHEN
    "\u2212": "-",  # MINUS SIGN
    "\u2122": "(TM)",  # TRADE MARK SIGN
}


def safe_text(text: str) -> str:
    """Transliterate text so it survives latin-1-only built-in PDF fonts.

    Smart quotes, dashes, ellipses and other typographic characters are
    replaced with ASCII equivalents; anything still outside latin-1 is
    replaced with "?" as a last resort.
    """
    if not text:
        return text
    for src, dst in _LATIN1_FALLBACK.items():
        if src in text:
            text = text.replace(src, dst)
    return text.encode("latin-1", errors="replace").decode("latin-1")


def extract_link_segments(text: str) -> list[Tuple[str, Optional[str]]]:
    """Split ``text`` into ``(chunk, url)`` segments around Markdown links.

    Plain text chunks get ``url=None``; linked chunks get a target
    URL.  Internal-anchor refs like ``#_Toc123`` come back verbatim
    (with the leading ``#``); the paragraph renderer converts them
    to fpdf2 link IDs lazily through :meth:`_link_target_for`.
    """
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
    """Uppercase *text* when the run requests ``w:caps`` / ``w:smallCaps``.

    Word's ``<w:caps/>`` / ``<w:smallCaps/>`` are display-only flags:
    the stored text keeps its original case but renders in all
    uppercase.  Applied at render time (not parse time) so the LDM
    continues to hold the literal source string.  Small caps are
    transliterated to full caps here -- rendering lowercase as a
    smaller uppercase glyph would need font-size shaping that the
    built-in PDF fonts don't support.
    """
    if text and (font.all_caps or font.small_caps):
        return text.upper()
    return text


def plain_text(para: ldm.Paragraph) -> str:
    """Return paragraph text with Markdown link syntax stripped.

    Per-run ``w:caps`` / ``w:smallCaps`` are applied here so heading
    and code-block paths (which render via a single joined string)
    honour the display-case flags without needing per-run font
    state.
    """
    return "".join(
        apply_caps(chunk, run.font)
        for run in para.runs
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
        for run in para.runs:
            for chunk, _ in extract_link_segments(run.text or ""):
                pieces.append(apply_caps(chunk, run.font))
        line = "".join(pieces)
        if line:
            lines.append(line)
    return "\n".join(lines)


def is_pure_page_break(para: ldm.Paragraph) -> bool:
    """Return True when *para* should produce a standalone page break.

    Matches a paragraph whose only visible content is one or more
    form-feed characters and which carries no inline images.  Word
    sometimes emits a real ``<w:br w:type='page'/>`` in a paragraph
    that also hosts a picture;
    those are intentionally excluded so the image stays on the
    current page.
    """
    for item in para.inline_extras or ():
        if isinstance(item, ldm.ShapeNode) and item.has_image:
            return False
    visible = False
    saw_form_feed = False
    for run in para.runs:
        text = run.text or ""
        for ch in text:
            if ch == "\f":
                saw_form_feed = True
            elif not ch.isspace():
                visible = True
    return saw_form_feed and not visible


def is_toc_style(style_name: str) -> bool:
    """Return True when *style_name* matches a Word Table of Contents style.

    Used to gate the tab-based two-column layout path: Word TOC entries
    encode "title <tab> page-number" lines that should render with the
    number pinned to the right margin, while ordinary numbered
    paragraphs (e.g. ``4.1<tab>Heading``) must not be split at the
    tab.
    """
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
    from aspose.words_foss.pdf_writer.color import parse_color

    for run in runs:
        rgb = parse_color(run.font.color)
        if rgb:
            return rgb
    return None
