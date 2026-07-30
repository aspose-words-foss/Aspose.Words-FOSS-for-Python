"""
Inline Markdown parsing.

Turns the raw text of a leaf block (paragraph, heading, table cell, ...)
into a flat-but-nestable list of inline ``Block`` nodes: plain text runs,
emphasis (bold/italic/strikethrough/underline), inline code, links,
autolinks, raw inline HTML, footnote references, and hard/soft line
breaks. Delimiter matching (``**bold**``, ``*italic*`` ...) follows
CommonMark's delimiter-stack algorithm (left/right-flanking, the "rule
of 3" for mixed same-character runs, intraword-"_" suppression, and
run-splitting so a longer run degrades to the matched length plus a
leftover marker) -- mirrors the C# reference's ``Delimiter.Link``/
``Cut``/``Split``.
"""

from __future__ import annotations

import html
import re
import unicodedata
from typing import Optional

from .blocks import (
    AutolinkBlock,
    Block,
    BoldInlineBlock,
    FootnoteReferenceBlock,
    HtmlTagBlock,
    InlineCodeBlock,
    ItalicInlineBlock,
    LineBreakBlock,
    LinkTextBlock,
    StrikethroughBlock,
    TextBlock,
    UnderlineBlock,
)

_EMPHASIS_CHARS = "*_~+"
_ASCII_PUNCTUATION = set("!\"#$%&'()*+,-./:;<=>?@[\\]^_`{|}~")
_ESCAPABLE = _ASCII_PUNCTUATION  # CommonMark: any ASCII punctuation character may be backslash-escaped


def _is_punctuation(c: str) -> bool:
    if c in _ASCII_PUNCTUATION:
        return True
    return ord(c) > 0x7F and unicodedata.category(c).startswith("P")


def _is_left_flanking(before: str, after: str) -> bool:
    return not after.isspace() and (not _is_punctuation(after) or before.isspace() or _is_punctuation(before))


def _is_right_flanking(before: str, after: str) -> bool:
    return not before.isspace() and (not _is_punctuation(before) or after.isspace() or _is_punctuation(after))

def unescape_destination(s: str) -> str:
    """Resolve backslash escapes and HTML entities in a link/image destination or title."""
    out = []
    i, n = 0, len(s)
    while i < n:
        ch = s[i]
        if ch == "\\" and i + 1 < n and s[i + 1] in _ESCAPABLE:
            out.append(s[i + 1])
            i += 2
            continue
        if ch == "&":
            match = _ENTITY_RE.match(s, i)
            if match is not None:
                decoded = html.unescape(match.group(0))
                if decoded != match.group(0):
                    out.append(decoded)
                    i = match.end()
                    continue
        out.append(ch)
        i += 1
    return "".join(out)


_AUTOLINK_URI_RE = re.compile(r"^[A-Za-z][A-Za-z0-9+.-]{1,31}:[^\s<>]*$")
_EMAIL_LOCAL_CHARS = set("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.!#$%&'*+/=?^_`{|}~-")


def is_valid_autolink(text: str) -> bool:
    """``AutolinkInlineBlock.IsValid``: a valid absolute URI or e-mail."""
    return bool(_AUTOLINK_URI_RE.match(text)) or _is_valid_autolink_email(text)


def _is_valid_autolink_email(text: str) -> bool:
    """Matches C#'s AutolinkInlineBlock.IsValidEmail (stricter than CommonMark's own email-autolink regex)."""
    at = text.find("@")
    if at == -1:
        return False
    local, domain = text[:at], text[at + 1 :]
    if local and local[0] == ".":
        return False
    for i, c in enumerate(local):
        if c not in _EMAIL_LOCAL_CHARS:
            return False
        if c == "." and i > 0 and local[i - 1] == ".":
            return False
    if not domain or domain[-1] in ".@":
        return False
    for segment in domain.split("."):
        if not (1 <= len(segment) <= 63 and segment[0].isascii() and segment[0].isalnum()):
            return False
        if segment[-1] == "-":
            return False
        if not all(c == "-" or (c.isascii() and c.isalnum()) for c in segment):
            return False
    return True


_HTML_TAG_RE = re.compile(r"<(/?)([A-Za-z][A-Za-z0-9-]*)((?:\s+[^<>]*?)?)(/?)>")
_ENTITY_RE = re.compile(r"&(#[0-9]{1,7}|#[xX][0-9a-fA-F]{1,6}|[A-Za-z][A-Za-z0-9]{1,31});")


class _DelimMark:
    """Placeholder for an unresolved emphasis/strike/underline delimiter run.

    Lives in ``nodes`` alongside real ``Block``s until it either becomes
    the boundary of a matched wrap or degrades to a literal ``TextBlock``
    (see ``_finalize_delims``) -- never returned from ``parse_inline``.
    """

    __slots__ = ("ch", "remaining", "can_open", "can_close")

    def __init__(self, ch: str, count: int, can_open: bool, can_close: bool) -> None:
        self.ch = ch
        self.remaining = count
        self.can_open = can_open
        self.can_close = can_close


def _min_run_len(ch: str) -> int:
    return 2 if ch in "~+" else 1


class _LinkOpenMark:
    """Placeholder for an unresolved "[" / "![" opener, mirrors C#'s LinkTextOpeningDelimiter."""

    __slots__ = ("is_image", "marker", "active", "start")

    def __init__(self, is_image: bool, start: int) -> None:
        self.is_image = is_image
        self.marker = "![" if is_image else "["
        self.active = True
        self.start = start


def _normalize_label(label: str) -> str:
    """Matches block_parser._normalize_label (duplicated locally to avoid a circular import)."""
    return " ".join(label.split()).upper()


def parse_inline(text: str, known_labels: Optional[dict] = None) -> list[Block]:
    """Parse *text* into inline nodes; *known_labels* is the pre-scanned reference-definition dict."""
    # Matches C#'s InlineContainerBlock.GetInlineText, which trims the fully-joined text
    # before any delimiter/hard-break processing sees it -- so trailing spaces on a
    # container's last line (nothing follows them) are insignificant, not a hard break.
    text = text.strip()
    nodes: list = []
    buf: list[str] = []
    i = 0
    n = len(text)

    def flush_text() -> None:
        if buf:
            nodes.append(TextBlock("".join(buf)))
            buf.clear()

    while i < n:
        ch = text[i]

        if ch == "\\" and i + 1 < n and text[i + 1] in _ESCAPABLE:
            buf.append(text[i + 1])
            i += 2
            continue

        if ch == "&":
            match = _ENTITY_RE.match(text, i)
            if match is not None:
                decoded = html.unescape(match.group(0))
                if decoded != match.group(0):
                    buf.append(decoded)
                    i = match.end()
                    continue

        if ch == "\n":
            prev_text = text[:i]
            trailing_backslashes = len(prev_text) - len(prev_text.rstrip("\\"))
            if trailing_backslashes % 2 == 1:
                # An odd trailing run means the last "\" is unescaped (matches C#'s
                # IsEscaped): e.g. "abc\\\n" is one escaped, literal "\" and no break,
                # while "abc\\\\\n" is a literal "\" plus a genuine break-triggering one.
                buf.pop()  # break syntax, not literal text
                hard = True
            else:
                hard = re.search(r"  +$", prev_text) is not None
                while buf and buf[-1] == " ":
                    buf.pop()  # trailing spaces are break syntax, never literal text
            flush_text()
            nodes.append(LineBreakBlock(is_hard=bool(hard)))
            i += 1
            while i < n and text[i] == " ":
                i += 1
            continue

        if ch == "`":
            flush_text()
            run_len = 1
            while i + run_len < n and text[i + run_len] == "`":
                run_len += 1
            new_i = _try_code_span(text, i, run_len, nodes)
            if new_i is not None:
                i = new_i
                continue
            buf.append(ch * run_len)  # a whole backtick run is one atomic token, unmatched or not
            i += run_len
            continue

        if ch == "<":
            new_i = _try_autolink_or_html(text, i, nodes, flush_text)
            if new_i is not None:
                i = new_i
                continue
            buf.append(ch)
            i += 1
            continue

        if ch == "[" and i + 1 < n and text[i + 1] == "^":
            footnote = try_footnote_reference(text, i)
            if footnote is not None:
                flush_text()
                node, new_i = footnote
                nodes.append(node)
                i = new_i
                continue

        if ch == "!" and i + 1 < n and text[i + 1] == "[":
            flush_text()
            nodes.append(_LinkOpenMark(True, i))
            i += 2
            continue

        if ch == "[":
            flush_text()
            nodes.append(_LinkOpenMark(False, i))
            i += 1
            continue

        if ch == "]":
            flush_text()
            i = _resolve_link_closer(text, i, nodes, known_labels)
            continue

        if ch in _EMPHASIS_CHARS:
            run_len = 1
            while i + run_len < n and text[i + run_len] == ch:
                run_len += 1

            before = text[i - 1] if i > 0 else " "
            after = text[i + run_len] if i + run_len < n else " "

            if ch == "_" and before.isalnum() and after.isalnum():  # intraword "_" never opens/closes
                buf.append(ch * run_len)
                i += run_len
                continue

            can_open = _is_left_flanking(before, after)
            can_close = _is_right_flanking(before, after)
            if ch in "~+" and run_len < _min_run_len(ch):
                can_open = can_close = False

            if not can_open and not can_close:
                buf.append(ch * run_len)
                i += run_len
                continue

            flush_text()
            mark = _DelimMark(ch, run_len, can_open, can_close)
            nodes.append(mark)
            if can_close:
                _resolve_closer(nodes, mark)
                if mark.remaining <= 0:
                    nodes.remove(mark)
                elif not mark.can_open or mark.remaining < _min_run_len(ch):
                    nodes[nodes.index(mark)] = TextBlock(mark.ch * mark.remaining)
            i += run_len
            continue

        buf.append(ch)
        i += 1

    flush_text()
    return _finalize_delims(nodes)


def _make_wrapper(ch: str, count: int) -> Block:
    if ch in "*_":
        return ItalicInlineBlock() if count == 1 else BoldInlineBlock()
    if ch == "~":
        return StrikethroughBlock()
    return UnderlineBlock()


def _finalize_delims(items: list) -> list[Block]:
    """Degrade any never-matched ``_DelimMark``s left in *items* to literal text."""
    out: list[Block] = []
    for item in items:
        if isinstance(item, _DelimMark):
            if item.remaining:
                out.append(TextBlock(item.ch * item.remaining))
        elif isinstance(item, _LinkOpenMark):
            out.append(TextBlock(item.marker))
        else:
            out.append(item)
    return out


def _mod3_blocks(opener: _DelimMark, closer: _DelimMark) -> bool:
    """CommonMark rule 9: reject the pairing if it would violate the "multiple of 3" rule.

    Only applies to "*"/"_" (matches C#'s EmphasesDelimiter.CanBeOpeningFor;
    strikethrough/underline derive from EmphasesDelimiterBase, which skips it).
    """
    if opener.ch not in "*_":
        return False
    if not ((opener.can_open and opener.can_close) or (closer.can_open and closer.can_close)):
        return False
    total = opener.remaining + closer.remaining
    if total % 3 != 0:
        return False
    return not (opener.remaining % 3 == 0 and closer.remaining % 3 == 0)


def _resolve_closer(nodes: list, closer: _DelimMark) -> None:
    """Match *closer* against openers already in *nodes*, wrapping and truncating runs (like C#'s Delimiter.Link)."""
    min_len = _min_run_len(closer.ch)
    while closer.remaining >= min_len:
        opener = None
        opener_idx = None
        for idx in range(len(nodes) - 2, -1, -1):
            cand = nodes[idx]
            if isinstance(cand, _LinkOpenMark) and cand.active:
                break  # a still-open link bracket blocks emphasis from reaching past it
            if not isinstance(cand, _DelimMark):
                continue
            if cand.ch != closer.ch or not cand.can_open or cand.remaining < min_len:
                continue
            if _mod3_blocks(cand, closer):
                continue
            opener, opener_idx = cand, idx
            break
        if opener is None:
            break

        use = min(opener.remaining, closer.remaining, 2)
        children = nodes[opener_idx + 1 : -1]
        del nodes[opener_idx + 1 : -1]
        wrapper = _make_wrapper(closer.ch, use)
        for child in _finalize_delims(children):
            wrapper.add_child(child)
        nodes.insert(opener_idx + 1, wrapper)

        opener.remaining -= use
        closer.remaining -= use
        if opener.remaining <= 0:
            del nodes[opener_idx]
        elif not opener.can_open or opener.remaining < min_len:
            nodes[opener_idx] = TextBlock(opener.ch * opener.remaining)


def _try_code_span(text: str, i: int, run_len: int, nodes: list[Block]) -> Optional[int]:
    n = len(text)
    fence = "`" * run_len
    close = text.find(fence, i + run_len)
    while close != -1 and close + run_len < n and text[close + run_len] == "`":
        close = text.find(fence, close + run_len + 1)
    if close == -1:
        return None
    content = text[i + run_len : close].replace("\n", " ")
    if content.startswith(" ") and content.endswith(" ") and content.strip() != "":
        content = content[1:-1]
    nodes.append(InlineCodeBlock(content, delimiter_length=run_len))
    return close + run_len


def _try_autolink_or_html(text: str, i: int, nodes: list[Block], flush_text) -> Optional[int]:
    close = text.find(">", i + 1)
    if close != -1:
        inner = text[i + 1 : close]
        if " " not in inner and "\n" not in inner and inner:
            if _AUTOLINK_URI_RE.match(inner):
                flush_text()
                nodes.append(AutolinkBlock(inner, inner))
                return close + 1
            if _is_valid_autolink_email(inner):
                flush_text()
                nodes.append(AutolinkBlock(inner, f"mailto:{inner}"))
                return close + 1

    match = _HTML_TAG_RE.match(text, i)
    if match:
        flush_text()
        is_closing, tag_name, _attrs, self_close = match.groups()
        nodes.append(
            HtmlTagBlock(
                tag_name,
                match.group(0),
                is_self_closing=bool(self_close),
                is_closing=bool(is_closing),
            )
        )
        return match.end()

    if text.startswith("<!--", i):
        close = text.find("-->", i + 4)
        if close != -1:
            content = text[i + 4 : close]
            if not content.startswith(">") and not content.startswith("->") and not content.endswith("-") and "--" not in content:
                return close + 3

    return None


def _find_matching_bracket(text: str, open_idx: int) -> Optional[int]:
    depth = 0
    i = open_idx
    n = len(text)
    while i < n:
        if text[i] == "\\" and i + 1 < n:
            i += 2
            continue
        if text[i] == "[":
            depth += 1
        elif text[i] == "]":
            depth -= 1
            if depth == 0:
                return i
        i += 1
    return None


def _find_unescaped(text: str, ch: str, start: int) -> int:
    i, n = start, len(text)
    while i < n:
        if text[i] == "\\" and i + 1 < n:
            i += 2
            continue
        if text[i] == ch:
            return i
        i += 1
    return -1


def _parse_angle_bracket_destination(text: str, j: int) -> Optional[tuple[str, int]]:
    """Parse ``<...>`` at *j*; ``None`` if it contains a raw "<"/newline or never closes."""
    n = len(text)
    k = j + 1
    while k < n:
        c = text[k]
        if c == "\\" and k + 1 < n:
            k += 2
            continue
        if c == "\n" or c == "<":
            return None
        if c == ">":
            return text[j + 1 : k], k + 1
        k += 1
    return None


def _parse_link_destination(text: str, i: int) -> Optional[tuple[str, str, int]]:
    """Parse ``(url "title")`` at *i*; ``None`` if malformed (matches C# ``LinkDestinationBlock.IsValid``)."""
    n = len(text)
    j = i + 1
    while j < n and text[j].isspace():
        j += 1
    uri = ""
    if j < n and text[j] == "<":
        result = _parse_angle_bracket_destination(text, j)
        if result is None:
            return None
        uri, j = result
        if j < n and text[j] != ")" and not text[j].isspace():
            return None  # a title must be separated from an angle-bracket URI by whitespace
    else:
        start = j
        depth = 0
        while j < n:
            c = text[j]
            if c == "\\" and j + 1 < n:
                j += 2
                continue
            if depth == 0 and (c.isspace() or c == ")"):
                break
            if c == "(":
                depth += 1
            elif c == ")":
                depth -= 1
            j += 1
        uri = text[start:j]
    while j < n and text[j].isspace():
        j += 1
    title = ""
    if j < n and text[j] in "\"'(":
        quote = ")" if text[j] == "(" else text[j]
        end = _find_unescaped(text, quote, j + 1)
        if end == -1:
            return None
        title = text[j + 1 : end]
        j = end + 1
        while j < n and text[j].isspace():
            j += 1
    if j >= n or text[j] != ")":
        return None
    return unescape_destination(uri), unescape_destination(title), j + 1


def _resolve_link_closer(text: str, i: int, nodes: list, known_labels: Optional[dict]) -> int:
    """Match "]" at *i* against the nearest "[" / "![" opener still in *nodes* (like C#'s LinkTextClosingDelimiter)."""
    opener_idx = None
    for idx in range(len(nodes) - 1, -1, -1):
        if isinstance(nodes[idx], _LinkOpenMark):
            opener_idx = idx
            break
    if opener_idx is None:
        nodes.append(TextBlock("]"))
        return i + 1

    opener = nodes[opener_idx]
    if not opener.active:
        nodes[opener_idx] = TextBlock(opener.marker)
        nodes.append(TextBlock("]"))
        return i + 1

    n = len(text)
    j = i + 1
    label_text = text[opener.start + len(opener.marker) : i]

    uri: Optional[str] = None
    title = ""
    definition_label: Optional[str] = None

    if j < n and text[j] == "(":
        result = _parse_link_destination(text, j)
        if result is not None:
            uri, title, j = result
        else:
            definition_label = label_text  # malformed "(...)"; falls back to shortcut reference
    elif j < n and text[j] == "[":
        end = _find_matching_bracket(text, j)
        if end is not None:
            ref_label = text[j + 1 : end]
            definition_label = ref_label if ref_label else label_text
            j = end + 1
    else:
        definition_label = label_text

    is_image = opener.is_image
    link = LinkTextBlock(
        uri=uri, title=title, is_image=is_image, definition_label=definition_label,
        raw_text=text[opener.start : j],
    )
    for child in _finalize_delims(nodes[opener_idx + 1 :]):
        link.add_child(child)
    del nodes[opener_idx:]
    nodes.append(link)

    will_resolve = uri is not None or (
        definition_label is not None and known_labels is not None
        and _normalize_label(definition_label) in known_labels
    )
    if not is_image and will_resolve:  # links may not contain other links, at any level of nesting
        for node in nodes[:opener_idx]:
            if isinstance(node, _LinkOpenMark) and not node.is_image:
                node.active = False

    return j


_FOOTNOTE_LABEL_CHARS = r"[^\s~`!:&*_<\\\[\]{}]+"  # matches C#'s FootnoteReferenceBlock.gUnallowableSpecialChars
_FOOTNOTE_REF_RE = re.compile(r"\[\^(" + _FOOTNOTE_LABEL_CHARS + r")\]")


def try_footnote_reference(text: str, i: int) -> Optional[tuple[FootnoteReferenceBlock, int]]:
    match = _FOOTNOTE_REF_RE.match(text, i)
    if match is None:
        return None
    return FootnoteReferenceBlock(match.group(1)), match.end()
