"""
Block-level Markdown parsing.

``parse_document`` is the FOSS equivalent of feeding a Markdown string
through Markdig: it turns raw text into the ``Block`` tree defined in
``blocks.py`` (headings, paragraphs, block quotes, lists, fenced/indented
code, tables, footnote definitions, raw HTML blocks). Leaf-block text is
handed off to ``inline_parser.parse_inline`` for emphasis/links/code-span
parsing. This is a pragmatic, well-tested subset of CommonMark + GFM
rather than a full spec-conformant implementation — it covers the
constructs Aspose.Words' ``MarkdownReaderContext`` (added in a later step)
needs to drive. Indentation width (list/quote markers, indented code)
follows C#'s ``MarkdownUtil.GetLength``: a tab is flat +4 regardless of
column position, not CommonMark's column-based tab-stop rounding.
"""

from __future__ import annotations

import re
from typing import Optional

from .blocks import (
    HTML_BLOCK_TAG_NAMES,
    AtxHeadingBlock,
    Block,
    BulletListItemBlock,
    CellBlock,
    DocumentBlock,
    FencedCodeBlock,
    FootnoteDefinitionBlock,
    HorizontalRuleBlock,
    HtmlTagBlock,
    IndentedCodeBlock,
    LinkTextBlock,
    ListBlock,
    ListItemBlock,
    ListMarker,
    OrderedListItemBlock,
    ParagraphBlock,
    QuoteBlock,
    RowBlock,
    SetextHeadingBlock,
    TableBlock,
    TextBlock,
)
from ._reader_constants import _SELF_CONTAINED_HTML_RE
from .inline_parser import _FOOTNOTE_LABEL_CHARS, _HTML_TAG_RE, parse_inline, unescape_destination

_BULLET_MARKER_CHARS = {"-": ListMarker.DASH, "*": ListMarker.STAR, "+": ListMarker.PLUS}

_ATX_RE = re.compile(r"^ {0,3}(#{1,6})(?:[ \t]+(.*))?$")
_HR_RE = re.compile(r"^ {0,3}([-*_])(?: *\1){2,} *$")
_SETEXT_RE = re.compile(r"^ {0,3}(=+|-+)[ \t]*$")
_BLOCKQUOTE_RE = re.compile(r"^ {0,3}>[ \t]?(.*)$")
_BULLET_ITEM_RE = re.compile(r"^( {0,3})([-*+])(?:([ \t]+)(.*)|())$")
_ORDERED_ITEM_RE = re.compile(r"^( {0,3})([0-9]{1,9})([.)])(?:([ \t]+)(.*)|())$")
_FENCE_RE = re.compile(r"^ {0,3}(`{3,}|~{3,})[ \t]*(.*)$")
_TABLE_DELIM_CELL_RE = re.compile(r"^:?-+:?$")
_FOOTNOTE_DEF_RE = re.compile(r"^ {0,3}\[\^(" + _FOOTNOTE_LABEL_CHARS + r")\]:[ \t]*(.*)$")
_REF_DEF_RE = re.compile(
    r'^ {0,3}\[((?:\\.|[^\]\\])+)\]:[ \t]*(?:<([^<>]*)>|([^\s>]+))'
    r'(?:[ \t]+(?:"((?:\\.|[^"\\])*)"|\'((?:\\.|[^\'\\])*)\'|\(((?:\\.|[^)\\])*)\)))?[ \t]*$'
)
_REF_DEF_LOOSE_START_RE = re.compile(r"^ {0,3}\[")
_HTML_BLOCK_START_RE = re.compile(r"^ {0,3}</?([A-Za-z][A-Za-z0-9-]*)")


def _fence_match(line: str) -> Optional[re.Match]:
    """Match *line* as a fenced-code opener, rejecting the one case the
    regex alone can't express: a backtick fence whose info string itself
    contains a backtick (CommonMark treats that as not a valid fence)."""
    match = _FENCE_RE.match(line)
    if match is None:
        return None
    fence, info = match.group(1), match.group(2)
    if fence[0] == "`" and "`" in info:
        return None
    return match


def parse_document(text: str, preserve_empty_lines: bool = False) -> DocumentBlock:
    """Parse *text* into a ``DocumentBlock`` tree.

    With *preserve_empty_lines*, each blank line between blocks becomes an
    empty ``ParagraphBlock`` instead of being silently skipped (mirrors
    ``MarkdownLoadOptions.preserve_empty_lines``).
    """
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    if lines and lines[-1] == "":
        lines.pop()  # trailing newline, not an extra blank line
    definitions: dict[str, tuple[str, str]] = {}
    _parse_blocks(lines, DocumentBlock(), definitions, preserve_empty_lines)  # pre-scan: collect all labels first
    doc = DocumentBlock()
    _parse_blocks(lines, doc, definitions, preserve_empty_lines)
    _resolve_references(doc, definitions)
    return doc


def _normalize_label(label: str) -> str:
    """Collapse whitespace runs (including embedded newlines) to one space, then uppercase -- matches C#'s
    LinkDefinitionBlock.NormalizeString + ReferenceBlock.Label, both of which use plain
    ToUpper(InvariantCulture), not full Unicode case folding (e.g. "SS" and the sharp-S "ẞ" don't match)."""
    return " ".join(label.split()).upper()


def _resolve_references(block: Block, definitions: dict[str, tuple[str, str]]) -> None:
    """Resolve reference-style ``LinkTextBlock``s; splice unresolved ones back to their literal source text."""
    new_children: list[Block] = []
    for child in block.children:
        if isinstance(child, LinkTextBlock):
            if child.uri is None and child.definition_label:
                entry = definitions.get(_normalize_label(child.definition_label))
                if entry is not None:
                    child.uri, child.title = entry
            if child.uri is not None:
                _resolve_references(child, definitions)
                new_children.append(child)
            else:
                fallback = TextBlock(child.raw_text)
                fallback.parent = block
                new_children.append(fallback)
        else:
            _resolve_references(child, definitions)
            new_children.append(child)
    block.children = new_children


def _leading_ws_flat_len(line: str) -> int:
    """Flat length of leading spaces/tabs, matches C#'s GetLength (tab=4 always, not column-based)."""
    n = 0
    for ch in line:
        if ch not in (" ", "\t"):
            break
        n += 4 if ch == "\t" else 1
    return n


def _trim_start_flat(line: str, target: int) -> str:
    """Drop leading spaces/tabs until flat length reaches *target*, matches C#'s GetCharsCount."""
    pos = 0
    n = len(line)
    remaining = target
    while remaining > 0 and pos < n and line[pos] in (" ", "\t"):
        remaining -= 4 if line[pos] == "\t" else 1
        pos += 1
    return line[pos:]


def _opening_run_flat_len(line: str, search_limit: int) -> tuple[int, int]:
    """Matches C#'s Block.GetOpening: consumes spaces/tabs, stopping at a non-ws char or exact match."""
    pos = 0
    n = len(line)
    length = 0
    while pos < n and line[pos] in (" ", "\t"):
        length += 4 if line[pos] == "\t" else 1
        pos += 1
        if length == search_limit:
            break
    return pos, length


def _is_blank(line: str) -> bool:
    return line.strip() == ""


def _looks_like_new_block(line: str) -> bool:
    """Whether *line* is allowed to interrupt an in-progress paragraph."""
    if _is_blank(line):
        return True
    if _ATX_RE.match(line):
        return True
    if _HR_RE.match(line):
        return True
    if _BLOCKQUOTE_RE.match(line):
        return True
    if _fence_match(line):
        return True
    html_match = _HTML_BLOCK_START_RE.match(line)
    if html_match and html_match.group(1).lower() in HTML_BLOCK_TAG_NAMES:  # HTML block types 1-6 interrupt
        return True
    if _FOOTNOTE_DEF_RE.match(line):
        return True
    bullet = _BULLET_ITEM_RE.match(line)
    if bullet and (bullet.group(4) or "").strip() != "":
        return True
    ordered = _ORDERED_ITEM_RE.match(line)
    if ordered and ordered.group(2) == "1" and (ordered.group(5) or "").strip() != "":
        return True
    return False


def is_block_start(line: str) -> bool:
    """``BlockParser.Parse(text, 0).Type != BlockType.Paragraph``."""
    line = line.lstrip(" \t")
    if not line:
        return False
    return bool(
        _ATX_RE.match(line)
        or _HR_RE.match(line)
        or _SETEXT_RE.match(line)
        or _BLOCKQUOTE_RE.match(line)
        or _fence_match(line)
        or _BULLET_ITEM_RE.match(line)
        or _ORDERED_ITEM_RE.match(line)
        or _FOOTNOTE_DEF_RE.match(line)
        or _REF_DEF_RE.match(line)
    )


def is_ordered_list_item(line: str) -> bool:
    """``MarkdownUtil.IsOrderedListItem``."""
    return bool(_ORDERED_ITEM_RE.match(line.lstrip(" \t")))


def _split_row_cells(line: str) -> list[str]:
    line = line.strip()
    if line.startswith("|"):
        line = line[1:]
    if line.endswith("|") and not line.endswith("\\|"):
        line = line[:-1]
    cells: list[str] = []
    buf: list[str] = []
    i = 0
    while i < len(line):
        if line[i] == "\\" and i + 1 < len(line) and line[i + 1] == "|":
            buf.append("|")
            i += 2
            continue
        if line[i] == "|":
            cells.append("".join(buf).strip())
            buf = []
            i += 1
            continue
        buf.append(line[i])
        i += 1
    cells.append("".join(buf).strip())
    return cells


def _delimiter_alignment(cell: str) -> Optional[str]:
    left = cell.startswith(":")
    right = cell.endswith(":")
    if left and right:
        return "center"
    if right:
        return "right"
    if left:
        return "left"
    return None


def _is_table_start(lines: list[str], i: int) -> bool:
    if i + 1 >= len(lines) or "|" not in lines[i] or "|" not in lines[i + 1]:
        return False
    if _leading_ws_flat_len(lines[i]) > 3 or _leading_ws_flat_len(lines[i + 1]) > 3:
        return False
    header_cells = _split_row_cells(lines[i])
    delim_cells = _split_row_cells(lines[i + 1])
    if not delim_cells or len(delim_cells) != len(header_cells):
        return False
    return all(_TABLE_DELIM_CELL_RE.match(c) for c in delim_cells)


def _parse_table(lines: list[str], i: int, parent: Block, definitions: dict[str, tuple[str, str]]) -> int:
    header_cells = _split_row_cells(lines[i])
    alignments = [_delimiter_alignment(c) for c in _split_row_cells(lines[i + 1])]
    table = TableBlock(alignments)
    parent.add_child(table)
    _add_table_row(table, header_cells, alignments, is_header=True, definitions=definitions)

    i += 2
    while i < len(lines) and not _is_blank(lines[i]) and "|" in lines[i]:
        _add_table_row(table, _split_row_cells(lines[i]), alignments, is_header=False, definitions=definitions)
        i += 1
    return i


def _add_table_row(
    table: TableBlock,
    cells: list[str],
    alignments: list[Optional[str]],
    is_header: bool,
    definitions: dict[str, tuple[str, str]],
) -> None:
    row = RowBlock(is_header=is_header)
    table.add_child(row)
    for idx, cell_text in enumerate(cells):
        alignment = alignments[idx] if idx < len(alignments) else None
        cell = CellBlock(alignment)
        row.add_child(cell)
        for node in parse_inline(cell_text, definitions):
            cell.add_child(node)


def _parse_blocks(
    lines: list[str], parent: Block, definitions: dict[str, tuple[str, str]], preserve_empty_lines: bool = False
) -> None:
    i = 0
    n = len(lines)
    while i < n:
        line = lines[i]

        if _is_blank(line):
            if preserve_empty_lines:
                parent.add_child(ParagraphBlock())
            i += 1
            continue

        if _fence_match(line):
            i = _parse_fenced_code(lines, i, parent)
            continue

        atx_match = _ATX_RE.match(line)
        if atx_match:
            level = len(atx_match.group(1))
            content = (atx_match.group(2) or "").strip()
            content = re.sub(r"(?:^|\s)#+\s*$", "", content).strip()
            heading = AtxHeadingBlock(level)
            parent.add_child(heading)
            for node in parse_inline(content, definitions):
                heading.add_child(node)
            i += 1
            continue

        if _HR_RE.match(line):
            parent.add_child(HorizontalRuleBlock())
            i += 1
            continue

        if _is_table_start(lines, i):
            i = _parse_table(lines, i, parent, definitions)
            continue

        if _BLOCKQUOTE_RE.match(line):
            i = _parse_blockquote(lines, i, parent, definitions, preserve_empty_lines)
            continue

        if _FOOTNOTE_DEF_RE.match(line):
            i = _parse_footnote_definition(lines, i, parent, definitions, preserve_empty_lines)
            continue

        ref_result = _try_reference_definition(lines, i)
        if ref_result is not None:
            label, uri, title, next_i = ref_result
            # CommonMark: when several definitions share a label, the first one wins.
            definitions.setdefault(_normalize_label(label), (unescape_destination(uri), unescape_destination(title)))
            i = next_i
            continue

        if _BULLET_ITEM_RE.match(line) or _ORDERED_ITEM_RE.match(line):
            i = _parse_list(lines, i, parent, definitions, preserve_empty_lines)
            continue

        if _opening_run_flat_len(line, 4)[1] > 3:
            i = _parse_indented_code(lines, i, parent)
            continue

        html_match = _HTML_BLOCK_START_RE.match(line)
        if html_match and html_match.group(1).lower() in HTML_BLOCK_TAG_NAMES:
            i = _parse_html_block(lines, i, parent)
            continue

        i = _parse_paragraph(lines, i, parent, definitions)


def _parse_fenced_code(lines: list[str], i: int, parent: Block) -> int:
    match = _fence_match(lines[i])
    fence_char = match.group(1)[0]
    fence_len = len(match.group(1))
    info = (match.group(2) or "").strip()
    fence_indent = _leading_ws_flat_len(lines[i])
    i += 1
    code_lines: list[str] = []
    close_re = re.compile(rf"^ {{0,3}}{re.escape(fence_char)}{{{fence_len},}}[ \t]*$")
    while i < len(lines):
        if close_re.match(lines[i]):
            i += 1
            break
        code_lines.append(_trim_start_flat(lines[i], fence_indent) if fence_indent else lines[i])
        i += 1
    parent.add_child(FencedCodeBlock("\n".join(code_lines), info, fence_char))
    return i


def _parse_indented_code(lines: list[str], i: int, parent: Block) -> int:
    code_lines: list[str] = []
    while i < len(lines):
        line = lines[i]
        if _is_blank(line):
            code_lines.append(_trim_start_flat(line, 4))  # keeps whitespace past the threshold
            i += 1
            continue
        chars, flat = _opening_run_flat_len(line, 4)
        if flat > 3:
            code_lines.append(line[chars:])
            i += 1
            continue
        break
    while code_lines and code_lines[-1].strip() == "":
        code_lines.pop()
    parent.add_child(IndentedCodeBlock("\n".join(code_lines)))
    return i


def _parse_blockquote(
    lines: list[str],
    i: int,
    parent: Block,
    definitions: dict[str, tuple[str, str]],
    preserve_empty_lines: bool = False,
) -> int:
    quote_lines: list[str] = []
    while i < len(lines):
        match = _BLOCKQUOTE_RE.match(lines[i])
        if match:
            quote_lines.append(match.group(1))
            i += 1
            continue
        # lazy continuation: no '>' but still extends the open paragraph
        prev = quote_lines[-1] if quote_lines else None
        can_continue = (
            prev is not None
            and prev.strip() != ""
            and not _BLOCKQUOTE_RE.match(prev)
            and not _looks_like_new_block(prev)
            and _opening_run_flat_len(prev, 4)[1] <= 3  # indented code has no lazy continuation
            and not _looks_like_new_block(lines[i])
        )
        if not can_continue:
            break
        quote_lines.append(lines[i])
        i += 1
    quote = QuoteBlock()
    parent.add_child(quote)
    _parse_blocks(quote_lines, quote, definitions, preserve_empty_lines)
    if not quote.children:
        quote.add_child(ParagraphBlock())  # keep the marker visible even with no content
    return i


def _try_reference_definition(lines: list[str], i: int) -> Optional[tuple[str, str, str, int]]:
    """Greedily tries growing line spans since the label/destination/title may span lines."""
    if _FOOTNOTE_DEF_RE.match(lines[i]) or not _REF_DEF_LOOSE_START_RE.match(lines[i]):
        return None
    parts = [lines[i].strip()]
    j = i + 1
    best = None
    while True:
        match = _REF_DEF_RE.match(" ".join(parts))
        if match is not None:
            best = (match, j)
        if j >= len(lines) or _is_blank(lines[j]) or len(parts) >= 8:
            break
        parts.append(lines[j].strip())
        j += 1
    if best is None:
        return None
    match, next_i = best
    label, uri_angle, uri_plain, title_dq, title_sq, title_paren = match.groups()
    if not _normalize_label(label):  # a label must contain at least one non-whitespace character
        return None
    uri = uri_angle if uri_angle is not None else uri_plain
    return label, uri, title_dq or title_sq or title_paren or "", next_i


def _parse_footnote_definition(
    lines: list[str],
    i: int,
    parent: Block,
    definitions: dict[str, tuple[str, str]],
    preserve_empty_lines: bool = False,
) -> int:
    match = _FOOTNOTE_DEF_RE.match(lines[i])
    label = match.group(1)
    item_lines = [match.group(2) or ""]
    i += 1
    while i < len(lines):
        line = lines[i]
        if _is_blank(line):
            item_lines.append("")
            i += 1
            continue
        chars, flat = _opening_run_flat_len(line, 4)
        if flat > 3:
            item_lines.append(line[chars:])
            i += 1
            continue
        break
    while item_lines and item_lines[-1] == "":
        item_lines.pop()
    definition = FootnoteDefinitionBlock(label)
    parent.add_child(definition)
    _parse_blocks(item_lines, definition, definitions, preserve_empty_lines)
    return i


def _marker_text(bullet_match, ordered_match) -> str:
    if bullet_match:
        return bullet_match.group(2)
    return ordered_match.group(2) + ordered_match.group(3)


def _parse_list(
    lines: list[str],
    i: int,
    parent: Block,
    definitions: dict[str, tuple[str, str]],
    preserve_empty_lines: bool = False,
) -> int:
    current_list: Optional[ListBlock] = None
    is_first_item = True

    while i < len(lines):
        line = lines[i]
        if _HR_RE.match(line):
            break
        bullet_match = _BULLET_ITEM_RE.match(line)
        ordered_match = None if bullet_match else _ORDERED_ITEM_RE.match(line)
        if not bullet_match and not ordered_match:
            break

        if bullet_match:
            indent, marker_char, spaces, content = bullet_match.groups()[:4]
            marker = _BULLET_MARKER_CHARS[marker_char]
            ordered = False
            start_at = 1
        else:
            indent, number, _punct, spaces, content = ordered_match.groups()[:5]
            marker = ListMarker.DOT if ordered_match.group(3) == "." else ListMarker.PARENTHESIS
            ordered = True
            start_at = int(number)

        if current_list is None or current_list.ordered != ordered or current_list.marker != marker:
            current_list = ListBlock(ordered, marker)
            parent.add_child(current_list)
            is_first_item = True

        item: ListItemBlock = (
            OrderedListItemBlock(start_at, marker) if ordered else BulletListItemBlock(marker)
        )
        item.is_level_start = is_first_item
        current_list.add_child(item)
        is_first_item = False

        # Matches C#'s ListLabelLength: only marker + 1 separator char count as the item's own width.
        # The floor is 2 for a bullet marker, 3 for an ordered one (digit + delimiter + space).
        marker_flat = len(_marker_text(bullet_match, ordered_match))  # never a tab
        sep_flat = (4 if spaces[0] == "\t" else 1) if spaces else 0
        base_indent = len(indent) + max(3 if ordered else 2, marker_flat + sep_flat)
        remainder = (spaces[1:] if spaces else "") + (content or "")
        _, remainder_flat = _opening_run_flat_len(remainder, 4)
        content_indent = base_indent if remainder_flat > 3 else base_indent + remainder_flat
        item_lines = [remainder]
        first_is_blank = not item_lines[0].strip()
        i += 1
        while i < len(lines):
            next_line = lines[i]
            if _is_blank(next_line):
                if first_is_blank and len(item_lines) == 1:
                    break  # marker line + one blank -> empty item, not lazy content
                item_lines.append("")
                i += 1
                continue
            if _leading_ws_flat_len(next_line) >= content_indent:
                item_lines.append(_trim_start_flat(next_line, content_indent))
                i += 1
                continue
            # lazy continuation, like _parse_blockquote
            prev = item_lines[-1] if item_lines else None
            next_is_list_marker = bool(_BULLET_ITEM_RE.match(next_line) or _ORDERED_ITEM_RE.match(next_line))
            can_continue = (
                prev is not None
                and prev.strip() != ""
                and not _looks_like_new_block(prev)
                and not next_is_list_marker
                and not _looks_like_new_block(next_line)
            )
            if not can_continue:
                break
            item_lines.append(next_line)
            i += 1

        while item_lines and item_lines[-1] == "":
            item_lines.pop()
        _parse_blocks(item_lines, item, definitions, preserve_empty_lines)
        if not item.children:
            item.add_child(ParagraphBlock())  # keep the marker visible even with no content

    return i


def _parse_html_block(lines: list[str], i: int, parent: Block) -> int:
    match = _HTML_BLOCK_START_RE.match(lines[i])
    tag_name = match.group(1).lower()
    closing = f"</{tag_name}>"
    html_lines = [lines[i]]
    i += 1
    if closing not in html_lines[0].lower():
        while i < len(lines) and not _is_blank(lines[i]):
            html_lines.append(lines[i])
            found_close = closing in lines[i].lower()
            i += 1
            if found_close:
                break
    raw_text = "\n".join(html_lines).strip()
    is_known = tag_name in HTML_BLOCK_TAG_NAMES
    if is_known and _SELF_CONTAINED_HTML_RE.match(raw_text):
        parent.add_child(HtmlTagBlock(tag_name, raw_text, is_self_contained=True))
        return i
    tag_events = list(_HTML_TAG_RE.finditer(raw_text))
    if is_known and not _tags_balanced(tag_events):
        # unclosed within its own block; open+close together so state doesn't leak into later blocks
        parent.add_child(HtmlTagBlock(tag_name, raw_text, is_self_contained=True))
        return i
    pos = 0
    for tag_match in tag_events:
        if tag_match.start() > pos:
            parent.add_child(TextBlock(raw_text[pos:tag_match.start()]))
        is_closing, name, _attrs, self_close = tag_match.groups()
        parent.add_child(
            HtmlTagBlock(name, tag_match.group(0), is_self_closing=bool(self_close), is_closing=bool(is_closing))
        )
        pos = tag_match.end()
    if pos < len(raw_text):
        parent.add_child(TextBlock(raw_text[pos:]))
    return i


def _tags_balanced(tag_events: list[re.Match]) -> bool:
    depth = 0
    for tag_match in tag_events:
        is_closing, _name, _attrs, self_close = tag_match.groups()
        if self_close:
            continue
        depth += -1 if is_closing else 1
        if depth < 0:
            return False
    return depth == 0


def _parse_paragraph(lines: list[str], i: int, parent: Block, definitions: dict[str, tuple[str, str]]) -> int:
    para_lines = [lines[i]]
    i += 1
    setext_level: Optional[int] = None

    while i < len(lines):
        line = lines[i]
        if _is_blank(line):
            break
        setext_match = _SETEXT_RE.match(line)
        if setext_match:
            setext_level = 1 if setext_match.group(1)[0] == "=" else 2
            i += 1
            break
        if _looks_like_new_block(line):
            break
        para_lines.append(line)
        i += 1

    text = "\n".join(l.lstrip() for l in para_lines)

    if setext_level is not None:
        heading = SetextHeadingBlock(setext_level)
        parent.add_child(heading)
        for node in parse_inline(text, definitions):
            heading.add_child(node)
    else:
        paragraph = ParagraphBlock()
        parent.add_child(paragraph)
        for node in parse_inline(text, definitions):
            paragraph.add_child(node)

    return i
