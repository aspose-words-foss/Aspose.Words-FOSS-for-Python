"""Tiny XML element builder.

Used instead of :mod:`xml.etree.ElementTree` because OOXML uses fixed
``w:``/``r:`` prefixes and ET tends to rewrite them as ``ns0:`` /
``ns1:`` unless every namespace is registered globally.  Building the
strings directly keeps the writer self-contained and the diff readable.
"""


import re
from typing import Iterable, Optional, Union

# Characters that must be escaped inside element text and attribute values.
_TEXT_ESCAPE = {
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
}
_ATTR_ESCAPE = {
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
}


def _is_xml_char(c: str) -> bool:
    """Return True if ``c`` is allowed by XML 1.0 char production.

    Word legitimately uses 0x0B (vertical tab) and 0x0C (form feed) as
    soft line breaks and page breaks inside its OOXML, but these are
    NOT valid in XML — they break round-trips through standard
    parsers.  We drop them silently; the structural ``<w:br/>`` /
    ``<w:lastRenderedPageBreak/>`` siblings carry the same semantics.
    """
    cp = ord(c)
    return cp in (0x09, 0x0A, 0x0D) or 0x20 <= cp <= 0xD7FF or 0xE000 <= cp <= 0xFFFD


def escape_text(text: str) -> str:
    """Escape ``&`` ``<`` ``>`` and strip XML-illegal control chars."""
    return "".join(_TEXT_ESCAPE.get(c, c) for c in text if _is_xml_char(c))


def escape_attr(text: str) -> str:
    """Escape characters illegal inside a double-quoted attribute value."""
    return "".join(_ATTR_ESCAPE.get(c, c) for c in text if _is_xml_char(c))


AttrValue = Union[str, int, float, None]


def el(
    tag: str,
    attrs: Optional[dict[str, AttrValue]] = None,
    children: Optional[Union[str, Iterable[str]]] = None,
    *,
    self_closing_when_empty: bool = True,
) -> str:
    """Build a single XML element string.

    ``attrs`` whose value is ``None`` are silently dropped — this lets
    callers pass conditional attributes without a wrapping ``if``.
    ``children`` may be a single pre-rendered XML string, an iterable
    of strings, or ``None`` for an empty element.
    """
    parts = ["<", tag]
    if attrs:
        for k, v in attrs.items():
            if v is None:
                continue
            parts.append(f' {k}="{escape_attr(str(v))}"')

    body = ""
    if children is None:
        pass
    elif isinstance(children, str):
        body = children
    else:
        body = "".join(children)

    if not body and self_closing_when_empty:
        parts.append("/>")
        return "".join(parts)
    parts.append(">")
    parts.append(body)
    parts.append(f"</{tag}>")
    return "".join(parts)


def _t_element(text: str, *, instr: bool = False) -> str:
    """Build a single ``<w:t>`` (or ``<w:instrText>``) element.

    Outer whitespace gets ``xml:space="preserve"``.  ``instr=True``
    switches the tag to ``<w:instrText>`` for runs that carry field
    code (the text between ``<w:fldChar w:fldCharType="begin"/>`` and
    ``<w:fldChar w:fldCharType="separate"/>``); Aspose's model keeps
    field code in regular runs whose element name is ``<w:instrText>``
    rather than ``<w:t>``.
    """
    needs_preserve = text != text.strip()
    attrs: dict[str, AttrValue] = {}
    if needs_preserve:
        attrs["xml:space"] = "preserve"
    tag = "w:instrText" if instr else "w:t"
    return el(tag, attrs, escape_text(text))


def text_run(text: str, *, instr: bool = False) -> str:
    """Render raw run text as the children of a ``<w:r>`` element.

    OOXML represents soft line breaks, page breaks, column breaks and
    tab stops as standalone sibling elements (``<w:br/>``, ``<w:br
    w:type="page"/>``, ``<w:br w:type="column"/>``, ``<w:tab/>``)
    rather than literal control characters inside ``<w:t>``.  We split
    on those four whitespace types and emit the structural form so
    Word renders the layout the way the source intended (and so the
    output is well-formed XML — ``\\x0c`` / ``\\x0b`` are illegal
    inside an XML element); everything else stays in ``<w:t>`` with
    ``xml:space="preserve"`` whenever leading or trailing spaces matter.

    ``instr=True`` emits ``<w:instrText>`` instead of ``<w:t>`` for
    runs that sit between a field's ``begin`` and ``separate`` markers.
    Break / tab / cr children are not introduced in field code in
    practice, so the structural splitting still runs the same way.
    """
    if not text:
        return _t_element("", instr=instr)

    parts: list[str] = []
    buf: list[str] = []

    def flush() -> None:
        if buf:
            parts.append(_t_element("".join(buf), instr=instr))
            buf.clear()

    for ch in text:
        if ch == "\n":
            flush()
            parts.append("<w:br/>")
        elif ch == "\t":
            flush()
            parts.append("<w:tab/>")
        elif ch == "\f":
            flush()
            parts.append('<w:br w:type="page"/>')
        elif ch == "\v":
            flush()
            parts.append('<w:br w:type="column"/>')
        else:
            buf.append(ch)
    flush()
    if not parts:
        # Run with only control characters that got stripped — emit an
        # empty ``<w:t>`` so the writer still produces a well-formed
        # ``<w:r>`` (Word treats a run with zero children as malformed).
        return _t_element("", instr=instr)
    return "".join(parts)


XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'


# Inline content that MUST NOT be re-indented.  Word inherits
# ``xml:space="preserve"`` for run content, so any whitespace introduced
# *inside* ``<w:r>`` becomes significant text and gets rendered as gaps —
# frequently bad enough that Word refuses to open the document.  We
# snapshot these blocks before indenting and splice them back in
# afterwards so the pretty output stays flat where the schema demands it.
_PRESERVE_RE = re.compile(
    r"<(w:r|w:t|w:instrText|w:delText)\b[^>]*?(?:/>|>.*?</\1>)",
    re.DOTALL,
)

# Containers whose direct children are block-level enough to deserve a
# line of their own when pretty-printing — everything else stays packed.
# Word is genuinely picky about whitespace anywhere inside paragraph
# / property elements (the OOXML schema lets it interpret stray text
# nodes as significant), so we only break at structural boundaries and
# leave the inside of each block on a single line.
_BLOCK_CONTAINERS: frozenset[str] = frozenset({
    "w:body",
    "w:hdr",
    "w:ftr",
    "w:tbl",
    "w:tr",
    "w:tc",
    "w:styles",
    "w:numbering",
    "w:abstractNum",
    "w:lvl",
    "w:num",
    "w:lvlOverride",
    "w:style",
    "w:settings",
    "w:fonts",
    "w:font",
    "w:latentStyles",
    "w:compat",
    "w:docDefaults",
    "w:rPrDefault",
    "w:pPrDefault",
    "Types",
    "Relationships",
    "cp:coreProperties",
    "Properties",
})

_TAG_OPEN_RE = re.compile(r"<([A-Za-z_][\w:.-]*)\b[^>]*?(/?)>")


def indent_xml(xml: str, indent_str: str = "  ") -> str:
    """Add structural newlines to an XML payload that Word will still open.

    Word does *not* tolerate indented whitespace inside ``<w:r>`` (or any
    element where ``xml:space="preserve"`` is inherited), and several
    common loaders reject inline whitespace inside property containers
    (``<w:pPr>``, ``<w:rPr>``, ``<w:tcPr>``, ``<w:tblPr>``, ...) as well.
    To stay safe we only break the line between siblings of a small set
    of structural containers (body / table / numbering / style / ...);
    everything inside a paragraph or a property block stays packed onto
    a single line, exactly like the compact output Word accepts.

    ``indent_str`` controls how deep each new line is pushed.  Self-closing
    and content-bearing tags are treated identically: a newline is
    emitted before each direct child of a structural container, and the
    container's own closing tag is pushed onto its own line.
    """
    decl = ""
    body = xml
    if body.startswith("<?xml"):
        end = body.index("?>") + 2
        decl = body[:end]
        body = body[end:].lstrip("\n")

    # Stash run-level content (and the text-bearing leaves) so we never
    # re-tokenise it; collapse any pre-existing whitespace between its
    # children so the transform is idempotent on already-pretty input.
    preserved: list[str] = []

    def _save(m: re.Match) -> str:
        snippet = re.sub(r">\s+<", "><", m.group(0))
        i = len(preserved)
        preserved.append(snippet)
        return f"\x00P{i}\x00"

    body = _PRESERVE_RE.sub(_save, body)

    out_parts: list[str] = []
    stack: list[tuple[str, bool]] = []  # (tag, is_block_container)
    pos = 0
    pending_newline = False

    def _pad(level: int) -> str:
        return indent_str * level

    while pos < len(body):
        ch = body[pos]
        if ch == "<":
            close_idx = body.index(">", pos)
            tag_text = body[pos:close_idx + 1]
            m = _TAG_OPEN_RE.match(tag_text)
            if tag_text.startswith("</"):
                # closing tag
                if stack:
                    tag, was_block = stack.pop()
                    if was_block:
                        # newline + indent at the container's own depth
                        out_parts.append("\n" + _pad(len(stack)))
                out_parts.append(tag_text)
                pending_newline = False
            elif m and m.group(2) == "/":
                # self-closing
                if stack and stack[-1][1]:
                    out_parts.append("\n" + _pad(len(stack)))
                out_parts.append(tag_text)
                pending_newline = False
            elif m:
                tag = m.group(1)
                is_block = tag in _BLOCK_CONTAINERS
                if stack and stack[-1][1]:
                    out_parts.append("\n" + _pad(len(stack)))
                out_parts.append(tag_text)
                stack.append((tag, is_block))
                pending_newline = False
            else:
                # processing instruction / comment / unrecognised — pass through
                out_parts.append(tag_text)
            pos = close_idx + 1
        elif ch == "\x00":
            # preserved block placeholder — bring it in on a fresh line if the
            # parent is a block container.
            end_idx = body.index("\x00", pos + 1) + 1
            placeholder = body[pos:end_idx]
            if stack and stack[-1][1]:
                out_parts.append("\n" + _pad(len(stack)))
            out_parts.append(placeholder)
            pos = end_idx
        else:
            # plain text — copy it verbatim, but compress runs of pure
            # whitespace away (we'll re-add them ourselves where needed).
            run_end = pos
            while run_end < len(body) and body[run_end] not in ("<", "\x00"):
                run_end += 1
            chunk = body[pos:run_end]
            if chunk.strip():
                out_parts.append(chunk)
            pos = run_end

    result = "".join(out_parts)
    for i, p in enumerate(preserved):
        result = result.replace(f"\x00P{i}\x00", p)

    if decl:
        return decl + "\n" + result + "\n"
    return result + "\n"
