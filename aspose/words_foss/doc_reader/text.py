"""
Text extraction, field evaluation, and hyperlink parsing for DOC files.
"""

import re
import struct

from aspose.words_foss.doc_reader.constants import IDX_CLX
from aspose.words_foss.doc_reader.fib import FibData, get_fc_lcb
from aspose.words_foss.docx_reader import PAGE_FIELD_SENTINEL

# Characters to strip from run text (shape anchors, field codes, etc.)
# Includes \x13-\x15 (field start/sep/end).
# Note: \x0c (form-feed / page break) is NOT stripped here — it is
# handled by _build_ldm_body_children as a page-break marker.
_CONTROL_CHARS = re.compile(r"[\x01-\x06\x08\x0e\x0f\x13-\x15\x1a-\x1f]")


def clean_control_chars(text: str) -> str:
    """Remove Word-internal control characters from text."""
    return _CONTROL_CHARS.sub("", text)


def extract_text_via_piece_table(wd: bytes, table: bytes, fib: FibData) -> str:
    """Extract document text using the CLX piece table."""
    fc_clx, lcb_clx = get_fc_lcb(wd, fib, IDX_CLX)

    if lcb_clx == 0:
        # No CLX: try reading text from fixed offset 0x800
        text_offset = 0x800
        if text_offset + fib.ccp_text * 2 <= len(wd):
            text_bytes = wd[text_offset : text_offset + fib.ccp_text * 2]
            return text_bytes.decode("utf-16-le", errors="replace")
        return ""

    clx = table[fc_clx : fc_clx + lcb_clx]
    pos = 0
    text_parts: list[str] = []

    while pos < len(clx):
        clxt = clx[pos]
        if clxt == 0x02:  # Piece table
            pos += 1
            pcdt_size = struct.unpack_from("<I", clx, pos)[0]
            pos += 4
            pt = clx[pos : pos + pcdt_size]

            n = (pcdt_size - 4) // 12
            cps = [struct.unpack_from("<I", pt, i * 4)[0] for i in range(n + 1)]

            for i in range(n):
                pcd_start = (n + 1) * 4 + i * 8
                fc_val = struct.unpack_from("<I", pt, pcd_start + 2)[0]
                f_compressed = bool(fc_val & 0x40000000)
                fc_real = fc_val & 0x3FFFFFFF
                char_count = cps[i + 1] - cps[i]

                if f_compressed:
                    text_bytes = wd[fc_real // 2 : fc_real // 2 + char_count]
                    text_parts.append(text_bytes.decode("cp1252", errors="replace"))
                else:
                    text_bytes = wd[fc_real : fc_real + char_count * 2]
                    text_parts.append(text_bytes.decode("utf-16-le", errors="replace"))
            break
        elif clxt == 0x01:  # Grpprl
            pos += 1
            cb = struct.unpack_from("<H", clx, pos)[0]
            pos += 2 + cb
        else:
            pos += 1

    return "".join(text_parts)


def parse_hyperlink(text: str) -> tuple[str, str]:
    """Parse a Word field code hyperlink.

    Field codes: \\x13 HYPERLINK "url" \\x01\\x14display_text\\x15
    Returns (display_text, url).
    """
    url = ""
    display = ""

    # Extract URL from between quotes after HYPERLINK
    begin_idx = text.find("\x13")
    sep_idx = text.find("\x14")
    end_idx = text.find("\x15")

    if begin_idx >= 0 and sep_idx >= 0:
        field_code = text[begin_idx + 1 : sep_idx]
        # Remove the \x01 separator if present
        field_code = field_code.replace("\x01", "").strip()

        # Extract URL from HYPERLINK "url"
        if "HYPERLINK" in field_code:
            parts = field_code.split('"')
            if len(parts) >= 2:
                url = parts[1]

    if sep_idx >= 0 and end_idx >= 0:
        display = text[sep_idx + 1 : end_idx]
    elif sep_idx >= 0:
        display = text[sep_idx + 1 :]

    return display, url


def evaluate_fields(text: str) -> str:
    """Recursively evaluate Word field codes in text.

    Handles nested fields (e.g. HYPERLINK containing PAGEREF).
    Returns clean text with:
    - HYPERLINK fields → ``[display](#bookmark)``
    - PAGEREF / PAGE / other fields → their result text only
    - TOC fields → their result text (which contains nested entries)
    """
    result: list[str] = []
    pos = 0
    while pos < len(text):
        if text[pos] == "\x13":
            # Find matching \x14 separator, accounting for nesting
            depth = 1
            sep_pos = -1
            end_pos = -1
            scan = pos + 1
            while scan < len(text) and depth > 0:
                ch = text[scan]
                if ch == "\x13":
                    depth += 1
                elif ch == "\x14":
                    if depth == 1 and sep_pos < 0:
                        sep_pos = scan
                elif ch == "\x15":
                    depth -= 1
                    if depth == 0:
                        end_pos = scan
                scan += 1

            if sep_pos < 0:
                # Malformed: no separator, skip the field start
                pos += 1
                continue

            field_code = text[pos + 1 : sep_pos].replace("\x01", "").strip()
            if end_pos < 0:
                end_pos = len(text)
            result_text = text[sep_pos + 1 : end_pos]

            # Recursively evaluate nested fields in the result
            result_text = evaluate_fields(result_text)

            if "HYPERLINK" in field_code:
                # Extract URL/bookmark from HYPERLINK field code
                parts = field_code.split('"')
                url = parts[1] if len(parts) >= 2 else ""
                # Check for \l switch (local bookmark)
                is_local = "\\l" in field_code.split('"')[0]
                if is_local and url:
                    url = "#" + url
                display = clean_control_chars(result_text)
                # TOC entries embed PAGEREF inside the HYPERLINK result
                # as "Title\tPageNum".  Split so the page number lands
                # outside the link: [Title](#bkmk)\tPageNum.
                suffix = ""
                if "\t" in display:
                    link_part, tab_rest = display.rsplit("\t", 1)
                    if tab_rest.strip().isdigit():
                        display = link_part
                        suffix = "\t" + tab_rest
                if display and url:
                    result.append(f"[{display}]({url}){suffix}")
                elif display:
                    result.append(display + suffix)
            elif field_code.split()[0] == "PAGE" if field_code.strip() else False:
                # Emit the same sentinel the DOCX reader uses so the
                # PDF writer can substitute the live page number.

                result.append(PAGE_FIELD_SENTINEL)
            else:
                # TOC, PAGEREF, NUMPAGES, etc. — use result text only
                cleaned = clean_control_chars(result_text)
                if cleaned:
                    result.append(cleaned)

            pos = end_pos + 1
        else:
            result.append(text[pos])
            pos += 1

    return "".join(result)
