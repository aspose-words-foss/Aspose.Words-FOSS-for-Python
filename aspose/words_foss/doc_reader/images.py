"""
OfficeArt / image extraction for DOC files.

Parses Escher BSE (Blip Store Entry) records, shape-to-blip mappings,
and PlcSpaMom/PlcSpaHdr shape anchors.
"""

from __future__ import annotations

import struct

from aspose.words_foss.doc_reader.constants import (
    ESCHER_BLIP_JPEG,
    ESCHER_BLIP_JPEG2,
)


class BlipInfo:
    """Parsed BSE (Blip Store Entry) with location of image data."""

    __slots__ = ("blip_type", "fo_delay", "img_offset", "img_size")

    def __init__(self, blip_type: int, fo_delay: int, img_offset: int, img_size: int):
        self.blip_type = blip_type  # Escher fbt of the blip record
        self.fo_delay = fo_delay
        self.img_offset = img_offset  # byte offset in WordDocument stream
        self.img_size = img_size  # image payload size (after blip header)


class ShapeLineProps:
    """Escher line/fill properties for a shape (from FOpt records)."""

    __slots__ = ("line_color_rgb", "line_width_pt", "fill_color_rgb")

    def __init__(
        self,
        line_color_rgb: tuple[int, int, int] | None = None,
        line_width_pt: float = 0.0,
        fill_color_rgb: tuple[int, int, int] | None = None,
    ):
        self.line_color_rgb = line_color_rgb  # (R, G, B) or None
        self.line_width_pt = line_width_pt
        self.fill_color_rgb = fill_color_rgb  # (R, G, B) or None


class ShapeAnchor:
    """Parsed SPA (Shape Address) from PlcSpaMom / PlcSpaHdr."""

    __slots__ = ("cp", "spid", "left", "top", "width", "height", "wr", "bx", "by")

    def __init__(
        self,
        cp: int,
        spid: int,
        left: float,
        top: float,
        width: float,
        height: float,
        wr: int = 0,
        bx: int = 0,
        by: int = 0,
    ):
        self.cp = cp
        self.spid = spid
        self.left = left  # points
        self.top = top  # points
        self.width = width  # points
        self.height = height  # points
        self.wr = wr  # SPA wrap type (bits 5-8 of wFlags)
        self.bx = bx  # X reference: 0=margin, 1=page, 2=char
        self.by = by  # Y reference: 0=margin, 1=page, 2=paragraph


def parse_blip_store(table: bytes, wd: bytes, dgg_offset: int, dgg_size: int) -> list[BlipInfo]:
    """Extract image blip info by scanning the DggInfo region for BSE
    records (fbt = 0xF007).  Each BSE's ``foDelay`` field gives the
    byte offset of the actual blip in the WordDocument stream.
    """
    blips: list[BlipInfo] = []
    data = table[dgg_offset : dgg_offset + dgg_size]
    # BSE signature: fbt = 0xF007 → bytes [+2]=0x07, [+3]=0xF0
    for pos in range(0, len(data) - 44):
        if data[pos + 2] == 0x07 and data[pos + 3] == 0xF0:
            cb = struct.unpack_from("<I", data, pos + 4)[0]
            if not (36 <= cb <= 500000):
                continue
            bse = data[pos + 8 : pos + 8 + min(cb, 44)]
            if len(bse) < 36:
                continue
            fo_delay = struct.unpack_from("<I", bse, 28)[0]
            if fo_delay == 0 or fo_delay + 8 > len(wd):
                continue
            blip_fbt = struct.unpack_from("<H", wd, fo_delay + 2)[0]
            blip_cb = struct.unpack_from("<I", wd, fo_delay + 4)[0]
            img_offset = fo_delay + 8 + 17  # header + rgbUid + tag
            img_size = blip_cb - 17
            if img_offset + img_size <= len(wd) and img_size > 0:
                blips.append(BlipInfo(blip_fbt, fo_delay, img_offset, img_size))
    return blips


def parse_shape_blip_map(table: bytes, dgg_offset: int, dgg_size: int) -> dict[int, int]:
    """Build spid -> pib mapping by raw-scanning the DggInfo region.

    Scans for Sp records (fbt=0xF00A, ver=2, cb=8) to collect shape
    IDs and for FOpt pib properties (raw pid = 0x4104) to find blip
    references.  Each pib is associated with the nearest preceding Sp.

    Returns ``{shape_id: blip_index}`` where blip_index is 1-based.
    """
    data = table[dgg_offset : dgg_offset + dgg_size]

    # Pass 1: find Sp records — fbt=0xF00A at bytes [+2,+3]
    sp_records: list[tuple[int, int]] = []
    for pos in range(0, len(data) - 12):
        if data[pos + 2] == 0x0A and data[pos + 3] == 0xF0:
            ver = data[pos] & 0x0F
            cb = struct.unpack_from("<I", data, pos + 4)[0]
            if ver == 0x02 and cb == 8:
                spid = struct.unpack_from("<I", data, pos + 8)[0]
                sp_records.append((pos, spid))

    # Pass 2: find FOpt pib properties (pid = 0x4104).
    # Scan every byte — FOpt property entries are 6 bytes so the
    # pid bytes can land on odd offsets.
    mapping: dict[int, int] = {}
    for pos in range(0, len(data) - 6):
        if data[pos] == 0x04 and data[pos + 1] == 0x41:
            pib = struct.unpack_from("<I", data, pos + 2)[0]
            if 0 < pib <= 100:
                for sp_off, spid in reversed(sp_records):
                    if sp_off < pos:
                        mapping[spid] = pib
                        break
    return mapping


def parse_shape_line_map(table: bytes, dgg_offset: int, dgg_size: int) -> dict[int, ShapeLineProps]:
    """Build spid -> ShapeLineProps mapping from Escher FOpt records.

    Properly parses FOpt container records (fbt=0xF00B) to extract
    line color (pid 0x01C0), line width (pid 0x01CB, in EMUs),
    and fill color (pid 0x0181) for each shape.

    Returns ``{shape_id: ShapeLineProps}``.
    """
    data = table[dgg_offset : dgg_offset + dgg_size]

    # Pass 1: find Sp records (fbt=0xF00A, ver=2, cb=8)
    sp_records: list[tuple[int, int]] = []
    for pos in range(0, len(data) - 12):
        if data[pos + 2] == 0x0A and data[pos + 3] == 0xF0:
            ver = data[pos] & 0x0F
            cb = struct.unpack_from("<I", data, pos + 4)[0]
            if ver == 0x02 and cb == 8:
                spid = struct.unpack_from("<I", data, pos + 8)[0]
                sp_records.append((pos, spid))

    # Pass 2: find FOpt containers (fbt=0xF00B) and parse their
    # property tables.  Each entry is 6 bytes: 2-byte pid + 4-byte value.
    result: dict[int, ShapeLineProps] = {}
    for pos in range(0, len(data) - 8):
        if data[pos + 2] != 0x0B or data[pos + 3] != 0xF0:
            continue
        nprops = (data[pos] | (data[pos + 1] << 8)) >> 4
        cb = struct.unpack_from("<I", data, pos + 4)[0]
        if nprops <= 0 or nprops > 50 or cb < nprops * 6:
            continue

        # Find the Sp that owns this FOpt (nearest preceding Sp record)
        owner_spid = None
        for sp_off, spid in reversed(sp_records):
            if sp_off < pos:
                owner_spid = spid
                break
        if owner_spid is None:
            continue

        line_color_cr: int | None = None
        line_width_emu = 0
        fill_color_cr: int | None = None

        for p in range(nprops):
            off = pos + 8 + p * 6
            if off + 6 > len(data):
                break
            pid = struct.unpack_from("<H", data, off)[0]
            val = struct.unpack_from("<I", data, off + 2)[0]
            base_pid = pid & 0x3FFF
            if base_pid == 0x01C0:  # lineColor (COLORREF)
                line_color_cr = val
            elif base_pid == 0x01CB:  # lineType / lineWidth (EMU)
                line_width_emu = val
            elif base_pid == 0x0181:  # fillColor (COLORREF)
                fill_color_cr = val

        if line_color_cr is not None or fill_color_cr is not None:
            line_rgb = None
            if line_color_cr is not None:
                line_rgb = (
                    line_color_cr & 0xFF,
                    (line_color_cr >> 8) & 0xFF,
                    (line_color_cr >> 16) & 0xFF,
                )
            fill_rgb = None
            if fill_color_cr is not None:
                fill_rgb = (
                    fill_color_cr & 0xFF,
                    (fill_color_cr >> 8) & 0xFF,
                    (fill_color_cr >> 16) & 0xFF,
                )
            lw_pt = line_width_emu / 12700.0 if line_width_emu else 0.0
            result[owner_spid] = ShapeLineProps(
                line_color_rgb=line_rgb,
                line_width_pt=lw_pt,
                fill_color_rgb=fill_rgb,
            )
    return result


def parse_plc_spa(table: bytes, fc: int, lcb: int) -> list[ShapeAnchor]:
    """Parse PlcSpaMom or PlcSpaHdr into a list of shape anchors."""
    if lcb == 0:
        return []
    data = table[fc : fc + lcb]
    n = (lcb - 4) // 30
    if n <= 0:
        return []

    anchors: list[ShapeAnchor] = []
    for i in range(n):
        cp = struct.unpack_from("<I", data, i * 4)[0]
        spa_off = (n + 1) * 4 + i * 26
        spid = struct.unpack_from("<I", data, spa_off)[0]
        xa_l = struct.unpack_from("<i", data, spa_off + 4)[0]
        ya_t = struct.unpack_from("<i", data, spa_off + 8)[0]
        xa_r = struct.unpack_from("<i", data, spa_off + 12)[0]
        ya_b = struct.unpack_from("<i", data, spa_off + 16)[0]
        # wFlags at offset +20 in the SPA entry (2 bytes).
        wr = bx = by = 0
        if spa_off + 22 <= len(data):
            w_flags = struct.unpack_from("<H", data, spa_off + 20)[0]
            bx = (w_flags >> 1) & 0x03  # bits 1-2: X ref (0=margin,1=page,2=char)
            by = (w_flags >> 3) & 0x03  # bits 3-4: Y ref (0=margin,1=page,2=para)
            wr = (w_flags >> 5) & 0x0F  # bits 5-8: wrap type
        # SPA coordinates are in twips (1 pt = 20 twips)
        anchors.append(
            ShapeAnchor(
                cp=cp,
                spid=spid,
                left=xa_l / 20.0,
                top=ya_t / 20.0,
                width=(xa_r - xa_l) / 20.0,
                height=(ya_b - ya_t) / 20.0,
                wr=wr,
                bx=bx,
                by=by,
            )
        )
    return anchors


# ---------------------------------------------------------------------------
# PlcftxbxTxt (FTXBXS) — textbox story → shape ID mapping
# ---------------------------------------------------------------------------


def parse_plcf_txbx_txt(table: bytes, fc: int, lcb: int) -> dict[int, tuple[int, int]]:
    """Parse PlcftxbxTxt to map each textbox story to its owning shape.

    The PLC contains (n+1) CPs (4 bytes each) followed by n FTXBXS
    data entries (22 bytes each).  Each FTXBXS has a ``lid`` field at
    offset 14 that holds the shape ID (spid).

    Returns ``{shape_id: (cp_start, cp_end)}`` where CPs are relative
    to the start of the textbox text stream.
    """
    if lcb == 0:
        return {}
    data = table[fc : fc + lcb]
    cb_data = 22  # FTXBXS struct size per MS-DOC 2.9.291
    # n = (lcb - 4) / (4 + 22) = (lcb - 4) / 26
    n = (lcb - 4) // (4 + cb_data)
    if n <= 0:
        return {}

    result: dict[int, tuple[int, int]] = {}
    for i in range(n):
        cp_start = struct.unpack_from("<I", data, i * 4)[0]
        cp_end = struct.unpack_from("<I", data, (i + 1) * 4)[0]
        ftxbxs_off = (n + 1) * 4 + i * cb_data
        if ftxbxs_off + 18 > len(data):
            break
        # lid at offset 14 within FTXBXS = shape ID
        lid = struct.unpack_from("<I", data, ftxbxs_off + 14)[0]
        if lid > 0 and cp_start < cp_end:
            result[lid] = (cp_start, cp_end)
    return result


# ---------------------------------------------------------------------------
# Escher child shape positions — ChildAnchor + Spgr records
# ---------------------------------------------------------------------------


class ChildAnchorInfo:
    """Position of a child shape within its parent group's coordinate system."""

    __slots__ = ("left", "top", "right", "bottom")

    def __init__(self, left: int, top: int, right: int, bottom: int):
        self.left = left
        self.top = top
        self.right = right
        self.bottom = bottom


class GroupShapeInfo:
    """Coordinate system of a shape group (from Spgr record)."""

    __slots__ = ("coord_left", "coord_top", "coord_right", "coord_bottom")

    def __init__(self, left: int, top: int, right: int, bottom: int):
        self.coord_left = left
        self.coord_top = top
        self.coord_right = right
        self.coord_bottom = bottom


def parse_escher_child_positions(
    table: bytes, dgg_offset: int, dgg_size: int
) -> tuple[dict[int, ChildAnchorInfo], dict[int, GroupShapeInfo], dict[int, int]]:
    """Parse Escher records to extract child shape positions and group info.

    Uses raw byte scanning (like ``parse_shape_blip_map``) to avoid
    record-walking issues with misaligned bytes between DggContainer
    and DgContainer.

    Returns:
        (child_anchors, group_info, child_to_parent) where:
        - child_anchors: ``{spid: ChildAnchorInfo}``
        - group_info: ``{spid: GroupShapeInfo}``
        - child_to_parent: ``{child_spid: group_spid}``
    """
    data = table[dgg_offset : dgg_offset + dgg_size]

    child_anchors: dict[int, ChildAnchorInfo] = {}
    group_info: dict[int, GroupShapeInfo] = {}
    child_to_parent: dict[int, int] = {}

    # Pass 1: find all Sp records (fbt=0xF00A, ver=2, cb=8)
    sp_records: list[tuple[int, int]] = []  # (offset, spid)
    for pos in range(0, len(data) - 12):
        if data[pos + 2] == 0x0A and data[pos + 3] == 0xF0:
            ver = data[pos] & 0x0F
            cb = struct.unpack_from("<I", data, pos + 4)[0]
            if ver == 0x02 and cb == 8:
                spid = struct.unpack_from("<I", data, pos + 8)[0]
                sp_records.append((pos, spid))

    # Pass 2: find Spgr records (fbt=0xF009) — group coordinate systems.
    # Spgr always immediately precedes its owning Sp in the stream.
    current_group_spid: int | None = None
    for pos in range(0, len(data) - 24):
        if data[pos + 2] == 0x09 and data[pos + 3] == 0xF0:
            cb = struct.unpack_from("<I", data, pos + 4)[0]
            if cb < 16:
                continue
            x1 = struct.unpack_from("<i", data, pos + 8)[0]
            y1 = struct.unpack_from("<i", data, pos + 12)[0]
            x2 = struct.unpack_from("<i", data, pos + 16)[0]
            y2 = struct.unpack_from("<i", data, pos + 20)[0]
            if x1 == 0 and y1 == 0 and x2 == 0 and y2 == 0:
                continue  # skip degenerate root groups
            # The next Sp record is the group shape that owns these bounds
            for sp_off, spid in sp_records:
                if sp_off > pos:
                    group_info[spid] = GroupShapeInfo(x1, y1, x2, y2)
                    current_group_spid = spid
                    break

    # Pass 3: find ChildAnchor records (fbt=0xF00F, cb=16).
    # ChildAnchor follows its owning Sp in the same SpContainer.
    for pos in range(0, len(data) - 24):
        if data[pos + 2] == 0x0F and data[pos + 3] == 0xF0:
            ver = data[pos] & 0x0F
            cb = struct.unpack_from("<I", data, pos + 4)[0]
            if cb != 16 or ver != 0:
                continue
            x1 = struct.unpack_from("<i", data, pos + 8)[0]
            y1 = struct.unpack_from("<i", data, pos + 12)[0]
            x2 = struct.unpack_from("<i", data, pos + 16)[0]
            y2 = struct.unpack_from("<i", data, pos + 20)[0]
            # Associate with the nearest preceding Sp
            for sp_off, spid in reversed(sp_records):
                if sp_off < pos:
                    child_anchors[spid] = ChildAnchorInfo(x1, y1, x2, y2)
                    if current_group_spid is not None and spid != current_group_spid:
                        child_to_parent[spid] = current_group_spid
                    break

    return child_anchors, group_info, child_to_parent


def parse_shape_txid_map(table: bytes, dgg_offset: int, dgg_size: int) -> dict[int, int]:
    """Build spid → textbox_story_index mapping from FOpt txid properties.

    The FOpt property ``lTxid`` (pid 0x0080) stores the textbox story
    index as ``(story_index + 1) << 16``.

    Returns ``{shape_id: story_index}`` (0-based).
    """
    data = table[dgg_offset : dgg_offset + dgg_size]

    # Find Sp records
    sp_records: list[tuple[int, int]] = []
    for pos in range(0, len(data) - 12):
        if data[pos + 2] == 0x0A and data[pos + 3] == 0xF0:
            ver = data[pos] & 0x0F
            cb = struct.unpack_from("<I", data, pos + 4)[0]
            if ver == 0x02 and cb == 8:
                spid = struct.unpack_from("<I", data, pos + 8)[0]
                sp_records.append((pos, spid))

    # Find FOpt records with txid (pid 0x0080)
    result: dict[int, int] = {}
    for pos in range(0, len(data) - 8):
        if data[pos + 2] != 0x0B or data[pos + 3] != 0xF0:
            continue
        nprops = (data[pos] | (data[pos + 1] << 8)) >> 4
        cb = struct.unpack_from("<I", data, pos + 4)[0]
        if nprops <= 0 or nprops > 50 or cb < nprops * 6:
            continue

        # Find owning Sp
        owner_spid = None
        for sp_off, spid in reversed(sp_records):
            if sp_off < pos:
                owner_spid = spid
                break
        if owner_spid is None:
            continue

        for p in range(nprops):
            off = pos + 8 + p * 6
            if off + 6 > len(data):
                break
            pid = struct.unpack_from("<H", data, off)[0]
            val = struct.unpack_from("<I", data, off + 2)[0]
            base_pid = pid & 0x3FFF
            if base_pid == 0x0080 and val > 0:
                story_idx = (val >> 16) - 1
                if story_idx >= 0:
                    result[owner_spid] = story_idx
                break  # Only one txid per shape

    return result
