"""
FIB (File Information Block) parser for Word 97-2003 binary format.
"""

import struct


class FibData:
    """Parsed FIB (File Information Block) data."""

    def __init__(self):
        self.f_which_tbl_stm: bool = False
        self.ccp_text: int = 0
        self.ccp_ftn: int = 0
        self.ccp_hdd: int = 0
        self.ccp_txbx: int = 0
        self.fc_start: int = 0  # Start of FibRgFcLcb in WordDocument stream
        self.cb_rg_fc_lcb: int = 0


def parse_fib(wd: bytes) -> FibData:
    """Parse the FIB from the WordDocument stream."""
    fib = FibData()

    # Check magic number
    magic = struct.unpack_from("<H", wd, 0)[0]
    if magic != 0xA5EC:
        raise ValueError(f"Not a valid Word document (magic=0x{magic:04X})")

    # Flags at offset 10
    flags = struct.unpack_from("<H", wd, 10)[0]
    fib.f_which_tbl_stm = bool(flags & 0x0200)

    # FibRgW97
    csw = struct.unpack_from("<H", wd, 32)[0]
    rgw_end = 34 + csw * 2

    # FibRgLw97
    cslw = struct.unpack_from("<H", wd, rgw_end)[0]
    lw_start = rgw_end + 2
    fib.ccp_text = struct.unpack_from("<I", wd, lw_start + 12)[0]
    fib.ccp_ftn = struct.unpack_from("<I", wd, lw_start + 16)[0]
    fib.ccp_hdd = struct.unpack_from("<I", wd, lw_start + 20)[0]
    fib.ccp_txbx = struct.unpack_from("<I", wd, lw_start + 36)[0]
    lw_end = lw_start + cslw * 4

    # FibRgFcLcb
    fib.cb_rg_fc_lcb = struct.unpack_from("<H", wd, lw_end)[0]
    fib.fc_start = lw_end + 2

    return fib


def get_fc_lcb(wd: bytes, fib: FibData, idx: int) -> tuple[int, int]:
    """Get an (fc, lcb) pair from FibRgFcLcb at the given index."""
    if idx >= fib.cb_rg_fc_lcb:
        return 0, 0
    offset = fib.fc_start + idx * 8
    fc = struct.unpack_from("<I", wd, offset)[0]
    lcb = struct.unpack_from("<I", wd, offset + 4)[0]
    return fc, lcb
