"""Best-effort local image loading for Markdown ``![]()`` import."""

from __future__ import annotations

import base64
import binascii
import re
import struct
from pathlib import Path
from typing import Optional
from urllib.parse import urlparse

from aspose.words_foss.model.enums.image import ImageType

_EXT_TO_IMAGE_TYPE = {
    "png": ImageType.PNG, "jpg": ImageType.JPEG, "jpeg": ImageType.JPEG,
    "gif": ImageType.GIF, "bmp": ImageType.BMP, "tif": ImageType.TIFF,
    "tiff": ImageType.TIFF, "svg": ImageType.SVG, "webp": ImageType.WEB_P,
    "emf": ImageType.EMF, "wmf": ImageType.WMF,
}

_MIME_TO_IMAGE_TYPE = {
    "image/png": ImageType.PNG, "image/jpeg": ImageType.JPEG,
    "image/gif": ImageType.GIF, "image/bmp": ImageType.BMP,
    "image/tiff": ImageType.TIFF, "image/webp": ImageType.WEB_P,
}

# "data:<mediatype>;base64,<data>" -- RFC 2397, restricted to the ";base64," form.
_DATA_URI_RE = re.compile(r"^data:(image/[a-zA-Z0-9.+-]+);base64,(.*)$", re.DOTALL)


def is_local_path(uri: str) -> bool:
    """Whether *uri* looks like a local file path rather than a remote/data URL."""
    scheme = urlparse(uri).scheme
    return len(scheme) <= 1 or scheme == "file"  # "" or a Windows drive letter counts as local


def guess_image_type(path: str) -> int:
    return _EXT_TO_IMAGE_TYPE.get(Path(path).suffix.lstrip(".").lower(), ImageType.UNKNOWN)


def sniff_dimensions(data: bytes) -> tuple[Optional[float], Optional[float]]:
    """Best-effort width/height in points (assuming 96 DPI) from a PNG/GIF/JPEG header."""
    pixels = _sniff_pixel_size(data)
    if pixels is None:
        return None, None
    width, height = pixels
    return width * 72.0 / 96.0, height * 72.0 / 96.0


def _sniff_pixel_size(data: bytes) -> Optional[tuple[int, int]]:
    if data[:8] == b"\x89PNG\r\n\x1a\n" and len(data) >= 24:
        return struct.unpack(">II", data[16:24])
    if data[:6] in (b"GIF87a", b"GIF89a") and len(data) >= 10:
        return struct.unpack("<HH", data[6:10])
    if data[:2] == b"\xff\xd8":
        i, n = 2, len(data)
        while i + 9 < n:
            if data[i] != 0xFF:
                i += 1
                continue
            marker = data[i + 1]
            if 0xC0 <= marker <= 0xC3:
                height, width = struct.unpack(">HH", data[i + 5 : i + 9])
                return width, height
            if marker in (0xD8, 0xD9) or 0xD0 <= marker <= 0xD7:
                i += 2
                continue
            if i + 4 > n:
                break
            seg_len = struct.unpack(">H", data[i + 2 : i + 4])[0]
            i += 2 + seg_len
    return None


def load_local_image(uri: str, base_dir: Optional[Path]) -> Optional[tuple[bytes, int, Optional[float], Optional[float]]]:
    """Read *uri* (resolved against *base_dir*) if it's a local, existing, readable file.

    Returns (bytes, ImageType, width_pt, height_pt), or ``None`` if it can't be loaded."""
    if not is_local_path(uri) or base_dir is None:
        return None
    path = Path(uri)
    full_path = path if path.is_absolute() else base_dir / path
    try:
        data = full_path.read_bytes()
    except OSError:
        return None
    image_type = guess_image_type(uri)
    width, height = sniff_dimensions(data)
    return data, image_type, width, height


def load_data_uri_image(uri: str) -> Optional[tuple[bytes, int, Optional[float], Optional[float]]]:
    """Decode a ``data:image/<type>;base64,<data>`` URI. No network access is involved --
    the bytes are already inline in the source text, same as a local file's own bytes.

    Returns (bytes, ImageType, width_pt, height_pt), or ``None`` if it isn't a
    recognized base64-encoded raster image (SVG data URIs are left as literal text,
    matching the source's existing behavior for a non-local, non-decoded URI).
    """
    match = _DATA_URI_RE.match(uri)
    if match is None:
        return None
    mime, payload = match.group(1).lower(), match.group(2)
    image_type = _MIME_TO_IMAGE_TYPE.get(mime)
    if image_type is None:
        return None
    try:
        data = base64.b64decode(payload, validate=True)
    except (binascii.Error, ValueError):
        return None
    if not data:
        return None
    width, height = sniff_dimensions(data)
    return data, image_type, width, height
