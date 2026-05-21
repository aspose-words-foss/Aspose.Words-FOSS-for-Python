"""Loader for the Aspose blank docx template.

Holds zero state — every accessor extracts the requested block from
the bundled ``resources/blank.docx`` archive on demand and lru-caches
the parsed strings.  Keeps the binary out of the importable namespace
and lets the template be swapped by replacing the file.
"""


import re
import zipfile
from functools import lru_cache
from importlib import resources


def _read_part(part_name: str) -> str:
    """Read ``part_name`` from the bundled blank docx zip as text."""
    pkg = resources.files(__package__) / "resources" / "blank.docx"
    with pkg.open("rb") as fp, zipfile.ZipFile(fp) as zf:
        return zf.read(part_name).decode("utf-8")


@lru_cache(maxsize=1)
def latent_styles() -> str:
    """Return the ``<w:latentStyles>`` block from blank/styles.xml."""
    return _extract_latent_styles(_read_part("word/styles.xml"))


def _extract_latent_styles(blob: str) -> str:
    """Return the ``<w:latentStyles>`` element verbatim, both CT_LatentStyles
    forms (self-closing and content-bearing) supported."""
    m = re.search(r"<w:latentStyles\b[^>]*?/>", blob)
    if m:
        return m.group(0)
    m = re.search(r"<w:latentStyles\b[^>]*?>.*?</w:latentStyles>", blob, re.DOTALL)
    return m.group(0) if m else ""


@lru_cache(maxsize=1)
def settings_compat() -> str:
    """Return the ``<w:compat>`` block from blank/settings.xml."""
    blob = _read_part("word/settings.xml")
    m = re.search(r"<w:compat>.*?</w:compat>", blob, re.DOTALL)
    return m.group(0) if m else ""


@lru_cache(maxsize=1)
def font_definitions() -> str:
    """Return ``<w:font>`` children of ``<w:fonts>`` from blank/fontTable.xml."""
    blob = _read_part("word/fontTable.xml")
    m = re.search(r"<w:fonts[^>]*>(.*)</w:fonts>", blob, re.DOTALL)
    if not m:
        return ""
    return re.sub(r">\s+<", "><", m.group(1).strip())
