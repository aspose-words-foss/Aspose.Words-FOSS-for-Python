"""
Document Reader - Pure Python DOCX parser (no python-docx dependency).

This module parses DOCX files using only the standard library.
You can replace this with your own implementation.
"""

from __future__ import annotations

import posixpath
import re
import zipfile
from xml.etree import ElementTree as ET
from io import BytesIO
from pathlib import Path
from typing import Optional, Union, BinaryIO, Iterator, TYPE_CHECKING
from dataclasses import dataclass, field

if TYPE_CHECKING:
    from aspose.words_foss import light_document_model as ldm

# =============================================================================
# XML NAMESPACE
# =============================================================================

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
WP_NS = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
PIC_NS = "{http://schemas.openxmlformats.org/drawingml/2006/picture}"
MC_NS = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"

_PKG_RELS_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"

_EXT_TO_CONTENT_TYPE: dict[str, str] = {
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "gif": "image/gif",
    "bmp": "image/bmp",
    "tiff": "image/tiff",
    "tif": "image/tiff",
    "webp": "image/webp",
}

# =============================================================================
# MAPPING CONSTANTS FOR LIGHT DOCUMENT MODEL
# =============================================================================

_ALIGNMENT_MAP = {
    "left": 0,
    "start": 0,
    "center": 1,
    "right": 2,
    "end": 2,
    "both": 3,
    "distribute": 3,
}

_UNDERLINE_MAP = {
    "none": 0,
    "single": 1,
    "words": 2,
    "double": 3,
    "dotted": 4,
    "thick": 6,
    "dash": 7,
    "dotDash": 9,
    "dotDotDash": 10,
    "wavy": 11,
}

_BORDER_STYLE_MAP = {
    "none": 0,
    "nil": 0,
    "single": 1,
    "thick": 2,
    "double": 3,
    "dotted": 4,
    "dashed": 5,
    "dotDash": 6,
    "dotDotDash": 7,
    "triple": 8,
    "thinThickSmallGap": 9,
    "thickThinSmallGap": 10,
    "thinThickMediumGap": 13,
    "thickThinMediumGap": 14,
    "thinThickLargeGap": 16,
    "thickThinLargeGap": 17,
    "wave": 18,
    "doubleWave": 19,
    "dashSmallGap": 20,
    "dashDotStroked": 21,
    "threeDEmboss": 22,
    "threeDEngrave": 23,
    "outset": 24,
    "inset": 25,
}

_STYLE_TYPE_MAP = {"paragraph": 1, "character": 2, "table": 3, "numbering": 4}

_SECTION_START_MAP = {
    "continuous": 0,
    "newColumn": 1,
    "newPage": 2,
    "evenPage": 3,
    "oddPage": 4,
}

_LINE_RULE_MAP = {"atLeast": 0, "exactly": 1, "auto": 2}

_NUMBER_STYLE_MAP = {
    "decimal": 0,
    "upperRoman": 1,
    "lowerRoman": 2,
    "upperLetter": 3,
    "lowerLetter": 4,
    "ordinal": 5,
    "bullet": 23,
    "none": 255,
}


# =============================================================================
# DATA CLASSES
# =============================================================================


@dataclass
class RunData:
    """Text run with formatting."""

    text: str = ""
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strikethrough: bool = False
    style_name: str = ""
    is_code_style: bool = False


@dataclass
class ParagraphData:
    """Paragraph with style and content."""

    text: str = ""
    style_name: str = "Normal"
    runs: list[RunData] = field(default_factory=list)
    is_list_item: bool = False
    list_level: int = 0
    list_id: Optional[int] = None
    alignment: str = "left"
    has_bottom_border: bool = False
    border_size: int = 0


@dataclass
class CellData:
    """Table cell."""

    paragraphs: list[ParagraphData] = field(default_factory=list)
    alignment: str = "left"


@dataclass
class RowData:
    """Table row."""

    cells: list[CellData] = field(default_factory=list)


@dataclass
class TableData:
    """Table structure."""

    rows: list[RowData] = field(default_factory=list)


@dataclass
class NumberingLevel:
    """List level definition."""

    format: str = "bullet"
    start: int = 1
    text: str = ""


@dataclass
class NumberingInfo:
    """Numbering definition."""

    num_id: int
    abstract_num_id: int
    levels: dict[int, NumberingLevel] = field(default_factory=dict)


# =============================================================================
# DOCUMENT READER
# =============================================================================


class DocumentReader:
    """
    Reads DOCX documents and produces abstracted data structures.

    This implementation uses only the standard library (zipfile, xml.etree).
    No python-docx dependency required.
    """

    def __init__(self):
        self._document_xml: Optional[ET.Element] = None
        self._numbering_xml: Optional[ET.Element] = None
        self._styles_xml: Optional[ET.Element] = None
        self._numbering_cache: dict[int, NumberingInfo] = {}
        self._rels: dict[str, str] = {}  # rId -> target URL for hyperlinks
        self._style_id_to_name: dict[str, str] = {}  # style ID → display name
        # Image support
        self._media: dict[str, bytes] = {}  # "word/media/..." -> raw bytes
        self._doc_image_rels: dict[str, str] = {}  # rId -> "word/media/..."
        self._header_data: list[tuple[ET.Element, dict[str, str]]] = []
        self._footer_data: list[tuple[ET.Element, dict[str, str]]] = []

    def load_file(self, filepath: Union[str, Path]) -> None:
        """Load DOCX from file path."""
        with zipfile.ZipFile(str(filepath), "r") as zf:
            self._load_from_zip(zf)

    def load_stream(self, stream: BinaryIO) -> None:
        """Load DOCX from stream."""
        with zipfile.ZipFile(stream, "r") as zf:
            self._load_from_zip(zf)

    def load_bytes(self, data: bytes) -> None:
        """Load DOCX from bytes."""
        with zipfile.ZipFile(BytesIO(data), "r") as zf:
            self._load_from_zip(zf)

    def _load_from_zip(self, zf: zipfile.ZipFile) -> None:
        """Extract and parse XML from DOCX archive."""
        namelist = set(zf.namelist())

        # Load all media files (images)
        for name in namelist:
            if name.startswith("word/media/"):
                self._media[name] = zf.read(name)

        # Parse document.xml
        with zf.open("word/document.xml") as f:
            self._document_xml = ET.parse(f).getroot()

        # Parse styles.xml (optional)
        if "word/styles.xml" in namelist:
            with zf.open("word/styles.xml") as f:
                self._styles_xml = ET.parse(f).getroot()
            self._build_style_id_map()

        # Parse numbering.xml (optional)
        if "word/numbering.xml" in namelist:
            with zf.open("word/numbering.xml") as f:
                self._numbering_xml = ET.parse(f).getroot()
            self._parse_numbering()

        # Parse document relationships (hyperlinks + images + headers/footers)
        rels_path = "word/_rels/document.xml.rels"
        header_targets: list[str] = []
        footer_targets: list[str] = []
        if rels_path in namelist:
            with zf.open(rels_path) as f:
                rels_root = ET.parse(f).getroot()
            for rel in rels_root.findall(f"{_PKG_RELS_NS}Relationship"):
                rid = rel.get("Id", "")
                target = rel.get("Target", "")
                rel_type = rel.get("Type", "")
                target_mode = rel.get("TargetMode", "")
                if not rid or not target:
                    continue
                if target_mode == "External":
                    self._rels[rid] = target
                elif rel_type.endswith("/image"):
                    raw = f"word/{target}" if not target.startswith("/") else target.lstrip("/")
                    self._doc_image_rels[rid] = posixpath.normpath(raw)
                elif rel_type.endswith("/header"):
                    header_targets.append(f"word/{target}")
                elif rel_type.endswith("/footer"):
                    footer_targets.append(f"word/{target}")

        # Parse header XML files and their rels
        for hdr_path in header_targets:
            if hdr_path in namelist:
                with zf.open(hdr_path) as f:
                    hdr_xml = ET.parse(f).getroot()
                hdr_rels = self._parse_part_image_rels(zf, hdr_path, namelist)
                self._header_data.append((hdr_xml, hdr_rels))

        # Parse footer XML files and their rels
        for ftr_path in footer_targets:
            if ftr_path in namelist:
                with zf.open(ftr_path) as f:
                    ftr_xml = ET.parse(f).getroot()
                ftr_rels = self._parse_part_image_rels(zf, ftr_path, namelist)
                self._footer_data.append((ftr_xml, ftr_rels))

    def _parse_numbering(self) -> None:
        """Parse numbering definitions from numbering.xml."""
        self._numbering_cache.clear()
        if self._numbering_xml is None:
            return

        # Parse abstract numbering definitions
        abstract_nums: dict[int, dict[int, NumberingLevel]] = {}
        for abstract in self._numbering_xml.findall(f".//{W_NS}abstractNum"):
            abs_id = int(abstract.get(f"{W_NS}abstractNumId", "0"))
            levels: dict[int, NumberingLevel] = {}

            for lvl in abstract.findall(f"{W_NS}lvl"):
                ilvl = int(lvl.get(f"{W_NS}ilvl", "0"))

                num_fmt = lvl.find(f"{W_NS}numFmt")
                start = lvl.find(f"{W_NS}start")
                lvl_text = lvl.find(f"{W_NS}lvlText")

                levels[ilvl] = NumberingLevel(
                    format=num_fmt.get(f"{W_NS}val", "bullet") if num_fmt is not None else "bullet",
                    start=int(start.get(f"{W_NS}val", "1")) if start is not None else 1,
                    text=lvl_text.get(f"{W_NS}val", "") if lvl_text is not None else "",
                )
            abstract_nums[abs_id] = levels

        # Parse num elements (concrete instances)
        for num in self._numbering_xml.findall(f".//{W_NS}num"):
            num_id = int(num.get(f"{W_NS}numId", "0"))
            abstract_num_id_elem = num.find(f"{W_NS}abstractNumId")

            if abstract_num_id_elem is not None:
                abs_id = int(abstract_num_id_elem.get(f"{W_NS}val", "0"))
                self._numbering_cache[num_id] = NumberingInfo(
                    num_id=num_id, abstract_num_id=abs_id, levels=abstract_nums.get(abs_id, {})
                )

    def _build_style_id_map(self) -> None:
        """Build mapping from style ID to display name from styles.xml."""
        if self._styles_xml is None:
            return
        for style_elem in self._styles_xml.findall(f"{W_NS}style"):
            style_id = style_elem.get(f"{W_NS}styleId", "")
            name_elem = style_elem.find(f"{W_NS}name")
            if name_elem is not None and style_id:
                self._style_id_to_name[style_id] = name_elem.get(f"{W_NS}val", style_id)

    def _resolve_style_name(self, style_id: str) -> str:
        """Resolve a style ID to its display name."""
        return self._style_id_to_name.get(style_id, style_id)

    def _parse_part_image_rels(
        self, zf: zipfile.ZipFile, part_path: str, namelist: set
    ) -> dict[str, str]:
        """Return {rId: media_path} for image rels of a given part (header/footer)."""
        dirname, basename = part_path.rsplit("/", 1)
        rels_path = f"{dirname}/_rels/{basename}.rels"
        result: dict[str, str] = {}
        if rels_path not in namelist:
            return result
        with zf.open(rels_path) as f:
            rels_root = ET.parse(f).getroot()
        for rel in rels_root.findall(f"{_PKG_RELS_NS}Relationship"):
            rid = rel.get("Id", "")
            target = rel.get("Target", "")
            rel_type = rel.get("Type", "")
            if rid and target and rel_type.endswith("/image"):
                full = posixpath.normpath(f"{dirname}/{target}")
                result[rid] = full
        return result

    @staticmethod
    def _ext_to_content_type(filename: str) -> str:
        ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
        return _EXT_TO_CONTENT_TYPE.get(ext, f"image/{ext}")

    def _build_drawing_shape(
        self, drawing_elem: ET.Element, image_rels: dict[str, str]
    ) -> "Optional[ldm.ShapeNode]":
        """Parse a <w:drawing> element and return a ShapeNode if it contains an image."""
        from aspose.words_foss import light_document_model as ldm

        inline = drawing_elem.find(f"{WP_NS}inline")
        anchor = drawing_elem.find(f"{WP_NS}anchor")
        container = inline if inline is not None else anchor
        if container is None:
            return None

        is_inline = inline is not None

        # Dimensions from wp:extent (EMU → points; 1 pt = 12700 EMU)
        extent = container.find(f"{WP_NS}extent")
        width_pt: Optional[float] = None
        height_pt: Optional[float] = None
        if extent is not None:
            cx = extent.get("cx")
            cy = extent.get("cy")
            if cx:
                width_pt = int(cx) / 12700.0
            if cy:
                height_pt = int(cy) / 12700.0

        # Find <a:blip r:embed="rIdN"/>
        blip = container.find(f".//{A_NS}blip")
        if blip is None:
            return None
        r_id = blip.get(f"{R_NS}embed")
        if not r_id:
            return None

        media_path = image_rels.get(r_id)
        if not media_path:
            return None

        image_bytes = self._media.get(media_path)
        if image_bytes is None:
            return None

        # Prefer the original name from pic:cNvPr[@name] if available
        cNvPr = container.find(f".//{PIC_NS}cNvPr")
        if cNvPr is not None and cNvPr.get("name"):
            filename = cNvPr.get("name", "")
        else:
            filename = media_path.rsplit("/", 1)[-1]
        content_type = self._ext_to_content_type(filename)

        shape = ldm.ShapeNode()
        shape.has_image = True
        shape.is_inline = is_inline
        shape.width = width_pt
        shape.height = height_pt
        shape.image_data = ldm.ImageData(
            source_filename=filename,
            content_type=content_type,
            image_bytes=image_bytes,
        )
        return shape

    # -- mc:AlternateContent / w:sdt helpers ----------------------------------

    @staticmethod
    def _resolve_body_children(parent: ET.Element) -> Iterator[ET.Element]:
        """Yield effective body-level children, unwrapping mc:AlternateContent
        and w:sdt elements.

        * ``mc:AlternateContent`` — only the ``mc:Choice`` branch is used;
          ``mc:Fallback`` (which typically contains degraded template /
          placeholder text) is discarded.
        * ``w:sdt`` (Structured Document Tag / Content Control) — the content
          inside ``w:sdtContent`` is unwrapped and its children yielded
          directly.  If the SDT has ``w:showingPlcHdr`` in its properties the
          entire block is skipped because it is placeholder text only visible
          when the content control has not been filled in.
        """
        for child in parent:
            if child.tag == f"{MC_NS}AlternateContent":
                choice = child.find(f"{MC_NS}Choice")
                if choice is not None:
                    yield from DocumentReader._resolve_body_children(choice)
            elif child.tag == f"{W_NS}sdt":
                sdt_pr = child.find(f"{W_NS}sdtPr")
                if sdt_pr is not None and sdt_pr.find(f"{W_NS}showingPlcHdr") is not None:
                    continue  # skip placeholder-only content controls
                sdt_content = child.find(f"{W_NS}sdtContent")
                if sdt_content is not None:
                    yield from DocumentReader._resolve_body_children(sdt_content)
            else:
                yield child

    @staticmethod
    def _resolve_paragraph_children(p_elem: ET.Element) -> Iterator[ET.Element]:
        """Yield effective inline children of a paragraph, unwrapping
        ``mc:AlternateContent`` and inline ``w:sdt`` elements.

        Same semantics as ``_resolve_body_children`` but operates on the
        inline (run-level) children of a ``w:p`` element.
        """
        for child in p_elem:
            if child.tag == f"{MC_NS}AlternateContent":
                choice = child.find(f"{MC_NS}Choice")
                if choice is not None:
                    yield from DocumentReader._resolve_paragraph_children(choice)
            elif child.tag == f"{W_NS}sdt":
                sdt_pr = child.find(f"{W_NS}sdtPr")
                if sdt_pr is not None and sdt_pr.find(f"{W_NS}showingPlcHdr") is not None:
                    continue
                sdt_content = child.find(f"{W_NS}sdtContent")
                if sdt_content is not None:
                    yield from DocumentReader._resolve_paragraph_children(sdt_content)
            else:
                yield child

    def _iterate_body_elements(self) -> Iterator[Union[ParagraphData, TableData]]:
        """Iterate over document body elements in order."""
        if self._document_xml is None:
            return

        body = self._document_xml.find(f"{W_NS}body")
        if body is None:
            return

        for element in self._resolve_body_children(body):
            if element.tag == f"{W_NS}p":
                yield self._parse_paragraph(element)
            elif element.tag == f"{W_NS}tbl":
                yield self._parse_table(element)

    def _parse_paragraph(self, p_elem: ET.Element) -> ParagraphData:
        """Parse a <w:p> paragraph element."""
        data = ParagraphData()

        # Parse paragraph properties
        pPr = p_elem.find(f"{W_NS}pPr")
        if pPr is not None:
            # Style name
            pStyle = pPr.find(f"{W_NS}pStyle")
            if pStyle is not None:
                data.style_name = pStyle.get(f"{W_NS}val", "Normal")

            # Alignment
            jc = pPr.find(f"{W_NS}jc")
            if jc is not None:
                data.alignment = jc.get(f"{W_NS}val", "left")

            # List properties
            numPr = pPr.find(f"{W_NS}numPr")
            if numPr is not None:
                data.is_list_item = True
                ilvl = numPr.find(f"{W_NS}ilvl")
                numId = numPr.find(f"{W_NS}numId")
                if ilvl is not None:
                    data.list_level = int(ilvl.get(f"{W_NS}val", "0"))
                if numId is not None:
                    data.list_id = int(numId.get(f"{W_NS}val", "0"))

            # Border (for horizontal rule detection)
            pBdr = pPr.find(f"{W_NS}pBdr")
            if pBdr is not None:
                bottom = pBdr.find(f"{W_NS}bottom")
                if bottom is not None:
                    sz = bottom.get(f"{W_NS}sz")
                    if sz:
                        data.has_bottom_border = True
                        data.border_size = int(sz)

        # Parse runs and hyperlinks from direct children
        for child in self._resolve_paragraph_children(p_elem):
            if child.tag == f"{W_NS}r":
                run_data = self._parse_run(child)
                if run_data.text:
                    data.runs.append(run_data)
            elif child.tag == f"{W_NS}hyperlink":
                # Get the hyperlink URL from relationships
                r_id = child.get(f"{R_NS}id", "")
                url = self._rels.get(r_id, "")
                # Append anchor/fragment if present
                anchor = child.get(f"{W_NS}anchor", "")
                if anchor:
                    url = url + "#" + anchor if url else "#" + anchor
                # Collect text from runs inside the hyperlink
                link_text_parts = []
                for r_elem in child.findall(f"{W_NS}r"):
                    for t_elem in r_elem.findall(f"{W_NS}t"):
                        link_text_parts.append(t_elem.text or "")
                link_text = "".join(link_text_parts)
                if link_text and url:
                    run_data = RunData(text=f"[{link_text}]({url})")
                    data.runs.append(run_data)
                elif link_text:
                    run_data = RunData(text=link_text)
                    data.runs.append(run_data)

        # Build full text
        data.text = "".join(run.text for run in data.runs)

        return data

    def _parse_run(self, r_elem: ET.Element) -> RunData:
        """Parse a <w:r> run element."""
        data = RunData()

        # Get text from all <w:t> elements
        for t_elem in r_elem.findall(f"{W_NS}t"):
            data.text += t_elem.text or ""

        # Parse run properties
        rPr = r_elem.find(f"{W_NS}rPr")
        if rPr is not None:
            # Bold
            b = rPr.find(f"{W_NS}b")
            if b is not None:
                val = b.get(f"{W_NS}val")
                data.bold = val is None or val not in ("false", "0")

            # Italic
            i = rPr.find(f"{W_NS}i")
            if i is not None:
                val = i.get(f"{W_NS}val")
                data.italic = val is None or val not in ("false", "0")

            # Underline
            u = rPr.find(f"{W_NS}u")
            if u is not None:
                val = u.get(f"{W_NS}val")
                data.underline = val is not None and val != "none"

            # Strikethrough
            strike = rPr.find(f"{W_NS}strike")
            dstrike = rPr.find(f"{W_NS}dstrike")
            if strike is not None:
                val = strike.get(f"{W_NS}val")
                data.strikethrough = val is None or val not in ("false", "0")
            elif dstrike is not None:
                val = dstrike.get(f"{W_NS}val")
                data.strikethrough = val is None or val not in ("false", "0")

            # Character style
            rStyle = rPr.find(f"{W_NS}rStyle")
            if rStyle is not None:
                data.style_name = rStyle.get(f"{W_NS}val", "")
                data.is_code_style = "code" in data.style_name.lower()

        return data

    def _parse_table(self, tbl_elem: ET.Element) -> TableData:
        """Parse a <w:tbl> table element."""
        data = TableData()

        for tr_elem in tbl_elem.findall(f"{W_NS}tr"):
            row = RowData()

            for tc_elem in tr_elem.findall(f"{W_NS}tc"):
                cell = CellData()

                for p_elem in tc_elem.findall(f"{W_NS}p"):
                    cell.paragraphs.append(self._parse_paragraph(p_elem))

                # Get alignment from first paragraph
                if cell.paragraphs:
                    cell.alignment = cell.paragraphs[0].alignment

                row.cells.append(cell)

            data.rows.append(row)

        return data

    def _get_list_format(self, num_id: int, level: int) -> tuple[str, int]:
        """Get list format and start number for a given numId and level."""
        info = self._numbering_cache.get(num_id)
        if info and level in info.levels:
            lvl = info.levels[level]
            return lvl.format, lvl.start
        return "bullet", 1

    def _get_numbering_info(self, num_id: int) -> Optional[NumberingInfo]:
        """Get full numbering info (for advanced usage)."""
        return self._numbering_cache.get(num_id)

    # =========================================================================
    # LIGHT DOCUMENT MODEL BUILDER
    # =========================================================================

    def to_light_document(self) -> ldm.Document:
        """Build a light_document_model.Document from the loaded DOCX.

        Returns a fully populated Pydantic model representing the document
        structure including styles, lists, sections, paragraphs, runs,
        tables, and page setup.
        """
        from aspose.words_foss import light_document_model as ldm

        doc = ldm.Document()

        # Parse styles
        if self._styles_xml is not None:
            doc.styles = self._build_ldm_styles()

        # Parse lists
        if self._numbering_xml is not None:
            doc.lists = self._build_ldm_lists()

        # Parse sections (body + page setup)
        doc.sections = self._build_ldm_sections()

        # Parse header paragraphs — only direct <w:p> children of the root,
        # not paragraphs nested inside table cells within the header.
        hdr_paras: list[ldm.Paragraph] = []
        for hdr_xml, hdr_rels in self._header_data:
            for p_elem in hdr_xml:
                if p_elem.tag == f"{W_NS}p":
                    hdr_paras.append(self._build_ldm_paragraph(p_elem, hdr_rels))
        doc.header_paragraphs = hdr_paras

        # Parse footer paragraphs — same rule: top-level paragraphs only.
        ftr_paras: list[ldm.Paragraph] = []
        for ftr_xml, ftr_rels in self._footer_data:
            for p_elem in ftr_xml:
                if p_elem.tag == f"{W_NS}p":
                    ftr_paras.append(self._build_ldm_paragraph(p_elem, ftr_rels))
        doc.footer_paragraphs = ftr_paras

        return doc

    # -- Styles ---------------------------------------------------------------

    def _build_ldm_styles(self) -> list[ldm.Style]:
        from aspose.words_foss import light_document_model as ldm

        styles: list[ldm.Style] = []
        if self._styles_xml is None:
            return styles

        for style_elem in self._styles_xml.findall(f"{W_NS}style"):
            s = ldm.Style()
            # Name
            name_elem = style_elem.find(f"{W_NS}name")
            if name_elem is not None:
                s.name = name_elem.get(f"{W_NS}val", "")

            # Type
            st = style_elem.get(f"{W_NS}type", "")
            s.type = _STYLE_TYPE_MAP.get(st, 0)

            # Is heading
            heading_match = re.search(r"[Hh]eading\s*(\d+)", s.name)
            s.is_heading = heading_match is not None

            # Base style
            based_on = style_elem.find(f"{W_NS}basedOn")
            if based_on is not None:
                s.base_style_name = based_on.get(f"{W_NS}val", "")

            # Next style
            next_style = style_elem.find(f"{W_NS}next")
            if next_style is not None:
                s.next_paragraph_style_name = next_style.get(f"{W_NS}val", "")

            # Paragraph format
            pPr = style_elem.find(f"{W_NS}pPr")
            if pPr is not None:
                s.paragraph_format = self._build_ldm_paragraph_format(pPr)
                if s.is_heading and heading_match:
                    s.paragraph_format.is_heading = True
                    s.paragraph_format.outline_level = int(heading_match.group(1)) - 1

            # Font (from rPr)
            rPr = style_elem.find(f"{W_NS}rPr")
            if rPr is not None:
                s.font = self._build_ldm_font(rPr)

            styles.append(s)

        return styles

    # -- Lists ----------------------------------------------------------------

    def _build_ldm_lists(self) -> list[ldm.DocList]:
        from aspose.words_foss import light_document_model as ldm

        lists: list[ldm.DocList] = []
        if self._numbering_xml is None:
            return lists

        # Parse abstract numbering definitions
        abstract_defs: dict[int, list[ldm.ListLevel]] = {}
        abstract_multi: dict[int, bool] = {}

        for abstract in self._numbering_xml.findall(f".//{W_NS}abstractNum"):
            abs_id = int(abstract.get(f"{W_NS}abstractNumId", "0"))
            levels: list[ldm.ListLevel] = []
            multi = abstract.get(f"{W_NS}multiLevelType", "")

            for lvl in abstract.findall(f"{W_NS}lvl"):
                ll = ldm.ListLevel()

                num_fmt_elem = lvl.find(f"{W_NS}numFmt")
                num_fmt = (
                    num_fmt_elem.get(f"{W_NS}val", "bullet")
                    if num_fmt_elem is not None
                    else "bullet"
                )
                ll.number_style = _NUMBER_STYLE_MAP.get(num_fmt, 0)

                lvl_text_elem = lvl.find(f"{W_NS}lvlText")
                ll.number_format = (
                    lvl_text_elem.get(f"{W_NS}val", "") if lvl_text_elem is not None else ""
                )

                start_elem = lvl.find(f"{W_NS}start")
                if start_elem is not None:
                    ll.start_at = int(start_elem.get(f"{W_NS}val", "1"))

                # Alignment
                lvl_jc = lvl.find(f"{W_NS}lvlJc")
                if lvl_jc is not None:
                    ll.alignment = _ALIGNMENT_MAP.get(lvl_jc.get(f"{W_NS}val", "left"), 0)

                # Indentation
                pPr = lvl.find(f"{W_NS}pPr")
                if pPr is not None:
                    ind = pPr.find(f"{W_NS}ind")
                    if ind is not None:
                        left = ind.get(f"{W_NS}left", "")
                        if left:
                            ll.number_position = int(left) / 20.0
                        hanging = ind.get(f"{W_NS}hanging", "")
                        first_line = ind.get(f"{W_NS}firstLine", "")
                        if left and hanging:
                            ll.text_position = int(left) / 20.0
                            ll.number_position = (int(left) - int(hanging)) / 20.0
                        elif first_line:
                            ll.text_position = int(left) / 20.0 if left else 0.0

                levels.append(ll)

            abstract_defs[abs_id] = levels
            abstract_multi[abs_id] = multi in (
                "multilevel",
                "hybridMultilevel",
            )

        # Map num -> abstract
        for num in self._numbering_xml.findall(f".//{W_NS}num"):
            num_id = int(num.get(f"{W_NS}numId", "0"))
            abs_id_elem = num.find(f"{W_NS}abstractNumId")
            if abs_id_elem is None:
                continue
            abs_id = int(abs_id_elem.get(f"{W_NS}val", "0"))

            dl = ldm.DocList()
            dl.list_id = num_id
            dl.levels = list(abstract_defs.get(abs_id, []))
            dl.is_multi_level = abstract_multi.get(abs_id, False)
            lists.append(dl)

        return lists

    # -- Sections -------------------------------------------------------------

    def _build_ldm_sections(self) -> list[ldm.Section]:
        from aspose.words_foss import light_document_model as ldm

        sections: list[ldm.Section] = []
        if self._document_xml is None:
            return sections

        body = self._document_xml.find(f"{W_NS}body")
        if body is None:
            return sections

        # Collect body children, splitting at sectPr boundaries
        # In DOCX, sections are delimited by w:sectPr inside w:pPr (for
        # non-final sections) and by w:sectPr directly in w:body (final
        # section).
        current_children: list[ldm.Paragraph | ldm.Table | ldm.UnknownNode] = []
        section_props: list[Optional[ET.Element]] = []

        for element in self._resolve_body_children(body):
            if element.tag == f"{W_NS}p":
                para = self._build_ldm_paragraph(element)
                current_children.append(para)
                # Check for section break in paragraph properties
                pPr = element.find(f"{W_NS}pPr")
                if pPr is not None:
                    sect_pr = pPr.find(f"{W_NS}sectPr")
                    if sect_pr is not None:
                        section_props.append(sect_pr)
                        sec = ldm.Section()
                        sec.page_setup = self._build_ldm_page_setup(sect_pr)
                        sec.body = ldm.Body(children=current_children)
                        sections.append(sec)
                        current_children = []
            elif element.tag == f"{W_NS}tbl":
                tbl = self._build_ldm_table(element)
                current_children.append(tbl)
            elif element.tag == f"{W_NS}sectPr":
                # Final section properties
                sec = ldm.Section()
                sec.page_setup = self._build_ldm_page_setup(element)
                sec.body = ldm.Body(children=current_children)
                sections.append(sec)
                current_children = []

        # If there are remaining children without a sectPr, add a default section
        if current_children:
            sec = ldm.Section()
            sec.body = ldm.Body(children=current_children)
            sections.append(sec)

        return sections

    # -- Page Setup -----------------------------------------------------------

    def _build_ldm_page_setup(self, sect_pr: ET.Element) -> ldm.PageSetup:
        from aspose.words_foss import light_document_model as ldm

        ps = ldm.PageSetup()

        pg_sz = sect_pr.find(f"{W_NS}pgSz")
        if pg_sz is not None:
            w = pg_sz.get(f"{W_NS}w", "")
            h = pg_sz.get(f"{W_NS}h", "")
            if w:
                ps.page_width = int(w) / 20.0
            if h:
                ps.page_height = int(h) / 20.0
            orient = pg_sz.get(f"{W_NS}orient", "")
            ps.orientation = 1 if orient == "landscape" else 0
            # Derive paper_size from dimensions
            ps.paper_size = self._detect_paper_size(ps.page_width, ps.page_height)

        pg_mar = sect_pr.find(f"{W_NS}pgMar")
        if pg_mar is not None:
            for attr, field_name in [
                ("top", "top_margin"),
                ("bottom", "bottom_margin"),
                ("left", "left_margin"),
                ("right", "right_margin"),
                ("header", "header_distance"),
                ("footer", "footer_distance"),
            ]:
                val = pg_mar.get(f"{W_NS}{attr}", "")
                if val:
                    setattr(ps, field_name, int(val) / 20.0)

        type_elem = sect_pr.find(f"{W_NS}type")
        if type_elem is not None:
            ps.section_start = _SECTION_START_MAP.get(type_elem.get(f"{W_NS}val", ""), 2)

        title_pg = sect_pr.find(f"{W_NS}titlePg")
        if title_pg is not None:
            ps.different_first_page_header_footer = True

        return ps

    @staticmethod
    def _detect_paper_size(width: float, height: float) -> int:
        """Detect standard paper size from dimensions in points."""
        w, h = min(width, height), max(width, height)
        # Letter: 612 x 792
        if abs(w - 612) < 2 and abs(h - 792) < 2:
            return 1
        # Legal: 612 x 1008
        if abs(w - 612) < 2 and abs(h - 1008) < 2:
            return 5
        # A4: 595.28 x 841.89
        if abs(w - 595.28) < 3 and abs(h - 841.89) < 3:
            return 9
        # A3: 841.89 x 1190.55
        if abs(w - 841.89) < 3 and abs(h - 1190.55) < 3:
            return 8
        return 0

    # -- Paragraph ------------------------------------------------------------

    def _build_ldm_paragraph(
        self,
        p_elem: ET.Element,
        image_rels: Optional[dict[str, str]] = None,
    ) -> ldm.Paragraph:
        from aspose.words_foss import light_document_model as ldm

        # Default to document-level image relationships
        if image_rels is None:
            image_rels = self._doc_image_rels

        para = ldm.Paragraph()

        pPr = p_elem.find(f"{W_NS}pPr")
        if pPr is not None:
            para.paragraph_format = self._build_ldm_paragraph_format(pPr)
            # If heading not detected from outlineLvl, infer from style name
            pf = para.paragraph_format
            if not pf.is_heading and pf.style_name:
                heading_match = re.search(r"[Hh]eading\s*(\d+)", pf.style_name)
                if heading_match:
                    pf.is_heading = True
                    pf.outline_level = int(heading_match.group(1)) - 1

            # List format (numId=0 means explicitly no list)
            numPr = pPr.find(f"{W_NS}numPr")
            if numPr is not None:
                numId_elem = numPr.find(f"{W_NS}numId")
                num_id_val = int(numId_elem.get(f"{W_NS}val", "0")) if numId_elem is not None else 0
                if num_id_val > 0:
                    lf = ldm.ListFormat()
                    lf.is_list_item = True
                    lf.list_id = num_id_val
                    ilvl = numPr.find(f"{W_NS}ilvl")
                    if ilvl is not None:
                        lf.list_level_number = int(ilvl.get(f"{W_NS}val", "0"))
                    para.list_format = lf
                    para.paragraph_format.is_list_item = True

        # Parse runs and hyperlinks
        text_parts: list[str] = []
        for child in self._resolve_paragraph_children(p_elem):
            if child.tag == f"{W_NS}r":
                # Extract shapes from any embedded <w:drawing> elements.
                for drawing in child.findall(f"{W_NS}drawing"):
                    shape = self._build_drawing_shape(drawing, image_rels)
                    if shape is not None:
                        para.inline_extras.append(shape)
                        para.content_sequence.append(shape)
                # Always extract run text — a <w:r> may contain both <w:drawing>
                # and <w:t> children, and non-image drawings (charts, SmartArt,
                # text boxes) must not cause the run's text content to be lost.
                run = self._build_ldm_run(child)
                if run.text:
                    para.runs.append(run)
                    para.content_sequence.append(run)
                    text_parts.append(run.text)
            elif child.tag == f"{W_NS}hyperlink":
                r_id = child.get(f"{R_NS}id", "")
                url = self._rels.get(r_id, "")
                anchor = child.get(f"{W_NS}anchor", "")
                if anchor:
                    url = url + "#" + anchor if url else "#" + anchor
                link_text_parts: list[str] = []
                for r_elem in child.findall(f"{W_NS}r"):
                    for t_elem in r_elem.findall(f"{W_NS}t"):
                        link_text_parts.append(t_elem.text or "")
                link_text = "".join(link_text_parts)
                if link_text and url:
                    run = ldm.Run()
                    run.text = f"[{link_text}]({url})"
                    para.runs.append(run)
                    para.content_sequence.append(run)
                    text_parts.append(run.text)
                elif link_text:
                    run = ldm.Run()
                    run.text = link_text
                    para.runs.append(run)
                    para.content_sequence.append(run)
                    text_parts.append(run.text)

        para.text = "".join(text_parts)
        return para

    # -- Paragraph Format -----------------------------------------------------

    def _build_ldm_paragraph_format(self, pPr: ET.Element) -> ldm.ParagraphFormat:
        from aspose.words_foss import light_document_model as ldm

        pf = ldm.ParagraphFormat()

        # Style name (resolve ID to display name)
        pStyle = pPr.find(f"{W_NS}pStyle")
        if pStyle is not None:
            style_id = pStyle.get(f"{W_NS}val", "")
            pf.style_name = self._resolve_style_name(style_id)

        # Alignment
        jc = pPr.find(f"{W_NS}jc")
        if jc is not None:
            pf.alignment = _ALIGNMENT_MAP.get(jc.get(f"{W_NS}val", "left"), 0)

        # Indentation
        ind = pPr.find(f"{W_NS}ind")
        if ind is not None:
            left = ind.get(f"{W_NS}left", "")
            if left:
                pf.left_indent = int(left) / 20.0
            right = ind.get(f"{W_NS}right", "")
            if right:
                pf.right_indent = int(right) / 20.0
            first_line = ind.get(f"{W_NS}firstLine", "")
            hanging = ind.get(f"{W_NS}hanging", "")
            if first_line:
                pf.first_line_indent = int(first_line) / 20.0
            elif hanging:
                pf.first_line_indent = -int(hanging) / 20.0

        # Spacing
        spacing = pPr.find(f"{W_NS}spacing")
        if spacing is not None:
            before = spacing.get(f"{W_NS}before", "")
            if before:
                pf.space_before = int(before) / 20.0
            after = spacing.get(f"{W_NS}after", "")
            if after:
                pf.space_after = int(after) / 20.0
            line = spacing.get(f"{W_NS}line", "")
            if line:
                pf.line_spacing = int(line) / 20.0
            line_rule = spacing.get(f"{W_NS}lineRule", "")
            if line_rule:
                pf.line_spacing_rule = _LINE_RULE_MAP.get(line_rule, 0)
            if spacing.get(f"{W_NS}beforeAutospacing", "") in (
                "1",
                "true",
            ):
                pf.space_before_auto = True
            if spacing.get(f"{W_NS}afterAutospacing", "") in (
                "1",
                "true",
            ):
                pf.space_after_auto = True

        # Keep with next
        kwn = pPr.find(f"{W_NS}keepNext")
        if kwn is not None:
            val = kwn.get(f"{W_NS}val")
            pf.keep_with_next = val is None or val not in ("false", "0")

        # Page break before
        pbb = pPr.find(f"{W_NS}pageBreakBefore")
        if pbb is not None:
            val = pbb.get(f"{W_NS}val")
            pf.page_break_before = val is None or val not in ("false", "0")

        # Outline level
        outline = pPr.find(f"{W_NS}outlineLvl")
        if outline is not None:
            ol_val = outline.get(f"{W_NS}val", "9")
            pf.outline_level = int(ol_val)
            if pf.outline_level < 9:
                pf.is_heading = True

        # Shading
        shd = pPr.find(f"{W_NS}shd")
        if shd is not None:
            pf.shading = self._build_ldm_shading(shd)

        # Borders
        pBdr = pPr.find(f"{W_NS}pBdr")
        if pBdr is not None:
            pf.borders = self._build_ldm_borders(pBdr)

        return pf

    # -- Run / Font -----------------------------------------------------------

    def _build_ldm_run(self, r_elem: ET.Element) -> ldm.Run:
        from aspose.words_foss import light_document_model as ldm

        run = ldm.Run()
        # Collect text
        parts: list[str] = []
        for t_elem in r_elem.findall(f"{W_NS}t"):
            parts.append(t_elem.text or "")
        run.text = "".join(parts)

        # Font from rPr
        rPr = r_elem.find(f"{W_NS}rPr")
        if rPr is not None:
            run.font = self._build_ldm_font(rPr)

        return run

    def _build_ldm_font(self, rPr: ET.Element) -> ldm.Font:
        from aspose.words_foss import light_document_model as ldm

        font = ldm.Font()

        # Font name
        rFonts = rPr.find(f"{W_NS}rFonts")
        if rFonts is not None:
            font.name = (
                rFonts.get(f"{W_NS}ascii", "")
                or rFonts.get(f"{W_NS}hAnsi", "")
                or rFonts.get(f"{W_NS}cs", "")
            )

        # Size (half-points → points)
        sz = rPr.find(f"{W_NS}sz")
        if sz is not None:
            val = sz.get(f"{W_NS}val", "")
            if val:
                font.size = int(val) / 2.0

        # Bold
        b = rPr.find(f"{W_NS}b")
        if b is not None:
            val = b.get(f"{W_NS}val")
            font.bold = val is None or val not in ("false", "0")

        # Italic
        i = rPr.find(f"{W_NS}i")
        if i is not None:
            val = i.get(f"{W_NS}val")
            font.italic = val is None or val not in ("false", "0")

        # Underline
        u = rPr.find(f"{W_NS}u")
        if u is not None:
            u_val = u.get(f"{W_NS}val", "none")
            font.underline = _UNDERLINE_MAP.get(u_val, 1 if u_val != "none" else 0)

        # Color
        color = rPr.find(f"{W_NS}color")
        if color is not None:
            font.color = color.get(f"{W_NS}val", "")

        # Strikethrough
        strike = rPr.find(f"{W_NS}strike")
        if strike is not None:
            val = strike.get(f"{W_NS}val")
            font.strike_through = val is None or val not in ("false", "0")

        # Superscript / subscript
        vert_align = rPr.find(f"{W_NS}vertAlign")
        if vert_align is not None:
            va_val = vert_align.get(f"{W_NS}val", "")
            font.superscript = va_val == "superscript"
            font.subscript = va_val == "subscript"

        # Highlight color
        highlight = rPr.find(f"{W_NS}highlight")
        if highlight is not None:
            font.highlight_color = highlight.get(f"{W_NS}val", "")

        # Caps
        caps = rPr.find(f"{W_NS}caps")
        if caps is not None:
            val = caps.get(f"{W_NS}val")
            font.all_caps = val is None or val not in ("false", "0")

        small_caps = rPr.find(f"{W_NS}smallCaps")
        if small_caps is not None:
            val = small_caps.get(f"{W_NS}val")
            font.small_caps = val is None or val not in ("false", "0")

        # Hidden
        vanish = rPr.find(f"{W_NS}vanish")
        if vanish is not None:
            val = vanish.get(f"{W_NS}val")
            font.hidden = val is None or val not in ("false", "0")

        # Style name (character style - resolve ID to display name)
        rStyle = rPr.find(f"{W_NS}rStyle")
        if rStyle is not None:
            style_id = rStyle.get(f"{W_NS}val", "")
            font.style_name = self._resolve_style_name(style_id)

        # Shading
        shd = rPr.find(f"{W_NS}shd")
        if shd is not None:
            font.shading = self._build_ldm_shading(shd)

        return font

    # -- Table ----------------------------------------------------------------

    def _build_ldm_table(self, tbl_elem: ET.Element) -> ldm.Table:
        from aspose.words_foss import light_document_model as ldm

        tbl = ldm.Table()

        # Table properties
        tblPr = tbl_elem.find(f"{W_NS}tblPr")
        if tblPr is not None:
            jc = tblPr.find(f"{W_NS}jc")
            if jc is not None:
                tbl.alignment = _ALIGNMENT_MAP.get(jc.get(f"{W_NS}val", "left"), 0)

            tblW = tblPr.find(f"{W_NS}tblW")
            if tblW is not None:
                w_type = tblW.get(f"{W_NS}type", "")
                w_val = tblW.get(f"{W_NS}w", "0")
                if w_type == "pct":
                    tbl.preferred_width = f"{int(w_val) / 50}%"
                elif w_type == "dxa":
                    tbl.preferred_width = f"{int(w_val) / 20.0}pt"
                elif w_type == "auto":
                    tbl.preferred_width = "Auto"

            tblInd = tblPr.find(f"{W_NS}tblInd")
            if tblInd is not None:
                val = tblInd.get(f"{W_NS}w", "0")
                tbl.left_indent = int(val) / 20.0

            # Table cell margins (default paddings)
            tblCellMar = tblPr.find(f"{W_NS}tblCellMar")
            if tblCellMar is not None:
                for side, attr in [
                    ("left", "left_padding"),
                    ("right", "right_padding"),
                    ("top", "top_padding"),
                    ("bottom", "bottom_padding"),
                ]:
                    side_elem = tblCellMar.find(f"{W_NS}{side}")
                    if side_elem is not None:
                        val = side_elem.get(f"{W_NS}w", "0")
                        setattr(tbl, attr, int(val) / 20.0)

        # Rows
        for tr_elem in tbl_elem.findall(f"{W_NS}tr"):
            row = self._build_ldm_row(tr_elem)
            tbl.rows.append(row)

        return tbl

    def _build_ldm_row(self, tr_elem: ET.Element) -> ldm.Row:
        from aspose.words_foss import light_document_model as ldm

        row = ldm.Row()

        trPr = tr_elem.find(f"{W_NS}trPr")
        if trPr is not None:
            rf = ldm.RowFormat()
            trHeight = trPr.find(f"{W_NS}trHeight")
            if trHeight is not None:
                val = trHeight.get(f"{W_NS}val", "0")
                rf.height = int(val) / 20.0
                rule = trHeight.get(f"{W_NS}hRule", "")
                if rule == "exact":
                    rf.height_rule = 1
                elif rule == "auto":
                    rf.height_rule = 2
                else:
                    rf.height_rule = 0

            tblHeader = trPr.find(f"{W_NS}tblHeader")
            if tblHeader is not None:
                rf.heading_format = True

            cant_split = trPr.find(f"{W_NS}cantSplit")
            if cant_split is not None:
                rf.allow_break_across_pages = False

            row.row_format = rf

        for tc_elem in tr_elem.findall(f"{W_NS}tc"):
            cell = self._build_ldm_cell(tc_elem)
            row.cells.append(cell)

        return row

    def _build_ldm_cell(self, tc_elem: ET.Element) -> ldm.Cell:
        from aspose.words_foss import light_document_model as ldm

        cell = ldm.Cell()

        tcPr = tc_elem.find(f"{W_NS}tcPr")
        if tcPr is not None:
            cf = ldm.CellFormat()

            tcW = tcPr.find(f"{W_NS}tcW")
            if tcW is not None:
                w_type = tcW.get(f"{W_NS}type", "")
                w_val = tcW.get(f"{W_NS}w", "0")
                if w_type == "dxa":
                    cf.width = int(w_val) / 20.0
                elif w_type == "pct":
                    cf.preferred_width = f"{int(w_val) / 50}%"
                elif w_type == "auto":
                    cf.preferred_width = "Auto"

            vAlign = tcPr.find(f"{W_NS}vAlign")
            if vAlign is not None:
                va = vAlign.get(f"{W_NS}val", "top")
                cf.vertical_alignment = {"top": 0, "center": 1, "bottom": 2}.get(va, 0)

            vMerge = tcPr.find(f"{W_NS}vMerge")
            if vMerge is not None:
                val = vMerge.get(f"{W_NS}val", "continue")
                cf.vertical_merge = 1 if val == "restart" else 2

            gridSpan = tcPr.find(f"{W_NS}gridSpan")
            if gridSpan is not None:
                span = int(gridSpan.get(f"{W_NS}val", "1"))
                if span > 1:
                    cf.horizontal_merge = 1

            # Cell margins
            tcMar = tcPr.find(f"{W_NS}tcMar")
            if tcMar is not None:
                for side, attr in [
                    ("left", "left_padding"),
                    ("right", "right_padding"),
                    ("top", "top_padding"),
                    ("bottom", "bottom_padding"),
                ]:
                    side_elem = tcMar.find(f"{W_NS}{side}")
                    if side_elem is not None:
                        val = side_elem.get(f"{W_NS}w", "0")
                        setattr(cf, attr, int(val) / 20.0)

            # Shading
            shd = tcPr.find(f"{W_NS}shd")
            if shd is not None:
                cf.shading = self._build_ldm_shading(shd)

            # Borders
            tcBorders = tcPr.find(f"{W_NS}tcBorders")
            if tcBorders is not None:
                cf.borders = self._build_ldm_borders(tcBorders)

            cell.cell_format = cf

        # Paragraphs in cell
        for p_elem in tc_elem.findall(f"{W_NS}p"):
            cell.paragraphs.append(self._build_ldm_paragraph(p_elem))

        # Nested tables
        for tbl_elem in tc_elem.findall(f"{W_NS}tbl"):
            cell.tables.append(self._build_ldm_table(tbl_elem))

        return cell

    # -- Shared helpers -------------------------------------------------------

    def _build_ldm_shading(self, shd: ET.Element) -> ldm.Shading:
        from aspose.words_foss import light_document_model as ldm

        s = ldm.Shading()
        fill = shd.get(f"{W_NS}fill", "")
        if fill:
            s.background_color = fill
        return s

    def _build_ldm_borders(self, bdr_elem: ET.Element) -> list[ldm.Border]:
        from aspose.words_foss import light_document_model as ldm

        borders: list[ldm.Border] = []
        # Standard border sides in order: top, left, bottom, right, between, bar
        for side in ("top", "left", "bottom", "right", "between", "bar"):
            side_elem = bdr_elem.find(f"{W_NS}{side}")
            if side_elem is not None:
                b = ldm.Border()
                style_val = side_elem.get(f"{W_NS}val", "none")
                b.line_style = _BORDER_STYLE_MAP.get(style_val, 0)
                sz = side_elem.get(f"{W_NS}sz", "")
                if sz:
                    b.line_width = int(sz) / 8.0
                color = side_elem.get(f"{W_NS}color", "")
                if color:
                    b.color = color
                borders.append(b)
            else:
                borders.append(ldm.Border())
        return borders
