"""
Core DocumentReader class for DOCX parsing.

Reads DOCX documents and produces abstracted data structures using
only the standard library (zipfile, xml.etree).  LDM construction
and shape parsing are provided by the LdmBuilderMixin and
ShapeParserMixin, respectively.
"""

from __future__ import annotations

import posixpath
import zipfile
from io import BytesIO
from pathlib import Path
from typing import Optional, Union, BinaryIO, Iterator, TYPE_CHECKING
from xml.etree import ElementTree as ET

from aspose.words_foss.docx_reader.constants import (
    A_NS,
    MC_NS,
    R_NS,
    W_NS,
    _PKG_RELS_NS,
)
from aspose.words_foss.docx_reader.data_classes import (
    CellData,
    NumberingInfo,
    NumberingLevel,
    ParagraphData,
    RowData,
    RunData,
    TableData,
)
from aspose.words_foss.docx_reader.utils import (
    _apply_theme_color_modifiers,
    _canonicalize_style_name,
)
from aspose.words_foss.docx_reader.ldm_builder import LdmBuilderMixin
from aspose.words_foss.docx_reader.shapes import ShapeParserMixin

if TYPE_CHECKING:
    from aspose.words_foss import light_document_model as ldm


class DocumentReader(LdmBuilderMixin, ShapeParserMixin):
    """
    Reads DOCX documents and produces abstracted data structures.

    This implementation uses only the standard library (zipfile, xml.etree).
    No python-docx dependency required.
    """

    def __init__(self):
        self._document_xml: Optional[ET.Element] = None
        self._numbering_xml: Optional[ET.Element] = None
        self._styles_xml: Optional[ET.Element] = None
        self._settings_xml: Optional[ET.Element] = None
        self._numbering_cache: dict[int, NumberingInfo] = {}
        self._rels: dict[str, str] = {}  # rId -> target URL for hyperlinks
        self._style_id_to_name: dict[str, str] = {}  # style ID → display name
        # Image support
        self._media: dict[str, bytes] = {}  # "word/media/..." -> raw bytes
        self._doc_image_rels: dict[str, str] = {}  # rId -> "word/media/..."
        self._header_data: list[tuple[ET.Element, dict[str, str]]] = []
        self._footer_data: list[tuple[ET.Element, dict[str, str]]] = []
        # Theme font mapping: theme name -> resolved font name
        self._theme_fonts: dict[str, str] = {}
        # Theme color mapping: scheme color name (accent1, dk1, lt1, hlink,
        # bg1, tx1…) -> "RRGGBB" hex.  Populated from theme1.xml/clrScheme
        # so ``a:schemeClr`` references in shape fills, borders and run
        # colors can be resolved to concrete RGB values.
        self._theme_colors: dict[str, str] = {}
        # Document defaults from styles.xml/docDefaults
        self._doc_default_rPr: Optional[ET.Element] = None
        self._doc_default_pPr: Optional[ET.Element] = None
        # First-section page setup, cached during body parsing so that
        # anchored-shape extraction can resolve margin-/column-relative
        # offsets into absolute page coordinates.
        self._current_page_setup: "Optional[ldm.PageSetup]" = None
        # Cached styleId → <w:style> element for O(1) lookups
        self._style_elem_cache: dict[str, ET.Element] = {}
        # Reverse map: display name → styleId (for style chain resolution)
        self._name_to_style_id: dict[str, str] = {}

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

        # Parse theme (for font resolution)
        if "word/theme/theme1.xml" in namelist:
            with zf.open("word/theme/theme1.xml") as f:
                self._parse_theme(ET.parse(f).getroot())

        # Parse styles.xml (optional)
        if "word/styles.xml" in namelist:
            with zf.open("word/styles.xml") as f:
                self._styles_xml = ET.parse(f).getroot()
            self._build_style_id_map()
            self._parse_doc_defaults()

        # Parse settings.xml (optional)
        if "word/settings.xml" in namelist:
            with zf.open("word/settings.xml") as f:
                self._settings_xml = ET.parse(f).getroot()

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
            if style_id:
                # Cache element for O(1) lookups by styleId
                self._style_elem_cache[style_id] = style_elem
            name_elem = style_elem.find(f"{W_NS}name")
            if name_elem is not None and style_id:
                raw_name = name_elem.get(f"{W_NS}val", style_id)
                # Resolve built-in style names to canonical form
                is_custom = style_elem.get(f"{W_NS}customStyle", "") == "1"
                self._style_id_to_name[style_id] = (
                    raw_name if is_custom else _canonicalize_style_name(raw_name)
                )

    def _resolve_style_name(self, style_id: str) -> str:
        """Resolve a style ID to its display name."""
        return self._style_id_to_name.get(style_id, style_id)

    def _parse_theme(self, theme_root: ET.Element) -> None:
        """Parse theme XML to resolve theme fonts and colors."""
        theme_elements = theme_root.find(f"{A_NS}themeElements")
        if theme_elements is None:
            return

        font_scheme = theme_elements.find(f"{A_NS}fontScheme")
        if font_scheme is not None:
            for group_tag, prefix in [("majorFont", "major"), ("minorFont", "minor")]:
                group = font_scheme.find(f"{A_NS}{group_tag}")
                if group is None:
                    continue
                latin = group.find(f"{A_NS}latin")
                if latin is not None:
                    typeface = latin.get("typeface", "")
                    if typeface:
                        self._theme_fonts[f"{prefix}HAnsi"] = typeface
                        self._theme_fonts[f"{prefix}Ascii"] = typeface
                        self._theme_fonts[f"{prefix}Bidi"] = typeface

        clr_scheme = theme_elements.find(f"{A_NS}clrScheme")
        if clr_scheme is not None:
            # Every clrScheme child is one named color slot (dk1, lt1,
            # accent1…).  The slot's single child is either <a:srgbClr>
            # (literal RGB) or <a:sysClr> (with a ``lastClr`` attribute
            # holding the resolved value from the last save).  Either
            # way we end up with an "RRGGBB" hex string.
            for slot in clr_scheme:
                tag = slot.tag.split("}", 1)[-1]  # strip namespace
                if not tag:
                    continue
                srgb = slot.find(f"{A_NS}srgbClr")
                sys_clr = slot.find(f"{A_NS}sysClr")
                hex_val = ""
                if srgb is not None:
                    hex_val = (srgb.get("val") or "").upper()
                elif sys_clr is not None:
                    hex_val = (sys_clr.get("lastClr") or "").upper()
                if hex_val:
                    self._theme_colors[tag] = hex_val
            # OOXML aliases: ``bg1``/``bg2`` map to ``lt1``/``lt2`` and
            # ``tx1``/``tx2`` map to ``dk1``/``dk2``.  Documents may use
            # either form in ``w:color w:themeColor="…"`` references.
            for alias, target in (
                ("bg1", "lt1"),
                ("bg2", "lt2"),
                ("tx1", "dk1"),
                ("tx2", "dk2"),
                ("background1", "lt1"),
                ("background2", "lt2"),
                ("text1", "dk1"),
                ("text2", "dk2"),
            ):
                if target in self._theme_colors and alias not in self._theme_colors:
                    self._theme_colors[alias] = self._theme_colors[target]

    def _resolve_theme_font(self, theme_name: str) -> str:
        """Resolve a theme font reference like 'minorHAnsi' to actual name."""
        return self._theme_fonts.get(theme_name, "")

    def _resolve_theme_color(self, scheme_name: str) -> str:
        """Resolve a theme color slot (accent1, dk1, bg1…) to 'RRGGBB'."""
        return self._theme_colors.get(scheme_name, "")

    @staticmethod
    def _apply_theme_color_modifiers(
        base_hex: str,
        *,
        tint: "str | None" = None,
        shade: "str | None" = None,
    ) -> str:
        """Apply ``w:themeTint`` / ``w:themeShade`` modifiers to a base RGB.

        Delegates to :func:`docx_reader.utils._apply_theme_color_modifiers`.
        """
        return _apply_theme_color_modifiers(base_hex, tint=tint, shade=shade)

    def _parse_doc_defaults(self) -> None:
        """Parse docDefaults from styles.xml for default font/paragraph props."""
        if self._styles_xml is None:
            return
        doc_defaults = self._styles_xml.find(f"{W_NS}docDefaults")
        if doc_defaults is None:
            return
        rPr_default = doc_defaults.find(f"{W_NS}rPrDefault")
        if rPr_default is not None:
            self._doc_default_rPr = rPr_default.find(f"{W_NS}rPr")
        pPr_default = doc_defaults.find(f"{W_NS}pPrDefault")
        if pPr_default is not None:
            self._doc_default_pPr = pPr_default.find(f"{W_NS}pPr")

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
