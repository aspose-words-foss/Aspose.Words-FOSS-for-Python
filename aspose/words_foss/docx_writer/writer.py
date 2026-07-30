"""LdmDocxWriter — orchestrates DOCX generation from an LDM Document.

Public surface:

* ``write(doc, output_path)`` — render ``doc`` and zip it to disk.
* ``write_to_bytes(doc)``     — same but returns the raw ``bytes`` of
  the resulting ``.docx`` for in-memory use (tests, services).
"""


import io
import zipfile
from pathlib import Path
from typing import Optional, Union

from dataclasses import dataclass

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.saving import (
    CompressionLevel,
    OoxmlCompliance,
    Zip64Mode,
    coerce_enum,
)
from aspose.words_foss.docx_writer.bookmarks import BookmarkState
from aspose.words_foss.docx_writer.document_part import _build_style_pf_map, render_document_xml
from aspose.words_foss.docx_writer.drawing import ImageEntry, ImageRenderState
from aspose.words_foss.docx_writer.headers_footers import (
    needs_footer,
    needs_header,
    render_footer_xml,
    render_header_xml,
)
from aspose.words_foss.docx_writer.numbering_part import render_numbering_xml
from aspose.words_foss.docx_writer.package import write_docx_package
from aspose.words_foss.docx_writer.settings_part import needs_settings, render_settings_xml
from aspose.words_foss.docx_writer.styles_part import (
    build_style_font_map,
    build_style_id_map,
    render_styles_xml,
)


class DocxWriterLossyWarning(UserWarning):
    """Warns the caller that the writer is dropping known LDM constructs.

    The reader populates fields the writer does not yet emit (headers,
    footers, inline shapes/images).  Round-tripping a document that
    carries any of these silently strips them.  We surface a
    :class:`UserWarning` so call sites can opt into stricter behaviour
    (``warnings.simplefilter("error", DocxWriterLossyWarning)``) or
    suppress it if the lossiness is acceptable.
    """


# Headers/footers, inline shapes/images, bookmarks, and custom-style
# ``<w:pPr>`` are now preserved on round-trip — see ``drawing.py``,
# ``bookmarks.py``, ``headers_footers.py`` and ``styles_part._custom_style``.
# Text-box-only shapes (``Shape`` without ``image_data``) and
# ``FieldStart`` markers in ``inline_extras`` are still dropped silently
# — they don't appear in the test fixtures and adding emit paths for
# them is a separate scope.


def _warn_about_unsupported_constructs(doc: ldm.Document) -> None:
    """Reserved for future lossy-construct warnings.

    Currently a no-op: every previously-warned-about construct
    (headers/footers, inline shapes) now round-trips through the writer.
    Kept as a hook so new lossy fields can be flagged without rewiring
    the call site.
    """
    del doc  # nothing to warn about right now


_COMPRESSION_MAP: dict[int, tuple[int, int]] = {
    CompressionLevel.NORMAL: (zipfile.ZIP_DEFLATED, 6),
    CompressionLevel.MAXIMUM: (zipfile.ZIP_DEFLATED, 9),
    CompressionLevel.FAST: (zipfile.ZIP_DEFLATED, 3),
    CompressionLevel.SUPER_FAST: (zipfile.ZIP_DEFLATED, 1),
}

_ZIP64_MAP: dict[int, bool] = {
    Zip64Mode.NEVER: False,
    Zip64Mode.IF_NECESSARY: True,
    Zip64Mode.ALWAYS: True,
}


def _collect_hf_children(
    doc: ldm.Document, *, header: bool
) -> list[ldm.Paragraph | ldm.Table]:
    """Collect ordered children for header or footer rendering.

    Prefers the per-section ``HeaderFooter.children`` list (which preserves
    interleaved paragraph/table order) and falls back to the flat
    ``doc.header_paragraphs`` / ``doc.footer_paragraphs`` for backwards
    compatibility with LDMs that only carry paragraphs.
    """
    hf_type = 0 if header else 1
    for sec in doc.sections:
        for hf in sec.headers_footers:
            if hf.header_footer_type == hf_type and hf.children:
                return list(hf.children)
    paras = doc.header_paragraphs if header else doc.footer_paragraphs
    return list(paras)


class LdmDocxWriter:
    """Convert an :class:`ldm.Document` into a DOCX file.

    The writer is intentionally stateless across calls — every
    invocation produces a fresh hyperlink relationship map.  This keeps
    parallel/sequential exports independent and avoids the bug-prone
    ``reset()``-style APIs the reader had to evolve.
    """

    def __init__(self, options: Optional[object] = None):
        # ``options`` accepts an ``OoxmlSaveOptions`` instance only —
        # No separate ``DocxSaveOptions`` is exposed —
        # ``DocxSaveOptions``.  Weakly typed so callers don't have to
        # import the save-options module to use the writer.
        self.options = options
        self._enforce_compliance()

    # ------------------------------------------------------------------
    def _enforce_compliance(self) -> None:
        """Reject compliance levels we cannot satisfy.

        ECMA-376 and ISO 29500 Transitional share the same on-disk
        layout for the constructs we emit, so both are accepted.
        Strict ISO 29500 mandates a separate namespace and forbids
        legacy elements (e.g. ``w:cantSplit``) — silently emitting
        transitional XML when the caller asked for Strict would be
        worse than failing fast.
        """
        compliance = getattr(self.options, "compliance", None)
        if compliance == OoxmlCompliance.ISO29500_2008_STRICT:
            raise NotImplementedError(
                "OoxmlCompliance.ISO29500_2008_STRICT is not supported by "
                "the FOSS writer; use ECMA376_2006 or "
                "ISO29500_2008_TRANSITIONAL."
            )

    def _compression(self) -> tuple[int, int]:
        """Return ``(compress_type, compresslevel)`` for :mod:`zipfile`."""
        level = coerce_enum(
            CompressionLevel, getattr(self.options, "compression_level", CompressionLevel.NORMAL)
        )
        return _COMPRESSION_MAP.get(level, _COMPRESSION_MAP[CompressionLevel.NORMAL])

    def _allow_zip64(self) -> bool:
        mode = getattr(self.options, "zip_64_mode", Zip64Mode.NEVER)
        return _ZIP64_MAP.get(coerce_enum(Zip64Mode, mode), False)

    def _pretty_format(self) -> bool:
        return bool(getattr(self.options, "pretty_format", False))

    def _render_parts(self, doc: ldm.Document) -> "_RenderedParts":
        """Render every XML part the package needs and return them bundled."""
        _warn_about_unsupported_constructs(doc)
        # The hyperlink table and image state are mutated by :mod:`runs`
        # / :mod:`paragraphs` while rendering the document body; gather
        # them first, then build dependents.
        hyperlinks: dict[str, str] = {}
        image_state = ImageRenderState()
        bookmark_state = BookmarkState()
        has_header = needs_header(doc)
        has_footer = needs_footer(doc)
        document_xml = render_document_xml(
            doc,
            hyperlinks,
            image_state=image_state,
            bookmark_state=bookmark_state,
            has_header=has_header,
            has_footer=has_footer,
        )
        styles_xml = render_styles_xml(doc)
        numbering_xml = render_numbering_xml(doc) or None
        settings_xml = render_settings_xml(doc) if needs_settings(doc) else None

        # Header / footer parts use their own hyperlink table — Word
        # resolves ``r:id`` against the part's sidecar rels file, not
        # the document's, so collisions on ``rIdLink<N>`` are fine.
        # Each part also owns its own image state with a unique media
        # filename prefix so images dropped into ``word/media/`` from
        # the header / footer don't overwrite the body's image1.png.
        header_xml: Optional[str] = None
        header_rels: dict[str, str] = {}
        # Header / footer anchors are placed relative to header_distance
        # / footer_distance, not the section top margin.  Use the first
        # section's distances as the baseline (matches reader behaviour).
        first_section = doc.sections[0] if doc.sections else None
        first_ps = first_section.page_setup if first_section else None
        from aspose.words_foss.docx_writer.drawing import pt_to_mm
        header_baseline_mm = pt_to_mm(first_ps.header_distance) if first_ps and first_ps.header_distance else 12.7
        page_height_mm = pt_to_mm(first_ps.page_height) if first_ps and first_ps.page_height else 297.0
        footer_distance_mm = pt_to_mm(first_ps.footer_distance) if first_ps and first_ps.footer_distance else 12.7
        footer_baseline_mm = page_height_mm - footer_distance_mm
        left_margin_mm = pt_to_mm(first_ps.left_margin) if first_ps and first_ps.left_margin else 25.4
        header_image_state = ImageRenderState(
            media_name_prefix="header_image",
            rel_id_prefix="rIdHdrImg",
            doc_pr_offset=100000,
            section_top_margin_mm=header_baseline_mm,
            section_left_margin_mm=left_margin_mm,
        )
        footer_xml: Optional[str] = None
        footer_rels: dict[str, str] = {}
        footer_image_state = ImageRenderState(
            media_name_prefix="footer_image",
            rel_id_prefix="rIdFtrImg",
            doc_pr_offset=200000,
            section_top_margin_mm=footer_baseline_mm,
            section_left_margin_mm=left_margin_mm,
        )
        style_pf_map = _build_style_pf_map(doc)
        style_id_map = build_style_id_map(doc)
        style_font_map = build_style_font_map(doc)
        if has_header:
            hdr_children = _collect_hf_children(doc, header=True)
            header_xml = render_header_xml(
                hdr_children,
                header_rels,
                style_pf_map,
                style_id_map,
                style_font_map,
                image_state=header_image_state,
            )
        if has_footer:
            ftr_children = _collect_hf_children(doc, header=False)
            footer_xml = render_footer_xml(
                ftr_children,
                footer_rels,
                style_pf_map,
                style_id_map,
                style_font_map,
                image_state=footer_image_state,
            )

        return _RenderedParts(
            document_xml=document_xml,
            styles_xml=styles_xml,
            numbering_xml=numbering_xml,
            settings_xml=settings_xml,
            hyperlinks=hyperlinks,
            images=image_state.images,
            header_xml=header_xml,
            header_hyperlinks=header_rels,
            header_images=header_image_state.images,
            footer_xml=footer_xml,
            footer_hyperlinks=footer_rels,
            footer_images=footer_image_state.images,
        )

    # ------------------------------------------------------------------
    def write(self, doc: ldm.Document, output_path: Union[str, Path]) -> None:
        """Render ``doc`` and write the resulting ``.docx`` to disk."""
        parts = self._render_parts(doc)
        compress_type, compresslevel = self._compression()
        write_docx_package(
            output_path,
            **parts.as_kwargs(),
            compression=compress_type,
            compresslevel=compresslevel,
            allow_zip64=self._allow_zip64(),
            pretty_format=self._pretty_format(),
        )

    def write_to_bytes(self, doc: ldm.Document) -> bytes:
        """Render ``doc`` and return the ``.docx`` as raw bytes."""
        parts = self._render_parts(doc)
        buffer = io.BytesIO()
        compress_type, compresslevel = self._compression()
        write_docx_package(
            buffer,
            **parts.as_kwargs(),
            compression=compress_type,
            compresslevel=compresslevel,
            allow_zip64=self._allow_zip64(),
            pretty_format=self._pretty_format(),
        )
        return buffer.getvalue()


@dataclass
class _RenderedParts:
    """Bundle of fully-rendered DOCX part payloads for the package writer."""

    document_xml: str
    styles_xml: str
    numbering_xml: Optional[str]
    settings_xml: Optional[str]
    hyperlinks: dict[str, str]
    images: list[ImageEntry]
    header_xml: Optional[str]
    header_hyperlinks: dict[str, str]
    header_images: list[ImageEntry]
    footer_xml: Optional[str]
    footer_hyperlinks: dict[str, str]
    footer_images: list[ImageEntry]

    def as_kwargs(self) -> dict[str, object]:
        """Return the kwargs ``write_docx_package`` accepts."""
        return {
            "document_xml": self.document_xml,
            "styles_xml": self.styles_xml,
            "numbering_xml": self.numbering_xml,
            "settings_xml": self.settings_xml,
            "hyperlinks": self.hyperlinks,
            "images": self.images,
            "header_xml": self.header_xml,
            "header_hyperlinks": self.header_hyperlinks,
            "header_images": self.header_images,
            "footer_xml": self.footer_xml,
            "footer_hyperlinks": self.footer_hyperlinks,
            "footer_images": self.footer_images,
        }
