"""DOCX writer — converts an LDM Document to a .docx file.

Public entry point is :class:`LdmDocxWriter`.  Re-exported here so callers
can simply do::

    from aspose.words_foss.docx_writer import LdmDocxWriter

The module is split into single-responsibility submodules: namespace
constants, XML helpers, OPC packaging, and per-element renderers
(paragraphs, runs, tables, styles, numbering).
"""

from aspose.words_foss.docx_writer.writer import (
    DocxWriterLossyWarning,
    LdmDocxWriter,
)

__all__ = ["DocxWriterLossyWarning", "LdmDocxWriter"]
