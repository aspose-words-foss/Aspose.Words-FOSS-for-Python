from aspose.words_foss.document import (
    Document,
    SaveFormat,
    LoadFormat,
    LoadOptions,
    MarkdownLoadOptions,
)
from aspose.words_foss import loading, saving
from aspose.words_foss.model import wrap_type, enums
from aspose.words_foss.light_document_model import NodeType  # noqa: F401
from aspose.words_foss.model.enums import (  # noqa: F401 — re-export for aw.ParagraphAlignment etc.
    CellMerge,
    CellVerticalAlignment,
    HeightRule,
    LineSpacingRule,
    LineStyle,
    NumberStyle,
    Orientation,
    ParagraphAlignment,
    SectionStart,
    StyleType,
    Underline,
)

__version__ = "26.7.0"
__all__ = [
    "Document",
    "SaveFormat",
    "LoadFormat",
    "LoadOptions",
    "MarkdownLoadOptions",
    "NodeType",
    "loading",
    "saving",
]
