from aspose.words_foss.document import Document, SaveFormat, LoadFormat
from aspose.words_foss import saving
from aspose.words_foss.model import wrap_type, enums
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

__version__ = "0.1.0"
__all__ = [
    "Document",
    "SaveFormat",
    "LoadFormat",
    "saving",
]
