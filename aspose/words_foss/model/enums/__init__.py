"""Document-model enums mirroring Aspose.Words core enumerations.

Each class uses plain integer constants (not ``enum.Enum``) to match
the Aspose.Words API convention where enum values are bare ints
serialisable to/from JSON without adapter logic.

Enums are organised into domain-specific modules:

- ``paragraph`` — ParagraphAlignment, LineSpacingRule
- ``style`` — StyleType, NumberStyle
- ``font`` — Underline
- ``layout`` — HeightRule, SectionStart, Orientation
- ``border`` — LineStyle
- ``table`` — CellMerge, CellVerticalAlignment
"""

from aspose.words_foss.model.enums.paragraph import (  # noqa: F401
    ParagraphAlignment,
    LineSpacingRule,
)
from aspose.words_foss.model.enums.style import (  # noqa: F401
    StyleType,
    NumberStyle,
)
from aspose.words_foss.model.enums.font import Underline  # noqa: F401
from aspose.words_foss.model.enums.layout import (  # noqa: F401
    HeightRule,
    SectionStart,
    Orientation,
)
from aspose.words_foss.model.enums.border import LineStyle  # noqa: F401
from aspose.words_foss.model.enums.table import (  # noqa: F401
    CellMerge,
    CellVerticalAlignment,
)
