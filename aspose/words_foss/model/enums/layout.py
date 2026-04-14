"""Page and section layout enums mirroring Aspose.Words."""


class HeightRule:
    """Specifies the rule for determining the height of an object.

    Mirrors ``Aspose.Words.HeightRule``.
    """

    AT_LEAST = 0
    """The height will be at least the specified height in points."""

    EXACTLY = 1
    """The height is specified exactly in points."""

    AUTO = 2
    """The height will grow automatically to accommodate all text."""


class SectionStart:
    """Specifies the type of break at the beginning of the section.

    Mirrors ``Aspose.Words.SectionStart``.
    """

    CONTINUOUS = 0
    NEW_COLUMN = 1
    NEW_PAGE = 2
    EVEN_PAGE = 3
    ODD_PAGE = 4


class Orientation:
    """Specifies page orientation.

    Mirrors ``Aspose.Words.Orientation``.
    """

    PORTRAIT = 0
    LANDSCAPE = 1
