"""Table-related enums."""


class CellMerge:
    """Specifies how a cell in a table is merged with other cells.
    """

    NONE = 0
    FIRST = 1
    PREVIOUS = 2


class CellVerticalAlignment:
    """Specifies vertical justification of text inside a table cell.
    """

    TOP = 0
    CENTER = 1
    BOTTOM = 2
