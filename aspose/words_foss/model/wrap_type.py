"""Drawing-related enums mirroring ``Aspose.Words.Drawing``."""


class WrapType:
    """Specifies how text is wrapped around a shape or picture.

    Mirrors ``Aspose.Words.Drawing.WrapType``.
    """

    INLINE = 0
    """The shape remains on the same layer as text and treated as a character."""

    TOP_BOTTOM = 1
    """The text stops at the top of the shape and restarts on the line below the shape."""

    SQUARE = 2
    """Wraps text around all sides of the square bounding box of the shape."""

    NONE = 3
    """No text wrapping around the shape. The shape is placed behind or in front of text."""

    TIGHT = 4
    """Wraps tightly around the edges of the shape, instead of wrapping around the bounding box."""

    THROUGH = 5
    """Same as Tight, but wraps inside any parts of the shape that are open."""
