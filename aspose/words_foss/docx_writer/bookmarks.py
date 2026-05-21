"""Document-wide bookmark id counter and name → id registry.

Word requires every ``<w:bookmarkStart>`` / ``<w:bookmarkEnd>`` pair to
share a unique numeric id within the document.  The LDM tracks both
markers via ``BookmarkStart`` / ``BookmarkEnd`` entries in
``inline_extras``, and Word allows a bookmark to span multiple
paragraphs — so the writer keeps a document-wide
``name → id`` registry to pair each end with the start that opened it,
no matter how many paragraphs apart they appear.
"""


class BookmarkState:
    """Hands out monotonically increasing bookmark ids and pairs
    starts/ends across paragraphs.

    The ``open`` mapping survives between paragraph renderings so a
    ``BookmarkStart`` in paragraph A and the matching ``BookmarkEnd``
    in paragraph B share the same numeric id.  After ``close()`` the
    entry leaves the map; subsequent ``open()`` calls for the same
    name (Word allows this) issue a fresh id.
    """

    def __init__(self) -> None:
        self._next = 0
        self._open: dict[str, int] = {}

    def next_id(self) -> int:
        cur = self._next
        self._next += 1
        return cur

    def open(self, name: str) -> int:
        bm_id = self.next_id()
        self._open[name] = bm_id
        return bm_id

    def close(self, name: str) -> int | None:
        """Return the id paired with ``name``'s open marker, or None
        when no start is currently open under that name."""
        return self._open.pop(name, None)

    def drain_open(self) -> list[tuple[str, int]]:
        """Return the currently-open ``(name, id)`` pairs and clear them.

        Used at the end of the document body to synthesise closing
        markers for bookmarks the source never closed — Word treats
        unclosed bookmarks as a corruption error."""
        items = list(self._open.items())
        self._open.clear()
        return items
