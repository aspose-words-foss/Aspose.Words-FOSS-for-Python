"""List handling for :class:`MarkdownReaderContext`.

Adapts the reference's ``ApplyList``/``SetList``/``RemoveList``/
``FetchBulletList``/``FetchOrderedList`` to the LDM: list membership lives
entirely on ``Paragraph.list_format`` (``list_id`` + ``list_level_number``),
resolved against ``Document.lists`` (``DocList``/``ListLevel``).
"""

from __future__ import annotations

from typing import Optional

from aspose.words_foss import light_document_model as ldm
from aspose.words_foss.model.enums import NumberStyle, ParagraphAlignment
from aspose.words_foss.model.list_limits import MAX_LIST_LEVELS

from ._reader_constants import _CHAR_TO_MARKER, _MARKER_CHAR
from .blocks import Block, BlockType, ListItemBlock, ListMarker


class _ListMixin:
    """List-format application and ``DocList``/``ListLevel`` bookkeeping."""

    def _apply_list(self, block: Block) -> None:
        item = (
            block
            if isinstance(block, ListItemBlock)
            else block.get_parent(BlockType.BULLET_LIST_ITEM, BlockType.ORDERED_LIST_ITEM)
        )
        if item is not None:
            self._set_list(block, item)
        else:
            self._remove_list()

    def _set_list(self, block: Block, item: ListItemBlock) -> None:
        b = block
        if block.type == BlockType.HTML_TAG and block.is_para_or_heading:
            parent_para = block.get_parent(BlockType.PARAGRAPH)
            b = parent_para if parent_para is not None else block

        if b is item:
            marker = item.marker
        elif b.parent is item and b.is_first_child:
            marker = item.marker
        else:
            marker = ListMarker.NONE

        if marker in (ListMarker.DOT, ListMarker.PARENTHESIS):
            doclist = self._fetch_ordered_list(item)
        else:
            doclist = self._fetch_bullet_list(marker)

        level = min(item.get_level(), MAX_LIST_LEVELS - 1)
        if marker != ListMarker.NONE:
            self._level_lists[level] = doclist

        lf = ldm.ListFormat()
        lf.is_list_item = True
        lf.list_id = doclist.list_id
        lf.list_level_number = level
        para = self._builder.current_paragraph
        para.list_format = lf
        para.paragraph_format.is_list_item = True

    def _remove_list(self) -> None:
        para = self._builder.current_paragraph
        if para.list_format is not None and para.list_format.list_id:
            self._list_containers.pop(para.list_format.list_id, None)
        para.list_format = None
        para.paragraph_format.is_list_item = False

    def _next_list_id(self) -> int:
        self._list_id_counter += 1
        return self._list_id_counter

    def _fetch_bullet_list(self, marker: ListMarker) -> ldm.DocList:
        doclist = self._bullet_lists.get(marker.value)
        if doclist is None:
            doclist = self._create_bullet_list(marker)
            self._bullet_lists[marker.value] = doclist
        return doclist

    def _create_bullet_list(self, marker: ListMarker) -> ldm.DocList:
        doclist = ldm.DocList()
        doclist.list_id = self._next_list_id()
        doclist.is_multi_level = False
        char = _MARKER_CHAR.get(marker, "-")
        doclist.list_levels = [
            ldm.ListLevel(
                number_style=NumberStyle.BULLET,
                number_format=char,
                start_at=1,
                alignment=ParagraphAlignment.LEFT,
                number_position=18.0 * (i + 1),
                text_position=18.0 * (i + 1) + 18.0,
            )
            for i in range(MAX_LIST_LEVELS)
        ]
        self._builder.document.lists.append(doclist)
        return doclist

    def _fetch_ordered_list(self, item: ListItemBlock) -> ldm.DocList:
        doclist = self._get_ordered_list(item)
        if doclist is None:
            doclist = ldm.DocList()
            doclist.list_id = self._next_list_id()
            doclist.is_multi_level = True
            self._list_containers[doclist.list_id] = item.get_list_container()
            self._builder.document.lists.append(doclist)

        level = min(item.get_level(), MAX_LIST_LEVELS - 1)
        while len(doclist.list_levels) <= level:
            doclist.list_levels.append(ldm.ListLevel())
        if not doclist.list_levels[level].number_format:
            self._init_ordered_list_level(doclist, level, item)
        return doclist

    def _get_ordered_list(self, item: ListItemBlock) -> Optional[ldm.DocList]:
        level = min(item.get_level(), MAX_LIST_LEVELS - 1)
        doclist = self._level_lists[level]
        if doclist is None and level > 0:
            doclist = self._level_lists[level - 1]
        if not self._is_suitable_list(doclist, item):
            return None
        return doclist

    @staticmethod
    def _init_ordered_list_level(doclist: ldm.DocList, level: int, item: ListItemBlock) -> None:
        ll = doclist.list_levels[level]
        ll.number_style = NumberStyle.ARABIC
        ll.start_at = item.start_at
        ll.alignment = ParagraphAlignment.LEFT
        marker_char = _MARKER_CHAR.get(item.marker, ".")
        ll.number_format = f"%{level + 1}{marker_char}"
        indent = 36.0 * (level + 1)
        ll.number_position = indent
        ll.text_position = indent + 18.0

    def _is_suitable_list(self, doclist: Optional[ldm.DocList], item: ListItemBlock) -> bool:
        if doclist is None:
            return False
        if item.type == BlockType.ORDERED_LIST_ITEM:
            if self._list_containers.get(doclist.list_id) is not item.get_list_container():
                return False
        level = (
            min(item.get_level(), MAX_LIST_LEVELS - 1)
            if item.type == BlockType.ORDERED_LIST_ITEM
            else 0
        )
        if level >= len(doclist.list_levels):
            return True
        ll = doclist.list_levels[level]
        if not ll.number_format:
            return True
        marker = _CHAR_TO_MARKER.get(ll.number_format[-1])
        if item.marker != marker:
            return False
        if item.type == BlockType.ORDERED_LIST_ITEM and item.is_level_start and ll.start_at != item.start_at:
            return False
        return True

    # -- Type-only declarations of host-provided state (see reader_context.py) --
    _bullet_lists: dict
    _list_containers: dict
    _level_lists: list
    _list_id_counter: int
    _builder: object
