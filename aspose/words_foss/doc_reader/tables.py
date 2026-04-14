"""
Table text parser for DOC files.
"""

from __future__ import annotations

from aspose.words_foss.docx_reader import (
    CellData,
    ParagraphData,
    RowData,
    RunData,
    TableData,
)


def parse_table_text(text: str) -> TableData:
    """Parse Word table text (using \\x07 cell separators) into TableData."""
    data = TableData()
    rows_text = text.split("\x07\r")

    for row_text in rows_text:
        row_text = row_text.strip("\x07").strip("\r")
        if not row_text:
            continue

        row = RowData()
        cell_texts = row_text.split("\x07")
        # Last element after split is often empty or the row-end marker
        for ct in cell_texts:
            cell = CellData()
            para = ParagraphData()
            para.text = ct
            run = RunData()
            run.text = ct
            para.runs = [run]
            cell.paragraphs = [para]
            row.cells.append(cell)
        data.rows.append(row)

    return data
