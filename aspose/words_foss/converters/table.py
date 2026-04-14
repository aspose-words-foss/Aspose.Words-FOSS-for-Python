"""Table conversion utilities."""

from typing import Optional

from aspose.words_foss.models import Table, TableRow, TableCell, ConversionOptions
from aspose.words_foss.docx_reader import TableData, CellData, ParagraphData, RunData


class TableConverter:
    """Handles conversion of tables to Markdown.

    This class works with TableData from the reader module,
    not with python-docx objects directly.
    """

    def __init__(self, options: ConversionOptions):
        self.options = options

    def convert(self, table: TableData) -> str:
        """Convert a table to Markdown format."""
        if not table.rows:
            return ""

        md_table = self._parse_table(table)
        return self._render(md_table)

    def _parse_table(self, table: TableData) -> Table:
        """Parse table data into internal Table structure."""
        md_table = Table()

        for row_idx, row in enumerate(table.rows):
            md_row = TableRow(is_header=(row_idx == 0))

            for cell in row.cells:
                cell_text = self._extract_cell_text(cell)
                alignment = cell.alignment
                md_row.cells.append(TableCell(text=cell_text, alignment=alignment))

            md_table.rows.append(md_row)

        return md_table

    def _extract_cell_text(self, cell: CellData) -> str:
        """Extract and format text from a table cell."""
        parts = []
        for para in cell.paragraphs:
            para_parts = []
            for run in para.runs:
                text = run.text or ""
                if text:
                    text = self._apply_inline_formatting(run, text)
                    para_parts.append(text)
            if para_parts:
                parts.append("".join(para_parts))

        result = " ".join(parts)
        result = result.replace("|", "\\|")
        result = result.replace("\n", " ").replace("\r", " ")
        return result.strip()

    def _apply_inline_formatting(self, run: RunData, text: str) -> str:
        """Apply inline formatting to run text."""
        result = text

        if run.bold and run.italic:
            result = f"***{result}***"
        elif run.bold:
            result = f"**{result}**"
        elif run.italic:
            result = f"*{result}*"

        if run.strikethrough and self.options.export_strikethrough:
            result = f"~~{result}~~"

        if run.is_code_style:
            result = f"`{result}`"

        return result

    def _render(self, table: Table) -> str:
        """Render Table structure as Markdown string."""
        if not table.rows:
            return ""

        num_cols = max(len(row.cells) for row in table.rows)
        col_widths = self._calculate_column_widths(table, num_cols)
        alignments = self._get_column_alignments(table, num_cols)

        lines = []
        for row_idx, row in enumerate(table.rows):
            line = self._render_row(row, num_cols, col_widths)
            lines.append(line)

            if row_idx == 0:
                separator = self._render_separator(num_cols, col_widths, alignments)
                lines.append(separator)

        return "\n".join(lines)

    def _calculate_column_widths(self, table: Table, num_cols: int) -> list[int]:
        """Calculate minimum column widths based on content."""
        widths = [3] * num_cols
        for row in table.rows:
            for col_idx, cell in enumerate(row.cells):
                if col_idx < num_cols:
                    widths[col_idx] = max(widths[col_idx], len(cell.text))
        return widths

    def _get_column_alignments(self, table: Table, num_cols: int) -> list[str]:
        """Get alignment for each column (from header row)."""
        alignments = ["left"] * num_cols
        if table.rows:
            header_row = table.rows[0]
            for col_idx, cell in enumerate(header_row.cells):
                if col_idx < num_cols:
                    alignments[col_idx] = cell.alignment
        return alignments

    def _render_row(self, row: TableRow, num_cols: int, col_widths: list[int]) -> str:
        """Render a single table row."""
        cells = []
        for col_idx in range(num_cols):
            if col_idx < len(row.cells):
                text = row.cells[col_idx].text
            else:
                text = ""
            cells.append(text.ljust(col_widths[col_idx]))
        return "| " + " | ".join(cells) + " |"

    def _render_separator(self, num_cols: int, col_widths: list[int], alignments: list[str]) -> str:
        """Render the header separator row."""
        separators = []
        for col_idx in range(num_cols):
            width = col_widths[col_idx]
            align = alignments[col_idx]

            if align == "center":
                sep = ":" + "-" * max(1, width - 2) + ":"
            elif align == "right":
                sep = "-" * max(1, width - 1) + ":"
            else:
                sep = "-" * width

            separators.append(sep)
        return "| " + " | ".join(separators) + " |"
