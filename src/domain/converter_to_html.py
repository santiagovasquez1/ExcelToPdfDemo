from src.domain.models.models import CellData, MergedRange, SheetData


class ConverterToHtml:
    def __init__(self, sheet_data: SheetData):
        self.sheet_data = sheet_data

    def _get_merged_range(self, cell: CellData) -> MergedRange | None:
        """
        Devuelve el rango fusionado que incluye la celda, si existe.
        """
        for merge in self.sheet_data.merged_ranges:
            if (
                merge.min_row <= cell.row <= merge.max_row
                and merge.min_col <= cell.col <= merge.max_col
            ):
                return merge
        return None

    def to_html(self) -> str:
        """
        Convierte los datos de la hoja (contenido en SheetData) en una tabla HTML,
        manejando correctamente las celdas fusionadas.
        """
        html = '<table border="1" style="border-collapse: collapse; padding: 5px;">\n'
        visited = (
            set()
        )  # Para no procesar varias veces las celdas de un rango fusionado

        for row in self.sheet_data.table_data:
            html += "  <tr>\n"
            for cell in row:
                cell_id = (cell.row, cell.col)
                if cell_id in visited:
                    continue

                merge = self._get_merged_range(cell)
                if merge:
                    rowspan = merge.max_row - merge.min_row + 1
                    colspan = merge.max_col - merge.min_col + 1

                    # Marcar todas las celdas del rango fusionado como procesadas
                    for r in range(merge.min_row, merge.max_row + 1):
                        for c in range(merge.min_col, merge.max_col + 1):
                            visited.add((r, c))
                    html += f'    <td rowspan="{rowspan}" colspan="{colspan}">{cell.value or ""}</td>\n'
                else:
                    html += f'    <td>{cell.value or ""}</td>\n'
            html += "  </tr>\n"
        html += "</table>\n"
        return html
