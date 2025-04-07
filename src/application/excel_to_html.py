from src.domain.converter_to_html import ConverterToHtml
from src.infraestructure.excel_reader import ExcelReader


class ExcelToHtmlConverter:
    def __init__(self, excel_path: str):
        self.excel_path = excel_path

    def convert(self, output_html_path: str) -> str:
        """
        Lee el archivo Excel, procesa cada hoja convirtiéndola a HTML y
        genera un archivo HTML con todas las hojas.
        """
        reader = ExcelReader(self.excel_path)
        sheets_data = reader.read_excel()

        # Iniciar el HTML global
        html = (
            "<html><head>"
            '<meta charset="UTF-8">'
            "<style>table, th, td { border: 1px solid black; border-collapse: collapse; padding: 5px; }</style>"
            "</head><body>\n"
        )

        # Procesar cada hoja
        for sheet in sheets_data:
            html += f"<h2>{sheet.sheet_name}</h2>\n"
            converter = ConverterToHtml(sheet)
            html += converter.to_html()
            html += "<br/>\n"

        html += "</body></html>"

        # Escribir el HTML a archivo
        with open(output_html_path, "w", encoding="utf-8") as f:
            f.write(html)

        return output_html_path
