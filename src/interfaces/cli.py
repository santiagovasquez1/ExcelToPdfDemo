import sys

from src.application.excel_to_html import ExcelToHtmlConverter


def main():
    if len(sys.argv) < 3:
        print("Uso: python cli.py <ruta_excel> <ruta_salida_html>")
        sys.exit(1)

    excel_path = sys.argv[1]
    output_html_path = sys.argv[2]

    converter = ExcelToHtmlConverter(excel_path)
    result = converter.convert(output_html_path)
    print(f"Conversión completa. Archivo HTML generado: {result}")


if __name__ == "__main__":
    main()
