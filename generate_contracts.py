import json
import os
import re
from typing import List, Dict

import gspread
from jinja2 import Environment, FileSystemLoader
import pdfkit

TEMPLATE_DIR = 'templates'
TEMPLATE_NAME = 'contract_template.html'
OUTPUT_DIR = 'PDFs'
SERVICE_ACCOUNT_FILE = 'service_account.json'  # ruta del archivo de la cuenta de servicio
SPREADSHEET_NAME = 'CENCO_DOCUMENTACION_PERSONAS'  # nombre del archivo de Google Sheets
SHEET_NAME = 'Personas'  # nombre de la hoja dentro del archivo

# Lista de columnas que quieres usar (deben coincidir exactamente con los nombres en Google Sheets)
COLUMNS_TO_USE = [
    'Nombre',
    'Nombre completo',
    'Rut',
    'Cargo',
    'Sueldo Liquido',
    'Fono',
    'Direccion',
    'Comuna',
    'Ciudad',
    'Tipo de contrato'  # <-- Añadida para el filtro
]

WKHTMLTOPDF_PATH = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'


def limpiar_sueldo(valor):
    """
    Limpia el sueldo: quita el símbolo $ y reemplaza la coma por punto.
    Ejemplo: "$750,000" -> "750.000"
    """
    if not isinstance(valor, str):
        valor = str(valor)
    valor = valor.replace("$", "").replace(",", ".").replace(" ", "")
    return valor


def sanitize_filename(filename):
    """
    Elimina caracteres no válidos para nombres de archivo.
    """
    filename = re.sub(r'[\\/*?:"<>|]', "_", filename)
    filename = filename.replace(' ', '_')
    return filename


def load_data() -> List[Dict]:
    """
    Lee los datos desde Google Sheets, filtra solo los empleados con 'Contrato por obra',
    mapea las columnas a las variables del HTML y agrega las variables opcionales como listas vacías.
    """
    gc = gspread.service_account(filename=SERVICE_ACCOUNT_FILE)
    files = gc.list_spreadsheet_files()
    print("Archivos accesibles por la cuenta de servicio:")
    for f in files:
        print(f["name"])
    # Luego sigue con:
    sh = gc.open(SPREADSHEET_NAME).worksheet(SHEET_NAME)
    records = sh.get_all_records()
    empleados = []
    for record in records:
        # Solo procesar si la columna 'Tipo de contrato' es 'Contrato por obra'
        if record.get('Tipo de contrato', '').strip().lower() == 'contrato por obra':
            empleado = {
                'nombre_completo': record.get('Nombre completo', ''),
                'rut': record.get('Rut', ''),
                'direccion': record.get('Direccion', ''),
                'comuna': record.get('Comuna', ''),
                'ciudad': record.get('Ciudad', ''),
                'nombre_del_cargo': record.get('Cargo', ''),
                'sueldo': limpiar_sueldo(record.get('Sueldo Liquido', '')),
                # Variables opcionales: listas vacías si no existen
                'sueldos_base': [],
                'bonos': [],
                'no_imponibles': []
            }
            # Guardamos también el nombre para el archivo
            empleado['nombre_archivo'] = sanitize_filename(record.get('Nombre', 'Desconocido')) + '_Contrato.html'
            empleados.append(empleado)
    return empleados


def render_contract(empleado: Dict, template):
    """
    Rellena la plantilla HTML con los datos del empleado.
    """
    return template.render(empleado=empleado)


def save_contract(html: str, filename: str):
    """
    Guarda el contrato HTML generado en la carpeta de salida.
    """
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    path = os.path.join(OUTPUT_DIR, filename)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(html)


def html_a_pdf(html_path, pdf_path):
    """
    Convierte un archivo HTML a PDF usando pdfkit y wkhtmltopdf.
    """
    config = pdfkit.configuration(wkhtmltopdf=WKHTMLTOPDF_PATH)
    pdfkit.from_file(html_path, pdf_path, configuration=config)


def main():
    """
    Función principal: carga los datos, genera y guarda solo los contratos en PDF.
    """
    env = Environment(loader=FileSystemLoader(TEMPLATE_DIR))
    template = env.get_template(TEMPLATE_NAME)
    records = load_data()
    for rec in records:
        html = render_contract(rec, template)
        pdf_path = os.path.join(OUTPUT_DIR, rec['nombre_archivo'].replace('.html', '.pdf'))
        # Generar PDF directamente desde el HTML en memoria
        config = pdfkit.configuration(wkhtmltopdf=WKHTMLTOPDF_PATH)
        pdfkit.from_string(html, pdf_path, configuration=config, options={
            'disable-smart-shrinking': '',
            'no-print-media-type': ''
        })
    print(f'Se generaron {len(records)} contratos en PDF en la carpeta {OUTPUT_DIR}')


if __name__ == '__main__':
    main()
