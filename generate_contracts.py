import json
import os
import re
from typing import List, Dict
from datetime import datetime

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


def limpiar_nombre(nombre):
    """
    Limpia el nombre: elimina espacios dobles, finales y capitaliza cada palabra.
    Si el nombre está repetido (ej: 'MELA MELA'), lo deja solo una vez.
    """
    if not isinstance(nombre, str):
        return nombre
    nombre = ' '.join(nombre.strip().split())  # quita espacios dobles y finales
    partes = nombre.lower().split()
    # Elimina repeticiones consecutivas
    partes_sin_repetir = []
    for i, parte in enumerate(partes):
        if i == 0 or parte != partes[i-1]:
            partes_sin_repetir.append(parte)
    nombre_limpio = ' '.join([p.capitalize() for p in partes_sin_repetir])
    return nombre_limpio


def load_data() -> List[Dict]:
    """
    Lee los datos desde Google Sheets, limpia y corrige los datos,
    y devuelve una lista de empleados listos para usar en la plantilla.
    """
    gc = gspread.service_account(filename=SERVICE_ACCOUNT_FILE)
    sh = gc.open(SPREADSHEET_NAME).worksheet(SHEET_NAME)
    records = sh.get_all_records()
    empleados = []
    for record in records:
        if record.get('Tipo de contrato', '').strip().lower() == 'contrato por obra':
            empleado = {
                'nombre_completo': limpiar_nombre(record.get('Nombre completo', '')),
                'rut': record.get('Rut', ''),
                'direccion': limpiar_nombre(record.get('Direccion', '')),
                'comuna': limpiar_nombre(record.get('Comuna', '')),
                'ciudad': limpiar_nombre(record.get('Ciudad', '')),
                'nombre_del_cargo': limpiar_nombre(record.get('Cargo', '')),
                'sueldo': limpiar_sueldo(record.get('Sueldo Liquido', '')),
                'fecha_contrato': fecha_formateada(),
                'sueldos_base': [],
                'bonos': [],
                'no_imponibles': []
            }
            empleado['nombre_archivo'] = sanitize_filename(limpiar_nombre(record.get('Nombre', 'Desconocido'))) + '_Contrato.html'
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


def fecha_formateada():
    """
    Devuelve la fecha actual en formato 'Lunes 23 de junio del 2025'
    """
    meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio",
             "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
    dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
    hoy = datetime.now()
    nombre_dia = dias[hoy.weekday()]
    nombre_mes = meses[hoy.month - 1]
    return f"{nombre_dia} {hoy.day} de {nombre_mes} del {hoy.year}"


def main():
    """
    Función principal: carga los datos, genera y guarda solo los contratos en PDF.
    """
    env = Environment(loader=FileSystemLoader(TEMPLATE_DIR))
    template = env.get_template(TEMPLATE_NAME)
    records = load_data()
    for rec in records:
        try:
            html = render_contract(rec, template)
            pdf_path = os.path.join(OUTPUT_DIR, rec['nombre_archivo'].replace('.html', '.pdf'))
            config = pdfkit.configuration(wkhtmltopdf=WKHTMLTOPDF_PATH)
            pdfkit.from_string(
                html,
                pdf_path,
                configuration=config,
                options={
                    'disable-smart-shrinking': '',
                    'no-print-media-type': '',
                    'enable-local-file-access': ''  # <-- agrega esta línea
                }
            )
        except Exception as e:
            print(f"Error generando PDF para {rec['nombre_completo']}: {e}")
    print(f'Se generaron {len(records)} contratos en PDF en la carpeta {OUTPUT_DIR}')


if __name__ == '__main__':
    main()
