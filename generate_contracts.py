import json
import os
from typing import List, Dict

import gspread
from jinja2 import Environment, FileSystemLoader

TEMPLATE_DIR = 'templates'
TEMPLATE_NAME = 'contract_template.html'
OUTPUT_DIR = 'Contratos'
SERVICE_ACCOUNT_FILE = 'service_account.json'  # ruta del archivo de la cuenta de servicio
SPREADSHEET_NAME = 'Contratos'  # nombre de la hoja de cálculo a abrir


def load_data() -> List[Dict]:
    gc = gspread.service_account(filename=SERVICE_ACCOUNT_FILE)
    sh = gc.open(SPREADSHEET_NAME).sheet1
    records = sh.get_all_records()
    # convertir columnas en listas si vienen en formato JSON
    for record in records:
        for key in ['sueldos_base', 'bonos', 'no_imponibles']:
            value = record.get(key)
            if isinstance(value, str) and value:
                try:
                    record[key] = json.loads(value)
                except json.JSONDecodeError:
                    record[key] = []
            else:
                record[key] = []
    return records


def render_contract(empleado: Dict, template):
    return template.render(empleado=empleado)


def save_contract(html: str, filename: str):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    path = os.path.join(OUTPUT_DIR, filename)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(html)


def main():
    env = Environment(loader=FileSystemLoader(TEMPLATE_DIR))
    template = env.get_template(TEMPLATE_NAME)
    records = load_data()
    for rec in records:
        nombre = rec.get('nombre_completo', 'contrato').replace(' ', '_')
        html = render_contract(rec, template)
        save_contract(html, f'{nombre}.html')
    print(f'Se generaron {len(records)} contratos en la carpeta {OUTPUT_DIR}')


if __name__ == '__main__':
    main()
