# Contratos

Este repositorio contiene plantillas de contratos HTML generados automáticamente.

## Generación de contratos

Se incluye un script `generate_contracts.py` que toma los datos de una hoja de cálculo de Google y rellena la plantilla `templates/contract_template.html` para crear un archivo HTML por cada trabajador.

### Requisitos

Instalar las dependencias de Python:

```bash
pip install -r requirements.txt
```

Debe existir un archivo `service_account.json` con las credenciales de una cuenta de servicio de Google que tenga acceso a la planilla. Ajuste en el script el nombre de la planilla si es necesario.

### Uso

```bash
python generate_contracts.py
```

Los contratos generados se guardarán en la carpeta `Contratos/` con el nombre de cada trabajador.
