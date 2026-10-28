# CORPOSALUD-TOOLS

Utilidades para consolidar archivos de Excel, TXT y documentos.

## Combinador de Excel (maqueta)

El notebook `notebooks/01_maqueta_combinador.ipynb` es el banco de pruebas y la
guía de uso del combinador. La lógica reutilizable vive en
`excel_merge_core.py`: usa la **fila 1 del archivo de referencia** como
encabezado del consolidado, permite **elegir la hoja de origen de cada archivo**
después de escanear sus pestañas y permite mapear a mano las columnas de los
demás archivos.

### Puesta en marcha

El entorno ya está creado en `.venv/`. Para recrearlo desde cero:

```bash
uv venv .venv
uv pip install -r requirements.txt
```

Para generar los Excel de ejemplo:

```bash
uv run python scripts/generar_datos_demo.py
```

Para abrir el notebook:

```bash
uv run jupyter lab notebooks/01_maqueta_combinador.ipynb
```

Los archivos que quieras combinar se colocan en `data/entrada/` (o se indican
en `ARCHIVOS_MANUALES` dentro del notebook). El consolidado se guarda en
`data/salida/`.

Para correr las pruebas de la lógica:

```bash
uv run python -m unittest discover -s tests -v
```

### Estructura

```
excel_merge_core.py                     # lógica reutilizable, sin interfaz
notebooks/01_maqueta_combinador.ipynb   # guía paso a paso y banco de pruebas
scripts/generar_datos_demo.py           # datos falsos para probar
tests/test_excel_merge_core.py          # pruebas automáticas
data/entrada/                           # Excel de entrada (no versionado)
data/salida/                            # consolidados generados (no versionado)
requirements.txt
```

### Selección de hoja por archivo

`escanear_hojas` lista las pestañas de cada libro y `hojas_por_archivo` permite
elegir una distinta a la primera:

```python
from excel_merge_core import combinar_excels, escanear_hojas

rutas = ["data/entrada/marzo.xlsx", "data/entrada/abril.xlsx"]
print(escanear_hojas(rutas))
# {"marzo.xlsx": ["DATOS", "HOJA_ALTERNATIVA"], "abril.xlsx": ["DATOS"]}

df = combinar_excels(
    rutas,
    archivo_referencia="marzo.xlsx",
    hojas_por_archivo={"marzo.xlsx": "HOJA_ALTERNATIVA"},
)
```

### Aviso sobre datos sensibles

`data/` está en `.gitignore`. Si ejecutás el notebook con datos reales, **limpiá
las salidas antes de hacer commit** (`Kernel > Restart & Clear Output` en
JupyterLab) para no dejar información de pacientes en el historial del notebook.

### Siguiente fase

La lógica ya está extraída y probada. Falta armar la GUI Tkinter con
desplegables para elegir el archivo de referencia y la hoja de cada archivo,
reutilizando el mapeo manual y el aviso de columnas extra. En Linux, Tkinter
requiere el paquete del sistema `python3-tk`.
