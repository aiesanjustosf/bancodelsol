# AIE San Justo - Resumen Operativo Bancario

App **Streamlit** para cargar un resumen bancario PDF del mismo formato que `MONTEGROSSO JUL-SET26.pdf` y obtener dos descargas independientes:

1. **PDF «Resumen Operativo: Registración Módulo IVA»**, tabla de Concepto / Importe siguiendo el formato usado por AIE en Galicia.
2. **Excel `.xlsx` real** con hojas `Resumen operativo`, `Movimientos`, `Gastos`, `Créditos bancarios`, `Control y conciliación` y `Para revisar`.

## Reglas contables

- Netos, IVA y brutos por alícuota 21%, 10,5% y 27%. La base neta es el movimiento de gasto identificado, no una división de importes sin sustento.
- Gastos exentos y no gravados **solo** cuando lo indica el concepto o el responsable los reclasifica; gastos sin IVA identificado se muestran aparte.
- IVA se vincula a una base cuando coinciden **fecha, referencia e importe para la alícuota**, sin usar fecha como único identificador. La alícuota inferida se marca como pendiente de respaldo documental.
- Percepciones de IVA, otras percepciones, Ley 25.413 Débito y Crédito, SIRCREB y retenciones independientes.
- El **TOTAL** operativo suma cada bruto y tributo **una sola vez**: no duplica netos e IVA ni subtotales de Ley 25.413.
- Créditos/préstamos se muestran aparte y **no se consideran gastos** salvo intereses cobrados identificados.
- `IVA potencial` no equivale a crédito fiscal legalmente computable, y el descubierto no permite conocer el límite crediticio contratado.
- Cada movimiento conserva fecha, descripción, débito, crédito y saldo; se controlan los saldos movimiento a movimiento.
- Los importes en Excel se exportan como valores numéricos compatibles con Excel.

## GitHub / Streamlit Community Cloud

1. Crear un nuevo repositorio GitHub y subir **el contenido** del ZIP, no el ZIP mismo.
2. En [share.streamlit.io](https://share.streamlit.io) elegir el repositorio y la rama `main`.
3. Archivo principal: `app.py`. Python 3.12 (`runtime.txt` incluido).
4. Streamlit instalará `requirements.txt` y ejecutará la app.

## Archivos principales

- `app.py`: interfaz, descarga de PDF y Excel, control de créditos.
- `processor.py`: parser probado para el PDF Montegrosso y auditoría de saldo.
- `operativo.py`: cálculo de conceptos y total contable.
- `excel_export.py`: genera Excel `.xlsx` compatible mediante `openpyxl`.
- `pdf_export.py`: resumen PDF de una página con tabla simple.
- `tests/test_app.py`: validación sobre casos sintéticos + validación real con PDF opcional.

## Pruebas

En el repositorio, con las dependencias instaladas:

```bash
python -m unittest discover -s tests -v
```

La app **no** acepta silenciosamente resúmenes de otro banco o formato. Si el parser no reconoce las líneas, informa el error; habría que incorporar un lector específico para ese banco. No adjuntar PDFs con datos de clientes al repositorio público.

**Herramienta para uso interno AIE San Justo | Developer Alfonso Alderete**
