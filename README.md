# Control Bancario AIE — Streamlit

App para analizar resúmenes de cuenta corriente en PDF **del formato aportado por el usuario** (banco con CBU que comienza en 310), con control individual de saldos, resumen operativo, préstamos, descubierto y controles fiscales potenciales.

## Instalación local

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# Linux/macOS: source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## GitHub + Streamlit Community Cloud

1. Crear un repositorio en GitHub y subir **los archivos del ZIP a la raíz**, no el directorio contenedor.
2. En https://share.streamlit.io/ elegir **Create app**, seleccionar repositorio y rama (p. ej. `main`).
3. En **Main file path** escribir `app.py` y desplegar.
4. Seleccionar Python 3.12 en la configuración avanzada si está disponible. `runtime.txt` documenta la versión objetivo; la opción efectiva la controla Streamlit Cloud.
5. Cargar el PDF en la web. Evitar colocar resúmenes reales en el repositorio público.

## Hojas del Excel

- Resumen operativo — importes por categoría sin duplicar débitos.
- Movimientos — original, referencia, débitos/créditos, saldo, reconstrucción y diferencia.
- Por mes — acumulados mensuales.
- Créditos fiscales — movimientos potencialmente computables, sujetos a revisión.
- Prestamos — préstamos y gastos por descubierto, separados.
- Controles — conciliación con totales del PDF.
- Auditoria — líneas no extraídas, movimientos inconsistentes y sin clasificar.
- Criterios — alcance, limitaciones y forma de cálculo.

## Alcance y advertencias

- No es un lector universal de resúmenes de cualquier banco. El patrón de columnas y fecha se validó con el PDF de Montegrosso julio-septiembre 2026.
- Una operación se clasifica como débito o crédito **por la variación efectiva del saldo**, no por la posición visual de los importes. Si no coincide, se muestra una diferencia y NO se asigna arbitrariamente.
- Los conceptos fiscales se exponen como *a verificar*. Los porcentajes del impuesto Ley 25.413 configurados por el usuario son una **estimación**, no determinan de modo automático beneficios fiscales, ni la contabilización de IVA.
- Los desembolsos de préstamos **no son ingresos operativos**, y las cobranzas de cuotas **no son íntegramente gastos**. Para el desglose de capital/intereses, tasa, vencimientos y deuda remanente se necesita el contrato o cuadro de amortización.
- Los saldos deudores son descubiertos en la cuenta; no miden el límite autorizado ni el crédito disponible.
- Sin OCR: los PDFs escaneados sin texto son rechazados explícitamente.
- Si se adjuntan PDFs que comparten movimientos, se bloquea la exportación para evitar duplicaciones.
- Los datos se procesan durante la sesión; no hay base de datos ni persistencia programada.

**Uso interno AIE · Developer Alfonso Alderete**
