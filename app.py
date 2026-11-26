"""AIE San Justo: PDF banco -> Resumen Operativo (PDF) + Excel auditable."""
from copy import deepcopy
from html import escape
from pathlib import Path
import re
import streamlit as st
from PIL import Image

from processor import extract_pdf, apply_review, ALICUOTAS
from operativo import make_summary, operation_rows, money_ar
from excel_export import export_excel
from pdf_export import export_pdf

BASE_DIR = Path(__file__).resolve().parent
ASSETS = BASE_DIR / 'assets'
LOGO_PATH = ASSETS / 'aie_logo.png'
FAVICON_PATH = ASSETS / 'aie_favicon.png'
page_icon = Image.open(FAVICON_PATH) if FAVICON_PATH.exists() else '🏦'

st.set_page_config(
    page_title='AIE | Resumen Operativo Bancario',
    page_icon=page_icon,
    layout='centered',
    initial_sidebar_state='collapsed'
)

st.markdown('''
<style>
.block-container {max-width: 960px; padding-top: 1.25rem; padding-bottom: 2rem;}
.aie-hero {text-align:center; margin-bottom: 1rem;}
.aie-hero img {max-width: 170px; margin: 0 auto 0.6rem auto; display:block;}
.aie-title {font-size: 2.3rem; font-weight: 800; color:#16344E; line-height:1.12; margin: 0.2rem 0 0.35rem 0; text-align:center;}
.aie-subtitle {font-size: 1rem; color:#4C647A; text-align:center; margin-bottom: 1.15rem;}
div[data-testid="stFileUploader"] {background:#F5F8FB; border:1px solid #D9E2EC; border-radius:14px; padding:0.6rem 0.85rem;}
.stDownloadButton button, .stButton button {border-radius: 10px;}
[data-testid="stMetricValue"] {font-size: 1.35rem;}
h1,h2,h3 {color:#16344E}
.aie-tbl {border-collapse:collapse;width:100%;font-size:14px;margin-top:6px}
.aie-tbl th {background:#dedede;text-align:left;padding:9px 12px;border:1px solid #bdbdbd}
.aie-tbl td {padding:7px 12px;border:1px solid #c9c9c9}
.aie-tbl td:last-child {text-align:right;font-variant-numeric:tabular-nums}
.aie-tbl .subtotal td {background:#f4f6f7;font-weight:600}
.aie-tbl .total td {background:#e7eff4;font-weight:750;font-size:16px}
.aie-foot {margin-top:32px;padding-top:12px;border-top:1px solid #ddd;color:#777;font-size:13px;text-align:center}
</style>
''', unsafe_allow_html=True)

with st.container():
    if LOGO_PATH.exists():
        left, center, right = st.columns([1.4, 2.2, 1.4])
        with center:
            st.image(str(LOGO_PATH), width=150)
    st.markdown('<div class="aie-title">AIE | Resumen Operativo Bancario</div>', unsafe_allow_html=True)
    st.markdown('<div class="aie-subtitle">Procesamiento de resúmenes bancarios · Registración Módulo IVA</div>', unsafe_allow_html=True)

up = st.file_uploader(
    'Subir resumen bancario PDF',
    type=['pdf'],
    help='Lector probado con la estructura del resumen MONTEGROSSO JUL-SET26. Cada carga se procesa de manera independiente.'
)
if not up:
    st.info('Cargá el PDF para obtener el Resumen Operativo como el de Galicia y el Excel de movimientos y controles.')
    st.markdown('<div class="aie-foot">Herramienta para uso interno AIE San Justo | Developer Alfonso Alderete</div>', unsafe_allow_html=True)
    st.stop()

try:
    rep = extract_pdf(up.getvalue(), up.name)
except Exception as ex:
    st.error('No fue posible leer el resumen. ' + str(ex))
    st.stop()

periods = sorted(set(m.periodo for m in rep['movimientos']))
selected = st.selectbox('Período', ['Período completo'] + periods)
all_period = selected == 'Período completo'
movements = deepcopy([m for m in rep['movimientos'] if all_period or m.periodo == selected])
if not movements:
    st.error('No hay movimientos para el período elegido.')
    st.stop()

with st.expander('Ajustes manuales de clasificación (opcional)', expanded=False):
    st.caption('Corregí únicamente operaciones respaldadas. Sin IVA identificado no equivale a exento.')
    editables = [m for m in movements if m.debito > 0 and m.naturaleza in ('Gasto', 'Pendiente revisión')]
    if editables:
        entries = [{
            'ID': m.id,
            'Fecha': m.fecha,
            'Concepto': m.detalle,
            'Débito': float(m.debito),
            'Naturaleza': m.naturaleza,
            'Rubro de gasto': m.rubro or m.categoria,
            'Tratamiento IVA': m.alicuota
        } for m in editables]
        edited = st.data_editor(
            entries,
            hide_index=True,
            num_rows='fixed',
            use_container_width=True,
            disabled=['ID', 'Fecha', 'Concepto', 'Débito'],
            key='ajuste_operativo',
            column_config={
                'Naturaleza': st.column_config.SelectboxColumn('Naturaleza', options=['Gasto', 'Pendiente revisión', 'Movimiento no operativo']),
                'Tratamiento IVA': st.column_config.SelectboxColumn('Tratamiento IVA', options=ALICUOTAS),
                'Débito': st.column_config.NumberColumn('Débito', format='%.2f')
            }
        )
        for w in apply_review(movements, edited):
            st.warning(w)

s = make_summary(movements)
st.subheader('Resumen Operativo: Registración Módulo IVA')
html = '<table class="aie-tbl"><thead><tr><th>Concepto</th><th>Importe</th></tr></thead><tbody>'
for key, label, val, subtotal in operation_rows(s):
    typ = 'total' if key == 'total' else 'subtotal' if subtotal else ''
    html += f'<tr class="{typ}"><td>{escape(label)}</td><td>{money_ar(val)}</td></tr>'
html += '</tbody></table>'
st.markdown(html, unsafe_allow_html=True)
st.caption('Los importes Bruto y el subtotal Ley 25.413 son informativos: no se duplican en el TOTAL.')

period_title = (f'{periods[0]} a {periods[-1]}' if all_period and len(periods) > 1 else periods[0] if all_period else selected)
name = re.sub(r'[^a-zA-Z0-9_-]+', '_', rep.get('cuenta', 'cuenta')) + '_' + selected.replace('Período completo', 'completo')
try:
    pdf = export_pdf(s, period_title, rep.get('titular', ''), rep.get('cuenta', ''))
    excel = export_excel(rep, movements, period_title, selected_all=all_period)
except Exception as ex:
    st.error(f'No se pudieron crear las descargas: {ex}')
    st.stop()

l, r = st.columns(2)
with l:
    st.download_button(
        'Descargar resumen operativo PDF',
        data=pdf,
        file_name=f'AIE_Resumen_Operativo_{name}.pdf',
        mime='application/pdf',
        type='primary',
        use_container_width=True
    )
with r:
    st.download_button(
        'Descargar movimientos Excel',
        data=excel,
        file_name=f'AIE_Control_Bancario_{name}.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        use_container_width=True
    )

with st.expander('¿Hay créditos bancarios o préstamos?', expanded=True):
    if s['tiene_credito_bancario']:
        st.success('Sí: el extracto contiene movimientos de préstamos o saldos deudores.')
    else:
        st.info('No se identificaron préstamos ni descubierto en los movimientos seleccionados.')
    a, b = st.columns(2)
    a.metric('Desembolsos de préstamos', '$ ' + money_ar(s['prestamos_desembolsados']))
    b.metric('Cobranza de créditos / cuotas', '$ ' + money_ar(s['prestamos_cobrados']))
    st.caption('El extracto no permite determinar el saldo adeudado del préstamo ni el límite de financiación disponible.')

with st.expander('Controles de integridad e IVA'):
    first, last = movements[0], movements[-1]
    saldo_ini = first.saldo - first.credito + first.debito
    computed = saldo_ini + s['creditos'] - s['debitos']
    st.write(f'**Movimientos procesados:** {s["movimientos"]}')
    st.write(f'**Saldo inicial:** $ {money_ar(saldo_ini)}')
    st.write(f'**Saldo final calculado:** $ {money_ar(computed)}')
    st.write(f'**Diferencias de saldos:** {s["diferencias"]}')
    if all_period and rep.get('saldo_final_informado') is not None:
        st.write(f'**Diferencia contra saldo final del PDF:** $ {money_ar(last.saldo - rep["saldo_final_informado"])}')
    st.write(f'**IVA bancario potencial a revisar:** $ {money_ar(sum(s["iva"].values()) + s["iva_sin_vincular"])}')
    st.caption('IVA informado o vinculado no implica automáticamente crédito fiscal computable; debe revisarse el respaldo fiscal.')
    if s['pendientes'] or s['diferencias'] or rep['no_extraidas']:
        st.warning(f'Pendientes: {s["pendientes"]} clasificaciones, {s["diferencias"]} diferencias de saldo y {len(rep["no_extraidas"])} líneas no extraídas. Revisar la hoja Para revisar.')
    else:
        st.success('Los movimientos seleccionados se extrajeron, clasificaron y conciliaron sin diferencias.')
    if not all_period:
        st.caption('Los acumulados impresos por el banco se comparan solamente cuando se procesa todo el archivo.')

st.markdown('<div class="aie-foot">Herramienta para uso interno AIE San Justo | Developer Alfonso Alderete</div>', unsafe_allow_html=True)
