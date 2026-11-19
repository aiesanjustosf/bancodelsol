import streamlit as st
from decimal import Decimal
from processor import extract_pdf, merge_reports, summarize
from exporter import export_excel

st.set_page_config(page_title='Control Bancario AIE',page_icon='🏦',layout='wide')
st.markdown('''<style>
.block-container{max-width:1380px;padding-top:1.6rem}
.stMetric{border:1px solid #dfe5ef;border-radius:11px;padding:13px}
h1,h2,h3{color:#172B4D}
.aie-foot{color:#657287;border-top:1px solid #dfe5ef;margin-top:40px;padding-top:14px;font-size:13px}
</style>''',unsafe_allow_html=True)

st.title('🏦 Control de resúmenes bancarios | AIE')
st.caption('Procesamiento, control de saldos, resumen operativo, préstamos y créditos fiscales potenciales.')
with st.sidebar:
    st.header('Configuración')
    st.info('Compatible con el formato PDF textual del resumen suministrado (cuenta 37151/3). No es un lector universal de todos los bancos.')
    debit_pct=st.number_input('% Ley 25.413 al débito a estimar como computable',0.0,100.0,0.0,1.0)
    credit_pct=st.number_input('% Ley 25.413 al crédito a estimar como computable',0.0,100.0,0.0,1.0)
    st.caption('Por defecto 0%: la computabilidad depende de normas, período, contribuyente y situación fiscal. No constituye asesoramiento tributario.')
    st.markdown('**Control de crédito bancario:** desembolsos, cobranzas de cuotas y descubierto. Sin contrato ni cuadro de amortización no se determina deuda pendiente, límite o interés implícito.')
files=st.file_uploader('Adjuntá uno o más resúmenes PDF de la misma cuenta (períodos no superpuestos)',type=['pdf'],accept_multiple_files=True)
if not files:
    st.info('Cargá un PDF para comenzar. La app no almacena tus archivos en una base de datos.')
    st.stop()
reports=[]
for f in files:
    try:
        report=extract_pdf(f.getvalue(),f.name)
        reports.append(report)
    except Exception as e:
        st.error(f'Error en {f.name}: {e}')
if len(reports)!=len(files) or not reports:st.stop()
try:
    grouped=merge_reports(reports)
except ValueError as e:
    st.error(str(e));st.stop()
allrows=[x for group in grouped.values() for x in group]
# Orden cronológico con estabilidad original (sort preserva orden de movimientos de igual día).
from datetime import datetime
allrows.sort(key=lambda x:datetime.strptime(x.fecha,'%d/%m/%Y'))
periods=list(grouped)
st.write('**Cuenta:**',reports[0]['cuenta'],' · **Titular:**',reports[0]['titular'],' · **Períodos:**',', '.join(periods))
selected=st.multiselect('Meses para análisis y exportación',periods,default=periods)
selectedrows=[x for x in allrows if x.periodo in selected]
if not selectedrows:st.warning('Seleccioná al menos un período.');st.stop()
summary=summarize(selectedrows,Decimal(str(debit_pct)),Decimal(str(credit_pct)))
def ars(n):
    if n is None:return '—'
    x=f'{n:,.2f}'.replace(',','X').replace('.',',').replace('X','.')
    return '$ '+x
r1=st.columns(4)
r1[0].metric('Débitos',ars(summary['debitos']))
r1[1].metric('Créditos',ars(summary['creditos']))
r1[2].metric('Movimiento neto',ars(summary['neto']))
r1[3].metric('Saldo del último movimiento',ars(summary['saldo_ultimo']))
r2=st.columns(4)
r2[0].metric('Intereses descubierto',ars(summary['intereses_descubierto']))
r2[1].metric('IVA a verificar',ars(summary['iva_a_verificar']))
r2[2].metric('Ley 25.413 débito',ars(summary['ley_debito']))
r2[3].metric('Estimación computable Ley 25.413',ars(summary['ley_credito_fiscal_estimado']))
if summary['inconsistencias']:
    st.error(f'⚠️ {summary["inconsistencias"]} movimientos con inconsistencias o sin signo validado. Revisar antes de utilizar los importes.')
else:st.success(f'✓ {summary["movimientos"]} movimientos sin diferencias en sus saldos individuales.')
for report in reports:
    for warning in report['avisos']:st.warning(f'{report["archivo"]}: {warning}')
if len(selected)!=len(periods):
    st.info('Vista filtrada: los controles contra totales impresos en el PDF solo corresponden al informe íntegro, no al subconjunto de meses.')
if len(reports)==1 and len(selected)==len(periods):
    bank=reports[0]['totales_banco']
    if 'ley_debito' in bank:
        delta=bank['ley_debito']-summary['ley_debito']
        if delta:st.error(f'Impuesto Ley 25.413: diferencia con total impreso {ars(delta)}.')
        else:st.success('✓ Ley 25.413 al débito coincide con el total impreso en el PDF.')

t1,t2,t3,t4,t5=st.tabs(['Resumen operativo','Movimientos','Crédito bancario','Créditos fiscales','Auditoría'])
with t1:
    st.subheader('Detalle operativo sin duplicar gastos e impuestos')
    st.caption('Los gastos financieros, IVA e impuestos forman parte de los débitos: no se suman nuevamente al total.')
    cats=[(k,ars(v)) for k,v in sorted(summary['clasificados'].items())]
    st.dataframe([{'Categoría':k,'Importe':v} for k,v in cats],hide_index=True,use_container_width=True)
    st.markdown('**Por período**')
    rows=[]
    for period in selected:
        s=summarize(grouped[period]);rows.append({'Período':period,'Débitos':ars(s['debitos']),'Créditos':ars(s['creditos']),'Intereses descubierto':ars(s['intereses_descubierto']),'IVA':ars(s['iva_a_verificar']),'Ley 25.413 débito':ars(s['ley_debito']),'Saldo cierre':ars(s['saldo_ultimo'])})
    st.dataframe(rows,hide_index=True,use_container_width=True)
with t2:
    st.subheader('Movimientos bancarios')
    st.dataframe([x.record() for x in selectedrows],hide_index=True,use_container_width=True)
with t3:
    col1,col2,col3=st.columns(3)
    col1.metric('Desembolsos préstamo',ars(summary['desembolsos_prestamo']))
    col2.metric('Cobranzas de créditos / cuotas',ars(summary['cuotas_prestamo']))
    col3.metric('Movimientos con saldo negativo',summary['movimientos_saldo_negativo'])
    st.warning('Una cobranza de crédito puede contener capital, interés, comisiones y otros conceptos. No se infiere descomposición, deuda vigente ni tasa sin contrato/cuadro de amortización. El saldo negativo corresponde a descubierto bancario.')
    st.dataframe([x.record() for x in selectedrows if x.categoria in ('Desembolso de préstamo','Cobro de préstamo/cuota','Intereses por descubierto','Otros intereses')],hide_index=True,use_container_width=True)
with t4:
    st.subheader('Importes fiscales potenciales a revisar')
    st.info('IVA, Ley 25.413, percepciones y SIRCREB se muestran por separado. El cálculo del impuesto computable depende de los porcentajes configurados en el panel lateral; el IVA no se presume computable automáticamente.')
    fiscal=[x.record() for x in selectedrows if x.tratamiento_fiscal]
    st.dataframe(fiscal,hide_index=True,use_container_width=True)
with t5:
    st.write('**Diferencias de saldo:**',summary['inconsistencias'])
    for report in reports:
        st.write(f'**{report["archivo"]}** – {report["paginas"]} páginas; {len(report["movimientos"])} movimientos')
        st.write('Saldo inicial:',ars(report['saldo_inicial']),' · Saldo final declarado:',ars(report['saldo_final_informado']))
        st.write('Totales del PDF:',{k:ars(v) for k,v in report['totales_banco'].items()})
        if report['no_extraidas']:st.dataframe(report['no_extraidas'],hide_index=True,use_container_width=True)
    st.dataframe([x.record() for x in selectedrows if x.observacion or x.categoria=='Sin clasificar'],hide_index=True,use_container_width=True)

# Si se combinan reportes, no atribuir controles impresos de un documento al total multi-PDF.
if len(reports)==1:
    export_report=reports[0]
else:
    export_report={'saldo_final_informado':None,'totales_banco':{},'avisos':['Archivo combinado: controles impresos de PDFs individuales consultables en pantalla.']}
if len(selected)!=len(periods):
    export_report={**export_report,'saldo_final_informado':None,'totales_banco':{},'avisos':export_report.get('avisos',[])+['Filtro aplicado: no comparar totales generales impresos en PDF con un período parcial.']}
content=export_excel(export_report,selectedrows,Decimal(str(debit_pct)),Decimal(str(credit_pct)))
st.download_button('📥 Descargar Excel: resumen operativo + movimientos + controles',data=content,file_name='control_bancario_AIE_'+('_'.join(selected))+'.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',use_container_width=True,type='primary')
st.markdown('<div class="aie-foot">Herramienta para uso interno AIE · Developer Alfonso Alderete · Resultados sujetos a verificación documental y contable.</div>',unsafe_allow_html=True)
