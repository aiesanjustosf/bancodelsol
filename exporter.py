"""Generación Excel con datos trazables, sin sumar subtotales al total de débitos."""
from io import BytesIO
from decimal import Decimal
from collections import defaultdict
import xlsxwriter
from processor import summarize

NAVY='#172B4D'; BLUE='#2457A6'; LIGHT='#E9F0FA'; ORANGE='#FFF1D8'
CURRENCY='#,##0.00;[Red](#,##0.00);–'


def export_excel(report, movements, pct_debit=Decimal('0'), pct_credit=Decimal('0')):
    stream=BytesIO()
    wb=xlsxwriter.Workbook(stream, {'in_memory':True})
    wb.set_properties({'title':'Control de resumen bancario AIE','comments':'Importes extraídos del PDF. Tratamiento fiscal sujeto a revisión.'})
    title=wb.add_format({'bold':True,'font_size':17,'font_color':NAVY})
    sub=wb.add_format({'bold':True,'font_color':'white','bg_color':NAVY,'border':0,'text_wrap':True,'valign':'vcenter'})
    val=wb.add_format({'num_format':CURRENCY})
    date=wb.add_format({'num_format':'dd/mm/yyyy'})
    warn=wb.add_format({'bg_color':ORANGE,'text_wrap':True})
    top=wb.add_format({'bold':True,'font_color':BLUE})
    plain=wb.add_format({'text_wrap':True,'valign':'top'})
    summary=summarize(movements,pct_debit,pct_credit)
    def make(name, headers, widths):
        sh=wb.add_worksheet(name);sh.freeze_panes(3,0);sh.hide_gridlines(2)
        sh.write(0,0,name,title)
        for i,h in enumerate(headers):sh.write(2,i,h,sub);sh.set_column(i,i,widths[i])
        sh.set_row(2,32)
        return sh
    sh=make('Resumen operativo',['Concepto','Importe / cantidad','Nota contable'],[39,23,76])
    metrics=[
        ('Cantidad de movimientos',summary['movimientos'],'Movimientos extraídos'),
        ('Total débitos',summary['debitos'],'Incluye egresos, impuestos, cuotas, intereses y comisiones'),
        ('Total créditos',summary['creditos'],'Incluye transferencias y desembolsos de préstamos'),
        ('Movimiento neto',summary['neto'],'Créditos menos débitos'),
        ('Intereses por descubierto',summary['intereses_descubierto'],'Gasto financiero separado del IVA'),
        ('Otros intereses',summary['otros_intereses'],'Revisar destino financiero'),
        ('Comisiones bancarias',summary['comisiones'],'Neto de IVA cuando se presenta desglosado'),
        ('IVA bancario a verificar',summary['iva_a_verificar'],'Potencial crédito fiscal; revisar vinculación y documentación'),
        ('Ley 25.413 - débito',summary['ley_debito'],'Impuesto bancario cobrado al débito'),
        ('Ley 25.413 - crédito',summary['ley_credito'],'Impuesto bancario cobrado al crédito'),
        ('SIRCREB',summary['sircreb'],'Potencial pago a cuenta de IIBB'),
        ('Percepciones',summary['percepciones'],'Sujetos a revisión'),
        ('Retenciones',summary['retenciones'],'Sujetos a revisión'),
        ('Cobranza de créditos / cuotas',summary['cuotas_prestamo'],'NO tratar íntegramente como gasto: capital + intereses no desagregados'),
        ('Desembolsos de préstamos',summary['desembolsos_prestamo'],'Financiamiento recibido, no ingreso operativo'),
        ('Transferencias recibidas',summary['transferencias'],'No necesariamente ingreso gravado'),
        ('Estimación Ley 25.413 computable',summary['ley_credito_fiscal_estimado'],f'Parámetros del usuario: débito {pct_debit}% / crédito {pct_credit}%. No constituye liquidación fiscal'),
        ('Movimientos sin clasificar',summary['movimientos_desconocidos'],'Revisión manual necesaria'),
        ('Inconsistencias de movimientos',summary['inconsistencias'],'No usar hasta resolver si es distinto de cero'),
        ('Saldo final según movimientos',summary['saldo_ultimo'],'Último saldo extraído'),
    ]
    for i,(key,value,note) in enumerate(metrics,3):
        sh.write(i,0,key)
        if isinstance(value,(int,float,Decimal)):sh.write_number(i,1,float(value),val if isinstance(value,Decimal) else None)
        sh.write(i,2,note,plain)
    sh.autofilter(2,0,2+len(metrics),2)
    headers=['Fecha','Período','Detalle original','Referencia','Débito','Crédito','Saldo PDF','Saldo reconstruido','Diferencia','Categoría','Tratamiento fiscal','Página','Archivo','Observaciones']
    widths=[14,13,58,16,20,20,20,22,18,31,35,9,32,63]
    m=make('Movimientos',headers,widths)
    for i,x in enumerate(movements,3):
        from datetime import datetime
        values=[datetime.strptime(x.fecha,'%d/%m/%Y'),x.periodo,x.detalle,x.referencia,x.debito,x.credito,x.saldo,x.saldo_calculado,x.diferencia,x.categoria,x.tratamiento_fiscal,x.pagina,x.archivo,x.observacion]
        for j,v in enumerate(values):
            if isinstance(v,datetime):m.write_datetime(i,j,v,date)
            elif isinstance(v,Decimal):m.write_number(i,j,float(v),val)
            elif isinstance(v,int):m.write_number(i,j,v)
            else:m.write_string(i,j,str(v),warn if j==13 and v else None)
    if movements:m.autofilter(2,0,2+len(movements),len(headers)-1)
    groups=defaultdict(list)
    for x in movements:groups[x.periodo].append(x)
    monthly=make('Por mes',['Período','Cantidad','Débitos','Créditos','Neto','Intereses descubierto','IVA','Ley 25.413 débito','Ley 25.413 crédito','SIRCREB','Cuotas / cobros créditos','Desembolsos préstamos','Saldo al cierre','Alertas'],[15,12,21,21,21,23,21,22,22,18,25,25,22,18])
    for i,(month,items) in enumerate(sorted(groups.items()),3):
        s=summarize(items,pct_debit,pct_credit)
        vals=[month,s['movimientos'],s['debitos'],s['creditos'],s['neto'],s['intereses_descubierto'],s['iva_a_verificar'],s['ley_debito'],s['ley_credito'],s['sircreb'],s['cuotas_prestamo'],s['desembolsos_prestamo'],s['saldo_ultimo'],s['inconsistencias']]
        for j,v in enumerate(vals):
            if isinstance(v,Decimal):monthly.write_number(i,j,float(v),val)
            elif isinstance(v,int):monthly.write_number(i,j,v)
            else:monthly.write(i,j,v)
    if groups:monthly.autofilter(2,0,2+len(groups),13)
    fiscal=make('Créditos fiscales',['Fecha','Período','Concepto','Débito','Crédito','A revisar como','Referencia','Página'],[15,14,50,22,22,34,20,10])
    candidates=[x for x in movements if x.tratamiento_fiscal]
    for i,x in enumerate(candidates,3):
        for j,v in enumerate([x.fecha,x.periodo,x.detalle,x.debito,x.credito,x.tratamiento_fiscal,x.referencia,x.pagina]):
            if isinstance(v,Decimal):fiscal.write_number(i,j,float(v),val)
            elif isinstance(v,int):fiscal.write_number(i,j,v)
            else:fiscal.write(i,j,v)
    if candidates:fiscal.autofilter(2,0,2+len(candidates),7)
    loans=make('Prestamos',['Fecha','Período','Movimiento','Débito (cuota/cobro)','Crédito (desembolso)','Referencia','Página','Advertencia'],[16,14,57,25,26,17,10,75])
    creditrows=[x for x in movements if x.categoria in ('Cobro de préstamo/cuota','Desembolso de préstamo','Intereses por descubierto','Otros intereses')]
    for i,x in enumerate(creditrows,3):
        row=[x.fecha,x.periodo,x.detalle,x.debito,x.credito,x.referencia,x.pagina,'Sin contrato/cuadro de amortización no puede separarse capital e interés' if x.categoria=='Cobro de préstamo/cuota' else 'Interés por descubierto: no atribuir a préstamo' if x.categoria=='Intereses por descubierto' else 'Confirmar naturaleza contractual']
        for j,v in enumerate(row):
            if isinstance(v,Decimal):loans.write_number(i,j,float(v),val)
            elif isinstance(v,int):loans.write_number(i,j,v)
            else:loans.write(i,j,v)
    if creditrows:loans.autofilter(2,0,2+len(creditrows),7)
    ctrl=make('Controles',['Control','Valor PDF','Valor calculado','Diferencia','Resultado'],[43,25,25,24,65])
    verifications=[('Saldo final',report.get('saldo_final_informado'),movements[-1].saldo if movements else None),('Ley 25.413 - débito',report.get('totales_banco',{}).get('ley_debito'),summary['ley_debito']),('Ley 25.413 - crédito',report.get('totales_banco',{}).get('ley_credito'),summary['ley_credito']),('SIRCREB',report.get('totales_banco',{}).get('sircreb'),summary['sircreb'])]
    for i,(name,bank,computed) in enumerate(verifications,3):
        ctrl.write(i,0,name)
        if bank is not None:ctrl.write_number(i,1,float(bank),val)
        if computed is not None:ctrl.write_number(i,2,float(computed),val)
        if bank is not None and computed is not None:
            diff=bank-computed;ctrl.write_number(i,3,float(diff),val);ctrl.write(i,4,'OK' if diff==0 else 'REVISAR diferencia',warn if diff else None)
        else:ctrl.write(i,4,'No informado / no verificable')
    start=9
    ctrl.write(start,0,'Verificación de movimientos',top)
    ctrl.write(start,1,'Con diferencias')
    ctrl.write_number(start,2,summary['inconsistencias'])
    if report.get('avisos'):
        for i,w in enumerate(report['avisos'],start+3):ctrl.write(i,0,'AVISO');ctrl.write(i,1,w,warn)
    pending=make('Auditoria',['Página','Línea no extraída / movimiento observado','Tipo'],[12,115,26])
    audit=[(str(x.pagina),x.observacion+' | '+x.fecha+' '+x.detalle,'Movimiento') for x in movements if x.observacion]
    audit += [(str(x['pagina']),x['texto'],'Línea sin extraer') for x in report.get('no_extraidas',[])]
    audit += [(str(x.pagina),x.fecha+' '+x.detalle,'Sin clasificar') for x in movements if x.categoria=='Sin clasificar']
    for i,(page,txt,kind) in enumerate(audit,3):
        pending.write(i,0,page);pending.write(i,1,txt,plain);pending.write(i,2,kind)
    note=make('Criterios',['Concepto','Criterio aplicado'],[34,111])
    notes=[
        ('Fuente','Datos obtenidos exclusivamente del PDF cargado, sin conexión con el banco.'),
        ('Formato','Compatibilidad comprobada para este formato de resumen textual. Otros bancos requieren adaptación.'),
        ('Comprobación','Cada operación se valida contra la diferencia entre el saldo previo y el saldo siguiente.'),
        ('IVA','IVA bancario expuesto separado. Su computabilidad depende de la documentación y la actividad del contribuyente.'),
        ('Ley 25.413','Estimación basada en porcentajes provistos manualmente. Requiere revisión normativa, condición del contribuyente y período.'),
        ('SIRCREB','Se expone por separado. La imputación contra IIBB está sujeta a jurisdicción y reglas aplicables.'),
        ('Préstamos','Desembolsos y cobranzas de cuotas no son ingresos o gastos operativos sin análisis. No se desglosan capital/intereses sin soporte.'),
        ('Saldo negativo','Indica utilización de descubierto en la cuenta, NO saldo de préstamo contratado ni límite disponible.'),
        ('Subtotal','Los renglones de impuestos/intereses son clasificaciones de movimientos incluidos en total débitos: NO deben sumarse otra vez.'),
        ('PDF incompleto','Si falta saldo inicial/final o existen líneas no extraídas, la auditoría lo advierte.'),
        ('Uso','Herramienta para uso interno AIE - Developer Alfonso Alderete.'),
    ]
    for i,(k,v) in enumerate(notes,3):note.write(i,0,k,top);note.write(i,1,v,plain);note.set_row(i,29)
    wb.close();stream.seek(0);return stream.getvalue()
