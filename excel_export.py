"""Generador XLSX real mediante Office Open XML (ZIP/XML de librería estándar).

Evita dependencias de escritoras de Excel en Streamlit Cloud. Incluye fórmulas
con valores precalculados, datos numéricos y detalle auditable.
"""
from __future__ import annotations
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED
from xml.sax.saxutils import escape
from decimal import Decimal
from operativo import make_summary, operation_rows

MONEY_STYLE=3
BOLD_STYLE=4
TOTAL_STYLE=5


def text(value, style=0):return ('s',str(value if value is not None else ''),style)
def num(value, style=MONEY_STYLE):return ('n',value,style)
def formula(expr, cached, style=MONEY_STYLE):return ('f',expr.lstrip('='),cached,style)

def col(n):
    out=''
    while n:
        n,r=divmod(n-1,26);out=chr(65+r)+out
    return out

def number(value):
    if isinstance(value, Decimal):return str(value.quantize(Decimal('.01')))
    return str(value)

def xcell(c,r,v):
    ref=f'{col(c)}{r}'
    kind=v[0];style=v[2] if kind!='f' else v[3]
    attrs=f'r="{ref}" s="{style}"'
    if kind=='n':return f'<c {attrs}><v>{number(v[1])}</v></c>'
    if kind=='f':return f'<c {attrs}><f>{escape(v[1])}</f><v>{number(v[2])}</v></c>'
    content=escape(v[1]).replace('\r',' ')
    return f'<c {attrs} t="inlineStr"><is><t xml:space="preserve">{content}</t></is></c>'

class Tab:
    def __init__(self,name,widths,freeze=4):
        self.name=name;self.widths=widths;self.freeze=freeze;self.rows={};self.merges=[];self.autofilter=None
    def put(self,r,values):self.rows[r]=values
    def merge(self,range_):self.merges.append(range_)
    def xml(self):
        highest=max(self.rows) if self.rows else 1
        data=[]
        for r,row in sorted(self.rows.items()):
            data.append(f'<row r="{r}">'+''.join(xcell(i+1,r,val) for i,val in enumerate(row) if val is not None)+'</row>')
        columns=''.join(f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>' for i,w in enumerate(self.widths,1))
        merges='' if not self.merges else '<mergeCells count="%d">%s</mergeCells>'%(len(self.merges),''.join(f'<mergeCell ref="{m}"/>' for m in self.merges))
        pane=f'<sheetViews><sheetView workbookViewId="0"><pane ySplit="{self.freeze}" topLeftCell="A{self.freeze+1}" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' if self.freeze else '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
        autofil=f'<autoFilter ref="{self.autofilter}"/>' if self.autofilter else ''
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                f'<dimension ref="A1:{col(len(self.widths))}{highest}"/>{pane}<sheetFormatPr defaultRowHeight="17"/>'
                f'<cols>{columns}</cols><sheetData>{"".join(data)}</sheetData>{merges}{autofil}</worksheet>')

def styles_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00;[Red](#,##0.00);0.00"/></numFmts>'
    '<fonts count="4">'
    '<font><sz val="10"/><name val="Aptos"/></font>'
    '<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Aptos"/></font>'
    '<font><b/><sz val="16"/><color rgb="FF153152"/><name val="Aptos"/></font>'
    '<font><b/><sz val="10"/><color rgb="FF153152"/><name val="Aptos"/></font></fonts>'
    '<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FF153152"/><bgColor indexed="64"/></patternFill></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FFE9F1F7"/><bgColor indexed="64"/></patternFill></fill></fills>'
    '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>'
    '<border><left style="thin"><color rgb="FFD2D8DD"/></left><right style="thin"><color rgb="FFD2D8DD"/></right>'
    '<top style="thin"><color rgb="FFD2D8DD"/></top><bottom style="thin"><color rgb="FFD2D8DD"/></bottom><diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="7">'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1"/>'
    '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1"/>'
    '<xf numFmtId="164" fontId="3" fillId="3" borderId="1" xfId="0" applyNumberFormat="1"/>'
    '<xf numFmtId="164" fontId="3" fillId="2" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1"/>'
    '<xf numFmtId="0" fontId="3" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1"/>'
    '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')

def package(tabs):
    stream=BytesIO();ns='http://schemas.openxmlformats.org/'
    with ZipFile(stream,'w',ZIP_DEFLATED) as z:
        defaults='<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
        overrides='<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        for i,t in enumerate(tabs,1):
            overrides+=f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            z.writestr(f'xl/worksheets/sheet{i}.xml',t.xml())
        z.writestr('[Content_Types].xml','<?xml version="1.0" encoding="utf-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'+defaults+overrides+'</Types>')
        z.writestr('_rels/.rels','<?xml version="1.0" encoding="utf-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="'+ns+'officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        z.writestr('xl/_rels/workbook.xml.rels','<?xml version="1.0" encoding="utf-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+''.join(f'<Relationship Id="rId{i}" Type="{ns}officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1,len(tabs)+1))+f'<Relationship Id="rId{len(tabs)+1}" Type="{ns}officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
        sheetxml=''.join(f'<sheet name="{escape(t.name)}" sheetId="{i}" r:id="rId{i}"/>' for i,t in enumerate(tabs,1))
        z.writestr('xl/workbook.xml','<?xml version="1.0" encoding="utf-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'+sheetxml+'</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')
        z.writestr('xl/styles.xml',styles_xml())
    return stream.getvalue()


def export_excel(report,movements,periodo='Todos los períodos',selected_all=True):
    s=make_summary(movements)
    account=report.get('cuenta','No identificada')
    holder=report.get('titular','No identificado')
    preamble=f'Cuenta {account} | {holder} | {periodo} | Valores en ARS'
    tabs=[]
    # Movimientos auditables. Débitos y créditos conservan signos en columnas distintas.
    mtab=Tab('Movimientos',[15,13,68,16,21,21,23,33,24,24,24,55,35,10])
    mtab.put(1,[text('AIE - Movimientos bancarios',2)]);mtab.merge('A1:N1')
    mtab.put(2,[text(preamble)]);mtab.merge('A2:N2')
    mtab.put(4,[text(h,1) for h in ['Fecha','Período','Concepto original','Referencia','Débito','Crédito','Saldo','Clasificación','Naturaleza','Alícuota','IVA asociado','Observación / control','Archivo','Página PDF']])
    for i,m in enumerate(movements,5):
        mtab.put(i,[text(m.fecha),text(m.periodo),text(m.detalle),text(m.referencia),num(m.debito),num(m.credito),num(m.saldo),
                    text(m.categoria),text(m.naturaleza),text(m.alicuota),num(m.iva_vinculado),text(m.observacion or m.criterio_iva),text(m.archivo),num(m.pagina,0)])
    mlast=max(5,len(movements)+4); mtab.autofilter=f'A4:N{mlast}'
    tabs.append(mtab)
    # Gastos: base efectivamente debitada + IVA separado, con vinculación por referencia.
    gtab=Tab('Gastos',[15,65,17,23,17,23,23,60,28])
    gtab.put(1,[text('AIE - Gastos netos e IVA',2)]);gtab.merge('A1:I1')
    gtab.put(2,[text(preamble)]);gtab.merge('A2:I2')
    gtab.put(4,[text(h,1) for h in ['Fecha','Detalle bancario','Referencia','Neto gasto','Alícuota','IVA asociado','Bruto','Criterio de alícuota','Vinculación IVA']])
    for i,m in enumerate(s['gastos'],5):
        gtab.put(i,[text(m.fecha),text(m.detalle),text(m.referencia),num(m.debito),text(m.alicuota),num(m.iva_vinculado),
                    formula(f'D{i}+F{i}',m.debito+m.iva_vinculado),text(m.criterio_iva),text(m.vinculacion_iva)])
    glast=max(5,len(s['gastos'])+4);gtab.autofilter=f'A4:I{glast}'
    tabs.append(gtab)
    # Resumen contable en A:B, fiel al PDF de AIE, con fórmulas que se recalculan en Excel.
    oper=Tab('Resumen operativo',[52,26],freeze=3)
    oper.put(1,[text('Resumen Operativo: Registración Módulo IVA',2)]);oper.merge('A1:B1')
    oper.put(2,[text(preamble)]);oper.merge('A2:B2')
    oper.put(3,[text('Concepto',1),text('Importe',1)])
    key_to_row={}
    operation=operation_rows(s)
    for r,(key,label,amount,subtotal) in enumerate(operation,4):
        key_to_row[key]=r
        oper.put(r,[text(label,6 if subtotal else 0),num(amount,TOTAL_STYLE if key=='total' else BOLD_STYLE if subtotal else MONEY_STYLE)])
    expense_net=f"'Gastos'!$D$5:$D${glast}"; expense_iva=f"'Gastos'!$F$5:$F${glast}"; expense_rate=f"'Gastos'!$E$5:$E${glast}"
    cats=f"'Movimientos'!$H$5:$H${mlast}"; debit=f"'Movimientos'!$E$5:$E${mlast}"; credit=f"'Movimientos'!$F$5:$F${mlast}"
    def tax(cat):return f'SUMIF({cats},"{cat}",{debit})-SUMIF({cats},"{cat}",{credit})'
    formulas={}
    for label,slug in [('21%','21'),('10,5%','105'),('27%','27')]:
        formulas[f'neto_{slug}']=f'SUMIF({expense_rate},"{label}",{expense_net})'
        formulas[f'iva_{slug}']=f'SUMIF({expense_rate},"{label}",{expense_iva})'
        formulas[f'bruto_{slug}']=f'B{key_to_row[f"neto_{slug}"]}+B{key_to_row[f"iva_{slug}"]}'
    formulas.update(exento=f'SUMIF({expense_rate},"Exento",{expense_net})',
                    no_gravado=f'SUMIF({expense_rate},"No gravado",{expense_net})',
                    sin_iva=f'SUM({expense_net})-SUM(B{key_to_row["neto_21"]},B{key_to_row["neto_105"]},B{key_to_row["neto_27"]},B{key_to_row["exento"]},B{key_to_row["no_gravado"]})',
                    iva_sin_vincular=f'{tax("IVA bancario")}-SUM({expense_iva})',
                    percepcion_iva=tax('Percepción de IVA'),
                    ley_debito=tax('Ley 25.413 - débito'),ley_credito=tax('Ley 25.413 - crédito'),
                    ley_neto=f'B{key_to_row["ley_debito"]}+B{key_to_row["ley_credito"]}',
                    sircreb=tax('SIRCREB'),otras_percepciones=tax('Otras percepciones'),
                    retenciones=tax('Retenciones'))
    all_total=['bruto_21','bruto_105','bruto_27','exento','no_gravado','sin_iva','iva_sin_vincular','percepcion_iva','ley_neto','sircreb','otras_percepciones','retenciones']
    formulas['total']='SUM('+','.join('B'+str(key_to_row[k]) for k in all_total)+')'
    for key,label,amount,subtotal in operation:
        r=key_to_row[key];oper.rows[r][1]=formula(formulas[key],amount,TOTAL_STYLE if key=='total' else BOLD_STYLE if subtotal else MONEY_STYLE)
    tabs.insert(0,oper)
    # Control de préstamos y descubiertos. No equivale a límite de crédito disponible.
    ctab=Tab('Créditos bancarios',[19,60,29,23,23,23,17],freeze=10)
    ctab.put(1,[text('Control de créditos bancarios y descubierto',2)]);ctab.merge('A1:G1')
    ctab.put(2,[text(preamble)]);ctab.merge('A2:G2')
    for r,title,value in [
        (4,'¿Hay indicios de créditos / préstamos?', 'SÍ' if s['tiene_credito_bancario'] else 'NO'),
        (5,'Desembolsos de préstamos (acreditaciones)',s['prestamos_desembolsados']),
        (6,'Cobranzas / cuotas de préstamos (débitos)',s['prestamos_cobrados']),
        (7,'Saldo de cuenta más negativo',s['descubierto_minimo']),
        (8,'Límite y deuda pendiente','No determinados por el resumen bancario'),
    ]:ctab.put(r,[text(title),num(value) if isinstance(value,Decimal) else text(value)])
    ctab.put(10,[text(h,1) for h in ['Fecha','Concepto original','Tipo','Débito','Crédito','Saldo','Referencia']])
    loans=[m for m in movements if m.categoria in ('Desembolso de préstamo','Pago de préstamo / cuota')]
    for i,m in enumerate(loans,11):ctab.put(i,[text(m.fecha),text(m.detalle),text(m.categoria),num(m.debito),num(m.credito),num(m.saldo),text(m.referencia)])
    ctab.autofilter=f'A10:G{max(11,10+len(loans))}'
    tabs.append(ctab)
    # Conciliación por movimientos y control de acumulados publicados por el banco.
    control=Tab('Control y conciliación',[53,30,68],freeze=3)
    control.put(1,[text('Control de saldos e impuestos',2)]);control.merge('A1:C1')
    control.put(2,[text(preamble)]);control.merge('A2:C2')
    control.put(3,[text('Concepto',1),text('Importe / resultado',1),text('Detalle',1)])
    first=movements[0] if movements else None
    last=movements[-1] if movements else None
    initial=(first.saldo-first.credito+first.debito) if first else Decimal('0')
    final=last.saldo if last else initial
    recalculated=initial+s['creditos']-s['debitos']
    bank_final=report.get('saldo_final_informado') if selected_all else None
    bank_ley=report.get('totales_banco',{}).get('ley_debito') if selected_all else None
    entries=[
        ('Saldo inicial del período',initial,'Derivado del primer movimiento de la selección'),
        ('Créditos de la cuenta',s['creditos'],'No equivale a ingresos operativos'),
        ('Débitos de la cuenta',s['debitos'],'No equivale al total de gastos bancarios'),
        ('Saldo final calculado',recalculated,'Saldo inicial + créditos - débitos'),
        ('Saldo final del último movimiento',final,'Tomado del extracto'),
        ('Diferencia de saldos',recalculated-final,'Debe ser cero'),
        ('Saldo final informado por el banco',bank_final,'Disponible para período completo' if selected_all else 'No aplicable al filtro'),
        ('Diferencia contra saldo final del banco',final-bank_final if bank_final is not None else None,'Debe ser cero para el período completo'),
        ('Ley 25.413 débito calculado',s['ley_debito'],'Suma neta de cargos identificados'),
        ('Ley 25.413 débito informado por el banco',bank_ley,'Tomado del pie del resumen'),
        ('Diferencia Ley 25.413 débito',s['ley_debito']-bank_ley if bank_ley is not None else None,'Debe ser cero para el período completo'),
        ('Movimientos con diferencia',s['diferencias'],'Debe ser cero'),
        ('Movimientos no clasificados',s['pendientes'],'Revisión manual'),
        ('Líneas del PDF no extraídas',len(report.get('no_extraidas',[])),'Revisión obligatoria si es mayor que cero'),
        ('IVA sin vincular',s['iva_sin_vincular'],'Comprobar documentación / alícuota'),
    ]
    for r,(label,value,obs) in enumerate(entries,4):
        control.put(r,[text(label),num(value) if isinstance(value,Decimal) else text('—' if value is None else value),text(obs)])
    tabs.append(control)
    pending=Tab('Para revisar',[15,66,19,23,24,60],freeze=4)
    pending.put(1,[text('Operaciones pendientes de revisar',2)]);pending.merge('A1:F1')
    pending.put(2,[text('Si no hay registros, los movimientos del PDF quedaron clasificados y conciliados')]);pending.merge('A2:F2')
    pending.put(4,[text(h,1) for h in ['Fecha','Concepto','Referencia','Débito','Categoría','Motivo']])
    warns=[(m,'Sin categoría confiable') for m in movements if m.naturaleza=='Pendiente revisión']+[(m,'IVA sin vinculación inequívoca') for m in movements if m.categoria=='IVA bancario' and not m.vinculacion_iva]+[(m,'Diferencia de saldos') for m in movements if m.diferencia]
    for r,(m,reason) in enumerate(warns,5):pending.put(r,[text(m.fecha),text(m.detalle),text(m.referencia),num(m.debito),text(m.categoria),text(reason)])
    pending.autofilter=f'A4:F{max(5,len(warns)+4)}'
    tabs.append(pending)
    return package(tabs)
