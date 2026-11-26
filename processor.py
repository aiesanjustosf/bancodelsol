"""Lector auditable de la familia de resúmenes bancarios PDF aportada a AIE.

El impuesto no se considera crédito fiscal por defecto. No se infiere exención por
la ausencia de un cargo de IVA. Vinculación IVA <-> gasto por fecha+referencia+importe.
"""
from __future__ import annotations
from dataclasses import asdict, dataclass
from collections import defaultdict
from decimal import Decimal, ROUND_HALF_UP
from datetime import datetime
import re
import fitz

CENT = Decimal('0.01')
ZERO = Decimal('0.00')
RATES = {'21%': Decimal('0.21'), '10,5%': Decimal('0.105'), '27%': Decimal('0.27')}
ALICUOTAS = ['21%', '10,5%', '27%', 'Exento', 'No gravado', 'Sin IVA identificado', 'Pendiente revisión']
CATEGORIAS_GASTO = ('Intereses por descubierto', 'Otros intereses', 'Comisiones', 'Gastos bancarios', 'Seguros y cargos', 'Gastos exentos')
DATE = re.compile(r'^\s*(\d{2}/\d{2}/\d{4})\s+(.+?)\s{2,}(\d{4,})\s+([\d,]+\.\d{2})\s+(-?[\d,]+\.\d{2})\s*$')
INITIAL = re.compile(r'^\s*\d{2}/\d{2}/\d{4}\s+SALDO INICIAL\s+(-?[\d,]+\.\d{2})\s*$', re.I)
FINAL = re.compile(r'SALDO\s+(?:ACREEDOR|DEUDOR)\s+CONTABLE\s+AL\s+\d{2}/\d{2}/\d{4}\s+(-?[\d,]+\.\d{2})', re.I)
DEBIT_TAX = re.compile(r'Idem,\s*al\s*D[eé]bito:\s*([\d,]+\.\d{2})', re.I)
CREDIT_TAX = re.compile(r'Impuesto\s+Ley\s+25\.413\s+al\s+Cr[eé]dito:\s*([\d,]+\.\d{2})', re.I)
SIRCREB_TAX = re.compile(r'Acumulado\s+por\s+Impuesto\s+Sircreb:\s*([\d,]+\.\d{2})', re.I)


def money(value: str) -> Decimal:
    return Decimal(value.replace(',', '')).quantize(CENT, rounding=ROUND_HALF_UP)


def normal(text: str) -> str:
    import unicodedata
    return ''.join(c for c in unicodedata.normalize('NFKD', text.upper()) if not unicodedata.combining(c))


def category(detail: str) -> str:
    t = normal(detail)
    if 'LEY NO 25.413' in t or 'LEY N° 25.413' in t or 'LEY 25.413' in t:
        return 'Ley 25.413 - crédito' if ('DE CREDITO' in t or 'AL CREDITO' in t or 'CREDITOS' in t) else 'Ley 25.413 - débito'
    if 'SIRCREB' in t: return 'SIRCREB'
    if ('PERCEP' in t or 'PERC.' in t) and ('IVA' in t or 'VALOR AGREGADO' in t): return 'Percepción de IVA'
    if 'PERCEP' in t: return 'Otras percepciones'
    if 'RETEN' in t: return 'Retenciones'
    if 'IMPUESTO AL VALOR AGREGADO' in t or re.search(r'\bIVA\b', t): return 'IVA bancario'
    if 'INTERESES POR GIRO EN DESCUBIERTO' in t: return 'Intereses por descubierto'
    if 'INTERES' in t: return 'Otros intereses'
    if 'DESEMBOLSO PRESTAMO' in t or 'ACREDITACION PRESTAMO' in t: return 'Desembolso de préstamo'
    if 'COBRANZAS CREDITOS' in t or 'CUOTA PRESTAMO' in t or 'PAGO PRESTAMO' in t or 'DEBITO PRESTAMO' in t: return 'Pago de préstamo / cuota'
    if 'TRANSFERENCIA RECIBIDA' in t: return 'Transferencia recibida'
    if 'TRANSFERENCIA' in t: return 'Otras transferencias'
    if 'COMISION' in t or 'MANTENIMIENTO CUENTA' in t: return 'Comisiones'
    if 'GASTO BANCARIO' in t or 'GASTOS ADMINISTRATIVOS' in t: return 'Gastos bancarios'
    if 'SEGURO' in t or 'CARGO POR SERVICIO' in t: return 'Seguros y cargos'
    if 'GASTO EXENTO' in t: return 'Gastos exentos'
    return 'Sin clasificar'


def naturaleza(cat: str) -> str:
    if cat in CATEGORIAS_GASTO: return 'Gasto'
    if cat in ('IVA bancario', 'Percepción de IVA', 'Otras percepciones', 'SIRCREB', 'Retenciones') or cat.startswith('Ley 25.413'): return 'Impuesto'
    if cat in ('Desembolso de préstamo', 'Pago de préstamo / cuota'): return 'Crédito bancario'
    if cat in ('Transferencia recibida', 'Otras transferencias'): return 'Movimiento no operativo'
    return 'Pendiente revisión'


@dataclass
class Movement:
    id: str
    fecha: str
    periodo: str
    detalle: str
    referencia: str
    debito: Decimal
    credito: Decimal
    saldo: Decimal
    saldo_calculado: Decimal
    diferencia: Decimal
    categoria: str
    naturaleza: str
    rubro: str
    alicuota: str
    criterio_iva: str
    iva_vinculado: Decimal
    vinculacion_iva: str
    pagina: int
    archivo: str
    observacion: str = ''

    def record(self):
        data=asdict(self)
        for k in ('debito','credito','saldo','saldo_calculado','diferencia','iva_vinculado'):
            data[k]=float(data[k])
        return data


def pair_iva(movements):
    """No vincular por fecha sola ni repartir un IVA entre bases ambiguas."""
    by_key=defaultdict(list)
    for x in movements:
        if x.naturaleza=='Gasto' and x.debito>0:
            by_key[(x.archivo,x.fecha,x.referencia)].append(x)
    paired=0
    pending=0
    used=set()
    for iva in (x for x in movements if x.categoria=='IVA bancario' and x.debito>0):
        candidates=[]
        for expense in by_key.get((iva.archivo,iva.fecha,iva.referencia),[]):
            if expense.id in used: continue
            for rate, ratio in RATES.items():
                if (expense.debito*ratio).quantize(CENT, rounding=ROUND_HALF_UP)==iva.debito:
                    candidates.append((expense, rate))
        if len(candidates)==1:
            expense, rate=candidates[0]
            used.add(expense.id)
            expense.alicuota=rate
            expense.criterio_iva='Inferida matemáticamente; verificar comprobante'
            expense.iva_vinculado=iva.debito
            expense.vinculacion_iva=iva.id
            iva.vinculacion_iva=expense.id
            iva.criterio_iva=f'IVA vinculado a gasto {expense.id} (alícuota inferida {rate})'
            paired+=1
        else:
            iva.criterio_iva='IVA sin asociación unívoca a gasto; revisar documento'
            iva.observacion='No se vinculó IVA con ningún gasto único' if not candidates else 'Vinculación ambigua; varias bases posibles'
            pending+=1
    return paired,pending


def extract_pdf(raw: bytes, filename='resumen.pdf'):
    try:
        pdf=fitz.open(stream=raw, filetype='pdf')
    except Exception as e:
        raise ValueError(f'No se pudo abrir el PDF ({filename}): {e}') from e
    texts=[page.get_text(sort=True) for page in pdf]
    full='\n'.join(texts)
    if not full.strip(): raise ValueError('PDF sin texto seleccionable. No se inventan movimientos mediante OCR.')
    acct=re.search(r'Cuenta:\s*([\d/]+)',full)
    holder=re.search(r'Cuenta:\s*[\d/]+\s*\[[^]]+\]\s*-\s*(.+)',full)
    cbu=re.search(r'Clave Bancaria Uniforme:\s*(.+)',full)
    initial=None; final=None; extracted=[]; missing=[]; warnings=[]
    for page_num,txt in enumerate(texts,1):
        for line in txt.splitlines():
            if initial is None:
                match=INITIAL.match(line)
                if match: initial=money(match.group(1))
            match=DATE.match(line)
            if match:
                dat,det,ref,amount,balance=match.groups()
                extracted.append((dat,det.strip(),ref,money(amount),money(balance),page_num))
            elif re.match(r'^\s*\d{2}/\d{2}/\d{4}\s+',line) and 'SALDO INICIAL' not in line.upper():
                missing.append({'pagina':page_num,'linea':line.strip()})
    if not extracted: raise ValueError('No se localizaron movimientos compatibles. No usar con otro formato bancario sin adaptación.')
    match=FINAL.search(full)
    if match: final=money(match.group(1))
    totals={}
    for key,pat in [('ley_credito',CREDIT_TAX),('ley_debito',DEBIT_TAX),('sircreb',SIRCREB_TAX)]:
        match=pat.search(full)
        if match: totals[key]=money(match.group(1))
    if initial is None: warnings.append('No se identificó saldo inicial: el signo del primer movimiento no puede certificarse.')
    if final is None: warnings.append('No se identificó saldo final contable: verificar integridad del PDF.')
    if missing: warnings.append(f'{len(missing)} líneas con fecha no extraídas; revisar antes de utilizar.')
    mov=[]; prev=initial
    for i,(dat,det,ref,amount,balance,page_num) in enumerate(extracted,1):
        dt=datetime.strptime(dat,'%d/%m/%Y')
        debit=credit=ZERO; observation=''
        if prev is None:
            computed=balance
            observation='Sin saldo anterior para verificar signo del movimiento.'
        else:
            change=(balance-prev).quantize(CENT)
            if change==-amount: debit=amount
            elif change==amount: credit=amount
            else: observation=f'Importe no coincide con variación de saldo: {change} vs {amount}.'
            computed=(prev+credit-debit).quantize(CENT)
        delta=(balance-computed).quantize(CENT)
        cat=category(det); nat=naturaleza(cat)
        rate='Exento' if (nat=='Gasto' and re.search(r'\bEXENT[OA]\b', normal(det)) is not None) else ('Sin IVA identificado' if nat=='Gasto' else 'No aplica')
        if nat=='Gasto' and re.search(r'\bNO GRAVAD[OA]\b', normal(det)) is not None:rate='No gravado'
        mov.append(Movement(
            f'{filename}:{i:05d}',dat,dt.strftime('%Y-%m'),det,ref,debit,credit,balance,computed,delta,
            cat,nat,cat if nat=='Gasto' else '',rate,
            'Exención indicada en concepto' if rate=='Exento' else ('No gravado según concepto' if rate=='No gravado' else ''),
            ZERO,'',page_num,filename,observation
        ))
        prev=balance
    paired,unpaired=pair_iva(mov)
    if final is not None and final!=extracted[-1][4]:warnings.append('El saldo final informado difiere del último movimiento.')
    if any(x.diferencia for x in mov):warnings.append('Existen diferencias en el control de saldos.')
    if unpaired: warnings.append(f'{unpaired} cargos de IVA sin vinculación inequívoca.')
    if any(x.categoria=='Sin clasificar' for x in mov):warnings.append('Hay movimientos sin clasificar; revisar.')
    return {'archivo':filename,'cuenta':acct.group(1) if acct else 'No identificada',
            'titular':holder.group(1).strip() if holder else 'No identificado',
            'cbu':cbu.group(1).strip() if cbu else 'No identificada',
            'saldo_inicial':initial,'saldo_final_informado':final,
            'movimientos':mov,'no_extraidas':missing,'avisos':warnings,
            'totales_banco':totals,'paginas':len(texts),
            'iva_vinculados':paired,'iva_pendientes':unpaired}


def merge_reports(reports):
    if not reports: raise ValueError('No se adjuntaron documentos.')
    accounts={(r['cuenta'],r['titular']) for r in reports}
    if len(accounts)>1: raise ValueError('Distintas cuentas o titulares: procesar por separado.')
    seen=set(); allmov=[]
    for rep in reports:
        for m in rep['movimientos']:
            key=(m.fecha,m.detalle,m.referencia,m.debito,m.credito,m.saldo)
            if key in seen: raise ValueError('Se detectaron movimientos repetidos entre archivos. Evitar resúmenes superpuestos.')
            seen.add(key)
            allmov.append(m)
    allmov.sort(key=lambda x:datetime.strptime(x.fecha,'%d/%m/%Y'))
    months=defaultdict(list)
    for x in allmov: months[x.periodo].append(x)
    return dict(sorted(months.items()))


def apply_review(movements, reviewed_rows):
    """Aplica correcciones explícitas del operador; nunca asigna IVA por simple porcentaje."""
    index={x.id:x for x in movements}
    alerts=[]
    for row in reviewed_rows:
        mov=index.get(row.get('ID'))
        if not mov:continue
        nat=row.get('Naturaleza',mov.naturaleza)
        rate=row.get('Tratamiento IVA',mov.alicuota)
        rubro=(row.get('Rubro de gasto') or mov.rubro).strip()
        if nat not in ('Gasto','Impuesto','Crédito bancario','Movimiento no operativo','Pendiente revisión'):continue
        if rate not in ALICUOTAS+['No aplica']:continue
        if nat=='Gasto' and mov.debito<=0:
            alerts.append(f'{mov.id}: no se convirtió un crédito bancario en gasto.');continue
        if mov.categoria in ('IVA bancario','Percepción de IVA','Otras percepciones','SIRCREB','Retenciones') or mov.categoria.startswith('Ley 25.413'):
            if nat!='Impuesto':
                alerts.append(f'{mov.id}: los impuestos no pueden reclasificarse como gasto neto.');continue
        if nat!=mov.naturaleza: mov.naturaleza=nat
        if nat=='Gasto':
            if rate=='No aplica':rate='Pendiente revisión'
            mov.rubro=rubro or mov.categoria
            if mov.iva_vinculado and rate!=mov.alicuota:
                alerts.append(f'{mov.id}: tratamiento {rate} contradice IVA emparejado; se conserva {mov.alicuota}.')
            else:
                if rate!=mov.alicuota:
                    mov.alicuota=rate
                    mov.criterio_iva='Definido manualmente; verificar respaldo documental'
        else:
            if mov.iva_vinculado:alerts.append(f'{mov.id}: se mantuvo como gasto porque tiene IVA asociado; revisar operación.');mov.naturaleza='Gasto'
            else:
                mov.rubro='';mov.alicuota='No aplica'
    return alerts


def sum_dec(items): return sum(items,ZERO).quantize(CENT)


def summarize(movements,pct_debito=Decimal('0'),pct_credito=Decimal('0')):
    bycat=defaultdict(lambda:ZERO)
    by_rate=defaultdict(lambda:ZERO)
    by_rubro=defaultdict(lambda:ZERO)
    table=defaultdict(lambda:{'neto':ZERO,'iva':ZERO,'cantidad':0})
    for x in movements:
        bycat[x.categoria]+=x.debito if x.debito else x.credito
        if x.naturaleza=='Gasto' and x.debito:
            rate=x.alicuota if x.alicuota in ALICUOTAS else 'Pendiente revisión'
            by_rate[rate]+=x.debito
            by_rubro[x.rubro or x.categoria]+=x.debito
            key=(x.rubro or x.categoria,rate)
            table[key]['neto']+=x.debito
            table[key]['iva']+=x.iva_vinculado
            table[key]['cantidad']+=1
    ley_d=bycat['Ley 25.413 - débito'];ley_c=bycat['Ley 25.413 - crédito']
    return {
        'movimientos':len(movements),
        'debitos':sum_dec(x.debito for x in movements),'creditos':sum_dec(x.credito for x in movements),
        'neto':sum_dec(x.credito-x.debito for x in movements),
        'gastos_neto':sum_dec(x.debito for x in movements if x.naturaleza=='Gasto'),
        'iva_bancario':bycat['IVA bancario'],
        'iva_asociado':sum_dec(x.iva_vinculado for x in movements if x.naturaleza=='Gasto'),
        'ley_debito':ley_d,'ley_credito':ley_c,
        'percepcion_iva':bycat['Percepción de IVA'],
        'percepciones_otras':bycat['Otras percepciones'],
        'sircreb':bycat['SIRCREB'],'retenciones':bycat['Retenciones'],
        'desembolsos_prestamo':bycat['Desembolso de préstamo'],
        'cuotas_prestamo':bycat['Pago de préstamo / cuota'],
        'intereses_descubierto':bycat['Intereses por descubierto'],
        'ley_computable_estimado':(ley_d*Decimal(str(pct_debito))/100+ley_c*Decimal(str(pct_credito))/100).quantize(CENT,rounding=ROUND_HALF_UP),
        'saldo_ultimo':movements[-1].saldo if movements else None,
        'saldo_minimo':min((m.saldo for m in movements),default=None),
        'saldo_negativo_count':sum(x.saldo<0 for x in movements),
        'inconsistencias':sum(x.diferencia!=0 or (bool(x.observacion) and x.categoria!='IVA bancario') for x in movements),
        'iva_sin_vincular':sum(1 for x in movements if x.categoria=='IVA bancario' and not x.vinculacion_iva),
        'sin_clasificar':sum(1 for x in movements if x.naturaleza=='Pendiente revisión'),
        'gastos_por_alicuota':dict(by_rate),
        'gastos_por_rubro':dict(by_rubro),
        'matriz_operativa':dict(sorted(table.items())),
        'categorias':dict(bycat),
        'credito_bancario_detectado':bool(bycat['Desembolso de préstamo'] or bycat['Pago de préstamo / cuota'] or any(x.saldo<0 for x in movements)),
    }
