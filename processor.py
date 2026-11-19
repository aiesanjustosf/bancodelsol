"""Extractor auditable de resúmenes de cuenta corriente del formato Banco 310 (PDF texto)."""
from __future__ import annotations
from dataclasses import dataclass, asdict
from decimal import Decimal, ROUND_HALF_UP
from datetime import datetime
from collections import defaultdict
import re
import fitz

CENT = Decimal('0.01')
DATE = re.compile(r'^\s*(\d{2}/\d{2}/\d{4})\s+(.+?)\s{2,}(\d{4,})\s+([\d,]+\.\d{2})\s+(-?[\d,]+\.\d{2})\s*$')
INITIAL = re.compile(r'^\s*\d{2}/\d{2}/\d{4}\s+SALDO INICIAL\s+(-?[\d,]+\.\d{2})\s*$')
FINAL = re.compile(r'SALDO\s+ACREEDOR\s+CONTABLE\s+AL\s+\d{2}/\d{2}/\d{4}\s+(-?[\d,]+\.\d{2})', re.I)
DEBIT_TAX = re.compile(r'Idem,\s*al\s*D[eé]bito:\s*([\d,]+\.\d{2})', re.I)
CREDIT_TAX = re.compile(r'Impuesto\s+Ley\s+25\.413\s+al\s+Cr[eé]dito:\s*([\d,]+\.\d{2})', re.I)
SIRCREB_TAX = re.compile(r'Acumulado\s+por\s+Impuesto\s+Sircreb:\s*([\d,]+\.\d{2})', re.I)


def money(s):
    return Decimal(s.replace(',', '')).quantize(CENT)


def category(text):
    t = text.upper()
    if 'LEY NO 25.413' in t or 'LEY N° 25.413' in t or 'LEY 25.413' in t:
        return 'Ley 25.413 - crédito' if 'CREDITO' in t or 'CRÉDITO' in t else 'Ley 25.413 - débito'
    if 'SIRCREB' in t: return 'SIRCREB'
    if 'IMPUESTO AL VALOR AGREGADO' in t or re.search(r'\bIVA\b',t): return 'IVA bancario'
    if 'INTERESES POR GIRO EN DESCUBIERTO' in t: return 'Intereses por descubierto'
    if 'INTERES' in t or 'INTERÉS' in t: return 'Otros intereses'
    if 'DESEMBOLSO PRESTAMO' in t or 'DESEMBOLSO PRÉSTAMO' in t: return 'Desembolso de préstamo'
    if 'COBRANZAS CREDITOS' in t or 'COBRANZAS CRÉDITOS' in t or 'CUOTA PRESTAMO' in t: return 'Cobro de préstamo/cuota'
    if 'TRANSFERENCIA RECIBIDA' in t: return 'Transferencia recibida'
    if 'TRANSFERENCIA' in t: return 'Otras transferencias'
    if 'COMISION' in t or 'COMISIÓN' in t: return 'Comisiones'
    if 'PERCEPCION' in t or 'PERCEPCIÓN' in t: return 'Percepciones'
    if 'RETENCION' in t or 'RETENCIÓN' in t: return 'Retenciones'
    return 'Sin clasificar'


def tax_role(cat):
    if cat == 'IVA bancario': return 'IVA a verificar'
    if cat.startswith('Ley 25.413'): return 'Impuesto 25.413 a verificar'
    if cat == 'SIRCREB': return 'Pago a cuenta IIBB a verificar'
    if cat == 'Percepciones': return 'Percepción a verificar'
    if cat == 'Retenciones': return 'Retención a verificar'
    return ''


@dataclass
class Movement:
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
    tratamiento_fiscal: str
    pagina: int
    archivo: str
    observacion: str = ''

    def record(self):
        r=asdict(self)
        for k in ('debito','credito','saldo','saldo_calculado','diferencia'):
            r[k]=float(r[k])
        return r


def extract_pdf(raw: bytes, filename='resumen.pdf'):
    try:
        pdf = fitz.open(stream=raw,filetype='pdf')
    except Exception as e:
        raise ValueError('No se pudo abrir el PDF: '+str(e)) from e
    alltext='\n'.join(page.get_text(sort=True) for page in pdf)
    if not alltext.strip():
        raise ValueError('El PDF no contiene texto extraíble. No se aplicó OCR: requiere revisión manual.')
    warnings=[]
    bank_match=re.search(r'Clave Bancaria Uniforme:\s*(.+)',alltext)
    acct=re.search(r'Cuenta:\s*([\d/]+)',alltext)
    holder=re.search(r'Cuenta:\s*[\d/]+\s*\[[^]]+\]\s*-\s*(.+)',alltext)
    initial=None; final=None; rows=[]; unmatched=[]
    for i,page in enumerate(pdf,1):
        for line in page.get_text(sort=True).splitlines():
            if initial is None:
                m=INITIAL.match(line)
                if m:initial=money(m.group(1))
            m=DATE.match(line)
            if not m:
                if re.match(r'^\s*\d{2}/\d{2}/\d{4}\s+',line) and 'SALDO INICIAL' not in line:
                    unmatched.append({'pagina':i,'texto':line.strip()})
                continue
            ds, detail, ref, amount, balance=m.groups()
            dt=datetime.strptime(ds,'%d/%m/%Y')
            rows.append((dt,ds,detail.strip(),ref,money(amount),money(balance),i))
    m=FINAL.search(alltext)
    if m: final=money(m.group(1))
    tax_totals={}
    for key, pat in [('ley_credito',CREDIT_TAX),('ley_debito',DEBIT_TAX),('sircreb',SIRCREB_TAX)]:
        m=pat.search(alltext)
        if m:tax_totals[key]=money(m.group(1))
    if initial is None: warnings.append('Saldo inicial no encontrado; no se certifica el primer movimiento.')
    if unmatched: warnings.append(f'{len(unmatched)} líneas con fecha no pudieron extraerse; revisar antes de usar resultados.')
    if not rows: raise ValueError('No se encontraron movimientos en el formato esperado. Verificá que sea el resumen compatible.')
    output=[]
    previous=initial
    for dt,ds,detail,ref,amount,balance,page in rows:
        debit=credit=Decimal('0.00')
        diff=Decimal('0.00'); obs=''
        if previous is None:
            obs='No hay saldo inicial para determinar el signo; primer movimiento sin validar.'
            # No asignar a débito ni crédito si falta contexto.
            calculated=balance
        else:
            change=(balance-previous).quantize(CENT)
            if change == -amount: debit=amount
            elif change == amount: credit=amount
            else:
                # Conservar importe sin asignar y marcar discrepancia.
                obs=f'No coincide variación de saldo ({change}) con importe de movimiento ({amount}).'
            calculated=(previous+credit-debit).quantize(CENT)
            diff=(balance-calculated).quantize(CENT)
        cat=category(detail)
        output.append(Movement(ds,dt.strftime('%Y-%m'),detail,ref,debit,credit,balance,calculated,diff,cat,tax_role(cat),page,filename,obs))
        previous=balance
    if final is not None and final!=rows[-1][5]: warnings.append('Saldo final informado distinto al último saldo de los movimientos.')
    if final is None: warnings.append('Saldo final contable no localizado; verificar si el PDF es completo.')
    if any(r.diferencia for r in output):warnings.append('Se detectaron diferencias en la conciliación de saldos.')
    if any(r.categoria=='Sin clasificar' for r in output):warnings.append('Hay movimientos sin clasificar; revisar en la pestaña de auditoría.')
    return {
        'archivo':filename,'cuenta':acct.group(1) if acct else 'No identificada',
        'titular':holder.group(1).strip() if holder else 'No identificado',
        'cbu':bank_match.group(1).strip() if bank_match else 'No identificada',
        'saldo_inicial':initial,'saldo_final_informado':final,
        'movimientos':output,'no_extraidas':unmatched,'avisos':warnings,
        'totales_banco':tax_totals,'paginas':len(pdf),
    }


def merge_reports(reports):
    if not reports:raise ValueError('No hay archivos.')
    accounts={(x['cuenta'],x['titular']) for x in reports}
    if len(accounts)>1:raise ValueError('Los PDFs corresponden a diferentes cuentas o titulares. Procesá cada cuenta por separado.')
    # PDFs repetidos y períodos superpuestos: bloquear para no duplicar transacciones.
    seen=set(); mov=[]
    for report in reports:
        for r in report['movimientos']:
            key=(r.fecha,r.detalle,r.referencia,r.debito,r.credito,r.saldo)
            if key in seen:raise ValueError('Se encontraron movimientos duplicados al combinar PDFs; revisá períodos repetidos.')
            seen.add(key);mov.append(r)
    periods=defaultdict(list)
    for m in mov: periods[m.periodo].append(m)
    return dict(sorted(periods.items()))


def summarize(movements, pct_debit=Decimal('0'), pct_credit=Decimal('0')):
    """Montos fiscales estimativos configurables; ninguna presunción de computabilidad legal."""
    bycat=defaultdict(lambda:Decimal('0'))
    for m in movements:
        bycat[m.categoria]+=m.debito if m.debito else m.credito
    d=sum((x.debito for x in movements),Decimal('0'))
    c=sum((x.credito for x in movements),Decimal('0'))
    ley_d=bycat['Ley 25.413 - débito']; ley_c=bycat['Ley 25.413 - crédito']
    return {
        'movimientos':len(movements),'debitos':d,'creditos':c,'neto':c-d,
        'iva_a_verificar':bycat['IVA bancario'],
        'ley_debito':ley_d,'ley_credito':ley_c,'sircreb':bycat['SIRCREB'],
        'intereses_descubierto':bycat['Intereses por descubierto'],
        'otros_intereses':bycat['Otros intereses'],
        'comisiones':bycat['Comisiones'],
        'percepciones':bycat['Percepciones'],
        'retenciones':bycat['Retenciones'],
        'cuotas_prestamo':bycat['Cobro de préstamo/cuota'],
        'desembolsos_prestamo':bycat['Desembolso de préstamo'],
        'transferencias':bycat['Transferencia recibida'],
        'ley_credito_fiscal_estimado':(ley_d*pct_debit/100 + ley_c*pct_credit/100).quantize(CENT,rounding=ROUND_HALF_UP),
        'movimientos_desconocidos':sum(x.categoria=='Sin clasificar' for x in movements),
        'inconsistencias':sum(x.diferencia!=0 or x.observacion!='' for x in movements),
        'saldo_minimo':min((x.saldo for x in movements),default=None),
        'saldo_ultimo':movements[-1].saldo if movements else None,
        'movimientos_saldo_negativo':sum(x.saldo<0 for x in movements),
        'clasificados':dict(bycat),
    }
