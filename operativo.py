"""Resumen Operativo AIE: reproduce la registración del módulo IVA del ejemplo Galicia.

Brutos son subtotales INFORMATIVOS: el total incluye cada gasto y cada tributo una vez.
Sin respaldo documental no se presume exención, alícuota fiscal definitiva o crédito computable.
"""
from collections import defaultdict
from decimal import Decimal
from processor import ZERO, CENT, sum_dec

RATES=('21%', '10,5%', '27%')

def make_summary(movements):
    expenses=[m for m in movements if m.naturaleza=='Gasto' and m.debito>0]
    bases={rate:sum_dec(m.debito for m in expenses if m.alicuota==rate) for rate in RATES}
    iva={rate:sum_dec(m.iva_vinculado for m in expenses if m.alicuota==rate) for rate in RATES}
    exento=sum_dec(m.debito for m in expenses if m.alicuota=='Exento')
    no_gravado=sum_dec(m.debito for m in expenses if m.alicuota=='No gravado')
    no_identificado=sum_dec(m.debito for m in expenses if m.alicuota not in (*RATES,'Exento','No gravado'))
    def taxed(category):
        selected=[m for m in movements if m.categoria==category]
        return sum_dec(m.debito-m.credito for m in selected)
    ley_deb=taxed('Ley 25.413 - débito')
    ley_cred=taxed('Ley 25.413 - crédito')
    percep=taxed('Percepción de IVA')
    sircreb=taxed('SIRCREB')
    otras=taxed('Otras percepciones')
    retenciones=taxed('Retenciones')
    iva_total=taxed('IVA bancario')
    iva_paired=sum_dec(iva.values())
    iva_unpaired=iva_total-iva_paired
    total=sum_dec([*bases.values(),*iva.values(),exento,no_gravado,no_identificado,iva_unpaired,percep,ley_deb,ley_cred,sircreb,otras,retenciones])
    # El total debe ser igual a la suma de todos los débitos/ajustes operativos.
    return dict(bases=bases,iva=iva,exento=exento,no_gravado=no_gravado,
                sin_identificar=no_identificado,iva_sin_vincular=iva_unpaired,
                percepcion_iva=percep,sircreb=sircreb,otras_percepciones=otras,
                retenciones=retenciones,ley_debito=ley_deb,ley_credito=ley_cred,
                ley_neto=ley_deb+ley_cred,total=total,
                tiene_credito_bancario=any(m.categoria in ('Desembolso de préstamo','Pago de préstamo / cuota') or m.saldo<0 for m in movements),
                prestamos_desembolsados=sum_dec(m.credito for m in movements if m.categoria=='Desembolso de préstamo'),
                prestamos_cobrados=sum_dec(m.debito for m in movements if m.categoria=='Pago de préstamo / cuota'),
                descubierto_minimo=min((m.saldo for m in movements),default=ZERO),
                movimientos=len(movements),debitos=sum_dec(m.debito for m in movements),
                creditos=sum_dec(m.credito for m in movements),
                pendientes=sum(1 for m in movements if m.naturaleza=='Pendiente revisión'),
                diferencias=sum(1 for m in movements if m.diferencia),
                gastos=expenses)

def operation_rows(s):
    """Tuplas (clave, rótulo, importe, es_subtotal), orden exacto AIE.

    Los valores 'Bruto' y los desglores Ley Débito/Crédito son informativos:
    TOTAL incluye los brutos y el total neto Ley una sola vez.
    """
    r=[]
    for rate,slug in [('21%','21'),('10,5%','105'),('27%','27')]:
        r += [(f'neto_{slug}',f'Neto gastos / comisiones {rate}',s['bases'][rate],False),
              (f'iva_{slug}',f'IVA {rate}',s['iva'][rate],False),
              (f'bruto_{slug}',f'Bruto {rate}',s['bases'][rate]+s['iva'][rate],True)]
    r += [('exento','Gastos exentos',s['exento'],False),
          ('no_gravado','Gastos no gravados',s['no_gravado'],False),
          ('sin_iva','Gastos sin IVA identificado',s['sin_identificar'],False),
          ('iva_sin_vincular','IVA sin vincular (revisar)',s['iva_sin_vincular'],False),
          ('percepcion_iva','Percepciones de IVA (RG 3337 / RG 2408)',s['percepcion_iva'],False),
          ('ley_debito','Ley 25.413 - débito',s['ley_debito'],False),
          ('ley_credito','Ley 25.413 - crédito',s['ley_credito'],False),
          ('ley_neto','Ley 25.413 (total D y C)',s['ley_neto'],True),
          ('sircreb','SIRCREB',s['sircreb'],False),
          ('otras_percepciones','Otras percepciones',s['otras_percepciones'],False),
          ('retenciones','Retenciones',s['retenciones'],False),
          ('total','TOTAL',s['total'],True)]
    return r


def money_ar(value):
    return f'{value:,.2f}'.replace(',','\u00a7').replace('.',',').replace('\u00a7','.')
