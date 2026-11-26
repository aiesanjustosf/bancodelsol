"""Exportación PDF en la estética simple del resumen AIE Galicia."""
from io import BytesIO
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, KeepTogether
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from operativo import operation_rows, money_ar

DARK=colors.HexColor('#1E2730')
GREY=colors.HexColor('#D7D7D7')

def export_pdf(summary, period='', titular='', cuenta=''):
    buf=BytesIO()
    doc=SimpleDocTemplate(buf,pagesize=A4,topMargin=36,bottomMargin=42,leftMargin=48,rightMargin=48,
                          title='Resumen Operativo: Registración Módulo IVA',author='AIE San Justo')
    title=ParagraphStyle('title',fontName='Helvetica-Bold',fontSize=18,leading=22,textColor=DARK,spaceAfter=14)
    small=ParagraphStyle('small',fontName='Helvetica',fontSize=8,leading=11,textColor=colors.HexColor('#666666'))
    h=ParagraphStyle('head',fontName='Helvetica-Bold',fontSize=10.5,leading=12)
    cell=ParagraphStyle('cell',fontName='Helvetica',fontSize=9,leading=11)
    val=ParagraphStyle('val',fontName='Helvetica',fontSize=9,leading=11,alignment=TA_RIGHT)
    bold=ParagraphStyle('bold',fontName='Helvetica-Bold',fontSize=9,leading=11)
    boldval=ParagraphStyle('boldval',fontName='Helvetica-Bold',fontSize=9,leading=11,alignment=TA_RIGHT)
    elements=[Paragraph('Resumen Operativo: Registración Módulo IVA',title)]
    meta=[]
    if titular:meta.append(f'Titular: {titular}')
    if cuenta:meta.append(f'Cuenta: {cuenta}')
    if period:meta.append(f'Período: {period}')
    if meta:
        elements.extend([Paragraph(' | '.join(meta),small),Spacer(1,12)])
    table_data=[[Paragraph('Concepto',h),Paragraph('Importe',h)]]
    ops=operation_rows(summary)
    for key,label,amount,subtotal in ops:
        main=key=='total'
        # Los brutos y el total de Ley son subtotales informativos; se destacan tenuemente.
        table_data.append([Paragraph(label,bold if main else cell),Paragraph(money_ar(amount),boldval if main else val)])
    t=Table(table_data,colWidths=[365,135],repeatRows=1,hAlign='LEFT')
    styles=[('BACKGROUND',(0,0),(-1,0),GREY),('BOX',(0,0),(-1,-1),.7,colors.HexColor('#989898')),
            ('INNERGRID',(0,0),(-1,-1),.35,colors.HexColor('#A6A6A6')),
            ('VALIGN',(0,0),(-1,-1),'MIDDLE'),('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5),
            ('LEFTPADDING',(0,0),(-1,-1),8),('RIGHTPADDING',(0,0),(-1,-1),8),
            ('BACKGROUND',(0,len(ops)),(-1,len(ops)),colors.HexColor('#E7ECF0'))]
    t.setStyle(TableStyle(styles))
    elements += [t, Spacer(1,12),Paragraph('TOTAL: suma de brutos y demás conceptos una sola vez. Los subtotales no se duplican.',small),
                 Spacer(1,17),Paragraph('Herramienta para uso interno - AIE San Justo',small)]
    doc.build(elements)
    return buf.getvalue()
