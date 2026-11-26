import unittest
from decimal import Decimal
from io import BytesIO
from zipfile import ZipFile
from xml.etree import ElementTree as ET
from pathlib import Path
from sys import path
path.insert(0,str(Path(__file__).resolve().parents[1]))
from operativo import make_summary
from excel_export import export_excel
from pdf_export import export_pdf
from openpyxl import load_workbook

class TestExports(unittest.TestCase):
    def test_totals(self):
        from processor import Movement, ZERO
        def mk(i,cat,nature,debit,credit=0,rate='No aplica',iva=0):
            return Movement(str(i),'01/07/2026','2026-07',cat,str(i),Decimal(str(debit)),Decimal(str(credit)),ZERO,ZERO,ZERO,cat,nature,cat,rate,'',Decimal(str(iva)),'',1,'example.pdf')
        rows=[mk(1,'Intereses por descubierto','Gasto','100','0','10,5%','10.50'),
              mk(2,'IVA bancario','Impuesto','10.50'),
              mk(3,'Ley 25.413 - débito','Impuesto','5'),
              mk(4,'Desembolso de préstamo','Crédito bancario','0','1000')]
        s=make_summary(rows)
        self.assertEqual(s['total'],Decimal('115.50'))
        self.assertEqual(s['prestamos_desembolsados'],Decimal('1000.00'))
        self.assertEqual(s['bases']['10,5%'],Decimal('100.00'))
        self.assertEqual(s['iva']['10,5%'],Decimal('10.50'))
        self.assertEqual(s['ley_debito'],Decimal('5.00'))
        self.assertEqual(s['iva_sin_vincular'],Decimal('0.00'))
        report={'cuenta':'37151/3','titular':'PRUEBA','saldo_final_informado':None,'totales_banco':{},'no_extraidas':[]}
        xlsx=export_excel(report,rows)
        with ZipFile(BytesIO(xlsx)) as archive:
            self.assertEqual(archive.testzip(),None)
            self.assertEqual(len([x for x in archive.namelist() if x.startswith('xl/worksheets/sheet')]),6)
            for x in archive.namelist():
                if x.endswith('.xml'):
                    ET.fromstring(archive.read(x))
        wb=load_workbook(BytesIO(xlsx), data_only=True)
        self.assertEqual(wb.sheetnames[0], 'Resumen operativo')
        ws=wb['Resumen operativo']
        self.assertEqual(ws['A1'].value, 'Resumen Operativo: Registración Módulo IVA')
        self.assertAlmostEqual(float(ws['B24'].value), 115.50, places=2)
        pdf=export_pdf(s,'2026-07','PRUEBA','37151/3')
        self.assertTrue(pdf.startswith(b'%PDF'))
        import fitz
        d=fitz.open(stream=pdf,filetype='pdf')
        self.assertEqual(len(d),1)
        self.assertIn('Registraci',d[0].get_text())
        d.close()

if __name__=='__main__':
    unittest.main()
