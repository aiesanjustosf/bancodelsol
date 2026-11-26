import unittest
from decimal import Decimal
from io import BytesIO
from zipfile import ZipFile
from xml.etree import ElementTree as ET
from pathlib import Path
from sys import path
path.insert(0,str(Path(__file__).resolve().parents[1]))
from processor import extract_pdf
from operativo import make_summary, operation_rows
from excel_export import export_excel
from pdf_export import export_pdf

class TestExports(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        # Tests con un resumen bancario sintético mínimo del formato Montegrosso.
        import fitz
        doc=fitz.open();page=doc.new_page();page.insert_text((30,30),
            'Cuenta: 37151/3 [20-10685397-6] - MONTEGROSSO, LUIS ALBERTO\n'
            '01/07/2026 SALDO INICIAL 100.00\n'
            '02/07/2026 INTERESES POR GIRO EN DESCUBIERTO    10188    100.00    0.00\n'
            '02/07/2026 IMPUESTO AL VALOR AGREGADO    10188    10.50    -10.50\n'
            '03/07/2026 LEY No 25.413 POR OPERACIONES DE DEBITO    10202    5.00    -15.50\n'
            'SALDO ACREEDOR CONTABLE AL 03/07/2026 -15.50', fontsize=9)
        # Synthetic PDF may be rejected if bank format differs; use assembled movements instead.
        cls.fakepdf=doc.tobytes();doc.close()
    def test_totals(self):
        from processor import Movement,ZERO
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
                if x.endswith('.xml'):ET.fromstring(archive.read(x))
            ops=archive.read('xl/worksheets/sheet1.xml')
            self.assertIn(b'115.50',ops)
            self.assertIn(b'SUM(',ops)
        pdf=export_pdf(s,'2026-07','PRUEBA','37151/3')
        self.assertTrue(pdf.startswith(b'%PDF'))
        import fitz
        d=fitz.open(stream=pdf,filetype='pdf')
        self.assertEqual(len(d),1)
        self.assertIn('Registraci',d[0].get_text())
        d.close()

if __name__=='__main__':unittest.main()
