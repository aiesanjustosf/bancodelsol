"""Exportación XLSX robusta mediante openpyxl.

Genera un archivo compatible con Excel sin construir XML manualmente.
"""
from __future__ import annotations
from io import BytesIO
from decimal import Decimal
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from operativo import make_summary, operation_rows

DARK = '153152'
LIGHT = 'E9F1F7'
HEADER_FILL = PatternFill('solid', fgColor=DARK)
SUB_FILL = PatternFill('solid', fgColor=LIGHT)
TOTAL_FILL = PatternFill('solid', fgColor=DARK)
WHITE_FONT = Font(color='FFFFFF', bold=True)
TITLE_FONT = Font(color=DARK, bold=True, size=16)
BOLD_FONT = Font(color=DARK, bold=True)
THIN = Side(style='thin', color='D2D8DD')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY_FORMAT = '#,##0.00;[Red](#,##0.00);0.00'


def _as_excel_number(value):
    if isinstance(value, Decimal):
        return float(value)
    return value


def _title(ws, title, preamble, end_col):
    ws['A1'] = title
    ws['A1'].font = TITLE_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_col)
    ws['A2'] = preamble
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=end_col)


def _header(ws, row_idx, labels):
    for col_idx, label in enumerate(labels, 1):
        c = ws.cell(row_idx, col_idx, label)
        c.fill = HEADER_FILL
        c.font = WHITE_FONT
        c.border = BORDER
        c.alignment = Alignment(horizontal='left', vertical='center')


def _money_cell(cell, value, bold=False, fill=None):
    cell.value = _as_excel_number(value)
    cell.number_format = MONEY_FORMAT
    cell.border = BORDER
    cell.alignment = Alignment(horizontal='right')
    if bold:
        cell.font = BOLD_FONT
    if fill is not None:
        cell.fill = fill


def _text_cell(cell, value, bold=False, fill=None):
    cell.value = value
    cell.border = BORDER
    cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    if bold:
        cell.font = BOLD_FONT
    if fill is not None:
        cell.fill = fill


def _autowidth(ws, widths):
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width


def _add_table(ws, start_row, end_row, end_col):
    if end_row <= start_row:
        return
    ref = f'A{start_row}:{get_column_letter(end_col)}{end_row}'
    name = f'Tabla{abs(hash((ws.title, ref))) % 10**8}'
    tab = Table(displayName=name, ref=ref)
    style = TableStyleInfo(name='TableStyleMedium2', showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    tab.tableStyleInfo = style
    ws.add_table(tab)


def export_excel(report, movements, periodo='Todos los períodos', selected_all=True):
    s = make_summary(movements)
    account = report.get('cuenta', 'No identificada')
    holder = report.get('titular', 'No identificado')
    preamble = f'Cuenta {account} | {holder} | {periodo} | Valores en ARS'

    wb = Workbook()
    ws = wb.active
    ws.title = 'Resumen operativo'

    # Hoja 1: Resumen operativo
    _title(ws, 'Resumen Operativo: Registración Módulo IVA', preamble, 2)
    _header(ws, 3, ['Concepto', 'Importe'])
    row = 4
    for key, label, amount, subtotal in operation_rows(s):
        _text_cell(ws.cell(row, 1), label, bold=subtotal or key == 'total', fill=SUB_FILL if subtotal and key != 'total' else TOTAL_FILL if key == 'total' else None)
        if key == 'total':
            ws.cell(row, 1).font = WHITE_FONT
        _money_cell(ws.cell(row, 2), amount, bold=subtotal or key == 'total', fill=SUB_FILL if subtotal and key != 'total' else TOTAL_FILL if key == 'total' else None)
        if key == 'total':
            ws.cell(row, 2).font = WHITE_FONT
        row += 1
    _autowidth(ws, [48, 22])
    ws.freeze_panes = 'A4'

    # Hoja 2: Movimientos
    ws = wb.create_sheet('Movimientos')
    _title(ws, 'AIE - Movimientos bancarios', preamble, 14)
    headers = ['Fecha', 'Período', 'Concepto original', 'Referencia', 'Débito', 'Crédito', 'Saldo', 'Clasificación', 'Naturaleza', 'Alícuota', 'IVA asociado', 'Observación / control', 'Archivo', 'Página PDF']
    _header(ws, 4, headers)
    row = 5
    for m in movements:
        vals = [m.fecha, m.periodo, m.detalle, m.referencia, m.debito, m.credito, m.saldo, m.categoria, m.naturaleza, m.alicuota, m.iva_vinculado, m.observacion or m.criterio_iva, m.archivo, m.pagina]
        for col, val in enumerate(vals, 1):
            cell = ws.cell(row, col)
            if col in (5, 6, 7, 11):
                _money_cell(cell, val)
            else:
                _text_cell(cell, val)
        row += 1
    _autowidth(ws, [13, 12, 55, 14, 16, 16, 16, 26, 20, 15, 15, 40, 28, 10])
    ws.freeze_panes = 'A5'
    _add_table(ws, 4, max(5, row - 1), len(headers))

    # Hoja 3: Gastos
    ws = wb.create_sheet('Gastos')
    _title(ws, 'AIE - Gastos netos e IVA', preamble, 9)
    headers = ['Fecha', 'Detalle bancario', 'Referencia', 'Neto gasto', 'Alícuota', 'IVA asociado', 'Bruto', 'Criterio de alícuota', 'Vinculación IVA']
    _header(ws, 4, headers)
    row = 5
    for m in s['gastos']:
        bruto = m.debito + m.iva_vinculado
        vals = [m.fecha, m.detalle, m.referencia, m.debito, m.alicuota, m.iva_vinculado, bruto, m.criterio_iva, m.vinculacion_iva]
        for col, val in enumerate(vals, 1):
            cell = ws.cell(row, col)
            if col in (4, 6, 7):
                _money_cell(cell, val)
            else:
                _text_cell(cell, val)
        row += 1
    _autowidth(ws, [13, 50, 14, 16, 14, 16, 16, 38, 22])
    ws.freeze_panes = 'A5'
    _add_table(ws, 4, max(5, row - 1), len(headers))

    # Hoja 4: Créditos bancarios
    ws = wb.create_sheet('Créditos bancarios')
    _title(ws, 'Control de créditos bancarios y descubierto', preamble, 7)
    info_rows = [
        ('¿Hay indicios de créditos / préstamos?', 'SÍ' if s['tiene_credito_bancario'] else 'NO'),
        ('Desembolsos de préstamos (acreditaciones)', s['prestamos_desembolsados']),
        ('Cobranzas / cuotas de préstamos (débitos)', s['prestamos_cobrados']),
        ('Saldo de cuenta más negativo', s['descubierto_minimo']),
        ('Límite y deuda pendiente', 'No determinados por el resumen bancario'),
    ]
    row = 4
    for label, value in info_rows:
        _text_cell(ws.cell(row, 1), label, bold=True)
        if isinstance(value, Decimal):
            _money_cell(ws.cell(row, 2), value)
        else:
            _text_cell(ws.cell(row, 2), value)
        row += 1
    row += 1
    headers = ['Fecha', 'Concepto original', 'Tipo', 'Débito', 'Crédito', 'Saldo', 'Referencia']
    _header(ws, row, headers)
    table_start = row
    row += 1
    loans = [m for m in movements if m.categoria in ('Desembolso de préstamo', 'Pago de préstamo / cuota')]
    for m in loans:
        vals = [m.fecha, m.detalle, m.categoria, m.debito, m.credito, m.saldo, m.referencia]
        for col, val in enumerate(vals, 1):
            cell = ws.cell(row, col)
            if col in (4, 5, 6):
                _money_cell(cell, val)
            else:
                _text_cell(cell, val)
        row += 1
    _autowidth(ws, [13, 54, 28, 16, 16, 16, 14])
    ws.freeze_panes = f'A{table_start + 1}'
    _add_table(ws, table_start, max(table_start + 1, row - 1), len(headers))

    # Hoja 5: Control y conciliación
    ws = wb.create_sheet('Control y conciliación')
    _title(ws, 'Control de saldos e impuestos', preamble, 3)
    _header(ws, 3, ['Concepto', 'Importe / resultado', 'Detalle'])
    first = movements[0] if movements else None
    last = movements[-1] if movements else None
    initial = (first.saldo - first.credito + first.debito) if first else Decimal('0')
    final = last.saldo if last else initial
    recalculated = initial + s['creditos'] - s['debitos']
    bank_final = report.get('saldo_final_informado') if selected_all else None
    bank_ley = report.get('totales_banco', {}).get('ley_debito') if selected_all else None
    entries = [
        ('Saldo inicial del período', initial, 'Derivado del primer movimiento de la selección'),
        ('Créditos de la cuenta', s['creditos'], 'No equivale a ingresos operativos'),
        ('Débitos de la cuenta', s['debitos'], 'No equivale al total de gastos bancarios'),
        ('Saldo final calculado', recalculated, 'Saldo inicial + créditos - débitos'),
        ('Saldo final del último movimiento', final, 'Tomado del extracto'),
        ('Diferencia de saldos', recalculated - final, 'Debe ser cero'),
        ('Saldo final informado por el banco', bank_final, 'Disponible para período completo' if selected_all else 'No aplicable al filtro'),
        ('Diferencia contra saldo final del banco', (final - bank_final) if bank_final is not None else None, 'Debe ser cero para el período completo'),
        ('Ley 25.413 débito calculado', s['ley_debito'], 'Suma neta de cargos identificados'),
        ('Ley 25.413 débito informado por el banco', bank_ley, 'Tomado del pie del resumen'),
        ('Diferencia Ley 25.413 débito', (s['ley_debito'] - bank_ley) if bank_ley is not None else None, 'Debe ser cero para el período completo'),
        ('Movimientos con diferencia', Decimal(s['diferencias']), 'Debe ser cero'),
        ('Movimientos no clasificados', Decimal(s['pendientes']), 'Revisión manual'),
        ('Líneas del PDF no extraídas', Decimal(len(report.get('no_extraidas', []))), 'Revisión obligatoria si es mayor que cero'),
        ('IVA sin vincular', Decimal(s['iva_sin_vincular']), 'Comprobar documentación / alícuota'),
    ]
    row = 4
    for label, value, obs in entries:
        _text_cell(ws.cell(row, 1), label)
        if isinstance(value, Decimal):
            _money_cell(ws.cell(row, 2), value)
        elif value is None:
            _text_cell(ws.cell(row, 2), '—')
        else:
            _text_cell(ws.cell(row, 2), value)
        _text_cell(ws.cell(row, 3), obs)
        row += 1
    _autowidth(ws, [42, 18, 50])
    ws.freeze_panes = 'A4'

    # Hoja 6: Para revisar
    ws = wb.create_sheet('Para revisar')
    _title(ws, 'Operaciones pendientes de revisar', 'Si no hay registros, los movimientos del PDF quedaron clasificados y conciliados', 6)
    _header(ws, 4, ['Fecha', 'Concepto', 'Referencia', 'Débito', 'Categoría', 'Motivo'])
    warns = [(m, 'Sin categoría confiable') for m in movements if m.naturaleza == 'Pendiente revisión']
    warns += [(m, 'IVA sin vinculación inequívoca') for m in movements if m.categoria == 'IVA bancario' and not m.vinculacion_iva]
    warns += [(m, 'Diferencia de saldos') for m in movements if m.diferencia]
    row = 5
    for m, reason in warns:
        vals = [m.fecha, m.detalle, m.referencia, m.debito, m.categoria, reason]
        for col, val in enumerate(vals, 1):
            cell = ws.cell(row, col)
            if col == 4:
                _money_cell(cell, val)
            else:
                _text_cell(cell, val)
        row += 1
    _autowidth(ws, [13, 55, 14, 16, 24, 32])
    ws.freeze_panes = 'A5'
    _add_table(ws, 4, max(5, row - 1), 6)

    # Opciones generales
    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = True

    out = BytesIO()
    wb.save(out)
    return out.getvalue()
