"""Reporte mensual por avión en Excel (doc 5.10 / etapa 7).

Reutiliza el mismo payload del reparto de utilidades (RepartoPdfRequest) y arma
un libro con: una hoja resumen por avión con las 6 secciones (ingresos, gastos
directos/indirectos, permisos, otros, reserva overhaul, saldo) y el reparto por
socio de cada avión.
"""

from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from app.schemas.reparto import RepartoPdfRequest

BRAND = "0F4C81"
LIGHT = "EEF2F7"
WHITE = "FFFFFF"
MONEY = '"$"#,##0.00'

_thin = Side(style="thin", color="D5DBE3")
_border = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)


def _title(ws, text: str, row: int, span: int, size: int = 14):
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = Font(bold=True, size=size, color=BRAND)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)


def _header_row(ws, row: int, headers: list[str]):
    for col, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=col, value=h)
        c.font = Font(bold=True, color=WHITE)
        c.fill = PatternFill("solid", fgColor=BRAND)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = _border


def render_reparto_xlsx(req: RepartoPdfRequest) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumen por avión"

    _title(ws, "VuelaTour — Reporte mensual por avión", 1, 11, size=16)
    sub = ws.cell(row=2, column=1, value=f"Periodo: {req.periodo_desde} a {req.periodo_hasta}  ·  Generado: {req.generado}")
    sub.font = Font(italic=True, size=10, color="5B6470")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=11)

    headers = [
        "Matrícula", "Modelo", "Horas voladas", "Ingresos cobrado", "Pendiente cobro",
        "Gastos directos", "Gastos indirectos", "Permisos", "Otros (prorrateo)",
        "Reserva overhaul", "Saldo disponible",
    ]
    hrow = 4
    _header_row(ws, hrow, headers)

    money_cols = list(range(4, 12))  # columnas D..K son montos (C = horas)
    totals = {c: 0.0 for c in money_cols}

    r = hrow + 1
    for a in req.aviones:
        valores = [
            a.matricula, a.modelo, a.horas_voladas_hr, a.ingresos_cobrado_usd, a.pendiente_cobro_usd,
            a.gastos_directos_usd, a.gastos_indirectos_usd, a.permisos_usd,
            a.otros_usd, a.reserva_overhaul_usd, a.saldo_usd,
        ]
        for col, v in enumerate(valores, start=1):
            c = ws.cell(row=r, column=col, value=v)
            c.border = _border
            if col == 3:
                c.number_format = "0.0"
            if col in money_cols:
                c.number_format = MONEY
                totals[col] += float(v or 0)
        r += 1

    # Fila de totales.
    tcell = ws.cell(row=r, column=1, value="TOTAL")
    tcell.font = Font(bold=True)
    tcell.fill = PatternFill("solid", fgColor=LIGHT)
    ws.cell(row=r, column=2).fill = PatternFill("solid", fgColor=LIGHT)
    for col in money_cols:
        c = ws.cell(row=r, column=col, value=totals[col])
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor=LIGHT)
        c.number_format = MONEY
        c.border = _border

    # Anchos de columna.
    widths = [12, 16, 13, 16, 15, 15, 16, 12, 16, 16, 16]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Reparto por socio (debajo del resumen).
    r += 2
    _title(ws, "Reparto por socio", r, 4, size=12)
    r += 1
    _header_row(ws, r, ["Avión", "Socio", "Porcentaje", "Monto USD"])
    r += 1
    for a in req.aviones:
        for linea in a.reparto:
            ws.cell(row=r, column=1, value=a.matricula).border = _border
            ws.cell(row=r, column=2, value=linea.socio_nombre).border = _border
            pc = ws.cell(row=r, column=3, value=(linea.porcentaje or 0) / 100)
            pc.number_format = "0.00%"
            pc.border = _border
            mc = ws.cell(row=r, column=4, value=linea.monto_usd)
            mc.number_format = MONEY
            mc.border = _border
            r += 1

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()
