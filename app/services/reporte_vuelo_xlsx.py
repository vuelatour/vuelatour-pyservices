"""Reporte consolidado de UN vuelo en Excel (openpyxl), por secciones."""

from __future__ import annotations

from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from app.schemas.reportes import ReporteVueloRequest

BRAND = "DC2626"
NAVY = "102A43"
LIGHT = "F0F4F8"
MONEY = '"$"#,##0.00'


def render_reporte_vuelo_xlsx(r: ReporteVueloRequest) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = f"Vuelo {r.folio}"[:31]
    ws.column_dimensions["A"].width = 22
    for col in ("B", "C", "D", "E"):
        ws.column_dimensions[col].width = 18

    row = 1

    def titulo(texto: str) -> None:
        nonlocal row
        c = ws.cell(row=row, column=1, value=texto)
        c.font = Font(bold=True, color="FFFFFF", size=12)
        c.fill = PatternFill("solid", fgColor=BRAND)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
        row += 1

    def kv(label: str, value) -> None:
        nonlocal row
        ws.cell(row=row, column=1, value=label).font = Font(color="627D98")
        ws.cell(row=row, column=2, value=value)
        row += 1

    def header(cols: list[str]) -> None:
        nonlocal row
        for i, h in enumerate(cols, start=1):
            c = ws.cell(row=row, column=i, value=h)
            c.font = Font(bold=True, color=NAVY)
            c.fill = PatternFill("solid", fgColor=LIGHT)
        row += 1

    def money_cell(rr: int, cc: int, v) -> None:
        cell = ws.cell(row=rr, column=cc, value=v)
        cell.number_format = MONEY
        cell.alignment = Alignment(horizontal="right")

    # ===== Encabezado =====
    h = ws.cell(row=row, column=1, value=f"Reporte de vuelo #{r.folio}")
    h.font = Font(bold=True, size=14, color=BRAND)
    row += 1
    ws.cell(row=row, column=1, value=f"Generado {r.generado} · hora de Cancún (UTC-5)").font = Font(
        color="627D98", italic=True
    )
    row += 2

    # ===== Resumen =====
    titulo("Resumen del vuelo")
    kv("Cliente", r.cliente or "—")
    kv("Ruta", r.ruta or "—")
    kv("Tipo / Estado", f"{r.tipo} · {r.estado}".strip(" ·"))
    kv("Aeronave", r.aeronave or "—")
    pil = r.piloto or "—"
    if r.copiloto:
        pil += f" / {r.copiloto} (copiloto)"
    kv("Piloto", pil)
    kv("Pasajeros", r.pasajeros)
    if r.pasajeros_nombres:
        kv("Nombres pasajeros", r.pasajeros_nombres)
    kv("Fecha de vuelo", r.fecha_vuelo or "—")
    kv("Traslado final", r.fecha_traslado_final or "—")
    row += 1

    # ===== Cotización =====
    titulo("Cotización")
    if r.tarifa_tipo:
        kv("Tarifa", r.tarifa_tipo)
    if r.tarifa_hora_usd:
        kv("USD / hora", r.tarifa_hora_usd)
    if r.tiempo_cobrable_hr:
        kv("Tiempo cobrable (hr)", round(r.tiempo_cobrable_hr, 2))
    for label, val in [
        ("Subtotal USD", r.subtotal_usd),
        ("TUAS USD", r.tuas_usd),
        ("Pernocta USD", r.viaticos_pernocta_usd),
        ("Extras USD", r.extras_total_usd),
        ("Ajuste USD", r.ajuste_final_usd),
        ("IVA USD", r.iva_usd),
        ("Total USD", r.total_usd),
    ]:
        ws.cell(row=row, column=1, value=label).font = Font(color="627D98")
        money_cell(row, 2, val)
        row += 1
    if r.total_mxn:
        ws.cell(row=row, column=1, value="Total MXN").font = Font(color="627D98")
        money_cell(row, 2, r.total_mxn)
        row += 1
    if r.metodo_cobro:
        kv("Método de cobro", r.metodo_cobro)
    row += 1

    # ===== Ingreso (cobros) =====
    titulo("Ingreso (cobros)")
    header(["Fecha", "Método", "Moneda", "Monto"])
    for c in r.cobros:
        ws.cell(row=row, column=1, value=c.fecha or "")
        ws.cell(row=row, column=2, value=c.concepto or "")
        ws.cell(row=row, column=3, value=c.moneda or "")
        money_cell(row, 4, c.monto)
        row += 1
    ws.cell(row=row, column=1, value="Cobrado USD").font = Font(bold=True)
    money_cell(row, 2, r.total_cobrado_usd)
    ws.cell(row=row, column=3, value="Saldo USD").font = Font(bold=True)
    money_cell(row, 4, r.saldo_usd)
    row += 2

    # ===== Tacómetro =====
    titulo("Tacómetro por tramo")
    header(["#", "Tramo", "Salida", "Llegada", "Horas"])
    for t in r.tramos:
        ws.cell(row=row, column=1, value=t.orden)
        ws.cell(row=row, column=2, value=t.ruta)
        ws.cell(row=row, column=3, value=t.taco_salida)
        ws.cell(row=row, column=4, value=t.taco_llegada)
        ws.cell(row=row, column=5, value=t.horas)
        row += 1
    row += 1

    # ===== Horas cotizadas vs voladas (utilidad operativa) =====
    if r.horas_cotizadas_hr is not None or r.horas_voladas_hr is not None:
        titulo("Horas cotizadas vs voladas")
        header(["Horas cotizadas", "Horas voladas", "Diferencia"])
        ws.cell(row=row, column=1, value=r.horas_cotizadas_hr)
        ws.cell(row=row, column=2, value=r.horas_voladas_hr)
        ws.cell(row=row, column=3, value=r.horas_delta_hr)
        row += 1
        for nota in r.notas_horas:
            ws.cell(row=row, column=1, value=nota)
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
            row += 1
        row += 1

    # ===== Combustible =====
    titulo("Combustible")
    header(["Fecha", "Detalle", "Moneda", "Monto"])
    for c in r.combustible:
        ws.cell(row=row, column=1, value=c.fecha or "")
        ws.cell(row=row, column=2, value=c.detalle or c.concepto or "")
        ws.cell(row=row, column=3, value=c.moneda or "")
        money_cell(row, 4, c.monto)
        row += 1
    row += 1

    # ===== Gastos =====
    titulo("Gastos")
    header(["Fecha", "Categoría", "Proveedor", "Moneda", "Monto"])
    for g in r.gastos:
        ws.cell(row=row, column=1, value=g.fecha or "")
        ws.cell(row=row, column=2, value=g.concepto or "")
        ws.cell(row=row, column=3, value=g.detalle or "")
        ws.cell(row=row, column=4, value=g.moneda or "")
        money_cell(row, 5, g.monto)
        row += 1

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()
