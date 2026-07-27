"""Balance mensual por avión en Excel (openpyxl) — réplica sistematizada del
control del equipo ("Balance N990GG.xlsx").

Seis hojas en el orden del libro original:
  1. reporte horas <MATRÍCULA> — maestro: 1 fila = 1 vuelo, TOTALES al final.
  2. gastos indirectos  3. otros gastos  4. permisos — resumen + ledger.
  5. balance — bloque A–I del original + reparto real de socios.
  6. pendientes de captura — lo que falta para que el libro quede completo.

Los montos vienen YA calculados del API (aquí solo se pintan; jamás se
recalcula dinero). None = celda vacía — nunca un 0 falso.
"""

from __future__ import annotations

from datetime import datetime
from io import BytesIO

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from app.schemas.reportes import (
    BalanceAvionCobro,
    BalanceAvionHojaGastos,
    BalanceAvionRequest,
)

BRAND = "0F4C81"
NAVY = "102A43"
MUTED = "627D98"
LIGHT = "EEF2F7"
GREEN = "15803D"
RED = "DC2626"
AMBER = "FCD34D"  # horas cobradas < voladas (regla: nunca cobrar de menos)
# Colores suaves por bloque (ayuda visual pedida en el contrato).
FILL_VENTA = "E8F0FB"  # azul suave
FILL_COSTOS = "FDF3E7"  # ámbar suave
FILL_COBROS = "E9F6EC"  # verde suave

MONEY = "#,##0.00"
HORAS = "0.00"
TACO = "0.0"
TC = "0.0000"

_thin = Side(style="thin", color="D5DBE3")
_border = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)


def _fecha(s: str | None) -> str | None:
    """ISO date → dd/mm/aaaa; texto libre (multi-día '9-10 sep') tal cual."""
    if not s:
        return None
    try:
        if len(s) == 10:
            return datetime.fromisoformat(s).strftime("%d/%m/%Y")
        return datetime.fromisoformat(s.replace("Z", "+00:00")).strftime("%d/%m/%Y")
    except ValueError:
        return s


def _title(ws: Worksheet, text: str, row: int, span: int, size: int = 14) -> None:
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = Font(bold=True, size=size, color=BRAND)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)


def _header_row(ws: Worksheet, row: int, headers: list[str], start: int = 1) -> None:
    for col, h in enumerate(headers, start=start):
        c = ws.cell(row=row, column=col, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=BRAND)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = _border


def _num(ws: Worksheet, row: int, col: int, value: float | None, fmt: str = MONEY, **font):
    """Pinta un numérico; None → celda vacía (nunca 0 falso)."""
    cell = ws.cell(row=row, column=col)
    cell.number_format = fmt
    cell.alignment = Alignment(horizontal="right")
    if value is not None:
        cell.value = value
    if font:
        cell.font = Font(**font)
    return cell


# ===== Hoja maestra: definición de columnas (orden del Excel original) =====
# (grupo, encabezado, atributo del vuelo, formato numérico o None=texto)
_COLS: list[tuple[str, str, str | None, str | None]] = [
    ("", "CLAVE", "clave", None),
    ("", "FECHA", "fecha", None),
    ("", "RUTA", "ruta", None),
    ("", "ESTADO", "estado", None),
    ("VENTA", "HORAS\nCOBRADAS", "horas_cobradas", HORAS),
    ("VENTA", "TARIFA\nUSD/HR S/IVA", "tarifa_usd", MONEY),
    ("VENTA", "IVA\nUSD/HR", "iva_hr_usd", MONEY),
    ("VENTA", "TOTAL COBRADO\nAL CLIENTE USD *", "total_usd", MONEY),
    ("VENTA", "IVA TOTAL\nUSD", "iva_usd", MONEY),
    ("VENTA", "TIPO CAMBIO\nVENTA", "tc_venta", TC),
    ("VENTA", "TOTAL COBRADO\nAL CLIENTE MXN", "total_mxn", MONEY),
    ("VENTA", "IVA TOTAL\nMXN", "iva_mxn", MONEY),
    ("VENTA", "TOTAL S/IVA\nMXN", "subtotal_mxn", MONEY),
    ("TIEMPO / TACÓMETRO", "TIEMPO\nVUELO HR", "tiempo_vuelo", HORAS),
    ("TIEMPO / TACÓMETRO", "TACO\nINICIO", "taco_inicio", TACO),
    ("TIEMPO / TACÓMETRO", "TACO\nFINAL", "taco_fin", TACO),
    ("COSTOS DIRECTOS (MXN)", "GAS TOTAL\nMXN", "gas_mxn", MONEY),
    ("COSTOS DIRECTOS (MXN)", "GAS\nLITROS", "gas_litros", MONEY),
    ("COSTOS DIRECTOS (MXN)", "GAS PRECIO\nX LITRO", "gas_precio_litro", MONEY),
    ("COSTOS DIRECTOS (MXN)", "OPERACIONES", "op_mxn", MONEY),
    ("COSTOS DIRECTOS (MXN)", "PILOTO", "piloto_mxn", MONEY),
    ("COSTOS DIRECTOS (MXN)", "OTROS", "otros_mxn", MONEY),
    ("COSTOS DIRECTOS (MXN)", "PERMISO AFAC\n(PROVISIÓN)", "permiso_afac_mxn", MONEY),
    ("COSTOS DIRECTOS (MXN)", "COSTO TOTAL\nMXN", "costo_total_mxn", MONEY),
    ("COSTOS DIRECTOS (MXN)", "TIPO CAMBIO\nCOSTOS", "tc_costos", TC),
    ("INDICADORES USD / IVA", "COSTO TOTAL\nUSD", "costo_usd", MONEY),
    ("INDICADORES USD / IVA", "COSTO TOTAL\nUSD S/IVA", "costo_usd_siva", MONEY),
    ("INDICADORES USD / IVA", "IVA PAGADO\nUSD", "iva_pagado_usd", MONEY),
    ("INDICADORES USD / IVA", "IVA PAGADO\nMXN", "iva_pagado_mxn", MONEY),
    ("INDICADORES USD / IVA", "REMANENTE\nVENTA−COSTO MXN", "remanente_mxn", MONEY),
    ("INDICADORES USD / IVA", "DIF. IVA\nHACIENDA MXN", "dif_iva_mxn", MONEY),
    ("INDICADORES USD / IVA", "COMISIÓN\nVENDEDOR MXN", "comision_vendedor_mxn", MONEY),
    ("INDICADORES USD / IVA", "GANANCIA\nMXN", "ganancia_mxn", MONEY),
    ("INDICADORES USD / IVA", "GANANCIA\nUSD", "ganancia_usd", MONEY),
    ("INDICADORES USD / IVA", "COSTO X HORA\nUSD", "costo_hr_usd", MONEY),
    ("INDICADORES USD / IVA", "COSTO X HORA\nUSD S/IVA", "costo_hr_usd_siva", MONEY),
    ("STATUS DE COBROS", "STATUS", "status_cobro", None),
    ("STATUS DE COBROS", "COBRO 1\nFECHA", None, None),
    ("STATUS DE COBROS", "COBRO 1\nMXN", None, MONEY),
    ("STATUS DE COBROS", "COBRO 2\nFECHA", None, None),
    ("STATUS DE COBROS", "COBRO 2\nMXN", None, MONEY),
    ("STATUS DE COBROS", "COBRO 3\nFECHA", None, None),
    ("STATUS DE COBROS", "COBRO 3\nMXN", None, MONEY),
    ("STATUS DE COBROS", "COBRO 4\nFECHA", None, None),
    ("STATUS DE COBROS", "COBRO 4\nMXN", None, MONEY),
    ("STATUS DE COBROS", "TOTAL COBRADO\nMXN", "cobrado_mxn", MONEY),
    ("STATUS DE COBROS", "POR COBRAR\nMXN", "por_cobrar_mxn", MONEY),
    ("STATUS DE COBROS", "POR COBRAR\nUSD", "por_cobrar_usd", MONEY),
]
_COBRO1_COL = next(i for i, c in enumerate(_COLS, start=1) if c[1] == "COBRO 1\nFECHA")
# Celdas de costos con NOTA de desglose (comentario de Excel): al pasar el
# cursor se ve qué gastos componen el total ("Comida · Starbucks — $206.00").
_DETALLE_ATTR = {
    "gas_mxn": "gas_detalle",
    # GAS LITROS comparte la nota del gas: cada línea trae sus litros
    # ("Gas · ASA — $4,857.30 · 164 L"), así se ve cuántas cargas son.
    "gas_litros": "gas_detalle",
    "op_mxn": "op_detalle",
    "piloto_mxn": "piloto_detalle",
    "otros_mxn": "otros_detalle",
}
_GROUP_FILLS = {"VENTA": FILL_VENTA, "COSTOS DIRECTOS (MXN)": FILL_COSTOS,
                "STATUS DE COBROS": FILL_COBROS}
# Columnas de la fila TOTALES: atributo del vuelo → atributo de totales.
_TOTAL_MAP = {
    "horas_cobradas": "horas_cobradas", "tiempo_vuelo": "tiempo_vuelo",
    "total_mxn": "total_mxn", "iva_mxn": "iva_mxn", "subtotal_mxn": "subtotal_mxn",
    "gas_mxn": "gas_mxn", "gas_litros": "gas_litros", "op_mxn": "op_mxn",
    "piloto_mxn": "piloto_mxn", "otros_mxn": "otros_mxn",
    "permiso_afac_mxn": "permiso_afac_mxn", "costo_total_mxn": "costo_total_mxn",
    "remanente_mxn": "remanente_mxn", "dif_iva_mxn": "dif_iva_mxn",
    "comision_vendedor_mxn": "comision_vendedor_mxn", "ganancia_mxn": "ganancia_mxn",
    "ganancia_usd": "ganancia_usd", "cobrado_mxn": "cobrado_mxn",
    "por_cobrar_mxn": "por_cobrar_mxn", "por_cobrar_usd": "por_cobrar_usd",
    # Promedios (se pintan en su columna con etiqueta "prom." implícita):
    "tc_costos": "tc_promedio", "costo_hr_usd": "costo_hr_prom_usd",
}


def _cobros_a_4(cobros: list[BalanceAvionCobro]) -> list[BalanceAvionCobro]:
    """Hasta 4 parcialidades; si hay más, la 4ª agrega el resto (solo suma de
    columna para mostrar, no cálculo de negocio)."""
    if len(cobros) <= 4:
        return cobros
    resto = cobros[3:]
    montos = [c.monto_mxn for c in resto if c.monto_mxn is not None]
    agg = round(sum(montos), 2) if montos else None
    ultima = next((c.fecha for c in reversed(resto) if c.fecha), None)
    etiqueta = f"{_fecha(ultima)} (+{len(resto)})" if ultima else f"(+{len(resto)} cobros)"
    return [*cobros[:3], BalanceAvionCobro(fecha=etiqueta, monto_mxn=agg)]


def _hoja_maestra(ws: Worksheet, req: BalanceAvionRequest) -> None:
    ws.title = f"reporte horas {req.matricula}".strip()[:31]
    n = len(_COLS)

    # Encabezado compacto de 2 filas: grupo (merged) / columna.
    col = 1
    while col <= n:
        grupo = _COLS[col - 1][0]
        fin = col
        while fin < n and _COLS[fin][0] == grupo:
            fin += 1
        if grupo:
            c = ws.cell(row=1, column=col, value=grupo)
            c.font = Font(bold=True, color="FFFFFF", size=10)
            c.fill = PatternFill("solid", fgColor=NAVY)
            c.alignment = Alignment(horizontal="center", vertical="center")
            if fin > col:
                ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=fin)
        col = fin + 1
    for i, (grupo, header, _attr, _fmt) in enumerate(_COLS, start=1):
        c = ws.cell(row=2, column=i, value=header)
        c.font = Font(bold=True, color=NAVY, size=8)
        c.fill = PatternFill("solid", fgColor=_GROUP_FILLS.get(grupo, LIGHT))
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = _border
    ws.row_dimensions[2].height = 30

    # Datos: 1 fila por vuelo.
    row = 3
    for v in req.vuelos:
        cobros = _cobros_a_4(v.cobros)
        # Regla del cliente: lo cobrado nunca puede ser menor a lo volado.
        # Resalta HORAS COBRADAS en ámbar cuando el taco registró más horas
        # (solo señal visual; el pendiente del API explica el caso).
        horas_menores = (
            v.horas_cobradas is not None
            and v.tiempo_vuelo is not None
            and v.horas_cobradas > 0
            and v.tiempo_vuelo - v.horas_cobradas > 0.01
        )
        for i, (grupo, _header, attr, fmt) in enumerate(_COLS, start=1):
            fill = _GROUP_FILLS.get(grupo)
            if attr is not None:
                val = getattr(v, attr)
                if fmt is None:
                    cell = ws.cell(row=row, column=i, value=_fecha(val) if attr == "fecha" else val)
                    if attr == "estado" and val == "CANCELADO":
                        cell.font = Font(color=RED, size=9)
                    elif attr == "estado":
                        cell.font = Font(color=MUTED, size=9)
                else:
                    cell = _num(ws, row, i, val, fmt)
                if attr == "horas_cobradas" and horas_menores:
                    fill = AMBER
                # Nota con el desglose del total de la celda (gastos que la
                # componen), visible al pasar el cursor en Excel.
                det_attr = _DETALLE_ATTR.get(attr)
                if det_attr:
                    lineas = getattr(v, det_attr, None) or []
                    if lineas:
                        nota = Comment("\n".join(lineas), "VuelaTour")
                        nota.width = 340
                        nota.height = min(260, 40 + 16 * len(lineas))
                        cell.comment = nota
            else:  # parcialidades de cobro (pares fecha/monto desde _COBRO1_COL)
                idx = (i - _COBRO1_COL) // 2
                cobro = cobros[idx] if idx < len(cobros) else None
                if fmt is None:
                    cell = ws.cell(row=row, column=i, value=_fecha(cobro.fecha) if cobro else None)
                else:
                    cell = _num(ws, row, i, cobro.monto_mxn if cobro else None, fmt)
            cell.border = _border
            if fill:
                cell.fill = PatternFill("solid", fgColor=fill)
        row += 1

    # Fila TOTALES al final (sumas y promedios YA calculados por el API).
    t = req.totales
    ws.cell(row=row, column=1, value="TOTALES").font = Font(bold=True)
    for i, (_grupo, _header, attr, fmt) in enumerate(_COLS, start=1):
        cell = ws.cell(row=row, column=i)
        total_attr = _TOTAL_MAP.get(attr) if attr else None
        if total_attr is not None:
            cell = _num(ws, row, i, getattr(t, total_attr), fmt or MONEY, bold=True)
        cell.fill = PatternFill("solid", fgColor=LIGHT)
        cell.border = _border
        if cell.value is not None or i == 1:
            cell.font = Font(bold=True)
    row += 1

    # Otros ingresos del periodo (TUAs/extras/pernocta cobrados al cliente):
    # informativo — NO suman en las columnas de la fila por vuelo; se
    # trasladan al control GENERAL (regla del libro del cliente).
    otros = getattr(t, "otros_ingresos_usd", None)
    if otros is not None and otros != 0:
        ws.cell(
            row=row,
            column=1,
            value="OTROS INGRESOS del periodo (TUAs, extras, pernocta, transportes)"
            " — trasladar al control general:",
        ).font = Font(bold=True, size=9)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
        _num(ws, row, 9, otros, MONEY, bold=True)
        ws.cell(row=row, column=10, value="USD").font = Font(bold=True, size=9)
        row += 1
    row += 1

    # Notas al pie.
    for nota in (
        "* VENTA de cada fila = horas cobradas × tarifa (con o sin IVA según el "
        "vuelo). TUAs/extras/pernocta NO van en la fila: son OTROS INGRESOS y se "
        "registran en el control general.",
        "TIPO CAMBIO COSTOS y COSTO X HORA USD de la fila TOTALES son PROMEDIOS "
        "(los demás son sumas).",
    ):
        ws.cell(row=row, column=1, value=nota).font = Font(color=MUTED, size=9, italic=True)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=14)
        row += 1

    # Anchos + paneles congelados (bajo el encabezado, a la derecha de RUTA).
    anchos = {1: 16, 2: 12, 3: 24, 4: 12}
    for i in range(1, n + 1):
        ws.column_dimensions[get_column_letter(i)].width = anchos.get(i, 13)
    ws.freeze_panes = "D3"


def _hoja_gastos(ws: Worksheet, titulo: str, hoja: BalanceAvionHojaGastos,
                 req: BalanceAvionRequest) -> None:
    _title(ws, f"{titulo} — {req.matricula}".strip(" —"), 1, 5)
    periodo = f"Periodo: {req.periodo_desde or '—'} a {req.periodo_hasta or '—'}"
    ws.cell(row=2, column=1, value=periodo).font = Font(italic=True, size=10, color=MUTED)

    _header_row(ws, 4, ["TOTAL MXN", "TC PROMEDIO", "TOTAL USD", "HRS VOLADAS", "USD X HR"])
    _num(ws, 5, 1, hoja.total_mxn, MONEY, bold=True)
    _num(ws, 5, 2, req.totales.tc_promedio, TC)
    _num(ws, 5, 3, hoja.usd, MONEY, bold=True)
    _num(ws, 5, 4, req.totales.tiempo_vuelo, HORAS)
    _num(ws, 5, 5, hoja.usd_hr, MONEY)
    for c in range(1, 6):
        ws.cell(row=5, column=c).border = _border

    _header_row(ws, 7, ["FECHA", "DETALLE", "MONTO MXN", "MONEDA ORIGINAL", "MONTO ORIGINAL"])
    row = 8
    if hoja.filas:
        for f in hoja.filas:
            ws.cell(row=row, column=1, value=_fecha(f.fecha)).border = _border
            ws.cell(row=row, column=2, value=f.detalle).border = _border
            _num(ws, row, 3, f.monto_mxn).border = _border
            # Moneda/monto original solo cuando el gasto NO se capturó en MXN.
            ws.cell(row=row, column=4, value=f.moneda_original).border = _border
            _num(ws, row, 5, f.monto_original).border = _border
            row += 1
    else:
        ws.cell(row=row, column=1, value="Sin gastos registrados en el periodo.").font = Font(
            color=MUTED, italic=True
        )
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)

    for i, w in enumerate([14, 52, 15, 16, 15], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A8"


def _hoja_balance(ws: Worksheet, req: BalanceAvionRequest) -> None:
    b = req.balance
    _title(ws, f"Balance — {req.matricula}".strip(" —"), 1, 3)
    periodo = f"Periodo: {req.periodo_desde or '—'} a {req.periodo_hasta or '—'} (todo en USD)"
    ws.cell(row=2, column=1, value=periodo).font = Font(italic=True, size=10, color=MUTED)

    filas: list[tuple[str, float | None, bool]] = [
        ("UTILIDAD ANTES DE GASTOS USD", b.utilidad_antes_usd, False),
        ("(−) GASTOS INDIRECTOS USD", b.gastos_indirectos_usd, False),
        ("(−) OTROS GASTOS USD", b.otros_usd, False),
        ("(−) PERMISOS USD", b.permisos_usd, False),
        ("UTILIDAD DESPUÉS DE GASTOS USD", b.utilidad_despues_usd, True),
        ("(−) PENDIENTE DE PAGO (COBRANZA PENDIENTE) USD", b.por_cobrar_usd, False),
        ("UTILIDAD COBRADA USD", b.utilidad_cobrada_usd, True),
    ]
    row = 4
    for label, val, bold in filas:
        lc = ws.cell(row=row, column=1, value=label)
        lc.font = Font(bold=bold, color=NAVY if bold else MUTED)
        lc.border = _border
        color = None
        if label.startswith("UTILIDAD COBRADA") and val is not None:
            color = GREEN if val >= 0 else RED
        cell = _num(ws, row, 2, val, MONEY, bold=bold, color=color or ("000000" if bold else MUTED))
        cell.border = _border
        row += 1

    row += 1
    _title(ws, "Reparto a socios (utilidad COBRADA × % vigente)", row, 3, size=11)
    row += 1
    _header_row(ws, row, ["SOCIO", "PORCENTAJE", "MONTO USD"])
    row += 1
    if b.socios:
        for s in b.socios:
            ws.cell(row=row, column=1, value=s.nombre).border = _border
            pc = _num(ws, row, 2, s.porcentaje, '0.00"%"')
            pc.border = _border
            mc = _num(ws, row, 3, s.monto_usd, MONEY, bold=True)
            mc.border = _border
            row += 1
    else:
        ws.cell(
            row=row, column=1, value="Sin socios configurados para este avión (ver pendientes)."
        ).font = Font(color=RED, italic=True)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)

    for i, w in enumerate([46, 14, 16], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _hoja_pendientes(ws: Worksheet, req: BalanceAvionRequest) -> None:
    _title(ws, f"Pendientes de captura — {req.matricula}".strip(" —"), 1, 2)
    ws.cell(
        row=2, column=1,
        value="Genera de nuevo el balance después de capturar; esta hoja debe quedar vacía.",
    ).font = Font(italic=True, size=10, color=MUTED)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=2)

    row = 4
    if req.pendientes:
        for i, texto in enumerate(req.pendientes, start=1):
            nc = ws.cell(row=row, column=1, value=i)
            nc.font = Font(bold=True, color=RED)
            nc.alignment = Alignment(horizontal="right", vertical="top")
            ws.cell(row=row, column=2, value=texto).alignment = Alignment(
                wrap_text=True, vertical="top"
            )
            row += 1
    else:
        c = ws.cell(
            row=row, column=2, value="Sin pendientes — la captura del periodo está completa."
        )
        c.font = Font(bold=True, color=GREEN)

    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 95
    ws.freeze_panes = "A4"


def render_balance_avion_xlsx(req: BalanceAvionRequest) -> bytes:
    wb = Workbook()
    _hoja_maestra(wb.active, req)
    _hoja_gastos(wb.create_sheet("gastos indirectos"), "Gastos indirectos",
                 req.gastos_indirectos, req)
    _hoja_gastos(wb.create_sheet("otros gastos"), "Otros gastos", req.otros_gastos, req)
    _hoja_gastos(wb.create_sheet("permisos"), "Permisos", req.permisos, req)
    _hoja_balance(wb.create_sheet("balance"), req)
    _hoja_pendientes(wb.create_sheet("pendientes de captura"), req)

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()
