"""Reporte mensual por avión en Excel (doc 5.10 / etapa 7).

Reutiliza el mismo payload del reparto de utilidades (RepartoPdfRequest) y arma
un libro con: una hoja resumen por avión con las 6 secciones (ingresos, gastos
directos/indirectos, permisos, otros, reserva overhaul, saldo) y el reparto por
socio de cada avión.
"""

from io import BytesIO

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from app.schemas.reparto import (
    RepartoCierre,
    RepartoOtrosIngresosDesglose,
    RepartoPdfRequest,
    RepartoVueloLinea,
)
from app.services import cierre_mes
from app.services._formato import _tc_txt
from app.services.reparto_comisiones import (
    ETIQUETA_COLUMNA,
    comisiones_al_avion,
    frase_regla,
    regla,
    solo_gasto_real,
)

BRAND = "0F4C81"
LIGHT = "EEF2F7"
WHITE = "FFFFFF"
MONEY = '"$"#,##0.00'

_thin = Side(style="thin", color="D5DBE3")
_border = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)


def _pct(factor: float | None) -> str:
    """0.5 → '50 %' (hasta 2 decimales, sin ceros de más)."""
    return "—" if factor is None else f"{round(factor * 100, 2):g} %"


def _usd(value: float) -> str:
    return f"${value:,.2f}"


def _desglose_txt(d: RepartoOtrosIngresosDesglose | None, moneda: str = "") -> str | None:
    """'TUAs $100.00 · comisión del vendedor $1,000.00 · IVA $176.00' con la
    composición COTIZADA de otros ingresos VuelaTour; None si no hay nada.
    `moneda` (reparto EN PESOS): « USD» tras cada monto; sin él, el de
    siempre."""
    if d is None:
        return None
    partes = [
        f"{label} {_usd(val)}{moneda}"
        for label, val in (
            ("TUAs", d.tuas_usd),
            ("extras", d.extras_usd),
            ("pernocta", d.pernocta_usd),
            ("comisión del vendedor", d.comision_usd),
            ("IVA", d.iva_usd),
        )
        if val
    ]
    return " · ".join(partes) or None


def _nota_tc_oficial(avion) -> str | None:
    """Nota del TC oficial de respaldo del avión (29-ago-2026), o None."""
    partes = []
    g = getattr(avion, "gastos_tc_oficial", None)
    if g is not None and g.count > 0:
        partes.append(f"{g.count} gasto(s) MXN sin TC capturado (${g.monto_mxn:,.2f} MXN)")
    c = getattr(avion, "cobros_tc_oficial_count", 0) or 0
    if c > 0:
        partes.append(f"{c} vuelo(s) con cobros MXN sin TC")
    if not partes:
        return None
    return (
        " y ".join(partes)
        + " se convirtieron con el TC oficial de referencia del día "
        "(open.er-api / BCE); ya están dentro de las cifras."
    )


def _nota_tc_oficial_global(tc) -> str | None:
    """Nota del periodo (resumen): vuelos/gastos convertidos con TC oficial."""
    if tc is None:
        return None
    gastos = tc.gastos.count if tc.gastos is not None else 0
    if tc.vuelos <= 0 and gastos <= 0:
        return None
    partes = []
    if gastos > 0:
        partes.append(f"{gastos} gasto(s) MXN (${tc.gastos.monto_mxn:,.2f} MXN)")
    if tc.vuelos > 0:
        partes.append(f"{tc.vuelos} vuelo(s) con cobros MXN")
    fuentes = f" Fuente(s): {' / '.join(tc.fuentes)}." if tc.fuentes else ""
    return (
        "Tipo de cambio de referencia: "
        + " y ".join(partes)
        + " sin TC capturado se convirtieron con el TC oficial del día de la "
        "cotización o del gasto (open.er-api / BCE); ya están dentro de las cifras."
        + fuentes
    )


def _folio_vuelo(v: RepartoVueloLinea) -> str:
    """'#105 · 50 %' cuando el vuelo fue MULTI-AVIÓN (regla B, 28-ago-2026:
    la venta del avión se repartió por tramo); '#105' si no. Con la marca del
    cierre de mes (16-oct-2026, texto del API) al final: '#105 · Viene de
    septiembre 2026'."""
    texto = f"#{v.folio}" if v.folio is not None else "—"
    if v.participacion is not None and v.participacion < 0.9999:
        texto += f" · {_pct(v.participacion)}"
    if v.cierre_marca:
        texto += f" · {v.cierre_marca}"
    return texto


# ===== CIERRE DE MES (16-oct-2026, API 0.0.88, invariante 89 del API) =====
# Sección «Cierre del mes» (después del detalle de vuelos): los vuelos que
# PASAN al mes siguiente (pendientes de cobro: no se reparten), los que
# VIENEN de un mes anterior (se reparten aquí) y, en un mes abierto, los que
# PASARÍAN al cerrar. Montos y leyendas YA calculados por el API, como
# valores (no entran a ningún total). Sin `cierre` no existe: libro de
# siempre.
ENCABEZADOS_CIERRE = [
    "Avión", "Vuelo", "Cliente", "Venta USD", "Cobrado USD", "Por cobrar USD", "Mes", "Leyenda",
]  # fmt: skip
# Con el `motivo` del API (caso #351, 16-oct-2026): lo cobrado y lo por
# cobrar son los del CORTE de cada lista; sin «al cierre» se leían como de
# hoy («aparece que aún no se cobran pero sí están cobrados»).
ENCABEZADOS_CIERRE_AL_CIERRE = [
    "Avión", "Vuelo", "Cliente", "Venta USD", "Cobrado USD al cierre",
    "Por cobrar USD al cierre", "Mes", "Leyenda",
]  # fmt: skip
_FILL_CIERRE = {"diferidos_salen": cierre_mes.FILL_SALE,
                "diferidos_entran": cierre_mes.FILL_VIENE,
                "pendientes_que_pasarian": cierre_mes.FILL_SALE}  # fmt: skip


def _hay_listas_cierre(cierre: RepartoCierre | None) -> bool:
    return cierre is not None and any(
        getattr(cierre, attr) for attr, _p, _m in cierre_mes.LISTAS_REPARTO
    )


def _seccion_cierre(ws, r: int, cierre: RepartoCierre) -> int:
    """Pinta la sección desde la fila `r` (ya con el espacio previo) y
    devuelve la fila siguiente a la última escrita."""
    titulo = "Cierre del mes"
    sello = cierre_mes.texto_estado_cierre(cierre)
    if sello:
        titulo += f" — {sello}"
    _title(ws, titulo, r, len(ENCABEZADOS_CIERRE), size=12)
    r += 1
    con_motivo = any(
        cierre_mes.hay_motivo(getattr(cierre, attr)) for attr, _p, _m in cierre_mes.LISTAS_REPARTO
    )
    _header_row(ws, r, ENCABEZADOS_CIERRE_AL_CIERRE if con_motivo else ENCABEZADOS_CIERRE)
    r += 1
    for attr, plantilla, mes_attr in cierre_mes.LISTAS_REPARTO:
        items = getattr(cierre, attr)
        if not items:
            continue
        sub = ws.cell(
            row=r, column=1, value=cierre_mes.titulo_lista_reparto(plantilla, items, mes_attr)
        )
        sub.font = Font(bold=True, size=10)
        sub.fill = PatternFill("solid", fgColor=_FILL_CIERRE[attr])
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=len(ENCABEZADOS_CIERRE))
        r += 1
        for v in items:
            mes = getattr(v, mes_attr)
            valores = [
                v.matricula, f"#{v.folio}" if v.folio else "—", v.cliente,
                v.venta_usd, v.cobrado_usd, v.por_cobrar_usd,
                cierre_mes.etiqueta_mes(mes), v.leyenda,
            ]  # fmt: skip
            for col, val in enumerate(valores, start=1):
                c = ws.cell(row=r, column=col, value=val)
                c.border = _border
                if col in (4, 5, 6):
                    c.number_format = MONEY
            r += 1
    return r


# Grupo del desglose de gastos (mismas etiquetas que el panel y el PDF).
_GRUPO_LABEL = {
    "DIRECTO": "Directo",
    "INDIRECTO": "Indirecto",
    "PERMISO": "Permiso",
    "FIJO": "Fijo (manual)",
    "EXCLUIDO": "Excluido",
}


# Nota bajo la tabla. La cláusula de en medio depende de la regla de
# comisiones (6-oct-2026): sin comisiones absorbidas, la de siempre.
_NOTA_INICIO = (
    "Venta del avión cobrada = tiempo de vuelo + ajuste + IVA proporcional "
    "(en vuelos multi-avión, la parte de cada matrícula en partes iguales por "
    "tramo vendido; los ferries/tramos operativos no reparten). "
    "Otros ingr. VuelaTour = TUAs/extras/pernocta/comisión del vendedor cobrados "
    "(con su IVA): ingreso de VuelaTour, fuera del saldo y del reparto"
)
_NOTA_COMISION_VENDEDOR = (
    "; el pago de la comisión al vendedor sale de VuelaTour, no del avión (ver "
    "'otros movimientos' del Balance general)."
)
_NOTA_FIN = (
    " Pendiente de cobro = parte del avión (se reparte al cobrar); Deuda total "
    "del cliente = lo que debe COMPLETO (con TUAs/extras/pernocta/comisión), "
    "solo cuando es mayor a la parte del avión."
)


# Otros ingresos ASIGNADOS a un avión (13-oct-2026, API 0.0.75, invariante
# «INGRESO ASIGNADO A UN AVIÓN» del API): SUMAN al saldo (ya dentro de
# `saldo_usd`). Columna 15 y su nota SOLO con dato; sin él, las 14 columnas
# de siempre. Textos EXACTOS del contrato (es-MX).
ENCABEZADO_OTROS_INGRESOS_AVION = "Otros ingresos del avión (suman al saldo)"
NOTA_OTROS_INGRESOS_AVION = (
    "Otros ingresos del avión = reembolsos, ventas e intereses registrados en "
    "Ingresos con ese avión, netos de comisión bancaria; SUMAN al saldo (Saldo "
    "disponible = venta cobrada − comisiones − gastos − fijos − reserva + otros "
    "ingresos del avión)."
)


def _nota_tabla(req: RepartoPdfRequest) -> str:
    """Nota bajo la tabla. Con comisiones a cargo del avión (regla de
    septiembre 2026) ya no dice que el pago de la comisión «no es del avión»:
    explica la columna Comisiones (banco + vendedor)."""
    if not comisiones_al_avion(req):
        return _NOTA_INICIO + _NOTA_COMISION_VENDEDOR + _NOTA_FIN
    return (
        f"{_NOTA_INICIO}. {ETIQUETA_COLUMNA}: {frase_regla(req, 'con')}; restan "
        f"antes del saldo (ver 'otros movimientos' del Balance general).{_NOTA_FIN}"
    )


def _nota_comision_por_avion(req: RepartoPdfRequest, por_avion: list[str]) -> str:
    """«Comisión del vendedor cotizada por avión (…): XB-TST $80.00.» Con la
    bandera del API 0.0.74 (`solo_gasto_real`, dos categorías) dice que el
    avión absorbe su pago si se captura como «Comisión del vendedor (avión)»
    —en cualquier fecha: sin nombrar la regla— y que el de «(VuelaTour)» lo
    paga VuelaTour; sin ella, el texto de siempre."""
    if not comisiones_al_avion(req):
        aclaracion = "ingreso de VuelaTour, su pago no es costo del avión"
    elif solo_gasto_real(req):
        aclaracion = (
            "ingreso de VuelaTour; su pago, si se captura como «Comisión del "
            "vendedor (avión)», lo absorbe el avión en Comisiones (el de "
            "«(VuelaTour)» lo paga VuelaTour)"
        )
    else:
        aclaracion = (
            f"ingreso de VuelaTour; con la {regla(req)}, por fecha del vuelo, "
            "su provisión la absorbe el avión en Comisiones"
        )
    return (
        f"Comisión del vendedor cotizada por avión ({aclaracion}): " + " · ".join(por_avion) + "."
    )


def _title(ws, text: str, row: int, span: int, size: int = 14):
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = Font(bold=True, size=size, color=BRAND)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)


def _header_row(ws, row: int, headers: list[str], start: int = 1):
    for col, h in enumerate(headers, start=start):
        c = ws.cell(row=row, column=col, value=h)
        c.font = Font(bold=True, color=WHITE)
        c.fill = PatternFill("solid", fgColor=BRAND)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = _border


def render_reparto_xlsx(req: RepartoPdfRequest) -> bytes:
    # Reparto EN PESOS (18-oct-2026, API 0.0.93): con la bandera, el libro en
    # pesos (dólares informativos); sin ella, el de siempre, byte a byte.
    if req.reparto_en_pesos:
        return _render_reparto_xlsx_pesos(req)
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumen por avión"

    # Columna 15 «Otros ingresos del avión» (13-oct-2026): solo si algún
    # avión trae otros ingresos asignados. NO entra a `money_cols` (D..L
    # intactas); el saldo del API ya los incluye.
    con_oi = any(a.otros_ingresos_avion_usd for a in req.aviones)
    n_cols = 15 if con_oi else 14
    _title(ws, "VuelaTour — Reporte mensual por avión", 1, n_cols, size=16)
    # Sello del cierre de mes (16-oct-2026) al final del subtítulo; sin
    # `cierre` (o mes abierto), el de siempre.
    sello_cierre = cierre_mes.texto_estado_cierre(req.cierre)
    sub = ws.cell(row=2, column=1, value=f"Periodo: {req.periodo_desde} a {req.periodo_hasta}  ·  Generado: {req.generado}"
                  + (f"  ·  {sello_cierre}" if sello_cierre else ""))  # fmt: skip
    sub.font = Font(italic=True, size=10, color="5B6470")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols)

    headers = [
        "Matrícula", "Modelo", "Horas voladas", "Venta del avión cobrada", ETIQUETA_COLUMNA,
        "Pendiente de cobro (parte avión)", "Gastos directos", "Gastos indirectos", "Permisos",
        "Fijos (prorrateo + reparto)", "Reserva overhaul", "Saldo disponible",
        # Informativas (regla 28-ago), FUERA de las columnas que suman el saldo:
        # TUAs/extras/pernocta cobrados = ingreso de VuelaTour; y la deuda
        # COMPLETA del cliente cuando es mayor a la parte del avión.
        "Otros ingr. VuelaTour (informativo)",
        "Deuda total del cliente (informativo)",
        *([ENCABEZADO_OTROS_INGRESOS_AVION] if con_oi else []),
    ]
    # Columna E con las comisiones que absorbe el avión (regla de septiembre
    # 2026: banco + lo del vendedor —su provisión hasta el API 0.0.73, el
    # pago «Comisión del vendedor (avión)» desde el 0.0.74—). Desde el
    # 13-oct-2026 se llama así SIEMPRE: sin comisiones absorbidas la columna
    # va oculta (abajo), pero quien la mostraba leía «Comisiones venta», el
    # rótulo de la comisión del vendedor de antes del 28-ago-2026.
    hrow = 4
    _header_row(ws, hrow, headers)

    money_cols = list(range(4, 13))  # columnas D..L son montos (C = horas)
    totals = {c: 0.0 for c in money_cols}
    info_col = 13  # M: otros ingresos VuelaTour (total propio, no suma al saldo)
    total_info = 0.0
    # N: deuda total del cliente — solo cuando supera la parte del avión
    # (pendiente_bruto_usd > pendiente_cobro_usd); en gris, sin total.
    deuda_col = 14
    # O: otros ingresos del avión (13-oct-2026) — SUMAN al saldo; total propio.
    oi_col = 15
    total_oi = 0.0

    r = hrow + 1
    for a in req.aviones:
        valores = [
            a.matricula, a.modelo, a.horas_voladas_hr, a.ingresos_cobrado_usd,
            a.comisiones_venta_usd, a.pendiente_cobro_usd,
            a.gastos_directos_usd, a.gastos_indirectos_usd, a.permisos_usd,
            a.otros_usd, a.reserva_overhaul_usd, a.saldo_usd,
            a.otros_ingresos_vuelatour_usd,
        ]
        for col, v in enumerate(valores, start=1):
            c = ws.cell(row=r, column=col, value=v)
            c.border = _border
            if col == 3:
                c.number_format = "0.0"
            if col in money_cols:
                c.number_format = MONEY
                totals[col] += float(v or 0)
            if col == info_col:
                c.number_format = MONEY
                c.font = Font(italic=True, color="5B6470")
                total_info += float(v or 0)
        pendiente = a.pendiente_cobro_usd or 0.0
        bruto = a.pendiente_bruto_usd or 0.0
        dc = ws.cell(
            row=r, column=deuda_col,
            value=bruto if bruto > pendiente + 0.005 else None,
        )
        dc.number_format = MONEY
        dc.font = Font(italic=True, color="5B6470")
        dc.border = _border
        if con_oi:
            oc = ws.cell(row=r, column=oi_col, value=a.otros_ingresos_avion_usd)
            oc.number_format = MONEY
            oc.border = _border
            total_oi += float(a.otros_ingresos_avion_usd or 0)
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
    # Total propio de la columna informativa (no entra al saldo).
    ci = ws.cell(row=r, column=info_col, value=total_info)
    ci.font = Font(bold=True, italic=True, color="5B6470")
    ci.fill = PatternFill("solid", fgColor=LIGHT)
    ci.number_format = MONEY
    ci.border = _border
    cd = ws.cell(row=r, column=deuda_col)
    cd.fill = PatternFill("solid", fgColor=LIGHT)
    cd.border = _border
    if con_oi:
        co = ws.cell(row=r, column=oi_col, value=total_oi)
        co.font = Font(bold=True)
        co.fill = PatternFill("solid", fgColor=LIGHT)
        co.number_format = MONEY
        co.border = _border
    r += 1
    # Columna E (comisiones que restan al avión): con la regla A del
    # 28-ago-2026 el API mandó 0 y desde la regla de septiembre 2026 trae lo
    # que absorbe cada avión (banco + lo del vendedor: provisión hasta el API
    # 0.0.73, pago «(avión)» desde el 0.0.74). Se conserva por índice
    # (money_cols, widths) y se OCULTA cuando nadie trae monto.
    if not any(a.comisiones_venta_usd for a in req.aviones):
        ws.column_dimensions[get_column_letter(5)].hidden = True
    nota = ws.cell(row=r, column=1, value=_nota_tabla(req))
    nota.font = Font(italic=True, size=9, color="5B6470")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=n_cols)
    if con_oi:
        r += 1
        no = ws.cell(row=r, column=1, value=NOTA_OTROS_INGRESOS_AVION)
        no.font = Font(italic=True, size=9, color="5B6470")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=n_cols)
    # Composición cotizada de "Otros ingr. VuelaTour" (regla A: incluye la
    # comisión del vendedor) — global y por avión, solo si el API la manda.
    desglose_txt = _desglose_txt(req.otros_ingresos_vuelatour_desglose)
    if desglose_txt:
        r += 1
        nd = ws.cell(
            row=r,
            column=1,
            value="Otros ingr. VuelaTour del periodo (cotizado, no se reparte): "
            f"{desglose_txt}.",
        )
        nd.font = Font(italic=True, size=9, color="5B6470")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=n_cols)
    por_avion = [
        f"{a.matricula} {_usd(a.otros_ingresos_vuelatour_desglose.comision_usd)}"
        for a in req.aviones
        if a.otros_ingresos_vuelatour_desglose is not None
        and a.otros_ingresos_vuelatour_desglose.comision_usd
    ]
    if por_avion:
        r += 1
        nc = ws.cell(row=r, column=1, value=_nota_comision_por_avion(req, por_avion))
        nc.font = Font(italic=True, size=9, color="5B6470")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=n_cols)
    # Nota global del TC oficial de respaldo (29-ago-2026): misma redacción
    # que el PDF.
    nota_global = _nota_tc_oficial_global(req.tc_oficial)
    if nota_global:
        r += 1
        nt = ws.cell(row=r, column=1, value=nota_global)
        nt.font = Font(italic=True, size=9, color="5B6470")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=n_cols)
    # Cierre de mes (16-oct-2026): los números de hoy ya no son los del
    # cierre (el API lo detecta) ⇒ aviso en rojo, como en el PDF.
    if req.cierre is not None and req.cierre.coincide_con_snapshot is False:
        r += 1
        na = ws.cell(row=r, column=1, value=f"AVISO: {cierre_mes.AVISO_SNAPSHOT}")
        na.font = Font(bold=True, size=9, color="B91C1C")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=n_cols)

    # Anchos de columna.
    widths = [12, 16, 13, 16, 16, 17, 15, 16, 12, 16, 16, 16, 18, 18]
    if con_oi:
        widths.append(18)
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
        # Vigencias traslapadas/incompletas: mismo aviso que el PDF y el badge
        # de la web — el Excel dispersable no puede salir doble en silencio.
        if a.reparto and round(a.reparto_porcentaje_total, 2) != 100:
            aviso = ws.cell(
                row=r,
                column=1,
                value=(
                    f"AVISO {a.matricula}: los porcentajes suman "
                    f"{a.reparto_porcentaje_total:g}% (no 100%) — revisar vigencias de socios."
                ),
            )
            aviso.font = Font(color="B45309", italic=True, size=9)
            r += 1
        # Nota informativa (29-ago-2026): montos MXN sin TC capturado que SÍ
        # entraron, convertidos con el TC oficial de referencia del día.
        nota_tc = _nota_tc_oficial(a)
        if nota_tc:
            nc = ws.cell(row=r, column=1, value=f"NOTA {a.matricula}: {nota_tc}")
            nc.font = Font(italic=True, size=9, color="5B6470")
            r += 1

    # Detalle de vuelos (solo si el API lo manda): el folio lleva ' · 50 %'
    # cuando el vuelo fue MULTI-AVIÓN (regla B, 28-ago-2026: la venta del
    # avión se repartió por tramo entre las matrículas).
    if any(a.vuelos for a in req.aviones):
        r += 2
        _title(ws, "Detalle de vuelos (venta del avión cobrada)", r, 6, size=12)
        r += 1
        _header_row(ws, r, ["Avión", "Vuelo", "Cliente", "Cobrado avión USD",
                            "Pendiente avión USD", "Tramos del avión"])
        r += 1
        for a in req.aviones:
            for v in a.vuelos:
                ws.cell(row=r, column=1, value=a.matricula).border = _border
                ws.cell(row=r, column=2, value=_folio_vuelo(v)).border = _border
                ws.cell(row=r, column=3, value=v.cliente).border = _border
                for col, val in ((4, v.cobrado_usd), (5, v.pendiente_usd)):
                    mc = ws.cell(row=r, column=col, value=val)
                    mc.number_format = MONEY
                    mc.border = _border
                ws.cell(row=r, column=6, value=v.tramos_avion).border = _border
                r += 1

    if _hay_listas_cierre(req.cierre):
        r = _seccion_cierre(ws, r + 2, req.cierre)

    # Gastos por categoría (solo si el API lo manda, 2-sep-2026): la categoría
    # se imprime con su etiqueta amable (homologada con panel/app/API) o con
    # el código del enum tal cual si no viene. Sección propia al final: no
    # mueve columnas ni índices de las tablas de arriba.
    if any(a.gastos_por_categoria for a in req.aviones):
        r += 2
        _title(ws, "Gastos por categoría", r, 5, size=12)
        r += 1
        _header_row(ws, r, ["Avión", "Categoría de gasto", "Grupo", "# gastos", "USD"])
        r += 1
        for a in req.aviones:
            for g in a.gastos_por_categoria:
                ws.cell(row=r, column=1, value=a.matricula).border = _border
                ws.cell(row=r, column=2, value=g.etiqueta or g.categoria).border = _border
                grupo = _GRUPO_LABEL.get(g.grupo or "", g.grupo or "")
                ws.cell(row=r, column=3, value=grupo).border = _border
                ws.cell(row=r, column=4, value=g.count).border = _border
                mc = ws.cell(row=r, column=5, value=g.usd)
                mc.number_format = MONEY
                mc.border = _border
                r += 1

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ===== REPARTO EN PESOS (18-oct-2026, API 0.0.93) =====
# Pedido del cliente: «… todo en pesos porque casi siempre se les transfiere
# a los socios en pesos mexicanos, y que los dólares solo sean informativos».
# Con la bandera `reparto_en_pesos` la tabla principal, el reparto por socio y
# el detalle de vuelos van en PESOS (los `*_mxn` del API) y, debajo, la tabla
# de siempre en dólares como INFORMATIVA (gris, cursiva). Aquí no se convierte
# nada: un `*_mxn` que no llega se escribe «—»; los totales solo re-suman la
# columna (con un «—» en la columna, el total también es «—»). Sin la
# bandera, nada de esto se usa.

LEYENDA_PESOS = "Cifras en pesos (MXN); dólares solo informativos"
GRIS_INFO = "94A3B8"
_FONT_INFO = Font(italic=True, size=9, color=GRIS_INFO)
_FILL_INFO = PatternFill("solid", fgColor="F1F5F9")
TITULO_USD_INFORMATIVO = "Las mismas cifras en dólares (informativo)"
ENCABEZADO_OTROS_INGRESOS_AVION_PESOS = "Otros ingresos del avión (MXN, suman al saldo)"
SIN_PESOS = "—"


def encabezados_pesos(con_oi: bool) -> list[str]:
    """Encabezados de la tabla principal en pesos (mismas 14/15 columnas y
    mismo orden que la de dólares)."""
    return [
        "Matrícula",
        "Modelo",
        "Horas voladas",
        "Venta del avión cobrada (MXN)",
        f"{ETIQUETA_COLUMNA} (MXN)",
        "Pendiente de cobro (parte avión) (MXN)",
        "Gastos directos (MXN)",
        "Gastos indirectos (MXN)",
        "Permisos (MXN)",
        "Fijos (prorrateo + reparto) (MXN)",
        "Reserva overhaul (MXN)",
        "Saldo disponible (MXN)",
        "Otros ingr. VuelaTour (USD, informativo)",
        "Deuda total del cliente (USD, informativo)",
        *([ENCABEZADO_OTROS_INGRESOS_AVION_PESOS] if con_oi else []),
    ]


def aviso_pesos_sin_tc(matricula: str, usd: float) -> str:
    return (
        f"AVISO {matricula}: {_usd(usd)} USD no se pudieron pasar a pesos (sin "
        "tipo de cambio): el saldo en pesos está incompleto."
    )


def _celda_mxn(ws, row: int, col: int, valor: float | None, *, bold: bool = False):
    """Celda principal en pesos: el número del API o «—» (jamás un 0)."""
    if valor is None:
        c = ws.cell(row=row, column=col, value=SIN_PESOS)
        c.alignment = Alignment(horizontal="right")
    else:
        c = ws.cell(row=row, column=col, value=valor)
        c.number_format = MONEY
    if bold:
        c.font = Font(bold=True)
    c.border = _border
    return c


def _celda_info(ws, row: int, col: int, valor: float | None, *, bold: bool = False):
    """Celda en dólares INFORMATIVA: gris, cursiva, letra 9."""
    c = ws.cell(row=row, column=col, value=valor)
    c.number_format = MONEY
    c.font = Font(italic=True, size=9, color=GRIS_INFO, bold=bold)
    c.border = _border
    return c


def _header_info(ws, row: int, headers: list[str], start: int = 1) -> None:
    """Encabezados de las columnas informativas en dólares: grises, sin el
    azul de los principales."""
    for col, h in enumerate(headers, start=start):
        c = ws.cell(row=row, column=col, value=h)
        c.font = Font(bold=True, italic=True, size=9, color="64748B")
        c.fill = _FILL_INFO
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = _border


def _suma_columna(valores: list[float | None]) -> float | None:
    if any(v is None for v in valores):
        return None
    return round(sum(v for v in valores if v is not None), 2)


def _nota(ws, r: int, texto: str, n_cols: int, *, color: str = "5B6470", bold: bool = False):
    c = ws.cell(row=r, column=1, value=texto)
    c.font = Font(italic=not bold, bold=bold, size=9, color=color)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=n_cols)


def _render_reparto_xlsx_pesos(req: RepartoPdfRequest) -> bytes:
    """El Excel del reparto EN PESOS (bandera `reparto_en_pesos`)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumen por avión"

    con_oi = any(a.otros_ingresos_avion_usd or a.otros_ingresos_avion_mxn for a in req.aviones)
    n_cols = 15 if con_oi else 14
    _title(ws, "VuelaTour — Reporte mensual por avión", 1, n_cols, size=16)
    sello_cierre = cierre_mes.texto_estado_cierre(req.cierre)
    sub = ws.cell(
        row=2,
        column=1,
        value=f"Periodo: {req.periodo_desde} a {req.periodo_hasta}  ·  Generado: {req.generado}"
        f"  ·  {LEYENDA_PESOS}" + (f"  ·  {sello_cierre}" if sello_cierre else ""),
    )
    sub.font = Font(italic=True, size=10, color="5B6470")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols)

    hrow = 4
    _header_row(ws, hrow, encabezados_pesos(con_oi))
    _header_info(ws, hrow, [encabezados_pesos(con_oi)[12], encabezados_pesos(con_oi)[13]], 13)

    # Columnas en pesos (D..L) y su atributo del API.
    attrs_mxn = {
        4: "ingresos_cobrado_mxn",
        5: "comisiones_venta_mxn",
        6: "pendiente_cobro_mxn",
        7: "gastos_directos_mxn",
        8: "gastos_indirectos_mxn",
        9: "permisos_mxn",
        10: "otros_mxn",
        11: "reserva_overhaul_mxn",
        12: "saldo_mxn",
    }
    col_vt, col_deuda, col_oi = 13, 14, 15
    r = hrow + 1
    for a in req.aviones:
        ws.cell(row=r, column=1, value=a.matricula).border = _border
        ws.cell(row=r, column=2, value=a.modelo).border = _border
        hc = ws.cell(row=r, column=3, value=a.horas_voladas_hr)
        hc.number_format = "0.0"
        hc.border = _border
        for col, attr in attrs_mxn.items():
            c = _celda_mxn(ws, r, col, getattr(a, attr), bold=col == 12)
            if col == 11 and a.reserva_tc:
                c.comment = Comment(
                    f"Reserva pasada a pesos al T.C. oficial del cierre: {_tc_txt(a.reserva_tc)}.",
                    "VuelaTour",
                )
        _celda_info(ws, r, col_vt, a.otros_ingresos_vuelatour_usd)
        pendiente = a.pendiente_cobro_usd or 0.0
        bruto = a.pendiente_bruto_usd or 0.0
        _celda_info(ws, r, col_deuda, bruto if bruto > pendiente + 0.005 else None)
        if con_oi:
            _celda_mxn(ws, r, col_oi, a.otros_ingresos_avion_mxn)
        r += 1

    tcell = ws.cell(row=r, column=1, value="TOTAL")
    tcell.font = Font(bold=True)
    tcell.fill = PatternFill("solid", fgColor=LIGHT)
    ws.cell(row=r, column=2).fill = PatternFill("solid", fgColor=LIGHT)
    totales_cols = dict(attrs_mxn)
    if con_oi:
        totales_cols[col_oi] = "otros_ingresos_avion_mxn"
    for col, attr in totales_cols.items():
        c = _celda_mxn(
            ws, r, col, _suma_columna([getattr(a, attr) for a in req.aviones]), bold=True
        )
        c.fill = PatternFill("solid", fgColor=LIGHT)
    ci = _celda_info(
        ws, r, col_vt, sum(a.otros_ingresos_vuelatour_usd for a in req.aviones), bold=True
    )
    ci.fill = PatternFill("solid", fgColor=LIGHT)
    cd = ws.cell(row=r, column=col_deuda)
    cd.fill = PatternFill("solid", fgColor=LIGHT)
    cd.border = _border
    r += 1
    if not any(a.comisiones_venta_usd or a.comisiones_venta_mxn for a in req.aviones):
        ws.column_dimensions[get_column_letter(5)].hidden = True

    # Avisos y notas (lo que no llegó a pesos va primero, en rojo).
    for a in req.aviones:
        if (a.pesos_sin_tc_usd or 0) > 0:
            _nota(ws, r, aviso_pesos_sin_tc(a.matricula, a.pesos_sin_tc_usd), n_cols,
                  color="B91C1C", bold=True)  # fmt: skip
            r += 1
    reservas = [f"{a.matricula} {_tc_txt(a.reserva_tc)}" for a in req.aviones if a.reserva_tc]
    if reservas:
        _nota(ws, r, "Reserva overhaul pasada a pesos al T.C. oficial del cierre: "
              + " · ".join(reservas) + ".", n_cols)  # fmt: skip
        r += 1
    _nota(ws, r, _nota_tabla(req), n_cols)
    if con_oi:
        r += 1
        _nota(ws, r, NOTA_OTROS_INGRESOS_AVION, n_cols)
    desglose_txt = _desglose_txt(req.otros_ingresos_vuelatour_desglose, " USD")
    if desglose_txt:
        r += 1
        _nota(ws, r, "Otros ingr. VuelaTour del periodo (cotizado, no se reparte): "
              f"{desglose_txt}.", n_cols)  # fmt: skip
    por_avion = [
        f"{a.matricula} {_usd(a.otros_ingresos_vuelatour_desglose.comision_usd)} USD"
        for a in req.aviones
        if a.otros_ingresos_vuelatour_desglose is not None
        and a.otros_ingresos_vuelatour_desglose.comision_usd
    ]
    if por_avion:
        r += 1
        _nota(ws, r, _nota_comision_por_avion(req, por_avion), n_cols)
    nota_global = _nota_tc_oficial_global(req.tc_oficial)
    if nota_global:
        r += 1
        _nota(ws, r, nota_global, n_cols)
    if req.cierre is not None and req.cierre.coincide_con_snapshot is False:
        r += 1
        _nota(ws, r, f"AVISO: {cierre_mes.AVISO_SNAPSHOT}", n_cols, color="B91C1C", bold=True)

    # La tabla de siempre en dólares, INFORMATIVA (gris, cursiva).
    r += 2
    tit = ws.cell(row=r, column=1, value=TITULO_USD_INFORMATIVO)
    tit.font = Font(bold=True, italic=True, size=11, color="64748B")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=n_cols)
    r += 1
    headers_usd = [
        "Matrícula", "Modelo", "Horas voladas", "Venta del avión cobrada", ETIQUETA_COLUMNA,
        "Pendiente de cobro (parte avión)", "Gastos directos", "Gastos indirectos", "Permisos",
        "Fijos (prorrateo + reparto)", "Reserva overhaul", "Saldo disponible",
        "Otros ingr. VuelaTour (informativo)",
        "Deuda total del cliente (informativo)",
        *([ENCABEZADO_OTROS_INGRESOS_AVION] if con_oi else []),
    ]  # fmt: skip
    _header_info(ws, r, headers_usd)
    r += 1
    attrs_usd = {
        4: "ingresos_cobrado_usd",
        5: "comisiones_venta_usd",
        6: "pendiente_cobro_usd",
        7: "gastos_directos_usd",
        8: "gastos_indirectos_usd",
        9: "permisos_usd",
        10: "otros_usd",
        11: "reserva_overhaul_usd",
        12: "saldo_usd",
        13: "otros_ingresos_vuelatour_usd",
    }
    if con_oi:
        attrs_usd[col_oi] = "otros_ingresos_avion_usd"
    for a in req.aviones:
        for col, val in ((1, a.matricula), (2, a.modelo)):
            c = ws.cell(row=r, column=col, value=val)
            c.font = _FONT_INFO
            c.border = _border
        hc = ws.cell(row=r, column=3, value=a.horas_voladas_hr)
        hc.number_format = "0.0"
        hc.font = _FONT_INFO
        hc.border = _border
        for col, attr in attrs_usd.items():
            _celda_info(ws, r, col, getattr(a, attr))
        pendiente = a.pendiente_cobro_usd or 0.0
        bruto = a.pendiente_bruto_usd or 0.0
        _celda_info(ws, r, col_deuda, bruto if bruto > pendiente + 0.005 else None)
        r += 1
    t2 = ws.cell(row=r, column=1, value="TOTAL USD")
    t2.font = Font(bold=True, italic=True, size=9, color="64748B")
    for col, attr in attrs_usd.items():
        _celda_info(ws, r, col, round(sum(float(getattr(a, attr) or 0) for a in req.aviones), 2),
                    bold=True)  # fmt: skip
    r += 1

    widths = [12, 16, 13, 18, 18, 19, 16, 17, 14, 18, 17, 18, 20, 20]
    if con_oi:
        widths.append(20)
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Reparto por socio: lo que se transfiere en pesos.
    r += 2
    _title(ws, "Reparto por socio (en pesos)", r, 5, size=12)
    r += 1
    _header_row(ws, r, ["Avión", "Socio", "Porcentaje", "Monto MXN"])
    _header_info(ws, r, ["Monto USD (informativo)"], 5)
    r += 1
    for a in req.aviones:
        for linea in a.reparto:
            ws.cell(row=r, column=1, value=a.matricula).border = _border
            ws.cell(row=r, column=2, value=linea.socio_nombre).border = _border
            pc = ws.cell(row=r, column=3, value=(linea.porcentaje or 0) / 100)
            pc.number_format = "0.00%"
            pc.border = _border
            _celda_mxn(ws, r, 4, linea.monto_mxn, bold=True)
            _celda_info(ws, r, 5, linea.monto_usd)
            r += 1
        if a.reparto and round(a.reparto_porcentaje_total, 2) != 100:
            aviso = ws.cell(
                row=r,
                column=1,
                value=(
                    f"AVISO {a.matricula}: los porcentajes suman "
                    f"{a.reparto_porcentaje_total:g}% (no 100%) — revisar vigencias de socios."
                ),
            )
            aviso.font = Font(color="B45309", italic=True, size=9)
            r += 1
        nota_tc = _nota_tc_oficial(a)
        if nota_tc:
            nc = ws.cell(row=r, column=1, value=f"NOTA {a.matricula}: {nota_tc}")
            nc.font = Font(italic=True, size=9, color="5B6470")
            r += 1

    if any(a.vuelos for a in req.aviones):
        r += 2
        _title(ws, "Detalle de vuelos (venta del avión cobrada)", r, 8, size=12)
        r += 1
        _header_row(ws, r, ["Avión", "Vuelo", "Cliente", "Cobrado avión MXN",
                            "Pendiente avión MXN"])  # fmt: skip
        _header_info(ws, r, ["Cobrado avión USD (informativo)",
                             "Pendiente avión USD (informativo)"], 6)  # fmt: skip
        _header_row(ws, r, ["Tramos del avión"], start=8)
        r += 1
        for a in req.aviones:
            for v in a.vuelos:
                ws.cell(row=r, column=1, value=a.matricula).border = _border
                ws.cell(row=r, column=2, value=_folio_vuelo(v)).border = _border
                ws.cell(row=r, column=3, value=v.cliente).border = _border
                _celda_mxn(ws, r, 4, v.cobrado_mxn)
                _celda_mxn(ws, r, 5, v.pendiente_mxn)
                _celda_info(ws, r, 6, v.cobrado_usd)
                _celda_info(ws, r, 7, v.pendiente_usd)
                ws.cell(row=r, column=8, value=v.tramos_avion).border = _border
                r += 1

    if _hay_listas_cierre(req.cierre):
        r = _seccion_cierre(ws, r + 2, req.cierre)

    if any(a.gastos_por_categoria for a in req.aviones):
        con_mxn = any(g.mxn is not None for a in req.aviones for g in a.gastos_por_categoria)
        r += 2
        _title(ws, "Gastos por categoría", r, 6 if con_mxn else 5, size=12)
        r += 1
        if con_mxn:
            _header_row(ws, r, ["Avión", "Categoría de gasto", "Grupo", "# gastos", "MXN"])
            _header_info(ws, r, ["USD (informativo)"], 6)
        else:
            _header_row(ws, r, ["Avión", "Categoría de gasto", "Grupo", "# gastos", "USD"])
        r += 1
        for a in req.aviones:
            for g in a.gastos_por_categoria:
                ws.cell(row=r, column=1, value=a.matricula).border = _border
                ws.cell(row=r, column=2, value=g.etiqueta or g.categoria).border = _border
                grupo = _GRUPO_LABEL.get(g.grupo or "", g.grupo or "")
                ws.cell(row=r, column=3, value=grupo).border = _border
                ws.cell(row=r, column=4, value=g.count).border = _border
                if con_mxn:
                    _celda_mxn(ws, r, 5, g.mxn)
                    _celda_info(ws, r, 6, g.usd)
                else:
                    mc = ws.cell(row=r, column=5, value=g.usd)
                    mc.number_format = MONEY
                    mc.border = _border
                r += 1

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()
