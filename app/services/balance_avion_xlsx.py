"""Balance mensual por avión en Excel (openpyxl) — réplica sistematizada del
control del equipo ("Balance N990GG.xlsx").

Ocho hojas en el orden del libro:
  1. reporte horas <MATRÍCULA> — maestro: 1 fila = 1 vuelo, TOTALES al final.
     VENTA AVIÓN = tiempo de vuelo + ajuste + IVA proporcional (regla
     28-ago-2026: TUAs/extras/pernocta y la comisión del vendedor son
     ingreso de VuelaTour y viven en 'otros movimientos' del Balance
     general). Desde la regla de septiembre 2026 (API 0.0.65, 6-oct-2026)
     la columna que era COMISIÓN VENDEDOR se llama COMISIONES y lleva lo
     que absorbe el avión: su parte de la comisión bancaria de los cobros +
     la comisión del vendedor como PROVISIÓN (nota por celda); GANANCIA =
     REMANENTE − COMISIONES. Vuelos MULTI-AVIÓN (regla B): la venta se
     reparte entre las matrículas por tramo; los gastos van al avión de su
     tramo. Costos SIN combustible; el TUA pagado es SOLO nota en
     OPERACIONES (no resta), y desde el 1-oct-2026 (API 0.0.47) la extensión
     y/o antelación de horario pagada al aeropuerto también
     (`extension_pagada_mxn`).
     STATUS DE COBROS trae los depósitos REALES (COBRO 1..4 y su Σ; desde el
     6-oct-2026 cada COBRO n MXN lleva la NOTA «cómo se cobró» y STATUS la
     nota resumen de todos los cobros — sin columnas nuevas —, y en los
     vuelos de la regla de septiembre 2026 cada COBRO es el NETO que entró a
     la cuenta) y, aparte, COBRADO AVIÓN = la parte prorrateada al avión
     (antes de comisiones); cierra con FACTURA VUELATOUR (folio de la
     factura del servicio que el API ya resolvió, 30-sep-2026 — texto tal
     cual, vacía sin dato).
  2. cobranza — estatus de cobro por vuelo: venta avión prorrateada, total
     cotización (c/extras), depósitos reales (netos con la regla de
     septiembre 2026), comisión y cuenta del banco.
  3. combustible — el gas del avión POR MES (litros y $/L), 26-ago-2026.
  4. Gastos Indirectos — UNA sola pestaña (pedido del cliente 2-sep-2026:
     'gastos indirectos' y 'otros gastos' se confundían) con las DOS listas
     que el API sigue mandando aparte: `gastos_indirectos` (gastos del avión
     sin vuelo) + `otros_gastos` (la parte de este avión de los gastos
     administrativos repartidos a mano — filas "reparto manual: $X de $Y",
     con tinte suave). Fusión de PRESENTACIÓN (_fusionar_hojas_gastos): el
     contrato del API no cambia ni un campo.
  5. refacciones (salidas de inventario — 29-ago-2026: antes dentro de
     indirectos; solo si el API la manda)  6. permisos (todo AFAC).
  7. balance — cascada (− combustible − GASTOS INDIRECTOS [una sola fila =
     gastos_indirectos_usd + otros_usd] − refacciones − permisos) + reparto
     real de socios. La cascada la calcula el API (utilidad_despues_usd ya
     resta ambas listas UNA vez); aquí solo se suman las dos celdas para
     mostrarlas juntas — restar otros además de indirectos contaría doble.
  8. pendientes de captura — lo que falta para que el libro quede completo.
El Balance GENERAL comparte estas funciones con otro juego de hojas (ver
render_balance_general_xlsx); ahí la hoja 'refacciones' fue sustituida por
'inventario' (tiendita, 30-ago-2026: resumen por ítem + detalle de salidas;
25-sep-2026: utilidad de la tienda también en DÓLARES, en columnas aparte;
y, con el API 0.0.36, costo = último precio de compra y pesos al T.C.
oficial del día — `regla_costo`, solo cambian los textos) — el libro
INDIVIDUAL conserva la suya.

El libro de flota sale en DOS variantes (6-oct-2026, API 0.0.64,
`BalanceGeneralRequest.variante`): «mensual» (o sin la llave) = la hoja de
vuelos de siempre (`_COLS_FLOTA`); «general» = la hoja «reporte horas FLOTA»
con el juego de columnas `_COLS_GENERAL` — COSTO TOTAL y el bloque COSTO POR
HORA (las 12 columnas de la hoja «utilidades» del cliente) en lugar de
operaciones / piloto / otros / AFAC e indicadores. Mismos números del API y
las demás hojas iguales (sus citas a la hoja de vuelos, por llave; la nota
de 'repartidos a aviones' nombra la columna de cada variante). Desde el
7-oct-2026 (API 0.0.66, pedido: «estas comisiones se están duplicando en la
general, ya que las tenemos en "Otros movimientos"») la hoja de vuelos de las
DOS variantes va SIN la columna COMISIONES y su GANANCIA / REMANENTE VENTA
MENOS COMPRA va antes de comisiones, como antes del 6-oct-2026
(`_antes_de_comisiones`): las que absorbe cada avión siguen en la columna
COMISIONES de su libro INDIVIDUAL y restan en la cascada de su bloque de la
hoja 'balance' (dos líneas visibles con el API 0.0.66). En 'otros
movimientos' (las dos variantes) las filas con PROVISIÓN van en amarillo con
su leyenda al pie.

Los montos vienen YA calculados del API (aquí solo se pintan; jamás se
recalcula dinero). None = celda vacía — nunca un 0 falso.

Fórmulas visibles (5-oct-2026, pedido del cliente con la hoja «reporte
horas»): toda celda que es aritmética de OTRAS celdas del libro (sumas de
TOTALES, VENTA/IVA MXN, PERMISO AFAC, COSTO TOTAL, indicadores USD/IVA,
remanente, ganancia, por cobrar, resúmenes de las hojas, cascada y reparto)
se escribe como FÓRMULA que reproduce el número del API, y ese número va en
la caché de la celda (`xlsx_formulas`). Si la fórmula no lo reproduce al
centavo, la celda se queda como valor.
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import datetime

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from app.schemas.reportes import (
    REGLA_COMISIONES_RESPALDO,
    REGLA_COSTO_ULTIMO_PRECIO,
    VARIANTE_BALANCE_GENERAL,
    BalanceAvionBalanceBloque,
    BalanceAvionCobro,
    BalanceAvionHojaCombustible,
    BalanceAvionHojaGastos,
    BalanceAvionRequest,
    BalanceAvionTotales,
    BalanceAvionVuelo,
    BalanceEmpresaBloque,
    BalanceGeneralRequest,
    BalanceGeneralResumenFila,
    BalanceHojaInventario,
    BalanceHojaOtrosMovimientos,
    BalanceOtroMovimientoFila,
    TcHoy,
)
from app.services import xlsx_formulas
from app.services._formato import _tc_txt
from app.services.caja_chica_xlsx import _texto_literal
from app.services.tabla_xlsx import sheet_title

BRAND = "0F4C81"
NAVY = "102A43"
MUTED = "627D98"
LIGHT = "EEF2F7"
GREEN = "15803D"
RED = "DC2626"
AMBER = "FCD34D"  # horas cobradas < voladas (regla: nunca cobrar de menos)
# TC oficial de referencia autocompletado (la cotización no traía TC): azul
# claro, sutil (pedido 27-ago). La nota de la celda dice fuente y fecha.
TC_OFICIAL_FILL = "DCE9F8"
# Colores suaves por bloque (ayuda visual pedida en el contrato).
FILL_VENTA = "E8F0FB"  # azul suave
FILL_COSTOS = "FDF3E7"  # ámbar suave
FILL_COBROS = "E9F6EC"  # verde suave

MONEY = "#,##0.00"
# Columna en DÓLARES de la hoja 'inventario' (22-sep-2026): lleva el "$"
# pegado al número para que a simple vista no se confunda con las columnas
# en pesos, que van sin símbolo.
MONEY_USD = '"$"#,##0.00'
HORAS = "0.00"
TACO = "0.0"
# T.C. USD→MXN: 2 decimales mínimo y hasta 6 (17-sep-2026). El API los
# persiste con 6 (`numeric(12,6)`) porque con ellos cuadra el total en
# pesos; mostrar «16.9916» escondía los que hacían cuadrar la cotización.
TC = "0.00####"

_thin = Side(style="thin", color="D5DBE3")
_border = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)


# Formato condicional «es menor que 0» (8-oct-2026, pedido del cliente con la
# captura del RESUMEN: la GANANCIA negativa del N58BT en rojo). Es una REGLA
# de Excel (cfRule), no un relleno fijo: Excel la evalúa con el número que la
# celda muestre, también cuando la oficina edita la fórmula. Colores del
# preset «Resaltar reglas de celdas › Es menor que… › Relleno rojo claro con
# texto rojo oscuro». Se aplica a TODAS las celdas de ganancia / utilidad /
# remanente de los dos libros (`_rojo_si_negativo`): hoja maestra
# (`_COLUMNAS_NEGATIVO_ROJO`, filas + TOTALES), RESUMEN del general, cascada y
# socios de 'balance' (y el RESULTADO del bloque VUELATOUR), 'otros
# movimientos' (remanente) y GANANCIA de 'refacciones'. Las huellas de layout
# de los tests no miran las reglas condicionales: las hojas SE VEN iguales
# mientras nada sea negativo.
_FILL_NEGATIVO = PatternFill("solid", fgColor="FFC7CE", bgColor="FFC7CE")
_FONT_NEGATIVO = Font(color="9C0006")
_COLUMNAS_NEGATIVO_ROJO = ("remanente_mxn", "ganancia_mxn", "ganancia_usd", "cph_remanente_mxn")


def _rojo_si_negativo(ws: Worksheet, *rangos: str) -> None:
    """Relleno rojo claro + letra rojo oscuro en las celdas de `rangos` cuyo
    valor sea < 0, con formato condicional de Excel (una regla por rango)."""
    for rango in rangos:
        ws.conditional_formatting.add(
            rango,
            CellIsRule(
                operator="lessThan", formula=["0"], fill=_FILL_NEGATIVO, font=_FONT_NEGATIVO
            ),
        )


def _hex(color: str | None) -> str | None:
    """#RRGGBB → RRGGBB tal cual (mismo criterio que el Libro Dinero: el
    equipo usa sus colores saturados por avión, no se aclaran)."""
    if not color:
        return None
    c = color.lstrip("#").strip()
    return c.upper() if len(c) == 6 else None


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


# Decimales que MUESTRA cada formato de número: xlsx_formulas exige que la
# fórmula se VEA igual que el número del API con esos decimales.
_DECIMALES_FMT = {
    MONEY: 2,
    MONEY_USD: 2,
    HORAS: 2,
    TACO: 1,
    TC: 6,
    "0": 0,
    "0.0%": 3,
    "General": 6,
}


def _formula(
    ws: Worksheet,
    row: int,
    col: int,
    formula: str | None,
    valor: float | None,
    fmt: str = MONEY,
    **font,
):
    """Como `_num`, pero la celda lleva una FÓRMULA visible (pedido del
    cliente, 5-oct-2026: «que tenga las fórmulas visibles, no que sea solo
    un texto con la cantidad») y el número del API en caché.

    `valor` None ⇒ celda vacía y SIN fórmula (nunca un 0 falso). `formula`
    None ⇒ el valor tal cual (sus insumos no están en el libro). Al cerrar el
    libro, `xlsx_formulas.finalizar` evalúa la fórmula como Excel y, si no
    reproduce `valor` con los decimales del formato, la vuelve VALOR: los
    números del libro no cambian ni un centavo."""
    cell = _num(ws, row, col, valor if formula is None else None, fmt, **font)
    if formula is not None and valor is not None:
        xlsx_formulas.escribir(cell, formula, valor, _DECIMALES_FMT.get(fmt, 2))
    return cell


def _ref_hoja(titulo: str, celda: str) -> str:
    """Referencia a otra hoja: 'nombre'!A1 (comillas siempre; una comilla del
    nombre se duplica, regla de Excel)."""
    return "'" + titulo.replace("'", "''") + "'!" + celda


def _suma_none(a: float | None, b: float | None) -> float | None:
    """Suma None-tolerante de dos montos que YA manda el API (solo para
    mostrarlos juntos, nunca negocio): None + None = None (celda vacía, no
    un 0 falso), None + x = x, x + y = round(x + y, 2)."""
    if a is None:
        return b
    if b is None:
        return a
    return round(a + b, 2)


# Marca que el API deja en el DETALLE de un parcial del reparto manual
# (aircraft-balance.service: "reparto manual: $X de $Y MXN").
_MARCA_PARCIAL = "reparto manual:"


def _es_parcial_reparto(detalle: str | None) -> bool:
    return bool(detalle) and _MARCA_PARCIAL in detalle


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
    ("VENTA", "VENTA AVIÓN\nUSD *", "total_usd", MONEY),
    ("VENTA", "IVA VENTA\nAVIÓN USD", "iva_usd", MONEY),
    ("VENTA", "TIPO CAMBIO\nVENTA", "tc_venta", TC),
    ("VENTA", "VENTA AVIÓN\nMXN", "total_mxn", MONEY),
    ("VENTA", "IVA VENTA\nAVIÓN MXN", "iva_mxn", MONEY),
    ("VENTA", "TOTAL S/IVA\nMXN", "subtotal_mxn", MONEY),
    ("TIEMPO / TACÓMETRO", "TIEMPO\nVUELO HR", "tiempo_vuelo", HORAS),
    ("TIEMPO / TACÓMETRO", "TACO\nINICIO", "taco_inicio", TACO),
    ("TIEMPO / TACÓMETRO", "TACO\nFINAL", "taco_fin", TACO),
    # GAS fuera de la hoja maestra (26-ago-2026): el combustible se controla
    # por avión/mes en su propia hoja "combustible" (los campos gas_* siguen
    # en el esquema por compat de contrato; un API viejo aún los manda).
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
    # Antes «COMISIÓN VENDEDOR MXN» (vacía desde el 28-ago-2026); desde el
    # 6-oct-2026 (API 0.0.65) COMISIONES = lo que absorbe el avión con la
    # regla de septiembre 2026 (`comisiones_mxn`), misma posición. Desde el
    # 7-oct-2026 solo en el libro INDIVIDUAL (el general usa `_COLS_FLOTA`).
    ("INDICADORES USD / IVA", "COMISIONES\nMXN", "comisiones_mxn", MONEY),
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
    # Regla 28-ago-2026: primero los depósitos REALES (Σ COBRO 1..4, tal
    # cual entraron; NETOS de la comisión del banco con la regla de
    # septiembre 2026) y aparte la parte de los cobros que corresponde al
    # AVIÓN (prorrateo de los BRUTOS hecho por el API; el resto es de
    # VuelaTour): va ANTES de comisiones, que restan en COMISIONES.
    ("STATUS DE COBROS", "COBRADO REAL\nMXN (Σ depósitos)", "cobrado_real_mxn", MONEY),
    (
        "STATUS DE COBROS",
        "COBRADO AVIÓN MXN\n(prorrateado, antes\nde comisiones) ****",
        "cobrado_mxn",
        MONEY,
    ),
    ("STATUS DE COBROS", "POR COBRAR\nMXN", "por_cobrar_mxn", MONEY),
    ("STATUS DE COBROS", "POR COBRAR\nUSD", "por_cobrar_usd", MONEY),
    # 30-sep-2026 (API 0.0.45, pedido de Marie): folio de la factura del
    # servicio emitida al cliente, YA resuelto por el API (misma etiqueta que
    # el Libro Dinero y 'otros movimientos'). Texto tal cual; sin dato o con
    # un API viejo, celda vacía. Va AL FINAL: no corre COBRO 1 ni COMISIONES
    # (`_DISP_MENSUAL.cobro1_col` / `.comision_col`; todo índice se busca por
    # atributo).
    ("STATUS DE COBROS", "FACTURA\nVUELATOUR", "factura_vuelatour", None),
]
# Hoja de vuelos del Balance GENERAL, variante «mensual» (7-oct-2026, API
# 0.0.66). Pedido del cliente: «estas comisiones se están duplicando en la
# general, ya que las tenemos en el apartado de "Otros movimientos". Las
# comisiones deben aparecer en ese apartado pero en el individual de los
# avioncitos, ya que en el reporte de los aviones no está la pestaña "Otros
# movimientos"». El juego de siempre SIN la columna COMISIONES MXN (tampoco
# la vieja COMISIÓN VENDEDOR MXN): todo lo que estaba a su derecha corre una
# columna y su GANANCIA va antes de comisiones (`_antes_de_comisiones`). El
# libro INDIVIDUAL conserva `_COLS` tal cual.
_COLS_FLOTA: list[tuple[str, str, str | None, str | None]] = [
    c for c in _COLS if c[2] != "comisiones_mxn"
]
# Base del reparto de la venta en vuelos MULTI-AVIÓN (participacion_fuente):
# SIEMPRE partes iguales por tramo vendido (nunca por horas); 'unico' no
# lleva base.
_FUENTE_PARTICIPACION = {
    "tramos": "partes iguales por tramo vendido (ferries/tramos operativos "
    "no reparten)",
}
# Fuente del TC oficial de referencia (TipoCambioService.oficialDetallePara).
_TC_FUENTE_LABEL = {
    "OPEN_ER_API": "open.er-api.com",
    "ECB_FRANKFURTER": "referencia BCE (frankfurter)",
    "BANXICO_FIX": "Banxico FIX",
}
# Celdas de costos con NOTA de desglose (comentario de Excel): al pasar el
# cursor se ve qué gastos componen el total ("Comida · Starbucks — $206.00").
_DETALLE_ATTR = {
    "op_mxn": "op_detalle",
    "piloto_mxn": "piloto_detalle",
    "otros_mxn": "otros_detalle",
}
_GROUP_FILLS = {"VENTA": FILL_VENTA, "COSTOS DIRECTOS (MXN)": FILL_COSTOS,
                "STATUS DE COBROS": FILL_COBROS}
# Columnas de la fila TOTALES: atributo del vuelo → atributo de totales.
_TOTAL_MAP = {
    "horas_cobradas": "horas_cobradas",
    "tiempo_vuelo": "tiempo_vuelo",
    "total_mxn": "total_mxn",
    "iva_mxn": "iva_mxn",
    "subtotal_mxn": "subtotal_mxn",
    "op_mxn": "op_mxn",
    "piloto_mxn": "piloto_mxn",
    "otros_mxn": "otros_mxn",
    "permiso_afac_mxn": "permiso_afac_mxn",
    "costo_total_mxn": "costo_total_mxn",
    "remanente_mxn": "remanente_mxn",
    "dif_iva_mxn": "dif_iva_mxn",
    # COMISIONES: `comisiones_mxn` de los totales; sin ella (API ≤ 0.0.64),
    # `comision_vendedor_mxn` como siempre (`_valor_total`).
    "comisiones_mxn": "comisiones_mxn",
    "ganancia_mxn": "ganancia_mxn",
    "ganancia_usd": "ganancia_usd",
    "cobrado_mxn": "cobrado_mxn",
    "cobrado_real_mxn": "cobrado_real_mxn",
    "por_cobrar_mxn": "por_cobrar_mxn",
    "por_cobrar_usd": "por_cobrar_usd",
    # Promedios (se pintan en su columna con etiqueta "prom." implícita):
    "tc_costos": "tc_promedio",
    "costo_hr_usd": "costo_hr_prom_usd",
}


# ===== Fórmulas visibles de la hoja maestra (5-oct-2026) =====
# Pedido del cliente: «¿me puedes apoyar poniendo con fórmulas las celdas que
# sean por suma? Por ejemplo los permisos de AFAC…». Cada fórmula reproduce
# EXACTAMENTE la aritmética del API (aircraft-balance.service.ts, letras del
# libro manual entre corchetes): ROUND(…, 2) justo donde el API hace round2 y
# en ningún otro lado. Lo que depende de datos que NO están en el libro
# (IVA % del cliente, ajustes de la cotización, prorrateos multi-avión, TC de
# cada gasto) sigue como valor. Las letras de las columnas salen SIEMPRE de
# la disposición de la hoja (`_Disposicion.letra`, `.cobro_mxn_letras`): desde
# el 6-oct-2026 hay dos juegos de columnas y un mapa fijo de letras apuntaría,
# en silencio, a las del Balance mensual dentro del Balance general.
# Constantes del encabezado: fila 1 de CLAVE..ESTADO, que no llevan título de
# grupo — no se mueve ni una fila ni una columna de datos. Se citan con $.
_CELDA_AFAC = "$B$1"
_CELDA_FACTOR_IVA = "$D$1"
# Factor con que el API quita el IVA a los costos (AF = AE / 1.16).
FACTOR_IVA_COSTOS = 1.16
# Totales que son PROMEDIO (el API: promedio de los no nulos); los demás de
# la fila TOTALES son suma de las filas de vuelo (round2(Σ no nulos)).
_TOTALES_PROMEDIO = ("tc_promedio", "costo_hr_prom_usd")


# ===== Variante «Balance general» de la hoja maestra (6-oct-2026, API 0.0.64)
# Pedido del cliente: «el que va a ser balance mensual está perfecto como
# está ahora; el que cambiaría sería el balance general: todos esos montos de
# operación, piloto, pagos AFAC se van a eliminar y se va a hacer un resumen
# del total de COSTO TOTAL. En lugar del apartado de Operaciones, piloto y
# otros se van a agregar estas columnas para sacar el costo por hora por cada
# vuelo» (las de su hoja «utilidades»), con la fórmula: «es el total de todos
# los gastos, entre el tiempo volado, entre el tipo de cambio del día, entre
# 1.16 (para sacar subtotal) y el resultado sería el costo por hora».
# Mismos campos del API (no hay números nuevos): CLAVE..ESTADO y VENTA
# iguales; «COSTO POR HORA» = las 12 columnas del cliente EN SU ORDEN y con
# sus nombres (repiten a propósito la venta s/IVA: el bloque se lee solo,
# como su hoja); STATUS DE COBROS y FACTURA VUELATOUR iguales. Fuera:
# OPERACIONES, PILOTO, OTROS, PERMISO AFAC (su desglose va en la NOTA de
# TOTAL PARA PROVEEDOR (PESOS)) e INDICADORES (cubiertos por el bloque nuevo;
# DIF. IVA y GANANCIA salen). **8-oct-2026 (pedido del cliente: «estos
# valores están duplicados… eliminar tacómetro inicial y final, solo en la
# general»): fuera también TIEMPO / TACÓMETRO (TIEMPO VUELO HR, TACO INICIO,
# TACO FINAL) y «COSTO TOTAL (MXN)» (COSTO TOTAL MXN, TIPO CAMBIO COSTOS)** —
# el bloque ya trae el tiempo (TIEMPO CALZOS/HOBS), el T.C. (TIPO CAMBIO) y
# el costo total (TOTAL PARA PROVEEDOR (PESOS)); los tacómetros de cada vuelo
# siguen en el Balance mensual. Esas tres columnas del bloque son ahora el
# ÚNICO lugar del dato y las demás hojas las citan por ALIAS de llave
# (`tiempo_vuelo` → `cph_tiempo_hr`, `tc_costos` → `cph_tc`,
# `costo_total_mxn` → `cph_proveedor_mxn`; `_disposicion`). La 3.ª posición
# de cada tupla es la LLAVE de la columna: el atributo del vuelo o, en las
# que repiten un dato o lo derivan, una llave `cph_*` (`_ORIGEN_GENERAL` /
# `_valor_columna`). Toda fórmula cita columnas por su llave
# (`_Disposicion.letra`), jamás por una letra fija.
# COMISIONES: del 6-oct-2026 (API 0.0.65) al 7-oct-2026 iban al final del
# bloque y REMANENTE VENTA MENOS COMPRA las restaba. Desde el API 0.0.66 la
# hoja de vuelos del Balance general ya no las lleva (pedido del cliente: se
# duplicaban con 'otros movimientos'; ver `_COLS_FLOTA`) y REMANENTE VENTA
# MENOS COMPRA vuelve a ser VENTA − PROVEEDOR, la definición del cliente: su
# llave `cph_remanente_mxn` repite el REMANENTE del API (antes de comisiones).
_GRUPO_COSTO_HORA = "COSTO POR HORA"
_COLS_GENERAL: list[tuple[str, str, str | None, str | None]] = [
    *(c for c in _COLS if c[0] in ("", "VENTA")),
    (_GRUPO_COSTO_HORA, "TOTAL COBRADO\nS/IVA (PESOS)", "cph_cobrado_siva_mxn", MONEY),
    (_GRUPO_COSTO_HORA, "TIEMPO\nCALZOS/HOBS (HR)", "cph_tiempo_hr", HORAS),
    (_GRUPO_COSTO_HORA, "COSTO X HORA\n(DLLS S/IVA)", "costo_hr_usd_siva", MONEY),
    (_GRUPO_COSTO_HORA, "IVA X HR\n(DLLS)", "cph_iva_hr_usd", MONEY),
    (_GRUPO_COSTO_HORA, "COSTO HR\nMÁS IVA (DLLS)", "costo_hr_usd", MONEY),
    (_GRUPO_COSTO_HORA, "TOTAL PARA\nPROVEEDOR (DLLS)", "costo_usd", MONEY),
    (_GRUPO_COSTO_HORA, "IVA TOTAL\nPAGADO (DLLS)", "iva_pagado_usd", MONEY),
    (_GRUPO_COSTO_HORA, "TIPO\nCAMBIO", "cph_tc", TC),
    (_GRUPO_COSTO_HORA, "TOTAL PARA\nPROVEEDOR (PESOS)", "cph_proveedor_mxn", MONEY),
    (_GRUPO_COSTO_HORA, "IVA TOTAL\nPAGADO (PESOS)", "iva_pagado_mxn", MONEY),
    (_GRUPO_COSTO_HORA, "TOTAL PAGADO\nS/IVA (PESOS)", "cph_pagado_siva_mxn", MONEY),
    (_GRUPO_COSTO_HORA, "REMANENTE VENTA\nMENOS COMPRA (PESOS)", "cph_remanente_mxn", MONEY),
    *(c for c in _COLS if c[0] == "STATUS DE COBROS"),
]
# Columnas del bloque COSTO POR HORA que REPITEN un dato del vuelo: llave →
# atributo del vuelo. Mismo número: TOTAL COBRADO S/IVA lleva la misma
# fórmula que TOTAL S/IVA MXN, REMANENTE VENTA MENOS COMPRA la del REMANENTE
# (VENTA − PROVEEDOR, antes de comisiones) y las otras tres son valor. Desde
# el 8-oct-2026 el tiempo, el T.C. y el costo total ya NO tienen otra columna
# en el general: el atributo entra a `_Disposicion.letra` como ALIAS de su
# columna del bloque (`_disposicion`), y así `maestra.total('tc_costos')`,
# `rango('tiempo_vuelo')` y las señales por origen siguen funcionando.
_ORIGEN_GENERAL = {
    "cph_cobrado_siva_mxn": "subtotal_mxn",
    "cph_tiempo_hr": "tiempo_vuelo",
    "cph_tc": "tc_costos",
    "cph_proveedor_mxn": "costo_total_mxn",
    "cph_remanente_mxn": "remanente_mxn",
}
# TOTALES de la variante: los del API por atributo (los mismos de siempre;
# las columnas repetidas llevan el total de su dato). Las del bloque que el
# API no totaliza van en `_CPH_TOTALES_PROPIOS` (abajo).
_TOTAL_MAP_GENERAL = {
    **{k: t for k, t in _TOTAL_MAP.items() if any(c[2] == k for c in _COLS_GENERAL)},
    **{k: _TOTAL_MAP[a] for k, a in _ORIGEN_GENERAL.items()},
}
# TOTALES del bloque COSTO POR HORA que el API NO manda (revisión 6-oct-2026:
# con huecos, la fila TOTALES «parece rota»). No es un número nuevo:
# - TOTAL PARA PROVEEDOR (DLLS), IVA TOTAL PAGADO (DLLS y PESOS) y TOTAL
#   PAGADO S/IVA (PESOS) = Σ de los números de las filas TAL COMO SE VEN (los
#   del API, a centavos). Sus celdas de fila son fórmulas SIN ROUND (el API
#   solo redondea al serializar), así que ROUND(SUM(rango),2) sumaría los
#   valores completos y no daría esa Σ (en el libro de pruebas, un centavo de
#   diferencia en las cuatro; en un mes real, casi siempre) y la celda se
#   quedaría como valor: la fórmula redondea cada fila antes de sumar,
#   ROUND(SUMPRODUCT(ROUND(rango,2)),2).
# - COSTO X HORA (DLLS S/IVA) = COSTO HR MÁS IVA de TOTALES (el promedio del
#   API) ÷ 1.16, como cada fila (AO = AN ÷ 1.16). No es AVERAGEIF de la
#   columna: en el consolidado el API promedia los promedios de cada avión
#   (no las filas) y la celda no lo reproduciría.
# - IVA X HR (DLLS) = COSTO HR MÁS IVA − COSTO X HORA de TOTALES.
_CPH_TOTALES_SUMA_FILAS = ("costo_usd", "iva_pagado_usd", "iva_pagado_mxn", "cph_pagado_siva_mxn")
_CPH_TOTALES_PROPIOS = frozenset((*_CPH_TOTALES_SUMA_FILAS, "costo_hr_usd_siva", "cph_iva_hr_usd"))
# Título del bloque nuevo en la fila de grupos: el encabezado de la hoja dice
# que es el «Balance general» (A1:D1 son las constantes y no se mueven).
_TITULO_GRUPO_COSTO_HORA = "BALANCE GENERAL · COSTO POR HORA"
# Ancho de las 12 columnas del bloque: sus encabezados son más largos.
_ANCHO_COSTO_HORA = 15


@dataclass(frozen=True)
class _Disposicion:
    """Juego de columnas de la hoja maestra (6-oct-2026): el de siempre
    (`_COLS`: libro individual), el de la variante «mensual» del Balance
    general (`_COLS_FLOTA`: el de siempre sin COMISIONES, 7-oct-2026) o el de
    la variante «general» (`_COLS_GENERAL`). Todo lo que la hoja y las demás
    pestañas necesitan saber de las columnas sale de aquí, por LLAVE — nunca
    una letra fija. Un juego sin COMISIONES (`comision_col` None) pinta la
    hoja antes de comisiones (`_antes_de_comisiones`)."""

    cols: tuple[tuple[str, str, str | None, str | None], ...]
    # Llave → letra de su columna. En un juego con COSTO POR HORA, un dato
    # cuya columna propia no está (tiempo, T.C., costo total; 8-oct-2026)
    # entra como ALIAS de la columna del bloque que lo repite.
    letra: dict[str, str]
    cobro1_col: int
    cobro_mxn_letras: tuple[str, ...]
    comision_col: int | None  # COMISIONES MXN (su encabezado lleva la nota)
    total_map: dict[str, str]
    fills: dict[str, str]
    detalle_attr: dict[str, str]
    titulos_grupo: dict[str, str]
    costo_por_hora: bool


def _disposicion(
    cols: list[tuple[str, str, str | None, str | None]],
    *,
    total_map: dict[str, str],
    fills: dict[str, str],
    detalle_attr: dict[str, str],
    titulos_grupo: dict[str, str] | None = None,
    costo_por_hora: bool = False,
) -> _Disposicion:
    claves = [c[2] for c in cols if c[2]]
    if len(claves) != len(set(claves)):
        raise ValueError("llave de columna repetida en la hoja maestra")
    cobro1 = next(i for i, c in enumerate(cols, start=1) if c[1] == "COBRO 1\nFECHA")
    letra = {c[2]: get_column_letter(i) for i, c in enumerate(cols, start=1) if c[2]}
    if costo_por_hora:
        # Alias (8-oct-2026): el atributo de un dato que SOLO vive en el
        # bloque COSTO POR HORA resuelve a la letra de esa columna; si el
        # juego trae la columna propia, esta manda. Sin la columna del bloque
        # tampoco hay alias (la cita cae al VALOR del API con nota).
        for cph, origen in _ORIGEN_GENERAL.items():
            if origen not in letra and cph in letra:
                letra[origen] = letra[cph]
    return _Disposicion(
        cols=tuple(cols),
        letra=letra,
        cobro1_col=cobro1,
        cobro_mxn_letras=tuple(get_column_letter(cobro1 + 1 + 2 * k) for k in range(4)),
        comision_col=next(
            (i for i, c in enumerate(cols, start=1) if c[2] == "comisiones_mxn"), None
        ),
        total_map=total_map,
        fills=fills,
        detalle_attr=detalle_attr,
        titulos_grupo=titulos_grupo or {},
        costo_por_hora=costo_por_hora,
    )


_DISP_MENSUAL = _disposicion(
    _COLS, total_map=_TOTAL_MAP, fills=_GROUP_FILLS, detalle_attr=_DETALLE_ATTR
)
# Balance general, variante «mensual» (7-oct-2026, API 0.0.66): el juego de
# siempre sin COMISIONES; mismos TOTALES del API (COMISIONES no tiene columna).
_DISP_FLOTA = _disposicion(
    _COLS_FLOTA, total_map=_TOTAL_MAP, fills=_GROUP_FILLS, detalle_attr=_DETALLE_ATTR
)
_DISP_GENERAL = _disposicion(
    _COLS_GENERAL,
    total_map=_TOTAL_MAP_GENERAL,
    fills={"VENTA": FILL_VENTA, "STATUS DE COBROS": FILL_COBROS},
    # El desglose de operación/piloto/otros/AFAC va en UNA nota, la de TOTAL
    # PARA PROVEEDOR (PESOS) (`_nota_desglose_costo`).
    detalle_attr={},
    titulos_grupo={_GRUPO_COSTO_HORA: _TITULO_GRUPO_COSTO_HORA},
    costo_por_hora=True,
)


def _resta(a: float | None, b: float | None) -> float | None:
    """Diferencia de dos montos que YA manda el API (dos columnas del bloque
    COSTO POR HORA son la resta de otras dos, como en la hoja del cliente):
    None si falta un lado (celda vacía, nunca un 0 falso)."""
    if a is None or b is None:
        return None
    return round(a - b, 2)


def _comisiones(x: BalanceAvionVuelo | BalanceAvionTotales) -> float | None:
    """Número de la columna COMISIONES MXN (fila o TOTALES): `comisiones_mxn`
    del API 0.0.65 o, sin él, `comision_vendedor_mxn` — la llave con que un
    API ≤ 0.0.64 llenaba la columna (vacía en las filas y 0 en TOTALES desde
    el 28-ago-2026): su libro sale como siempre."""
    return x.comisiones_mxn if x.comisiones_mxn is not None else x.comision_vendedor_mxn


def _valor_columna(clave: str, v: BalanceAvionVuelo) -> float | str | None:
    """Valor de la celda de la columna `clave` en la fila del vuelo `v`: el
    atributo del vuelo tal cual (o el que repite, `_ORIGEN_GENERAL`); IVA X
    HR y TOTAL PAGADO S/IVA son la resta de dos campos del API."""
    if clave == "cph_iva_hr_usd":  # COSTO HR MÁS IVA − COSTO X HORA
        return _resta(v.costo_hr_usd, v.costo_hr_usd_siva)
    if clave == "cph_pagado_siva_mxn":  # TOTAL PARA PROVEEDOR − IVA PAGADO
        return _resta(v.costo_total_mxn, v.iva_pagado_mxn)
    if clave == "comisiones_mxn":
        return _comisiones(v)
    return getattr(v, _ORIGEN_GENERAL.get(clave, clave))


def _valor_total(t: BalanceAvionTotales, total_attr: str) -> float | None:
    """Número de la celda TOTALES para el atributo de totales `total_attr`
    (el de `total_map`): tal cual, salvo COMISIONES (`_comisiones`)."""
    if total_attr == "comisiones_mxn":
        return _comisiones(t)
    return getattr(t, total_attr)


def _numero_de_celda(x: float) -> float:
    """El número tal como Excel lo lee de la celda: openpyxl lo escribe con
    «%.16g» (el evaluador de `xlsx_formulas` lee igual)."""
    return float(format(x, ".16g"))


def _fila_antes_de_comisiones(v: BalanceAvionVuelo) -> BalanceAvionVuelo:
    """La fila de un vuelo como se ve en una hoja SIN la columna COMISIONES
    (Balance general, 7-oct-2026): si el avión absorbe comisiones en ella
    (regla de septiembre 2026), GANANCIA MXN = el REMANENTE del API y
    GANANCIA USD = ese remanente ÷ el T.C. de costos de la fila — la fórmula
    visible de la celda, ROUND(GANANCIA MXN ÷ TIPO CAMBIO COSTOS, 2), con lo
    que el API mandaba antes de la regla; sin T.C. de costos va vacía (el API
    la convierte con el T.C. promedio del libro de SU avión, que esta hoja no
    tiene: mejor vacía que un número que no es). Sin comisiones, la fila tal
    cual (su GANANCIA ya es el remanente). Con COMISIONES en 0 explícito (el
    API manda null, pero el esquema admite la «regla con 0»), una copia SIN
    la llave, que se pinta igual que la de null: si no, GANANCIA MXN citaba
    la columna COMISIONES, que esta hoja no tiene, y la celda se iba como
    valor con la nota «calculado por el sistema» en su encabezado (revisión
    7-oct-2026)."""
    comisiones = _comisiones(v)
    if comisiones is None:
        return v
    if comisiones == 0:
        return v.model_copy(update={"comisiones_mxn": None, "comision_vendedor_mxn": None})
    ganancia_usd = None
    if v.remanente_mxn is not None and v.tc_costos:
        ganancia_usd = xlsx_formulas.redondear_excel(
            _numero_de_celda(v.remanente_mxn) / _numero_de_celda(v.tc_costos), 2
        )
    return v.model_copy(
        update={
            "comisiones_mxn": None,
            "comision_vendedor_mxn": None,
            "ganancia_mxn": v.remanente_mxn,
            "ganancia_usd": ganancia_usd,
        }
    )


def _antes_de_comisiones(req: BalanceAvionRequest) -> BalanceAvionRequest:
    """El libro tal como lo pinta una hoja de vuelos SIN la columna
    COMISIONES (las dos variantes del Balance general, 7-oct-2026, API
    0.0.66; pedido del cliente: «estas comisiones se están duplicando en la
    general, ya que las tenemos en "Otros movimientos"»): cada fila con
    comisiones va antes de ellas (`_fila_antes_de_comisiones`) y TOTALES
    GANANCIA MXN = el REMANENTE de TOTALES del API y GANANCIA USD = Σ de las
    filas tal como se ven (sumadas en orden, como SUM de Excel; la fórmula
    ROUND(SUM()) de la celda da ese número). Lo que absorbe cada avión sigue
    restando en la cascada de su bloque de la hoja 'balance' (el número del
    API) y, por vuelo, en la columna COMISIONES de su libro INDIVIDUAL. Sin
    filas con comisiones (vuelos anteriores a la regla o un API previo), el
    libro tal cual — byte a byte. Con filas de COMISIONES en 0 explícito y
    ninguna ≠ 0, solo esas filas cambian de objeto (sin la llave) y los
    TOTALES son los del API: su GANANCIA ya es el remanente, como con null.
    Solo PRESENTACIÓN: ningún número del API se recalcula, se elige el de
    antes de comisiones."""
    vuelos = [_fila_antes_de_comisiones(v) for v in req.vuelos]
    if all(nueva is vieja for nueva, vieja in zip(vuelos, req.vuelos, strict=True)):
        return req
    if not any(_monto_no_cero(_comisiones(v)) for v in req.vuelos):
        return req.model_copy(update={"vuelos": vuelos})
    ganancia_usd = 0.0
    for v in vuelos:
        if v.ganancia_usd is not None:
            ganancia_usd += _numero_de_celda(v.ganancia_usd)
    totales = req.totales.model_copy(
        update={
            "comisiones_mxn": None,
            "ganancia_mxn": req.totales.remanente_mxn,
            "ganancia_usd": xlsx_formulas.redondear_excel(ganancia_usd, 2),
        }
    )
    return req.model_copy(update={"vuelos": vuelos, "totales": totales})


class _ColumnaAusente(Exception):  # noqa: N818 - no es un error: es una señal
    """La fórmula de una celda cita una columna que no está en la variante
    de la hoja: la celda va como VALOR del API («calculado por el sistema»)."""


@dataclass(frozen=True)
class _Maestra:
    """Dónde quedó la hoja maestra: las demás hojas la citan en sus
    fórmulas (TC promedio, horas voladas, totales de venta y cobranza)."""

    titulo: str
    fila_ini: int
    fila_fin: int  # < fila_ini si el periodo no tiene vuelos
    fila_tot: int
    general: bool
    disp: _Disposicion  # juego de columnas de la hoja (6-oct-2026)

    @property
    def hay_filas(self) -> bool:
        return self.fila_fin >= self.fila_ini

    def total(self, attr: str) -> str | None:
        """Celda TOTALES de la columna `attr` (absoluta, con la hoja). None si
        la columna no está en esta variante de la hoja: quien la cite escribe
        el VALOR del API con nota «calculado por el sistema»
        (`_formula_cita`)."""
        letra = self.disp.letra.get(attr)
        if letra is None:
            return None
        return _ref_hoja(self.titulo, f"${letra}${self.fila_tot}")

    def rango(self, attr: str, *, local: bool = False) -> str:
        """Filas de vuelo de la columna `attr` (absoluto; con la hoja salvo
        `local`, para fórmulas de la propia hoja maestra)."""
        letra = self.disp.letra[attr]
        celdas = f"${letra}${self.fila_ini}:${letra}${self.fila_fin}"
        return celdas if local else _ref_hoja(self.titulo, celdas)

    def tc_promedio_crudo(self, *, local: bool = False) -> str | None:
        """El TC promedio SIN redondear con que el API convierte las hojas de
        gastos (`tcPromedio` = promedio de los TC de costos no nulos de las
        filas). La celda TOTALES lo muestra con round2: dividir entre ella
        no reproduciría al API. Solo en el libro INDIVIDUAL: en el general
        cada avión usó el suyo."""
        if self.general or not self.hay_filas:
            return None
        return f"AVERAGE({self.rango('tc_costos', local=local)})"


def _formula_fila(
    attr: str | None,
    v: BalanceAvionVuelo,
    r: int,
    *,
    hay_afac: bool,
    tc_prom_filas: str | None,
    disp: _Disposicion = _DISP_MENSUAL,
) -> str | None:
    """Fórmula de la celda `attr` en la fila `r` de un vuelo, o None si la
    celda va como valor. Espejo de la fila del API (paso 2 de buildPayload).
    Las columnas se citan por su llave en `disp`; si la fórmula necesita una
    que la variante no tiene, levanta `_ColumnaAusente` (la celda va como
    valor)."""

    def c(a: str) -> str:
        letra = disp.letra.get(a)
        if letra is None:
            raise _ColumnaAusente(a)
        return f"{letra}{r}"

    if disp.costo_por_hora:
        formula = _formula_costo_por_hora(attr, c)
        if formula is not None:
            return formula
    if attr == "total_mxn":  # [L] = round2(I × K)
        return f"ROUND({c('total_usd')}*{c('tc_venta')},2)"
    if attr == "iva_mxn":  # [M] = round2(J × K)
        return f"ROUND({c('iva_usd')}*{c('tc_venta')},2)"
    if attr == "subtotal_mxn":  # [N] = round2(L − M)
        return f"ROUND({c('total_mxn')}-{c('iva_mxn')},2)"
    if attr == "permiso_afac_mxn":  # [X] = round2(tarifa × z × D)
        if not hay_afac:
            return None
        return f"ROUND({_CELDA_AFAC}*{c('tc_costos')}*{c('horas_cobradas')},2)"
    if attr == "costo_total_mxn":  # [Y] = round2(op + piloto + otros + X)
        return f"ROUND({c('op_mxn')}+{c('piloto_mxn')}+{c('otros_mxn')}+{c('permiso_afac_mxn')},2)"
    if attr == "costo_usd":  # [AE] = Y / z (sin redondeo)
        return f"{c('costo_total_mxn')}/{c('tc_costos')}"
    if attr == "costo_usd_siva":  # [AF] = AE / 1.16
        return f"{c('costo_usd')}/{_CELDA_FACTOR_IVA}"
    if attr == "iva_pagado_usd":  # [AG] = AE − AF
        return f"{c('costo_usd')}-{c('costo_usd_siva')}"
    if attr == "iva_pagado_mxn":  # [AH] = AG × z
        return f"{c('iva_pagado_usd')}*{c('tc_costos')}"
    if attr == "remanente_mxn":  # [AI] = round2(L − Y)
        return f"ROUND({c('total_mxn')}-{c('costo_total_mxn')},2)"
    if attr == "dif_iva_mxn":  # [AJ] = round2(M − AH)
        return f"ROUND({c('iva_mxn')}-{c('iva_pagado_mxn')},2)"
    if attr == "ganancia_mxn":
        # [AL] = round2(AI − COMISIONES) (regla de septiembre 2026, API
        # 0.0.65). Fila sin comisiones (vuelo anterior o API previo): AL = AI,
        # la referencia de siempre.
        if _comisiones(v) is None:
            return c("remanente_mxn")
        return f"ROUND({c('remanente_mxn')}-{c('comisiones_mxn')},2)"
    if attr == "ganancia_usd":  # [AM] = round2(AL / (z ?? tcPromedio))
        if v.tc_costos is not None:
            return f"ROUND({c('ganancia_mxn')}/{c('tc_costos')},2)"
        if tc_prom_filas is None:
            return None
        return f"ROUND({c('ganancia_mxn')}/{tc_prom_filas},2)"
    if attr == "costo_hr_usd":  # [AN] = AE / O
        return f"{c('costo_usd')}/{c('tiempo_vuelo')}"
    if attr == "costo_hr_usd_siva":  # [AO] = AN / 1.16
        return f"{c('costo_hr_usd')}/{_CELDA_FACTOR_IVA}"
    if attr == "cobrado_real_mxn":  # Σ COBRO 1..4 MXN (el API cuadra la última)
        cobros = ",".join(f"{letra}{r}" for letra in disp.cobro_mxn_letras)
        return f"ROUND(SUM({cobros}),2)"
    if attr == "por_cobrar_mxn":  # round2(L − cobrado avión)
        return f"ROUND({c('total_mxn')}-{c('cobrado_mxn')},2)"
    if attr == "por_cobrar_usd":  # round2(por cobrar MXN / K)
        return f"ROUND({c('por_cobrar_mxn')}/{c('tc_venta')},2)"
    return None


def _formula_costo_por_hora(attr: str | None, c) -> str | None:
    """Fórmulas del bloque COSTO POR HORA de la variante general (6-oct-2026)
    o None si `attr` no es del bloque (o es un dato: TIEMPO, TIPO CAMBIO y
    TOTAL PARA PROVEEDOR (PESOS) van como valor). El bloque se cita a sí
    mismo, como la hoja «utilidades» del cliente, y cada fórmula es la MISMA
    aritmética del API (letras de su fila entre corchetes; ROUND solo donde
    el API hace round2): Y = costo total, z = T.C. de costos, O = tiempo de
    vuelo. `c(llave)` da la celda de la fila."""
    if attr == "cph_cobrado_siva_mxn":  # TOTAL S/IVA [N] = round2(L − M)
        return f"ROUND({c('total_mxn')}-{c('iva_mxn')},2)"
    if attr == "costo_hr_usd_siva":  # COSTO X HORA [AO] = Y / z / O / 1.16
        return f"{c('cph_proveedor_mxn')}/{c('cph_tc')}/{c('cph_tiempo_hr')}/{_CELDA_FACTOR_IVA}"
    if attr == "cph_iva_hr_usd":
        # IVA X HR = COSTO HR MÁS IVA − COSTO X HORA tal como el API los manda
        # (round2 de AN y de AO): se redondean donde el API redondea. Con la
        # resta de las celdas completas (AN − AO) una de cada cuatro filas no
        # daría la diferencia de los dos números a la vista y la celda se
        # quedaba como valor (medido con 20 000 filas, 6-oct-2026).
        return f"ROUND(ROUND({c('costo_hr_usd')},2)-ROUND({c('costo_hr_usd_siva')},2),2)"
    if attr == "costo_hr_usd":  # COSTO HR MÁS IVA [AN] = AE / O
        return f"{c('costo_usd')}/{c('cph_tiempo_hr')}"
    if attr == "costo_usd":  # TOTAL PARA PROVEEDOR DLLS [AE] = Y / z
        return f"{c('cph_proveedor_mxn')}/{c('cph_tc')}"
    if attr == "iva_pagado_usd":  # [AG] = AE − AF, con AF = AE / 1.16
        return f"{c('costo_usd')}-{c('costo_usd')}/{_CELDA_FACTOR_IVA}"
    if attr == "iva_pagado_mxn":  # [AH] = AG × z
        return f"{c('iva_pagado_usd')}*{c('cph_tc')}"
    if attr == "cph_pagado_siva_mxn":  # TOTAL PAGADO S/IVA = Y − AH
        return f"{c('cph_proveedor_mxn')}-{c('iva_pagado_mxn')}"
    if attr == "cph_remanente_mxn":
        # REMANENTE VENTA MENOS COMPRA [AI] = round2(L − Y): la definición del
        # cliente, antes de comisiones (7-oct-2026: el Balance general ya no
        # las resta en la hoja de vuelos).
        return f"ROUND({c('total_mxn')}-{c('cph_proveedor_mxn')},2)"
    return None


def _formula_total(attr: str, maestra: _Maestra) -> str | None:
    """Fórmula de la fila TOTALES para la columna `attr` (None = valor)."""
    if not maestra.hay_filas:
        return None
    letra = maestra.disp.letra[attr]
    rango = f"{letra}{maestra.fila_ini}:{letra}{maestra.fila_fin}"
    total_attr = maestra.disp.total_map.get(attr)
    if total_attr not in _TOTALES_PROMEDIO:  # round2(Σ filas)
        return f"ROUND(SUM({rango}),2)"
    if maestra.general:
        # Consolidado: promedio de los promedios de cada libro (no está en
        # las celdas de esta hoja) — valor.
        return None
    if total_attr == "tc_promedio":  # round2(promedio de z no nulos)
        return f"ROUND(AVERAGE({rango}),2)"
    if total_attr == "costo_hr_prom_usd":
        # costo_hr_prom_usd = round2(promedio de AN con costo > 0): AN > 0
        # ⟺ Y > 0 (z y horas siempre > 0 cuando AN existe).
        return f'ROUND(AVERAGEIF({rango},">0"),2)'
    return None


def _valor_total_costo_por_hora(attr: str, req: BalanceAvionRequest) -> float | None:
    """Número de la celda TOTALES de una columna de `_CPH_TOTALES_PROPIOS`
    (variante general), con los MISMOS números del API: la Σ de los de las
    filas tal como se ven (`_valor_columna`) o, en COSTO X HORA e IVA X HR, el
    promedio COSTO HR MÁS IVA de `totales` ÷ 1.16 y la diferencia de los dos
    — redondeados como los redondea la fórmula (ROUND de Excel). None (celda
    vacía) si el API no manda el promedio."""
    if attr in _CPH_TOTALES_SUMA_FILAS:
        valores = [_valor_columna(attr, v) for v in req.vuelos]
        return xlsx_formulas.redondear_excel(sum(x for x in valores if x is not None), 2)
    prom = req.totales.costo_hr_prom_usd
    if prom is None:
        return None
    siva = xlsx_formulas.redondear_excel(prom / FACTOR_IVA_COSTOS, 2)
    if attr == "costo_hr_usd_siva":
        return siva
    return xlsx_formulas.redondear_excel(prom - siva, 2)  # IVA X HR


def _formula_total_costo_por_hora(attr: str, maestra: _Maestra) -> str | None:
    """Fórmula de la celda TOTALES de una columna de `_CPH_TOTALES_PROPIOS`
    (None = valor: sin filas). Cita las columnas por su llave; si la que
    citaría no está en el juego, levanta `_ColumnaAusente` (la celda va como
    valor con la nota «calculado por el sistema»)."""
    disp = maestra.disp

    def c(a: str) -> str:  # celda TOTALES de la columna `a`
        letra = disp.letra.get(a)
        if letra is None:
            raise _ColumnaAusente(a)
        return f"{letra}{maestra.fila_tot}"

    if attr in _CPH_TOTALES_SUMA_FILAS:  # Σ de las filas tal como se ven
        if not maestra.hay_filas:
            return None
        letra = disp.letra[attr]
        return f"ROUND(SUMPRODUCT(ROUND({letra}{maestra.fila_ini}:{letra}{maestra.fila_fin},2)),2)"
    if attr == "costo_hr_usd_siva":  # COSTO HR MÁS IVA ÷ 1.16
        return f"ROUND({c('costo_hr_usd')}/{_CELDA_FACTOR_IVA},2)"
    if attr == "cph_iva_hr_usd":  # COSTO HR MÁS IVA − COSTO X HORA
        return f"ROUND({c('costo_hr_usd')}-{c('costo_hr_usd_siva')},2)"
    return None


def _cobros_a_4(cobros: list[BalanceAvionCobro]) -> list[BalanceAvionCobro]:
    """Hasta 4 parcialidades; si hay más, la 4ª agrega el resto (solo suma de
    columna para mostrar, no cálculo de negocio)."""
    if len(cobros) <= 4:
        return cobros
    resto = cobros[3:]
    montos = [c.monto_mxn for c in resto if c.monto_mxn is not None]
    agg = round(sum(montos), 2) if montos else None
    # Regla de septiembre 2026: la celda pinta el NETO (`_neto_cobro`) y la
    # 4.ª lo agrega igual; un cobro sin neto cuenta con su monto.
    agg_neto = None
    if any(c.neto_mxn is not None for c in resto):
        netos = [x for x in (_neto_cobro(c) for c in resto) if x is not None]
        agg_neto = round(sum(netos), 2)
    ultima = next((c.fecha for c in reversed(resto) if c.fecha), None)
    etiqueta = f"{_fecha(ultima)} (+{len(resto)})" if ultima else f"(+{len(resto)} cobros)"
    return [*cobros[:3], BalanceAvionCobro(fecha=etiqueta, monto_mxn=agg, neto_mxn=agg_neto)]


# ===== «Cómo se cobró»: NOTA de cada parcialidad (6-oct-2026, API 0.0.60) =====
# Pedido del cliente: «al lado de la columna STATUS, si ya se pagó, que venga
# la misma información de cómo se cobró, quién lo cobró y, si es posible, a
# qué cuenta»; y luego: «mejor, después de cada cobro venga la nota con los
# detalles, para no tener tantas columnas nuevas». SIN columnas nuevas: cada
# celda COBRO n MXN lleva una nota «fecha · monto · cómo se cobró» y la
# celda STATUS una nota resumen con todas las líneas. El «cómo se cobró» lo
# arma el API (`cobrado_con`, fuente única `etiquetaCobradoCon`); aquí solo
# se pinta. Un API ≤ 0.0.59 no lo manda: se compone con el método y la
# cuenta que ya mandaba (sin quién registró). Sin método, cuenta ni registro
# no hay nota — repetir fecha y monto no dice nada (y los libros de siempre
# no cambian).
# Regla de septiembre 2026 (6-oct-2026, API 0.0.65; pedido: «cuando hay una
# comisión de un banco, en la parte de total cobrado no refleja el monto real
# que entró a la cuenta»): con `neto_mxn` la celda pinta lo que ENTRÓ y su
# nota dice «Bruto $20,400.00 · comisión banco 5 % $1,020.00 · neto
# $19,380.00 · Transferencia → Scotiabank Pesos · Registró: Itzi». El API
# arma esa cadena dentro de `cobrado_con` (el % solo él lo sabe); si no la
# trae, aquí se compone con bruto, comisión y neto (`_bruto_neto`).
_ANCHO_NOTA_COBRO = 320
# Caracteres que caben por renglón en una nota de 320 de ancho (estimado,
# para que el alto no recorte las líneas largas del multi-avión).
_CHARS_RENGLON_NOTA_COBRO = 48


def _txt_nota(s: str | None) -> str | None:
    """Texto útil para la nota: en blanco o «—» (método sin dato) ⇒ None."""
    if not s:
        return None
    s = s.strip()
    return None if s in ("", "—", "-") else s


def _monto_nota(m: float | None) -> str | None:
    """12522.2 → «$12,522.20»; un reembolso (cobro negativo) → «-$500.00»."""
    if m is None:
        return None
    return f"-${-m:,.2f}" if m < 0 else f"${m:,.2f}"


def _cobrado_con(c: BalanceAvionCobro) -> str | None:
    """«Cómo se cobró» del cobro: `cobrado_con` del API tal cual. Sin él (API
    ≤ 0.0.59): método → cuenta, más «Registró: X» si llegara el nombre.
    None si no hay nada que decir."""
    armado = _txt_nota(c.cobrado_con)
    if armado:
        return armado
    metodo = _txt_nota(c.metodo_etiqueta) or _txt_nota(c.metodo)
    como = " → ".join(x for x in (metodo, _txt_nota(c.cuenta)) if x)
    registro = _txt_nota(c.registro)
    partes = [como] if como else []
    if registro:
        partes.append(f"Registró: {registro}")
    return " · ".join(partes) or None


# «Bruto $…» / «neto $…» dentro de `cobrado_con`: el API ya armó bruto ·
# comisión · neto (no se repite). Con el monto pegado: «Registró: Neto» (un
# apodo) no es la marca.
_MARCA_BRUTO = re.compile(r"\b(?:bruto|neto)s?:?\s*-?\$", re.IGNORECASE)


def _neto_cobro(c: BalanceAvionCobro) -> float | None:
    """Monto de la celda COBRO n MXN: `neto_mxn` (lo que entró a la cuenta,
    regla de septiembre 2026) o, sin él, `monto_mxn` como siempre."""
    return c.neto_mxn if c.neto_mxn is not None else c.monto_mxn


def _bruto_neto(c: BalanceAvionCobro) -> str | None:
    """«Bruto $20,400.00 · comisión banco $1,020.00 · neto $19,380.00» si el
    cobro trae un neto distinto del bruto; None si no (sin comisión o API
    ≤ 0.0.64)."""
    if c.neto_mxn is None or c.monto_mxn is None or abs(c.monto_mxn - c.neto_mxn) < 0.005:
        return None
    partes = [f"Bruto {_monto_nota(c.monto_mxn)}"]
    if c.comision_mxn is not None:
        partes.append(f"comisión banco {_monto_nota(c.comision_mxn)}")
    partes.append(f"neto {_monto_nota(c.neto_mxn)}")
    return " · ".join(partes)


def _linea_cobro(c: BalanceAvionCobro) -> str | None:
    """Una línea «15/09/2026 · $12,522.20 · Transferencia → Scotiabank Pesos
    · Itzi» (las partes que no vienen se omiten). Con neto distinto del bruto
    (regla de septiembre 2026) el monto suelto cede su lugar a «Bruto … ·
    comisión banco … · neto …»: el que el API armó dentro de `cobrado_con` o,
    si no lo trae, el de `_bruto_neto`."""
    como = _cobrado_con(c)
    if c.neto_mxn is not None and como and _MARCA_BRUTO.search(como):
        partes: tuple[str | None, ...] = (_fecha(c.fecha), como)
    else:
        monto = _bruto_neto(c) or _monto_nota(_neto_cobro(c))
        partes = (_fecha(c.fecha), monto, como)
    return " · ".join(x for x in partes if x) or None


def _cuenta_algo(c: BalanceAvionCobro) -> bool:
    """¿Hay algo que contar en la nota del cobro? Cómo se cobró (método,
    cuenta o quién lo registró) o, con la regla de septiembre 2026, que entró
    neto de una comisión."""
    return bool(_cobrado_con(c)) or _bruto_neto(c) is not None


def _nota_cobro(c: BalanceAvionCobro) -> str | None:
    """Nota de la celda COBRO n MXN de UNA parcialidad; None si el cobro no
    dice cómo se cobró (ni método, ni cuenta, ni quién lo registró) ni trae
    neto de comisión."""
    if not _cuenta_algo(c):
        return None
    return _linea_cobro(c)


def _nota_resumen_cobros(cobros: list[BalanceAvionCobro]) -> str | None:
    """Una línea por cobro (todas, en orden) para la nota de la celda STATUS
    y la de la 4.ª parcialidad agregada. None sin cobros o si NINGUNO dice
    cómo se cobró (ni trae neto de comisión)."""
    if not any(_cuenta_algo(c) for c in cobros):
        return None
    lineas = [linea for linea in (_linea_cobro(c) for c in cobros) if linea]
    return "\n".join(lineas) or None


def _notas_parcialidades(cobros: list[BalanceAvionCobro]) -> list[str | None]:
    """Nota de cada celda COBRO 1..4 MXN, en el mismo orden que `_cobros_a_4`:
    con más de 4 cobros la 4.ª agrega el resto y su nota lleva una línea por
    cobro agrupado."""
    if len(cobros) <= 4:
        return [_nota_cobro(c) for c in cobros]
    return [*(_nota_cobro(c) for c in cobros[:3]), _nota_resumen_cobros(cobros[3:])]


def _comentario_cobro(texto: str) -> Comment:
    """Nota de Excel con el tamaño de las de desglose (320 × ≥ 70) y alto
    según los renglones que envuelve."""
    lineas = texto.split("\n")
    renglones = sum(max(1, math.ceil(len(x) / _CHARS_RENGLON_NOTA_COBRO)) for x in lineas)
    com = Comment(texto, "VuelaTour")
    com.width = _ANCHO_NOTA_COBRO
    com.height = max(70, min(260, 24 + 16 * renglones))
    return com


# ===== COMISIONES a cargo del avión (6-oct-2026, API 0.0.65) =====
# Pedido del cliente: «la comisión del banco y vendedor se puede ir a la
# columna de (comisiones del vendedor) pero cambiar el nombre a "comisiones" y
# poner notas el tipo de comisión y si hay más de una. Y la comisión del
# vendedor también entra en el balance del avión, para que el monto real
# total cobrado ya sea después de cualquier comisión». Respuestas del 6-oct:
# la comisión del vendedor la absorbe el avión, lo cobrado al cliente por ese
# concepto se queda como ingreso de VuelaTour y la regla va desde septiembre
# 2026 (por fecha del vuelo: clave `comisiones_al_avion_desde` del API). Los
# números los calcula el API (fuente única `comisionesDelVuelo`, la misma del
# reparto a socios); aquí solo se pintan: la columna COMISIÓN VENDEDOR MXN
# pasa a COMISIONES MXN (misma posición) con la nota `comisiones_detalle`, y
# GANANCIA la resta cuando la fila la trae. Sin los campos nuevos (API ≤
# 0.0.64) el libro sale como siempre salvo los encabezados y sus textos.
# 7-oct-2026 (API 0.0.66, pedido: «estas comisiones se están duplicando en
# la general, ya que las tenemos en "Otros movimientos". Las comisiones deben
# aparecer en ese apartado pero en el individual de los avioncitos»): la
# columna vive SOLO en el libro INDIVIDUAL. La hoja de vuelos del Balance
# general (las dos variantes) va sin ella y antes de comisiones; sus notas
# (`_COMISIONES_FUERA_DE_LA_HOJA`, `_NOTA_COMISIONES_FLOTA`) dicen dónde
# quedaron: en el libro individual y en la cascada de la hoja 'balance'.
_NOTA_ENCABEZADO_COMISIONES = (
    "COMISIONES (regla de septiembre 2026, por fecha del vuelo) = lo que "
    "absorbe el avión: su parte de la comisión bancaria de los cobros del "
    "vuelo (en la misma proporción con que el cobro se prorratea al avión) + "
    "la comisión del vendedor como PROVISIÓN (lo cobrado al cliente por ese "
    "concepto, comisión + IVA). La nota de cada celda dice cuáles son. "
    "Vuelos anteriores: vacía (la comisión era de VuelaTour)."
)
# Hoja de vuelos del Balance GENERAL sin COMISIONES (7-oct-2026): las notas
# de encabezado de GANANCIA MXN (variante mensual) y de REMANENTE VENTA MENOS
# COMPRA (variante general) dicen que van antes de comisiones y dónde están.
_COMISIONES_FUERA_DE_LA_HOJA = (
    "Las comisiones que absorbe cada avión (regla de septiembre 2026, por "
    "fecha del vuelo: su parte de la comisión bancaria de sus cobros y la "
    "comisión del vendedor como PROVISIÓN) no van en esta hoja: están en la "
    "columna COMISIONES del libro individual de cada avión y ya van restadas "
    "en la cascada de cada avión de la hoja 'balance'."
)
_NOTA_ENCABEZADO_GANANCIA_FLOTA = (
    "GANANCIA (MXN y su equivalente en USD) = REMANENTE (VENTA − COSTO "
    "TOTAL), antes de comisiones. " + _COMISIONES_FUERA_DE_LA_HOJA
)
_NOTA_ENCABEZADO_REMANENTE = (
    "REMANENTE VENTA MENOS COMPRA = VENTA AVIÓN MXN − TOTAL PARA PROVEEDOR "
    "(PESOS), antes de comisiones. " + _COMISIONES_FUERA_DE_LA_HOJA
)
# Llave de la columna → nota de su encabezado en una hoja SIN COMISIONES.
_NOTAS_ENCABEZADO_SIN_COMISIONES = {
    "ganancia_mxn": _NOTA_ENCABEZADO_GANANCIA_FLOTA,
    "cph_remanente_mxn": _NOTA_ENCABEZADO_REMANENTE,
}
# Nota al pie de la hoja de vuelos del libro INDIVIDUAL: le sigue el
# fragmento del pago al vendedor (`_MAESTRA_PAGO_VENDEDOR_PROVISION`) y la
# frase de GANANCIA.
_NOTA_COMISIONES = (
    "COMISIONES MXN (regla de septiembre 2026, por fecha del vuelo) = lo que "
    "absorbe el avión: su parte de la comisión bancaria de los cobros y la "
    "comisión del vendedor como PROVISIÓN (lo cobrado al cliente por ese "
    "concepto, comisión + IVA); la nota de cada celda dice cuáles son y en los "
    "vuelos anteriores va vacía. Lo cobrado al cliente por la comisión del "
    "vendedor sigue siendo ingreso de VuelaTour y su pago al vendedor sale de "
    "VuelaTour; ambos viven en 'otros movimientos' "
)
_NOTA_GANANCIA = " GANANCIA de la fila = REMANENTE (VENTA − COSTO TOTAL) − COMISIONES."
# Su versión en el Balance GENERAL (las dos variantes, 7-oct-2026): la hoja no
# lleva COMISIONES; el pago al vendedor lo explica la leyenda de 'otros
# movimientos'. Le sigue la frase de GANANCIA (mensual) o de REMANENTE
# (general, `_CPH_NOTA_REMANENTE`).
_NOTA_COMISIONES_FLOTA = (
    "COMISIONES (regla de septiembre 2026, por fecha del vuelo): las que "
    "absorbe cada avión —su parte de la comisión bancaria de sus cobros y la "
    "comisión del vendedor como PROVISIÓN— no van en esta hoja: están en la "
    "columna COMISIONES del libro individual de cada avión (Reportes › Balance "
    "por avión) y ya van restadas en la cascada de cada avión de la hoja "
    "'balance'. Lo cobrado al cliente por la comisión del vendedor es ingreso "
    "de VuelaTour: ese ingreso, el pago al vendedor y la parte de VuelaTour de "
    "la comisión bancaria viven en 'otros movimientos' (su leyenda explica "
    "cada línea)."
)
_NOTA_GANANCIA_FLOTA = (
    " GANANCIA de la fila = REMANENTE (VENTA − COSTO TOTAL), antes de comisiones."
)
# Cómo se nombra la regla en los textos de arriba (y en `_RESUMEN_COMISIONES`,
# `_NOTA_COBRADO_REAL` y `_NOTA_COBRANZA_NETO`) cuando el API no manda su
# etiqueta: la vigencia por defecto. `_con_regla` pone la del API.
_DESDE_FIJO = "desde septiembre 2026"


def _con_regla(texto: str, regla: str | None) -> str:
    """El texto con el nombre de la regla que manda el API (`regla_comisiones`,
    «regla sep-2026»: sale de la vigencia CONFIGURADA, revisión 6-oct-2026)
    en lugar del mes fijo: «regla de septiembre 2026» ⇒ la etiqueta y «desde
    septiembre 2026» ⇒ «con la <etiqueta>». Sin etiqueta (API 0.0.65), el
    texto tal cual: libro byte-idéntico."""
    if not regla:
        return texto
    return (
        texto.replace(REGLA_COMISIONES_RESPALDO, regla)
        .replace(_DESDE_FIJO, f"con la {regla}")
        .replace(_DESDE_FIJO.capitalize(), f"Con la {regla}")
    )


def _regla_del_general(req: BalanceGeneralRequest) -> str | None:
    """Etiqueta de la regla en el Balance general: la del request o, si no
    la trae, la del `consolidado`."""
    if req.regla_comisiones:
        return req.regla_comisiones
    return req.consolidado.regla_comisiones if req.consolidado is not None else None


def _avion_cubre_al_vendedor(req: BalanceGeneralRequest) -> bool:
    """¿El payload es del API 0.0.66 (7-oct-2026)? Desde ese API la línea de
    INGRESO «comisión del vendedor a cargo del avión …» de 'otros
    movimientos' ya no viaja (duplicaba la comisión, pedido del cliente) y,
    en los vuelos COMPLETADOS de la regla, el pago al vendedor dice
    «cubierto por el avión» (solo resta lo que el pago real exceda la
    provisión). La señal es el campo que ese mismo API agrega a la cascada de
    los aviones, `balance.comisiones_usd`: sin él (API 0.0.65), las leyendas
    de antes."""
    return any(a.balance.comisiones_usd is not None for a in req.aviones)


def _regla_vigente(cons: BalanceAvionRequest) -> str | None:
    """Nombre de la regla si el libro trae vuelos de ella, o None. El API
    manda `comisiones_mxn` en los totales SOLO si alguna fila de avión cae en
    la vigencia (aunque sus comisiones sean 0); antes de ella, o con un API
    ≤ 0.0.64, no viene."""
    if cons.totales.comisiones_mxn is None:
        return None
    return cons.regla_comisiones or REGLA_COMISIONES_RESPALDO


# Aclaraciones cuando el periodo trae vuelos de la regla (el API manda
# COMISIONES en los totales del consolidado: `_regla_vigente`). Con el API
# 0.0.65 la pestaña 'otros movimientos' gana, por avión, una línea de INGRESO
# «comisión del vendedor a cargo del avión …» (la provisión que el avión le
# paga a VuelaTour, no el cliente) y su comisión bancaria es solo la parte de
# VuelaTour; la leyenda de la fila 2 y la base del bloque VUELATOUR (empresa)
# lo dicen. Sin vuelos de la regla, los textos de siempre. El API 0.0.66
# cambia esa pestaña: ver `_OM_REGLA_COMISIONES_CUBRE_AVION`.
_OM_REGLA_COMISIONES = (
    " Con la {regla} (por fecha del vuelo) la fila del vuelo trae además una "
    "línea de INGRESO «comisión del vendedor a cargo del avión …»: la "
    "provisión que cada avión le paga a VuelaTour (no la paga el cliente), y "
    "la comisión bancaria de aquí es solo la parte de VuelaTour (la del avión "
    "resta en la columna COMISIONES de su libro individual)."
)
_EMPRESA_REGLA_COMISIONES = (
    " Con la {regla} (por fecha del vuelo), los ingresos propios incluyen "
    "además la provisión de la comisión del vendedor que paga cada avión y, "
    "de la comisión bancaria, los egresos solo traen la parte de VuelaTour."
)
# API 0.0.66 (7-oct-2026; pedido: «en el reporte mensual y en el balance
# general, en "otros movimientos", en el ingreso estás duplicando la
# comisión»): la línea de INGRESO «… a cargo del avión» ya no viaja (la fila
# del vuelo trae lo cobrado al cliente UNA vez) y, en los vuelos COMPLETADOS
# de la regla, el pago al vendedor dice «cubierto por el avión»: solo resta lo
# que el pago real exceda la provisión que ya carga el avión. Las dos
# leyendas (y el RESUMEN, tras su fragmento del pago apareado) lo dicen
# cuando el payload es de ese API (`_avion_cubre_al_vendedor`); con uno
# anterior, las de arriba.
_OM_REGLA_COMISIONES_CUBRE_AVION = (
    " Con la {regla} (por fecha del vuelo), en los vuelos COMPLETADOS la "
    "comisión del vendedor la cubre el avión con su PROVISIÓN: el pago al "
    "vendedor dice «cubierto por el avión» y aquí solo resta lo que el pago "
    "real exceda esa provisión; la comisión bancaria de aquí es solo la parte "
    "de VuelaTour. Lo que absorbe cada avión (la provisión y su parte de la "
    "comisión bancaria) resta en la columna COMISIONES de su libro individual "
    "y en la cascada de la hoja 'balance'."
)
_EMPRESA_REGLA_COMISIONES_CUBRE_AVION = (
    " Con la {regla} (por fecha del vuelo), en los vuelos completados el pago "
    "al vendedor lo cubre cada avión con su provisión (los egresos propios "
    "solo traen lo que el pago real la exceda) y, de la comisión bancaria, los "
    "egresos solo traen la parte de VuelaTour."
)
_RESUMEN_REGLA_CUBRE_AVION = (
    " Con la {regla} (por fecha del vuelo), en los vuelos completados ese pago "
    "lo cubre el avión con su provisión: en 'otros movimientos' solo queda lo "
    "que el pago real la exceda."
)
# «N conceptos» que encabeza la nota cuando hay más de una comisión (si el API
# ya lo trae en el detalle, no se repite).
_ENCABEZADO_CONCEPTOS = re.compile(r"\b\d+\s+conceptos\b", re.IGNORECASE)


def _nota_comisiones(v: BalanceAvionVuelo) -> str | None:
    """Nota de la celda COMISIONES MXN: las líneas que arma el API
    (`comisiones_detalle`: «Comisión bancaria · Transferencia 5 % · $1,020.00
    (parte del avión 100 %)», «Comisión vendedor (Pablo Canales) · $2,360.00
    · provisión = cotizado + IVA») con «N conceptos» arriba cuando hay más de
    una. Sin detalle, una línea por cada parte que mande el API (banco /
    vendedor; la provisión sin «+ IVA»: no toda cotización lo cobra). None si
    no hay nada que decir."""
    lineas = list(v.comisiones_detalle)
    if not lineas:
        partes = (
            ("Comisión bancaria (parte del avión)", v.comision_banco_avion_mxn),
            ("Comisión del vendedor (provisión)", v.comision_vendedor_prov_mxn),
        )
        lineas = [f"{nombre} · {_monto_nota(m)}" for nombre, m in partes if _monto_no_cero(m)]
    if not lineas:
        return None
    if len(lineas) > 1 and not any(_ENCABEZADO_CONCEPTOS.search(x) for x in lineas):
        lineas = [f"{len(lineas)} conceptos", *lineas]
    return "\n".join(lineas)


def _pct(factor: float | None) -> str:
    """0.5 → '50 %' (hasta 2 decimales, sin ceros de más)."""
    return "—" if factor is None else f"{round(factor * 100, 2):g} %"


def _nota_participacion(v: BalanceAvionVuelo) -> str:
    """Nota de la celda VENTA AVIÓN USD de una fila de vuelo MULTI-AVIÓN
    (regla B, 28-ago-2026): cuánto le toca a esta matrícula y con qué base."""
    base = _FUENTE_PARTICIPACION.get(v.participacion_fuente or "")
    return (
        f"Vuelo multi-avión: esta fila trae el {_pct(v.participacion)} de la "
        "venta del avión (repartida entre las matrículas del vuelo"
        + (f" por {base}" if base else "")
        + "); horas cobradas, cobros y por cobrar van en la misma "
        "proporción. Los gastos de la fila son los de ESTE avión (por "
        "tramo), sin repartir. TUAs/extras/pernocta/comisión del vendedor "
        "son de VuelaTour y no se reparten."
    )


def _nota_tc_oficial(v: BalanceAvionVuelo) -> str:
    """Nota de la celda TC VENTA cuando la cotización no traía tipo de cambio
    y el API usó el oficial de referencia del día (fuente y fecha del dato,
    si el API las manda)."""
    detalle: list[str] = []
    if v.tc_venta_oficial_fuente:
        fuente = _TC_FUENTE_LABEL.get(
            v.tc_venta_oficial_fuente, v.tc_venta_oficial_fuente
        )
        detalle.append(f"fuente {fuente}")
    if v.tc_venta_oficial_fecha:
        detalle.append(f"dato del {v.tc_venta_oficial_fecha}")
    return (
        "TC oficial de referencia del día de la cotización"
        + (f" ({', '.join(detalle)})" if detalle else "")
        + ": la cotización no traía tipo de cambio."
    )


# Comisión del vendedor como GASTO (28-sep-2026, API 0.0.39, invariante 31
# del API): el gasto real «Comisión del vendedor» ligado al vuelo REEMPLAZA a
# la PROVISIÓN en 'otros movimientos' (concepto y nota llegan hechos del API).
# Las leyendas del Balance GENERAL que hablan de la provisión (fila 2 de
# 'otros movimientos' y nota del RESUMEN) cambian SOLO con
# `otros_movimientos.hay_pago_vendedor_real` (el API la manda solo si algún
# vuelo del periodo tiene pago real). Sin ella el texto de siempre
# —verdadero mientras no haya gasto real—. Hasta el 7-oct-2026 también la
# nota de la hoja de vuelos del general; desde entonces esa nota no habla
# del pago al vendedor (remite a la leyenda de 'otros movimientos',
# `_NOTA_COMISIONES_FLOTA`). El libro INDIVIDUAL no trae 'otros
# movimientos': su nota no cambia nunca (`_MAESTRA_PAGO_VENDEDOR_PROVISION`).
_OM_PAGO_VENDEDOR_PROVISION = (
    "el pago de la comisión al vendedor va apareado en la "
    "misma fila como PROVISIÓN a la fecha del vuelo mientras no exista el "
    "gasto real (lo dice la nota de la celda)."
)
_OM_PAGO_VENDEDOR_REAL = (
    "el pago de la comisión al vendedor va apareado en la misma fila: el "
    "GASTO REAL cuando se captura en Gastos con la categoría «Comisión del "
    "vendedor» ligado al vuelo (con «faltan $…» o «excede $…» si no cuadra "
    "con lo cobrado) o, mientras no se capture, una PROVISIÓN por el mismo "
    "monto a la fecha del vuelo (lo dice la nota de la celda). Ese pago NO se "
    "captura como «Otros gastos VuelaTour»: quedaría duplicado en la hoja "
    "'otros gastos'."
)
# Nota del RESUMEN sobre COMISIONES (6-oct-2026, API 0.0.65; antes
# «COMISIONES VENDEDOR va vacía a propósito»): le sigue el fragmento de la
# bandera (`_RESUMEN_PAGO_VENDEDOR_*`) y, con el API 0.0.66,
# `_RESUMEN_REGLA_CUBRE_AVION`. Desde el 7-oct-2026 la columna COMISIONES
# es la del libro INDIVIDUAL de cada avión: la hoja de vuelos de este libro ya
# no la lleva (el RESUMEN conserva la suya).
_RESUMEN_COMISIONES = (
    "COMISIONES = las de la columna COMISIONES del libro individual de cada "
    "avión (regla de septiembre 2026, por fecha del vuelo: su parte de la "
    "comisión bancaria de los cobros y la comisión del vendedor como "
    "PROVISIÓN; en vuelos anteriores no hay); en este libro no van en la hoja "
    "de vuelos (ahí la GANANCIA va antes de comisiones): restan aquí y en la "
    "cascada de cada avión de la hoja 'balance'. Lo cobrado al cliente por la "
    "comisión del vendedor sigue siendo INGRESO de VuelaTour; ese ingreso y el "
    "pago al vendedor viven en 'otros movimientos' "
)
_RESUMEN_PAGO_VENDEDOR_PROVISION = (
    "(ingreso cobrado y pago apareado como PROVISIÓN a la fecha del vuelo)."
)
_RESUMEN_PAGO_VENDEDOR_REAL = (
    "(ingreso cobrado y pago apareado: el gasto real «Comisión del vendedor» "
    "o, mientras no se capture, la PROVISIÓN a la fecha del vuelo)."
)
_MAESTRA_PAGO_VENDEDOR_PROVISION = (
    "(el pago apareado como PROVISIÓN a la fecha del vuelo)."
)


# Nota al pie de la hoja maestra (individual y general) que explica la
# columna FACTURA VUELATOUR (30-sep-2026, API 0.0.45). Se pinta SIEMPRE: la
# columna existe aunque el API no mande el dato.
_NOTA_FACTURA_VUELATOUR = (
    "FACTURA VUELATOUR = folio de la factura del servicio emitida al cliente "
    "(timbrada, registrada en Facturas emitidas o capturada en el vuelo); sin "
    "folio registrado dice el estatus del vuelo («Facturado» / «Factura "
    "elaborada y enviada») y vacía = sin factura. Es la factura del VUELO "
    "completo: en un vuelo multi-avión todas sus filas traen la misma."
)


# Notas al pie de la hoja maestra que la variante general repite tal cual
# (el texto es el de siempre; solo se nombró para no copiarlo).
_NOTA_VENTA_AVION = (
    "* VENTA AVIÓN de cada fila = tiempo de vuelo (tarifa × horas "
    "cobradas) + ajuste/descuento + su IVA proporcional. IVA VENTA "
    "AVIÓN (USD y MXN) es SOLO el IVA proporcional de la venta del "
    "avión — el IVA de TUAs/extras/pernocta/comisión viaja con ellos. "
    "TUAs, extras, viáticos de pernocta y la COMISIÓN DEL VENDEDOR NO "
    "son venta del avión: son ingreso de VuelaTour (pestaña 'otros "
    "movimientos' del Balance general). Sin cotización: horas × tarifa."
)
_NOTA_COBRADO_REAL_INICIO = (
    "**** COBRADO REAL = Σ de los depósitos tal cual entraron a la cuenta "
    "(COBRO 1..4): desde septiembre 2026 cada COBRO va NETO de la comisión "
    "del banco (la nota de la celda trae bruto, comisión y neto). COBRADO "
    "AVIÓN = cobros BRUTOS × (venta avión ÷ total cotización), ANTES de "
    "comisiones: la parte de los cobros que corresponde a TUAs/extras/"
    "pernocta/comisión del vendedor es de VuelaTour (ver 'otros "
    "movimientos') y las comisiones que absorbe el avión restan aparte, en "
)
_NOTA_COBRADO_REAL_FIN = (
    ". POR COBRAR es la parte del avión. COBRADO AVIÓN y POR "
    "COBRAR van al TC de venta (un depósito en pesos con su propio TC se "
    "pasa a USD con ese TC y se re-expresa al TC de venta), así que COBRADO "
    "AVIÓN puede diferir de COBRADO REAL sin que falte dinero."
)
_NOTA_COBRADO_REAL = _NOTA_COBRADO_REAL_INICIO + "COMISIONES" + _NOTA_COBRADO_REAL_FIN
# Balance GENERAL (7-oct-2026): su hoja de vuelos ya no lleva COMISIONES.
_COMISIONES_EN_EL_INDIVIDUAL = (
    "la columna COMISIONES del libro individual de cada avión y en la cascada de la hoja 'balance'"
)
_NOTA_COBRADO_REAL_FLOTA = (
    _NOTA_COBRADO_REAL_INICIO + _COMISIONES_EN_EL_INDIVIDUAL + _NOTA_COBRADO_REAL_FIN
)
_NOTA_ESTATUS_COBRO = (
    "El ESTATUS DE COBRO por vuelo (cuánto se cobró, con qué fechas y "
    "métodos, y cuánto falta) está al frente en la hoja 'cobranza' — el "
    "bloque STATUS DE COBROS del final de esta hoja trae lo mismo en "
    "columnas."
)
_NOTA_COMPARTIDO = (
    "*** Filas 'COMPARTIDO' / con porcentaje en la RUTA (vuelo "
    "MULTI-AVIÓN, regla 28-ago-2026): la fila trae la parte "
    "proporcional de la VENTA del avión (repartida entre las matrículas "
    "del vuelo en partes iguales por tramo vendido; los ferries/tramos "
    "operativos no reparten) y en la misma proporción sus "
    "horas cobradas, cobros y por cobrar; sus GASTOS son los de su "
    "avión (los del tramo que voló), sin repartir. TUAs/extras/pernocta/"
    "comisión del vendedor no son de ningún avión y no se reparten. La "
    "nota de la celda VENTA AVIÓN USD dice el porcentaje y la base."
)
_NOTA_TACOS_AMBAR = (
    "TACO INICIO / TACO FINAL en ámbar = salto en la cadena de "
    "tacómetros (el valor esperado está en la nota de la celda) y/u "
    "OBSERVACIÓN del equipo capturada en Tacómetros en vivo — pasa el "
    "cursor por la celda para leer el comentario (quién y cuándo). "
    "TIEMPO VUELO en ámbar = salto entre tramos DEL MISMO vuelo (el "
    "tramo culpable está en la nota) — mismo amarillo que el panel."
)


# Notas al pie de la hoja maestra sobre los TRASLADOS al pasajero (pie **).
# La versión «de siempre» habla solo del TUA; la ampliada (1-oct-2026, API
# 0.0.47, pedido de Ale sobre el #192: la extensión de horario de Chetumal
# caía en OPERACIONES) suma la extensión y/o antelación de horario y se pinta
# SOLO cuando alguna fila o los totales traen `extension_pagada_mxn` ≠ 0
# (`_hay_extension_pagada`). Sin la llave, el libro es byte-idéntico.
_NOTA_COMBUSTIBLE_FILAS_SIN_GAS = (
    "El COMBUSTIBLE ya no va por vuelo (26-ago-2026): se controla por "
    "avión y por MES en la hoja 'combustible'"
)
_NOTA_COMBUSTIBLE_COLA_TUA = (
    " (litros y $/L incluidos) y "
    "resta una sola vez en la hoja 'balance'. Por eso COSTO TOTAL, COSTO "
    "X HORA, REMANENTE y GANANCIA de las filas van SIN combustible (y el "
    "TUA pagado tampoco resta, ver **): la utilidad real del periodo se "
    "lee en la hoja 'balance', no en la fila. IVA PAGADO y DIF. IVA "
    "también derivan solo de los costos de la fila (sin gas ni TUA)."
)
_NOTA_COMBUSTIBLE_COLA_EXTENSION = (
    " (litros y $/L incluidos) y "
    "resta una sola vez en la hoja 'balance'. Por eso COSTO TOTAL, COSTO "
    "X HORA, REMANENTE y GANANCIA de las filas van SIN combustible (ni el "
    "TUA pagado ni la extensión de horario restan, ver **): la utilidad "
    "real del periodo se lee en la hoja 'balance', no en la fila. IVA "
    "PAGADO y DIF. IVA también derivan solo de los costos de la fila (sin "
    "gas, TUA ni extensión de horario)."
)
_NOTA_TRASLADOS_TUA = (
    "** El TUA pagado al aeropuerto NO es costo del avión ni resta en "
    "ningún lado de este libro: queda solo como nota en la celda "
    "OPERACIONES ('TUA $x**'); cobro y pago del TUA viven en 'otros "
    "movimientos' del Balance general. Los servicios FBO sí son costo "
    "(columna OTROS)."
)
_NOTA_TRASLADOS_EXTENSION = (
    "** TUA y extensión de horario (extensión y/o antelación de horario del "
    "aeropuerto) son traslados al pasajero: NO son costo del avión ni "
    "restan en ningún lado de este libro; quedan solo como nota en la celda "
    "OPERACIONES ('TUA $x**', 'Extensión de horario (IVA incluido) $x**'); "
    "lo cobrado y lo pagado viven en 'otros movimientos' del Balance "
    "general. Los servicios FBO sí son costo (columna OTROS)."
)


def _monto_no_cero(x: float | None) -> bool:
    return x is not None and x != 0


def _hay_extension_pagada(req: BalanceAvionRequest) -> bool:
    """True si alguna fila o los totales traen `extension_pagada_mxn` ≠ 0
    (API 0.0.47). El API solo manda la llave con monto ≠ 0; un None o un 0
    explícitos cuentan como «no vino» — mismo criterio que el renglón del
    TUA pagado en el bloque de totales."""
    return _monto_no_cero(req.totales.extension_pagada_mxn) or any(
        _monto_no_cero(v.extension_pagada_mxn) for v in req.vuelos
    )


def _hay_pago_vendedor_real(om: BalanceHojaOtrosMovimientos | None) -> bool:
    """True solo si el API marcó la hoja con `hay_pago_vendedor_real: true`
    (None/False/ausente ⇒ leyendas de siempre)."""
    return om is not None and om.hay_pago_vendedor_real is True


# ===== Textos y notas de la variante «Balance general» (6-oct-2026) =====
# Las notas de siempre nombran columnas que la variante ya no tiene
# (OPERACIONES, GANANCIA, COSTO TOTAL USD, COSTO X HORA USD…); estas dicen lo
# mismo con las columnas del bloque COSTO POR HORA.
_CPH_NOTA_AFAC = (
    "La tarifa AFAC es POR AVIÓN y este Balance general ya no lleva la "
    "columna PERMISO AFAC: el permiso (provisión) de cada vuelo va dentro de "
    "TOTAL PARA PROVEEDOR (PESOS) — su monto está en la nota de esa celda. La "
    "fórmula con la tarifa de cada avión está en su libro individual "
    "(Reportes › Balance por avión)."
)
_CPH_NOTA_FACTOR_IVA = (
    "IVA de los costos (16 %): COSTO X HORA (DLLS S/IVA) = TOTAL PARA "
    "PROVEEDOR (PESOS) ÷ TIPO CAMBIO ÷ TIEMPO CALZOS/HOBS ÷ este factor e IVA "
    "TOTAL PAGADO (DLLS) = TOTAL PARA PROVEEDOR (DLLS) − TOTAL PARA PROVEEDOR "
    "(DLLS) ÷ este factor."
)
_CPH_NOTA_BLOQUE = (
    "BALANCE GENERAL · COSTO POR HORA: COSTO X HORA (DLLS S/IVA) = TOTAL PARA "
    "PROVEEDOR (PESOS) ÷ TIPO CAMBIO ÷ TIEMPO CALZOS/HOBS ÷ 1.16 — el total de "
    "todos los gastos del vuelo, entre el tiempo volado, entre el tipo de "
    "cambio, entre 1.16 para sacar el subtotal. COSTO HR MÁS IVA = TOTAL PARA "
    "PROVEEDOR (DLLS) ÷ TIEMPO e IVA X HR = la diferencia de las dos. TOTAL "
    "PARA PROVEEDOR (PESOS) = COSTO TOTAL del vuelo (operación + piloto + "
    "otros + permiso AFAC, sin combustible): el desglose está en la nota de "
    "la celda y, por columna, en el Balance mensual. TIPO CAMBIO = el de los "
    "costos del vuelo."
)
# 7-oct-2026: la definición del cliente, antes de comisiones (la hoja ya no
# las lleva; `_NOTA_COMISIONES_FLOTA` dice dónde están).
_CPH_NOTA_REMANENTE = (
    " REMANENTE VENTA MENOS COMPRA = VENTA AVIÓN MXN − TOTAL PARA PROVEEDOR "
    "(PESOS), antes de comisiones."
)
_CPH_NOTA_TOTALES = (
    "Fila TOTALES: TIPO CAMBIO y COSTO HR MÁS IVA son "
    "PROMEDIOS (promedio de los promedios de cada avión); COSTO X HORA (DLLS "
    "S/IVA) = ese COSTO HR MÁS IVA ÷ 1.16 e IVA X HR = la diferencia de los "
    "dos; las demás son sumas (TOTAL PARA PROVEEDOR (DLLS), IVA TOTAL PAGADO y "
    "TOTAL PAGADO S/IVA suman las filas tal como se ven, a centavos)."
)
_CPH_COMBUSTIBLE_COLA_TUA = (
    " (litros y $/L incluidos) y resta una sola vez en la hoja 'balance'. Por "
    "eso TOTAL PARA PROVEEDOR, COSTO X HORA y REMANENTE de las "
    "filas van SIN combustible (y el TUA pagado tampoco resta, ver **): la "
    "utilidad real del periodo se lee en la hoja 'balance', no en la fila. "
    "IVA TOTAL PAGADO también deriva solo de los costos de la fila (sin gas "
    "ni TUA)."
)
_CPH_COMBUSTIBLE_COLA_EXTENSION = (
    " (litros y $/L incluidos) y resta una sola vez en la hoja 'balance'. Por "
    "eso TOTAL PARA PROVEEDOR, COSTO X HORA y REMANENTE de las "
    "filas van SIN combustible (ni el TUA pagado ni la extensión de horario "
    "restan, ver **): la utilidad real del periodo se lee en la hoja "
    "'balance', no en la fila. IVA TOTAL PAGADO también deriva solo de los "
    "costos de la fila (sin gas, TUA ni extensión de horario)."
)
_CPH_TRASLADOS_TUA = (
    "** El TUA pagado al aeropuerto NO es costo del avión ni resta en ningún "
    "lado de este libro: queda solo como nota en el desglose de TOTAL PARA "
    "PROVEEDOR (PESOS) (parte Operación: 'TUA $x**'); cobro y pago del TUA "
    "viven en 'otros movimientos'. Los servicios FBO sí son costo (parte "
    "Otros del desglose)."
)
_CPH_TRASLADOS_EXTENSION = (
    "** TUA y extensión de horario (extensión y/o antelación de horario del "
    "aeropuerto) son traslados al pasajero: NO son costo del avión ni restan "
    "en ningún lado de este libro; quedan solo como nota en el desglose de "
    "TOTAL PARA PROVEEDOR (PESOS) (parte Operación: 'TUA $x**', 'Extensión de "
    "horario (IVA incluido) $x**'); lo cobrado y lo pagado viven en 'otros "
    "movimientos'. Los servicios FBO sí son costo (parte Otros del desglose)."
)
# 8-oct-2026: el general ya no lleva TACO INICIO / TACO FINAL ni TIEMPO VUELO
# HR; la señal del salto interno vive en TIEMPO CALZOS/HOBS (HR).
_CPH_NOTA_TACOS_AMBAR = (
    "TIEMPO CALZOS/HOBS (HR) en ámbar = salto entre tramos DEL MISMO vuelo "
    "(el tramo culpable está en la nota) — mismo amarillo que el panel. Los "
    "tacómetros inicial y final de cada vuelo (y su salto en la cadena) "
    "están en el Balance mensual."
)
_CPH_RENGLON_TUA = (
    "TUA pagado del periodo (solo nota en el desglose de TOTAL PARA "
    "PROVEEDOR, no resta en este libro):"
)
_CPH_RENGLON_EXTENSION = (
    "Extensión de horario pagada del periodo (solo nota en el desglose de "
    "TOTAL PARA PROVEEDOR, no resta en este libro):"
)
# «Calculado por el sistema» (regla del contrato del 6-oct-2026): una celda
# cuya fórmula de siempre cita columnas que la variante no tiene va como
# VALOR del API con esta nota (en la hoja maestra, en el encabezado de su
# columna; en otra hoja, en la celda).
_NOTA_CALCULADO_POR_SISTEMA = {
    "cph_proveedor_mxn": (
        "Calculado por el sistema: TOTAL PARA PROVEEDOR (PESOS) = COSTO TOTAL "
        "del vuelo = operación + piloto + otros + permiso AFAC (provisión), sin "
        "combustible. Esas columnas no van en este Balance general: el "
        "desglose de cada vuelo está en la nota de su celda y en el Balance "
        "mensual la columna COSTO TOTAL MXN es la fórmula que las suma."
    ),
    "costo_total_mxn": (
        "Calculado por el sistema: COSTO TOTAL = operación + piloto + otros + "
        "permiso AFAC (provisión) del vuelo, sin combustible. Esas columnas no "
        "van en este Balance general: el desglose de cada vuelo está en la "
        "nota de TOTAL PARA PROVEEDOR (PESOS) y en el Balance mensual esta "
        "celda es la fórmula que las suma."
    ),
}
_NOTA_CALCULADO_POR_SISTEMA_GENERICA = (
    "Calculado por el sistema: la fórmula de esta columna cita columnas que "
    "no van en este Balance general (están en el Balance mensual)."
)
_NOTA_CALCULADO_POR_SISTEMA_CELDA = (
    "Calculado por el sistema: la columna de la hoja de vuelos que esta "
    "fórmula citaría no va en este Balance general (está en el Balance "
    "mensual)."
)
# Partes del COSTO TOTAL de un vuelo, en el orden de sus columnas del Balance
# mensual: nombre en la nota, monto del API y líneas de detalle de su nota.
_PARTES_COSTO_TOTAL = (
    ("Operación", "op_mxn", "op_detalle"),
    ("Piloto", "piloto_mxn", "piloto_detalle"),
    ("Otros", "otros_mxn", "otros_detalle"),
    ("Permiso AFAC", "permiso_afac_mxn", None),
)
# Caracteres por renglón de la nota del desglose (340 de ancho), para el alto.
_CHARS_RENGLON_DESGLOSE = 52


def _nota_desglose_costo(v: BalanceAvionVuelo) -> str | None:
    """Nota de TOTAL PARA PROVEEDOR (PESOS) en la variante general: arriba
    «Operación $x · Piloto $y · Otros $z · Permiso AFAC $w» (los montos que
    el API manda; None se omite) y debajo, por parte, las MISMAS líneas de
    las notas de OPERACIONES / PILOTO / OTROS del Balance mensual. Aquí no se
    suma nada. None si el vuelo no trae costo ni detalle (sin nota)."""
    partes = [
        (nombre, getattr(v, attr), list(getattr(v, det)) if det else [])
        for nombre, attr, det in _PARTES_COSTO_TOTAL
    ]
    con_monto = [(nombre, m) for nombre, m, _lineas in partes if m is not None]
    if not any(m for _n, m in con_monto) and not any(lineas for *_x, lineas in partes):
        return None
    bloques = []
    if con_monto:
        bloques.append(" · ".join(f"{nombre} {_monto_nota(m)}" for nombre, m in con_monto))
    for nombre, _m, lineas in partes:
        if lineas:
            bloques.append(f"{nombre}:\n" + "\n".join(lineas))
    return "\n\n".join(bloques)


def _comentario_desglose(texto: str) -> Comment:
    """Nota del desglose (340 de ancho, como las de OPERACIONES / PILOTO /
    OTROS del Balance mensual) con el alto según los renglones que envuelve."""
    renglones = sum(max(1, math.ceil(len(x) / _CHARS_RENGLON_DESGLOSE)) for x in texto.split("\n"))
    com = Comment(texto, "VuelaTour")
    com.width = 340
    com.height = max(70, min(300, 24 + 16 * renglones))
    return com


def _notas_pie_costo_por_hora(req: BalanceAvionRequest, hay_extension: bool) -> tuple[str, ...]:
    """Notas al pie de la hoja maestra en la variante general: las de siempre
    que siguen siendo verdad tal cual y, en lugar de las que nombran
    columnas que la variante ya no tiene, su versión con las columnas del
    bloque COSTO POR HORA, en el mismo orden; la del bloque nuevo va después
    de la de VENTA. Las de COMISIONES y COBRADO REAL son las del Balance
    general (7-oct-2026): la hoja no lleva COMISIONES y REMANENTE va antes de
    ellas."""
    return (
        _NOTA_VENTA_AVION,
        _CPH_NOTA_BLOQUE,
        _con_regla(_NOTA_COMISIONES_FLOTA, req.regla_comisiones) + _CPH_NOTA_REMANENTE,
        _con_regla(_NOTA_COBRADO_REAL_FLOTA, req.regla_comisiones),
        _CPH_NOTA_TOTALES,
        _NOTA_ESTATUS_COBRO,
        _NOTA_FACTURA_VUELATOUR,
        _NOTA_COMBUSTIBLE_FILAS_SIN_GAS
        + " del libro individual de cada avión"
        + (_CPH_COMBUSTIBLE_COLA_EXTENSION if hay_extension else _CPH_COMBUSTIBLE_COLA_TUA),
        _CPH_TRASLADOS_EXTENSION if hay_extension else _CPH_TRASLADOS_TUA,
        _NOTA_COMPARTIDO,
        _CPH_NOTA_TACOS_AMBAR,
    )


def _constantes_maestra(
    ws: Worksheet, req: BalanceAvionRequest, general: bool, *, costo_por_hora: bool = False
) -> bool:
    """Constantes del encabezado (5-oct-2026), en la fila 1 de CLAVE..ESTADO:
    «Permiso AFAC USD/hr» en B1 y «Factor IVA costos» (1.16) en D1. Devuelve
    True si hay tarifa AFAC (la columna PERMISO AFAC lleva fórmula). En el
    GENERAL la tarifa es por avión: celda vacía y la columna como valor.
    `costo_por_hora` (variante general, 6-oct-2026): las dos constantes se
    conservan (COSTO X HORA cita $D$1) y sus notas hablan de sus columnas."""
    afac = None if general else req.permiso_afac_usd_hr
    hay_afac = afac is not None and afac > 0
    for col, texto in ((1, "Permiso AFAC USD/hr"), (3, "Factor IVA costos")):
        c = ws.cell(row=1, column=col, value=texto)
        c.font = Font(bold=True, color=NAVY, size=8)
        c.alignment = Alignment(horizontal="right", vertical="center")
    ca = _num(ws, 1, 2, afac if hay_afac else None, MONEY, bold=True)
    ca.fill = PatternFill("solid", fgColor=FILL_COSTOS)
    ca.border = _border
    if hay_afac:
        texto_afac = (
            "Aportación AFAC USD/hr de la ficha del avión. PERMISO AFAC "
            "(PROVISIÓN) de cada vuelo = ROUND(esta tarifa × TIPO CAMBIO "
            "COSTOS × HORAS COBRADAS, 2)."
        )
    elif costo_por_hora:
        texto_afac = _CPH_NOTA_AFAC
    elif general:
        texto_afac = (
            "La tarifa AFAC es POR AVIÓN: en el Balance general la columna "
            "PERMISO AFAC (PROVISIÓN) va como valor; la fórmula con la tarifa "
            "de cada avión está en su libro individual (Reportes › Balance "
            "por avión)."
        )
    else:
        texto_afac = (
            "El avión no tiene configurada la «Aportación AFAC USD/hr» en su "
            "ficha: la columna PERMISO AFAC (PROVISIÓN) va vacía (ver la hoja "
            "'pendientes de captura')."
        )
    nota_afac = Comment(texto_afac, "VuelaTour")
    nota_afac.width = 320
    nota_afac.height = 110 if costo_por_hora else 80
    ca.comment = nota_afac
    cf = _num(ws, 1, 4, FACTOR_IVA_COSTOS, "0.00", bold=True)
    cf.fill = PatternFill("solid", fgColor=FILL_COSTOS)
    cf.border = _border
    nota_iva = Comment(
        _CPH_NOTA_FACTOR_IVA
        if costo_por_hora
        else "IVA de los costos (16 %): COSTO TOTAL USD S/IVA = COSTO TOTAL USD ÷ "
        "este factor y COSTO X HORA USD S/IVA = COSTO X HORA USD ÷ este "
        "factor.",
        "VuelaTour",
    )
    nota_iva.width = 300
    nota_iva.height = 100 if costo_por_hora else 70
    cf.comment = nota_iva
    return hay_afac


def _hoja_maestra(
    ws: Worksheet, req: BalanceAvionRequest, *, general: bool = False, costo_por_hora: bool = False
) -> _Maestra:
    # `general` solo ajusta las notas al pie (29-ago): en el Balance general
    # las hojas 'combustible'/'Gastos Indirectos'/'permisos' ya no existen —
    # viven en el libro individual de cada avión. Desde el 5-oct-2026 también
    # decide qué fórmulas caben (la tarifa AFAC y el TC promedio de cada libro
    # no están en el consolidado). Devuelve dónde quedó la hoja para que las
    # demás la citen en sus fórmulas.
    # `costo_por_hora` (6-oct-2026, API 0.0.64; solo con `general`): la
    # variante «Balance general» — columnas `_COLS_GENERAL` (COSTO TOTAL y el
    # bloque COSTO POR HORA en lugar de operaciones/piloto/otros/AFAC e
    # indicadores). Sin ella, el Balance mensual: las columnas de siempre.
    # Desde el 7-oct-2026 (API 0.0.66) las dos variantes del general van SIN
    # la columna COMISIONES (`_COLS_FLOTA` / `_COLS_GENERAL`) y la hoja se
    # pinta antes de comisiones (`_antes_de_comisiones`); el libro INDIVIDUAL
    # conserva `_COLS` tal cual.
    if not general:
        disp = _DISP_MENSUAL
    elif costo_por_hora:
        disp = _DISP_GENERAL
    else:
        disp = _DISP_FLOTA
    if disp.comision_col is None:
        req = _antes_de_comisiones(req)
    cols = disp.cols
    ws.title = sheet_title(f"reporte horas {req.matricula}")
    n = len(cols)
    maestra = _Maestra(
        titulo=ws.title,
        fila_ini=3,
        fila_fin=2 + len(req.vuelos),
        fila_tot=3 + len(req.vuelos),
        general=general,
        disp=disp,
    )

    # Encabezado compacto de 2 filas: grupo (merged) / columna.
    col = 1
    while col <= n:
        grupo = cols[col - 1][0]
        fin = col
        while fin < n and cols[fin][0] == grupo:
            fin += 1
        if grupo:
            c = ws.cell(row=1, column=col, value=disp.titulos_grupo.get(grupo, grupo))
            c.font = Font(bold=True, color="FFFFFF", size=10)
            c.fill = PatternFill("solid", fgColor=NAVY)
            c.alignment = Alignment(horizontal="center", vertical="center")
            if fin > col:
                ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=fin)
        col = fin + 1
    for i, (grupo, header, _attr, _fmt) in enumerate(cols, start=1):
        c = ws.cell(row=2, column=i, value=header)
        c.font = Font(bold=True, color=NAVY, size=8)
        c.fill = PatternFill("solid", fgColor=disp.fills.get(grupo, LIGHT))
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = _border
    ws.row_dimensions[2].height = 38
    # COMISIONES (6-oct-2026, API 0.0.65; antes COMISIÓN VENDEDOR, vacía
    # desde el 28-ago-2026): la nota del encabezado explica la regla de
    # septiembre 2026 en el libro INDIVIDUAL, el único que la lleva desde el
    # 7-oct-2026. En el Balance general, GANANCIA MXN (mensual) y REMANENTE
    # VENTA MENOS COMPRA (general) dicen que van antes de comisiones y dónde
    # quedaron.
    if disp.comision_col is not None:
        ws.cell(row=2, column=disp.comision_col).comment = _comentario_cobro(
            _con_regla(_NOTA_ENCABEZADO_COMISIONES, req.regla_comisiones)
        )
    else:
        for i, c in enumerate(cols, start=1):
            nota_encabezado = _NOTAS_ENCABEZADO_SIN_COMISIONES.get(c[2] or "")
            if nota_encabezado:
                ws.cell(row=2, column=i).comment = _comentario_cobro(
                    _con_regla(nota_encabezado, req.regla_comisiones)
                )
    hay_afac = _constantes_maestra(ws, req, general, costo_por_hora=disp.costo_por_hora)
    # Fila sin TC de costos: GANANCIA USD divide entre el TC promedio CRUDO
    # del libro (z ?? tcPromedio del API).
    tc_prom_filas = maestra.tc_promedio_crudo(local=True)

    # Datos: 1 fila por vuelo. `calculadas` = columnas cuya fórmula de siempre
    # cita columnas que la variante no tiene (van como valor del API).
    calculadas: set[int] = set()
    row = 3
    for v in req.vuelos:
        cobros = _cobros_a_4(v.cobros)
        # «Cómo se cobró» (6-oct-2026): nota de cada COBRO n MXN y resumen en
        # STATUS. Sin método/cuenta/registro no hay nota.
        notas_cobro = _notas_parcialidades(v.cobros)
        nota_status = _nota_resumen_cobros(v.cobros)
        # Regla del cliente: lo cobrado nunca puede ser menor a lo volado.
        # Resalta HORAS COBRADAS en ámbar cuando el taco registró más horas
        # (solo señal visual; el pendiente del API explica el caso).
        horas_menores = (
            v.horas_cobradas is not None
            and v.tiempo_vuelo is not None
            and v.horas_cobradas > 0
            and v.tiempo_vuelo - v.horas_cobradas > 0.01
        )
        # Balance GENERAL: SOLO la celda de la CLAVE se tiñe con el color del
        # avión (así lo maneja el equipo en su libro — el resto de la fila
        # conserva los colores por bloque); en el individual viene vacío.
        avion_fill = _hex(v.avion_color)
        # Vuelo EXTERNO (operador ajeno, 28-ago): un vuelo más de la flota;
        # la clave va en GRIS e itálica para distinguirlo de las matrículas.
        if v.es_externo and not avion_fill:
            avion_fill = "E5E7EB"
        for i, (grupo, _header, attr, fmt) in enumerate(cols, start=1):
            fill = disp.fills.get(grupo)
            if i == 1 and avion_fill:
                fill = avion_fill
            if attr is not None:
                val = _valor_columna(attr, v)
                # Dato que la celda muestra: en las columnas del bloque COSTO
                # POR HORA que REPITEN otra de la fila (`cph_*`, variante
                # general) es el de su columna de origen, y hereda sus señales
                # (azul del T.C. oficial, ámbar del salto de taco). En el
                # juego de siempre es la propia llave.
                origen = _ORIGEN_GENERAL.get(attr, attr)
                if fmt is None:
                    cell = ws.cell(row=row, column=i, value=_fecha(val) if attr == "fecha" else val)
                    if attr == "factura_vuelatour":
                        # El folio lo teclea la oficina: un «=A-12» sería
                        # FÓRMULA para openpyxl (guarda de la caja chica,
                        # 24-sep-2026) — se queda como TEXTO.
                        _texto_literal(cell)
                    if attr == "estado" and val == "CANCELADO":
                        cell.font = Font(color=RED, size=9)
                    elif attr == "estado":
                        cell.font = Font(color=MUTED, size=9)
                    elif attr == "clave" and v.es_externo:
                        cell.font = Font(italic=True, color="374151")
                    # STATUS no cambia de texto: la nota junta TODOS los
                    # cobros (uno por renglón) para verlos de un vistazo.
                    if attr == "status_cobro" and nota_status:
                        cell.comment = _comentario_cobro(nota_status)
                else:
                    try:
                        formula = _formula_fila(
                            attr,
                            v,
                            row,
                            hay_afac=hay_afac,
                            tc_prom_filas=tc_prom_filas,
                            disp=disp,
                        )
                    except _ColumnaAusente:
                        # Variante general: COSTO TOTAL = operación + piloto +
                        # otros + AFAC, columnas que ya no están ⇒ VALOR del
                        # API, «calculado por el sistema» (nota en su
                        # encabezado, abajo).
                        formula = None
                        if val is not None:
                            calculadas.add(i)
                    if attr == "cph_proveedor_mxn" and val is not None:
                        # TOTAL PARA PROVEEDOR (PESOS) es el COSTO TOTAL del
                        # vuelo y, desde el 8-oct-2026, su único lugar en el
                        # general: la nota «calculado por el sistema» va en
                        # su encabezado (antes en COSTO TOTAL MXN).
                        calculadas.add(i)
                    cell = _formula(ws, row, i, formula, val, fmt)
                if attr == "horas_cobradas" and horas_menores:
                    fill = AMBER
                # Vuelo MULTI-AVIÓN (regla B, 28-ago-2026): la VENTA de la
                # fila es la parte proporcional de esta matrícula — la nota
                # de la celda dice cuánto y con qué base se repartió.
                if attr == "total_usd" and v.multi_avion:
                    nota_part = Comment(_nota_participacion(v), "VuelaTour")
                    nota_part.width = 320
                    nota_part.height = 110
                    cell.comment = nota_part
                # TC no capturado en la cotización → se usó el oficial: la
                # celda del TC y las que derivan de él (MXN) se marcan
                # (también TOTAL COBRADO S/IVA (PESOS), que repite TOTAL
                # S/IVA MXN en la variante general).
                if v.tc_venta_oficial and origen in (
                    "tc_venta", "total_mxn", "iva_mxn", "subtotal_mxn",
                ):
                    fill = TC_OFICIAL_FILL
                    if attr == "tc_venta":
                        nota_tc = Comment(_nota_tc_oficial(v), "VuelaTour")
                        nota_tc.width = 320
                        nota_tc.height = 80
                        cell.comment = nota_tc
                # Salto INTERNO entre tramos del MISMO vuelo: infla las horas
                # sin romper la cadena entre vuelos — se pinta en la celda de
                # horas voladas con el tramo culpable en la nota. En la
                # variante general también TIEMPO CALZOS/HOBS (HR): es el
                # divisor de COSTO X HORA y unas horas infladas lo bajan.
                if origen == "tiempo_vuelo" and v.salto_taco_interno:
                    fill = AMBER
                    nota_int = Comment(
                        "Salto entre tramos del vuelo (las horas pueden "
                        "salir infladas): "
                        + (v.salto_taco_interno_detalle or "revisar tacos"),
                        "VuelaTour",
                    )
                    nota_int.width = 280
                    nota_int.height = 70
                    cell.comment = nota_int
                # Salto en la cadena de tacómetros y/u OBSERVACIONES del
                # equipo (Tacómetros en vivo): mismo amarillo que el panel;
                # la nota de la celda junta el salto y los comentarios.
                if attr == "taco_inicio" and (
                    v.salto_taco_inicio or v.taco_inicio_obs
                ):
                    fill = AMBER
                    lineas_ti: list[str] = []
                    if v.salto_taco_inicio and v.salto_taco_esperado is not None:
                        lineas_ti.append(
                            "Salto en la cadena de tacómetros: no empalma "
                            f"con la llegada anterior ({v.salto_taco_esperado:g})."
                        )
                    elif v.salto_taco_inicio:
                        lineas_ti.append("Salto en la cadena de tacómetros.")
                    lineas_ti.extend(v.taco_inicio_obs)
                    if lineas_ti:
                        nota_salto = Comment("\n".join(lineas_ti), "VuelaTour")
                        nota_salto.width = 320
                        nota_salto.height = min(220, 45 + 30 * len(lineas_ti))
                        cell.comment = nota_salto
                # Observaciones sobre la LLEGADA → celda TACO FINAL.
                if attr == "taco_fin" and v.taco_fin_obs:
                    fill = AMBER
                    nota_tf = Comment("\n".join(v.taco_fin_obs), "VuelaTour")
                    nota_tf.width = 320
                    nota_tf.height = min(220, 45 + 30 * len(v.taco_fin_obs))
                    cell.comment = nota_tf
                # Nota con el desglose del total de la celda (gastos que la
                # componen), visible al pasar el cursor en Excel.
                det_attr = disp.detalle_attr.get(attr)
                if det_attr:
                    lineas = getattr(v, det_attr, None) or []
                    if lineas:
                        nota = Comment("\n".join(lineas), "VuelaTour")
                        nota.width = 340
                        nota.height = min(260, 40 + 16 * len(lineas))
                        cell.comment = nota
                # Variante general: el desglose operación / piloto / otros /
                # AFAC (columnas del Balance mensual) va en UNA nota, la de
                # TOTAL PARA PROVEEDOR (PESOS).
                if attr == "cph_proveedor_mxn":
                    nota_costo = _nota_desglose_costo(v)
                    if nota_costo:
                        cell.comment = _comentario_desglose(nota_costo)
                # COMISIONES (regla de septiembre 2026): qué comisiones son,
                # una línea por concepto (las dos variantes).
                if attr == "comisiones_mxn":
                    nota_comisiones = _nota_comisiones(v)
                    if nota_comisiones:
                        cell.comment = _comentario_cobro(nota_comisiones)
            else:  # parcialidades de cobro (pares fecha/monto desde COBRO 1)
                idx = (i - disp.cobro1_col) // 2
                cobro = cobros[idx] if idx < len(cobros) else None
                if fmt is None:
                    cell = ws.cell(row=row, column=i, value=_fecha(cobro.fecha) if cobro else None)
                else:
                    # Lo que entró a la cuenta: el NETO con la regla de
                    # septiembre 2026 (sin él, el monto de siempre).
                    cell = _num(ws, row, i, _neto_cobro(cobro) if cobro else None, fmt)
                    nota_cobro = notas_cobro[idx] if idx < len(notas_cobro) else None
                    if nota_cobro:
                        cell.comment = _comentario_cobro(nota_cobro)
            cell.border = _border
            if fill:
                cell.fill = PatternFill("solid", fgColor=fill)
        row += 1

    # «Calculado por el sistema» (variante general): la nota va UNA vez, en
    # el encabezado de la columna.
    for i in sorted(calculadas):
        nota_sis = Comment(
            _NOTA_CALCULADO_POR_SISTEMA.get(
                cols[i - 1][2] or "", _NOTA_CALCULADO_POR_SISTEMA_GENERICA
            ),
            "VuelaTour",
        )
        nota_sis.width = 320
        nota_sis.height = 120
        ws.cell(row=2, column=i).comment = nota_sis

    # Fila TOTALES al final (sumas y promedios YA calculados por el API; desde
    # el 5-oct-2026 como fórmulas SUM/AVERAGE visibles con ese número). En la
    # variante general, las columnas del bloque COSTO POR HORA que el API no
    # totaliza llevan la Σ de sus filas o el promedio sin IVA
    # (`_CPH_TOTALES_PROPIOS`, revisión 6-oct-2026).
    t = req.totales
    ws.cell(row=row, column=1, value="TOTALES").font = Font(bold=True)
    for i, (_grupo, _header, attr, fmt) in enumerate(cols, start=1):
        cell = ws.cell(row=row, column=i)
        total_attr = disp.total_map.get(attr) if attr else None
        if total_attr is not None:
            cell = _formula(
                ws,
                row,
                i,
                _formula_total(attr, maestra),
                _valor_total(t, total_attr),
                fmt or MONEY,
                bold=True,
            )
        elif disp.costo_por_hora and attr in _CPH_TOTALES_PROPIOS:
            valor_tot = _valor_total_costo_por_hora(attr, req)
            try:
                formula_tot = _formula_total_costo_por_hora(attr, maestra)
                ausente = False
            except _ColumnaAusente:
                formula_tot, ausente = None, True
            cell = _formula(ws, row, i, formula_tot, valor_tot, fmt or MONEY, bold=True)
            if ausente and valor_tot is not None:
                cell.comment = _nota_celda(_NOTA_CALCULADO_POR_SISTEMA_CELDA)
        cell.fill = PatternFill("solid", fgColor=LIGHT)
        cell.border = _border
        if cell.value is not None or i == 1:
            cell.font = Font(bold=True)
    row += 1

    # Renglones informativos (regla 28-ago-2026): lo que NO está en las filas.
    # 1) TUAs/extras/pernocta/comisión del vendedor COTIZADOS (+ su IVA) =
    #    ingreso de VuelaTour,
    #    EXCLUIDOS de VENTA AVIÓN; el detalle vive en 'otros movimientos'
    #    del Balance general. Es lo cotizado de TODOS los estados (misma
    #    base que VENTA AVIÓN), no lo cobrado.
    otros = t.otros_ingresos_usd
    if otros is not None and otros != 0:
        lc = ws.cell(
            row=row,
            column=1,
            value="TUAs/extras/pernocta/comisión del vendedor COTIZADOS en el "
            "periodo (con su IVA; todos los estados salvo CANCELADO, cuya "
            "venta es solo lo cobrado y retenido) — EXCLUIDOS de las filas: "
            "ingreso de VuelaTour, detalle en 'otros movimientos' del "
            "Balance general:",
        )
        lc.font = Font(bold=True, size=9)
        lc.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
        ws.row_dimensions[row].height = 30
        _num(ws, row, 9, otros, MONEY, bold=True)
        ws.cell(row=row, column=10, value="USD").font = Font(bold=True, size=9)
        row += 1
    # 2) TUA pagado al aeropuerto: solo nota en OPERACIONES, no resta aquí.
    #    El API lo manda como suma → 0 = no hubo; se omite igual que None.
    tua = t.tua_pagado_mxn
    if tua is not None and tua != 0:
        ws.cell(
            row=row,
            column=1,
            value=_CPH_RENGLON_TUA
            if disp.costo_por_hora
            else "TUA pagado del periodo (solo nota en OPERACIONES, no resta en este libro):",
        ).font = Font(bold=True, size=9)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
        _num(ws, row, 9, tua, MONEY, bold=True)
        ws.cell(row=row, column=10, value="MXN").font = Font(bold=True, size=9)
        row += 1
    # 3) Extensión y/o antelación de horario pagada al aeropuerto (1-oct-2026,
    #    API 0.0.47): traslado al cliente como el TUA — solo nota en
    #    OPERACIONES, no resta aquí. El API manda la llave solo con suma ≠ 0;
    #    sin ella no se pinta nada (libro byte-idéntico).
    extension = t.extension_pagada_mxn
    if _monto_no_cero(extension):
        ws.cell(
            row=row,
            column=1,
            value=_CPH_RENGLON_EXTENSION
            if disp.costo_por_hora
            else "Extensión de horario pagada del periodo (solo nota en "
            "OPERACIONES, no resta en este libro):",
        ).font = Font(bold=True, size=9)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
        _num(ws, row, 9, extension, MONEY, bold=True)
        ws.cell(row=row, column=10, value="MXN").font = Font(bold=True, size=9)
        row += 1
    row += 1
    hay_extension = _hay_extension_pagada(req)

    # Notas al pie (las de la variante general: `_notas_pie_costo_por_hora`).
    # COMISIONES / GANANCIA y COBRADO REAL: las del libro INDIVIDUAL, que lleva
    # la columna, o las del Balance general (7-oct-2026), que ya no la lleva.
    if disp.comision_col is not None:
        nota_comisiones = (
            _con_regla(_NOTA_COMISIONES, req.regla_comisiones)
            + _MAESTRA_PAGO_VENDEDOR_PROVISION
            + _NOTA_GANANCIA
        )
        nota_cobrado_real = _con_regla(_NOTA_COBRADO_REAL, req.regla_comisiones)
    else:
        nota_comisiones = (
            _con_regla(_NOTA_COMISIONES_FLOTA, req.regla_comisiones) + _NOTA_GANANCIA_FLOTA
        )
        nota_cobrado_real = _con_regla(_NOTA_COBRADO_REAL_FLOTA, req.regla_comisiones)
    notas_pie = (
        _NOTA_VENTA_AVION,
        nota_comisiones,
        nota_cobrado_real,
        "TIPO CAMBIO COSTOS y COSTO X HORA USD de la fila TOTALES son PROMEDIOS "
        "(los demás son sumas).",
        _NOTA_ESTATUS_COBRO,
        _NOTA_FACTURA_VUELATOUR,
        _NOTA_COMBUSTIBLE_FILAS_SIN_GAS
        + (" del libro individual de cada avión" if general else "")
        + (
            _NOTA_COMBUSTIBLE_COLA_EXTENSION
            if hay_extension
            else _NOTA_COMBUSTIBLE_COLA_TUA
        ),
        _NOTA_TRASLADOS_EXTENSION if hay_extension else _NOTA_TRASLADOS_TUA,
        _NOTA_COMPARTIDO,
        _NOTA_TACOS_AMBAR,
    )
    if disp.costo_por_hora:
        notas_pie = _notas_pie_costo_por_hora(req, hay_extension)
    for nota in notas_pie:
        ws.cell(row=row, column=1, value=nota).font = Font(color=MUTED, size=9, italic=True)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=14)
        row += 1

    # Anchos + paneles congelados (bajo el encabezado, a la derecha de RUTA).
    anchos = {1: 16, 2: 12, 3: 34, 4: 12}
    for i, (grupo, _h, attr, _f) in enumerate(cols, start=1):
        if attr in ("cobrado_real_mxn", "cobrado_mxn"):
            anchos[i] = 17
        elif attr == "factura_vuelatour":
            anchos[i] = 16
        elif grupo == _GRUPO_COSTO_HORA:
            anchos[i] = _ANCHO_COSTO_HORA
    for i in range(1, n + 1):
        ws.column_dimensions[get_column_letter(i)].width = anchos.get(i, 13)
    # GANANCIA / REMANENTE negativos en rojo (regla de Excel, 8-oct-2026):
    # filas de vuelo + TOTALES de cada columna que esta variante tenga.
    _rojo_si_negativo(
        ws,
        *dict.fromkeys(
            f"{letra}{maestra.fila_ini}:{letra}{maestra.fila_tot}"
            for attr in _COLUMNAS_NEGATIVO_ROJO
            if (letra := disp.letra.get(attr)) is not None
        ),
    )
    ws.freeze_panes = "D3"
    return maestra


# Cómo se reproduce TOTAL USD de una hoja de gastos con fórmula (5-oct-2026):
# 'tc_libro' = TOTAL MXN ÷ el TC promedio CRUDO del libro (buildHoja del API
# en el libro individual); 'tc_celda' = TOTAL MXN ÷ la celda TC PROMEDIO (los
# gastos de EMPRESA del general, que el API convierte con el promedio ya
# redondeado); None = valor (hojas de FLOTA: Σ de los USD de cada avión).
_USD_TC_LIBRO = "tc_libro"
_USD_TC_CELDA = "tc_celda"


_CAMPO_PLANTILLA = re.compile(r"\{(\w+)\}")


def _formula_cita(
    ws: Worksheet,
    row: int,
    col: int,
    maestra: _Maestra | None,
    plantilla: str,
    valor: float | None,
    fmt: str = MONEY,
    **font,
):
    """Celda de otra hoja que cita la fila TOTALES de la hoja maestra: cada
    `{llave}` de `plantilla` es la celda TOTALES de esa columna («{tc_costos}»,
    «ROUND(E5/{tiempo_vuelo},2)»). Sin maestra, el valor del API (como
    siempre). Si la variante de la hoja no tiene alguna de esas columnas
    (6-oct-2026), el VALOR del API con nota «calculado por el sistema» —
    jamás una referencia a una columna que no existe."""
    if maestra is None:
        return _formula(ws, row, col, None, valor, fmt, **font)
    refs = {campo: maestra.total(campo) for campo in _CAMPO_PLANTILLA.findall(plantilla)}
    if any(ref is None for ref in refs.values()):
        cell = _formula(ws, row, col, None, valor, fmt, **font)
        if valor is not None:
            cell.comment = _nota_celda(_NOTA_CALCULADO_POR_SISTEMA_CELDA)
        return cell
    return _formula(ws, row, col, plantilla.format(**refs), valor, fmt, **font)


def _usd_de_mxn(modo: str | None, mxn: str, tc: str, maestra: _Maestra | None) -> str | None:
    """Fórmula de un TOTAL USD = ROUND(mxn ÷ TC promedio, 2) según `modo`."""
    if modo == _USD_TC_CELDA:
        return f"ROUND({mxn}/{tc},2)"
    if modo == _USD_TC_LIBRO and maestra is not None:
        crudo = maestra.tc_promedio_crudo()
        return f"ROUND({mxn}/{crudo},2)" if crudo else None
    return None


def _hoja_gastos(ws: Worksheet, titulo: str, hoja: BalanceAvionHojaGastos,
                 req: BalanceAvionRequest, nota: str | None = None, *,
                 titulo_exacto: bool = False,
                 resaltar_parciales: bool = False,
                 maestra: _Maestra | None = None,
                 usd: str | None = None) -> None:
    # `maestra` + `usd` (5-oct-2026): el resumen de arriba va con FÓRMULAS —
    # TOTAL MXN = Σ de la columna MONTO MXN, TC PROMEDIO y HRS VOLADAS citan
    # la fila TOTALES de la hoja maestra, TOTAL USD según `usd` (ver
    # _USD_TC_LIBRO) y USD X HR = TOTAL USD ÷ HRS VOLADAS.
    # `titulo_exacto`: pinta `titulo` tal cual, sin el sufijo "— MATRÍCULA"
    # (lo usa la hoja 'otros gastos' del GENERAL, cuyo título ya trae
    # "VuelaTour" y el sufijo "FLOTA" solo estorba).
    # `resaltar_parciales` (2-sep-2026, SOLO la hoja "Gastos Indirectos" del
    # libro INDIVIDUAL): las filas del reparto manual (DETALLE con "reparto
    # manual: $X de $Y") llevan un tinte suave en la celda DETALLE para
    # distinguirlas de los gastos capturados directo al avión. El GENERAL
    # no lo usa: ahí esa celda lleva el color del avión.
    # Ancho de la columna DETALLE: el texto ENVUELVE (wrap) y el alto de la
    # fila se estima con él — mismo patrón que DETALLE DE COBROS en la hoja
    # cobranza (el cliente ajustaba estas filas a mano).
    ancho_detalle = 70
    # Hoja "refacciones" del GENERAL (29-ago): columnas COSTO VUELATOUR /
    # VENTA AL AVIÓN / GANANCIA solo cuando el payload las trae.
    con_costo_venta = any(
        f.costo_mxn is not None or f.venta_mxn is not None for f in hoja.filas
    )
    n_cols = 9 if con_costo_venta else 6
    _title(
        ws,
        titulo if titulo_exacto else f"{titulo} — {req.matricula}".strip(" —"),
        1,
        n_cols,
    )
    periodo = f"Periodo: {req.periodo_desde or '—'} a {req.periodo_hasta or '—'}"
    ws.cell(row=2, column=1, value=periodo).font = Font(italic=True, size=10, color=MUTED)

    _header_row(ws, 4, ["TOTAL MXN", "TC PROMEDIO", "TOTAL USD", "HRS VOLADAS", "USD X HR"])
    ultima = 7 + len(hoja.filas)
    _formula(
        ws, 5, 1, f"ROUND(SUM(D8:D{ultima}),2)" if hoja.filas else None,
        hoja.total_mxn, MONEY, bold=True,
    )
    _formula_cita(ws, 5, 2, maestra, "{tc_costos}", req.totales.tc_promedio, TC)
    _formula(ws, 5, 3, _usd_de_mxn(usd, "A5", "B5", maestra), hoja.usd, MONEY, bold=True)
    _formula_cita(ws, 5, 4, maestra, "{tiempo_vuelo}", req.totales.tiempo_vuelo, HORAS)
    _formula(ws, 5, 5, "ROUND(C5/D5,2)", hoja.usd_hr, MONEY)
    for c in range(1, 6):
        ws.cell(row=5, column=c).border = _border

    headers = ["FECHA", "CATEGORÍA", "DETALLE", "MONTO MXN",
               "MONEDA ORIGINAL", "MONTO ORIGINAL"]
    if con_costo_venta:
        headers += ["COSTO VUELATOUR\nMXN", "VENTA AL AVIÓN\nMXN",
                    "GANANCIA\nMXN"]
    _header_row(ws, 7, headers)
    row = 8
    if hoja.filas:
        for f in hoja.filas:
            ws.cell(row=row, column=1, value=_fecha(f.fecha)).border = _border
            ws.cell(row=row, column=2, value=f.categoria).border = _border
            # DETALLE ÍNTEGRO: se pinta tal cual llega del API — jamás se
            # recorta ni se re-arma. En la hoja 'otros gastos' del GENERAL el
            # API puede mandar el vuelo dentro del detalle («… · vuelo #123»,
            # 11-sep-2026): esa referencia tiene que llegar completa a la
            # celda (el ancho/alto de abajo solo la ENVUELVEN).
            dc = ws.cell(row=row, column=3, value=f.detalle)
            dc.border = _border
            dc.alignment = Alignment(wrap_text=True, vertical="top")
            # Tinte suave del parcial (sin columnas nuevas: anchos y merges
            # siguen por índice); el color del avión, si viene, gana abajo.
            if resaltar_parciales and _es_parcial_reparto(f.detalle):
                dc.fill = PatternFill("solid", fgColor=LIGHT)
            _num(ws, row, 4, f.monto_mxn).border = _border
            # Moneda/monto original solo cuando el gasto NO se capturó en MXN.
            ws.cell(row=row, column=5, value=f.moneda_original).border = _border
            _num(ws, row, 6, f.monto_original).border = _border
            if con_costo_venta:
                _num(ws, row, 7, f.costo_mxn).border = _border
                _num(ws, row, 8, f.venta_mxn).border = _border
                # GANANCIA = venta − costo (solo para MOSTRAR; 0 mientras la
                # salida se cargue a costo). Falta un lado → celda vacía.
                ganancia = (
                    round(f.venta_mxn - f.costo_mxn, 2)
                    if f.venta_mxn is not None and f.costo_mxn is not None
                    else None
                )
                _formula(ws, row, 9, f"ROUND(H{row}-G{row},2)", ganancia).border = _border
            # Balance GENERAL: SOLO la celda del DETALLE (lleva la matrícula
            # al frente) se tiñe con el color del avión — como su libro.
            avion_fill = _hex(f.avion_color)
            if avion_fill:
                dc.fill = PatternFill("solid", fgColor=avion_fill)
            # Alto de fila = renglones REALES que envuelve el DETALLE (mismo
            # patrón que la hoja cobranza): sin él Excel recorta el texto.
            detalle_txt = f.detalle or ""
            lineas = sum(
                max(1, math.ceil(len(linea) / ancho_detalle))
                for linea in detalle_txt.split("\n")
            ) if detalle_txt else 1
            if lineas > 1:
                ws.row_dimensions[row].height = 14 * lineas + 4
            row += 1
    else:
        ws.cell(row=row, column=1, value="Sin gastos registrados en el periodo.").font = Font(
            color=MUTED, italic=True
        )
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n_cols)
        row += 1

    # GANANCIA MXN negativa en rojo (hoja 'refacciones'), 8-oct-2026.
    if con_costo_venta and hoja.filas:
        _rojo_si_negativo(ws, f"I8:I{ultima}")

    if nota:
        row += 1
        ws.cell(row=row, column=1, value=nota).font = Font(
            color=MUTED, size=9, italic=True
        )
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n_cols)

    anchos = [14, 16, ancho_detalle, 15, 16, 15]
    if con_costo_venta:
        anchos += [16, 16, 14]
    for i, w in enumerate(anchos, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A8"


def _fusionar_hojas_gastos(
    a: BalanceAvionHojaGastos, b: BalanceAvionHojaGastos
) -> BalanceAvionHojaGastos:
    """Hoja "Gastos Indirectos" del libro INDIVIDUAL (2-sep-2026): UNA sola
    lista con `gastos_indirectos` + `otros_gastos`, que el API sigue
    mandando aparte (el contrato no cambia; `otros_gastos` alimenta además
    la hoja 'repartidos a aviones' del general). Función PURA de
    presentación: filas de ambas ordenadas por fecha (sort estable — a
    igual fecha primero las de `a`, cada lista en su orden) y totales =
    suma None-tolerante de los totales que YA vienen calculados del API.
    Jamás se recalcula el USD desde el MXN."""
    return BalanceAvionHojaGastos(
        filas=sorted([*a.filas, *b.filas], key=lambda f: f.fecha or ""),
        total_mxn=_suma_none(a.total_mxn, b.total_mxn),
        usd=_suma_none(a.usd, b.usd),
        usd_hr=_suma_none(a.usd_hr, b.usd_hr),
    )


# Cobranza con la regla de septiembre 2026 (6-oct-2026, API 0.0.65): COBRADO
# REAL llega NETO del API; el detalle pinta el neto con el bruto al lado y
# la nota de abajo lo explica — SOLO si algún cobro trae `neto_mxn` (sin él,
# la hoja de siempre).
_NOTA_COBRANZA_NETO_INICIO = (
    "Desde septiembre 2026 COBRADO REAL va NETO: lo que entró a la cuenta ya "
    "sin la comisión que retuvo el banco (el detalle trae neto y bruto; la "
    "comisión sigue en COMISIÓN BANCO MXN). COBRADO sigue siendo la parte del "
    "avión de los cobros BRUTOS, antes de comisiones: la parte de la comisión "
    "que absorbe el avión resta en la columna COMISIONES "
)
_NOTA_COBRANZA_NETO = _NOTA_COBRANZA_NETO_INICIO + (
    "de la hoja de vuelos "
    "y la de VuelaTour (toda, en un vuelo EXTERNOS) queda en 'otros "
    "movimientos' del Balance general."
)
# En el Balance general (7-oct-2026) su hoja de vuelos ya no lleva COMISIONES.
_NOTA_COBRANZA_NETO_FLOTA = _NOTA_COBRANZA_NETO_INICIO + (
    "del libro individual de cada avión (en este libro, en la cascada de la "
    "hoja 'balance') y la de VuelaTour (toda, en un vuelo EXTERNOS) queda en "
    "'otros movimientos'."
)


def _hay_netos(req: BalanceAvionRequest) -> bool:
    """¿Algún cobro del libro trae `neto_mxn` (regla de septiembre 2026)?"""
    return any(c.neto_mxn is not None for v in req.vuelos for c in v.cobros)


def _monto_detalle_cobranza(c: BalanceAvionCobro) -> str | None:
    """Monto del cobro en el DETALLE de 'cobranza': «$12,522.20» de siempre o,
    con un neto distinto del bruto, «$19,380.00 neto (bruto $20,400.00)»."""
    if c.monto_mxn is None:
        return None
    if _bruto_neto(c) is None:
        return f"${_neto_cobro(c):,.2f}"
    return f"${c.neto_mxn:,.2f} neto (bruto ${c.monto_mxn:,.2f})"


def _hoja_cobranza(ws: Worksheet, req: BalanceAvionRequest, *, general: bool = False) -> None:
    """Hoja 'cobranza' (26-ago-2026 · regla 28-ago): el estatus de cobro POR
    VUELO al frente — cuánto se debía cobrar (VENTA DEL AVIÓN), cuánto YA se
    cobró (prorrateado al avión) y cuánto falta; más el TOTAL COTIZACIÓN
    (c/extras), los depósitos REALES, la comisión que retuvo el banco y la
    cuenta que recibió cada parcialidad. Los mismos números del bloque
    STATUS DE COBROS del final de la hoja maestra, en formato legible.
    `general` (7-oct-2026) solo cambia la nota de los netos: en el Balance
    general la columna COMISIONES ya no está en su hoja de vuelos."""
    n_cols = 13
    # Ancho (en caracteres) de la columna DETALLE DE COBROS: sirve para
    # estimar cuántos renglones envuelve cada parcialidad y no recortar.
    ancho_detalle = 46
    _title(ws, f"Cobranza — {req.matricula}".strip(" —"), 1, n_cols)
    periodo = f"Periodo: {req.periodo_desde or '—'} a {req.periodo_hasta or '—'}"
    ws.cell(row=2, column=1, value=periodo).font = Font(italic=True, size=10, color=MUTED)

    t = req.totales

    def _comision_banco(v: BalanceAvionVuelo) -> float | None:
        """Σ comisiones bancarias de las parcialidades (None si ninguna trae)."""
        montos = [c.comision_mxn for c in v.cobros if c.comision_mxn is not None]
        return round(sum(montos), 2) if montos else None

    def _suma(valores: list[float | None]) -> float | None:
        nums = [x for x in valores if x is not None]
        return round(sum(nums), 2) if nums else None

    # Totales de las columnas nuevas: los manda el API; si un API viejo no
    # los trae, se re-suman las filas SOLO para mostrar (nunca negocio).
    tot_cotizacion = t.total_cotizacion_mxn
    if tot_cotizacion is None:
        tot_cotizacion = _suma([v.total_cotizacion_mxn for v in req.vuelos])
    tot_cobrado_real = t.cobrado_real_mxn
    if tot_cobrado_real is None:
        tot_cobrado_real = _suma([v.cobrado_real_mxn for v in req.vuelos])
    tot_comision = t.comision_banco_mxn
    if tot_comision is None:
        tot_comision = _suma([_comision_banco(v) for v in req.vuelos])

    _header_row(ws, 4, ["TOTAL A COBRAR\nMXN (venta avión)", "COBRADO\nMXN",
                        "POR COBRAR\nMXN", "POR COBRAR\nUSD",
                        "TOTAL COTIZACIÓN\nMXN (c/extras)", "COBRADO REAL\nMXN",
                        "% COBRADO"])
    ws.row_dimensions[4].height = 30
    # Fórmulas (5-oct-2026): el resumen de arriba cita la fila TOTALES de
    # esta hoja (que a su vez es Σ de las columnas) y % COBRADO = COBRADO ÷
    # TOTAL A COBRAR.
    fila_tot = 8 + len(req.vuelos)
    _formula(ws, 5, 1, f"F{fila_tot}", t.total_mxn, MONEY, bold=True)
    _formula(ws, 5, 2, f"G{fila_tot}", t.cobrado_mxn, MONEY, bold=True)
    pc = _formula(ws, 5, 3, f"H{fila_tot}", t.por_cobrar_mxn, MONEY, bold=True)
    if (t.por_cobrar_mxn or 0) > 0.005:
        pc.font = Font(bold=True, color=RED)
    _formula(ws, 5, 4, f"I{fila_tot}", t.por_cobrar_usd, MONEY)
    _formula(ws, 5, 5, f"J{fila_tot}", tot_cotizacion, MONEY, bold=True)
    _formula(ws, 5, 6, f"K{fila_tot}", tot_cobrado_real, MONEY, bold=True)
    # % COBRADO = COBRADO ÷ TOTAL A COBRAR: la MISMA base que POR COBRAR
    # (ambos al TC de venta). Nunca cobrado real ÷ total cotización: son
    # pesos a TCs distintos (el depósito trae su propio TC) y contradecían
    # el POR COBRAR de la misma fila (0 por cobrar y 97% cobrado).
    pct = (
        (t.cobrado_mxn or 0) / t.total_mxn
        if t.total_mxn and t.total_mxn > 0
        else None
    )
    _formula(ws, 5, 7, "B5/A5", pct, "0.0%")
    for c in range(1, 8):
        ws.cell(row=5, column=c).border = _border

    headers = ["CLAVE", "FECHA", "RUTA", "ESTADO", "STATUS\nCOBRO",
               "TOTAL A COBRAR\nMXN (venta avión)", "COBRADO\nMXN",
               "POR COBRAR\nMXN", "POR COBRAR\nUSD",
               "TOTAL COTIZACIÓN\nMXN (c/extras)", "COBRADO REAL\nMXN",
               "COMISIÓN\nBANCO MXN",
               "DETALLE DE COBROS\n(fecha · monto · método · cuenta · comisión)"]
    _header_row(ws, 7, headers)
    ws.row_dimensions[7].height = 42
    row = 8
    for v in req.vuelos:
        cc = ws.cell(row=row, column=1, value=v.clave)
        avion_fill = _hex(v.avion_color)
        if avion_fill:
            cc.fill = PatternFill("solid", fgColor=avion_fill)
        ws.cell(row=row, column=2, value=_fecha(v.fecha))
        ws.cell(row=row, column=3, value=v.ruta)
        ec = ws.cell(row=row, column=4, value=v.estado)
        ec.font = Font(color=RED if v.estado == "CANCELADO" else MUTED, size=9)
        st = ws.cell(row=row, column=5, value=v.status_cobro or "—")
        if v.status_cobro == "Cobrado":
            st.fill = PatternFill("solid", fgColor="D1FAE5")
            st.font = Font(color="065F46", bold=True, size=9)
        elif v.status_cobro == "Parcial":
            st.fill = PatternFill("solid", fgColor=AMBER)
            st.font = Font(bold=True, size=9)
        elif v.status_cobro == "Pendiente":
            st.fill = PatternFill("solid", fgColor="FECACA")
            st.font = Font(color="991B1B", bold=True, size=9)
        _num(ws, row, 6, v.total_mxn)
        _num(ws, row, 7, v.cobrado_mxn)
        # POR COBRAR = round2(venta avión − cobrado), igual que el API.
        pcc = _formula(ws, row, 8, f"ROUND(F{row}-G{row},2)", v.por_cobrar_mxn)
        if (v.por_cobrar_mxn or 0) > 0.005:
            pcc.font = Font(color=RED)
        _num(ws, row, 9, v.por_cobrar_usd)
        cot = _num(ws, row, 10, v.total_cotizacion_mxn)
        if v.estado == "CANCELADO":
            # Solo referencia de lo cotizado: no se cobró ni se cobrará y NO
            # suma en el total del periodo (el API la excluye).
            cot.font = Font(color=MUTED, italic=True)
        _num(ws, row, 11, v.cobrado_real_mxn)
        _num(ws, row, 12, _comision_banco(v))
        # Detalle: depósitos REALES (fecha · monto · método · cuenta ·
        # comisión) — las partes que no vienen se omiten. Con la regla de
        # septiembre 2026 el monto es el NETO que entró, con el bruto al lado.
        detalle = "\n".join(
            " · ".join(
                x
                for x in (
                    _fecha(c.fecha),
                    _monto_detalle_cobranza(c),
                    c.metodo or None,
                    c.cuenta or None,
                    (
                        f"comisión ${c.comision_mxn:,.2f}"
                        if c.comision_mxn is not None
                        else None
                    ),
                )
                if x
            )
            for c in v.cobros
        )
        dc = ws.cell(row=row, column=n_cols, value=detalle or None)
        dc.alignment = Alignment(wrap_text=True, vertical="top")
        # Alto de fila = renglones REALES del detalle: cada parcialidad
        # trae cuenta + comisión (~80 caracteres) y envuelve en 2 líneas
        # dentro de la columna; con la altura por cobro Excel recortaba.
        lineas = sum(
            max(1, math.ceil(len(linea) / ancho_detalle))
            for linea in detalle.split("\n")
        ) if detalle else 1
        if lineas > 1:
            ws.row_dimensions[row].height = 14 * lineas + 4
        for c in range(1, n_cols + 1):
            ws.cell(row=row, column=c).border = _border
        row += 1

    ws.cell(row=row, column=1, value="TOTALES").font = Font(bold=True)
    # Σ de cada columna (5-oct-2026); TOTAL COTIZACIÓN deja fuera a los
    # CANCELADOS (su cotización es solo referencia: el API no la suma).
    ult = row - 1

    def _suma_col(letra: str) -> str | None:
        return f"ROUND(SUM({letra}8:{letra}{ult}),2)" if req.vuelos else None

    _formula(ws, row, 6, _suma_col("F"), t.total_mxn, MONEY, bold=True)
    _formula(ws, row, 7, _suma_col("G"), t.cobrado_mxn, MONEY, bold=True)
    _formula(ws, row, 8, _suma_col("H"), t.por_cobrar_mxn, MONEY, bold=True)
    _formula(ws, row, 9, _suma_col("I"), t.por_cobrar_usd, MONEY, bold=True)
    _formula(
        ws, row, 10,
        f'ROUND(SUMIF(D8:D{ult},"<>CANCELADO",J8:J{ult}),2)' if req.vuelos else None,
        tot_cotizacion, MONEY, bold=True,
    )
    _formula(ws, row, 11, _suma_col("K"), tot_cobrado_real, MONEY, bold=True)
    _formula(ws, row, 12, _suma_col("L"), tot_comision, MONEY, bold=True)
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.border = _border
        cell.fill = PatternFill("solid", fgColor=LIGHT)
    row += 2

    for nota in (
        "TOTAL A COBRAR = venta del avión (tiempo + ajuste + IVA "
        "proporcional; en vuelos multi-avión, la parte de esta matrícula). "
        "COBRADO = parcialidades reales × (venta avión ÷ total cotización): "
        "la parte de los cobros que corresponde a TUAs/extras/pernocta/"
        "comisión del vendedor es de VuelaTour. El detalle trae los "
        "depósitos reales y la comisión que retuvo el banco.",
        "TOTAL COTIZACIÓN = lo cobrado al cliente COMPLETO (con TUAs/extras/"
        "pernocta/comisión del vendedor y su IVA); en un vuelo CANCELADO la "
        "celda (gris) es solo "
        "referencia de lo cotizado y NO suma en el total. COBRADO REAL = "
        "depósitos tal cual entraron. "
        "% COBRADO = COBRADO ÷ TOTAL A COBRAR (misma base que POR COBRAR). "
        "COBRADO y POR COBRAR van al TC de venta: un depósito en pesos con "
        "su propio TC se convierte a USD con ese TC y se re-expresa al TC de "
        "venta, por lo que COBRADO puede diferir de COBRADO REAL sin que "
        "falte dinero.",
        "Filas COMPARTIDO (vuelo multi-avión) traen SU parte proporcional "
        "de la venta, cobros y por cobrar (partes iguales por tramo vendido; "
        "ferries/tramos operativos no reparten — regla 28-ago-2026); los "
        "clientes INTERNOS van sin venta a "
        "propósito (el interno no cobra). CANCELADO: su venta es lo "
        "realmente cobrado y retenido (cargo por cancelación / anticipo no "
        "reembolsado; también repartido si fue multi-avión) — sin "
        "pendiente; sus gastos cuentan igual.",
        *(
            (
                _con_regla(
                    _NOTA_COBRANZA_NETO_FLOTA if general else _NOTA_COBRANZA_NETO,
                    req.regla_comisiones,
                ),
            )
            if _hay_netos(req)
            else ()
        ),
    ):
        ws.cell(row=row, column=1, value=nota).font = Font(
            color=MUTED, size=9, italic=True
        )
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n_cols)
        row += 1

    for i, w in enumerate(
        [22, 11, 26, 12, 11, 16, 15, 15, 13, 16, 15, 14, ancho_detalle], start=1
    ):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A8"


def _hoja_combustible(ws: Worksheet, hoja: BalanceAvionHojaCombustible,
                      req: BalanceAvionRequest, *,
                      titulo: str | None = None,
                      maestra: _Maestra | None = None,
                      usd: str | None = None) -> None:
    """Pestaña 'combustible' (26-ago-2026): el gas del avión POR MES —
    ya no va por vuelo. Ledger con litros + resumen con $/L promedio.
    En el balance GENERAL (10-sep-2026) `req` es el consolidado de flota y
    las filas traen matrícula: se pintan por secciones con subtotal.
    Columna PAGO al final (11-sep-2026, pedido del cliente: con qué se pagó
    cada carga) — 7 columnas en encabezados, subtotales, merges y anchos; el
    API viejo (sin `pago`) deja la celda vacía.
    Fórmulas (5-oct-2026): subtotales y TOTAL DEL PERIODO = Σ de las cargas;
    arriba, TOTAL MXN cita el TOTAL DEL PERIODO, LITROS = Σ litros, $ X LITRO
    = TOTAL MXN ÷ LITROS, TC PROMEDIO cita la hoja maestra, TOTAL USD según
    `usd` (ver _USD_TC_LIBRO) y USD X HR = TOTAL USD ÷ horas voladas."""
    _title(ws, titulo or f"Combustible — {req.matricula}".strip(" —"), 1, 7)
    periodo = f"Periodo: {req.periodo_desde or '—'} a {req.periodo_hasta or '—'}"
    ws.cell(row=2, column=1, value=periodo).font = Font(italic=True, size=10, color=MUTED)

    _header_row(ws, 4, ["TOTAL MXN", "LITROS", "$ X LITRO PROM", "TC PROMEDIO",
                        "TOTAL USD", "USD X HR"])

    _header_row(ws, 7, ["FECHA", "DETALLE", "LITROS", "MONTO MXN",
                        "MONEDA ORIGINAL", "MONTO ORIGINAL", "PAGO"])
    row = 8
    # Filas de cargas (sin títulos de sección ni subtotales) y la fila del
    # TOTAL DEL PERIODO: el resumen de arriba las cita en sus fórmulas.
    rangos_cargas: list[tuple[int, int]] = []
    fila_total: int | None = None
    if hoja.filas:
        # SECCIONES por matrícula (pedido del cliente 28-ago, igual que la
        # pantalla Combustibles del panel): todas las cargas de un avión
        # juntas y, al terminar, su SUBTOTAL (cargas, litros, MXN, $/L); al
        # final el TOTAL del periodo. En el libro individual hay una sola
        # matrícula: solo se pinta el total.
        grupos: dict[str, list] = {}
        for f in hoja.filas:
            grupos.setdefault(f.matricula or req.matricula or "", []).append(f)
        varias = len(grupos) > 1

        def _subtotal(filas: list, etiqueta: str, fill: str,
                      rangos: list[tuple[int, int]]) -> None:
            nonlocal row
            litros = sum(f.litros or 0 for f in filas)
            mxn = sum(f.monto_mxn or 0 for f in filas)
            con_l = [f for f in filas if (f.litros or 0) > 0 and f.monto_mxn is not None]
            ppl = (
                sum(f.monto_mxn for f in con_l) / sum(f.litros for f in con_l)
                if con_l else None
            )
            c = ws.cell(
                row=row, column=2,
                value=f"{etiqueta} · {len(filas)} carga(s)"
                + (f" · $/L prom {ppl:,.2f}" if ppl is not None else ""),
            )
            c.font = Font(bold=True, size=9)
            # Σ de las cargas de `rangos` (fórmula visible, 5-oct-2026); los
            # litros se muestran a 1 decimal, como siempre.
            celdas = ",".join(f"{{L}}{a}:{{L}}{b}" for a, b in rangos)
            _formula(ws, row, 3, f"ROUND(SUM({celdas.format(L='C')}),1)",
                     round(litros, 1), HORAS, bold=True)
            _formula(ws, row, 4, f"ROUND(SUM({celdas.format(L='D')}),2)",
                     round(mxn, 2), MONEY, bold=True)
            for col in range(1, 8):
                cell = ws.cell(row=row, column=col)
                cell.border = _border
                cell.fill = PatternFill("solid", fgColor=fill)
            row += 1

        for mat, filas in sorted(grupos.items(), key=lambda kv: kv[0]):
            primera = row + (1 if varias else 0)
            if varias:
                h = ws.cell(row=row, column=1, value=mat)
                h.font = Font(bold=True, color=NAVY)
                sw = _hex(filas[0].avion_color)
                if sw:
                    h.fill = PatternFill("solid", fgColor=sw)
                ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
                row += 1
            for f in filas:
                ws.cell(row=row, column=1, value=_fecha(f.fecha)).border = _border
                dc = ws.cell(row=row, column=2, value=f.detalle)
                dc.border = _border
                # Mismo WRAP del DETALLE que las hojas de gastos (29-ago):
                # el texto envuelve y el alto de fila se estima con el ancho
                # de la columna (patrón de la hoja cobranza).
                dc.alignment = Alignment(wrap_text=True, vertical="top")
                _num(ws, row, 3, f.litros, HORAS).border = _border
                _num(ws, row, 4, f.monto_mxn).border = _border
                ws.cell(row=row, column=5, value=f.moneda_original).border = _border
                _num(ws, row, 6, f.monto_original).border = _border
                # PAGO (medio de pago ya legible del API): última columna,
                # con borde y wrap como el DETALLE. Sin dato = celda vacía.
                pc = ws.cell(row=row, column=7, value=f.pago)
                pc.border = _border
                pc.alignment = Alignment(wrap_text=True, vertical="top")
                # Balance GENERAL: el DETALLE (lleva la matrícula al frente)
                # se tiñe con el color del avión — como su libro.
                avion_fill = _hex(f.avion_color)
                if avion_fill:
                    dc.fill = PatternFill("solid", fgColor=avion_fill)
                detalle_txt = f.detalle or ""
                lineas = sum(
                    max(1, math.ceil(len(linea) / 52))
                    for linea in detalle_txt.split("\n")
                ) if detalle_txt else 1
                # El PAGO también envuelve (columna angosta): el alto manda
                # el texto más alto de la fila.
                if f.pago:
                    lineas = max(lineas, math.ceil(len(f.pago) / 18))
                if lineas > 1:
                    ws.row_dimensions[row].height = 14 * lineas + 4
                row += 1
            rangos_cargas.append((primera, row - 1))
            if varias:
                _subtotal(filas, f"Subtotal {mat}", "F3F4F6", [rangos_cargas[-1]])
        fila_total = row
        _subtotal(list(hoja.filas), "TOTAL DEL PERIODO (con IVA)", LIGHT, rangos_cargas)
    else:
        ws.cell(
            row=row, column=1,
            value="Sin cargas de combustible en el periodo.",
        ).font = Font(color=MUTED, italic=True)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
        row += 1

    # Resumen de arriba (fila 5), con fórmulas que citan las cargas.
    litros_celdas = ",".join(f"C{a}:C{b}" for a, b in rangos_cargas)
    _formula(ws, 5, 1, f"D{fila_total}" if fila_total else None,
             hoja.total_mxn, MONEY, bold=True)
    _formula(ws, 5, 2, f"ROUND(SUM({litros_celdas}),2)" if rangos_cargas else None,
             hoja.litros_total, HORAS)
    _formula(ws, 5, 3, "ROUND(A5/B5,2)", hoja.precio_litro_prom, MONEY)
    _formula_cita(ws, 5, 4, maestra, "{tc_costos}", req.totales.tc_promedio, TC)
    _formula(ws, 5, 5, _usd_de_mxn(usd, "A5", "D5", maestra), hoja.usd, MONEY, bold=True)
    _formula_cita(ws, 5, 6, maestra, "ROUND(E5/{tiempo_vuelo},2)", hoja.usd_hr, MONEY)
    for c in range(1, 7):
        ws.cell(row=5, column=c).border = _border

    row += 1
    ws.cell(
        row=row, column=1,
        value="El combustible se controla POR AVIÓN y POR MES (fecha del "
        "gasto, con o sin vuelo) — mismo criterio que el reparto a socios. "
        "Resta una sola vez en la hoja 'balance'.",
    ).font = Font(color=MUTED, size=9, italic=True)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)

    for i, w in enumerate([14, 52, 10, 15, 16, 15, 18], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A8"


def _hoja_balance(ws: Worksheet, req: BalanceAvionRequest,
                  refs: dict[str, str] | None = None) -> None:
    _title(ws, f"Balance — {req.matricula}".strip(" —"), 1, 3)
    periodo = f"Periodo: {req.periodo_desde or '—'} a {req.periodo_hasta or '—'} (todo en USD)"
    ws.cell(row=2, column=1, value=periodo).font = Font(italic=True, size=10, color=MUTED)
    _bloque_balance(ws, req, 4, refs=refs)
    for i, w in enumerate([46, 14, 16], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


# Etiqueta de la fila ÚNICA de indirectos en la cascada (2-sep-2026).
_FILA_INDIRECTOS = "(−) GASTOS INDIRECTOS USD"
# COMISIONES visibles en la cascada del Balance GENERAL (7-oct-2026, API
# 0.0.66): la hoja de vuelos de ese libro va antes de comisiones, así que el
# bloque de cada avión arranca antes de ellas y las resta en una línea; de
# UTILIDAD ANTES DE GASTOS para abajo, los números de siempre del API (los
# mismos de su libro individual).
# Las dos notas NO prometen cuadres que no siempre se dan (revisión 7-oct-2026):
# - «(−) COMISIONES»: por vuelo es la cifra en USD del reparto a socios
#   (`comisionesDelVuelo`), pero el TOTAL puede no coincidir con el del
#   reparto: el anticipo con comisión bancaria de un vuelo que aún no se
#   completa cuenta en el balance y no en el reparto, que solo lee vuelos
#   COMPLETADOS o CANCELADOS (caso real: el #314 de N4142R, CONFIRMADO del
#   15-sep-2026, cobro de $100,000 con $5,000 de comisión bancaria);
# - «antes de comisiones»: puede no cuadrar al centavo con la Σ de GANANCIA
#   USD de los vuelos del avión en la hoja de vuelos, que convierte cada
#   remanente con su T.C. de costos (y queda vacía sin él), mientras que la
#   comisión va con la cifra USD del reparto (la bancaria al T.C. de su
#   cobro; la provisión del vendedor, en los USD cotizados).
_FILA_ANTES_COMISIONES = "UTILIDAD ANTES DE GASTOS USD (antes de comisiones)"
_FILA_COMISIONES = "(−) COMISIONES (banco + vendedor) USD"
_NOTA_CASCADA_ANTES_COMISIONES = (
    "Utilidad de los vuelos del avión ANTES de comisiones = UTILIDAD ANTES DE "
    "GASTOS + COMISIONES (la hoja de vuelos de este libro va antes de "
    "comisiones). Puede no cuadrar al centavo con la suma de GANANCIA USD de "
    "sus vuelos en esa hoja: ahí cada vuelo convierte su remanente con su T.C. "
    "de costos (sin él, la celda va vacía) y aquí las comisiones van con su "
    "cifra en USD por vuelo, la del reparto a socios."
)
_NOTA_CASCADA_COMISIONES = (
    "Comisiones que absorbe el avión (regla de septiembre 2026, por fecha del "
    "vuelo): su parte de la comisión bancaria de sus cobros y la provisión de "
    "la comisión del vendedor, en USD — las de la columna COMISIONES de su "
    "libro individual. Por vuelo es la misma cifra que descuenta el reparto a "
    "socios, pero el total puede no coincidir con el del reparto: por ejemplo, "
    "si un vuelo que aún no se completa ya tiene un anticipo con comisión "
    "bancaria, este balance cuenta ese cobro y su comisión, y el reparto no "
    "(solo lee vuelos completados o cancelados)."
)


def _cascada_con_comisiones(b: BalanceAvionBalanceBloque) -> bool:
    """¿La cascada del avión en el Balance GENERAL lleva las dos líneas de
    COMISIONES? Solo si el API las manda (0.0.66: `comisiones_usd` y
    `utilidad_antes_comisiones_usd`) y son ≠ 0: sin comisiones no hay nada
    que restar y la cascada es la de siempre (un API previo, agosto)."""
    return b.utilidad_antes_comisiones_usd is not None and _monto_no_cero(b.comisiones_usd)


def _nota_otros_en_indirectos(otros_usd: float, general: bool) -> Comment:
    """Nota de la celda "(−) GASTOS INDIRECTOS USD" cuando parte del monto
    es reparto manual (otros_usd ≠ 0): dice cuánto y dónde está el detalle."""
    # Signo delante del "$" (un reparto negativo —reembolso/ajuste— no
    # debe salir como "$-X"): $1,234.56 / -$1,234.56.
    monto = f"-${abs(otros_usd):,.2f}" if otros_usd < 0 else f"${otros_usd:,.2f}"
    texto = (
        f"Incluye {monto} USD repartidos a mano a este avión (detalle en la "
        "hoja repartidos a aviones)."
        if general
        else f"Incluye {monto} USD de gastos administrativos repartidos a "
        "mano (filas reparto manual de la hoja Gastos Indirectos)."
    )
    com = Comment(texto, "VuelaTour")
    com.width = 320
    com.height = 80
    return com


# Filas de la cascada que, en el libro INDIVIDUAL, citan otra hoja (5-oct-2026):
# clave de `refs` → etiqueta de la fila.
_CASCADA_REFS = {
    "UTILIDAD ANTES DE GASTOS USD": "antes",
    "(−) COMBUSTIBLE DEL MES USD": "combustible",
    "(−) GASTOS INDIRECTOS USD": "indirectos",
    "(−) REFACCIONES (INVENTARIO) USD": "refacciones",
    "(−) PERMISOS USD": "permisos",
    "(−) PENDIENTE DE PAGO (COBRANZA PENDIENTE) USD": "por_cobrar",
}


def _refs_balance_individual(maestra: _Maestra, con_refacciones: bool) -> dict[str, str]:
    """Celdas que la cascada del libro INDIVIDUAL cita: GANANCIA USD y POR
    COBRAR USD de la fila TOTALES de la hoja maestra y el TOTAL USD (fila 5)
    de las hojas combustible / Gastos Indirectos / refacciones / permisos —
    son los mismos números que el API resta (utilidad_antes = Σ ganancia
    USD; cada hoja resta su `usd`)."""
    refs = {
        "antes": maestra.total("ganancia_usd"),
        "combustible": _ref_hoja("combustible", "$E$5"),
        "indirectos": _ref_hoja("Gastos Indirectos", "$C$5"),
        "permisos": _ref_hoja("permisos", "$C$5"),
        "por_cobrar": maestra.total("por_cobrar_usd"),
    }
    if con_refacciones:
        refs["refacciones"] = _ref_hoja("refacciones", "$C$5")
    # Una columna que la hoja no tenga (`total` ⇒ None) deja su fila como
    # valor: nunca una referencia rota.
    return {clave: ref for clave, ref in refs.items() if ref is not None}


@dataclass(frozen=True)
class _CeldaSocioEmpresa:
    """Celda MONTO USD del socio `es_empresa` en el bloque de un avión de la
    hoja 'balance' del GENERAL: el bloque «VUELATOUR (empresa)» la cita."""

    matricula: str
    porcentaje: float | None
    celda: str  # «C15» (misma hoja)
    monto_usd: float | None


def _bloque_balance(
    ws: Worksheet,
    req: BalanceAvionRequest,
    row: int,
    *,
    general: bool = False,
    refs: dict[str, str] | None = None,
    celdas_empresa: list[_CeldaSocioEmpresa] | None = None,
    regla: str | None = None,
) -> int:
    """Bloque utilidad + reparto de socios de UN avión, empezando en `row`.
    Devuelve la siguiente fila libre (el balance GENERAL apila un bloque por
    avión — los socios son POR avión, no hay un reparto de flota).
    Desde el 2-sep-2026 la cascada pinta UNA sola fila "(−) GASTOS
    INDIRECTOS USD" = gastos_indirectos_usd + otros_usd (suma None-tolerante
    de dos celdas que YA manda el API). La cascada la calculó el API:
    utilidad_despues_usd ya resta ambas listas UNA vez — pintar otros
    además de indirectos contaría doble; por eso ya no hay fila de OTROS
    GASTOS. Si otros_usd ≠ 0 la celda lleva una NOTA con cuánto de ese
    total es reparto manual. `general` solo cambia el texto de esa nota:
    en el general el detalle de esa parte es la hoja 'repartidos a
    aviones' (la hoja 'otros gastos' de ese libro es la de EMPRESA y NO
    resta aquí); en el individual son las filas "reparto manual" de la
    hoja 'Gastos Indirectos'. Las demás filas (combustible, refacciones,
    permisos, utilidad después, pendiente de pago, utilidad cobrada,
    socios) no cambian.
    Fórmulas (5-oct-2026): UTILIDAD DESPUÉS = ANTES − las cinco restas,
    UTILIDAD COBRADA = DESPUÉS − PENDIENTE DE PAGO y cada socio = ROUND(% ÷
    100 × UTILIDAD COBRADA, 2) — la aritmética del API. Con `refs` (libro
    individual) las filas de arriba citan su hoja; en el GENERAL van como
    valor (son los números de cada avión, que no están en ese libro).
    `celdas_empresa` (6-oct-2026, solo el GENERAL): lista ACUMULADA a la que
    se agrega la celda MONTO USD de cada socio `es_empresa` de este bloque —
    el bloque «VUELATOUR (empresa)» del final la cita con fórmula. No cambia
    nada de lo que se pinta.
    COMISIONES en el GENERAL (7-oct-2026, API 0.0.66; pedido del cliente: las
    comisiones se duplicaban en el general): su hoja de vuelos va antes de
    comisiones, así que con `general` y comisiones del API
    (`_cascada_con_comisiones`) el bloque abre con «UTILIDAD ANTES DE GASTOS
    USD (antes de comisiones)» y «(−) COMISIONES (banco + vendedor) USD»
    (valores del API, con nota; `regla` = la etiqueta de la regla para la
    nota) y UTILIDAD ANTES DE GASTOS = ROUND(antes − comisiones, 2) — el
    número del API de siempre: de ahí para abajo nada cambia, ni un centavo.
    El libro INDIVIDUAL nunca las pinta (ya restan en su columna
    COMISIONES)."""
    b = req.balance
    refs = refs or {}
    fila_ini = row
    previas: list[tuple[str, float | None, bool]] = []
    if general and _cascada_con_comisiones(b):
        previas = [
            (_FILA_ANTES_COMISIONES, b.utilidad_antes_comisiones_usd, False),
            (_FILA_COMISIONES, b.comisiones_usd, False),
        ]
    notas_previas = {
        _FILA_ANTES_COMISIONES: _NOTA_CASCADA_ANTES_COMISIONES,
        _FILA_COMISIONES: _con_regla(_NOTA_CASCADA_COMISIONES, regla),
    }
    antes = f"ROUND(B{fila_ini}-B{fila_ini + 1},2)" if previas else None
    # Filas de la cascada desde UTILIDAD ANTES DE GASTOS (`k`): 0 antes, 1
    # combustible, 2 indirectos, 3 refacciones, 4 permisos, 5 después, 6
    # pendiente, 7 cobrada.
    k = fila_ini + len(previas)
    despues = f"ROUND(B{k}-B{k + 1}-B{k + 2}-B{k + 3}-B{k + 4},2)"
    cobrada = f"ROUND(B{k + 5}-B{k + 6},2)"
    fila_cobrada = k + 7
    filas: list[tuple[str, float | None, bool]] = [
        *previas,
        ("UTILIDAD ANTES DE GASTOS USD", b.utilidad_antes_usd, False),
        ("(−) COMBUSTIBLE DEL MES USD", b.combustible_usd, False),
        # Indirectos + otros (reparto manual) en UNA fila — solo se muestran
        # juntos; la resta ya la hizo el API una sola vez.
        (_FILA_INDIRECTOS, _suma_none(b.gastos_indirectos_usd, b.otros_usd), False),
        # Refacciones (29-ago): salidas de inventario, ANTES dentro de
        # gastos indirectos — indirectos + refacciones == lo de antes.
        # None = API viejo (celda vacía; ya venían dentro de indirectos).
        ("(−) REFACCIONES (INVENTARIO) USD", b.refacciones_usd, False),
        ("(−) PERMISOS USD", b.permisos_usd, False),
        ("UTILIDAD DESPUÉS DE GASTOS USD", b.utilidad_despues_usd, True),
        ("(−) PENDIENTE DE PAGO (COBRANZA PENDIENTE) USD", b.por_cobrar_usd, False),
        ("UTILIDAD COBRADA USD", b.utilidad_cobrada_usd, True),
    ]
    for label, val, bold in filas:
        lc = ws.cell(row=row, column=1, value=label)
        lc.font = Font(bold=bold, color=NAVY if bold else MUTED)
        lc.border = _border
        if label in notas_previas and len(label) > _ANCHO_ETIQUETA_BALANCE:
            # Etiqueta larga: envuelve, como las del bloque VUELATOUR (empresa).
            lc.alignment = Alignment(wrap_text=True, vertical="center")
            ws.row_dimensions[row].height = 15 * math.ceil(len(label) / _ANCHO_ETIQUETA_BALANCE) + 1
        color = None
        if label.startswith("UTILIDAD COBRADA") and val is not None:
            color = GREEN if val >= 0 else RED
        if label == "UTILIDAD DESPUÉS DE GASTOS USD":
            formula = despues
        elif label == "UTILIDAD COBRADA USD":
            formula = cobrada
        elif label == "UTILIDAD ANTES DE GASTOS USD" and antes is not None:
            formula = antes
        else:
            formula = refs.get(_CASCADA_REFS.get(label, ""))
        cell = _formula(
            ws, row, 2, formula, val, MONEY, bold=bold,
            color=color or ("000000" if bold else MUTED),
        )
        cell.border = _border
        if label.startswith("UTILIDAD"):
            _rojo_si_negativo(ws, cell.coordinate)
        if label == _FILA_INDIRECTOS and b.otros_usd is not None and b.otros_usd != 0:
            cell.comment = _nota_otros_en_indirectos(b.otros_usd, general)
        if label in notas_previas and val is not None:
            cell.comment = _comentario_cobro(notas_previas[label])
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
            mc = _formula(
                ws, row, 3, f"ROUND(B{row}/100*$B${fila_cobrada},2)",
                s.monto_usd, MONEY, bold=True,
            )
            mc.border = _border
            _rojo_si_negativo(ws, mc.coordinate)
            if celdas_empresa is not None and s.es_empresa:
                celdas_empresa.append(
                    _CeldaSocioEmpresa(
                        req.matricula or "—", s.porcentaje, mc.coordinate, s.monto_usd
                    )
                )
            row += 1
    else:
        ws.cell(
            row=row, column=1, value="Sin socios configurados para este avión (ver pendientes)."
        ).font = Font(color=RED, italic=True)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        row += 1
    return row


# Bloque «VUELATOUR (empresa)» al final de la hoja 'balance' del GENERAL
# (6-oct-2026, pedido del cliente: «en la hoja de balance falta, hasta el
# final, el balance de la empresa VuelaTour»). Etiquetas de sus filas:
_EMPRESA_TITULO = "VUELATOUR (empresa)"
_EMPRESA_PARTICIPACION = "(+) PARTICIPACIÓN COMO SOCIO EN LOS AVIONES USD"
# Revisión 6-oct-2026: las dos filas de 'otros movimientos' NO son caja. El
# ingreso es lo COTIZADO (desglose v1.3) de TODOS los vuelos del periodo no
# cancelados —COTIZADO y RESERVA incluidos— y el egreso es TODO lo que esa
# hoja resta (pago o PROVISIÓN al vendedor, TUAs pagadas, extensión de
# horario, comisión bancaria, gas sin avión, TUAS sin vuelo): las etiquetas
# lo dicen para que nadie lea el RESULTADO como dinero en caja.
_EMPRESA_INGRESOS = (
    "(+) INGRESOS PROPIOS USD (TUAs, extras, pernocta, comisión — cotizado del periodo)"
)
_EMPRESA_PAGOS = (
    "(−) EGRESOS PROPIOS USD (pago al vendedor, TUAs, extensión, comisión bancaria, sueltos)"
)
_EMPRESA_OTROS = "(−) OTROS GASTOS DE LA EMPRESA USD"
_EMPRESA_TIENDA = "(+) UTILIDAD TIENDA (INVENTARIO) USD"
_EMPRESA_RESULTADO = "RESULTADO VUELATOUR USD"
_EMPRESA_SIN_PARTICIPACION = "   sin participación registrada"
_NOTA_EMPRESA_OTROS_MOVIMIENTOS = (
    "MXN de la hoja 'otros movimientos' a USD con el TC de cada vuelo "
    "(sueltas: su TC o el oficial del día)"
)
_NOTA_EMPRESA_TIENDA = "utilidad de la hoja 'inventario' / TC promedio"
# Base de las filas de 'otros movimientos' (lo único que la `nota` del API no
# dice): va SIEMPRE al pie del bloque.
_NOTA_EMPRESA_BASE = (
    "Ingresos y egresos propios = hoja 'otros movimientos': lo cotizado de "
    "los vuelos del periodo (no cancelados) más los otros ingresos "
    "registrados, y lo pagado o provisionado — no es lo cobrado: el "
    "RESULTADO no es dinero en caja."
)
# Pie de RESPALDO para un payload sin `nota` (con ella, el API ya explica
# participación, otros gastos, tienda y gastos personales: repetirlo era
# pintar dos párrafos casi iguales).
_NOTA_BLOQUE_EMPRESA = (
    "Participación = utilidad COBRADA de cada avión × % de Aero Charter "
    "Cancún. " + _NOTA_EMPRESA_BASE + " Otros gastos = hoja 'otros gastos'. "
    "Tienda = margen de las salidas de inventario. Los gastos personales del "
    "dueño no entran."
)
# Línea del RESUMEN que apunta al bloque (solo con `empresa`). Corta: las
# notas del RESUMEN van combinadas en A:L SIN envolver.
_RESUMEN_BLOQUE_EMPRESA = (
    "Al final de la hoja 'balance': bloque VUELATOUR (empresa) — participación "
    "como socia, ingresos y egresos propios, otros gastos, tienda y su RESULTADO."
)
# Caracteres que caben en la columna A (ancho 46) de la hoja 'balance' con
# etiquetas en MAYÚSCULAS: arriba de eso la etiqueta se envuelve.
_ANCHO_ETIQUETA_BALANCE = 40
# Ancho total A:C de la hoja 'balance' (46 + 14 + 16) para las notas.
_ANCHO_HOJA_BALANCE = 76


def _nota_celda(texto: str) -> Comment:
    com = Comment(texto, "VuelaTour")
    com.width = 320
    com.height = 80
    return com


def _pct_socio(pct: float | None) -> str:
    return "" if pct is None else f" · {pct:.2f} %"


def _bloque_empresa(
    ws: Worksheet,
    emp: BalanceEmpresaBloque,
    row: int,
    celdas: list[_CeldaSocioEmpresa],
    regla_vigente: str | None = None,
    *,
    cubre_avion: bool = False,
) -> int:
    """Bloque «VUELATOUR (empresa)» de la hoja 'balance' del GENERAL, después
    del último bloque de avión (6-oct-2026). Mismo estilo que la cascada de
    los aviones; los números los calculó el API (`empresa`, 0.0.59) y aquí
    solo se pintan. Fórmulas verificadas con `xlsx_formulas` (si no
    reproducen el número del API, la celda queda como valor):
      PARTICIPACIÓN = ROUND(SUM(celda MONTO USD del socio `es_empresa` de
        cada bloque de arriba), 2) — `celdas`, que acumuló _bloque_balance
        (bloques con distinto número de socios: la celda es la real); debajo
        una línea gris por avión que cita esa celda.
      OTROS GASTOS = 'otros gastos'!$C$5 (TOTAL USD de esa hoja; sin la
        hoja, valor).
      RESULTADO = ROUND(part + ingresos − pagos − otros + tienda, 2).
    Ingresos propios, pagos al vendedor y tienda van como valor (salen de
    filas en MXN con el TC de cada vuelo, que no están en esta hoja). Con
    vuelos de la regla de comisiones a cargo del avión (`regla_vigente`) la
    base del pie lo aclara (`_EMPRESA_REGLA_COMISIONES`; con el API 0.0.66,
    `cubre_avion`: `_EMPRESA_REGLA_COMISIONES_CUBRE_AVION`).
    Devuelve la siguiente fila libre."""
    t = ws.cell(row=row, column=1, value=_EMPRESA_TITULO)
    t.font = Font(bold=True, size=12, color="FFFFFF")
    t.fill = PatternFill("solid", fgColor=NAVY)
    row += 1

    def etiqueta(r: int, texto: str, *, bold: bool = False, detalle: bool = False) -> None:
        c = ws.cell(row=r, column=1, value=texto)
        c.font = Font(bold=bold, italic=detalle, color=NAVY if bold else MUTED)
        c.border = _border
        lineas = math.ceil(len(texto) / _ANCHO_ETIQUETA_BALANCE)
        if lineas > 1:
            c.alignment = Alignment(wrap_text=True, vertical="center")
            ws.row_dimensions[r].height = 15 * lineas + 1

    def monto(
        r: int,
        formula: str | None,
        valor: float | None,
        *,
        bold: bool = False,
        detalle: bool = False,
        color: str | None = None,
        nota: str | None = None,
    ) -> None:
        color = color or ("000000" if bold else MUTED)
        cell = _formula(ws, r, 2, formula, valor, MONEY, bold=bold, italic=detalle, color=color)
        cell.border = _border
        if nota and valor is not None:
            cell.comment = _nota_celda(nota)

    # (+) Participación como socio: Σ de las celdas de los bloques de arriba.
    fila_part = row
    etiqueta(row, _EMPRESA_PARTICIPACION)
    if celdas:
        refs = ",".join(c.celda for c in celdas)
        monto(row, f"ROUND(SUM({refs}),2)", emp.participacion_usd)
    else:
        # Sin celda que citar: el número del API tal cual. Sin socio empresa
        # el API manda 0 (0.00); un null es un payload a medias y queda
        # VACÍO, como el RESULTADO —jamás un 0 falso (revisión 6-oct-2026).
        monto(row, None, emp.participacion_usd)
    row += 1
    if celdas:
        for c in celdas:
            etiqueta(row, f"   {c.matricula}{_pct_socio(c.porcentaje)}", detalle=True)
            monto(row, c.celda, c.monto_usd, detalle=True)
            row += 1
    elif emp.participaciones:
        # El API manda participaciones pero ningún socio trae `es_empresa`
        # (skew): las líneas van como valor, sin celda que citar.
        for p in emp.participaciones:
            etiqueta(row, f"   {p.matricula or '—'}{_pct_socio(p.porcentaje)}", detalle=True)
            monto(row, None, p.monto_usd, detalle=True)
            row += 1
    else:
        etiqueta(row, _EMPRESA_SIN_PARTICIPACION, detalle=True)
        ws.cell(row=row, column=2).border = _border
        row += 1

    fila_ing = row
    etiqueta(row, _EMPRESA_INGRESOS)
    monto(row, None, emp.ingresos_propios_usd, nota=_NOTA_EMPRESA_OTROS_MOVIMIENTOS)
    row += 1

    fila_pag = row
    etiqueta(row, _EMPRESA_PAGOS)
    monto(row, None, emp.pagos_vendedor_usd, nota=_NOTA_EMPRESA_OTROS_MOVIMIENTOS)
    row += 1

    fila_otros = row
    etiqueta(row, _EMPRESA_OTROS)
    hoja_otros = "otros gastos" in ws.parent.sheetnames
    nota_otros = None
    if emp.tc_usado is not None:
        nota_otros = (
            "TOTAL USD de la hoja 'otros gastos' (total MXN / TC promedio de "
            f"la flota, {_tc_txt(emp.tc_usado)})"
        )
    formula_otros = _ref_hoja("otros gastos", "$C$5") if hoja_otros else None
    monto(row, formula_otros, emp.otros_gastos_empresa_usd, nota=nota_otros)
    row += 1

    fila_tienda = row
    etiqueta(row, _EMPRESA_TIENDA)
    nota_tienda = _NOTA_EMPRESA_TIENDA + (
        f" ({_tc_txt(emp.tc_promedio)})" if emp.tc_promedio is not None else ""
    )
    monto(row, None, emp.tienda_utilidad_usd, nota=nota_tienda)
    row += 1

    etiqueta(row, _EMPRESA_RESULTADO, bold=True)
    res = emp.resultado_usd
    monto(
        row,
        f"ROUND(B{fila_part}+B{fila_ing}-B{fila_pag}-B{fila_otros}+B{fila_tienda},2)",
        res,
        bold=True,
        color=None if res is None else (GREEN if res >= 0 else RED),
    )
    _rojo_si_negativo(ws, f"B{row}")
    row += 2

    # Pie: la `nota` del API (base de cada línea con sus números) + la base
    # de 'otros movimientos'; sin `nota`, el pie fijo completo.
    regla_empresa = (
        _EMPRESA_REGLA_COMISIONES_CUBRE_AVION if cubre_avion else _EMPRESA_REGLA_COMISIONES
    )
    base = _NOTA_EMPRESA_BASE + (regla_empresa.format(regla=regla_vigente) if regla_vigente else "")
    if emp.nota:
        for texto in (emp.nota, base):
            _pinta_nota_envuelta(ws, row, texto, 3, _ANCHO_HOJA_BALANCE)
            row += 1
    else:
        pie = _NOTA_BLOQUE_EMPRESA.replace(_NOTA_EMPRESA_BASE, base)
        _pinta_nota_envuelta(ws, row, pie, 3, _ANCHO_HOJA_BALANCE)
        row += 1
    return row


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


# Definiciones de las hojas de gastos (regla del cliente, 28-ago-2026).
# Nota de la hoja "Gastos Indirectos" del libro INDIVIDUAL (2-sep-2026:
# fusiona la vieja nota de 'gastos indirectos' con la de 'otros gastos').
_NOTA_INDIRECTOS = (
    "GASTOS INDIRECTOS = todo lo que se cargó a este avión SIN un vuelo: (1) "
    "gastos capturados con la aeronave y sin vuelo, de cualquier categoría "
    "salvo combustible (hoja 'combustible'), permisos (hoja 'permisos') y "
    "salidas de inventario (hoja 'refacciones'); y (2) la parte de este "
    "avión de los gastos administrativos de la empresa repartidos a mano "
    "(nómina, IMSS, pensión, fijos…): son las filas cuyo DETALLE dice "
    "reparto manual: $X de $Y. Todo resta UNA sola vez en la hoja "
    "'balance'. El TUA pagado NO va aquí (regla 28-ago-2026): queda solo "
    "como nota en OPERACIONES y vive en 'otros movimientos' del Balance "
    "general VuelaTour. Hasta el 2-sep-2026 la parte (2) era la hoja 'otros "
    "gastos' de este libro."
)
_NOTA_PERMISOS = "PERMISOS = todo lo de AFAC (permisos y provisiones del avión)."
_NOTA_REFACCIONES = (
    "REFACCIONES = salidas de INVENTARIO (bodega) cargadas al avión: gasto "
    "REFACCION medio BODEGA ligado al cardex. Antes vivían dentro de 'gastos "
    "indirectos' (29-ago-2026: indirectos + refacciones = lo de antes; la "
    "utilidad no cambia) y restan con renglón propio en la hoja 'balance'. "
    "El dinero salió del banco al COMPRAR la pieza, no al consumirla — la "
    "conciliación bancaria no las cruza."
)
# Solo se pinta en el FALLBACK del general (payload SIN `inventario`, API
# anterior al 30-ago-2026, que costeaba FIFO): por eso conserva «costo FIFO».
# Con el API 0.0.36 (último precio) el general siempre trae `inventario` y
# esta hoja ya no existe.
_NOTA_REFACCIONES_GENERAL = (
    _NOTA_REFACCIONES
    + " En este Balance general cada fila trae además COSTO VUELATOUR "
    "(costo FIFO de la salida), VENTA AL AVIÓN (lo cargado al avión) y "
    "GANANCIA = venta − costo (0 mientras la salida se cargue a costo, aún "
    "sin precio de venta)."
)
# Hoja "inventario" del general (tiendita, 30-ago-2026): sustituye a la hoja
# 'refacciones' del general — su detalle de salidas es el BLOQUE 2 de esta.
_NOTA_INVENTARIO = (
    "INVENTARIO (tiendita) — BLOQUE POR ÍTEM: EXISTENCIA y VALOR A COSTO son "
    "A HOY (todo el cardex FIFO), NO una foto al corte del periodo; COMPRAS "
    "= solo ENTRADAs del periodo (una devolución o un ajuste regresan stock "
    "pero no son compra); VENDIDO = lo cargado a los aviones en salidas CON "
    "precio de venta y UTILIDAD = vendido − costo FIFO consumido (las "
    "salidas a costo no llevan venta ni utilidad). DETALLE DE SALIDAS "
    "(bloque 2) = las salidas de bodega cargadas a aviones en el periodo: "
    "gasto REFACCION medio BODEGA ligado al cardex, con GANANCIA = venta − "
    "costo (0 mientras la salida se cargue a costo); antes era la hoja "
    "'refacciones' de este libro (el libro INDIVIDUAL de cada avión la "
    "conserva). El dinero salió del banco al COMPRAR la pieza, no al "
    "consumirla — la conciliación bancaria no cruza estos cargos."
)
# Aclaración que se AÑADE a la nota anterior solo cuando la hoja trae la
# columna en dólares (22-sep-2026): el cliente vio «30 · $3,300.00 MXN» en
# un ítem cuya única entrada fueron 30 × 110 USD sin tipo de cambio — la
# columna sumaba dólares y los rotulaba pesos. Invariante: jamás un USD
# sumado como MXN.
_NOTA_INVENTARIO_USD = (
    " VALOR A COSTO MXN suma SOLO las capas compradas en pesos (o en USD "
    "con tipo de cambio capturado): lo comprado en dólares que todavía no "
    "tiene T.C. se muestra aparte, EN DÓLARES, en la columna VALOR A COSTO "
    "USD (sin T.C.) y no se convierte ni entra al total en pesos. Al "
    "capturar el tipo de cambio de esas entradas su valor pasa solo a la "
    "columna en pesos."
)
# Aclaración que se AÑADE a la nota cuando la hoja trae VENDIDO/UTILIDAD en
# dólares (25-sep-2026, utilidad de la tienda: todas las ventas de prod son
# USD sobre costo USD sin T.C. y antes su utilidad no aparecía en ninguna
# columna). Misma invariante: jamás un USD sumado como MXN.
_NOTA_INVENTARIO_VENTA_USD = (
    " VENDIDO USD y UTILIDAD USD: salidas cobradas al avión en dólares sobre "
    "costo en dólares; se muestran en su moneda y no se suman con las "
    "columnas en pesos."
)

# ---------------------------------------------------------------------------
# Regla ÚLTIMO PRECIO DE COMPRA + T.C. OFICIAL DEL DÍA (25-sep-2026, API
# 0.0.36, payload con `regla_costo='ULTIMO_PRECIO'`). Pedido del cliente:
# «que los precios se ajusten en automático al último registrado» y «en el
# tipo de cambio, que sea los mismos que usan en las cotizaciones (tipo de
# cambio del día de la venta)». Los NÚMEROS los manda el API ya convertidos;
# aquí solo cambian los TEXTOS (ninguna nota dice «FIFO»). Payload previo
# (sin `regla_costo`) ⇒ textos de siempre y hoja idéntica (huellas
# congeladas en tests/test_balance_inventario_xlsx.py).
# ---------------------------------------------------------------------------
_NOTA_INVENTARIO_ULTIMO_PRECIO = (
    "INVENTARIO (tiendita) — BLOQUE POR ÍTEM: EXISTENCIA y VALOR A COSTO son "
    "A HOY (todo el cardex), NO una foto al corte del periodo. {valor} Las "
    "compras y ventas en dólares se convierten con el T.C. oficial de su día "
    "(el mismo de las cotizaciones): la compra con el del día de la compra; "
    "la venta y el costo de esa pieza con el del día de la venta. COMPRAS = "
    "solo ENTRADAs del periodo (una devolución o un ajuste regresan stock "
    "pero no son compra); VENDIDO = lo cargado a los aviones en salidas CON "
    "precio de venta y UTILIDAD = vendido − costo de la pieza (el último "
    "precio de compra vigente el día de la salida; las salidas a costo no "
    "llevan venta ni utilidad). DETALLE DE SALIDAS (bloque 2) = las salidas "
    "de bodega cargadas a aviones en el periodo: gasto REFACCION medio BODEGA "
    "ligado al cardex, con GANANCIA = venta − costo (0 si la salida se cargó "
    "a costo); antes era la hoja 'refacciones' de este libro (el libro "
    "INDIVIDUAL de cada avión la conserva). El dinero salió del banco al "
    "COMPRAR la pieza, no al consumirla — la conciliación bancaria no cruza "
    "estos cargos."
)
# Con la regla nueva el valorizado usa el T.C. de HOY: una columna en
# dólares solo aparece si algún producto con último precio en USD no se pudo
# pasar a pesos. No se afirma la causa (revisión adversaria del contrato:
# «sin T.C.» puede venir de más de un lado).
_NOTA_INVENTARIO_USD_ULTIMO_PRECIO = (
    " VALOR A COSTO USD (sin T.C.): productos cuyo último precio de compra "
    "es en dólares y que no se pudieron pasar a pesos por falta de tipo de "
    "cambio; se muestran aparte, EN DÓLARES, y no entran al total en pesos."
)
# Con la regla nueva VENDIDO/UTILIDAD USD son SOLO el respaldo de ventas que
# siguen sin T.C.; las demás ya cuentan en pesos (jamás las dos a la vez).
_NOTA_INVENTARIO_VENTA_USD_ULTIMO_PRECIO = (
    " VENDIDO USD y UTILIDAD USD: ventas en dólares que todavía no tienen "
    "tipo de cambio; se muestran en su moneda y no se suman con las columnas "
    "en pesos (las que sí lo tienen ya cuentan en pesos)."
)
# Nota bajo el BLOQUE 2 (solo con la regla nueva y si hay salidas): el
# detalle de salidas sale de los GASTOS del balance (T.C. de cada gasto o el
# promedio del libro, igual que la hoja 'balance') y la utilidad por ítem
# del bloque 1 del T.C. oficial del día de la venta. En las 10 salidas del
# 01-sep-2026 (gastos sin T.C., convertidos al T.C. promedio del libro) las
# dos utilidades en pesos difieren unos pesos (≈ $63 MXN en septiembre): la
# propia hoja lo explica en vez de que el cliente lo descubra. «Del
# periodo», no «del mes» (revisión adversaria 25-sep-2026): el respaldo es
# el TC PROMEDIO del libro (`tc_promedio`) para el rango pedido, que no
# siempre es un mes.
_NOTA_BLOQUE2_TC = (
    "El detalle de salidas convierte con el T.C. de cada gasto (o el T.C. "
    "promedio del periodo si el gasto no lo trae), igual que la hoja balance; "
    "la utilidad por ítem de arriba usa el T.C. oficial del día de la venta. "
    "En salidas registradas antes del 25-sep-2026 pueden diferir unos pesos."
)


def _es_ultimo_precio(inv: BalanceHojaInventario | None) -> bool:
    """¿El API mandó la regla de costo nueva (último precio de compra)?
    Ausente / otro valor = API previo (FIFO) ⇒ textos de siempre."""
    if inv is None or not inv.regla_costo:
        return False
    return inv.regla_costo.strip().upper() == REGLA_COSTO_ULTIMO_PRECIO


def _frase_valor_a_costo(tc_hoy: TcHoy | None) -> str:
    """Cómo se valorizó la bodega: existencia × último precio de compra al
    T.C. oficial de HOY. Cita el T.C. y la fecha de SU dato tal como los
    mandó el API (el T.C. con `_tc_txt`, fuente única del texto de un T.C.;
    la fecha con `_fecha`, el dd/mm/aaaa de todo este libro). Sin dato no
    se inventa un número."""
    tc = tc_hoy.tc if tc_hoy is not None else None
    if tc is None or tc <= 0:
        return ("VALOR A COSTO = existencia × último precio de compra, en "
                "pesos al T.C. oficial de hoy.")
    fecha = _fecha(tc_hoy.fecha_dato) if tc_hoy is not None else None
    cita = _tc_txt(tc) + (f", {fecha}" if fecha else "")
    return ("VALOR A COSTO = existencia × último precio de compra, al T.C. "
            f"oficial de hoy ({cita}).")


def _alto_nota(texto: str, ancho_total: float) -> float | None:
    """Alto de fila para una nota ENVUELTA (letra 9) que abarca columnas de
    `ancho_total` caracteres: mismo patrón de estimación que DETALLE en
    _hoja_gastos (renglones × alto de renglón). Un renglón ⇒ None (alto por
    omisión). Estima de más, nunca de menos: sobra aire, no se corta."""
    lineas = max(1, math.ceil(len(texto) / max(ancho_total, 1)))
    return 12 * lineas + 4 if lineas > 1 else None


def _pinta_nota_envuelta(ws: Worksheet, row: int, texto: str, n_cols: int,
                         ancho_total: float) -> None:
    """Nota tenue (gris, cursiva, 9) combinada en `n_cols` columnas y
    ENVUELTA, con el alto de fila estimado — sin el wrap, una nota combinada
    se corta al ancho de las celdas (lo importante suele ir al final)."""
    c = ws.cell(row=row, column=1, value=texto)
    c.font = Font(color=MUTED, size=9, italic=True)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n_cols)
    alto = _alto_nota(texto, ancho_total)
    if alto:
        ws.row_dimensions[row].height = alto
# Nota de la hoja 'otros gastos' del GENERAL (payload `gastos_empresa`;
# 1-sep-2026: la pestaña se llamaba 'gastos VuelaTour' — solo cambió el
# NOMBRE de la hoja, el campo del contrato sigue igual).
_NOTA_GASTOS_EMPRESA = (
    "OTROS GASTOS = gastos de la EMPRESA del periodo NO asignados a ningún "
    "avión: sin vuelo y sin avión (sin PERSONAL_DUENO ni GAS) que nadie "
    "reparte, más el REMANENTE de los repartidos a mano. Son egresos de "
    "VuelaTour: NO restan en la cascada de ningún avión. Se pueden repartir "
    "a aviones desde 'Otros gastos' del panel — la parte asignada a cada "
    "avión aparece en la hoja 'repartidos a aviones' de este libro (y dentro "
    "de la hoja 'Gastos Indirectos' del libro individual del avión). Antes "
    "salían como "
    "'MOVIMIENTOS SIN AVIÓN / SIN VUELO' en 'otros movimientos' (allá "
    "quedan 'gas sin avión', 'tuas sin vuelo' y los otros ingresos "
    "registrados en Ingresos, clave ING-n)."
)


# Nota de la hoja 'repartidos a aviones' del GENERAL (1-sep-2026: antes se
# llamaba 'otros gastos' y confundía — se veía vacía sin repartos). Solo
# existe en el general: en el libro INDIVIDUAL esa lista va fusionada en la
# hoja "Gastos Indirectos" (2-sep-2026), por eso ya no hay variante
# individual de esta nota.
_NOTA_REPARTIDOS_A_AVIONES_BASE = (
    "REPARTIDOS A AVIONES = gastos administrativos de la empresa (nómina, "
    "IMSS, pensión, fijos…) repartidos a mano entre aviones — aquí solo la "
    "parte asignada a cada avión (fila con su color); en el libro INDIVIDUAL "
    "de cada avión esta parte está dentro de su hoja 'Gastos Indirectos' "
    "(filas 'reparto manual') y NO resta aparte: va sumada en la fila "
    "'(−) GASTOS INDIRECTOS USD' de la hoja 'balance'. Lo que NADIE ha "
    "repartido (y el remanente) vive en la hoja 'otros gastos' de este "
    "libro. El TUA pagado NO va aquí (regla 28-ago-2026): queda solo como "
)
_NOTA_REPARTIDOS_A_AVIONES = (
    _NOTA_REPARTIDOS_A_AVIONES_BASE + "nota en OPERACIONES y vive en 'otros movimientos'."
)
# Variante «Balance general» (6-oct-2026): su hoja de vuelos ya no tiene la
# columna OPERACIONES; el TUA pagado queda en la nota del desglose de TOTAL
# PARA PROVEEDOR (PESOS), como dicen sus notas al pie (`_CPH_TRASLADOS_*`).
_CPH_NOTA_REPARTIDOS_A_AVIONES = _NOTA_REPARTIDOS_A_AVIONES_BASE + (
    "nota en el desglose de TOTAL PARA PROVEEDOR (PESOS) de la hoja de vuelos "
    "(parte Operación) y vive en 'otros movimientos'."
)


def _hay_usd_sin_tc(inv: BalanceHojaInventario) -> bool:
    """¿La hoja debe llevar la columna en DÓLARES? (22-sep-2026)

    Solo cuando el API manda el desglose nuevo Y hay algo que mostrar en él:
    así un API viejo (que no manda ninguno de los campos) sigue produciendo
    la hoja de siempre, y un inventario 100 % en pesos no gana una columna
    de ceros."""
    if inv.filas_sin_tc:
        return True
    if inv.total_valor_usd:  # None o 0 → no hay nada que mostrar
        return True
    return any(f.valor_costo_usd or f.sin_tc for f in inv.filas)


def _cuenta_sin_tc(inv: BalanceHojaInventario) -> int:
    """Cuántos productos traen costo en USD sin T.C. Manda el número del API
    (`filas_sin_tc`); si no viene se cuentan las filas (solo para el texto de
    la nota — aquí no se recalcula dinero)."""
    if inv.filas_sin_tc:
        return inv.filas_sin_tc
    return sum(1 for f in inv.filas if f.valor_costo_usd or f.sin_tc)


def _nota_sin_tc(n: int) -> str:
    """Nota al pie de la tabla, en el idioma del operador."""
    plural = n != 1
    return (
        f"{n} producto{'s' if plural else ''} "
        f"{'tienen' if plural else 'tiene'} costo en USD sin tipo de cambio: "
        "su valor se muestra en dólares y no entra al total en pesos."
    )


def _hay_venta_usd(inv: BalanceHojaInventario) -> bool:
    """¿La hoja lleva VENDIDO USD / UTILIDAD USD? (25-sep-2026)

    Solo si el API manda utilidad de la tienda en dólares: un API viejo (sin
    los campos) o una tienda que solo vende en pesos dejan la hoja como
    antes. Un 0.0 SÍ cuenta (es dato: vendido a costo en dólares); None no."""
    if inv.total_utilidad_usd is not None or inv.total_vendido_usd is not None:
        return True
    return any(f.vendido_usd is not None or f.utilidad_usd is not None
               for f in inv.filas)


def _total_usd_de_columna(total: float | None, valores: list[float | None]
                          ) -> float | None:
    """Total de una columna en DÓLARES: el que manda el API; si faltara, la
    re-suma de ESA MISMA columna (misma moneda — solo para mostrar). Sin
    ningún valor ⇒ None (celda vacía, nunca un 0 falso)."""
    if total is not None:
        return total
    vals = [v for v in valores if v is not None]
    return round(sum(vals), 2) if vals else None


def _cuenta_utilidad_incompleta(inv: BalanceHojaInventario) -> int:
    """Productos con ventas cuya utilidad no se puede calcular. Manda el
    número del API; si no viene se cuentan las filas (solo para el texto)."""
    if inv.filas_utilidad_incompleta:
        return inv.filas_utilidad_incompleta
    return sum(1 for f in inv.filas if f.ventas_sin_utilidad > 0)


def _nota_utilidad_incompleta(n: int) -> str:
    """Nota roja bajo la tabla: venta en pesos sobre costo en dólares sin
    T.C. ⇒ no hay utilidad que mostrar en ninguna moneda."""
    plural = n != 1
    return (
        f"{n} producto{'s' if plural else ''} con ventas en pesos sobre costo "
        "en dólares sin tipo de cambio: su utilidad no se puede calcular y no "
        "aparece en ninguna columna."
    )


def _pct_txt(pct: float) -> str:
    """25.0 → «25»; 12.5 → «12.5» (hasta 4 decimales, sin ceros de cola)."""
    txt = f"{float(pct):.4f}".rstrip("0").rstrip(".")
    return txt or "0"


def _nota_margen(pct: float, *, ultimo_precio: bool = False) -> str:
    """Cómo se cobra hoy una salida sin precio (config del API
    `inventario_margen_venta_pct`). Aquí solo se redacta: el número llega.
    `ultimo_precio` (API 0.0.36): la base es el último precio de compra;
    sin él, el texto de siempre (costo FIFO del API previo)."""
    base = "al último precio de compra" if ultimo_precio else "a costo FIFO"
    if pct <= 0:
        return (f" Toda salida sin precio se cobra al avión {base}, sin "
                "utilidad (margen de la tienda en 0 %).")
    return (" Desde el 25-sep-2026 toda salida sin precio se cobra al avión "
            f"{base} + {_pct_txt(pct)} % (utilidad de la tienda).")


def _hoja_inventario(ws: Worksheet, inv: BalanceHojaInventario,
                     refacciones: BalanceAvionHojaGastos | None,
                     req: BalanceAvionRequest) -> None:
    """Hoja 'inventario' del Balance GENERAL (tiendita, 30-ago-2026).

    BLOQUE 1 = resumen POR ÍTEM del periodo (existencia y valor a costo A
    HOY; compras / salidas / vendido / utilidad del periodo) con fila
    TOTALES. BLOQUE 2 = el detalle de salidas que antes pintaba la hoja
    'refacciones' del general (`refacciones.filas`; el libro INDIVIDUAL
    conserva su hoja). Todo viene YA calculado del API — aquí solo se pinta
    (los totales del bloque 2 son re-suma de columnas para mostrar);
    None = celda vacía, nunca un 0 falso.

    22-sep-2026 (invariante: jamás un USD sumado como MXN): cuando el API
    manda el desglose nuevo, VALOR A COSTO MXN lleva solo pesos reales y a
    su derecha aparece VALOR A COSTO USD (sin T.C.) con lo comprado en
    dólares que aún no tiene tipo de cambio, con su propio total y una nota
    bajo la tabla. Sin esos campos (API viejo) la hoja es la de siempre.

    25-sep-2026 (utilidad de la tienda, costo + margen): cuando el API
    manda ventas en DÓLARES entran VENDIDO USD y UTILIDAD USD entre UTILIDAD
    MXN y MATRÍCULAS, con sus totales APARTE (jamás sumados con los de
    pesos); nota roja si hay productos cuya utilidad no se puede calcular y,
    al pie, el margen vigente. Es un segundo desplazamiento, independiente
    del de VALOR A COSTO USD: las dos columnas en dólares conviven.

    25-sep-2026, API 0.0.36 (`regla_costo='ULTIMO_PRECIO'`): costo = último
    precio de compra y todo en PESOS con el T.C. oficial del día (lo
    convierte el API). Aquí cambian SOLO los textos: nota al pie sin «FIFO»
    que cita el T.C. de hoy (`tc_hoy`), nota bajo el bloque 2 (el detalle
    de salidas usa el T.C. de cada gasto, la utilidad por ítem el del día
    de la venta) y las dos notas ENVUELTAS con su alto de fila. Las columnas
    en dólares siguen la misma regla de siempre: tras la migración de T.C.
    no llega ningún USD sin T.C. y se apagan solas. Sin `regla_costo` la
    hoja es byte a byte la de antes."""
    usd = _hay_usd_sin_tc(inv)
    venta_usd = _hay_venta_usd(inv)
    ultimo_precio = _es_ultimo_precio(inv)
    # Desplazamiento de las columnas del periodo cuando entra la de dólares.
    off = 1 if usd else 0
    # VENDIDO USD + UTILIDAD USD (25-sep) van justo antes de MATRÍCULAS: solo
    # esa columna corre (dos lugares) — las del periodo no se mueven.
    off_venta = 2 if venta_usd else 0
    col_matriculas = 9 + off + off_venta
    n_cols = col_matriculas
    # Anchos para estimar el alto de las filas que envuelven texto (mismo
    # patrón que _hoja_gastos): ÍTEM del bloque 1 y detalle del bloque 2.
    ancho_item = 34
    ancho_detalle = 42
    # Anchos compartidos entre los dos bloques (col 1 = ÍTEM/FECHA, col 2 =
    # números del bloque 1 y detalle del bloque 2 — por eso va ancha). Con
    # la columna en dólares (22-sep) se inserta su ancho en la posición 4 y
    # las del periodo corren un lugar; con VENDIDO/UTILIDAD USD (25-sep) se
    # insertan dos anchos antes del de MATRÍCULAS. Se calculan aquí (se
    # aplican al final) porque las notas envueltas estiman su alto con ellos.
    anchos = [ancho_item, ancho_detalle, 16, *([18] if usd else []),
              12, 14, 12, 14, 14, *([14, 14] if venta_usd else []), 24]
    _title(ws, f"Inventario (tiendita) — {req.matricula}".strip(" —"), 1, n_cols)
    periodo = f"Periodo: {req.periodo_desde or '—'} a {req.periodo_hasta or '—'}"
    ws.cell(row=2, column=1, value=periodo).font = Font(italic=True, size=10, color=MUTED)

    # ===== Bloque 1: resumen por ítem =====
    _header_row(ws, 4, [
        "ÍTEM", "EXISTENCIA\nACTUAL", "VALOR A\nCOSTO MXN",
        *(["VALOR A COSTO\nUSD (sin T.C.)"] if usd else []),
        "COMPRADAS\nCANT",
        "COMPRAS\nMXN", "SALIDAS\nCANT", "VENDIDO\nMXN", "UTILIDAD\nMXN",
        *(["VENDIDO\nUSD", "UTILIDAD\nUSD"] if venta_usd else []),
        "MATRÍCULAS",
    ])
    row = 5
    if inv.filas:
        for f in inv.filas:
            ic = ws.cell(row=row, column=1, value=f.nombre)
            ic.border = _border
            ic.alignment = Alignment(wrap_text=True, vertical="top")
            _num(ws, row, 2, f.existencia, "General").border = _border
            _num(ws, row, 3, f.valor_costo_mxn).border = _border
            if usd:
                _num(ws, row, 4, f.valor_costo_usd, MONEY_USD).border = _border
            _num(ws, row, 4 + off, f.compradas_cant, "General").border = _border
            _num(ws, row, 5 + off, f.compradas_costo_mxn).border = _border
            _num(ws, row, 6 + off, f.salidas_cant, "General").border = _border
            _num(ws, row, 7 + off, f.vendido_mxn).border = _border
            _num(ws, row, 8 + off, f.utilidad_mxn).border = _border
            if venta_usd:
                # Dólares en SU columna, con "$" (MONEY_USD); None = celda
                # vacía (ese producto no vendió en dólares).
                _num(ws, row, 9 + off, f.vendido_usd, MONEY_USD).border = _border
                _num(ws, row, 10 + off, f.utilidad_usd, MONEY_USD).border = _border
            mc = ws.cell(row=row, column=col_matriculas, value=f.matriculas)
            mc.border = _border
            mc.alignment = Alignment(wrap_text=True, vertical="top")
            lineas = max(
                math.ceil(len(f.nombre or "") / ancho_item),
                math.ceil(len(f.matriculas or "") / 24),
                1,
            )
            if lineas > 1:
                ws.row_dimensions[row].height = 14 * lineas + 4
            row += 1
        # Fila TOTALES (piezas en stock, valorizado, compras, ventas,
        # utilidad — las columnas de cantidades del periodo van vacías).
        # Fórmulas (5-oct-2026): Σ de cada columna, con el redondeo del API
        # (piezas a 3 decimales, dinero a 2).
        ws.cell(row=row, column=1, value="TOTALES").font = Font(bold=True)

        def _suma_col(col: int, dec: int = 2) -> str:
            letra = get_column_letter(col)
            return f"ROUND(SUM({letra}5:{letra}{row - 1}),{dec})"

        _formula(ws, row, 2, _suma_col(2, 3), inv.total_piezas, "General", bold=True)
        _formula(ws, row, 3, _suma_col(3), inv.total_valor_mxn, MONEY, bold=True)
        if usd:
            # Total en dólares APARTE. Lo manda el API; si faltara (solo
            # llegaron las filas) se re-suma la columna para mostrarla — es
            # la misma moneda, nunca se mezcla con pesos.
            total_usd = _total_usd_de_columna(
                inv.total_valor_usd, [f.valor_costo_usd for f in inv.filas])
            _formula(ws, row, 4, _suma_col(4), total_usd, MONEY_USD, bold=True)
        _formula(ws, row, 5 + off, _suma_col(5 + off), inv.total_compras_mxn, MONEY,
                 bold=True)
        _formula(ws, row, 7 + off, _suma_col(7 + off), inv.total_vendido_mxn, MONEY,
                 bold=True)
        _formula(ws, row, 8 + off, _suma_col(8 + off), inv.total_utilidad_mxn, MONEY,
                 bold=True)
        if venta_usd:
            # Totales de la tienda en dólares, cada uno bajo SU columna.
            _formula(ws, row, 9 + off, _suma_col(9 + off), _total_usd_de_columna(
                inv.total_vendido_usd, [f.vendido_usd for f in inv.filas]),
                MONEY_USD, bold=True)
            _formula(ws, row, 10 + off, _suma_col(10 + off), _total_usd_de_columna(
                inv.total_utilidad_usd, [f.utilidad_usd for f in inv.filas]),
                MONEY_USD, bold=True)
        for c in range(1, n_cols + 1):
            cell = ws.cell(row=row, column=c)
            cell.fill = PatternFill("solid", fgColor=LIGHT)
            cell.border = _border
        row += 1
        # Nota bajo la tabla: por qué el total en pesos "no cuadra" con la
        # existencia de esos productos.
        if usd:
            n_sin_tc = _cuenta_sin_tc(inv)
            if n_sin_tc > 0:
                nc = ws.cell(row=row, column=1, value=_nota_sin_tc(n_sin_tc))
                nc.font = Font(color=RED, size=9, italic=True)
                ws.merge_cells(start_row=row, start_column=1,
                               end_row=row, end_column=n_cols)
                row += 1
        # Ventas cuya utilidad no cabe en ninguna moneda (venta en pesos
        # sobre costo en dólares sin T.C.): se dice, no se inventa un número.
        n_incompletas = _cuenta_utilidad_incompleta(inv)
        if n_incompletas > 0:
            nc = ws.cell(row=row, column=1,
                         value=_nota_utilidad_incompleta(n_incompletas))
            nc.font = Font(color=RED, size=9, italic=True)
            ws.merge_cells(start_row=row, start_column=1,
                           end_row=row, end_column=n_cols)
            row += 1
    else:
        ws.cell(
            row=row, column=1,
            value="Sin inventario con existencia ni movimientos en el periodo.",
        ).font = Font(color=MUTED, italic=True)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n_cols)
        row += 1

    # ===== Bloque 2: detalle de salidas (la vieja hoja 'refacciones') =====
    row += 2
    bc = ws.cell(row=row, column=1,
                 value="DETALLE DE SALIDAS DEL PERIODO (cargos a aviones)")
    bc.font = Font(bold=True, size=12, color=NAVY)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n_cols)
    row += 1
    _header_row(ws, row, ["FECHA", "ÍTEM", "AVIÓN", "COSTO VUELATOUR\nMXN",
                          "VENTA AL AVIÓN\nMXN", "GANANCIA\nMXN"])
    row += 1
    filas_det = refacciones.filas if refacciones is not None else []
    if filas_det:
        # Σ de columnas SOLO para mostrar (None si la columna vino vacía).
        tot_costo: float | None = None
        tot_venta: float | None = None
        tot_ganancia: float | None = None
        for f in filas_det:
            ws.cell(row=row, column=1, value=_fecha(f.fecha)).border = _border
            # El detalle del general llega con 'MATRÍCULA · ' al frente
            # (hoja de flota); aquí el avión tiene su columna — se limpia
            # para no duplicar.
            detalle = f.detalle or ""
            if f.matricula and detalle.startswith(f"{f.matricula} · "):
                detalle = detalle[len(f.matricula) + 3:]
            dc = ws.cell(row=row, column=2, value=detalle)
            dc.border = _border
            dc.alignment = Alignment(wrap_text=True, vertical="top")
            ac = ws.cell(row=row, column=3, value=f.matricula)
            ac.border = _border
            avion_fill = _hex(f.avion_color)
            if avion_fill:
                ac.fill = PatternFill("solid", fgColor=avion_fill)
            _num(ws, row, 4, f.costo_mxn).border = _border
            _num(ws, row, 5, f.venta_mxn).border = _border
            # GANANCIA = venta − costo (solo para MOSTRAR; 0 mientras la
            # salida se cargue a costo). Falta un lado → celda vacía.
            ganancia = (
                round(f.venta_mxn - f.costo_mxn, 2)
                if f.venta_mxn is not None and f.costo_mxn is not None
                else None
            )
            _formula(ws, row, 6, f"ROUND(E{row}-D{row},2)", ganancia).border = _border
            if f.costo_mxn is not None:
                tot_costo = round((tot_costo or 0.0) + f.costo_mxn, 2)
            if f.venta_mxn is not None:
                tot_venta = round((tot_venta or 0.0) + f.venta_mxn, 2)
            if ganancia is not None:
                tot_ganancia = round((tot_ganancia or 0.0) + ganancia, 2)
            lineas = max(1, math.ceil(len(detalle) / ancho_detalle))
            if lineas > 1:
                ws.row_dimensions[row].height = 14 * lineas + 4
            row += 1
        ws.cell(row=row, column=1, value="TOTALES").font = Font(bold=True)
        ini_det = row - len(filas_det)
        for col, total in ((4, tot_costo), (5, tot_venta), (6, tot_ganancia)):
            letra = get_column_letter(col)
            _formula(ws, row, col, f"ROUND(SUM({letra}{ini_det}:{letra}{row - 1}),2)",
                     total, MONEY, bold=True)
        for c in range(1, 7):
            cell = ws.cell(row=row, column=c)
            cell.fill = PatternFill("solid", fgColor=LIGHT)
            cell.border = _border
        row += 1
        if ultimo_precio:
            # Por qué la GANANCIA de aquí y la UTILIDAD del bloque 1 pueden
            # no coincidir al peso (T.C. del gasto vs T.C. del día de la
            # venta): dicho en la hoja, no descubierto por el cliente.
            _pinta_nota_envuelta(ws, row, _NOTA_BLOQUE2_TC, n_cols,
                                 sum(anchos[:n_cols]))
            row += 1
    else:
        ws.cell(
            row=row, column=1,
            value="Sin salidas de inventario cargadas a aviones en el periodo.",
        ).font = Font(color=MUTED, italic=True)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        row += 1

    row += 1
    if ultimo_precio:
        # Regla nueva: mismas piezas de la nota, sin «FIFO», citando el T.C.
        # de hoy con que el API valorizó; ENVUELTA (la de siempre se corta
        # al ancho de la hoja — se conserva así solo para el payload previo,
        # cuya hoja no debe moverse ni un estilo).
        nota = (
            _NOTA_INVENTARIO_ULTIMO_PRECIO.format(
                valor=_frase_valor_a_costo(inv.tc_hoy))
            + (_NOTA_INVENTARIO_USD_ULTIMO_PRECIO if usd else "")
            + (_NOTA_INVENTARIO_VENTA_USD_ULTIMO_PRECIO if venta_usd else "")
            + (_nota_margen(inv.margen_venta_pct, ultimo_precio=True)
               if inv.margen_venta_pct is not None else "")
        )
        _pinta_nota_envuelta(ws, row, nota, n_cols, sum(anchos[:n_cols]))
    else:
        nota = (
            _NOTA_INVENTARIO
            + (_NOTA_INVENTARIO_USD if usd else "")
            + (_NOTA_INVENTARIO_VENTA_USD if venta_usd else "")
            + (_nota_margen(inv.margen_venta_pct)
               if inv.margen_venta_pct is not None else "")
        )
        ws.cell(row=row, column=1, value=nota).font = Font(
            color=MUTED, size=9, italic=True
        )
        ws.merge_cells(start_row=row, start_column=1, end_row=row,
                       end_column=n_cols)

    for i, w in enumerate(anchos, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A5"


def render_balance_avion_xlsx(req: BalanceAvionRequest) -> bytes:
    """Libro INDIVIDUAL de un avión. Pestañas, en orden: reporte horas,
    cobranza, combustible, Gastos Indirectos, refacciones (solo si el API
    la manda), permisos, balance, pendientes de captura."""
    wb = Workbook()
    maestra = _hoja_maestra(wb.active, req)
    _hoja_cobranza(wb.create_sheet("cobranza"), req)
    _hoja_combustible(wb.create_sheet("combustible"), req.combustible, req,
                      maestra=maestra, usd=_USD_TC_LIBRO)
    # Pestaña ÚNICA "Gastos Indirectos" (pedido del cliente 2-sep-2026: las
    # hojas 'gastos indirectos' y 'otros gastos' se confundían): las DOS
    # listas del API pintadas juntas, ordenadas por fecha y con los
    # parciales del reparto manual resaltados. Fusión de PRESENTACIÓN — el
    # contrato (gastos_indirectos / otros_gastos) y la cascada no cambian.
    _hoja_gastos(
        wb.create_sheet("Gastos Indirectos"), "Gastos indirectos",
        _fusionar_hojas_gastos(req.gastos_indirectos, req.otros_gastos), req,
        nota=_NOTA_INDIRECTOS, resaltar_parciales=True,
        maestra=maestra, usd=_USD_TC_LIBRO,
    )
    # Hoja "refacciones" (29-ago): solo si el API la manda (skew tolerante).
    if req.refacciones is not None:
        _hoja_gastos(wb.create_sheet("refacciones"), "Refacciones",
                     req.refacciones, req, nota=_NOTA_REFACCIONES,
                     maestra=maestra, usd=_USD_TC_LIBRO)
    _hoja_gastos(wb.create_sheet("permisos"), "Permisos", req.permisos, req,
                 nota=_NOTA_PERMISOS, maestra=maestra, usd=_USD_TC_LIBRO)
    _hoja_balance(
        wb.create_sheet("balance"), req,
        refs=_refs_balance_individual(maestra, req.refacciones is not None),
    )
    _hoja_pendientes(wb.create_sheet("pendientes de captura"), req)

    # Fórmulas verificadas contra el API + su valor en caché (5-oct-2026).
    return xlsx_formulas.guardar(wb)


def _fila_resumen(ws: Worksheet, row: int, f: BalanceGeneralResumenFila,
                  *, bold: bool = False,
                  filas_avion: tuple[int, int] | None = None) -> None:
    """Fila del RESUMEN. Fórmulas (5-oct-2026): en la fila de un avión,
    GANANCIA = VENTA − COSTO TOTAL − COMBUSTIBLE (la leyenda impresa y la
    aritmética del API) − COMISIONES cuando el avión las trae (regla de
    septiembre 2026, API 0.0.65: sin ellas, la fórmula de siempre); en la
    fila TOTALES (`filas_avion` = primera y última fila de aviones) cada
    columna es la Σ de las de arriba, salvo VUELOS, que cuenta vuelos
    DISTINTOS (un multi-avión sale en dos filas) y queda como valor."""
    c = ws.cell(row=row, column=1, value=f.matricula)
    if bold:
        c.font = Font(bold=True)
    # La celda de la matrícula teñida con el color del avión = la LEYENDA del
    # libro (tabla "Color calendario" del equipo).
    swatch = _hex(f.color)
    if swatch:
        c.fill = PatternFill("solid", fgColor=swatch)
    negrita = {"bold": True} if bold else {}

    def _suma(col: int, dec: int = 2) -> str | None:
        if filas_avion is None:
            return None
        letra = get_column_letter(col)
        return f"ROUND(SUM({letra}{filas_avion[0]}:{letra}{filas_avion[1]}),{dec})"

    _num(ws, row, 2, f.vuelos, "0", **negrita)
    _formula(ws, row, 3, _suma(3), f.horas, HORAS, **negrita)
    _formula(ws, row, 4, _suma(4), f.horas_cobradas, HORAS, **negrita)
    _formula(ws, row, 5, _suma(5), f.venta_mxn, MONEY, **negrita)
    _formula(ws, row, 6, _suma(6), f.costo_mxn, MONEY, **negrita)
    _formula(ws, row, 7, _suma(7), f.combustible_mxn, MONEY, **negrita)
    _formula(ws, row, 8, _suma(8), f.comisiones_mxn, MONEY, **negrita)
    if filas_avion is not None:
        ganancia = _suma(9)
    elif _monto_no_cero(f.comisiones_mxn):
        ganancia = f"ROUND(E{row}-F{row}-G{row}-H{row},2)"
    else:
        ganancia = f"ROUND(E{row}-F{row}-G{row},2)"
    _formula(ws, row, 9, ganancia, f.ganancia_mxn, MONEY, **negrita)
    _formula(ws, row, 10, _suma(10), f.cobrado_mxn, MONEY, **negrita)
    _formula(ws, row, 11, _suma(11), f.por_cobrar_mxn, MONEY, **negrita)
    _formula(ws, row, 12, _suma(12, 0), f.pendientes, "0", **negrita)


def _hoja_resumen_general(
    ws: Worksheet, req: BalanceGeneralRequest, regla_cubre_avion: str | None = None
) -> None:
    """Índice de la flota: una fila por avión (los números vienen del API —
    son los TOTALES del libro de cada avión, jamás se recalculan aquí).
    `regla_cubre_avion` (7-oct-2026): la etiqueta de la regla si el payload es
    del API 0.0.66 y el periodo trae vuelos de ella — la nota de COMISIONES
    dice entonces que el pago al vendedor lo cubre el avión."""
    ws.title = "RESUMEN flota"
    _title(
        ws,
        f"Balance general VuelaTour · {req.periodo_desde or ''} a {req.periodo_hasta or ''}",
        1,
        12,
    )
    _header_row(
        ws,
        3,
        [
            "AVIÓN",
            "VUELOS",
            "HORAS\nVOLADAS",
            "HORAS\nCOBRADAS",
            "VENTA\nMXN",
            "COSTO TOTAL\nMXN",
            "COMBUSTIBLE\nMXN",
            "COMISIONES\nMXN",
            "GANANCIA\nMXN",
            "COBRADO\nMXN",
            "POR COBRAR\nMXN",
            "PENDIENTES",
        ],
    )
    row = 4
    for f in req.resumen:
        _fila_resumen(ws, row, f)
        row += 1
    if req.resumen_totales is not None:
        _fila_resumen(
            ws, row, req.resumen_totales, bold=True,
            filas_avion=(4, row - 1) if req.resumen else None,
        )
        row += 1
    # GANANCIA MXN negativa en rojo (aviones + TOTALES), 8-oct-2026.
    if row > 4:
        _rojo_si_negativo(ws, f"I4:I{row - 1}")
    row += 1
    # Cómo se nombra el costo de las salidas en el índice de hojas: con el
    # API 0.0.36 (`regla_costo`) es el último precio de compra; con uno
    # previo, el costo FIFO de siempre (texto intacto).
    costo_salidas = (
        "costo al último precio de compra"
        if _es_ultimo_precio(req.inventario) else "costo FIFO"
    )
    for nota in (
        "VENTA = tiempo de vuelo + ajuste + IVA proporcional (sin TUAs/extras/"
        "pernocta/comisión del vendedor: ver 'otros movimientos'). GANANCIA = "
        "VENTA − COSTO TOTAL − COMBUSTIBLE − COMISIONES; los TUAs pagados no "
        "restan a ningún avión. La GANANCIA va antes de GASTOS INDIRECTOS "
        "(incluida la parte repartida a mano), refacciones y permisos — la "
        "utilidad final por avión está en la hoja 'balance'.",
        _con_regla(_RESUMEN_COMISIONES, _regla_del_general(req))
        + (
            _RESUMEN_PAGO_VENDEDOR_REAL
            if req.consolidado is not None
            and _hay_pago_vendedor_real(req.consolidado.otros_movimientos)
            else _RESUMEN_PAGO_VENDEDOR_PROVISION
        )
        + (_RESUMEN_REGLA_CUBRE_AVION.format(regla=regla_cubre_avion) if regla_cubre_avion else ""),
        "Vuelos multi-avión (regla 28-ago-2026): la VENTA se reparte entre "
        "las matrículas del vuelo en partes iguales por tramo vendido — los "
        "ferries/tramos operativos no reparten — (cada fila COMPARTIDO "
        "trae su parte proporcional de venta, cobros y por cobrar) y los "
        "GASTOS van al avión del tramo que los generó — así el cruce de "
        "columnas cuadra por avión.",
        "El detalle está en las hojas siguientes: los datos de TODOS los "
        "aviones juntos — cada fila se identifica por su CLAVE y el color de "
        "su avión. El detalle de COMBUSTIBLE, GASTOS INDIRECTOS (incluida la "
        "parte repartida a mano) y PERMISOS por avión vive en el libro "
        "INDIVIDUAL de cada avión (Reportes › Balance por avión); aquí "
        "restan igual en la hoja 'balance' (indirectos y reparto manual en "
        "una sola fila). Hojas propias de este libro: 'inventario' "
        "(tiendita: resumen por ítem del periodo + detalle de salidas con "
        f"{costo_salidas} vs venta al avión), 'otros gastos' (gastos de la EMPRESA "
        "sin avión ni vuelo, fuera de toda cascada por avión) y 'repartidos "
        "a aviones' (la parte de esos gastos asignada a mano a cada avión; "
        "en el libro individual va dentro de su hoja 'Gastos Indirectos').",
        # Bloque «VUELATOUR (empresa)» (6-oct-2026): solo con `empresa`; sin
        # él el RESUMEN es el de siempre.
        *((_RESUMEN_BLOQUE_EMPRESA,) if req.empresa is not None else ()),
    ):
        ws.cell(row=row, column=1, value=nota).font = Font(
            color=MUTED, size=9, italic=True
        )
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=12)
        row += 1
    for i in range(1, 13):
        ws.column_dimensions[get_column_letter(i)].width = 10 if i == 1 else 14
    ws.freeze_panes = "A4"


# Estados de vuelo que NO se anotan junto a la clave en 'otros movimientos'
# (los normales); cualquier otro (RESERVA, CANCELADO…) se marca.
_ESTADOS_NORMALES = ("CONFIRMADO", "EN_VUELO", "COMPLETADO")

# PROVISIÓN en amarillo (6-oct-2026, pedido del cliente: «los que están
# provisión podrían estar en amarillo por fi»). El API escribe la marca EN
# MAYÚSCULAS como calificador del concepto del egreso, «· PROVISIÓN (»
# («pago comisión vendedor (X) · PROVISIÓN (mismo monto que lo cobrado…)»,
# aircraft-balance.service.ts) y, cuando la fila junta varios egresos del
# vuelo, en la línea de la nota de esa celda («<concepto> = $x»). Solo esa
# marca cuenta: el GASTO REAL dice «reemplaza la provisión» (minúsculas) y
# una fila suelta lleva «<categoría> · <proveedor>» — un proveedor como
# «PROVISIONES AEREAS DEL SURESTE» NO es provisión (revisión 6-oct-2026: la
# búsqueda de la palabra suelta lo pintaba de amarillo).
_MARCA_PROVISION = re.compile(r"·\s*PROVISI[OÓ]N\s*\(")
_LEYENDA_PROVISION = "Amarillo = provisión (sin gasto real capturado)"
# Columnas que se pintan: concepto y monto del egreso.
_COLUMNAS_PROVISION = (3, 4)


def _es_provision(f: BalanceOtroMovimientoFila) -> bool:
    """¿El egreso de la fila es (o incluye) una PROVISIÓN sin gasto real?
    (la marca del API en el concepto o en una línea de su nota)."""
    return any(
        _MARCA_PROVISION.search(texto) for texto in (f.concepto_egreso, f.nota_egreso) if texto
    )


def _hoja_otros_movimientos(
    ws: Worksheet,
    hoja: BalanceHojaOtrosMovimientos,
    regla_vigente: str | None = None,
    *,
    cubre_avion: bool = False,
) -> None:
    """Pestaña "Otros movimientos" (28-ago, réplica de la hoja manual
    "dinero otros ingresos"): conceptos cobrados al cliente vs pagados, por
    clave de vuelo (celda teñida con el color del avión), más la sección de
    movimientos SIN avión/SIN vuelo. Remanente negativo en ROJO. Incluye
    TODOS los estados del periodo (igual que la hoja maestra): la clave
    lleva ' · ESTADO' cuando no es un estado normal y va en ROJO si es
    CANCELADO. Los montos vienen YA calculados del API — aquí solo se
    pinta. Desde el 6-oct-2026 el egreso con PROVISIÓN (pago al vendedor sin
    gasto real) va en amarillo — concepto y monto — con la leyenda
    `_LEYENDA_PROVISION` al pie (solo si alguna fila la trae). Con vuelos de
    la regla de comisiones a cargo del avión (`regla_vigente`, 6-oct-2026)
    la leyenda de la fila 2 explica la línea «a cargo del avión» y la parte
    de VuelaTour de la comisión bancaria (`_OM_REGLA_COMISIONES`). Con el API
    0.0.66 (`cubre_avion`, 7-oct-2026) esa línea ya no viaja y la leyenda
    explica el pago «cubierto por el avión» (`_OM_REGLA_COMISIONES_CUBRE_AVION`)."""
    ws.cell(row=1, column=1, value="OTROS MOVIMIENTOS VUELATOUR").font = Font(
        bold=True, size=12, color=NAVY
    )
    ws.cell(
        row=2, column=1,
        value="Ingreso de VuelaTour (no del avión): TUAs, extras, viáticos de "
        "pernocta y comisión del vendedor cobrados al cliente (con su IVA) vs "
        "lo pagado — "
        + (
            _OM_PAGO_VENDEDOR_REAL
            if _hay_pago_vendedor_real(hoja)
            else _OM_PAGO_VENDEDOR_PROVISION
        )
        + (
            (_OM_REGLA_COMISIONES_CUBRE_AVION if cubre_avion else _OM_REGLA_COMISIONES).format(
                regla=regla_vigente
            )
            if regla_vigente
            else ""
        )
        + " UNA FILA POR "
        "VUELO: los ingresos van sumados en una celda y los egresos en otra; "
        "el desglose concepto por concepto está en el COMENTARIO de cada "
        "celda (pasa el cursor sobre el triángulo rojo). Una fila puede traer "
        "solo ingreso o solo egreso. Todo en MXN. Incluye todos "
        "los estados del periodo (igual que la hoja maestra); los cancelados "
        "se marcan (clave · CANCELADO en rojo) y los demás estados no "
        "normales llevan su estado junto a la clave. Los MOVIMIENTOS SIN "
        "AVIÓN / SIN VUELO de abajo son 'gas sin avión', 'tuas sin vuelo' y "
        "los OTROS INGRESOS registrados en Ingresos (clave ING-n: intereses, "
        "reembolsos, ventas de activos…); los anticipos y las aportaciones NO "
        "aparecen: no son resultado (29-ago: los gastos de empresa viven en "
        "la hoja 'otros gastos' — antes llamada 'gastos VuelaTour').",
    ).font = Font(italic=True, size=9, color=MUTED)
    headers = [
        "clave", "fecha\nvuelo", "concepto", "egreso", "fecha", "concepto",
        "ingreso", "fecha", "remanente", "factura\nvuelatour",
    ]
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=3, column=i, value=h)
        c.font = Font(bold=True, color="FFFFFF", size=9)
        c.fill = PatternFill("solid", fgColor=BRAND)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = _border
        ws.column_dimensions[get_column_letter(i)].width = 15
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["C"].width = 28
    ws.column_dimensions["F"].width = 32
    row = 4
    tot_e = tot_i = 0.0
    hay_provision = False

    def pinta(f: BalanceOtroMovimientoFila) -> None:
        nonlocal row, tot_e, tot_i, hay_provision
        # Estado del vuelo (API nuevo): se anota junto a la clave cuando NO
        # es normal; CANCELADO además en rojo (igual que la hoja maestra).
        estado = (f.estado or "").strip().upper()
        clave = f.clave or ""
        if estado and estado not in _ESTADOS_NORMALES:
            clave = f"{clave} · {estado}" if clave else estado
        vals = [
            clave, _fecha(f.fecha_vuelo), f.concepto_egreso, f.egreso_mxn,
            _fecha(f.fecha_egreso), f.concepto_ingreso, f.ingreso_mxn,
            _fecha(f.fecha_ingreso), f.remanente_mxn, f.factura,
        ]
        for i, v in enumerate(vals, start=1):
            cell = ws.cell(row=row, column=i, value=v if v != "" else None)
            cell.border = _border
            if isinstance(v, (int, float)):
                cell.number_format = MONEY
                if i == 9:
                    # REMANENTE = round2(ingreso − egreso), igual que el API
                    # (5-oct-2026: fórmula visible).
                    xlsx_formulas.escribir(cell, f"ROUND(G{row}-D{row},2)", v)
            if i == 10:
                # FACTURA: la MISMA etiqueta que la columna FACTURA
                # VUELATOUR de la hoja maestra (30-sep-2026) — un folio
                # tecleado «=A-12» se queda como TEXTO aquí también; si no,
                # el Balance general entero abre con error por esta pestaña.
                _texto_literal(cell)
        swatch = _hex(f.avion_color)
        if swatch:
            ws.cell(row=row, column=1).fill = PatternFill("solid", fgColor=swatch)
        if estado == "CANCELADO":
            ws.cell(row=row, column=1).font = Font(color=RED)
        if isinstance(f.remanente_mxn, (int, float)) and f.remanente_mxn < 0:
            ws.cell(row=row, column=9).font = Font(color=RED)
        # PROVISIÓN (pago al vendedor sin gasto real): concepto y monto del
        # egreso en amarillo; la leyenda va al pie.
        if _es_provision(f):
            hay_provision = True
            for col in _COLUMNAS_PROVISION:
                ws.cell(row=row, column=col).fill = PatternFill("solid", fgColor=AMBER)
        # Una fila por vuelo: el desglose de cada celda va en su COMENTARIO
        # (el equipo lo pidió así para ahorrar espacio y entenderlo mejor).
        for col, nota in ((7, f.nota_ingreso), (4, f.nota_egreso)):
            if nota:
                lineas = nota.count("\n") + 1
                com = Comment(nota, "VuelaTour")
                com.width = 420
                com.height = max(60, 18 * lineas + 20)
                ws.cell(row=row, column=col).comment = com
        if isinstance(f.egreso_mxn, (int, float)):
            tot_e += f.egreso_mxn
        if isinstance(f.ingreso_mxn, (int, float)):
            tot_i += f.ingreso_mxn
        row += 1

    for f in hoja.filas:
        pinta(f)
    if not hoja.filas and not hoja.filas_sueltas:
        ws.cell(
            row=row, column=1, value="Sin otros movimientos en el periodo."
        ).font = Font(color=MUTED, italic=True)
        row += 1
    if hoja.filas_sueltas:
        row += 1
        t = ws.cell(row=row, column=1, value="MOVIMIENTOS SIN AVIÓN / SIN VUELO")
        t.font = Font(bold=True, size=10, color=NAVY)
        t.fill = PatternFill("solid", fgColor=LIGHT)
        for i in range(2, 11):
            ws.cell(row=row, column=i).fill = PatternFill("solid", fgColor=LIGHT)
        row += 1
        for f in hoja.filas_sueltas:
            pinta(f)
    ws.cell(row=row, column=1, value="TOTALES").font = Font(bold=True)
    rem = tot_i - tot_e
    # Fórmulas (5-oct-2026): Σ egresos, Σ ingresos y remanente = ingresos −
    # egresos (los encabezados de sección no traen montos: SUM los ignora).
    formulas = {
        4: f"ROUND(SUM(D4:D{row - 1}),2)",
        7: f"ROUND(SUM(G4:G{row - 1}),2)",
        9: f"ROUND(G{row}-D{row},2)",
    }
    for col, val in ((4, round(tot_e, 2)), (7, round(tot_i, 2)), (9, round(rem, 2))):
        cell = ws.cell(row=row, column=col, value=val)
        cell.font = Font(bold=True, color=RED if col == 9 and rem < 0 else "000000")
        cell.number_format = MONEY
        cell.fill = PatternFill("solid", fgColor=LIGHT)
        if row > 4:
            xlsx_formulas.escribir(cell, formulas[col], val)
    # REMANENTE negativo en rojo (filas + TOTALES), 8-oct-2026.
    _rojo_si_negativo(ws, f"I4:I{row}")
    # Leyenda del amarillo (solo si alguna fila lo trae: sin provisiones la
    # hoja es la de siempre).
    if hay_provision:
        fila_ley = row + 2
        ley = ws.cell(row=fila_ley, column=1, value=_LEYENDA_PROVISION)
        ley.font = Font(italic=True, size=9, color=NAVY)
        ley.fill = PatternFill("solid", fgColor=AMBER)
        ws.merge_cells(start_row=fila_ley, start_column=1, end_row=fila_ley, end_column=3)
    ws.freeze_panes = "A4"


def render_balance_general_xlsx(req: BalanceGeneralRequest) -> bytes:
    """Balance general VuelaTour (regla del cliente, 18-ago; hojas 29-ago;
    apartado propio en Reportes desde el 2-sep-2026 — antes era la opción
    "Toda la flota" del balance por avión): RESUMEN + los datos de TODOS
    los aviones JUNTOS — 1 reporte de horas, 1 otros movimientos (TUAs/
    extras/pernocta/comisión del vendedor cobrados y pagados: dinero de
    VuelaTour), 1 cobranza, 1 otros gastos (gastos de EMPRESA sin vuelo ni
    avión, payload `gastos_empresa`), 1 repartidos a aviones (parte de esos
    gastos repartida a mano a cada avión, payload `otros_gastos`; sin
    TUAs), 1 inventario (tiendita, 30-ago: resumen por ítem del periodo +
    detalle de salidas con su costo vs venta al avión — sustituye a la
    hoja 'refacciones' del general, que solo se pinta como fallback de un
    API viejo), 1 balance (bloques por avión: los socios son por avión;
    indirectos + reparto manual en UNA fila, 2-sep; y, hasta el final, el
    bloque «VUELATOUR (empresa)» con el payload `empresa`, 6-oct-2026) y 1
    pendientes.
    Renombre 1-sep-2026 (modelo mental del equipo: lo sin avión ni vuelo
    "cae en otros gastos"): la hoja de empresa se llamaba 'gastos
    VuelaTour' y la de parciales 'otros gastos'. Este libro CONSERVA sus
    dos hojas ('otros gastos' de empresa y 'repartidos a aviones'): la
    fusión del 2-sep-2026 es solo del libro INDIVIDUAL (hoja "Gastos
    Indirectos"); los campos del contrato del API no cambian. Desde el
    29-ago el general ya NO pinta 'combustible', 'Gastos Indirectos' ni
    'permisos' (viven en el libro INDIVIDUAL de cada avión); el API los
    sigue calculando y restan igual en la cascada de la hoja 'balance'.
    Cada fila se identifica por su clave y el COLOR del avión
    (aeronave.color_calendario, editable en el apartado del avión).
    `req.variante` (6-oct-2026, API 0.0.64): «general» cambia la hoja
    maestra (costo total + costo por hora, `_COLS_GENERAL`), las letras con
    que la citan las demás hojas y la nota de 'repartidos a aviones' (que
    nombraba OPERACIONES); «mensual» o ausente = la hoja de vuelos de
    siempre.
    7-oct-2026 (API 0.0.66; pedido: «estas comisiones se están duplicando en
    la general»): la hoja de vuelos de las dos variantes va SIN la columna
    COMISIONES y antes de comisiones (`_COLS_FLOTA` / `_COLS_GENERAL`,
    `_antes_de_comisiones`); la cascada de cada avión de 'balance' las
    muestra en dos líneas cuando el API manda `balance.comisiones_usd`, y las
    leyendas de 'otros movimientos', RESUMEN y bloque VUELATOUR (empresa)
    explican el pago «cubierto por el avión» de ese API."""
    wb = Workbook()
    cons = req.consolidado
    # Nombre de la regla de comisiones (6-oct-2026, ADITIVO): si el API lo
    # manda solo en el request, las hojas de la flota lo usan igual.
    if cons is not None and cons.regla_comisiones is None and req.regla_comisiones:
        cons = cons.model_copy(update={"regla_comisiones": req.regla_comisiones})
    # ¿El periodo trae vuelos de la regla? ⇒ aclaraciones en 'otros
    # movimientos', en el RESUMEN y en el bloque VUELATOUR (empresa); con el
    # API 0.0.66 (`_avion_cubre_al_vendedor`), las del pago al vendedor
    # «cubierto por el avión».
    regla_vigente = _regla_vigente(cons) if cons is not None else None
    cubre_avion = regla_vigente is not None and _avion_cubre_al_vendedor(req)
    _hoja_resumen_general(wb.active, req, regla_vigente if cubre_avion else None)
    if cons is not None:
        # La hoja maestra se titula sola ("reporte horas FLOTA"). Variante
        # «general» (6-oct-2026, API 0.0.64): costo total + costo por hora en
        # lugar del desglose de costos; «mensual» (o sin la llave): la de
        # siempre. Las demás hojas son las mismas en las dos.
        maestra = _hoja_maestra(
            wb.create_sheet(),
            cons,
            general=True,
            costo_por_hora=req.variante == VARIANTE_BALANCE_GENERAL,
        )
        # Pestaña "Otros movimientos" JUNTO a "reporte horas" (pedido del
        # cliente 28-ago): se crea aquí para que quede al lado; solo si el
        # API la manda.
        if cons.otros_movimientos is not None:
            _hoja_otros_movimientos(
                wb.create_sheet("otros movimientos"),
                cons.otros_movimientos,
                regla_vigente,
                cubre_avion=cubre_avion,
            )
        _hoja_cobranza(wb.create_sheet("cobranza"), cons, general=True)
        # Hoja 'combustible' de la FLOTA (10-sep-2026, pedido del cliente:
        # «ya no me sale el apartado de combustibles»): se había quitado del
        # general el 29-ago junto con 'Gastos Indirectos' y 'permisos'. Vuelve
        # SOLO combustible: mismas filas por avión que su libro (secciones
        # por matrícula con subtotal, litros y $/L) + total de flota. El API
        # ya la mandaba (`consolidado.combustible`); sigue restando UNA vez en
        # los bloques de 'balance'. 'Gastos Indirectos' y 'permisos' siguen
        # solo en el libro individual.
        _hoja_combustible(
            wb.create_sheet("combustible"), cons.combustible, cons,
            titulo="Combustible — flota (por avión y por mes)",
            maestra=maestra,
        )
        # Hoja 'otros gastos' (payload `gastos_empresa`, 29-ago): gastos de
        # EMPRESA sin avión ni vuelo — egresos de VuelaTour, fuera de toda
        # cascada por avión (1-sep-2026: antes 'gastos VuelaTour'; solo
        # cambió el nombre de la hoja, el campo del contrato sigue igual).
        # Va en la posición que el equipo conoce como 'otros gastos', JUNTO
        # a 'repartidos a aviones'.
        if req.gastos_empresa is not None:
            # TOTAL USD = TOTAL MXN ÷ la celda TC PROMEDIO: el API convierte
            # estos gastos con el TC promedio de FLOTA ya redondeado.
            _hoja_gastos(wb.create_sheet("otros gastos"),
                         "Otros gastos — VuelaTour (sin avión ni vuelo)",
                         req.gastos_empresa, cons,
                         nota=_NOTA_GASTOS_EMPRESA, titulo_exacto=True,
                         maestra=maestra, usd=_USD_TC_CELDA)
        # Hoja 'repartidos a aviones' (payload `otros_gastos`; 1-sep-2026:
        # antes 'otros gastos' — se veía vacía sin repartos y ese nombre
        # ahora es de la hoja de gastos de empresa). Se CONSERVA en el
        # general (2-sep): la fusión en "Gastos Indirectos" es solo del
        # libro individual. Sin `resaltar_parciales`: aquí la celda DETALLE
        # lleva el color del avión. Su nota dice dónde queda el TUA pagado
        # con las columnas de la hoja de vuelos de cada variante.
        _hoja_gastos(
            wb.create_sheet("repartidos a aviones"),
            "Otros gastos repartidos a aviones",
            cons.otros_gastos, cons,
            nota=(
                _CPH_NOTA_REPARTIDOS_A_AVIONES
                if maestra.disp.costo_por_hora
                else _NOTA_REPARTIDOS_A_AVIONES
            ),
            maestra=maestra,
        )
        # Hoja "inventario" (tiendita, 30-ago): resumen POR ÍTEM del periodo
        # + el detalle de salidas que antes pintaba la hoja "refacciones" del
        # general (el libro INDIVIDUAL conserva la suya; la cascada de
        # 'balance' no cambia — la hoja era informativa). Fallback: un API
        # viejo sin `inventario` sigue pintando "refacciones" como antes
        # (skew tolerante).
        if req.inventario is not None:
            _hoja_inventario(wb.create_sheet("inventario"), req.inventario,
                             cons.refacciones, cons)
        elif cons.refacciones is not None:
            _hoja_gastos(wb.create_sheet("refacciones"), "Refacciones",
                         cons.refacciones, cons,
                         nota=_NOTA_REFACCIONES_GENERAL, maestra=maestra)
        # Hoja balance: un BLOQUE por avión (título teñido con su color).
        ws_b = wb.create_sheet("balance")
        _title(ws_b, "Balance por avión — Balance general VuelaTour", 1, 3)
        ws_b.cell(
            row=2, column=1,
            value=f"Periodo: {req.periodo_desde or '—'} a "
            f"{req.periodo_hasta or '—'} (todo en USD) · los socios son POR avión",
        ).font = Font(italic=True, size=10, color=MUTED)
        row = 4
        # Celdas MONTO USD de los socios `es_empresa` de cada bloque: las cita
        # el bloque «VUELATOUR (empresa)» del final (6-oct-2026).
        celdas_empresa: list[_CeldaSocioEmpresa] = []
        for avion in req.aviones:
            tc = ws_b.cell(row=row, column=1, value=avion.matricula or "—")
            tc.font = Font(bold=True, size=12, color=NAVY)
            # Solo la celda de la matrícula lleva el color (como su libro).
            swatch = _hex(avion.avion_color)
            if swatch:
                tc.fill = PatternFill("solid", fgColor=swatch)
            # Con el API 0.0.66 la cascada abre antes de comisiones y las
            # resta en una línea (`_cascada_con_comisiones`, 7-oct-2026).
            row = (
                _bloque_balance(
                    ws_b,
                    avion,
                    row + 1,
                    general=True,
                    celdas_empresa=celdas_empresa,
                    regla=_regla_del_general(req),
                )
                + 2
            )
        # Balance de la EMPRESA hasta el final (6-oct-2026): solo con el API
        # 0.0.59+; sin `empresa` el libro es byte-idéntico al de antes.
        if req.empresa is not None:
            _bloque_empresa(
                ws_b, req.empresa, row, celdas_empresa, regla_vigente, cubre_avion=cubre_avion
            )
        for i, w in enumerate([46, 14, 16], start=1):
            ws_b.column_dimensions[get_column_letter(i)].width = w

        _hoja_pendientes(wb.create_sheet("pendientes de captura"), cons)

    # Fórmulas verificadas contra el API + su valor en caché (5-oct-2026).
    return xlsx_formulas.guardar(wb)
