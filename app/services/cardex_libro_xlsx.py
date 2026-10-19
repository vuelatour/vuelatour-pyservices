"""Cardex de un ítem en formato LIBRO (réplica del cuaderno del cliente).

Un solo sheet con dos bloques lado a lado — ENTRADAS | SALIDAS — cada uno
con sus columnas (venta, remanente y ganancia del lado de salidas). Todo
viene YA calculado del API: aquí solo se pinta el libro con los estilos
BRAND de la casa (tabla_xlsx).

25-sep-2026 (API 0.0.36): costo = último precio de compra y montos en pesos
con el T.C. oficial de su día. Lo único que cambia aquí es el SUBTÍTULO:
la `nota` del API va tras «Montos en MXN» para que quien lea el libro sepa
con qué T.C. se convirtió cada lado. Sin `nota` (API previo) el libro sale
idéntico.

18-oct-2026 (API 0.0.92; pedido del cliente: el cardex «en forma de lista
por orden de fecha … con sus entradas, salidas y movimientos» y, en el
Excel, «stock inicio del mes y stock final del mes y cantidad de unidades
vendidas»): con `movimientos` en el payload el libro trae TRES hojas —la
LISTA por fecha con una banda por mes, el RESUMEN POR MES y, al final, la
hoja ENTRADAS | SALIDAS de siempre—. Sin `movimientos` (API previo), el
libro de siempre byte a byte. Aquí tampoco se calcula nada: filas, stock y
meses llegan del API (`cardexCronologicoDe` / `mesesCardexDe`); solo se
re-suman columnas para la fila TOTAL del resumen.
"""

from __future__ import annotations

from datetime import date
from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from app.schemas.cardex_libro import (
    CardexLibroMes,
    CardexLibroMovimiento,
    CardexLibroRequest,
)
from app.services.tabla_xlsx import BRAND, LIGHT, MONEY, WHITE, sheet_title

_thin = Side(style="thin", color="D5DBE3")
_border = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
# Separación visual entre el bloque ENTRADAS y el de SALIDAS.
_sep = Side(style="medium", color=BRAND)

# Columnas del bloque ENTRADAS (1-6) y del bloque SALIDAS (7-14).
_COLS_ENTRADAS = [
    ("Fecha entrada", 12),
    ("Cantidad", 10),
    ("Descripción", 34),
    ("Valor compra", 13),
    ("Valor compra total", 14),
    ("Cantidad en stock", 11),
]
_COLS_SALIDAS = [
    ("Fecha de salida", 12),
    ("Cantidad", 10),
    ("Descripción del producto", 34),
    ("Valor unitario al que se vendió", 14),
    ("Valor total", 13),
    ("Remanente", 11),
    ("Ganancia", 13),
    ("A quién se le vendió", 16),
]
_N_ENT = len(_COLS_ENTRADAS)  # 6
_NCOLS = _N_ENT + len(_COLS_SALIDAS)  # 14
# Columnas con formato de dinero (1-based).
_MONEY_COLS = {4, 5, _N_ENT + 4, _N_ENT + 5, _N_ENT + 7}


def _celda(ws, row: int, col: int, value, *, money: bool = False, wrap: bool = False):
    cell = ws.cell(row=row, column=col, value=value)
    cell.border = _border
    if money and value is not None:
        cell.number_format = MONEY
    if wrap:
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    # Borde grueso entre bloques (última col de ENTRADAS | primera de SALIDAS).
    if col == _N_ENT:
        cell.border = Border(left=_thin, right=_sep, top=_thin, bottom=_thin)
    return cell


def _subtitulo(req: CardexLibroRequest) -> str:
    partes = []
    if req.numero_parte:
        partes.append(f"Parte {req.numero_parte}")
    if req.unidad:
        partes.append(f"Unidad: {req.unidad}")
    partes.append(f"Montos en {req.moneda or 'MXN'}")
    # Nota del API (ADITIVA): con qué T.C. se convirtieron compras y ventas y
    # de dónde sale el costo. Tal cual (solo sin el punto final: en el
    # subtítulo los tramos van separados por « · »); aquí no se redacta
    # negocio.
    nota = (req.nota or "").strip().rstrip(".").rstrip()
    if nota:
        partes.append(nota)
    if req.generado:
        partes.append(f"Generado {req.generado}")
    return " · ".join(partes)


def render_cardex_libro_xlsx(req: CardexLibroRequest) -> bytes:
    wb = Workbook()
    ws = wb.active
    if req.movimientos is None:
        # API previo: el libro de siempre (una hoja ENTRADAS | SALIDAS).
        ws.title = sheet_title("Cardex libro")
        _hoja_entradas_salidas(ws, req)
    else:
        ws.title = sheet_title("Cardex")
        _hoja_lista(ws, req)
        _hoja_meses(wb.create_sheet(sheet_title("Resumen por mes")), req)
        _hoja_entradas_salidas(
            wb.create_sheet(sheet_title("Entradas y salidas")), req
        )

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _hoja_entradas_salidas(ws, req: CardexLibroRequest) -> None:
    """La hoja de siempre: ENTRADAS | SALIDAS lado a lado (cuaderno)."""
    # Título + subtítulo.
    t = ws.cell(row=1, column=1, value=req.titulo or f"Cardex — {req.item_nombre}")
    t.font = Font(bold=True, size=14, color=BRAND)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=_NCOLS)
    s = ws.cell(row=2, column=1, value=_subtitulo(req))
    s.font = Font(italic=True, size=10, color="5B6470")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=_NCOLS)

    # Encabezado DOBLE: fila de bloques (ENTRADAS | SALIDAS) + fila de columnas.
    hrow = 4
    for etiqueta, c1, c2 in (
        ("ENTRADAS", 1, _N_ENT),
        ("SALIDAS", _N_ENT + 1, _NCOLS),
    ):
        cell = ws.cell(row=hrow, column=c1, value=etiqueta)
        cell.font = Font(bold=True, color=WHITE, size=12)
        cell.fill = PatternFill("solid", fgColor=BRAND)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        for col in range(c1, c2 + 1):
            x = ws.cell(row=hrow, column=col)
            x.fill = PatternFill("solid", fgColor=BRAND)
            x.border = _border
        ws.merge_cells(start_row=hrow, start_column=c1, end_row=hrow, end_column=c2)

    crow = hrow + 1
    for col, (label, width) in enumerate(_COLS_ENTRADAS + _COLS_SALIDAS, start=1):
        cell = ws.cell(row=crow, column=col, value=label)
        cell.font = Font(bold=True, color=WHITE)
        cell.fill = PatternFill("solid", fgColor=BRAND)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = _border
        ws.column_dimensions[get_column_letter(col)].width = width

    # Filas: cada lado avanza por su cuenta (como en el cuaderno físico).
    r = crow + 1
    nfilas = max(len(req.entradas), len(req.salidas))
    for i in range(nfilas):
        e = req.entradas[i] if i < len(req.entradas) else None
        v = req.salidas[i] if i < len(req.salidas) else None
        _celda(ws, r, 1, e.fecha if e else None)
        _celda(ws, r, 2, e.cantidad if e else None)
        _celda(ws, r, 3, e.descripcion if e else None, wrap=True)
        _celda(ws, r, 4, e.valor_compra_unitario if e else None, money=True)
        _celda(ws, r, 5, e.valor_compra_total if e else None, money=True)
        _celda(ws, r, 6, e.stock_despues if e else None)
        _celda(ws, r, 7, v.fecha if v else None)
        _celda(ws, r, 8, v.cantidad if v else None)
        _celda(ws, r, 9, v.descripcion if v else None, wrap=True)
        _celda(ws, r, 10, v.venta_unitaria if v else None, money=True)
        _celda(ws, r, 11, v.venta_total if v else None, money=True)
        _celda(ws, r, 12, v.remanente if v else None)
        _celda(ws, r, 13, v.ganancia if v else None, money=True)
        _celda(ws, r, 14, v.vendido_a if v else None)
        r += 1

    # Totales (los manda el API): compra | venta | ganancia.
    tot = {
        3: "TOTAL COMPRA",
        5: req.total_compra,
        9: "TOTAL VENTA / GANANCIA",
        11: req.total_venta,
        13: req.total_ganancia,
    }
    for col in range(1, _NCOLS + 1):
        cell = ws.cell(row=r, column=col, value=tot.get(col))
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor=LIGHT)
        cell.border = _border
        if col in _MONEY_COLS and tot.get(col) is not None:
            cell.number_format = MONEY

    ws.freeze_panes = ws.cell(row=crow + 1, column=1)


# ---------------------------------------------------------------------------
# Formato nuevo (18-oct-2026, API 0.0.92): LISTA por fecha + RESUMEN POR MES.
# ---------------------------------------------------------------------------

_VERDE = "1F7A3F"
_ROJO = "B42318"
_GRIS = "5B6470"

# Lista por fecha: (encabezado, ancho).
_COLS_LISTA = [
    ("Fecha", 12),
    ("Movimiento", 15),
    ("Detalle", 42),
    ("Entra", 10),
    ("Sale", 10),
    ("Existencia\n(después)", 12),
    ("Precio unitario", 14),
    ("Total", 14),
    ("Ganancia", 13),
]
_N_LISTA = len(_COLS_LISTA)  # 9

# Resumen por mes: (encabezado, ancho).
_COLS_MESES = [
    ("Mes", 24),
    ("Stock al inicio\ndel mes", 14),
    ("Entradas", 11),
    ("Unidades\nvendidas", 11),
    ("Stock al final\ndel mes", 14),
    ("Movimientos", 12),
    ("Vendido", 14),
    ("Utilidad", 14),
]
_N_MESES = len(_COLS_MESES)  # 8

_MESES_ES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
    "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]

_NOTA_LISTA = (
    "Movimientos en orden de fecha (el orden del cardex). «Entra» y «Sale» en "
    "unidades; «Existencia» = stock después del movimiento. En una venta, "
    "«Precio unitario» es el precio al que se vendió (o el costo, si salió a "
    "costo) y «Ganancia» = venta − costo."
)
_NOTA_MESES = (
    "Stock al final del mes = stock al inicio + entradas − unidades vendidas. "
    "«Unidades vendidas» cuenta todas las salidas a los aviones (con precio o a "
    "costo); «Entradas» incluye compras, devoluciones y ajustes. Un mes sin "
    "movimientos arrastra el stock del anterior; el mes en curso muestra el "
    "stock de hoy, no un cierre."
)


def _etiqueta_mes(mes: str) -> str:
    """'2026-08' ⇒ 'Agosto 2026' (respaldo si el API no manda la etiqueta)."""
    try:
        anio, m = mes.split("-")[:2]
        return f"{_MESES_ES[int(m) - 1]} {int(anio)}"
    except (ValueError, IndexError):
        return mes or "Sin fecha"


def _fecha_celda(fecha: str):
    """'2026-08-29' ⇒ date (Excel la filtra y ordena); otra cosa, tal cual."""
    try:
        return date.fromisoformat((fecha or "")[:10])
    except ValueError:
        return fecha or None


def _cantidad(n: float | None):
    """10.0 ⇒ 10 (entero a la vista); 7.5 se queda."""
    if n is None:
        return None
    return int(n) if float(n).is_integer() else n


def _movimiento_txt(m: CardexLibroMovimiento) -> str:
    if m.tipo == "SALIDA":
        return "Salida a costo" if m.a_costo else "Venta"
    return {
        "ENTRADA": "Compra",
        "DEVOLUCION": "Devolución",
        "AJUSTE": "Ajuste",
    }.get(m.tipo, m.tipo.capitalize() if m.tipo else "")


def _titulo(ws, req: CardexLibroRequest, titulo: str, ncols: int) -> None:
    t = ws.cell(row=1, column=1, value=titulo)
    t.font = Font(bold=True, size=14, color=BRAND)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    s = ws.cell(row=2, column=1, value=_subtitulo(req))
    s.font = Font(italic=True, size=10, color=_GRIS)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)


def _encabezado(ws, row: int, cols: list[tuple[str, int]]) -> None:
    for col, (label, width) in enumerate(cols, start=1):
        cell = ws.cell(row=row, column=col, value=label)
        cell.font = Font(bold=True, color=WHITE)
        cell.fill = PatternFill("solid", fgColor=BRAND)
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
        cell.border = _border
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.row_dimensions[row].height = 30


def _nota_pie(ws, row: int, texto: str, ncols: int) -> None:
    n = ws.cell(row=row, column=1, value=texto)
    n.font = Font(italic=True, size=9, color=_GRIS)
    n.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    # Alto estimado: ~150 caracteres por renglón en el ancho de la tabla.
    ws.row_dimensions[row].height = 13 * (len(texto) // 150 + 1)


def _texto_banda(mes: str, info: CardexLibroMes | None) -> str:
    if info is None:
        return _etiqueta_mes(mes)
    partes = [
        info.etiqueta or _etiqueta_mes(mes),
        f"Stock al inicio: {_cantidad(info.stock_inicio)}",
        f"Entradas: {_cantidad(info.entradas)}",
        f"Vendidas: {_cantidad(info.unidades_vendidas)}",
        f"Stock al final: {_cantidad(info.stock_final)}",
    ]
    if info.movimientos == 0:
        partes.append("sin movimientos")
    return " · ".join(partes)


def _hoja_lista(ws, req: CardexLibroRequest) -> None:
    """Hoja 1: todos los movimientos en orden de fecha, con una banda por mes."""
    nombre = req.item_nombre or ""
    _titulo(ws, req, req.titulo or f"Cardex — {nombre}", _N_LISTA)
    hrow = 4
    _encabezado(ws, hrow, _COLS_LISTA)

    movs = req.movimientos or []
    info_mes = {m.mes: m for m in req.meses if m.mes}
    # Meses en orden: los del resumen (también los que no tuvieron
    # movimientos) y, por si acaso, cualquier mes de un movimiento.
    claves = sorted(set(info_mes) | {m.mes for m in movs if m.mes})
    sin_mes = [m for m in movs if not m.mes]

    r = hrow + 1
    if not movs and not claves:
        c = ws.cell(row=r, column=1, value="Sin movimientos.")
        c.font = Font(italic=True, color=_GRIS)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=_N_LISTA)
        r += 1

    def banda(texto: str) -> None:
        nonlocal r
        c = ws.cell(row=r, column=1, value=texto)
        c.font = Font(bold=True, color=BRAND)
        c.alignment = Alignment(vertical="center")
        for col in range(1, _N_LISTA + 1):
            x = ws.cell(row=r, column=col)
            x.fill = PatternFill("solid", fgColor=LIGHT)
            x.border = _border
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=_N_LISTA)
        ws.row_dimensions[r].height = 18
        r += 1

    def fila(m: CardexLibroMovimiento) -> None:
        nonlocal r
        f = ws.cell(row=r, column=1, value=_fecha_celda(m.fecha))
        f.number_format = "dd/mm/yyyy"
        f.border = _border
        ws.cell(row=r, column=2, value=_movimiento_txt(m)).border = _border
        d = ws.cell(row=r, column=3, value=m.detalle or None)
        d.border = _border
        d.alignment = Alignment(wrap_text=True, vertical="top")
        e = ws.cell(row=r, column=4, value=_cantidad(m.entra))
        e.border = _border
        if m.entra is not None:
            e.font = Font(bold=True, color=_VERDE)
        s = ws.cell(row=r, column=5, value=_cantidad(m.sale))
        s.border = _border
        if m.sale is not None:
            s.font = Font(bold=True, color=_ROJO)
        x = ws.cell(row=r, column=6, value=_cantidad(m.existencia))
        x.border = _border
        x.font = Font(bold=True)
        for col, v in ((7, m.precio_unitario), (8, m.total), (9, m.ganancia)):
            c = ws.cell(row=r, column=col, value=v)
            c.border = _border
            if v is not None:
                c.number_format = MONEY
                if col == 9 and v < 0:
                    c.font = Font(color=_ROJO)
        r += 1

    for mes in claves:
        banda(_texto_banda(mes, info_mes.get(mes)))
        for m in movs:
            if m.mes == mes:
                fila(m)
    if sin_mes:
        banda("Sin fecha")
        for m in sin_mes:
            fila(m)

    # Totales (los manda el API) y la existencia de hoy.
    r += 1
    existencia = (
        req.meses[-1].stock_final
        if req.meses
        else (movs[-1].existencia if movs else 0)
    )
    for etiqueta, col, valor, money in (
        ("TOTAL COMPRAS", 8, req.total_compra, True),
        ("TOTAL VENTAS / GANANCIA", 8, req.total_venta, True),
        ("EXISTENCIA ACTUAL", 6, _cantidad(existencia), False),
    ):
        for c in range(1, _N_LISTA + 1):
            x = ws.cell(row=r, column=c)
            x.fill = PatternFill("solid", fgColor=LIGHT)
            x.border = _border
            x.font = Font(bold=True)
        ws.cell(row=r, column=3, value=etiqueta)
        v = ws.cell(row=r, column=col, value=valor)
        if money:
            v.number_format = MONEY
        if etiqueta.startswith("TOTAL VENTAS"):
            g = ws.cell(row=r, column=9, value=req.total_ganancia)
            g.number_format = MONEY
        r += 1

    _nota_pie(ws, r + 1, _NOTA_LISTA, _N_LISTA)
    ws.freeze_panes = ws.cell(row=hrow + 1, column=1)


def _hoja_meses(ws, req: CardexLibroRequest) -> None:
    """Hoja 2: stock al inicio, entradas, vendidas y stock al final, por mes."""
    nombre = req.item_nombre or ""
    _titulo(ws, req, f"Resumen por mes — {nombre}".rstrip(" —"), _N_MESES)
    hrow = 4
    _encabezado(ws, hrow, _COLS_MESES)

    r = hrow + 1
    if not req.meses:
        c = ws.cell(row=r, column=1, value="Sin movimientos.")
        c.font = Font(italic=True, color=_GRIS)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=_N_MESES)
        r += 1

    for m in req.meses:
        valores = [
            m.etiqueta or _etiqueta_mes(m.mes),
            _cantidad(m.stock_inicio),
            _cantidad(m.entradas),
            _cantidad(m.unidades_vendidas),
            _cantidad(m.stock_final),
            m.movimientos,
            m.vendido,
            m.ganancia,
        ]
        for col, v in enumerate(valores, start=1):
            c = ws.cell(row=r, column=col, value=v)
            c.border = _border
            if col >= 7 and v is not None:
                c.number_format = MONEY
                if v < 0:
                    c.font = Font(color=_ROJO)
            if col in (2, 5):
                c.font = Font(bold=True)
            if m.movimientos == 0 and col != 7 and col != 8:
                c.font = Font(color=_GRIS, bold=col in (2, 5))
            if m.en_curso and col == 1:
                c.font = Font(italic=True)
        r += 1

    # Fila TOTAL: re-suma de columnas (solo para mostrar; el stock no se suma).
    if req.meses:
        def suma(vals):
            vals = [v for v in vals if v is not None]
            return round(sum(vals), 2) if vals else None

        tot = {
            1: "TOTAL",
            3: _cantidad(round(sum(m.entradas for m in req.meses), 3)),
            4: _cantidad(round(sum(m.unidades_vendidas for m in req.meses), 3)),
            5: _cantidad(req.meses[-1].stock_final),
            6: sum(m.movimientos for m in req.meses),
            7: suma([m.vendido for m in req.meses]),
            8: suma([m.ganancia for m in req.meses]),
        }
        for col in range(1, _N_MESES + 1):
            cell = ws.cell(row=r, column=col, value=tot.get(col))
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor=LIGHT)
            cell.border = _border
            if col >= 7 and tot.get(col) is not None:
                cell.number_format = MONEY
        r += 1

    _nota_pie(ws, r + 1, _NOTA_MESES, _N_MESES)
    ws.freeze_panes = ws.cell(row=hrow + 1, column=1)
