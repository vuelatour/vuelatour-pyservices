"""Hoja 'inventario' (tiendita) del Balance GENERAL: bloque por ítem +
detalle de salidas; sustituye a 'refacciones' SOLO en el general (fallback
skew tolerante con API viejo) y el libro INDIVIDUAL conserva la suya.

25-sep-2026: utilidad de la tienda en DÓLARES (VENDIDO/UTILIDAD USD aparte
de los pesos), nota del margen, nota de utilidad incalculable, huella del
layout anterior (payload viejo ⇒ hoja idéntica) y el Excel del inventario
(export genérico) con Ubicación y Utilidad MXN/USD."""

import hashlib
import json
from io import BytesIO

from openpyxl import load_workbook

from app.schemas.reportes import (
    BalanceAvionRequest,
    BalanceGeneralRequest,
    BalanceHojaInventario,
)
from app.schemas.tabla import TablaXlsxRequest
from app.services.balance_avion_xlsx import (
    render_balance_avion_xlsx,
    render_balance_general_xlsx,
)
from app.services.tabla_xlsx import render_tabla_xlsx

_REFACCIONES = {
    "filas": [
        {
            "fecha": "2026-08-10",
            "categoria": "Refacción",
            "detalle": "XB-ABC · Salida de bodega: 2 × Filtro 108-1 (precio de venta)",
            "monto_mxn": 1600.0,
            "moneda_original": None,
            "monto_original": None,
            "matricula": "XB-ABC",
            "avion_color": "#FF0000",
            "costo_mxn": 1000.0,
            "venta_mxn": 1600.0,
        },
        {
            "fecha": "2026-08-12",
            "categoria": "Refacción",
            "detalle": "XB-DEF · Salida de bodega: 1 × Aceite (costo FIFO)",
            "monto_mxn": 500.0,
            "moneda_original": None,
            "monto_original": None,
            "matricula": "XB-DEF",
            "costo_mxn": 500.0,
            "venta_mxn": 500.0,
        },
    ],
    "total_mxn": 2100.0,
    "usd": 110.0,
    "usd_hr": None,
}

_INVENTARIO = {
    "filas": [
        {
            "nombre": "Filtro 108-1 · 108-1",
            "existencia": 8,
            "valor_costo_mxn": 4000.0,
            "compradas_cant": 10,
            "compradas_costo_mxn": 5000.0,
            "salidas_cant": 2,
            "vendido_mxn": 1600.0,
            "utilidad_mxn": 600.0,
            "matriculas": "XB-ABC + FLOTA",
        },
        {
            # Ítem con stock pero sin actividad del periodo: solo existencia.
            "nombre": "Bujía fina",
            "existencia": 4,
            "valor_costo_mxn": 800.0,
        },
    ],
    "total_piezas": 12,
    "total_valor_mxn": 4800.0,
    "total_compras_mxn": 5000.0,
    "total_vendido_mxn": 1600.0,
    "total_utilidad_mxn": 600.0,
}


# Caso REAL del cliente (22-sep-2026): «Aceite 15w 50» existencia 30 y una
# sola ENTRADA de 30 × 110 USD SIN tipo de cambio — la columna sumaba
# dólares y los rotulaba «$3,300.00 MXN». Ahora el valor viaja en su propia
# columna en dólares y el total en pesos NO lo incluye.
_INVENTARIO_USD = {
    "filas": [
        {
            "nombre": "Aceite 15w 50",
            "existencia": 30,
            "valor_costo_mxn": 0.0,  # 0 real (no hay ni un peso capturado)
            "valor_costo_usd": 3300.0,
            "sin_tc": True,
            "compradas_cant": 30,
            "compradas_costo_mxn": None,  # el API ya manda null sin T.C.
        },
        {
            # Ítem MIXTO: una capa en pesos + otra en USD sin T.C.
            "nombre": "Bujía fina",
            "existencia": 6,
            "valor_costo_mxn": 800.0,
            "valor_costo_usd": 220.0,
            "sin_tc": True,
        },
        {
            # Todo en pesos: sin valor en dólares ni bandera.
            "nombre": "Filtro 108-1 · 108-1",
            "existencia": 8,
            "valor_costo_mxn": 4000.0,
            "valor_costo_usd": None,
            "sin_tc": False,
            "compradas_cant": 10,
            "compradas_costo_mxn": 5000.0,
            "salidas_cant": 2,
            "vendido_mxn": 1600.0,
            "utilidad_mxn": 600.0,
            "matriculas": "XB-ABC + FLOTA",
        },
    ],
    "total_piezas": 44,
    "total_valor_mxn": 4800.0,  # SOLO pesos reales (800 + 4000)
    "total_valor_usd": 3520.0,  # aparte, en dólares (3300 + 220)
    "total_compras_mxn": 5000.0,
    "total_vendido_mxn": 1600.0,
    "total_utilidad_mxn": 600.0,
    "filas_sin_tc": 2,
}

_ENCABEZADOS_HOY = [
    "ÍTEM", "EXISTENCIA\nACTUAL", "VALOR A\nCOSTO MXN", "COMPRADAS\nCANT",
    "COMPRAS\nMXN", "SALIDAS\nCANT", "VENDIDO\nMXN", "UTILIDAD\nMXN",
    "MATRÍCULAS",
]


def _general(**extra) -> BalanceGeneralRequest:
    return BalanceGeneralRequest(
        periodo_desde="2026-08-01",
        periodo_hasta="2026-08-31",
        consolidado=BalanceAvionRequest(
            matricula="FLOTA",
            periodo_desde="2026-08-01",
            periodo_hasta="2026-08-31",
            refacciones=_REFACCIONES,
        ),
        **extra,
    )


def _sheet(data: bytes, nombre: str):
    return load_workbook(BytesIO(data))[nombre]


def test_general_con_inventario_pinta_hoja_y_no_refacciones():
    data = render_balance_general_xlsx(_general(inventario=_INVENTARIO))
    wb = load_workbook(BytesIO(data))
    assert "inventario" in wb.sheetnames
    assert "refacciones" not in wb.sheetnames

    ws = wb["inventario"]
    # Bloque 1: encabezados, fila por ítem y fila TOTALES.
    assert ws.cell(row=4, column=1).value == "ÍTEM"
    assert ws.cell(row=4, column=9).value == "MATRÍCULAS"
    assert ws.cell(row=5, column=1).value == "Filtro 108-1 · 108-1"
    assert ws.cell(row=5, column=2).value == 8
    assert ws.cell(row=5, column=5).value == 5000.0
    assert ws.cell(row=5, column=9).value == "XB-ABC + FLOTA"
    # None = celda vacía (ítem sin actividad del periodo), jamás 0 falso.
    assert ws.cell(row=6, column=4).value is None
    assert ws.cell(row=6, column=7).value is None
    assert ws.cell(row=7, column=1).value == "TOTALES"
    assert ws.cell(row=7, column=2).value == 12
    assert ws.cell(row=7, column=3).value == 4800.0
    assert ws.cell(row=7, column=8).value == 600.0

    # Bloque 2: detalle de salidas (la vieja hoja 'refacciones'): la columna
    # AVIÓN lleva la matrícula y el detalle pierde su prefijo 'MATRÍCULA · '.
    filas = {
        (c[0].value, c[1].value, c[2].value, c[3].value, c[4].value, c[5].value)
        for c in ws.iter_rows(min_col=1, max_col=6)
    }
    assert (
        "10/08/2026",
        "Salida de bodega: 2 × Filtro 108-1 (precio de venta)",
        "XB-ABC",
        1000.0,
        1600.0,
        600.0,  # GANANCIA = venta − costo (solo para mostrar)
    ) in filas
    assert (
        "12/08/2026",
        "Salida de bodega: 1 × Aceite (costo FIFO)",
        "XB-DEF",
        500.0,
        500.0,
        0.0,
    ) in filas
    # TOTALES del bloque 2 (re-suma de columnas, solo para mostrar).
    assert ("TOTALES", None, None, 1500.0, 2100.0, 600.0) in filas


def test_general_sin_inventario_conserva_refacciones():
    # API viejo (skew): sin `inventario` la hoja 'refacciones' sigue igual.
    data = render_balance_general_xlsx(_general())
    wb = load_workbook(BytesIO(data))
    assert "refacciones" in wb.sheetnames
    assert "inventario" not in wb.sheetnames


def _textos(ws) -> list[str]:
    return [c.value for fila in ws.iter_rows() for c in fila
            if isinstance(c.value, str)]


def test_usd_sin_tc_va_en_su_columna_en_dolares_y_no_al_total_en_pesos():
    """Caso REAL: 30 × 110 USD sin T.C. ⇒ MXN 0.00 y USD 3,300.00."""
    ws = _sheet(render_balance_general_xlsx(_general(inventario=_INVENTARIO_USD)),
                "inventario")
    # La columna nueva entra a la DERECHA de «VALOR A COSTO MXN»; las del
    # periodo corren un lugar (MATRÍCULAS pasa de la 9 a la 10).
    assert ws.cell(row=4, column=3).value == "VALOR A\nCOSTO MXN"
    assert ws.cell(row=4, column=4).value == "VALOR A COSTO\nUSD (sin T.C.)"
    assert ws.cell(row=4, column=5).value == "COMPRADAS\nCANT"
    assert ws.cell(row=4, column=10).value == "MATRÍCULAS"

    # Fila del aceite: el valor está en dólares, el peso es 0 REAL.
    assert ws.cell(row=5, column=1).value == "Aceite 15w 50"
    assert ws.cell(row=5, column=2).value == 30
    assert ws.cell(row=5, column=3).value == 0.0
    assert ws.cell(row=5, column=4).value == 3300.0
    # Formatos: dólares con "$" pegado, pesos como siempre (sin símbolo).
    assert ws.cell(row=5, column=4).number_format == '"$"#,##0.00'
    assert ws.cell(row=5, column=3).number_format == "#,##0.00"

    # Ítem MIXTO: los dos valores, cada uno en su moneda; jamás sumados.
    assert ws.cell(row=6, column=3).value == 800.0
    assert ws.cell(row=6, column=4).value == 220.0
    # Ítem todo en pesos: columna de dólares VACÍA (nunca un 0 falso).
    assert ws.cell(row=7, column=3).value == 4000.0
    assert ws.cell(row=7, column=4).value is None

    # TOTALES: Σ de cada columna por separado y el resto sin moverse.
    assert ws.cell(row=8, column=1).value == "TOTALES"
    assert ws.cell(row=8, column=2).value == 44
    assert ws.cell(row=8, column=3).value == 4800.0
    assert ws.cell(row=8, column=4).value == 3520.0
    assert ws.cell(row=8, column=4).number_format == '"$"#,##0.00'
    assert ws.cell(row=8, column=6).value == 5000.0  # compras MXN
    assert ws.cell(row=8, column=8).value == 1600.0  # vendido MXN
    assert ws.cell(row=8, column=9).value == 600.0  # utilidad MXN

    # Nota bajo la tabla (plural) + aclaración en la nota al pie.
    textos = _textos(ws)
    assert (
        "2 productos tienen costo en USD sin tipo de cambio: su valor se "
        "muestra en dólares y no entra al total en pesos."
    ) in textos
    assert any("VALOR A COSTO MXN suma SOLO las capas compradas en pesos" in t
               for t in textos)


def test_nota_sin_tc_en_singular_cuando_es_un_solo_producto():
    inv = {
        "filas": [_INVENTARIO_USD["filas"][0]],
        "total_piezas": 30,
        "total_valor_mxn": 0.0,
        "total_valor_usd": 3300.0,
        "filas_sin_tc": 1,
    }
    ws = _sheet(render_balance_general_xlsx(_general(inventario=inv)), "inventario")
    assert (
        "1 producto tiene costo en USD sin tipo de cambio: su valor se "
        "muestra en dólares y no entra al total en pesos."
    ) in _textos(ws)


def test_inventario_todo_en_pesos_no_gana_la_columna_usd():
    # API nuevo sin un solo movimiento en USD sin T.C.: la hoja es la de hoy.
    inv = {**_INVENTARIO, "total_valor_usd": 0.0, "filas_sin_tc": 0}
    ws = _sheet(render_balance_general_xlsx(_general(inventario=inv)), "inventario")
    assert [ws.cell(row=4, column=c).value for c in range(1, 10)] == _ENCABEZADOS_HOY
    assert ws.cell(row=4, column=10).value is None
    assert not any("USD" in t for t in _textos(ws))


def test_payload_sin_campos_nuevos_deja_la_hoja_identica():
    """Skew: API viejo (sin `valor_costo_usd` / `filas_sin_tc`) ⇒ misma hoja
    de siempre — encabezados y totales en sus columnas de hoy."""
    ws = _sheet(render_balance_general_xlsx(_general(inventario=_INVENTARIO)),
                "inventario")
    assert [ws.cell(row=4, column=c).value for c in range(1, 10)] == _ENCABEZADOS_HOY
    assert ws.cell(row=4, column=10).value is None
    # Fila TOTALES intacta (piezas, valorizado, compras, vendido, utilidad).
    assert ws.cell(row=7, column=1).value == "TOTALES"
    assert [ws.cell(row=7, column=c).value for c in range(2, 10)] == [
        12, 4800.0, None, 5000.0, None, 1600.0, 600.0, None,
    ]
    # Ni columna ni notas en dólares.
    assert not any("USD" in t for t in _textos(ws))


def test_individual_conserva_su_hoja_refacciones():
    data = render_balance_avion_xlsx(
        BalanceAvionRequest(matricula="XB-ABC", refacciones=_REFACCIONES)
    )
    wb = load_workbook(BytesIO(data))
    assert "refacciones" in wb.sheetnames
    assert "inventario" not in wb.sheetnames


# ---------------------------------------------------------------------------
# UTILIDAD DE LA TIENDA en dólares (25-sep-2026, API 0.0.35): toda salida sin
# precio se cobra al avión a costo FIFO + margen (25 %). Las ventas de prod
# son USD sobre costo USD sin T.C.: su utilidad viaja en VENDIDO USD /
# UTILIDAD USD, APARTE de las columnas en pesos — jamás sumadas.
# ---------------------------------------------------------------------------

MONEY_USD = '"$"#,##0.00'

# Septiembre 2026 tras el re-precio de las 10 salidas del 01-sep (tabla §0
# del contrato: venta y utilidad por salida con la regla del API). El aceite
# lleva además su existencia real (84 × 21.25 USD sin T.C.); el resto solo
# sus ventas del periodo.
_TIENDA_SEP26 = {
    "filas": [
        {
            "nombre": "Aceite multigrado semisintético 15W-50",
            "existencia": 84,
            "valor_costo_mxn": 0.0,
            "valor_costo_usd": 1785.0,
            "sin_tc": True,
            "salidas_cant": 36,
            "vendido_mxn": None,
            "utilidad_mxn": None,
            "vendido_usd": 956.25,  # 318.75 + 637.50
            "utilidad_usd": 191.25,  # 63.75 + 127.50
            "ventas_sin_utilidad": 0,
            "matriculas": "XA-VGV + N4142R",
        },
        {"nombre": "Filtro CH48108", "salidas_cant": 2, "vendido_usd": 115.15,
         "utilidad_usd": 23.03, "matriculas": "N4142R"},
        {"nombre": "Filtro CH48110", "salidas_cant": 1, "vendido_usd": 57.58,
         "utilidad_usd": 11.52, "matriculas": "XA-VGV"},
        {"nombre": "Cámara 6.00-6", "salidas_cant": 1, "vendido_usd": 194.93,
         "utilidad_usd": 38.99, "matriculas": "N4142R"},
        {"nombre": "Cámara 8.00-6", "salidas_cant": 1, "vendido_usd": 240.24,
         "utilidad_usd": 48.05, "matriculas": "XA-VGV"},
        {"nombre": "Llanta 6.00-6", "salidas_cant": 2, "vendido_usd": 934.38,
         "utilidad_usd": 186.88, "matriculas": "XA-VGV + N4142R"},
        {"nombre": "Balata 66-105", "salidas_cant": 4, "vendido_usd": 115.65,
         "utilidad_usd": 23.13, "matriculas": "XA-VGV"},
        {"nombre": "Cubre pitot", "salidas_cant": 1, "vendido_usd": 62.5,
         "utilidad_usd": 12.5, "matriculas": "N4142R"},
    ],
    "total_piezas": 84,
    "total_valor_mxn": 0.0,
    "total_valor_usd": 1785.0,
    "filas_sin_tc": 1,
    "total_compras_mxn": None,
    "total_vendido_mxn": None,
    "total_utilidad_mxn": None,
    "total_vendido_usd": 2676.68,
    "total_utilidad_usd": 535.35,
    "filas_utilidad_incompleta": 0,
    "margen_venta_pct": 25,
}

_ENCABEZADOS_TIENDA_USD = [
    "ÍTEM", "EXISTENCIA\nACTUAL", "VALOR A\nCOSTO MXN",
    "VALOR A COSTO\nUSD (sin T.C.)", "COMPRADAS\nCANT", "COMPRAS\nMXN",
    "SALIDAS\nCANT", "VENDIDO\nMXN", "UTILIDAD\nMXN", "VENDIDO\nUSD",
    "UTILIDAD\nUSD", "MATRÍCULAS",
]


def _firma_hoja(data: bytes, nombre: str = "inventario") -> str:
    """Huella del LAYOUT de una hoja (valores, formatos, fuentes, rellenos,
    alineación, bordes, merges, anchos, altos y panel fijo): cambia si
    cambia cualquier cosa visible, pero no si cambian OTRAS hojas del libro
    (los índices de estilo del XML sí se mueven con ellas)."""
    ws = load_workbook(BytesIO(data))[nombre]
    partes: list = []
    for fila in ws.iter_rows():
        for c in fila:
            f = c.font
            partes.append([
                c.coordinate, repr(c.value), c.number_format,
                bool(f.b), bool(f.i), f.sz,
                f.color.rgb if f.color is not None else None,
                c.fill.fgColor.rgb if c.fill.fill_type else None,
                c.alignment.horizontal, c.alignment.vertical,
                bool(c.alignment.wrap_text), c.border.left.style,
            ])
    partes.append(sorted(str(r) for r in ws.merged_cells.ranges))
    partes.append(sorted((k, d.width) for k, d in ws.column_dimensions.items()))
    partes.append(sorted((k, d.height) for k, d in ws.row_dimensions.items()
                         if d.height))
    partes.append(ws.freeze_panes)
    return hashlib.sha256(json.dumps(partes, default=str).encode()).hexdigest()


# Huellas calculadas con el código ANTERIOR a la utilidad en dólares (HEAD
# 6e89e37, 24-sep-2026). Además se verificó que los .xlsx completos salen
# byte-idénticos (todos los miembros del zip salvo docProps/core.xml, que
# lleva la hora). Si una cambia, la hoja de un API viejo YA NO es la de antes.
_FIRMA_PAYLOAD_VIEJO = "c0ab510e45b06cca6dc13f2cdd01b4365d5b855b60db48665930ae4044c36925"
_FIRMA_USD_SIN_TC_22SEP = "dcbc50c7f6e43939ff5c8441ced991d35b7505e861aec9e98d5b7f4df410b52b"


def _nota_pie(ws) -> str:
    return next(t for t in _textos(ws) if t.startswith("INVENTARIO (tiendita)"))


def _fila_de(ws, texto: str) -> int:
    return next(c.row for fila in ws.iter_rows(max_col=1) for c in fila
                if c.value == texto)


def test_payload_viejo_sigue_byte_identico_al_de_antes():
    """Skew: API 0.0.34 (sin campos de venta USD ni margen) ⇒ la hoja de
    SIEMPRE, también la del 22-sep con VALOR A COSTO USD."""
    viejo = render_balance_general_xlsx(_general(inventario=_INVENTARIO))
    assert _firma_hoja(viejo) == _FIRMA_PAYLOAD_VIEJO
    usd = render_balance_general_xlsx(_general(inventario=_INVENTARIO_USD))
    assert _firma_hoja(usd) == _FIRMA_USD_SIN_TC_22SEP


def test_campos_nuevos_en_null_no_cambian_nada():
    """API nuevo sin ventas en dólares (y NestJS mandando null en todo lo
    nuevo, conteos incluidos) ⇒ misma hoja que el payload viejo; nada de 422."""
    inv = {
        **_INVENTARIO,
        "filas": [
            {**f, "vendido_usd": None, "utilidad_usd": None,
             "ventas_sin_utilidad": None}
            for f in _INVENTARIO["filas"]
        ],
        "total_vendido_usd": None,
        "total_utilidad_usd": None,
        "filas_utilidad_incompleta": None,
        "margen_venta_pct": None,
    }
    modelo = BalanceHojaInventario.model_validate(inv)
    assert modelo.filas_utilidad_incompleta == 0
    assert all(f.ventas_sin_utilidad == 0 for f in modelo.filas)
    data = render_balance_general_xlsx(_general(inventario=inv))
    assert _firma_hoja(data) == _FIRMA_PAYLOAD_VIEJO


def test_tienda_real_septiembre_utilidad_en_dolares_en_su_columna():
    """Caso real tras el re-precio: 8 productos, 2,676.68 USD vendidos y
    535.35 USD de utilidad — en dólares, jamás rotulados como pesos."""
    ws = _sheet(render_balance_general_xlsx(_general(inventario=_TIENDA_SEP26)),
                "inventario")
    # Las DOS columnas en dólares conviven: VALOR A COSTO USD (posición 4) y
    # VENDIDO/UTILIDAD USD entre UTILIDAD MXN y MATRÍCULAS (12 columnas).
    assert [ws.cell(row=4, column=c).value for c in range(1, 13)] == \
        _ENCABEZADOS_TIENDA_USD
    assert ws.cell(row=4, column=13).value is None

    # Aceite: vendió en dólares; las celdas en pesos quedan VACÍAS.
    assert ws.cell(row=5, column=1).value == "Aceite multigrado semisintético 15W-50"
    assert ws.cell(row=5, column=4).value == 1785.0
    assert ws.cell(row=5, column=7).value == 36
    assert ws.cell(row=5, column=8).value is None
    assert ws.cell(row=5, column=9).value is None
    assert ws.cell(row=5, column=10).value == 956.25
    assert ws.cell(row=5, column=11).value == 191.25
    assert ws.cell(row=5, column=10).number_format == MONEY_USD
    assert ws.cell(row=5, column=11).number_format == MONEY_USD
    assert ws.cell(row=5, column=12).value == "XA-VGV + N4142R"
    # Llanta: dos salidas (una por avión) en una fila.
    assert ws.cell(row=10, column=1).value == "Llanta 6.00-6"
    assert ws.cell(row=10, column=11).value == 186.88

    # TOTALES: los de la tienda en dólares, cada uno en SU columna; los de
    # pesos vacíos (no hubo ni un peso vendido) — nunca la suma de ambos.
    fila_tot = _fila_de(ws, "TOTALES")
    assert fila_tot == 13
    assert ws.cell(row=fila_tot, column=8).value is None
    assert ws.cell(row=fila_tot, column=9).value is None
    assert ws.cell(row=fila_tot, column=10).value == 2676.68
    assert ws.cell(row=fila_tot, column=11).value == 535.35
    assert ws.cell(row=fila_tot, column=10).number_format == MONEY_USD
    assert ws.cell(row=fila_tot, column=11).number_format == MONEY_USD
    assert ws.cell(row=fila_tot, column=4).value == 1785.0
    # El relleno de la fila TOTALES cubre las 12 columnas.
    assert ws.cell(row=fila_tot, column=12).fill.fgColor.rgb.endswith("EEF2F7")

    # Nota al pie: VALOR A COSTO USD + VENDIDO/UTILIDAD USD + margen vigente.
    nota = _nota_pie(ws)
    assert "VALOR A COSTO MXN suma SOLO las capas compradas en pesos" in nota
    assert ("VENDIDO USD y UTILIDAD USD: salidas cobradas al avión en dólares "
            "sobre costo en dólares; se muestran en su moneda y no se suman "
            "con las columnas en pesos.") in nota
    assert nota.endswith(
        "Desde el 25-sep-2026 toda salida sin precio se cobra al avión a costo "
        "FIFO + 25 % (utilidad de la tienda).")
    # Sin productos con utilidad incalculable: no hay nota roja de eso.
    assert not any("su utilidad no se puede calcular" in t for t in _textos(ws))

    # Layout: título, notas y bloque 2 abarcan las 12 columnas; anchos con
    # los dos insertados (14, 14) antes del de MATRÍCULAS.
    rangos = {str(r) for r in ws.merged_cells.ranges}
    assert "A1:L1" in rangos
    fila_det = _fila_de(ws, "DETALLE DE SALIDAS DEL PERIODO (cargos a aviones)")
    assert f"A{fila_det}:L{fila_det}" in rangos
    fila_nota = _fila_de(ws, nota)
    assert f"A{fila_nota}:L{fila_nota}" in rangos
    anchos = {k: ws.column_dimensions[k].width for k in "DIJKL"}
    assert anchos == {"D": 18, "I": 14, "J": 14, "K": 14, "L": 24}


def test_bloque_2_no_se_mueve_con_las_columnas_en_dolares():
    ws = _sheet(render_balance_general_xlsx(_general(inventario=_TIENDA_SEP26)),
                "inventario")
    fila_det = _fila_de(ws, "DETALLE DE SALIDAS DEL PERIODO (cargos a aviones)")
    assert [ws.cell(row=fila_det + 1, column=c).value for c in range(1, 8)] == [
        "FECHA", "ÍTEM", "AVIÓN", "COSTO VUELATOUR\nMXN",
        "VENTA AL AVIÓN\nMXN", "GANANCIA\nMXN", None,
    ]
    assert ws.cell(row=fila_det + 2, column=3).value == "XB-ABC"
    assert ws.cell(row=fila_det + 2, column=6).value == 600.0


def test_venta_usd_sin_columna_de_valor_usd_desplaza_solo_matriculas():
    """Los dos desplazamientos son independientes: sin capas USD sin T.C.
    las columnas del periodo no se mueven y VENDIDO/UTILIDAD USD entran en
    la 9 y la 10 (MATRÍCULAS pasa a la 11)."""
    inv = {
        "filas": [{"nombre": "Balata 66-105", "existencia": 2,
                   "valor_costo_mxn": 900.0, "salidas_cant": 4,
                   "vendido_usd": 115.65, "utilidad_usd": 23.13,
                   "matriculas": "XA-VGV"}],
        "total_piezas": 2,
        "total_valor_mxn": 900.0,
        "total_valor_usd": 0.0,
        "filas_sin_tc": 0,
        "total_vendido_usd": 115.65,
        "total_utilidad_usd": 23.13,
    }
    ws = _sheet(render_balance_general_xlsx(_general(inventario=inv)), "inventario")
    assert [ws.cell(row=4, column=c).value for c in range(1, 12)] == [
        *_ENCABEZADOS_HOY[:8], "VENDIDO\nUSD", "UTILIDAD\nUSD", "MATRÍCULAS",
    ]
    assert ws.cell(row=5, column=6).value == 4  # SALIDAS CANT, sin moverse
    assert [ws.cell(row=5, column=c).value for c in (9, 10, 11)] == [
        115.65, 23.13, "XA-VGV"]
    assert [ws.cell(row=6, column=c).value for c in (1, 9, 10)] == [
        "TOTALES", 115.65, 23.13]
    assert "A1:K1" in {str(r) for r in ws.merged_cells.ranges}
    # Sin margen en el payload no se inventa la frase del margen.
    assert "costo FIFO +" not in _nota_pie(ws)


def test_pesos_y_dolares_juntos_cada_total_en_su_moneda():
    """Un producto vendido en pesos y otro en dólares: dos totales, jamás
    uno solo (600 + 23.13 o 1,600 + 115.65 no aparecen en ninguna celda)."""
    inv = {
        "filas": [
            {"nombre": "Filtro 108-1", "existencia": 8, "valor_costo_mxn": 4000.0,
             "salidas_cant": 2, "vendido_mxn": 1600.0, "utilidad_mxn": 600.0,
             "matriculas": "XB-ABC"},
            {"nombre": "Balata 66-105", "salidas_cant": 4, "vendido_usd": 115.65,
             "utilidad_usd": 23.13, "matriculas": "XA-VGV"},
        ],
        "total_piezas": 8,
        "total_valor_mxn": 4000.0,
        "total_vendido_mxn": 1600.0,
        "total_utilidad_mxn": 600.0,
        "total_vendido_usd": 115.65,
        "total_utilidad_usd": 23.13,
        "margen_venta_pct": 25,
    }
    ws = _sheet(render_balance_general_xlsx(_general(inventario=inv)), "inventario")
    fila_tot = _fila_de(ws, "TOTALES")
    assert [ws.cell(row=fila_tot, column=c).value for c in (7, 8, 9, 10)] == [
        1600.0, 600.0, 115.65, 23.13]
    assert ws.cell(row=fila_tot, column=8).number_format == "#,##0.00"
    assert ws.cell(row=fila_tot, column=10).number_format == MONEY_USD
    # La fila en pesos deja vacías las de dólares y viceversa.
    assert ws.cell(row=5, column=9).value is None
    assert ws.cell(row=6, column=8).value is None
    valores = {c.value for fila in ws.iter_rows() for c in fila
               if isinstance(c.value, (int, float))}
    assert not valores & {623.13, 1715.65, 1623.13, 715.65}


def test_sin_total_del_api_se_resuma_la_misma_columna():
    """Si el API mandara solo las filas, el total en dólares es la re-suma
    de ESA columna (misma moneda, solo para mostrar)."""
    inv = {"filas": list(_TIENDA_SEP26["filas"]), "total_piezas": 84, "total_valor_mxn": 0.0,
           "total_valor_usd": 1785.0, "filas_sin_tc": 1}
    ws = _sheet(render_balance_general_xlsx(_general(inventario=inv)), "inventario")
    fila_tot = _fila_de(ws, "TOTALES")
    assert ws.cell(row=fila_tot, column=10).value == 2676.68
    assert ws.cell(row=fila_tot, column=11).value == 535.35
    assert ws.cell(row=fila_tot, column=9).value is None


def test_utilidad_incompleta_nota_roja_singular_y_plural():
    """Venta en PESOS sobre costo en dólares sin T.C.: no hay utilidad en
    ninguna moneda ⇒ se dice bajo la tabla (en rojo), no se inventa."""
    fila = {"nombre": "Cámara 8.00-6", "existencia": 1, "valor_costo_mxn": 0.0,
            "valor_costo_usd": 192.19, "sin_tc": True, "salidas_cant": 1,
            "vendido_mxn": 4200.0, "utilidad_mxn": None,
            "ventas_sin_utilidad": 1, "matriculas": "XA-VGV"}
    inv = {"filas": [fila], "total_piezas": 1, "total_valor_mxn": 0.0,
           "total_valor_usd": 192.19, "filas_sin_tc": 1,
           "total_vendido_mxn": 4200.0, "filas_utilidad_incompleta": 1}
    ws = _sheet(render_balance_general_xlsx(_general(inventario=inv)), "inventario")
    esperado = ("1 producto con ventas en pesos sobre costo en dólares sin "
                "tipo de cambio: su utilidad no se puede calcular y no aparece "
                "en ninguna columna.")
    assert esperado in _textos(ws)
    celda = ws.cell(row=_fila_de(ws, esperado), column=1)
    assert celda.font.color.rgb.endswith("DC2626")
    # Sin ventas en dólares: la hoja NO gana VENDIDO/UTILIDAD USD.
    assert ws.cell(row=4, column=10).value == "MATRÍCULAS"
    # Va DESPUÉS de la nota de «sin tipo de cambio» del 22-sep.
    assert _fila_de(ws, esperado) == _fila_de(ws, _nota_sin_tc_1()) + 1

    # Sin el conteo del API se cuentan las filas (solo para el texto).
    inv2 = {"filas": [fila, {**fila, "nombre": "Llanta 6.00-6"}]}
    ws2 = _sheet(render_balance_general_xlsx(_general(inventario=inv2)), "inventario")
    assert ("2 productos con ventas en pesos sobre costo en dólares sin tipo "
            "de cambio: su utilidad no se puede calcular y no aparece en "
            "ninguna columna.") in _textos(ws2)


def _nota_sin_tc_1() -> str:
    return ("1 producto tiene costo en USD sin tipo de cambio: su valor se "
            "muestra en dólares y no entra al total en pesos.")


def test_nota_del_margen_redacta_el_numero_que_manda_el_api():
    def nota(pct):
        inv = {**_INVENTARIO, "margen_venta_pct": pct}
        ws = _sheet(render_balance_general_xlsx(_general(inventario=inv)),
                    "inventario")
        return _nota_pie(ws)

    assert nota(12.5).endswith("a costo FIFO + 12.5 % (utilidad de la tienda).")
    assert nota(30.0).endswith("a costo FIFO + 30 % (utilidad de la tienda).")
    assert nota(0).endswith(
        "Toda salida sin precio se cobra al avión a costo FIFO, sin utilidad "
        "(margen de la tienda en 0 %).")
    # Todo en pesos y con margen: la frase entra, las columnas USD no.
    ws = _sheet(render_balance_general_xlsx(_general(
        inventario={**_INVENTARIO, "margen_venta_pct": 25})), "inventario")
    assert [ws.cell(row=4, column=c).value for c in range(1, 10)] == _ENCABEZADOS_HOY
    assert ws.cell(row=4, column=10).value is None


# ---------------------------------------------------------------------------
# Excel del INVENTARIO (panel «Excel» → API GET /v1/inventory/items/export →
# /pdf/tabla-xlsx). Lo arma el export GENÉRICO: aquí no hay código propio,
# pero se congela que las columnas nuevas del API 0.0.35 (Ubicación y
# Utilidad MXN/USD separadas) salen cada una en su lugar, con None = celda
# vacía y totales solo en SU columna.
# ---------------------------------------------------------------------------

_COLUMNAS_EXPORT_INVENTARIO = [
    {"label": "Ítem"}, {"label": "Código"}, {"label": "No. parte"},
    {"label": "Categoría"}, {"label": "Ubicación"},
    {"label": "Stock", "tipo": "numero"}, {"label": "Unidad"},
    {"label": "Mínimo", "tipo": "numero"},
    {"label": "Costo FIFO (MXN)", "tipo": "money"},
    {"label": "Valor (MXN)", "tipo": "money"},
    {"label": "Valor USD (sin T.C.)", "tipo": "money"},
    {"label": "Utilidad (MXN)", "tipo": "money"},
    {"label": "Utilidad (USD)", "tipo": "money"},
]


def test_excel_del_inventario_ubicacion_y_utilidad_por_moneda():
    req = TablaXlsxRequest(
        titulo="Inventario valorizado",
        subtitulo=("Generado 2026-09-25 · utilidad: todo el historial · "
                   "margen vigente 25 %"),
        columnas=_COLUMNAS_EXPORT_INVENTARIO,
        filas=[
            ["Aceite multigrado semisintético 15W-50", "ACE-1550", "", "Aceites",
             "Bodega Cancún (anterior)", 84, "cuarto", 12, 0, 0, 1785.0,
             None, 191.25],
            ["Filtro 108-1", "FIL-108", "108-1", "Filtros", "Oficina nueva",
             8, "pieza", 2, 500.0, 4000.0, None, 600.0, None],
            ["Bujía fina", "", "", "Motor", "", 4, "pieza", None, 200.0,
             800.0, None, None, None],
        ],
        totales=["TOTAL", None, None, None, None, None, None, None, None,
                 4800.0, 1785.0, 600.0, 191.25],
    )
    ws = load_workbook(BytesIO(render_tabla_xlsx(req))).active
    # Título, subtítulo, blanco y encabezado en la fila 4.
    assert [ws.cell(row=4, column=c).value for c in range(1, 14)] == [
        c["label"] for c in _COLUMNAS_EXPORT_INVENTARIO]
    assert ws.cell(row=5, column=5).value == "Bodega Cancún (anterior)"
    assert ws.cell(row=6, column=5).value == "Oficina nueva"
    # Utilidad: cada moneda en SU columna; None = celda vacía (no 0).
    assert ws.cell(row=5, column=12).value is None
    assert ws.cell(row=5, column=13).value == 191.25
    assert ws.cell(row=6, column=12).value == 600.0
    assert ws.cell(row=6, column=13).value is None
    assert ws.cell(row=7, column=12).value is None
    assert ws.cell(row=7, column=13).value is None
    # Totales separados, jamás sumados (600 + 191.25 no existe).
    assert [ws.cell(row=8, column=c).value for c in (1, 12, 13)] == [
        "TOTAL", 600.0, 191.25]
    valores = {c.value for fila in ws.iter_rows() for c in fila
               if isinstance(c.value, (int, float))}
    assert 791.25 not in valores
    # Revisión adversaria (25-sep-2026): «Ubicación» e «Ítem» miden lo de su
    # texto más largo — con el ancho del encabezado (13 y 12) «Bodega Cancún
    # (anterior)» se cortaba justo en la marca «(anterior)».
    assert ws.column_dimensions["E"].width >= len("Bodega Cancún (anterior)") + 2
    assert ws.column_dimensions["A"].width >= len(
        "Aceite multigrado semisintético 15W-50")
    # Las de dinero conservan su regla (encabezado + 4).
    assert ws.column_dimensions["M"].width == len("Utilidad (USD)") + 4
