"""Hoja 'inventario' (tiendita) del Balance GENERAL: bloque por ítem +
detalle de salidas; sustituye a 'refacciones' SOLO en el general (fallback
skew tolerante con API viejo) y el libro INDIVIDUAL conserva la suya."""

from io import BytesIO

from openpyxl import load_workbook

from app.schemas.reportes import (
    BalanceAvionRequest,
    BalanceGeneralRequest,
)
from app.services.balance_avion_xlsx import (
    render_balance_avion_xlsx,
    render_balance_general_xlsx,
)

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
