"""Columna FACTURA VUELATOUR en la hoja maestra de los balances (30-sep-2026,
API 0.0.45).

Pedido de Marie (para Ale), con la foto del Excel «balance-XA-VGV-2026-09-01-
2026-09-30», hoja «reporte horas XA-VGV»: «en balance por avión el reporte de
excel se ocupa que diga el num de factura que nosotros emitimos del servicio,
no aparece en la columna, y si puede salir en el reporte general también».

El API manda por fila `factura_vuelatour` YA resuelto con su cascada única
(`etiquetasFacturaDeVuelos`: CFDI timbrado vivo → facturas EMITIDAS vigentes
«A-0424» → `vuelo.factura_folio` tecleado → etiqueta del estatus → nada).
Aquí solo se pinta, TAL CUAL, en una columna nueva AL FINAL del bloque STATUS
DE COBROS — en el libro individual y en la hoja «reporte horas FLOTA» del
Balance general (las dos salen de `_hoja_maestra`). Sin el campo (API viejo) o
en null, la celda va vacía; la fila TOTALES también. Nada más se mueve: la
columna va al final, así que `_COBRO1_COL`, `_COMISION_COL` y las columnas de
la izquierda quedan donde estaban.
"""

from io import BytesIO

import pytest
from fastapi.testclient import TestClient
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from app.config import get_settings
from app.main import app
from app.schemas.reportes import (
    BalanceAvionRequest,
    BalanceAvionVuelo,
    BalanceGeneralRequest,
    BalanceHojaOtrosMovimientos,
)
from app.services.balance_avion_xlsx import (
    _COBRO1_COL,
    _COLS,
    _COMISION_COL,
    _NOTA_FACTURA_VUELATOUR,
    FILL_COBROS,
    LIGHT,
    render_balance_avion_xlsx,
    render_balance_general_xlsx,
)

ENCABEZADO = "FACTURA\nVUELATOUR"
# Índice (1-based) de la columna nueva: la ÚLTIMA de la hoja maestra.
COL = len(_COLS)
LETRA = get_column_letter(COL)
FILA_DATOS = 3  # filas 1-2 = encabezado de grupo / columna


def _vuelo(clave: str, **extra) -> dict:
    return {
        "clave": clave,
        "fecha": "2026-09-05",
        "ruta": "CUN-PTU-CUN",
        "estado": "COMPLETADO",
        "status_cobro": "Cobrado",
        "cobros": [{"fecha": "2026-09-05", "monto_mxn": 30000.0}],
        "cobrado_real_mxn": 30000.0,
        "cobrado_mxn": 30000.0,
        "por_cobrar_mxn": 0.0,
        "por_cobrar_usd": 0.0,
        **extra,
    }


# Vuelos como los de la foto: uno con factura emitida, uno con dos facturas
# vigentes, uno solo con el estatus, uno SIN factura y un cancelado que sí la
# trae (los cancelados también la llevan).
_VUELOS = [
    _vuelo("XA-VGV-300", factura_vuelatour="A-0424"),
    _vuelo("XA-VGV-301", factura_vuelatour="A-0425, A-0430"),
    _vuelo("XA-VGV-302", factura_vuelatour="Facturado"),
    _vuelo("XA-VGV-303", factura_vuelatour=None),
    _vuelo("XA-VGV-304", estado="CANCELADO", factura_vuelatour="A-0399"),
]
_ESPERADO = ["A-0424", "A-0425, A-0430", "Facturado", None, "A-0399"]
_TOTALES = {"cobrado_real_mxn": 150000.0, "por_cobrar_usd": 0.0}


def _individual(vuelos=None) -> BalanceAvionRequest:
    return BalanceAvionRequest(
        matricula="XA-VGV",
        periodo_desde="2026-09-01",
        periodo_hasta="2026-09-30",
        vuelos=_VUELOS if vuelos is None else vuelos,
        totales=_TOTALES,
    )


def _general(vuelos=None) -> BalanceGeneralRequest:
    return BalanceGeneralRequest(
        periodo_desde="2026-09-01",
        periodo_hasta="2026-09-30",
        resumen=[{"matricula": "XA-VGV", "color": "#FF0000", "vuelos": 5}],
        consolidado=BalanceAvionRequest(
            matricula="FLOTA",
            periodo_desde="2026-09-01",
            periodo_hasta="2026-09-30",
            vuelos=_VUELOS if vuelos is None else vuelos,
            totales=_TOTALES,
        ),
    )


def _maestra_individual(req: BalanceAvionRequest | None = None):
    data = render_balance_avion_xlsx(req or _individual())
    wb = load_workbook(BytesIO(data), data_only=True)
    return wb["reporte horas XA-VGV"]


def _maestra_general(req: BalanceGeneralRequest | None = None):
    data = render_balance_general_xlsx(req or _general())
    wb = load_workbook(BytesIO(data), data_only=True)
    return wb["reporte horas FLOTA"]


def _columna(ws, n: int) -> list:
    return [ws.cell(row=FILA_DATOS + i, column=COL).value for i in range(n)]


# ---------------------------------------------------------------------------
# Layout: la columna va al FINAL de STATUS DE COBROS y no mueve nada más.
# ---------------------------------------------------------------------------


def test_cols_la_columna_nueva_es_la_ultima_de_status_de_cobros() -> None:
    assert _COLS[-1] == ("STATUS DE COBROS", ENCABEZADO, "factura_vuelatour", None)
    # La vecina de la izquierda es la que antes cerraba la hoja.
    assert _COLS[-2][1:3] == ("POR COBRAR\nUSD", "por_cobrar_usd")
    assert [c[2] for c in _COLS].count("factura_vuelatour") == 1
    # Los índices derivados siguen apuntando a su columna (se buscan en
    # `_COLS`, pero se congela que no se corrieron).
    assert _COLS[_COBRO1_COL - 1][1] == "COBRO 1\nFECHA"
    assert _COLS[_COMISION_COL - 1][2] == "comision_vendedor_mxn"
    assert (COL, _COBRO1_COL, _COMISION_COL) == (47, 35, 29)


@pytest.mark.parametrize("hoja", ["individual", "general"])
def test_encabezado_grupo_relleno_y_ancho(hoja) -> None:
    ws = _maestra_individual() if hoja == "individual" else _maestra_general()
    assert ws.max_column == COL
    assert ws.cell(row=2, column=COL).value == ENCABEZADO
    assert ws.cell(row=2, column=COL - 1).value == "POR COBRAR\nUSD"
    # El título de grupo STATUS DE COBROS (fila 1, combinado) la abarca.
    grupo = next(
        m
        for m in ws.merged_cells.ranges
        if m.min_row == 1 and ws.cell(row=1, column=m.min_col).value == "STATUS DE COBROS"
    )
    assert grupo.max_col == COL
    assert ws.cell(row=1, column=grupo.min_col).value == "STATUS DE COBROS"
    # Mismo verde suave del bloque en encabezado y datos.
    assert ws.cell(row=2, column=COL).fill.fgColor.rgb.endswith(FILL_COBROS)
    assert ws.cell(row=FILA_DATOS, column=COL).fill.fgColor.rgb.endswith(FILL_COBROS)
    assert ws.column_dimensions[LETRA].width == 16
    # Los anchos de la izquierda no cambian (COBRADO REAL/AVIÓN 17, resto 13).
    anchos = {
        h: ws.column_dimensions[get_column_letter(i)].width
        for i, (_g, h, _a, _f) in enumerate(_COLS, start=1)
    }
    assert anchos["POR COBRAR\nUSD"] == 13
    assert anchos["COBRADO REAL\nMXN (Σ depósitos)"] == 17
    assert ws.freeze_panes == "D3"


# ---------------------------------------------------------------------------
# Valores: la etiqueta del API TAL CUAL; sin dato ⇒ vacía; TOTALES vacía.
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("hoja", ["individual", "general"])
def test_valor_tal_cual_y_vacio_sin_factura(hoja) -> None:
    ws = _maestra_individual() if hoja == "individual" else _maestra_general()
    assert _columna(ws, len(_VUELOS)) == _ESPERADO
    # Las filas son las del payload, en orden (la clave lo confirma).
    assert [ws.cell(row=FILA_DATOS + i, column=1).value for i in range(len(_VUELOS))] == [
        v["clave"] for v in _VUELOS
    ]
    # Texto, no número ni fecha.
    celda = ws.cell(row=FILA_DATOS, column=COL)
    assert celda.data_type == "s"
    assert celda.number_format == "General"


@pytest.mark.parametrize("hoja", ["individual", "general"])
def test_fila_totales_va_vacia(hoja) -> None:
    ws = _maestra_individual() if hoja == "individual" else _maestra_general()
    fila_tot = FILA_DATOS + len(_VUELOS)
    assert ws.cell(row=fila_tot, column=1).value == "TOTALES"
    celda = ws.cell(row=fila_tot, column=COL)
    assert celda.value is None
    assert celda.fill.fgColor.rgb.endswith(LIGHT)
    # La vecina sí trae su total: la fila TOTALES no se corrió.
    assert ws.cell(row=fila_tot, column=COL - 1).value == 0.0


def test_api_viejo_sin_la_clave_deja_la_columna_vacia() -> None:
    """Un API 0.0.44 no manda el campo: el libro sale igual salvo la columna
    nueva, que va vacía (con su encabezado y su relleno)."""
    viejos = [{k: v for k, v in f.items() if k != "factura_vuelatour"} for f in _VUELOS]
    for ws in (_maestra_individual(_individual(viejos)), _maestra_general(_general(viejos))):
        assert ws.cell(row=2, column=COL).value == ENCABEZADO
        assert _columna(ws, len(viejos)) == [None] * len(viejos)


def test_nada_a_la_izquierda_cambia_con_o_sin_factura() -> None:
    """Con y sin facturas registradas, todas las celdas de las columnas 1..46
    (valores, formatos y rellenos) son idénticas: la columna es aditiva."""
    sin = [{**f, "factura_vuelatour": None} for f in _VUELOS]
    a, b = _maestra_individual(), _maestra_individual(_individual(sin))
    assert a.max_row == b.max_row
    for fila in range(1, a.max_row + 1):
        for col in range(1, COL):
            ca, cb = a.cell(row=fila, column=col), b.cell(row=fila, column=col)
            assert (ca.value, ca.number_format, ca.fill.fgColor.rgb) == (
                cb.value,
                cb.number_format,
                cb.fill.fgColor.rgb,
            ), (fila, col)


def test_multi_avion_todas_sus_filas_traen_la_misma_factura() -> None:
    """La factura es del VUELO: el API manda la misma etiqueta en cada fila
    COMPARTIDO del general; aquí se pinta en las dos."""
    filas = [
        _vuelo(
            "XA-VGV-310 · 50 %",
            multi_avion=True,
            participacion=0.5,
            participacion_fuente="tramos",
            avion_color="#FF0000",
            factura_vuelatour="A-0431",
        ),
        _vuelo(
            "XB-PEV-310 · 50 %",
            multi_avion=True,
            participacion=0.5,
            participacion_fuente="tramos",
            avion_color="#00FF00",
            factura_vuelatour="A-0431",
        ),
    ]
    ws = _maestra_general(_general(filas))
    assert _columna(ws, 2) == ["A-0431", "A-0431"]


def test_folio_que_empieza_con_igual_se_queda_como_texto() -> None:
    """El folio lo teclea la oficina (`vuelo.factura_folio`): openpyxl vuelve
    FÓRMULA toda cadena con «=» al inicio y Excel abriría el libro con error."""
    ws = _maestra_individual(_individual([_vuelo("XA-VGV-320", factura_vuelatour="=A-12")]))
    celda = ws.cell(row=FILA_DATOS, column=COL)
    assert celda.value == "=A-12"
    assert celda.data_type == "s"


def test_otros_movimientos_folio_con_igual_tambien_es_texto() -> None:
    """La MISMA etiqueta viaja en «factura vuelatour» de 'otros movimientos'
    del general: sin la guarda ahí, el libro entero abriría con error aunque
    la hoja maestra la tenga. Los folios normales no cambian."""
    req = _general([_vuelo("XA-VGV-321", factura_vuelatour="=A-12")])
    req.consolidado.otros_movimientos = BalanceHojaOtrosMovimientos(
        filas=[
            {
                "clave": "XA-VGV-321",
                "concepto_ingreso": "TUA",
                "ingreso_mxn": 500.0,
                "factura": "=A-12",
            },
            {
                "clave": "XA-VGV-322",
                "concepto_ingreso": "TUA",
                "ingreso_mxn": 500.0,
                "factura": "A-0424",
            },
        ],
        filas_sueltas=[{"concepto_egreso": "Renta", "egreso_mxn": 100.0, "factura": "=B-1"}],
    )
    wb = load_workbook(BytesIO(render_balance_general_xlsx(req)), data_only=True)
    ws = wb["otros movimientos"]
    celdas = {c.value: c for c in ws["J"] if c.value in ("=A-12", "A-0424", "=B-1")}
    assert set(celdas) == {"=A-12", "A-0424", "=B-1"}
    assert all(c.data_type == "s" for c in celdas.values())
    # Y la hoja maestra del mismo libro también.
    assert wb["reporte horas FLOTA"].cell(row=FILA_DATOS, column=COL).data_type == "s"


# ---------------------------------------------------------------------------
# Nota al pie: en los dos libros, una sola vez.
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("hoja", ["individual", "general"])
def test_nota_al_pie_explica_la_columna(hoja) -> None:
    ws = _maestra_individual() if hoja == "individual" else _maestra_general()
    notas = [
        c.value
        for c in ws["A"]
        if isinstance(c.value, str) and c.value.startswith("FACTURA VUELATOUR =")
    ]
    assert notas == [_NOTA_FACTURA_VUELATOUR]
    assert notas[0].startswith(
        "FACTURA VUELATOUR = folio de la factura del servicio emitida al "
        "cliente (timbrada, registrada en Facturas emitidas o capturada en el "
        "vuelo)"
    )
    # Va justo después de la nota del ESTATUS DE COBRO (bloque de cobros).
    fila = next(c.row for c in ws["A"] if c.value == _NOTA_FACTURA_VUELATOUR)
    assert ws.cell(row=fila - 1, column=1).value.startswith("El ESTATUS DE COBRO por vuelo")


# ---------------------------------------------------------------------------
# Esquema: ADITIVO y LIBERAL (jamás un 422 que tumbe el balance entero).
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    ("entrada", "esperado"),
    [
        ("A-0424", "A-0424"),
        ("  A-0424  ", "A-0424"),
        ("", None),
        ("   ", None),
        (None, None),
        (1234, "1234"),
        (1234.0, "1234"),
        (12.5, "12.5"),
        (float("nan"), None),
        (True, None),
        ({"serie": "A"}, None),
    ],
)
def test_esquema_normaliza_lo_que_llega(entrada, esperado) -> None:
    assert BalanceAvionVuelo(factura_vuelatour=entrada).factura_vuelatour == esperado


def test_esquema_sin_la_clave_es_none() -> None:
    assert BalanceAvionVuelo().factura_vuelatour is None


# ---------------------------------------------------------------------------
# Rutas: el payload del API 0.0.45 entra sin 422 en los dos libros.
# ---------------------------------------------------------------------------

TOKEN = "secreto-de-prueba"
client = TestClient(app)


def test_rutas_aceptan_el_payload_nuevo(monkeypatch, request) -> None:
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    # Aunque una aserción falle, el token de prueba no se queda en caché.
    request.addfinalizer(get_settings.cache_clear)
    headers = {"X-Internal-Token": TOKEN}
    # JSON crudo como lo manda NestJS (con un folio numérico de más).
    vuelos = [*_VUELOS, _vuelo("XA-VGV-305", factura_vuelatour=1234)]

    res = client.post(
        "/pdf/balance-avion-xlsx",
        json={
            "matricula": "XA-VGV",
            "periodo_desde": "2026-09-01",
            "periodo_hasta": "2026-09-30",
            "vuelos": vuelos,
            "totales": _TOTALES,
        },
        headers=headers,
    )
    assert res.status_code == 200, res.text
    ws = load_workbook(BytesIO(res.content), data_only=True)["reporte horas XA-VGV"]
    assert _columna(ws, len(vuelos)) == [*_ESPERADO, "1234"]

    res = client.post(
        "/pdf/balance-general-xlsx",
        json={
            "periodo_desde": "2026-09-01",
            "periodo_hasta": "2026-09-30",
            "consolidado": {"matricula": "FLOTA", "vuelos": vuelos, "totales": _TOTALES},
        },
        headers=headers,
    )
    assert res.status_code == 200, res.text
    ws = load_workbook(BytesIO(res.content), data_only=True)["reporte horas FLOTA"]
    assert _columna(ws, len(vuelos)) == [*_ESPERADO, "1234"]
