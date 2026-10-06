"""Extensión de horario pagada en la hoja maestra de los balances (1-oct-2026,
API 0.0.47).

Pedido de Ale, con la captura del balance N4142R, vuelo #192 (Srta. Mariana,
06-sep, CUN-CTM-CUN): «en este vuelo me se está poniendo la extensión de
servicios como Operación y no va en ese apartado». El gasto real es la
factura del aeropuerto de Chetumal «AE-Extension y/o antelacion de horario»
$4,549.06 (3,921.60 + IVA 627.46). Mismo caso en el #190 (XB-PEV, 25-ago,
$4,549.04) y, sin lectura IA, el #314 (N4142R, «extensión de servicio
inspector Baraona $500 efectivo»).

Regla (vive en el API): la extensión y/o antelación de horario de un
aeropuerto es un TRASLADO al cliente, igual que el TUA — no es costo de
operar el avión. El API ya la saca de OPERACIONES/OTROS y arma la nota de la
celda («Extensión de horario (IVA incluido) $X**» en `op_detalle`). A
pyservices le llegan dos llaves ADITIVAS, solo cuando hay monto ≠ 0:
`vuelos[].extension_pagada_mxn` y `totales.extension_pagada_mxn`. Aquí solo:
  1) un renglón informativo bajo TOTALES, después del «TUA pagado»;
  2) el pie ** ampliado (y la nota del combustible) SOLO con la llave.
Sin la llave (API ≤ 0.0.46 o periodo sin extensiones) el libro es el de
siempre, byte a byte.
"""

import hashlib
import json
import zipfile
from io import BytesIO

import pytest
from fastapi.testclient import TestClient
from openpyxl import load_workbook

from app.config import get_settings
from app.main import app
from app.schemas.reportes import (
    BalanceAvionRequest,
    BalanceAvionTotales,
    BalanceAvionVuelo,
    BalanceGeneralRequest,
)
from app.services.balance_avion_xlsx import (
    _COLS,
    _NOTA_TRASLADOS_EXTENSION,
    _NOTA_TRASLADOS_TUA,
    render_balance_avion_xlsx,
    render_balance_general_xlsx,
)

FILA_DATOS = 3  # filas 1-2 = encabezado de grupo / columna
COL_OP = next(i for i, c in enumerate(_COLS, start=1) if c[2] == "op_mxn")
RENGLON_TUA = (
    "TUA pagado del periodo (solo nota en OPERACIONES, no resta en este libro):"
)
RENGLON_EXT = (
    "Extensión de horario pagada del periodo (solo nota en OPERACIONES, no "
    "resta en este libro):"
)
# Las notas del combustible y del pie ** con y sin extensión.
COMBUSTIBLE_TUA = "(y el TUA pagado tampoco resta, ver **)"
COMBUSTIBLE_EXT = "(ni el TUA pagado ni la extensión de horario restan, ver **)"
IVA_TUA = "(sin gas ni TUA)."
IVA_EXT = "(sin gas, TUA ni extensión de horario)."

# Montos REALES de prod (gastos 8ab208c4, ccf37888 y 0d2c6f0c).
EXT_192 = 4549.06  # round2(3921.60 × 1.16 = 4549.056)
EXT_190 = 4549.04  # 3921.59 + 627.45
EXT_314 = 500.0  # sin IA: todo el monto por el texto de las notas
# Notas de la celda OPERACIONES tal como las arma el API (`op_detalle`).
NOTA_192 = (
    "Operaciones · GRUPO AEROPORTUARIO … OLMECA-MAYA-MEXICA — Extensión de "
    "horario (IVA incluido) $4,549.06**"
)
NOTA_314 = (
    "Operaciones · extensión de servicio inspector Baraona (efectivo) — "
    "Extensión de horario (IVA incluido) $500.00**"
)


def _vuelo(clave: str, **extra) -> dict:
    return {
        "clave": clave,
        "fecha": "2026-09-06",
        "ruta": "CUN-CTM-CUN",
        "estado": "COMPLETADO",
        "status_cobro": "Cobrado",
        "cobros": [{"fecha": "2026-09-06", "monto_mxn": 30000.0}],
        "cobrado_real_mxn": 30000.0,
        "cobrado_mxn": 30000.0,
        "por_cobrar_mxn": 0.0,
        "por_cobrar_usd": 0.0,
        **extra,
    }


def _vuelos_sep(*, con_llave: bool) -> list[dict]:
    """Septiembre de N4142R: el #192 (solo extensión: OPERACIONES vacía, la
    nota la arma el API), el #314 (operación propia + extensión por texto) y
    un vuelo con TUA y sin extensión. Sin la llave = lo que manda un API
    0.0.46 (o el 0.0.47 en un periodo sin extensiones)."""

    def ext(monto: float) -> dict:
        return {"extension_pagada_mxn": monto} if con_llave else {}

    return [
        _vuelo("N4142R-192", op_detalle=[NOTA_192], **ext(EXT_192)),
        _vuelo(
            "N4142R-314",
            fecha="2026-09-15",
            ruta="CUN-CZM-CUN",
            op_mxn=1200.0,
            op_detalle=["Operaciones · Aterrizaje CZM — $1,200.00", NOTA_314],
            **ext(EXT_314),
        ),
        _vuelo(
            "N4142R-320",
            fecha="2026-09-20",
            ruta="CUN-PTU-CUN",
            op_mxn=600.0,
            op_detalle=[
                "Operaciones · ASUR — Operación $600.00 · TUA (IVA incluido) $939.60**"
            ],
            tua_pagado_mxn=939.6,
        ),
    ]


def _totales(*, con_llave: bool, extension: float = EXT_192 + EXT_314) -> dict:
    base = {
        "op_mxn": 1800.0,
        "tua_pagado_mxn": 939.6,
        "otros_ingresos_usd": 1392.0,  # «Extensión de servicios» $1,200 + IVA
        "cobrado_real_mxn": 90000.0,
    }
    return {**base, "extension_pagada_mxn": round(extension, 2)} if con_llave else base


def _individual(vuelos=None, totales=None, *, con_llave: bool = True) -> BalanceAvionRequest:
    return BalanceAvionRequest(
        matricula="N4142R",
        periodo_desde="2026-09-01",
        periodo_hasta="2026-09-30",
        vuelos=_vuelos_sep(con_llave=con_llave) if vuelos is None else vuelos,
        totales=_totales(con_llave=con_llave) if totales is None else totales,
    )


def _vuelos_flota(*, con_llave: bool) -> list[dict]:
    ext = {"extension_pagada_mxn": EXT_190} if con_llave else {}
    previo = _vuelo(
        "XB-PEV-190",
        fecha="2026-08-25",
        avion_color="#2563EB",
        op_detalle=[
            "Operaciones · Aeropuerto Internacional de Chetumal — Extensión de "
            "horario (IVA incluido) $4,549.04**"
        ],
        **ext,
    )
    return [previo, *_vuelos_sep(con_llave=con_llave)]


def _general(*, con_llave: bool = True) -> BalanceGeneralRequest:
    return BalanceGeneralRequest(
        periodo_desde="2026-08-01",
        periodo_hasta="2026-09-30",
        resumen=[
            {"matricula": "N4142R", "color": "#FF0000", "vuelos": 3},
            {"matricula": "XB-PEV", "color": "#2563EB", "vuelos": 1},
        ],
        consolidado=BalanceAvionRequest(
            matricula="FLOTA",
            periodo_desde="2026-08-01",
            periodo_hasta="2026-09-30",
            vuelos=_vuelos_flota(con_llave=con_llave),
            totales=_totales(
                con_llave=con_llave, extension=EXT_190 + EXT_192 + EXT_314
            ),
        ),
    )


def _maestra(data: bytes, nombre: str = "reporte horas N4142R"):
    # data_only=True (5-oct-2026): las celdas calculadas son FÓRMULAS con el
    # número del API en caché — se lee el número, como lo ve el cliente.
    return load_workbook(BytesIO(data), data_only=True)[nombre]


def _fila_de(ws, texto: str) -> int | None:
    return next((c.row for c in ws["A"] if c.value == texto), None)


def _textos_a(ws) -> list[str]:
    return [c.value for c in ws["A"] if isinstance(c.value, str)]


def _nota_combustible(ws) -> str:
    return next(t for t in _textos_a(ws) if t.startswith("El COMBUSTIBLE ya no va por vuelo"))


def _nota_traslados(ws) -> str:
    return next(t for t in _textos_a(ws) if t.startswith("** "))


def _miembros(data: bytes) -> dict[str, bytes]:
    """Todos los miembros del .xlsx salvo docProps/core.xml (lleva la hora)."""
    z = zipfile.ZipFile(BytesIO(data))
    return {n: z.read(n) for n in sorted(z.namelist()) if n != "docProps/core.xml"}


def _firma_hoja(data: bytes, nombre: str) -> str:
    """Huella del LAYOUT de una hoja (misma receta que
    test_balance_inventario_xlsx): valores, formatos, fuentes, rellenos,
    alineación, bordes, merges, anchos, altos, panel fijo y comentarios.
    Fórmulas visibles (5-oct-2026): se lee la CACHÉ (`data_only=True`, el
    número que muestra la celda) y se quitan las constantes nuevas del
    encabezado (A1:D1: «Permiso AFAC USD/hr» y «Factor IVA costos»), que el
    código anterior no pintaba — así la huella de antes sigue valiendo y
    prueba que nada más de la hoja cambió."""
    ws = load_workbook(BytesIO(data), data_only=True)[nombre]
    for col in range(1, 5):
        ws._cells.pop((1, col), None)
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
                c.comment.text if c.comment else None,
            ])
    partes.append(sorted(str(r) for r in ws.merged_cells.ranges))
    partes.append(sorted((k, d.width) for k, d in ws.column_dimensions.items()))
    partes.append(sorted((k, d.height) for k, d in ws.row_dimensions.items() if d.height))
    partes.append(ws.freeze_panes)
    return hashlib.sha256(json.dumps(partes, default=str).encode()).hexdigest()


# Huellas de la hoja maestra calculadas con el código ANTERIOR a este cambio
# (HEAD ef4d1f4, 30-sep-2026) para los payloads SIN la llave. Además se
# verificó que los .xlsx completos salían byte-idénticos (todos los miembros
# del zip salvo docProps/core.xml). Si una cambia, un API 0.0.46 —o un
# periodo sin extensiones— ya no ve su libro de siempre.
_FIRMA_INDIVIDUAL_SIN_LLAVE = "caea831da37e6e8c97184dd2d8510f74a9ec725e2a969449038a1533d385efbf"
_FIRMA_GENERAL_SIN_LLAVE = "88a80db60c6a4638b0bbbe9d94b57fe99b4bdbee07942e4243b0bb1e5e3114d6"


# ---------------------------------------------------------------------------
# Skew: sin la llave el libro es el de siempre (byte a byte).
# ---------------------------------------------------------------------------


def test_sin_la_llave_la_hoja_maestra_es_la_de_siempre() -> None:
    ind = render_balance_avion_xlsx(_individual(con_llave=False))
    assert _firma_hoja(ind, "reporte horas N4142R") == _FIRMA_INDIVIDUAL_SIN_LLAVE
    gen = render_balance_general_xlsx(_general(con_llave=False))
    assert _firma_hoja(gen, "reporte horas FLOTA") == _FIRMA_GENERAL_SIN_LLAVE
    ws = _maestra(ind)
    assert _fila_de(ws, RENGLON_EXT) is None
    assert _nota_traslados(ws) == _NOTA_TRASLADOS_TUA
    assert COMBUSTIBLE_TUA in _nota_combustible(ws)
    assert _nota_combustible(ws).endswith(IVA_TUA)


@pytest.mark.parametrize("valor", [None, 0, 0.0])
def test_llave_en_null_o_cero_no_cambia_ni_un_byte(valor) -> None:
    """El API solo manda la llave con monto ≠ 0; si llegara en null o 0 (fila
    y totales), el .xlsx sale idéntico al de un payload sin ella."""
    sin = render_balance_avion_xlsx(_individual(con_llave=False))
    vuelos = [{**v, "extension_pagada_mxn": valor} for v in _vuelos_sep(con_llave=False)]
    totales = {**_totales(con_llave=False), "extension_pagada_mxn": valor}
    con = render_balance_avion_xlsx(_individual(vuelos, totales))
    assert _miembros(con) == _miembros(sin)


# ---------------------------------------------------------------------------
# Renglón informativo bajo TOTALES, después del «TUA pagado».
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("hoja", ["individual", "general"])
def test_renglon_de_extension_tras_el_del_tua(hoja) -> None:
    if hoja == "individual":
        ws, esperado = _maestra(render_balance_avion_xlsx(_individual())), 5049.06
    else:
        ws = _maestra(render_balance_general_xlsx(_general()), "reporte horas FLOTA")
        esperado = 9598.1  # #190 + #192 + #314
    fila_tua = _fila_de(ws, RENGLON_TUA)
    fila_ext = _fila_de(ws, RENGLON_EXT)
    assert fila_tua is not None and fila_ext == fila_tua + 1
    # Mismo formato que el del TUA: A:H combinadas, monto en I, «MXN» en J.
    for fila in (fila_tua, fila_ext):
        assert f"A{fila}:H{fila}" in {str(r) for r in ws.merged_cells.ranges}
        a, i, j = (ws.cell(row=fila, column=c) for c in (1, 9, 10))
        assert (a.font.b, a.font.sz) == (True, 9)
        assert i.font.b is True
        assert (j.value, j.font.b, j.font.sz) == ("MXN", True, 9)
    monto = ws.cell(row=fila_ext, column=9)
    assert monto.value == pytest.approx(esperado)
    assert monto.number_format == ws.cell(row=fila_tua, column=9).number_format
    # Una línea en blanco y luego las notas al pie, como siempre.
    assert ws.cell(row=fila_ext + 1, column=1).value is None
    assert ws.cell(row=fila_ext + 2, column=1).value.startswith("* VENTA AVIÓN")


def test_sin_tua_el_renglon_de_extension_va_solo() -> None:
    """Periodo con extensión y SIN TUA pagado: el renglón aparece igual (no
    depende del TUA) y el del TUA no se pinta."""
    vuelos = _vuelos_sep(con_llave=True)[:2]
    totales = {**_totales(con_llave=True), "tua_pagado_mxn": None}
    ws = _maestra(render_balance_avion_xlsx(_individual(vuelos, totales)))
    assert _fila_de(ws, RENGLON_TUA) is None
    fila = _fila_de(ws, RENGLON_EXT)
    assert fila is not None
    assert ws.cell(row=fila, column=9).value == pytest.approx(5049.06)
    # Va justo bajo el de «TUAs/extras/… COTIZADOS» (ingreso de VuelaTour).
    assert ws.cell(row=fila - 1, column=1).value.startswith(
        "TUAs/extras/pernocta/comisión del vendedor COTIZADOS"
    )


def test_la_extension_no_suma_en_ninguna_columna() -> None:
    """Las filas y TOTALES pintan lo que manda el API: con y sin la llave, las
    celdas de datos (valores, formatos, rellenos y comentarios) son las
    mismas — OPERACIONES del #192 sigue vacía y la nota llega tal cual."""
    con = _maestra(render_balance_avion_xlsx(_individual()))
    sin = _maestra(render_balance_avion_xlsx(_individual(con_llave=False)))
    fila_tot = FILA_DATOS + 3
    assert con.cell(row=fila_tot, column=1).value == "TOTALES"
    for fila in range(1, fila_tot + 1):
        for col in range(1, len(_COLS) + 1):
            a, b = con.cell(row=fila, column=col), sin.cell(row=fila, column=col)
            assert (a.value, a.number_format, a.fill.fgColor.rgb) == (
                b.value,
                b.number_format,
                b.fill.fgColor.rgb,
            ), (fila, col)
            assert (a.comment.text if a.comment else None) == (
                b.comment.text if b.comment else None
            ), (fila, col)
    # #192: OPERACIONES vacía; la nota del API va en el comentario tal cual.
    op_192 = con.cell(row=FILA_DATOS, column=COL_OP)
    assert op_192.value is None
    assert op_192.comment.text == NOTA_192
    # #314: OPERACIONES = solo su operación propia (1,200), no 1,700.
    assert con.cell(row=FILA_DATOS + 1, column=COL_OP).value == 1200.0
    assert con.cell(row=fila_tot, column=COL_OP).value == 1800.0


# ---------------------------------------------------------------------------
# Pie ** y nota del combustible: ampliados SOLO con la llave.
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("hoja", ["individual", "general"])
def test_pie_ampliado_con_la_llave(hoja) -> None:
    if hoja == "individual":
        ws = _maestra(render_balance_avion_xlsx(_individual()))
    else:
        ws = _maestra(render_balance_general_xlsx(_general()), "reporte horas FLOTA")
    pie = _nota_traslados(ws)
    assert pie == _NOTA_TRASLADOS_EXTENSION
    assert pie == (
        "** TUA y extensión de horario (extensión y/o antelación de horario "
        "del aeropuerto) son traslados al pasajero: NO son costo del avión ni "
        "restan en ningún lado de este libro; quedan solo como nota en la "
        "celda OPERACIONES ('TUA $x**', 'Extensión de horario (IVA incluido) "
        "$x**'); lo cobrado y lo pagado viven en 'otros movimientos' del "
        "Balance general. Los servicios FBO sí son costo (columna OTROS)."
    )
    assert [t for t in _textos_a(ws) if t.startswith("** ")] == [pie]
    comb = _nota_combustible(ws)
    assert COMBUSTIBLE_EXT in comb and COMBUSTIBLE_TUA not in comb
    assert comb.endswith(IVA_EXT)
    # El general conserva su remisión al libro individual de cada avión.
    assert ("del libro individual de cada avión" in comb) is (hoja == "general")
    # Mismo lugar del pie: entre la del combustible y la del multi-avión.
    fila = _fila_de(ws, pie)
    assert ws.cell(row=fila - 1, column=1).value == comb
    assert ws.cell(row=fila + 1, column=1).value.startswith("*** Filas 'COMPARTIDO'")


def test_pie_de_siempre_es_el_texto_previo() -> None:
    """Congela el texto del TUA solo: es el que ve todo libro sin extensiones."""
    assert _NOTA_TRASLADOS_TUA == (
        "** El TUA pagado al aeropuerto NO es costo del avión ni resta en "
        "ningún lado de este libro: queda solo como nota en la celda "
        "OPERACIONES ('TUA $x**'); cobro y pago del TUA viven en 'otros "
        "movimientos' del Balance general. Los servicios FBO sí son costo "
        "(columna OTROS)."
    )


def test_llave_solo_en_una_fila_amplia_el_pie_sin_renglon_de_totales() -> None:
    """Skew raro (la fila trae la llave y los totales no): el pie se amplía
    porque la nota de la celda dice «Extensión de horario …**», pero el
    renglón de totales solo sale con `totales.extension_pagada_mxn`."""
    ws = _maestra(
        render_balance_avion_xlsx(
            _individual(_vuelos_sep(con_llave=True), _totales(con_llave=False))
        )
    )
    assert _fila_de(ws, RENGLON_EXT) is None
    assert _nota_traslados(ws) == _NOTA_TRASLADOS_EXTENSION


def test_llave_solo_en_totales_tambien_amplia_el_pie() -> None:
    ws = _maestra(
        render_balance_avion_xlsx(
            _individual(_vuelos_sep(con_llave=False), _totales(con_llave=True))
        )
    )
    assert _fila_de(ws, RENGLON_EXT) is not None
    assert _nota_traslados(ws) == _NOTA_TRASLADOS_EXTENSION


def test_extension_negativa_tambien_se_pinta() -> None:
    """Una nota de crédito del aeropuerto (≠ 0) se pinta tal cual: aquí no se
    decide qué es válido, solo se muestra lo que calculó el API."""
    totales = {**_totales(con_llave=False), "extension_pagada_mxn": -500.0}
    ws = _maestra(render_balance_avion_xlsx(_individual(totales=totales)))
    fila = _fila_de(ws, RENGLON_EXT)
    assert ws.cell(row=fila, column=9).value == -500.0


# ---------------------------------------------------------------------------
# Esquema y rutas: ADITIVO; el JSON crudo del API 0.0.47 entra sin 422.
# ---------------------------------------------------------------------------


def test_esquema_aditivo() -> None:
    assert BalanceAvionVuelo().extension_pagada_mxn is None
    assert BalanceAvionTotales().extension_pagada_mxn is None
    assert BalanceAvionVuelo(extension_pagada_mxn=4549.06).extension_pagada_mxn == 4549.06
    assert BalanceAvionTotales(extension_pagada_mxn=5049).extension_pagada_mxn == 5049.0


TOKEN = "secreto-de-prueba"
client = TestClient(app)


def test_rutas_aceptan_el_payload_nuevo(monkeypatch, request) -> None:
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    request.addfinalizer(get_settings.cache_clear)
    headers = {"X-Internal-Token": TOKEN}

    res = client.post(
        "/pdf/balance-avion-xlsx",
        json={
            "matricula": "N4142R",
            "periodo_desde": "2026-09-01",
            "periodo_hasta": "2026-09-30",
            "vuelos": _vuelos_sep(con_llave=True),
            "totales": _totales(con_llave=True),
        },
        headers=headers,
    )
    assert res.status_code == 200, res.text
    ws = _maestra(res.content)
    assert ws.cell(row=_fila_de(ws, RENGLON_EXT), column=9).value == pytest.approx(5049.06)

    res = client.post(
        "/pdf/balance-general-xlsx",
        json={
            "periodo_desde": "2026-08-01",
            "periodo_hasta": "2026-09-30",
            "consolidado": {
                "matricula": "FLOTA",
                "vuelos": _vuelos_flota(con_llave=True),
                "totales": _totales(con_llave=True, extension=EXT_190 + EXT_192 + EXT_314),
            },
        },
        headers=headers,
    )
    assert res.status_code == 200, res.text
    ws = _maestra(res.content, "reporte horas FLOTA")
    assert ws.cell(row=_fila_de(ws, RENGLON_EXT), column=9).value == pytest.approx(9598.1)
    assert _nota_traslados(ws) == _NOTA_TRASLADOS_EXTENSION
