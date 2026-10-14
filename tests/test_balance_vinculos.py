"""Vínculos gasto → celda del balance (13-oct-2026, API 0.0.75, punto 5 del
lote del cliente: desde Conciliación, abrir el balance del avión o el general
EN LA CELDA DEL GASTO).

El API manda los ids de los gastos (`gasto_id` / `gasto_parte` en cada
renglón de las hojas de gastos y `op_vinculos` / `piloto_vinculos` /
`otros_vinculos` en cada fila de vuelo) y, SOLO en la petición del libro de la
vista en pantalla, `vinculos_gasto: true`. pyservices los transporta como
`vinculos` de la celda en el libro JSON y marca la raíz; el .xlsx no cambia ni
un byte. Aquí no se decide a qué libro ni a qué mes va un gasto: solo se
transporta lo que el API ya acreditó.

Fixtures: COPIAS de los dicts de `tests/test_balance_formulas.py` con los ids
inyectados ANTES de `model_validate` (ese archivo no cambia).
"""

from __future__ import annotations

import copy
import zipfile
from io import BytesIO

import pytest
from fastapi.testclient import TestClient

from app.config import get_settings
from app.main import app
from app.schemas.reportes import (
    BalanceAvionGastoFila,
    BalanceAvionRequest,
    BalanceAvionVuelo,
    BalanceGeneralRequest,
)
from app.services.balance_avion_xlsx import (
    libro_balance_avion,
    libro_balance_general,
    render_balance_avion_xlsx,
    render_balance_general_xlsx,
)
from tests import test_balance_formulas as t_formulas

# Ids con forma de uuid (el panel solo salta a un uuid).
G_OP = "11111111-1111-4111-8111-000000000001"
G_MIXTA = "11111111-1111-4111-8111-000000000002"  # factura repartida op + FBO
G_TUA = "11111111-1111-4111-8111-000000000003"
G_COMIDA = "11111111-1111-4111-8111-000000000004"
G_FBO = "11111111-1111-4111-8111-000000000005"
G_OP2 = "11111111-1111-4111-8111-000000000006"
G_USD = "11111111-1111-4111-8111-000000000007"  # USD sin T.C.
G_OP_DOS = "11111111-1111-4111-8111-000000000008"

CLAVE_401 = "#401 · Cliente Uno"
CLAVE_402 = "#402 · Cliente Dos"

# Fila #401: OPERACIONES 4,500.50 = 4,433.36 + 67.14 (operación de la
# factura repartida) · OTROS 350.75 = 87.00 (FBO de esa factura) + 263.75 ·
# PILOTO 1,200.00. La TUA va como 'nota' (no suma).
_VINCULOS_401 = {
    "op_vinculos": [
        {"id": G_OP, "rol": "suma", "monto_mxn": 4433.36},
        {"id": G_MIXTA, "rol": "suma", "monto_mxn": 67.14, "gasto_mxn": 154.14},
        {"id": G_TUA, "rol": "nota", "monto_mxn": None, "gasto_mxn": 250.0},
    ],
    "piloto_vinculos": [{"id": G_COMIDA, "rol": "suma", "monto_mxn": 1200.0}],
    "otros_vinculos": [
        {"id": G_MIXTA, "rol": "suma", "monto_mxn": 87.0, "gasto_mxn": 154.14},
        {"id": G_FBO, "rol": "suma", "monto_mxn": 263.75},
    ],
}
# Fila #402 de XA-TST: OPERACIONES 2,000 + un gasto USD sin ningún T.C.
_VINCULOS_402 = {
    "op_vinculos": [
        {"id": G_OP2, "rol": "suma", "monto_mxn": 2000.0},
        {"id": G_USD, "rol": "sin_tc", "monto_mxn": None},
    ],
    "piloto_vinculos": [],
    "otros_vinculos": [],
}


def _id_fila(prefijo: str, i: int) -> str:
    return f"22222222-2222-4222-8222-{prefijo}{i:08d}"


def _con_ids_libro(p: dict, prefijo: str) -> dict:
    """COPIA del libro de un avión con `gasto_id` en cada renglón de sus hojas
    de gastos y de combustible (el parcial del reparto manual con
    `gasto_parte: 'reparto'`) y los `*_vinculos` de las filas #401/#402."""
    p = copy.deepcopy(p)
    for i, hoja in enumerate(
        ("gastos_indirectos", "otros_gastos", "refacciones", "permisos", "combustible")
    ):
        for j, f in enumerate(p[hoja]["filas"]):
            f["gasto_id"] = _id_fila(f"{prefijo}{i}", j)
            if "reparto manual" in (f.get("detalle") or ""):
                f["gasto_parte"] = "reparto"
    for v in p["vuelos"]:
        if v["clave"] == CLAVE_401:
            v.update(copy.deepcopy(_VINCULOS_401))
        elif v["clave"] == CLAVE_402 and p["matricula"] == "XA-TST":
            v.update(copy.deepcopy(_VINCULOS_402))
        elif v["clave"] == CLAVE_402:
            v["op_vinculos"] = [{"id": G_OP_DOS, "rol": "suma", "monto_mxn": 1800.0}]
    return p


def _payload_tst_ids() -> dict:
    return _con_ids_libro(t_formulas._payload_tst(), "1")


def _general_ids(*, variante: str | None = None, bandera: bool = True) -> dict:
    """COPIA del general de `test_balance_formulas` con los ids: los libros de
    los aviones (el consolidado los hereda por el spread de sus filas, como el
    `hojaFlota` del API), 'otros gastos' de empresa (el último renglón es el
    REMANENTE de un reparto manual) y el detalle de salidas del inventario
    con su COSTO / VENTA AL AVIÓN."""
    tst = _con_ids_libro(t_formulas._payload_tst(), "1")
    dos = _con_ids_libro(t_formulas._payload_dos(), "2")
    payload = t_formulas._general_payload([tst, dos])
    for j, f in enumerate(payload["gastos_empresa"]["filas"]):
        f["gasto_id"] = _id_fila("30", j)
    payload["gastos_empresa"]["filas"][-1]["gasto_parte"] = "remanente"
    for f in payload["consolidado"]["refacciones"]["filas"]:
        f["costo_mxn"] = 960.0
        f["venta_mxn"] = f["monto_mxn"]
    if variante is not None:
        payload["variante"] = variante
    if bandera:
        payload["vinculos_gasto"] = True
    return payload


# ---------------------------------------------------------------------------
# Lectura del libro JSON
# ---------------------------------------------------------------------------


def _hoja(libro: dict, nombre: str) -> dict:
    return next(h for h in libro["hojas"] if h["nombre"] == nombre)


def _celda(hoja: dict, r: int, c: int) -> dict | None:
    return next((x for x in hoja["celdas"] if x["r"] == r and x["c"] == c), None)


def _col(hoja: dict, encabezado: str, *, fila: int = 2) -> int:
    """Columna por su ENCABEZADO leído del JSON (nunca con letra fija)."""
    return next(
        x["c"]
        for x in hoja["celdas"]
        if x["r"] == fila and isinstance(x["v"], str) and x["v"].replace("\n", " ") == encabezado
    )


def _filas_de_clave(hoja: dict, clave: str) -> list[int]:
    return sorted(x["r"] for x in hoja["celdas"] if x["c"] == 1 and x["v"] == clave)


def _con_vinculos(hoja: dict) -> list[dict]:
    return [x for x in hoja["celdas"] if x.get("vinculos")]


def _todos_los_vinculos(libro: dict) -> list[dict]:
    return [v for h in libro["hojas"] for x in h["celdas"] for v in x.get("vinculos", [])]


def _vinculo_vuelo(x: dict, *, clave: str, columna: str) -> dict:
    """Lo que se espera en la celda: el vínculo del API + tipo, fila y columna."""
    d = {"tipo": "gasto", "id": x["id"], "rol": x["rol"], "monto_mxn": x.get("monto_mxn")}
    if x.get("gasto_mxn") is not None:
        d["gasto_mxn"] = x["gasto_mxn"]
    d["fila"] = clave
    d["columna"] = columna
    return d


def _miembros(data: bytes) -> dict[str, bytes]:
    """Todos los miembros del .xlsx salvo docProps/core.xml (lleva la hora);
    copia del helper de test_balance_extension_horario."""
    z = zipfile.ZipFile(BytesIO(data))
    return {n: z.read(n) for n in sorted(z.namelist()) if n != "docProps/core.xml"}


# ---------------------------------------------------------------------------
# (a) Libro individual: OPERACIONES, PILOTO y OTROS de la hoja maestra
# ---------------------------------------------------------------------------


def test_a_individual_celdas_de_costos_de_la_fila_del_vuelo():
    payload = _payload_tst_ids()
    payload["vinculos_gasto"] = True
    libro = libro_balance_avion(BalanceAvionRequest.model_validate(payload))
    assert libro["vinculos_gasto"] is True
    hoja = _hoja(libro, "reporte horas XA-TST")
    (r401,) = _filas_de_clave(hoja, CLAVE_401)
    for encabezado, llave in (
        ("OPERACIONES", "op_vinculos"),
        ("PILOTO", "piloto_vinculos"),
        ("OTROS", "otros_vinculos"),
    ):
        celda = _celda(hoja, r401, _col(hoja, encabezado))
        esperado = [
            _vinculo_vuelo(x, clave=CLAVE_401, columna=encabezado) for x in _VINCULOS_401[llave]
        ]
        assert celda["vinculos"] == esperado, encabezado
        # La nota de desglose sigue ahí y lo que suma la celda cuadra con su
        # número (metadato del API: aquí solo se comprueba que viaja entero).
        suma = sum(v["monto_mxn"] for v in celda["vinculos"] if v["rol"] == "suma")
        assert suma == pytest.approx(celda["v"], abs=0.01), encabezado
    op = _celda(hoja, r401, _col(hoja, "OPERACIONES"))["vinculos"]
    # La TUA: rol 'nota', sin monto en la celda, con el gasto completo.
    assert op[2] == {
        "tipo": "gasto",
        "id": G_TUA,
        "rol": "nota",
        "monto_mxn": None,
        "gasto_mxn": 250.0,
        "fila": CLAVE_401,
        "columna": "OPERACIONES",
    }
    # `gasto_mxn` solo en la PARTE de la factura repartida (y en la nota).
    assert [v.get("gasto_mxn") for v in op] == [None, 154.14, 250.0]
    assert all("\n" not in v["columna"] for v in _todos_los_vinculos(libro))
    # #402: el gasto USD sin T.C. llega con 'sin_tc' y sin monto.
    (r402,) = _filas_de_clave(hoja, CLAVE_402)
    op402 = _celda(hoja, r402, _col(hoja, "OPERACIONES"))["vinculos"]
    assert op402[1] == {
        "tipo": "gasto",
        "id": G_USD,
        "rol": "sin_tc",
        "monto_mxn": None,
        "fila": CLAVE_402,
        "columna": "OPERACIONES",
    }
    # Ni la fila TOTALES, ni COSTO TOTAL, ni COMISIONES, ni el AFAC llevan.
    columnas = {(x["r"], x["c"]) for x in _con_vinculos(hoja)}
    cols_costos = {_col(hoja, e) for e in ("OPERACIONES", "PILOTO", "OTROS")}
    assert {c for _, c in columnas} <= cols_costos
    assert {r for r, _ in columnas} == {r401, r402}


# ---------------------------------------------------------------------------
# (b) Hoja «permisos» con `tc_por_fila`: la celda MONTO MXN es FÓRMULA
# ---------------------------------------------------------------------------


def test_b_permisos_con_tc_por_fila_la_celda_formula_lleva_el_vinculo():
    payload = t_formulas._payload_tst_permisos_tc(con_usd_sin_tc=True)
    ids = []
    for j, f in enumerate(payload["permisos"]["filas"]):
        f["gasto_id"] = _id_fila("40", j)
        ids.append((f["gasto_id"], f["monto_mxn"]))
    payload["vinculos_gasto"] = True
    libro = libro_balance_avion(BalanceAvionRequest.model_validate(payload))
    hoja = _hoja(libro, "permisos")
    for k, (gid, monto) in enumerate(ids):
        celda = _celda(hoja, 8 + k, 4)
        assert celda["vinculos"] == [
            {"tipo": "gasto", "id": gid, "rol": "fila", "monto_mxn": monto, "columna": "MONTO MXN"}
        ]
        if monto is not None:
            assert celda["f"].startswith("=ROUND(")
    assert len(_con_vinculos(hoja)) == len(ids)


# ---------------------------------------------------------------------------
# (c) «combustible» y «Gastos Indirectos» (fusionada) del libro individual
# ---------------------------------------------------------------------------


def test_c_combustible_gastos_indirectos_y_refacciones():
    payload = _payload_tst_ids()
    req = BalanceAvionRequest.model_validate({**payload, "vinculos_gasto": True})
    libro = libro_balance_avion(req)
    # Combustible: una celda MONTO MXN (columna D) por renglón, rol 'fila'.
    comb = _hoja(libro, "combustible")
    esperado = {f["gasto_id"]: f["monto_mxn"] for f in payload["combustible"]["filas"]}
    vistos = {}
    for x in _con_vinculos(comb):
        assert x["c"] == 4
        (v,) = x["vinculos"]
        assert v["rol"] == "fila" and v["columna"] == "MONTO MXN" and "parte" not in v
        vistos[v["id"]] = v["monto_mxn"]
    assert vistos == esperado
    # Gastos Indirectos = gastos_indirectos + otros_gastos (fusión): el
    # parcial del reparto manual lleva el id del PADRE y `parte: 'reparto'`.
    gi = _hoja(libro, "Gastos Indirectos")
    filas = [*payload["gastos_indirectos"]["filas"], *payload["otros_gastos"]["filas"]]
    por_id = {f["gasto_id"]: f for f in filas}
    vistos = {}
    for x in _con_vinculos(gi):
        assert x["c"] == 4
        (v,) = x["vinculos"]
        vistos[v["id"]] = v
    assert set(vistos) == set(por_id)
    parcial = next(f for f in filas if f.get("gasto_parte") == "reparto")
    assert vistos[parcial["gasto_id"]] == {
        "tipo": "gasto",
        "id": parcial["gasto_id"],
        "rol": "fila",
        "monto_mxn": 980.0,
        "columna": "MONTO MXN",
        "parte": "reparto",
    }
    assert sum("parte" in v for v in vistos.values()) == 1
    # Refacciones (hoja de gastos de siempre).
    (x,) = _con_vinculos(_hoja(libro, "refacciones"))
    assert x["vinculos"][0]["id"] == payload["refacciones"]["filas"][0]["gasto_id"]


# ---------------------------------------------------------------------------
# (d) General «mensual»
# ---------------------------------------------------------------------------


def test_d_general_mensual_flota_otros_gastos_repartidos_combustible_inventario():
    payload = _general_ids()
    libro = libro_balance_general(BalanceGeneralRequest.model_validate(payload))
    assert libro["vinculos_gasto"] is True
    # OPERACIONES de «reporte horas FLOTA» (el multi-avión #402 sale DOS
    # veces, una por avión, cada una con lo suyo).
    flota = _hoja(libro, "reporte horas FLOTA")
    c_op = _col(flota, "OPERACIONES")
    (r401,) = _filas_de_clave(flota, CLAVE_401)
    assert [v["id"] for v in _celda(flota, r401, c_op)["vinculos"]] == [G_OP, G_MIXTA, G_TUA]
    ids_402 = sorted(
        v["id"]
        for r in _filas_de_clave(flota, CLAVE_402)
        for v in _celda(flota, r, c_op)["vinculos"]
    )
    assert ids_402 == sorted([G_OP2, G_USD, G_OP_DOS])
    # 'otros gastos' (gastos de empresa): el remanente con `parte`.
    og = _hoja(libro, "otros gastos")
    filas_og = payload["gastos_empresa"]["filas"]
    vistos = {x["vinculos"][0]["id"]: x["vinculos"][0] for x in _con_vinculos(og)}
    assert set(vistos) == {f["gasto_id"] for f in filas_og}
    assert vistos[filas_og[-1]["gasto_id"]]["parte"] == "remanente"
    assert "parte" not in vistos[filas_og[0]["gasto_id"]]
    # 'repartidos a aviones': el parcial del consolidado con `parte`.
    rep = _con_vinculos(_hoja(libro, "repartidos a aviones"))
    assert [x["vinculos"][0].get("parte") for x in rep] == ["reparto"]
    # Combustible de la flota: los renglones de los DOS aviones.
    comb = _hoja(libro, "combustible")
    esperados = {f["gasto_id"] for f in payload["consolidado"]["combustible"]["filas"]}
    assert {x["vinculos"][0]["id"] for x in _con_vinculos(comb)} == esperados
    assert all(x["c"] == 4 for x in _con_vinculos(comb))
    # Inventario, bloque 2 (detalle de salidas): VENTA AL AVIÓN.
    inv = _hoja(libro, "inventario")
    (x,) = _con_vinculos(inv)
    salida = payload["consolidado"]["refacciones"]["filas"][0]
    assert x["c"] == 5
    assert x["vinculos"] == [
        {
            "tipo": "gasto",
            "id": salida["gasto_id"],
            "rol": "fila",
            "monto_mxn": salida["venta_mxn"],
            "columna": "VENTA AL AVIÓN",
        }
    ]
    assert _celda(inv, x["r"] - 1, 5)["v"] == "VENTA AL AVIÓN\nMXN"


# ---------------------------------------------------------------------------
# (e) Variante «general»: TOTAL PARA PROVEEDOR (PESOS) con la unión por id
# ---------------------------------------------------------------------------


def test_e_variante_general_total_para_proveedor_une_por_id():
    payload = _general_ids(variante="general")
    libro = libro_balance_general(BalanceGeneralRequest.model_validate(payload))
    flota = _hoja(libro, "reporte horas FLOTA")
    columna = "TOTAL PARA PROVEEDOR (PESOS)"
    c_tp = _col(flota, columna)
    (r401,) = _filas_de_clave(flota, CLAVE_401)
    vinculos = _celda(flota, r401, c_tp)["vinculos"]
    # Orden de primera aparición: operación, piloto, otros.
    assert [v["id"] for v in vinculos] == [G_OP, G_MIXTA, G_TUA, G_COMIDA, G_FBO]
    mixta = vinculos[1]
    # Operación 67.14 + FBO 87.00 de la MISMA factura ⇒ UN vínculo de 154.14.
    assert mixta == {
        "tipo": "gasto",
        "id": G_MIXTA,
        "rol": "suma",
        "monto_mxn": 154.14,
        "gasto_mxn": 154.14,
        "fila": CLAVE_401,
        "columna": columna,
    }
    assert vinculos[2]["rol"] == "nota" and vinculos[2]["monto_mxn"] is None
    # Σ de lo que suma == operación + piloto + otros de la fila (la provisión
    # AFAC que también entra al TOTAL PARA PROVEEDOR no es un gasto: sin
    # vínculo).
    suma = sum(v["monto_mxn"] for v in vinculos if v["rol"] == "suma")
    assert suma == pytest.approx(4500.5 + 1200.0 + 350.75, abs=0.01)
    # La variante general no tiene columnas OPERACIONES / PILOTO / OTROS: las
    # únicas celdas con vínculos de la hoja son las de esa columna.
    assert {x["c"] for x in _con_vinculos(flota)} == {c_tp}


# ---------------------------------------------------------------------------
# (f) INVARIANTE: el .xlsx no cambia ni un byte
# ---------------------------------------------------------------------------


def test_f_el_xlsx_es_byte_a_byte_el_mismo_con_y_sin_ids():
    con = BalanceAvionRequest.model_validate({**_payload_tst_ids(), "vinculos_gasto": True})
    sin = BalanceAvionRequest.model_validate(t_formulas._payload_tst())
    assert _miembros(render_balance_avion_xlsx(con)) == _miembros(render_balance_avion_xlsx(sin))
    for variante in (None, "general"):
        con_g = BalanceGeneralRequest.model_validate(_general_ids(variante=variante))
        sin_g = copy.deepcopy(_general_ids(variante=variante, bandera=False))
        _quitar_ids(sin_g)
        sin_g = BalanceGeneralRequest.model_validate(sin_g)
        assert _miembros(render_balance_general_xlsx(con_g)) == _miembros(
            render_balance_general_xlsx(sin_g)
        ), variante


def _quitar_ids(o):
    """Quita (recursivo) toda llave de vínculo de un payload."""
    if isinstance(o, dict):
        for k in ("gasto_id", "gasto_parte", "op_vinculos", "piloto_vinculos", "otros_vinculos"):
            o.pop(k, None)
        for v in o.values():
            _quitar_ids(v)
    elif isinstance(o, list):
        for v in o:
            _quitar_ids(v)


# ---------------------------------------------------------------------------
# (g) Sin ids ni bandera: el libro de siempre; con la bandera y sin ids, la
# raíz la trae y no hay ni un vínculo.
# ---------------------------------------------------------------------------


def test_g_sin_ids_ni_bandera_y_bandera_sin_ids():
    libro = libro_balance_avion(BalanceAvionRequest.model_validate(t_formulas._payload_tst()))
    assert "vinculos_gasto" not in libro
    assert _todos_los_vinculos(libro) == []
    gen = libro_balance_general(BalanceGeneralRequest.model_validate(t_formulas._general_payload()))
    assert "vinculos_gasto" not in gen
    assert _todos_los_vinculos(gen) == []
    marcado = libro_balance_avion(
        BalanceAvionRequest.model_validate({**t_formulas._payload_tst(), "vinculos_gasto": True})
    )
    assert list(marcado)[-1] == "vinculos_gasto" and marcado["vinculos_gasto"] is True
    assert _todos_los_vinculos(marcado) == []
    # Las celdas del libro marcado son las mismas que sin la bandera.
    assert marcado["hojas"] == libro["hojas"]
    # Con ids pero SIN la bandera (el .xlsx o un API que no la pide): las
    # celdas traen sus vínculos y la raíz no se marca.
    sin_bandera = libro_balance_avion(BalanceAvionRequest.model_validate(_payload_tst_ids()))
    assert "vinculos_gasto" not in sin_bandera
    assert _todos_los_vinculos(sin_bandera) != []


# ---------------------------------------------------------------------------
# (h) Validadores liberales: jamás un 422 por un metadato
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("valor", [None, "x", 7, {"id": "g1"}])
def test_h_vinculos_que_no_son_lista_quedan_vacios(valor):
    v = BalanceAvionVuelo.model_validate(
        {"op_vinculos": valor, "piloto_vinculos": valor, "otros_vinculos": valor}
    )
    assert v.op_vinculos == [] and v.piloto_vinculos == [] and v.otros_vinculos == []


def test_h_vinculos_basura_se_filtran_y_los_montos_no_numericos_son_none():
    v = BalanceAvionVuelo.model_validate(
        {
            "op_vinculos": [
                1,
                {},
                {"id": ""},
                {"id": "   "},
                {"id": 5},
                {"id": "g1", "monto_mxn": "abc", "rol": 7, "gasto_mxn": True},
            ],
            "otros_vinculos": [
                {"id": " g2 ", "monto_mxn": "12.5", "rol": " NOTA ", "gasto_mxn": float("nan")},
                {"id": "g3", "monto_mxn": float("inf"), "rol": "", "extra": "se ignora"},
            ],
        }
    )
    assert [x.model_dump() for x in v.op_vinculos] == [
        {"id": "g1", "rol": "suma", "monto_mxn": None, "gasto_mxn": None}
    ]
    assert [x.model_dump() for x in v.otros_vinculos] == [
        {"id": "g2", "rol": "nota", "monto_mxn": 12.5, "gasto_mxn": None},
        {"id": "g3", "rol": "suma", "monto_mxn": None, "gasto_mxn": None},
    ]


@pytest.mark.parametrize(
    ("gasto_id", "gasto_parte", "id_esperado", "parte_esperada"),
    [
        (" abc ", "reparto", "abc", "reparto"),
        ("abc", " REMANENTE ", "abc", "remanente"),
        ("abc", "otro", "abc", None),
        ("", "reparto", None, "reparto"),
        (7, 3, None, None),
        (None, None, None, None),
    ],
)
def test_h_gasto_id_y_parte_liberales(gasto_id, gasto_parte, id_esperado, parte_esperada):
    f = BalanceAvionGastoFila.model_validate({"gasto_id": gasto_id, "gasto_parte": gasto_parte})
    assert f.gasto_id == id_esperado and f.gasto_parte == parte_esperada


@pytest.mark.parametrize(
    ("valor", "esperado"), [(True, True), ("true", None), (1, None), (None, None)]
)
def test_h_bandera_vinculos_gasto_solo_con_true(valor, esperado):
    payload = {**t_formulas._payload_tst(), "vinculos_gasto": valor}
    assert BalanceAvionRequest.model_validate(payload).vinculos_gasto is esperado
    gen = {**t_formulas._general_payload(), "vinculos_gasto": valor}
    assert BalanceGeneralRequest.model_validate(gen).vinculos_gasto is esperado


# ---------------------------------------------------------------------------
# Por HTTP: las rutas del libro aceptan el payload del API 0.0.75 (también
# con vínculos basura: jamás un 422 por un metadato) y la .xlsx lo ignora.
# ---------------------------------------------------------------------------

TOKEN = "tok-vinculos"
client = TestClient(app)


def test_rutas_aceptan_los_vinculos_del_api_0075(monkeypatch, request):
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    request.addfinalizer(get_settings.cache_clear)
    headers = {"X-Internal-Token": TOKEN}
    individual = {**_payload_tst_ids(), "vinculos_gasto": True}
    # Un vuelo con vínculos basura y un renglón con `gasto_id` numérico.
    individual["vuelos"][2]["op_vinculos"] = "x"
    individual["vuelos"][3]["piloto_vinculos"] = [1, {"id": ""}, None]
    individual["permisos"]["filas"][0]["gasto_id"] = 7
    res = client.post("/pdf/balance-avion-libro", json=individual, headers=headers)
    assert res.status_code == 200, res.text
    libro = res.json()
    assert libro["vinculos_gasto"] is True
    assert G_MIXTA in {v["id"] for v in _todos_los_vinculos(libro)}
    res = client.post("/pdf/balance-avion-xlsx", json=individual, headers=headers)
    assert res.status_code == 200, res.text
    general = _general_ids()
    res = client.post("/pdf/balance-general-libro", json=general, headers=headers)
    assert res.status_code == 200, res.text
    assert res.json()["vinculos_gasto"] is True
    res = client.post("/pdf/balance-general-xlsx", json=general, headers=headers)
    assert res.status_code == 200, res.text
