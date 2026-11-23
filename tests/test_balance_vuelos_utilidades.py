"""VUELOS PAGADOS CON UTILIDADES en la hoja 'balance' (23-nov-2026, API
0.0.148).

Pedido del dueño: los socios a veces pagan un vuelo propio «descontándolo de
sus utilidades». El API lo registra como una ENTREGA al socio ligada a un
cobro del vuelo, y el balance por avión y el general lo muestran APARTE del
dinero entregado. Aquí no se calcula nada: el API manda —SOLO cuando el avión
tiene al menos un vuelo así en el periodo— por avión
`vuelos_utilidades_socios_mxn` (el DISCRIMINADOR),
`entregado_efectivo_socios_mxn` y el detalle `vuelos_utilidades`; por socio
`vuelos_utilidades_mxn`, `entregado_efectivo_mxn` y `entregado_efectivo_usd`.
Lo de siempre no cambia de significado (`entregado_mxn` = TODAS las entregas,
`por_entregar_mxn` = monto − entregado, ya neto).

Con el discriminador, y SOLO en la tabla de socios EN CASCADA de la hoja en
PESOS: las entregas pasan de tres a CUATRO columnas (G «(−) VUELOS PAGADOS CON
UTILIDADES (MXN)» · H ENTREGADO (MXN), ya solo el dinero · I POR ENTREGAR
(MXN) = ROUND(E − G − H, 2) · J ENTREGADO USD informativo), el pie suma las
cuatro con otro rótulo, hay una nota más y, tras las notas, el DETALLE vuelo
por vuelo con su TOTAL. Sin el discriminador (API ≤ 0.0.147, o un avión sin
vuelos así) y en los demás caminos (tabla en cascada en USD, tablas legadas)
el libro es BYTE A BYTE el de hoy.
"""

from __future__ import annotations

import copy
import re
from dataclasses import dataclass
from io import BytesIO

import pytest
from fastapi.testclient import TestClient
from openpyxl import load_workbook

from app.config import get_settings
from app.main import app
from app.schemas.reportes import (
    BalanceAvionBalanceBloque,
    BalanceAvionRequest,
    BalanceAvionSocio,
    BalanceGeneralRequest,
    BalanceVueloUtilidades,
)
from app.services import xlsx_formulas
from app.services.balance_avion_xlsx import (
    _ANCHO_COL_ENTREGAS,
    _ANCHOS_BALANCE_PESOS_EN_CASCADA,
    _COL_VUELOS_UTILIDADES,
    _FILA_RESERVA_CASCADA_PESOS,
    _FILA_TOTAL_ENTREGAS_CON_VUELOS,
    _FILA_TOTAL_VUELOS_UTILIDADES,
    _NOTA_BALANCE_PESOS_EN_CASCADA,
    _NOTA_ENTREGAS_SOCIOS,
    _NOTA_VUELOS_UTILIDADES,
    _TITULO_VUELOS_UTILIDADES,
    BRAND,
    MONEY,
    _texto_vuelo_utilidades,
    _vuelos_utilidades_en_tabla,
    libro_balance_avion,
    libro_balance_general,
    render_balance_avion_xlsx,
    render_balance_general_xlsx,
)
from tests import _pesos_payloads as pp
from tests import test_balance_en_pesos as tp
from tests import test_balance_entregas_socios as t_ent
from tests import test_balance_reserva_overhaul as t_res
from tests import test_balance_reserva_retenida as t_ret
from tests import test_libro_json as tlj
from tests._evaluador_formulas import verificar_libro

r2 = pp.r2

# Textos EXACTOS del contrato (los necesita el panel).
COL_VUELOS = "(−) VUELOS PAGADOS\nCON UTILIDADES (MXN)"
COLS_ENTREGAS = [
    COL_VUELOS,
    "ENTREGADO (MXN)",
    "POR ENTREGAR (MXN)",
    "ENTREGADO USD (informativo)",
]
FILA_TOTAL_CON_VUELOS = "Vuelos pagados con utilidades / Entregado a socios / Por entregar"
NOTA_VUELOS = (
    "Vuelos pagados con utilidades: el socio pagó su vuelo descontándolo de sus utilidades "
    "(no entró dinero). El vuelo cuenta como cobrado y aquí se le resta al socio; por eso no "
    "aparece en ENTREGADO."
)
TITULO_DETALLE = "VUELOS PAGADOS CON UTILIDADES"

LLAVES_SOCIO = ("vuelos_utilidades_mxn", "entregado_efectivo_mxn", "entregado_efectivo_usd")
LLAVES_BLOQUE = (
    "vuelos_utilidades_socios_mxn",
    "entregado_efectivo_socios_mxn",
    "vuelos_utilidades",
)
DISCRIMINADOR = "vuelos_utilidades_socios_mxn"

# El caso real (N621TX, octubre 2026): (nombre, %, es_empresa).
EMPRESA = "Aero Charter Cancun"
SOCIOS = [
    ("Hernan Garza", 51.15, False),
    (EMPRESA, 20.0, True),
    ("Carlos Moreno", 15.54, False),
    ("Macedo", 13.31, False),
]
# nombre → (vuelos pagados con utilidades MXN, dinero entregado MXN, USD
# informativo de ese dinero, nº de entregas). Macedo solo «recibió» su vuelo.
ENTREGAS = {
    "Hernan Garza": (0.0, 150000.0, 8152.17, 2),
    "Carlos Moreno": (32210.43, 94782.0, 5151.2, 3),
    "Macedo": (44174.3, 0.0, 0.0, 1),
}
VUELOS = [
    {
        "folio": 496,
        "fecha": "2026-10-29",
        "socio": "Carlos Moreno",
        "matricula": "N621TX",
        "cliente": "Carlos Moreno",
        "monto_mxn": 32210.43,
    },
    {
        "folio": 498,
        "fecha": "2026-10-28",
        "socio": "Macedo",
        "matricula": "N621TX",
        "cliente": "Macedo",
        "monto_mxn": 44174.3,
    },
]
TEXTOS_VUELOS = [
    "#496 · 29/10/2026 · Carlos Moreno · N621TX",
    "#498 · 28/10/2026 · Macedo · N621TX",
]
TOTAL_VUELOS = 76384.73
# T.C. con que el payload de prueba pasa los vuelos al USD informativo.
TC_INFO = 18.4

# Para el GENERAL (los socios de `_pesos_payloads`, que cuadran con el bloque
# VUELATOUR): por matrícula, las entregas y el detalle. El segundo vuelo lo
# voló OTRO avión y lo pagó un cliente que no es el socio; al tercero le
# faltan piezas.
GENERAL = {
    "XA-TST": (
        {"Socio A": (12000.0, 20000.0, 1157.41, 3), "Socio B": (5000.5, 3000.0, 173.61, 2)},
        [
            {
                "folio": 512,
                "fecha": "2026-09-12",
                "socio": "Socio A",
                "matricula": "XA-TST",
                "cliente": "socio a",
                "monto_mxn": 12000.0,
            },
            {
                "folio": 515,
                "fecha": "2026-09-20",
                "socio": "Socio B",
                "matricula": "XB-DOS",
                "cliente": "Familia B",
                "monto_mxn": 5000.5,
            },
        ],
    ),
    "XB-DOS": (
        {"Aero Charter Cancún S.A. de C.V.": (1500.0, 10000.0, 578.7, 2)},
        [
            {
                "folio": 520,
                "fecha": None,
                "socio": "Aero Charter Cancún S.A. de C.V.",
                "matricula": None,
                "cliente": None,
                "monto_mxn": 1500.0,
            }
        ],
    ),
}
TEXTOS_GENERAL = {
    "XA-TST": [
        "#512 · 12/09/2026 · Socio A · XA-TST",
        "#515 · 20/09/2026 · Socio B · XB-DOS · Familia B",
    ],
    "XB-DOS": ["#520 · Aero Charter Cancún S.A. de C.V."],
}


# ---------------------------------------------------------------------------
# Payloads: el libro EN CASCADA de `test_balance_reserva_retenida` (API
# 0.0.130) con las entregas (invariante 100) y las llaves del 0.0.148, espejo
# de la aritmética del API: entregado = vuelos + dinero; por entregar = monto
# − entregado; Σ por avión.
# ---------------------------------------------------------------------------


def con_vuelos(
    libro: dict,
    *,
    socios: list[tuple[str, float, bool]] | None = None,
    entregas: dict | None = None,
    vuelos: list[dict] | None = None,
) -> dict:
    """El libro de un avión (ya en cascada) con sus entregas partidas en
    vuelos pagados con utilidades y dinero. `socios`: cambia los socios del
    bloque por estos y los recalcula como el API 0.0.130."""
    p = copy.deepcopy(libro)
    b = p["balance"]
    if socios is not None:
        b["socios"] = [
            {"nombre": nombre, "porcentaje": pct, "es_empresa": es_empresa}
            for nombre, pct, es_empresa in socios
        ]
        t_ret._socios_0130(b)
    tabla = ENTREGAS if entregas is None else entregas
    for s in b["socios"]:
        en_vuelos, dinero, dinero_usd, n = tabla.get(s["nombre"], (0.0, 0.0, 0.0, 0))
        s["vuelos_utilidades_mxn"] = en_vuelos
        s["entregado_efectivo_mxn"] = dinero
        s["entregado_efectivo_usd"] = dinero_usd
        s["entregado_mxn"] = r2(en_vuelos + dinero)
        s["entregado_usd"] = r2(dinero_usd + en_vuelos / TC_INFO)
        s["entregas"] = n
        monto = s.get("monto_mxn")
        s["por_entregar_mxn"] = None if monto is None else r2(monto - s["entregado_mxn"])
    b["entregado_socios_mxn"] = r2(sum(s["entregado_mxn"] for s in b["socios"]))
    pendientes = [s["por_entregar_mxn"] for s in b["socios"]]
    b["por_entregar_socios_mxn"] = (
        None if any(x is None for x in pendientes) else r2(sum(pendientes))
    )
    b["vuelos_utilidades_socios_mxn"] = r2(sum(s["vuelos_utilidades_mxn"] for s in b["socios"]))
    b["entregado_efectivo_socios_mxn"] = r2(
        b["entregado_socios_mxn"] - b["vuelos_utilidades_socios_mxn"]
    )
    b["vuelos_utilidades"] = copy.deepcopy(VUELOS if vuelos is None else vuelos)
    return p


def sin_llaves(libro: dict) -> dict:
    """El MISMO payload como lo manda un API ≤ 0.0.147: sin ninguna de las
    llaves nuevas (las entregas de siempre se quedan)."""
    p = copy.deepcopy(libro)
    for b in t_res._bloques(p):
        for k in LLAVES_BLOQUE:
            b.pop(k, None)
        for s in b["socios"]:
            for k in LLAVES_SOCIO:
                s.pop(k, None)
    return p


def individual(*, bandera: bool | None = True) -> dict:
    """El caso real: el individual en cascada con los cuatro socios."""
    return con_vuelos(t_ret.individual(bandera=bandera), socios=SOCIOS)


def general(variante: str = "mensual", *, bandera: bool | None = True) -> dict:
    """El general en cascada con vuelos pagados con utilidades en los DOS
    aviones (los socios de siempre: el bloque VUELATOUR sigue cuadrando)."""
    p = t_ret.general(variante, bandera=bandera)
    p["aviones"] = [
        con_vuelos(a, entregas=GENERAL[a["matricula"]][0], vuelos=GENERAL[a["matricula"]][1])
        for a in p["aviones"]
    ]
    return p


@pytest.fixture
def registros(monkeypatch) -> list[xlsx_formulas.RegistroFormulas]:
    """Registro de fórmulas de cada libro que se cierra (cuáles volvieron a
    valor): con las llaves ninguna debe degradarse."""
    capturados: list[xlsx_formulas.RegistroFormulas] = []
    original = xlsx_formulas.finalizar

    def espia(wb):
        cache = original(wb)
        capturados.append(xlsx_formulas.registro(wb))
        return cache

    monkeypatch.setattr(xlsx_formulas, "finalizar", espia)
    return capturados


# ---------------------------------------------------------------------------
# Helpers de lectura del libro.
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class _Filas:
    """Dónde cae cada fila de la tabla de socios con los vuelos pagados con
    utilidades, a partir de la fila de UTILIDAD COBRADA de su cascada."""

    titulo: int
    enc: int
    ini: int
    fin: int
    total: int
    pie: int
    nota_entregas: int
    nota_vuelos: int
    nota_socios: int
    det_blanco: int
    det_titulo: int
    det_ini: int
    det_fin: int
    det_total: int

    @property
    def siguiente(self) -> int:
        """La primera fila libre después de todo (ahí empieza el bloque de
        la reserva retenida, con su fila en blanco)."""
        return self.det_total + 1


def _filas(cob: int, n_socios: int, n_vuelos: int) -> _Filas:
    ini = cob + 4
    fin = ini + n_socios - 1
    pie = fin + 2
    return _Filas(
        titulo=cob + 2,
        enc=cob + 3,
        ini=ini,
        fin=fin,
        total=fin + 1,
        pie=pie,
        nota_entregas=pie + 1,
        nota_vuelos=pie + 2,
        nota_socios=pie + 3,
        det_blanco=pie + 4,
        det_titulo=pie + 5,
        det_ini=pie + 6,
        det_fin=pie + 5 + n_vuelos,
        det_total=pie + 6 + n_vuelos,
    )


def _merges(ws) -> set[str]:
    return {str(m) for m in ws.merged_cells.ranges}


def _tabla_con_vuelos(ws, wsv, cob: int, b: dict, textos: list[str]) -> _Filas:
    """La tabla de socios EN CASCADA en pesos del bloque cuya UTILIDAD COBRADA
    está en `cob`, con las cuatro columnas G..J, el pie, las notas y el
    detalle. Lo de siempre (A..F y la fila TOTAL) lo revisa el helper del
    0.0.130. Devuelve dónde quedó cada fila."""
    socios, vuelos = b["socios"], b["vuelos_utilidades"]
    f = _filas(cob, len(socios), len(vuelos))
    res = max(r for r in tp._filas_de(ws, _FILA_RESERVA_CASCADA_PESOS) if r < cob)
    assert t_ret._tabla(ws, wsv, cob, b, pesos=True, res=res) == f.pie
    merges = _merges(ws)
    rojos = t_ent._rangos_rojo(ws)
    # Encabezados: G..I principales (blanco sobre azul), J informativo, K nada.
    assert [ws.cell(row=f.enc, column=c).value for c in range(1, 12)] == [
        *t_ret.COLS_MXN,
        *COLS_ENTREGAS,
        None,
    ]
    assert f"A{f.titulo}:J{f.titulo}" in merges
    for c in (7, 8, 9):
        h = ws.cell(row=f.enc, column=c)
        assert h.font.b and h.font.color.rgb.endswith("FFFFFF")
        assert h.fill.fgColor.rgb.endswith(BRAND) and h.alignment.wrap_text
    info = ws.cell(row=f.enc, column=10)
    assert info.font.i and info.font.size == 9
    # Renglones: G y H valores del API, I la resta visible, J el USD del dinero.
    for i, s in enumerate(socios):
        r = f.ini + i
        assert ws.cell(row=r, column=1).value == s["nombre"]
        g, h, pe, usd = (ws.cell(row=r, column=c) for c in (7, 8, 9, 10))
        assert g.value == s["vuelos_utilidades_mxn"] and g.number_format == MONEY
        assert h.value == s["entregado_efectivo_mxn"] and h.number_format == MONEY
        assert not g.font.b and not h.font.b
        assert {g.coordinate, h.coordinate} <= rojos
        assert pe.value == f"=ROUND(E{r}-G{r}-H{r},2)" and pe.font.b
        assert wsv.cell(row=r, column=9).value == s["por_entregar_mxn"]
        assert pe.coordinate in rojos
        assert usd.value == s["entregado_efectivo_usd"]
        assert usd.font.i and usd.font.size == 9
        assert ws.cell(row=r, column=11).value is None
        # El número del API: monto − TODAS las entregas (vuelos + dinero).
        assert s["por_entregar_mxn"] == r2(s["monto_mxn"] - s["entregado_mxn"])
        assert s["entregado_mxn"] == r2(s["vuelos_utilidades_mxn"] + s["entregado_efectivo_mxn"])
    # Pie: el rótulo nuevo y las CUATRO sumas con los Σ del API.
    rotulo = ws.cell(row=f.pie, column=1)
    assert rotulo.value == FILA_TOTAL_CON_VUELOS and rotulo.font.b
    assert f"A{f.pie}:F{f.pie}" in merges
    sumas = {
        7: b["vuelos_utilidades_socios_mxn"],
        8: b["entregado_efectivo_socios_mxn"],
        9: b["por_entregar_socios_mxn"],
        10: r2(sum(s["entregado_efectivo_usd"] for s in socios)),
    }
    for col, esperado in sumas.items():
        letra = "GHIJ"[col - 7]
        celda = ws.cell(row=f.pie, column=col)
        assert celda.value == f"=ROUND(SUM({letra}{f.ini}:{letra}{f.fin}),2)", col
        assert wsv.cell(row=f.pie, column=col).value == esperado, col
    assert all(ws.cell(row=f.pie, column=c).font.b for c in (7, 8, 9))
    assert {f"G{f.pie}", f"H{f.pie}", f"I{f.pie}"} <= rojos
    assert r2(sumas[7] + sumas[8]) == b["entregado_socios_mxn"]
    # Notas: la de las entregas, la nueva debajo y la de la tabla.
    assert ws.cell(row=f.nota_entregas, column=1).value == _NOTA_ENTREGAS_SOCIOS
    nota = ws.cell(row=f.nota_vuelos, column=1)
    assert nota.value == NOTA_VUELOS and nota.font.i and nota.alignment.wrap_text
    assert ws.cell(row=f.nota_socios, column=1).value == t_ret.NOTA_SOCIOS
    for r in (f.nota_entregas, f.nota_vuelos, f.nota_socios):
        assert f"A{r}:J{r}" in merges
    # Detalle: fila en blanco, título, una fila por vuelo y el TOTAL.
    assert all(ws.cell(row=f.det_blanco, column=c).value is None for c in range(1, 12))
    titulo = ws.cell(row=f.det_titulo, column=1)
    assert titulo.value == TITULO_DETALLE and titulo.font.b
    assert f"A{f.det_titulo}:J{f.det_titulo}" in merges
    assert len(textos) == len(vuelos)
    for i, (v, texto) in enumerate(zip(vuelos, textos, strict=True)):
        r = f.det_ini + i
        assert ws.cell(row=r, column=1).value == texto
        assert f"A{r}:F{r}" in merges
        monto = ws.cell(row=r, column=7)
        assert monto.value == v["monto_mxn"] and monto.number_format == MONEY
        assert not monto.font.b
        assert all(ws.cell(row=r, column=c).value is None for c in (8, 9, 10))
    total = ws.cell(row=f.det_total, column=1)
    assert total.value == _FILA_TOTAL_VUELOS_UTILIDADES == "TOTAL" and total.font.b
    assert f"A{f.det_total}:F{f.det_total}" in merges
    suma = ws.cell(row=f.det_total, column=7)
    assert suma.value == f"=ROUND(SUM(G{f.det_ini}:G{f.det_fin}),2)" and suma.font.b
    assert wsv.cell(row=f.det_total, column=7).value == r2(sum(v["monto_mxn"] for v in vuelos))
    # El detalle suma lo mismo que la columna de la tabla.
    assert wsv.cell(row=f.det_total, column=7).value == wsv.cell(row=f.pie, column=7).value
    return f


def _celdas_de_merges(ws) -> list[tuple[int, int]]:
    return [
        (r, c)
        for m in ws.merged_cells.ranges
        for r in range(m.min_row, m.max_row + 1)
        for c in range(m.min_col, m.max_col + 1)
    ]


# ---------------------------------------------------------------------------
# 1) Esquema aditivo y liberal; discriminador; textos.
# ---------------------------------------------------------------------------


def test_esquema_aditivo_con_defaults_y_discriminador() -> None:
    s = BalanceAvionSocio()
    assert all(getattr(s, k) is None for k in LLAVES_SOCIO)
    b = BalanceAvionBalanceBloque()
    assert b.vuelos_utilidades_socios_mxn is None and b.entregado_efectivo_socios_mxn is None
    assert b.vuelos_utilidades == []
    assert not _vuelos_utilidades_en_tabla(b, pesos=True)
    # Solo la tabla EN CASCADA en PESOS, con las entregas y con el
    # discriminador (0.0 también cuenta: es un número).
    base = {
        "reserva_overhaul_usd": 1.0,
        "reserva_overhaul_en_cascada": True,
        "entregado_socios_mxn": 0,
        DISCRIMINADOR: 0,
    }
    completo = BalanceAvionBalanceBloque.model_validate(base)
    assert _vuelos_utilidades_en_tabla(completo, pesos=True)
    assert not _vuelos_utilidades_en_tabla(completo, pesos=False)
    for falta in ("reserva_overhaul_en_cascada", "entregado_socios_mxn", DISCRIMINADOR):
        bloque = BalanceAvionBalanceBloque.model_validate({**base, falta: None})
        assert not _vuelos_utilidades_en_tabla(bloque, pesos=True), falta


def test_esquema_liberal_jamas_un_422() -> None:
    # Números: texto, booleano o NaN ⇒ None (y el discriminador se apaga).
    for basura in ("76384.73", True, float("nan"), [], {}):
        b = BalanceAvionBalanceBloque.model_validate(
            {DISCRIMINADOR: basura, "entregado_efectivo_socios_mxn": basura}
        )
        assert b.vuelos_utilidades_socios_mxn is None, basura
        assert b.entregado_efectivo_socios_mxn is None, basura
        s = BalanceAvionSocio.model_validate(dict.fromkeys(LLAVES_SOCIO, basura))
        assert all(getattr(s, k) is None for k in LLAVES_SOCIO), basura
    s = BalanceAvionSocio.model_validate(
        {"vuelos_utilidades_mxn": 44174.3, "entregado_efectivo_mxn": 0, "entregado_efectivo_usd": 0}
    )
    assert (s.vuelos_utilidades_mxn, s.entregado_efectivo_mxn, s.entregado_efectivo_usd) == (
        44174.3,
        0.0,
        0.0,
    )
    # El detalle: lo que no sea lista ⇒ [].
    for basura in (None, "no soy una lista", 7, True, {"folio": 1, "monto_mxn": 5}):
        b = BalanceAvionBalanceBloque.model_validate({"vuelos_utilidades": basura})
        assert b.vuelos_utilidades == [], basura
    # Filas rotas (no objeto, sin monto, monto que no es número) se saltan; un
    # folio que no es entero o un texto que no es texto ⇒ esa pieza en None.
    b = BalanceAvionBalanceBloque.model_validate(
        {
            "vuelos_utilidades": [
                {"folio": 498, "fecha": "2026-10-28", "socio": " Macedo ", "monto_mxn": 100},
                "texto",
                None,
                5,
                ["lista"],
                {"folio": 1},
                {"folio": 2, "monto_mxn": "cien"},
                {"folio": 3, "monto_mxn": None},
                {"folio": 4, "monto_mxn": True},
                {"folio": 5, "monto_mxn": float("nan")},
                {
                    "folio": "abc",
                    "fecha": 20261028,
                    "socio": 9,
                    "matricula": "  ",
                    "cliente": ["x"],
                    "monto_mxn": 50.5,
                    "llave_de_mas": {"a": 1},
                },
                {"folio": "499", "monto_mxn": 1},
                {"folio": 2.5, "monto_mxn": 2},
                {"folio": 0, "monto_mxn": 3},
            ]
        }
    )
    filas = [
        (v.folio, v.fecha, v.socio, v.matricula, v.cliente, v.monto_mxn)
        for v in b.vuelos_utilidades
    ]
    assert filas == [
        (498, "2026-10-28", "Macedo", None, None, 100.0),
        (None, None, "", None, None, 50.5),
        (499, None, "", None, None, 1.0),
        (None, None, "", None, None, 2.0),
        (None, None, "", None, None, 3.0),
    ]
    assert all(isinstance(v, BalanceVueloUtilidades) for v in b.vuelos_utilidades)


def test_textos_exactos_del_contrato() -> None:
    assert _COL_VUELOS_UTILIDADES == COL_VUELOS
    assert _FILA_TOTAL_ENTREGAS_CON_VUELOS == FILA_TOTAL_CON_VUELOS
    assert _NOTA_VUELOS_UTILIDADES == NOTA_VUELOS
    assert _TITULO_VUELOS_UTILIDADES == TITULO_DETALLE
    assert _ANCHO_COL_ENTREGAS == 20


@pytest.mark.parametrize(
    ("fila", "texto"),
    [
        (VUELOS[1], "#498 · 28/10/2026 · Macedo · N621TX"),
        # El cliente solo se añade si es OTRO que el socio (sin mayúsculas ni
        # espacios de más).
        ({**VUELOS[1], "cliente": " macedo "}, "#498 · 28/10/2026 · Macedo · N621TX"),
        (
            {**VUELOS[1], "cliente": "Familia Macedo"},
            "#498 · 28/10/2026 · Macedo · N621TX · Familia Macedo",
        ),
        # Las piezas que no vienen se omiten.
        ({**VUELOS[1], "fecha": None, "matricula": None}, "#498 · Macedo"),
        ({**VUELOS[1], "folio": None, "cliente": None}, "28/10/2026 · Macedo · N621TX"),
        ({"monto_mxn": 1}, "Vuelo sin datos"),
        ({"socio": "", "cliente": "Ana", "monto_mxn": 1}, "Ana"),
    ],
)
def test_texto_del_renglon_del_detalle(fila, texto) -> None:
    assert _texto_vuelo_utilidades(BalanceVueloUtilidades.model_validate(fila)) == texto


# ---------------------------------------------------------------------------
# 2) (a) Libro INDIVIDUAL en pesos, el caso real.
# ---------------------------------------------------------------------------


def test_individual_pesos_caso_real(registros) -> None:
    p = individual()
    b = p["balance"]
    macedo, carlos = b["socios"][3], b["socios"][2]
    assert (macedo["nombre"], carlos["nombre"]) == ("Macedo", "Carlos Moreno")
    assert macedo["vuelos_utilidades_mxn"] == macedo["entregado_mxn"] == 44174.3
    assert macedo["entregado_efectivo_mxn"] == 0
    assert macedo["por_entregar_mxn"] == r2(macedo["monto_mxn"] - 44174.3)
    assert (carlos["vuelos_utilidades_mxn"], carlos["entregado_efectivo_mxn"]) == (
        32210.43,
        94782.0,
    )
    assert b["vuelos_utilidades_socios_mxn"] == TOTAL_VUELOS
    data = render_balance_avion_xlsx(BalanceAvionRequest.model_validate(p))
    assert tp._degradadas_balance(registros[-1]) == []
    verificar_libro(data)
    wb, wv = tp._libros(data)
    ws, wsv = wb["balance"], wv["balance"]
    cob = tp._fila_de(ws, "UTILIDAD COBRADA (MXN)")
    f = _tabla_con_vuelos(ws, wsv, cob, b, TEXTOS_VUELOS)
    # Macedo: nada de dinero, todo su vuelo; le falta (o sobra) el resto.
    r = f.fin
    assert ws.cell(row=r, column=1).value == "Macedo"
    assert [wsv.cell(row=r, column=c).value for c in (7, 8, 9, 10)] == [
        44174.3,
        0,
        macedo["por_entregar_mxn"],
        0,
    ]
    assert wsv.cell(row=f.pie, column=7).value == TOTAL_VUELOS
    assert wsv.cell(row=f.det_total, column=7).value == TOTAL_VUELOS
    # Los socios sin vuelos llevan 0.00 en la columna, no una celda vacía.
    assert [ws.cell(row=f.ini + i, column=7).value for i in (0, 1)] == [0, 0]
    # El rótulo de siempre del pie ya no está; el nuevo, una vez.
    assert not tp._filas_de(ws, t_ent.FILA_TOTAL)
    assert tp._filas_de(ws, FILA_TOTAL_CON_VUELOS) == [f.pie]
    # Lo de DESPUÉS sigue en su sitio: fila en blanco, el bloque de la reserva
    # retenida y, al final, el pie de la hoja.
    t_ret._bloque_retenida(ws, f.siguiente, b, pesos=True, como=t_res.COMO_PESOS)
    assert tp._fila_de(ws, t_ret.TITULO_EN_CASCADA) == f.siguiente + 1
    assert t_ret._textos_a(ws)[-1] == _NOTA_BALANCE_PESOS_EN_CASCADA
    assert ws.max_row == f.siguiente + 4 + 2
    # Anchos: los de la tabla en cascada + CUATRO de 20; K, el de siempre.
    anchos = [ws.column_dimensions[c].width for c in "ABCDEFGHIJ"]
    assert anchos == [*_ANCHOS_BALANCE_PESOS_EN_CASCADA, *([_ANCHO_COL_ENTREGAS] * 4)]
    assert ws.column_dimensions["K"].width != _ANCHO_COL_ENTREGAS
    celdas = _celdas_de_merges(ws)
    assert len(celdas) == len(set(celdas))
    t_ret._ninguna_formula_cita_texto(ws)


def test_individual_pesos_respaldo_si_el_api_no_manda_el_dinero(registros) -> None:
    """Sin `entregado_efectivo_*` (no debería pasar), el respaldo del
    contrato: ENTREGADO = `entregado_mxn` − `vuelos_utilidades_mxn`, su Σ =
    `entregado_socios_mxn` − `vuelos_utilidades_socios_mxn` y el USD
    informativo, el `entregado_usd` de siempre. Un socio sin
    `vuelos_utilidades_mxn` lleva 0.00."""
    p = individual()
    b = p["balance"]
    del b["entregado_efectivo_socios_mxn"]
    for s in b["socios"]:
        del s["entregado_efectivo_mxn"], s["entregado_efectivo_usd"]
    del b["socios"][0]["vuelos_utilidades_mxn"]
    data = render_balance_avion_xlsx(BalanceAvionRequest.model_validate(p))
    assert tp._degradadas_balance(registros[-1]) == []
    verificar_libro(data)
    wb, wv = tp._libros(data)
    ws, wsv = wb["balance"], wv["balance"]
    f = _filas(tp._fila_de(ws, "UTILIDAD COBRADA (MXN)"), len(b["socios"]), len(VUELOS))
    completo = individual()["balance"]
    for i, (s, s_api) in enumerate(zip(b["socios"], completo["socios"], strict=True)):
        r = f.ini + i
        assert ws.cell(row=r, column=7).value == s_api["vuelos_utilidades_mxn"]
        assert ws.cell(row=r, column=8).value == s_api["entregado_efectivo_mxn"]
        assert ws.cell(row=r, column=9).value == f"=ROUND(E{r}-G{r}-H{r},2)"
        assert wsv.cell(row=r, column=9).value == s["por_entregar_mxn"]
        assert ws.cell(row=r, column=10).value == s["entregado_usd"]
    assert wsv.cell(row=f.pie, column=7).value == TOTAL_VUELOS
    assert wsv.cell(row=f.pie, column=8).value == completo["entregado_efectivo_socios_mxn"]
    assert wsv.cell(row=f.pie, column=10).value == r2(sum(s["entregado_usd"] for s in b["socios"]))


def test_individual_pesos_sin_filas_de_detalle_no_pinta_el_bloque(registros) -> None:
    """Con el discriminador y la lista vacía (o rota): columna, pie y nota
    sí; el bloque de detalle no, y la reserva retenida sube a su sitio."""
    for detalle in ([], None, "roto", [{"folio": 1}]):
        p = individual()
        b = p["balance"]
        b["vuelos_utilidades"] = detalle
        data = render_balance_avion_xlsx(BalanceAvionRequest.model_validate(p))
        assert tp._degradadas_balance(registros[-1]) == []
        verificar_libro(data)
        wb, wv = tp._libros(data)
        ws, wsv = wb["balance"], wv["balance"]
        f = _filas(tp._fila_de(ws, "UTILIDAD COBRADA (MXN)"), len(b["socios"]), 0)
        assert ws.cell(row=f.enc, column=7).value == COL_VUELOS
        assert wsv.cell(row=f.pie, column=7).value == TOTAL_VUELOS
        assert ws.cell(row=f.nota_vuelos, column=1).value == NOTA_VUELOS
        assert ws.cell(row=f.nota_socios, column=1).value == t_ret.NOTA_SOCIOS
        assert not tp._filas_de(ws, TITULO_DETALLE)
        assert len(tp._filas_de(ws, "TOTAL")) == 1
        t_ret._bloque_retenida(ws, f.nota_socios + 1, b, pesos=True, como=t_res.COMO_PESOS)


def test_detalle_con_un_nombre_que_empieza_con_igual_se_queda_en_texto() -> None:
    """Los nombres los teclea la oficina: un renglón del detalle que empezara
    con «=» (sin folio ni fecha) es TEXTO, nunca una fórmula."""
    p = individual()
    p["balance"]["vuelos_utilidades"] = [
        {"socio": "=Macedo", "matricula": "N621TX", "monto_mxn": TOTAL_VUELOS}
    ]
    data = render_balance_avion_xlsx(BalanceAvionRequest.model_validate(p))
    verificar_libro(data)
    ws = load_workbook(BytesIO(data))["balance"]
    (titulo,) = tp._filas_de(ws, TITULO_DETALLE)
    celda = ws.cell(row=titulo + 1, column=1)
    assert celda.value == "=Macedo · N621TX" and celda.data_type == "s"


def test_individual_pesos_sin_pesos_raya_por_entregar() -> None:
    """Sin T.C. oficial del cierre no hay MONTO en pesos: POR ENTREGAR va
    «—» (sin fórmula sobre texto) y los vuelos y el dinero se pintan igual."""
    p = con_vuelos(t_ret.individual(sin_tc=True), socios=SOCIOS)
    b = p["balance"]
    assert b["por_entregar_socios_mxn"] is None
    data = render_balance_avion_xlsx(BalanceAvionRequest.model_validate(p))
    verificar_libro(data)
    wb, wv = tp._libros(data)
    ws, wsv = wb["balance"], wv["balance"]
    f = _filas(tp._fila_de(ws, "UTILIDAD COBRADA (MXN)"), len(b["socios"]), len(VUELOS))
    for i, s in enumerate(b["socios"]):
        r = f.ini + i
        assert ws.cell(row=r, column=7).value == s["vuelos_utilidades_mxn"]
        assert ws.cell(row=r, column=8).value == s["entregado_efectivo_mxn"]
        assert ws.cell(row=r, column=9).value == "—"
    assert wsv.cell(row=f.pie, column=7).value == TOTAL_VUELOS
    assert ws.cell(row=f.pie, column=9).value == "—"
    assert ws.cell(row=f.det_titulo, column=1).value == TITULO_DETALLE
    t_ret._ninguna_formula_cita_texto(ws)


# ---------------------------------------------------------------------------
# 3) (b) Sin el discriminador: BYTE A BYTE el libro de hoy; y los caminos que
#    no son la tabla en cascada en pesos ignoran las llaves aunque vengan.
# ---------------------------------------------------------------------------


def _payload(libro: str, bandera: bool | None = True) -> dict:
    if libro == "individual":
        return individual(bandera=bandera)
    return general(libro.removeprefix("general-"), bandera=bandera)


@pytest.mark.parametrize("forma", ["ausente", "null", "texto", "booleano", "nan"])
@pytest.mark.parametrize("libro", ["individual", "general-mensual", "general-general"])
def test_sin_el_discriminador_el_libro_de_hoy_byte_a_byte(libro, forma) -> None:
    """El discriminador ausente, en null o con algo que no es un número —con
    TODAS las demás llaves nuevas presentes— ⇒ el .xlsx es miembro a miembro
    el del payload sin ninguna llave nueva, y la vista en pantalla también."""
    p = _payload(libro)
    base = sin_llaves(p)
    for b in t_res._bloques(p):
        if forma == "ausente":
            del b[DISCRIMINADOR]
        else:
            b[DISCRIMINADOR] = {
                "null": None,
                "texto": "76384.73",
                "booleano": True,
                "nan": float("nan"),
            }[forma]
    data_base, vista_base = tp._render(libro, base)
    data, vista = tp._render(libro, p)
    assert t_ret._miembros(data) == t_ret._miembros(data_base)
    assert vista == vista_base
    # Y ese libro es el de las entregas de siempre: tres columnas, su rótulo.
    ws = load_workbook(BytesIO(data))["balance"]
    assert len(tp._filas_de(ws, t_ent.FILA_TOTAL)) == len(t_res._bloques(p))
    assert not tp._filas_de(ws, FILA_TOTAL_CON_VUELOS) and not tp._filas_de(ws, TITULO_DETALLE)


def test_sin_las_entregas_el_discriminador_solo_no_pinta_nada() -> None:
    """Los vuelos pagados con utilidades SON entregas: sin
    `entregado_socios_mxn` no hay columnas de entregas ni, por tanto, la
    nueva."""
    base = t_ret.individual()
    p = con_vuelos(base)
    for k in ("entregado_socios_mxn", "por_entregar_socios_mxn"):
        del p["balance"][k]
    data_base, vista_base = tp._render("individual", base)
    data, vista = tp._render("individual", p)
    assert t_ret._miembros(data) == t_ret._miembros(data_base)
    assert vista == vista_base


@pytest.mark.parametrize(
    ("libro", "bandera"),
    [
        ("individual", None),
        ("general-mensual", None),
        ("general-general", None),
    ],
)
def test_la_tabla_en_cascada_en_usd_ignora_las_llaves(libro, bandera) -> None:
    """Hoja en dólares (sin `balance_en_pesos`), en cascada y con entregas:
    con las llaves nuevas y su discriminador, el libro de hoy."""
    p = _payload(libro, bandera)
    data_base, vista_base = tp._render(libro, sin_llaves(p))
    data, vista = tp._render(libro, p)
    assert t_ret._miembros(data) == t_ret._miembros(data_base)
    assert vista == vista_base


@pytest.mark.parametrize(("libro", "bandera"), list(t_res._HUELLAS_HEAD))
def test_las_tablas_legadas_ignoran_las_llaves_con_entregas(libro, bandera) -> None:
    """Tablas legadas (sin la reserva en la cascada), en USD y en pesos, con
    las entregas de siempre: las llaves nuevas —discriminador incluido— no
    cambian ni un byte."""
    base = tp._payload(libro, bandera)
    if libro == "individual":
        base = t_ent.con_entregas(base)
    else:
        base["aviones"] = [t_ent.con_entregas(a) for a in base["aviones"]]
    p = copy.deepcopy(base)
    for b in t_res._bloques(p):
        for s in b["socios"]:
            s.update(vuelos_utilidades_mxn=500.0, entregado_efectivo_mxn=1.0)
            s["entregado_efectivo_usd"] = 0.05
        b.update(vuelos_utilidades_socios_mxn=1500.0, entregado_efectivo_socios_mxn=3.0)
        b["vuelos_utilidades"] = copy.deepcopy(VUELOS)
    data_base, vista_base = tp._render(libro, base)
    data, vista = tp._render(libro, p)
    assert t_ret._miembros(data) == t_ret._miembros(data_base)
    assert vista == vista_base


@pytest.mark.parametrize(("libro", "bandera"), list(t_res._HUELLAS_HEAD))
def test_huellas_del_repo_con_las_llaves_nuevas(libro, bandera) -> None:
    """Los payloads de las huellas congeladas del repo (`_HUELLAS_HEAD` de
    `test_balance_reserva_overhaul`: sin cascada y sin entregas), con las
    llaves nuevas y su discriminador encima: el .xlsx y la vista en pantalla
    siguen siendo los de la huella."""
    p = tp._payload(libro, bandera)
    for b in t_res._bloques(p):
        for s in b["socios"]:
            s.update(vuelos_utilidades_mxn=500.0, entregado_efectivo_mxn=1.0)
            s["entregado_efectivo_usd"] = 0.05
        b.update(vuelos_utilidades_socios_mxn=1500.0, entregado_efectivo_socios_mxn=3.0)
        b["vuelos_utilidades"] = copy.deepcopy(VUELOS)
    data, vista = tp._render(libro, p)
    assert (tp._huella_xlsx(data), tp._huella_json(vista)) == t_res._HUELLAS_HEAD[(libro, bandera)]


# ---------------------------------------------------------------------------
# 4) (c) Balance GENERAL: cada bloque con SU número de columnas; lo que va
#    después (el otro avión, el bloque VUELATOUR, el pie) queda en su sitio.
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("variante", ["mensual", "general"])
def test_general_los_dos_aviones_con_vuelos(registros, variante) -> None:
    p = general(variante)
    data = render_balance_general_xlsx(BalanceGeneralRequest.model_validate(p))
    assert tp._degradadas_balance(registros[-1]) == []
    verificar_libro(data)
    wb, wv = tp._libros(data)
    ws, wsv = wb["balance"], wv["balance"]
    cobradas = tp._filas_de(ws, "UTILIDAD COBRADA (MXN)")
    titulos = tp._filas_de(ws, t_ret.TITULO_EN_CASCADA)
    assert len(cobradas) == len(titulos) == len(p["aviones"]) == 2
    for cob, titulo, avion in zip(cobradas, titulos, p["aviones"], strict=True):
        f = _tabla_con_vuelos(ws, wsv, cob, avion["balance"], TEXTOS_GENERAL[avion["matricula"]])
        assert titulo == f.siguiente + 1
    celdas = _celdas_de_merges(ws)
    assert len(celdas) == len(set(celdas))
    t_ret._ninguna_formula_cita_texto(ws)


@pytest.mark.parametrize("con_llaves", [0, 1], ids=["primer-avion", "segundo-avion"])
@pytest.mark.parametrize("variante", ["mensual", "general"])
def test_general_mixto_cada_bloque_con_sus_columnas_y_nada_se_encima(
    registros, variante, con_llaves
) -> None:
    """Un avión con vuelos pagados con utilidades (cuatro columnas, nota y
    detalle) y el otro solo con las entregas de siempre (tres). Contra el
    MISMO general sin las llaves nuevas: todo lo que no es la tabla del avión
    con llaves vale lo mismo, corrido exactamente las filas que se insertaron
    (la nota y el detalle), y el bloque VUELATOUR sigue citando a las socias."""
    p = general(variante)
    otro = 1 - con_llaves
    p["aviones"][otro] = sin_llaves(p["aviones"][otro])
    base = sin_llaves(p)
    data = render_balance_general_xlsx(BalanceGeneralRequest.model_validate(p))
    assert tp._degradadas_balance(registros[-1]) == []
    verificar_libro(data)
    data_base = render_balance_general_xlsx(BalanceGeneralRequest.model_validate(base))
    wb, wv = tp._libros(data)
    ws, wsv = wb["balance"], wv["balance"]
    ws_base, wsv_base = (x["balance"] for x in tp._libros(data_base))
    cobradas = tp._filas_de(ws, "UTILIDAD COBRADA (MXN)")
    assert len(cobradas) == 2
    avion = p["aviones"][con_llaves]
    f = _tabla_con_vuelos(
        ws, wsv, cobradas[con_llaves], avion["balance"], TEXTOS_GENERAL[avion["matricula"]]
    )
    # El otro bloque: las TRES columnas de hoy en G..I y nada en J.
    b_otro = p["aviones"][otro]["balance"]
    enc = cobradas[otro] + 3
    assert [ws.cell(row=enc, column=c).value for c in range(7, 12)] == [
        t_ent.COL_ENTREGADO,
        t_ent.COL_POR_ENTREGAR,
        t_ent.COL_USD_INFO,
        None,
        None,
    ]
    pie_otro = enc + len(b_otro["socios"]) + 2
    assert tp._filas_de(ws, t_ent.FILA_TOTAL) == [pie_otro]
    assert all(ws.cell(row=r, column=10).value is None for r in range(enc, pie_otro + 3))
    for i, s in enumerate(b_otro["socios"]):
        r = enc + 1 + i
        assert ws.cell(row=r, column=7).value == s["entregado_mxn"]
        assert ws.cell(row=r, column=8).value == f"=ROUND(E{r}-G{r},2)"
    assert wsv.cell(row=pie_otro, column=7).value == b_otro["entregado_socios_mxn"]
    assert tp._filas_de(ws, FILA_TOTAL_CON_VUELOS) == [f.pie]
    assert tp._filas_de(ws, NOTA_VUELOS) == [f.nota_vuelos]
    assert tp._filas_de(ws, TITULO_DETALLE) == [f.det_titulo]
    # Nada se encima: quitando las filas insertadas, cada celda vale lo que
    # valía (las columnas G..J de la tabla con llaves son las que cambian).
    insertadas = {f.nota_vuelos, *range(f.det_blanco, f.det_total + 1)}
    assert ws.max_row == ws_base.max_row + len(insertadas)
    corridas = 0
    for r in range(1, ws.max_row + 1):
        if r in insertadas:
            corridas += 1
            continue
        columnas = range(1, 7) if f.enc <= r <= f.pie else range(1, 12)
        for c in columnas:
            if (r, c) == (f.pie, 1):
                continue  # el rótulo del pie es el nuevo
            esperado = wsv_base.cell(row=r - corridas, column=c).value
            assert wsv.cell(row=r, column=c).value == esperado, (r, c)
    assert corridas == len(insertadas) == 1 + 3 + len(avion["balance"]["vuelos_utilidades"])
    # Las combinadas tampoco se pisan, y fuera de la tabla con llaves son las
    # de siempre, corridas.
    celdas = _celdas_de_merges(ws)
    assert len(celdas) == len(set(celdas))

    def fila_base(r: int) -> int:
        return r - sum(1 for x in insertadas if x < r)

    fuera = {
        (fila_base(m.min_row), m.min_col, fila_base(m.max_row), m.max_col)
        for m in ws.merged_cells.ranges
        if not f.titulo <= m.min_row <= f.det_total
    }
    fuera_base = {
        (m.min_row, m.min_col, m.max_row, m.max_col)
        for m in ws_base.merged_cells.ranges
        if not f.titulo <= m.min_row <= f.nota_socios - 1
    }
    assert fuera == fuera_base
    # VUELATOUR: la participación cita la celda MONTO (E) de cada socia y vale
    # lo del API; el pie de la hoja, al final.
    emp = p["empresa"]
    part = tp._fila_de(ws, "(+) PARTICIPACIÓN COMO SOCIO EN LOS AVIONES (MXN)")
    citadas = [int(x) for x in re.findall(r"E(\d+)", ws.cell(row=part, column=2).value)]
    assert len(citadas) == 2
    assert all(ws.cell(row=c, column=5).value == f"=ROUND(C{c}-D{c},2)" for c in citadas)
    assert [ws.cell(row=c, column=1).value for c in citadas] == [
        s["nombre"] for a in p["aviones"] for s in a["balance"]["socios"] if s["es_empresa"]
    ]
    assert wsv.cell(row=part, column=2).value == emp["participacion_mxn"]
    resultado = tp._fila_de(ws, "RESULTADO VUELATOUR (MXN)")
    assert wsv.cell(row=resultado, column=2).value == emp["resultado_mxn"]
    assert t_ret._textos_a(ws)[-1] == _NOTA_BALANCE_PESOS_EN_CASCADA
    # Anchos de la hoja: con UN bloque de cuatro columnas, G..J en 20.
    anchos = [ws.column_dimensions[c].width for c in "ABCDEFGHIJ"]
    assert anchos == [*_ANCHOS_BALANCE_PESOS_EN_CASCADA, *([_ANCHO_COL_ENTREGAS] * 4)]
    assert ws_base.column_dimensions["J"].width != _ANCHO_COL_ENTREGAS
    t_ret._ninguna_formula_cita_texto(ws)


# ---------------------------------------------------------------------------
# 5) (e) Vista en pantalla: el MISMO libro, con la columna y sus fórmulas.
# ---------------------------------------------------------------------------


def test_libro_json_exporta_la_columna_las_formulas_y_el_detalle() -> None:
    req = BalanceAvionRequest.model_validate(individual())
    libro = libro_balance_avion(req)
    assert tlj._comparar_con_xlsx(libro, render_balance_avion_xlsx(req)) > 0
    hoja = next(h for h in libro["hojas"] if h["nombre"] == "balance")
    celdas = {(c["r"], c["c"]): c for c in hoja["celdas"]}
    cob = next(r for (r, c), x in celdas.items() if c == 1 and x["v"] == "UTILIDAD COBRADA (MXN)")
    b = req.balance
    f = _filas(cob, len(b.socios), len(b.vuelos_utilidades))
    assert [celdas[(f.enc, c)]["v"] for c in range(7, 11)] == COLS_ENTREGAS
    assert (f.enc, 11) not in celdas
    for i, s in enumerate(b.socios):
        r = f.ini + i
        assert celdas[(r, 7)]["v"] == s.vuelos_utilidades_mxn and celdas[(r, 7)].get("f") is None
        assert celdas[(r, 8)]["v"] == s.entregado_efectivo_mxn
        assert celdas[(r, 9)]["f"] == f"=ROUND(E{r}-G{r}-H{r},2)"
        assert celdas[(r, 9)]["v"] == s.por_entregar_mxn
        assert celdas[(r, 10)]["v"] == s.entregado_efectivo_usd
    assert celdas[(f.pie, 1)]["v"] == FILA_TOTAL_CON_VUELOS
    assert celdas[(f.pie, 7)]["f"] == f"=ROUND(SUM(G{f.ini}:G{f.fin}),2)"
    assert celdas[(f.pie, 7)]["v"] == TOTAL_VUELOS
    assert celdas[(f.pie, 8)]["v"] == b.entregado_efectivo_socios_mxn
    assert celdas[(f.nota_vuelos, 1)]["v"] == NOTA_VUELOS
    assert celdas[(f.det_titulo, 1)]["v"] == TITULO_DETALLE
    assert [celdas[(f.det_ini + i, 1)]["v"] for i in range(2)] == TEXTOS_VUELOS
    assert [celdas[(f.det_ini + i, 7)]["v"] for i in range(2)] == [32210.43, 44174.3]
    assert celdas[(f.det_total, 7)]["f"] == f"=ROUND(SUM(G{f.det_ini}:G{f.det_fin}),2)"
    assert celdas[(f.det_total, 7)]["v"] == TOTAL_VUELOS
    # El general también (los dos bloques).
    g = BalanceGeneralRequest.model_validate(general("mensual"))
    lg = libro_balance_general(g)
    assert tlj._comparar_con_xlsx(lg, render_balance_general_xlsx(g)) > 0
    hoja_g = next(h for h in lg["hojas"] if h["nombre"] == "balance")
    valores = [x["v"] for x in hoja_g["celdas"]]
    assert valores.count(COL_VUELOS) == valores.count(TITULO_DETALLE) == 2
    assert valores.count(FILA_TOTAL_CON_VUELOS) == 2


# ---------------------------------------------------------------------------
# 6) (d) Rutas: el JSON del API 0.0.148 entra sin 422, también con basura en
#    el detalle.
# ---------------------------------------------------------------------------

TOKEN = "tok-vuelos-utilidades"
client = TestClient(app)


def test_rutas_aceptan_el_payload_y_la_basura_del_detalle(monkeypatch, request) -> None:
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    request.addfinalizer(get_settings.cache_clear)
    headers = {"X-Internal-Token": TOKEN}
    p = individual()
    res = client.post("/pdf/balance-avion-xlsx", json=p, headers=headers)
    assert res.status_code == 200, res.text
    ws = load_workbook(BytesIO(res.content))["balance"]
    assert tp._filas_de(ws, TEXTOS_VUELOS[0]) and tp._filas_de(ws, TEXTOS_VUELOS[1])
    res = client.post("/pdf/balance-avion-libro", json={**p, "archivo": "x.xlsx"}, headers=headers)
    assert res.status_code == 200, res.text
    hoja = next(h for h in res.json()["hojas"] if h["nombre"] == "balance")
    assert any(x["v"] == COL_VUELOS for x in hoja["celdas"])
    formulas = [str(x.get("f") or "") for x in hoja["celdas"]]
    assert any(re.fullmatch(r"=ROUND\(E(\d+)-G\1-H\1,2\)", f) for f in formulas)
    # Basura en el detalle: ni lista, ni filas con monto, ni folio entero.
    for basura in (
        "no soy una lista",
        {"folio": 498},
        [{"folio": "abc", "socio": 7, "monto_mxn": 10.5}, {"folio": 2}, "x", None, 3],
    ):
        sucio = individual()
        sucio["balance"]["vuelos_utilidades"] = basura
        for ruta in ("/pdf/balance-avion-xlsx", "/pdf/balance-avion-libro"):
            res = client.post(ruta, json=sucio, headers=headers)
            assert res.status_code == 200, (ruta, basura, res.text)
    # La última: una sola fila con monto (sin folio ni socio legibles).
    res = client.post("/pdf/balance-avion-xlsx", json=sucio, headers=headers)
    ws = load_workbook(BytesIO(res.content))["balance"]
    (titulo,) = tp._filas_de(ws, TITULO_DETALLE)
    assert ws.cell(row=titulo + 1, column=1).value == "Vuelo sin datos"
    assert ws.cell(row=titulo + 1, column=7).value == 10.5
    assert ws.cell(row=titulo + 2, column=1).value == "TOTAL"
    g = general("general")
    res = client.post("/pdf/balance-general-xlsx", json=g, headers=headers)
    assert res.status_code == 200, res.text
    ws = load_workbook(BytesIO(res.content))["balance"]
    assert len(tp._filas_de(ws, TITULO_DETALLE)) == 2
    res = client.post(
        "/pdf/balance-general-libro", json={**g, "archivo": "g.xlsx"}, headers=headers
    )
    assert res.status_code == 200, res.text
