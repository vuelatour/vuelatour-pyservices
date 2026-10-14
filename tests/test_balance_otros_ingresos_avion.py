"""Otros ingresos ASIGNADOS a un avión en los balances (13-oct-2026, API 0.0.75;
invariante «INGRESO ASIGNADO A UN AVIÓN» del API).

Pedido del cliente: un ingreso registrado en Ingresos con avión (el caso real
ING-4 de N990GG: «reembolso del pago del serv de 50 hrs en toledo, fact 53»,
$46,550.80 MXN) cuenta UNA vez, para ESE avión: pestaña «Otros ingresos» de su
libro, línea «(+) OTROS INGRESOS USD» que SUMA en la cascada (y entra al
reparto) y deja de ser ingreso de VuelaTour. El API manda todo calculado; aquí
se pinta con fórmulas que reproducen sus números:
- libro INDIVIDUAL: pestaña «Otros ingresos» entre 'permisos' y 'balance'
  (siempre que llegue la llave, aunque vacía) y la línea en la cascada, que
  cita su TOTAL USD;
- Balance GENERAL: la misma línea en cada bloque de 'balance' (VALOR, como el
  resto del bloque), hoja 'otros ingresos aviones' SOLO con filas y la cola de
  la fila 2 de 'otros movimientos' con la bandera `otros_ingresos_por_avion`;
- sin la llave (API ≤ 0.0.74): libros byte a byte (huellas de miembros
  calculadas con el código de HEAD d58faa1, antes de este cambio).
Los payloads salen del espejo del API de `tests/test_balance_formulas.py`.
"""

from __future__ import annotations

import copy
import hashlib
import zipfile
from io import BytesIO

import pytest
from fastapi.testclient import TestClient
from openpyxl import load_workbook

from app.config import get_settings
from app.main import app
from app.schemas.reportes import (
    BalanceAvionBalanceBloque,
    BalanceAvionHojaOtrosIngresos,
    BalanceAvionOtroIngresoFila,
    BalanceAvionRequest,
    BalanceGeneralRequest,
    BalanceHojaOtrosMovimientos,
)
from app.services import libro_json, xlsx_formulas
from app.services.balance_avion_xlsx import (
    _ENCABEZADOS_OTROS_INGRESOS,
    _FILA_OTROS_INGRESOS,
    _NOTA_CASCADA_OTROS_INGRESOS,
    _NOTA_OTROS_INGRESOS_AVION,
    _NOTA_OTROS_INGRESOS_AVIONES_GENERAL,
    _OM_OTROS_INGRESOS_POR_AVION,
    _SIN_OTROS_INGRESOS,
    _TITULO_OTROS_INGRESOS_GENERAL,
    libro_balance_general,
    render_balance_avion_xlsx,
    render_balance_general_xlsx,
)
from tests import test_balance_comisiones_general as tcg
from tests import test_balance_empresa_vuelatour as t_empresa
from tests import test_balance_formulas as tf
from tests._evaluador_formulas import verificar_libro

r2 = tf.r2
MAESTRA_TST = "reporte horas XA-TST"
HOJA = "Otros ingresos"
HOJA_GENERAL = "otros ingresos aviones"
DESPUES = "UTILIDAD DESPUÉS DE GASTOS USD"
COBRADA = "UTILIDAD COBRADA USD"

# ---------------------------------------------------------------------------
# Textos EXACTOS del contrato (§8 de la especificación).
# ---------------------------------------------------------------------------

NOTA_HOJA = (
    "Ingresos registrados en Ingresos (clave ING-n) y asignados a este avión: "
    "reembolsos de su servicio o de su aseguradora, ventas de piezas o activos, "
    "intereses… NETO MXN = ingreso − comisión del banco (en dólares, con el T.C. "
    "propio del ingreso). TOTAL USD = TOTAL NETO MXN ÷ TC PROMEDIO de las hojas "
    "de este libro. Ese TOTAL USD SUMA en la hoja 'balance' («(+) OTROS "
    "INGRESOS USD») y en el reparto a socios; el reparto convierte cada "
    "ingreso por separado (en dólares cuenta su neto tal cual; en pesos, con "
    "su T.C. propio o el oficial de su día), así que su monto en USD puede "
    "diferir del de esta hoja según la distancia entre esos T.C. y el promedio "
    "del libro (no son centavos), como pasa con los gastos. No cuentan como ingreso "
    "de VuelaTour: no están en 'otros movimientos' del Balance general ni en el "
    "total de 'Otros ingresos' del Libro Dinero. Los anticipos y las "
    "aportaciones nunca llevan avión."
)
NOTA_CASCADA = (
    "Ingresos registrados en Ingresos y asignados a este avión (hoja 'Otros "
    "ingresos' de su libro), netos de comisión bancaria, al T.C. promedio del "
    "libro. Suman a la utilidad del avión y a su reparto a socios; no cuentan "
    "como ingreso de VuelaTour."
)
NOTA_GENERAL = (
    "Ingresos registrados en Ingresos y asignados a un avión (clave ING-n), "
    "netos de comisión bancaria. Cada uno YA SUMA en el bloque de su avión en "
    "'balance' («(+) OTROS INGRESOS USD») y en su reparto a socios; por eso no "
    "están en 'otros movimientos' ni en el bloque VUELATOUR (empresa). TOTAL "
    "USD = suma de lo que suma cada avión (cada libro convierte con su propio "
    "T.C. promedio)."
)
OM_COLA = (
    " Los otros ingresos asignados a un avión (por ejemplo el reembolso de su "
    "servicio) no están aquí: cuentan en la hoja 'Otros ingresos' del libro de "
    "ese avión, en su bloque de 'balance' y en su reparto a socios (si hay "
    "alguno en el periodo, su detalle está en la hoja 'otros ingresos "
    "aviones')."
)
ENCABEZADOS = [
    "FECHA",
    "CATEGORÍA",
    "DETALLE",
    "MONTO ORIGINAL",
    "MONEDA",
    "T.C.",
    "INGRESO MXN",
    "COMISIÓN MXN",
    "NETO MXN",
    "CUENTA",
    "NOTAS",
]

# ---------------------------------------------------------------------------
# Filas como las arma el API (`filaHojaOtrosIngresos` sobre
# `filaLibroDeIngreso`).
# ---------------------------------------------------------------------------

_FILA_MXN = {
    "ingreso_id": "i-4",
    "clave": "ING-4",
    "fecha": "2026-09-04",
    "categoria": "Reembolsos y devoluciones recibidos",
    "detalle": "ING-4 · TRANSF INTERBANCARIA SPEI · angel alvarez",
    "moneda_original": "MXN",
    "monto_original": 46550.8,
    "tc_aplicado": None,
    "ingreso_mxn": 46550.8,
    "comision_mxn": None,
    "neto_mxn": 46550.8,
    "cuenta": "EJE SCOTIA · Scotiabank",
    "notas": "reembolso del pago del serv de 50 hrs en toledo, fact 53",
}
# USD 100 × 17.5 con comisión de 5 USD (= 87.50 MXN).
_FILA_USD = {
    "ingreso_id": "i-9",
    "clave": "ING-9",
    "fecha": "2026-09-20",
    "categoria": "Venta de activos",
    "detalle": "ING-9 · Venta de una hélice usada · Aero Partes",
    "moneda_original": "USD",
    "monto_original": 100.0,
    "tc_aplicado": 17.5,
    "ingreso_mxn": 1750.0,
    "comision_mxn": 87.5,
    "neto_mxn": 1662.5,
    "cuenta": "Efectivo / caja",
    "notas": None,
}
# USD SIN tipo de cambio: sin pesos ⇒ la hoja (y la utilidad) en null.
_FILA_USD_SIN_TC = {
    "ingreso_id": "i-11",
    "clave": "ING-11",
    "fecha": "2026-09-25",
    "categoria": "Otros ingresos",
    "detalle": "ING-11 · Intereses",
    "moneda_original": "USD",
    "monto_original": 200.0,
    "tc_aplicado": None,
    "ingreso_mxn": None,
    "comision_mxn": None,
    "neto_mxn": None,
    "cuenta": "Efectivo / caja",
    "notas": None,
}


def _tc_hojas(payload: dict) -> float | None:
    """`tcHojas` del API: promedio CRUDO de los T.C. de costos de las filas."""
    zs = [v["tc_costos"] for v in payload["vuelos"] if v["tc_costos"] is not None]
    return tf._suma(zs) / len(zs) if zs else None


def _hoja_api(filas: list[dict], tc_hojas: float | None) -> dict:
    """`hojaOtrosIngresosAvion` del API: total = Σ netos no nulos; usd con la
    regla de `buildHoja` (alguna fila sin pesos ⇒ null; total 0 ⇒ 0)."""
    total = r2(tf._suma(f["neto_mxn"] for f in filas if f["neto_mxn"] is not None))
    if any(f["neto_mxn"] is None for f in filas):
        usd = None
    elif total == 0:
        usd = 0
    else:
        usd = r2(total / tc_hojas) if tc_hojas is not None else None
    return {"filas": copy.deepcopy(filas), "total_mxn": total, "usd": usd}


def _con_otros_ingresos(payload: dict, filas: list[dict]) -> dict:
    """El libro de un avión como lo manda el API 0.0.75: la hoja y
    `balance.otros_ingresos_avion_usd`, con la utilidad después de gastos, la
    cobrada y los socios RE-ARMADOS (la línea SUMA; null se propaga)."""
    p = copy.deepcopy(payload)
    hoja = _hoja_api(filas, _tc_hojas(p))
    p["otros_ingresos_avion"] = hoja
    b = p["balance"]
    b["otros_ingresos_avion_usd"] = hoja["usd"]
    restas = [
        p["combustible"]["usd"],
        *(p[n]["usd"] for n in ("gastos_indirectos", "otros_gastos", "refacciones", "permisos")),
    ]
    antes = b["utilidad_antes_usd"]
    oi = hoja["usd"]
    despues = (
        r2(antes - tf._suma(restas) + oi)
        if oi is not None and all(x is not None for x in restas)
        else None
    )
    cobrada = r2(despues - p["totales"]["por_cobrar_usd"]) if despues is not None else None
    b["utilidad_despues_usd"] = despues
    b["utilidad_cobrada_usd"] = cobrada
    for s in b["socios"]:
        s["monto_usd"] = r2((s["porcentaje"] / 100) * cobrada) if cobrada is not None else None
    return p


def _individual(filas: list[dict] | None = None) -> BalanceAvionRequest:
    filas = [_FILA_MXN, _FILA_USD] if filas is None else filas
    return BalanceAvionRequest.model_validate(_con_otros_ingresos(tf._payload_tst(), filas))


def _libros(data: bytes):
    """(con fórmulas, con caché)."""
    return load_workbook(BytesIO(data)), load_workbook(BytesIO(data), data_only=True)


def _miembros(data: bytes) -> dict[str, bytes]:
    """Todos los miembros del .xlsx salvo docProps/core.xml (lleva la hora)."""
    z = zipfile.ZipFile(BytesIO(data))
    return {n: z.read(n) for n in sorted(z.namelist()) if n != "docProps/core.xml"}


def _huella(data: bytes) -> str:
    h = hashlib.sha256()
    for nombre, contenido in _miembros(data).items():
        h.update(nombre.encode())
        h.update(contenido)
    return h.hexdigest()


def _fila_de(ws, etiqueta: str) -> int:
    filas = [c.row for c in ws["A"] if c.value == etiqueta]
    assert len(filas) == 1, (etiqueta, filas)
    return filas[0]


@pytest.fixture
def registros(monkeypatch) -> list[xlsx_formulas.RegistroFormulas]:
    """Registro de fórmulas de cada libro que se cierra (cuáles volvieron a
    valor)."""
    capturados: list[xlsx_formulas.RegistroFormulas] = []
    original = xlsx_formulas.finalizar

    def espia(wb):
        cache = original(wb)
        capturados.append(xlsx_formulas.registro(wb))
        return cache

    monkeypatch.setattr(xlsx_formulas, "finalizar", espia)
    return capturados


# Las dos celdas que el espejo del API deja como VALOR a propósito en el libro
# XA-TST (redondeo intermedio de COSTO TOTAL del #405 y TOTAL USD de 'Gastos
# Indirectos', `test_balance_formulas`): los otros ingresos no agregan ninguna.
_DEGRADADAS_TST = [(MAESTRA_TST, "U7"), ("Gastos Indirectos", "C5")]


# ---------------------------------------------------------------------------
# 1) Esquema y textos.
# ---------------------------------------------------------------------------


def test_esquemas_aditivos_con_default() -> None:
    assert BalanceAvionRequest().otros_ingresos_avion is None
    assert BalanceAvionBalanceBloque().otros_ingresos_avion_usd is None
    assert BalanceHojaOtrosMovimientos().otros_ingresos_por_avion is None
    hoja = BalanceAvionHojaOtrosIngresos.model_validate({"filas": None})
    assert (hoja.filas, hoja.total_mxn, hoja.usd) == ([], None, None)
    fila = BalanceAvionOtroIngresoFila.model_validate({"clave": None, "monto_original": 5})
    assert (fila.clave, fila.monto_original, fila.neto_mxn) == ("", 5.0, None)
    # Las llaves desconocidas se ignoran (skew: un API más nuevo no da 422).
    req = BalanceAvionRequest.model_validate(
        {"otros_ingresos_avion": {"filas": [{**_FILA_MXN, "algo_nuevo": 1}], "usd": 2}}
    )
    assert req.otros_ingresos_avion.filas[0].clave == "ING-4"


def test_textos_son_los_del_contrato() -> None:
    assert _NOTA_OTROS_INGRESOS_AVION == NOTA_HOJA
    assert _NOTA_CASCADA_OTROS_INGRESOS == NOTA_CASCADA
    assert _NOTA_OTROS_INGRESOS_AVIONES_GENERAL == NOTA_GENERAL
    assert _OM_OTROS_INGRESOS_POR_AVION == OM_COLA
    assert _ENCABEZADOS_OTROS_INGRESOS == ENCABEZADOS
    assert _FILA_OTROS_INGRESOS == "(+) OTROS INGRESOS USD"
    assert _SIN_OTROS_INGRESOS == "Sin otros ingresos asignados a este avión en el periodo."
    assert _TITULO_OTROS_INGRESOS_GENERAL == "Otros ingresos asignados a aviones — flota"
    # Ninguno lleva la marca del amarillo de 'otros movimientos'.
    for t in (NOTA_HOJA, NOTA_CASCADA, NOTA_GENERAL, OM_COLA):
        assert "· PROVISIÓN (" not in t


# ---------------------------------------------------------------------------
# 2) Libro INDIVIDUAL.
# ---------------------------------------------------------------------------


def test_individual_pestana_entre_permisos_y_balance() -> None:
    wb = load_workbook(BytesIO(render_balance_avion_xlsx(_individual())))
    assert wb.sheetnames == [
        MAESTRA_TST,
        "cobranza",
        "combustible",
        "Gastos Indirectos",
        "refacciones",
        "permisos",
        HOJA,
        "balance",
        "pendientes de captura",
    ]
    # Sin la hoja 'refacciones' (API viejo): sigue entre 'permisos' y 'balance'.
    p = _con_otros_ingresos(tf._payload_tst(), [_FILA_MXN])
    p["refacciones"] = None
    wb = load_workbook(BytesIO(render_balance_avion_xlsx(BalanceAvionRequest.model_validate(p))))
    assert wb.sheetnames[-4:] == ["permisos", HOJA, "balance", "pendientes de captura"]


def test_individual_hoja_encabezados_filas_y_formulas(registros) -> None:
    req = _individual()
    data = render_balance_avion_xlsx(req)
    wb, wv = _libros(data)
    ws, wsv = wb[HOJA], wv[HOJA]
    assert ws["A1"].value == "Otros ingresos — XA-TST"
    assert ws["A2"].value == "Periodo: 2026-09-01 a 2026-09-30"
    assert [ws.cell(row=4, column=c).value for c in range(1, 4)] == [
        "TOTAL NETO MXN",
        "TC PROMEDIO",
        "TOTAL USD",
    ]
    assert [ws.cell(row=7, column=c).value for c in range(1, 12)] == ENCABEZADOS
    # Fila MXN (8) y USD (9).
    assert [ws.cell(row=8, column=c).value for c in range(1, 6)] == [
        "04/09/2026",
        "Reembolsos y devoluciones recibidos",
        "ING-4 · TRANSF INTERBANCARIA SPEI · angel alvarez",
        46550.8,
        "MXN",
    ]
    assert ws["F8"].value is None  # sin T.C. en pesos
    assert ws["G8"].value == "=ROUND(D8,2)"
    assert ws["H8"].value is None  # sin comisión
    assert ws["I8"].value == "=ROUND(G8,2)"
    assert (ws["J8"].value, ws["K8"].value) == (
        "EJE SCOTIA · Scotiabank",
        "reembolso del pago del serv de 50 hrs en toledo, fact 53",
    )
    assert (ws["D9"].value, ws["E9"].value, ws["F9"].value) == (100, "USD", 17.5)
    assert ws["G9"].value == "=ROUND(D9*F9,2)"
    assert ws["H9"].value == 87.5
    assert ws["I9"].value == "=ROUND(G9-H9,2)"
    # Resumen: fórmulas que citan la hoja y la maestra.
    assert ws["A5"].value == "=ROUND(SUM(I8:I9),2)"
    assert ws["B5"].value == f"='{MAESTRA_TST}'!$V$9"
    assert ws["C5"].value == f"=ROUND(A5/AVERAGE('{MAESTRA_TST}'!$V$3:$V$8),2)"
    # Los números son los del API (caché) y ninguna fórmula volvió a valor.
    hoja = req.otros_ingresos_avion
    assert (wsv["G8"].value, wsv["I8"].value) == (46550.8, 46550.8)
    assert (wsv["G9"].value, wsv["I9"].value) == (1750, 1662.5)
    assert wsv["A5"].value == hoja.total_mxn == r2(46550.8 + 1662.5)
    assert wsv["C5"].value == hoja.usd == r2(hoja.total_mxn / _tc_hojas(tf._payload_tst()))
    assert registros[-1].degradadas() == _DEGRADADAS_TST
    assert verificar_libro(data)
    # Nota al pie (envuelta, en A:K) y panel fijo bajo los encabezados.
    pie = ws.cell(row=11, column=1)
    assert pie.value == NOTA_HOJA and pie.alignment.wrap_text
    assert any(str(rg) == "A11:K11" for rg in ws.merged_cells.ranges)
    assert ws.freeze_panes == "A8"
    assert ws.column_dimensions["C"].width == 50 and ws.column_dimensions["K"].width == 36


def test_individual_cascada_suma_los_otros_ingresos(registros) -> None:
    req = _individual()
    data = render_balance_avion_xlsx(req)
    wb, wv = _libros(data)
    ws, wsv = wb["balance"], wv["balance"]
    k = _fila_de(ws, "UTILIDAD ANTES DE GASTOS USD")
    assert [ws.cell(row=k + i, column=1).value for i in range(9)] == [
        "UTILIDAD ANTES DE GASTOS USD",
        "(−) COMBUSTIBLE DEL MES USD",
        "(−) GASTOS INDIRECTOS USD",
        "(−) REFACCIONES (INVENTARIO) USD",
        "(−) PERMISOS USD",
        _FILA_OTROS_INGRESOS,
        DESPUES,
        "(−) PENDIENTE DE PAGO (COBRANZA PENDIENTE) USD",
        COBRADA,
    ]
    linea = ws.cell(row=k + 5, column=2)
    assert linea.value == "='Otros ingresos'!$C$5"
    assert linea.comment is not None and linea.comment.text == NOTA_CASCADA
    assert ws.cell(row=k + 6, column=2).value == (
        f"=ROUND(B{k}-B{k + 1}-B{k + 2}-B{k + 3}-B{k + 4}+B{k + 5},2)"
    )
    assert ws.cell(row=k + 8, column=2).value == f"=ROUND(B{k + 6}-B{k + 7},2)"
    b = req.balance
    assert wsv.cell(row=k + 5, column=2).value == b.otros_ingresos_avion_usd
    assert wsv.cell(row=k + 6, column=2).value == b.utilidad_despues_usd
    assert wsv.cell(row=k + 8, column=2).value == b.utilidad_cobrada_usd
    # Los socios citan la UTILIDAD COBRADA corrida.
    socio = _fila_de(ws, "Socio A")
    assert ws.cell(row=socio, column=3).value == f"=ROUND(B{socio}/100*$B${k + 8},2)"
    assert [wsv.cell(row=socio + i, column=3).value for i in range(2)] == [
        s.monto_usd for s in b.socios
    ]
    # La utilidad sube EXACTAMENTE lo que suman los otros ingresos.
    sin = tf._payload_tst()["balance"]
    assert b.utilidad_despues_usd == r2(sin["utilidad_despues_usd"] + b.otros_ingresos_avion_usd)
    assert registros[-1].degradadas() == _DEGRADADAS_TST


def test_individual_hoja_vacia_texto_y_linea_en_cero() -> None:
    req = _individual([])
    assert (req.otros_ingresos_avion.total_mxn, req.otros_ingresos_avion.usd) == (0, 0)
    data = render_balance_avion_xlsx(req)
    wb, wv = _libros(data)
    ws = wb[HOJA]
    assert ws["A8"].value == _SIN_OTROS_INGRESOS
    assert any(str(rg) == "A8:K8" for rg in ws.merged_cells.ranges)
    assert ws["A5"].value == 0  # sin filas no hay rango que sumar: valor
    assert wv[HOJA]["C5"].value == 0
    k = _fila_de(wb["balance"], _FILA_OTROS_INGRESOS)
    assert wb["balance"].cell(row=k, column=2).value == "='Otros ingresos'!$C$5"
    assert wv["balance"].cell(row=k, column=2).value == 0
    assert verificar_libro(data)


def test_individual_usd_sin_tc_celdas_vacias_pero_la_fila_existe() -> None:
    req = _individual([_FILA_MXN, _FILA_USD_SIN_TC])
    assert req.otros_ingresos_avion.usd is None
    assert req.balance.utilidad_despues_usd is None
    data = render_balance_avion_xlsx(req)
    wb, wv = _libros(data)
    ws = wb[HOJA]
    # La fila USD sin T.C.: sin pesos ni fórmula (jamás un USD como MXN).
    assert (ws["G9"].value, ws["I9"].value, ws["F9"].value) == (None, None, None)
    assert ws["D9"].value == 200 and ws["E9"].value == "USD"
    # TOTAL NETO = solo los pesos; TOTAL USD vacío.
    assert wv[HOJA]["A5"].value == 46550.8
    assert ws["C5"].value is None
    bal = wb["balance"]
    k = _fila_de(bal, _FILA_OTROS_INGRESOS)
    assert bal.cell(row=k, column=2).value is None
    assert bal.cell(row=k, column=2).comment is None
    assert bal.cell(row=k + 1, column=1).value == DESPUES
    assert bal.cell(row=k + 1, column=2).value is None
    assert verificar_libro(data)


@pytest.mark.parametrize("valor", [..., None], ids=["sin-llave", "null"])
def test_individual_sin_la_llave_libro_de_hoy(valor) -> None:
    p = tf._payload_tst()
    if valor is not ...:
        p["otros_ingresos_avion"] = valor
    data = render_balance_avion_xlsx(BalanceAvionRequest.model_validate(p))
    wb = load_workbook(BytesIO(data))
    assert HOJA not in wb.sheetnames
    assert _FILA_OTROS_INGRESOS not in [c.value for c in wb["balance"]["A"]]
    assert wb["balance"]["B9"].value == "=ROUND(B4-B5-B6-B7-B8,2)"
    assert _huella(data) == _HUELLA_INDIVIDUAL_HEAD


# ---------------------------------------------------------------------------
# 3) Balance GENERAL.
# ---------------------------------------------------------------------------

_COLOR_TST = "#3B82F6"


def _fusion_flota(libros: list[dict]) -> dict | None:
    """`fusionarOtrosIngresosFlota` del API: solo libros con filas; detalle con
    «MATRÍCULA · » al frente, matrícula y color por fila; orden por fecha
    (estable); totales sumados y null que se propaga."""
    con = [p for p in libros if (p.get("otros_ingresos_avion") or {}).get("filas")]
    if not con:
        return None
    filas = sorted(
        (
            {
                **f,
                "detalle": f"{p['matricula']} · {f['detalle']}",
                "matricula": p["matricula"],
                "avion_color": p.get("avion_color"),
            }
            for p in con
            for f in p["otros_ingresos_avion"]["filas"]
        ),
        key=lambda f: f["fecha"],
    )
    usds = [p["otros_ingresos_avion"]["usd"] for p in con]
    return {
        "filas": filas,
        "total_mxn": r2(tf._suma(p["otros_ingresos_avion"]["total_mxn"] for p in con)),
        "usd": None if any(u is None for u in usds) else r2(tf._suma(usds)),
    }


def _general_dict(*, bandera: bool = True, filas_tst=None, filas_dos=None) -> dict:
    tst = tf._payload_tst()
    tst["avion_color"] = _COLOR_TST
    tst = _con_otros_ingresos(tst, [_FILA_MXN, _FILA_USD] if filas_tst is None else filas_tst)
    dos = _con_otros_ingresos(tf._payload_dos(), [] if filas_dos is None else filas_dos)
    p = tf._general_payload([tst, dos])
    oi = _fusion_flota([tst, dos])
    if oi is not None:
        p["consolidado"]["otros_ingresos_avion"] = oi
    if bandera:
        p["consolidado"]["otros_movimientos"]["otros_ingresos_por_avion"] = True
    return p


def _general(**kw) -> BalanceGeneralRequest:
    return BalanceGeneralRequest.model_validate(_general_dict(**kw))


def _bloque(ws, matricula: str) -> int:
    """Fila de «UTILIDAD ANTES DE GASTOS USD» del bloque de `matricula` en la
    hoja 'balance' del general."""
    inicio = _fila_de(ws, matricula)
    return next(
        c.row for c in ws["A"] if c.row > inicio and c.value == "UTILIDAD ANTES DE GASTOS USD"
    )


def test_general_hoja_otros_ingresos_aviones(registros) -> None:
    req = _general()
    data = render_balance_general_xlsx(req)
    wb, wv = _libros(data)
    nombres = wb.sheetnames
    assert HOJA_GENERAL in nombres
    assert nombres.index(HOJA_GENERAL) == nombres.index("repartidos a aviones") + 1
    assert nombres.index("inventario") == nombres.index(HOJA_GENERAL) + 1
    ws, wsv = wb[HOJA_GENERAL], wv[HOJA_GENERAL]
    assert ws["A1"].value == _TITULO_OTROS_INGRESOS_GENERAL
    assert [ws.cell(row=4, column=c).value for c in range(1, 4)] == [
        "TOTAL NETO MXN",
        "TOTAL USD (suma de los aviones)",
        None,
    ]
    assert [ws.cell(row=7, column=c).value for c in range(1, 12)] == ENCABEZADOS
    # Detalle con la matrícula al frente y teñido con el color del avión.
    assert ws["C8"].value == "XA-TST · ING-4 · TRANSF INTERBANCARIA SPEI · angel alvarez"
    assert ws["C8"].fill.fgColor.rgb.endswith("3B82F6")
    assert ws["A5"].value == "=ROUND(SUM(I8:I9),2)"
    oi = req.consolidado.otros_ingresos_avion
    assert wsv["A5"].value == oi.total_mxn
    # TOTAL USD = VALOR (cada libro convirtió con su T.C.: sin fórmula).
    assert ws["B5"].value == oi.usd == req.aviones[0].otros_ingresos_avion.usd
    assert [c.value for c in ws["A"] if c.value == NOTA_GENERAL] == [NOTA_GENERAL]
    assert all(h != HOJA_GENERAL for h, _ in registros[-1].degradadas())
    assert verificar_libro(data)


def test_general_sin_filas_no_hay_hoja_pero_los_bloques_llevan_la_linea() -> None:
    req = _general(filas_tst=[])
    assert req.consolidado.otros_ingresos_avion is None
    data = render_balance_general_xlsx(req)
    wb, wv = _libros(data)
    assert HOJA_GENERAL not in wb.sheetnames
    ws = wb["balance"]
    for mat in ("XA-TST", "XB-DOS"):
        k = _bloque(ws, mat)
        assert ws.cell(row=k + 5, column=1).value == _FILA_OTROS_INGRESOS
        assert wv["balance"].cell(row=k + 5, column=2).value == 0
    assert verificar_libro(data)


def test_general_bloques_con_la_linea_como_valor_y_comentario() -> None:
    req = _general()
    data = render_balance_general_xlsx(req)
    wb, wv = _libros(data)
    ws, wsv = wb["balance"], wv["balance"]
    for mat, avion in (("XA-TST", req.aviones[0]), ("XB-DOS", req.aviones[1])):
        k = _bloque(ws, mat)
        linea = ws.cell(row=k + 5, column=2)
        assert ws.cell(row=k + 5, column=1).value == _FILA_OTROS_INGRESOS
        # VALOR (el general no tiene la hoja del avión) con su nota.
        assert linea.value == avion.balance.otros_ingresos_avion_usd
        assert linea.comment is not None and linea.comment.text == NOTA_CASCADA
        assert ws.cell(row=k + 6, column=2).value == (
            f"=ROUND(B{k}-B{k + 1}-B{k + 2}-B{k + 3}-B{k + 4}+B{k + 5},2)"
        )
        assert ws.cell(row=k + 8, column=2).value == f"=ROUND(B{k + 6}-B{k + 7},2)"
        assert wsv.cell(row=k + 6, column=2).value == avion.balance.utilidad_despues_usd
        assert wsv.cell(row=k + 8, column=2).value == avion.balance.utilidad_cobrada_usd
    assert verificar_libro(data)


def test_general_con_comisiones_la_linea_va_tras_las_dos_previas(registros) -> None:
    """Cascada del general con «antes de comisiones» y «(−) COMISIONES» (API
    0.0.66): la línea de otros ingresos sigue en k + 5 contando desde UTILIDAD
    ANTES DE GASTOS (que ya va corrida dos filas)."""
    p = tcg._payload()
    tst = next(a for a in p["aviones"] if a["matricula"] == "XA-TST")
    con = _con_otros_ingresos(tst, [_FILA_MXN])
    p["aviones"] = [con if a["matricula"] == "XA-TST" else a for a in p["aviones"]]
    p["empresa"] = t_empresa._empresa_api(p)
    req = BalanceGeneralRequest.model_validate(p)
    data = render_balance_general_xlsx(req)
    wb, wv = _libros(data)
    ws, wsv = wb["balance"], wv["balance"]
    k = _bloque(ws, "XA-TST")
    assert ws.cell(row=k - 2, column=1).value == (
        "UTILIDAD ANTES DE GASTOS USD (antes de comisiones)"
    )
    assert ws.cell(row=k + 5, column=1).value == _FILA_OTROS_INGRESOS
    assert ws.cell(row=k + 6, column=2).value == (
        f"=ROUND(B{k}-B{k + 1}-B{k + 2}-B{k + 3}-B{k + 4}+B{k + 5},2)"
    )
    b = req.aviones[[a.matricula for a in req.aviones].index("XA-TST")].balance
    assert wsv.cell(row=k + 6, column=2).value == b.utilidad_despues_usd
    assert wsv.cell(row=k + 8, column=2).value == b.utilidad_cobrada_usd
    # XB-DOS no trae la llave (API previo en ese libro): su cascada de siempre.
    kd = _bloque(ws, "XB-DOS")
    assert ws.cell(row=kd + 5, column=1).value == DESPUES
    assert verificar_libro(data)


@pytest.mark.parametrize("bandera", [True, None, False], ids=["true", "null", "false"])
def test_general_fila2_de_otros_movimientos(bandera) -> None:
    p = _general_dict(bandera=False)
    if bandera is not None:
        p["consolidado"]["otros_movimientos"]["otros_ingresos_por_avion"] = bandera
    wb = load_workbook(
        BytesIO(render_balance_general_xlsx(BalanceGeneralRequest.model_validate(p)))
    )
    a2 = wb["otros movimientos"]["A2"].value
    if bandera is True:
        assert a2.endswith(OM_COLA)
        assert a2.count(OM_COLA) == 1
    else:
        assert OM_COLA not in a2
        assert a2.endswith("antes llamada 'gastos VuelaTour').")


def test_general_sin_la_llave_libro_de_hoy() -> None:
    data = render_balance_general_xlsx(tf._general())
    assert HOJA_GENERAL not in load_workbook(BytesIO(data)).sheetnames
    assert _huella(data) == _HUELLA_GENERAL_HEAD


def test_libro_json_del_general_exporta_la_hoja() -> None:
    libro = libro_balance_general(_general())
    hojas = {h["nombre"]: h for h in libro["hojas"]}
    assert HOJA_GENERAL in hojas
    celdas = {(c["r"], c["c"]): c for c in hojas[HOJA_GENERAL]["celdas"]}
    assert celdas[(5, 1)]["f"] == "=ROUND(SUM(I8:I9),2)"
    assert celdas[(8, 7)]["f"] == "=ROUND(D8,2)"
    assert celdas[(9, 9)]["f"] == "=ROUND(G9-H9,2)"
    assert libro_json.VERSION_LIBRO == 1


# ---------------------------------------------------------------------------
# 4) Rutas con el JSON del API 0.0.75.
# ---------------------------------------------------------------------------

TOKEN = "tok-otros-ingresos"
client = TestClient(app)


def test_rutas_aceptan_el_payload_0075(monkeypatch, request) -> None:
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    request.addfinalizer(get_settings.cache_clear)
    headers = {"X-Internal-Token": TOKEN}
    individual = _con_otros_ingresos(tf._payload_tst(), [_FILA_MXN, _FILA_USD])
    res = client.post("/pdf/balance-avion-xlsx", json=individual, headers=headers)
    assert res.status_code == 200, res.text
    assert HOJA in load_workbook(BytesIO(res.content)).sheetnames
    res = client.post("/pdf/balance-general-xlsx", json=_general_dict(), headers=headers)
    assert res.status_code == 200, res.text
    assert HOJA_GENERAL in load_workbook(BytesIO(res.content)).sheetnames
    res = client.post("/pdf/balance-avion-libro", json=individual, headers=headers)
    assert res.status_code == 200, res.text
    assert HOJA in [h["nombre"] for h in res.json()["hojas"]]


# ---------------------------------------------------------------------------
# Huellas de miembros (todos salvo docProps/core.xml) de los libros SIN la
# llave, calculadas con el código de HEAD d58faa1 (13-oct-2026), ANTES de
# este cambio: si una se mueve, un API ≤ 0.0.74 ya no ve su libro de siempre.
# ---------------------------------------------------------------------------

_HUELLA_INDIVIDUAL_HEAD = "4bb86191e6fdd486fb46a332b612b3b763908f6ec8c14ff9e8d9eaee4e1a506f"
_HUELLA_GENERAL_HEAD = "dd8a3740c2ef71791c94cb35ac02a3f71123fc49979b3d15a91569cbfa56ba90"
