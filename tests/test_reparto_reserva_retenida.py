"""La RESERVA DE OVERHAUL SE RETIENE en el reparto a socios, PDF y Excel
(7-nov-2026, API 0.0.128).

Decisión del cliente: la reserva debe «retenerse bien … y verse también en el
reparto de utilidades». Con `saldo_a_repartir_usd` (el DISCRIMINADOR:
round2(saldo − reserva); 0.0 también cuenta) y `saldo_a_repartir_mxn`, y la
cascada del balance (`reparto_como_balance`), la reserva vuelve a RESTAR en la
cascada: «Saldo antes de la reserva de overhaul» (`saldo_*`) → «(-) Reserva de
overhaul (se retiene para el overhaul)» → «Saldo a repartir»; los socios y el
resumen dicen «Reserva retenida»; el pie explica; el resumen global suma el
saldo a repartir (y sus gastos incluyen la reserva: venta − gastos = saldo a
repartir). En el Excel la reserva deja de ser gris «informativa» y entra la
columna «Saldo a repartir» tras el saldo (rotulado «Saldo antes de reserva»).
`reparto[].monto_*` llega ya neto (Σ = saldo a repartir al centavo). Sin la
llave (API ≤ 0.0.127), o sin la cascada del balance, PDF y Excel BYTE A BYTE
los de siempre (`test_reparto_como_balance`, bloque «RESERVA PARA TODOS»).
"""

from __future__ import annotations

import copy
from io import BytesIO

import pytest
from fastapi.testclient import TestClient
from openpyxl import load_workbook
from reportlab.lib.units import mm
from reportlab.pdfbase.pdfmetrics import stringWidth

from app.config import get_settings
from app.main import app
from app.schemas.reparto import RepartoAvion, RepartoPdfRequest
from app.services import reparto_combustible as rc
from app.services import reparto_pdf as rp
from app.services import reparto_reserva as rr
from app.services import reparto_xlsx as rx
from app.services.reparto_pdf import render_reparto_pdf
from app.services.reparto_xlsx import render_reparto_xlsx
from tests import _pesos_payloads as pp
from tests import test_reparto_combustible_separado as tcs
from tests import test_reparto_como_balance as tcb
from tests.test_reparto_en_pesos import _texto_pdf

r2 = pp.r2
_d = tcb._d
K = pp.K_REPARTO
SALDO_TST, SALDO_DOS = tcb.SALDO_TST, tcb.SALDO_DOS
RESERVA_TST = 140.0
A_REPARTIR_TST = 914.30  # 1,054.30 − 140.00
UTIL_A, RESERVA_A, MONTO_A = tcb.UTIL_A, tcb.RESERVA_A, tcb.MONTO_A
UTIL_B, RESERVA_B = tcb.UTIL_B, tcb.RESERVA_B
MONTO_B = r2(UTIL_B - RESERVA_B)
_flujos, _miembros, _fila, _textos_col_a = tcb._flujos, tcb._miembros, tcb._fila, tcb._textos_col_a


# ---------------------------------------------------------------------------
# Payloads: el reparto donde todos descuentan (API 0.0.125) con las llaves
# del 0.0.128, espejo de la aritmética del API.
# ---------------------------------------------------------------------------


def con_retencion(p: dict) -> dict:
    """`saldo_a_repartir_*` = saldo − reserva por avión (los socios ya vienen
    netos: todos descuentan)."""
    p = copy.deepcopy(p)
    for a in p["aviones"]:
        a["saldo_a_repartir_usd"] = r2(a["saldo_usd"] - a["reserva_overhaul_usd"])
        a["saldo_a_repartir_mxn"] = (
            None
            if a.get("saldo_mxn") is None or a.get("reserva_overhaul_mxn") is None
            else r2(a["saldo_mxn"] - a["reserva_overhaul_mxn"])
        )
    return p


def _payload(*, pesos: bool | None = True) -> dict:
    return con_retencion(tcb._payload_todos(pesos=pesos))


def _req(**kw) -> RepartoPdfRequest:
    return RepartoPdfRequest.model_validate(_payload(**kw))


def _hoja(p: dict):
    return load_workbook(BytesIO(render_reparto_xlsx(RepartoPdfRequest.model_validate(p)))).active


def _n621tx_retenido(*, pesos: bool, otros_ingresos: float = 0.0) -> dict:
    """El caso real N621TX del combustible separado con TODOS descontando y
    la reserva retenida (saldo a repartir = saldo − 15,000)."""
    p = tcs._payload(pesos=pesos, otros_ingresos=otros_ingresos)
    for a in p["aviones"]:
        for s in a["reparto"]:
            s["descuenta_reserva"] = True
            s["monto_usd"] = r2(s["utilidad_usd"] - s["reserva_usd"])
            s["monto_mxn"] = r2(s["utilidad_mxn"] - s["reserva_mxn"])
        a["reserva_retenida_usd"] = a["reserva_overhaul_usd"]
        a["reserva_retenida_mxn"] = a["reserva_overhaul_mxn"]
    return con_retencion(p)


# ---------------------------------------------------------------------------
# 1) Números, esquema y helpers.
# ---------------------------------------------------------------------------


def test_los_numeros_del_payload_cuadran() -> None:
    req = _req()
    tst, dos = req.aviones
    assert tst.saldo_a_repartir_usd == A_REPARTIR_TST == r2(SALDO_TST - RESERVA_TST)
    assert tst.saldo_a_repartir_mxn == r2(tst.saldo_mxn - tst.reserva_overhaul_mxn)
    # Σ de lo que reciben los socios = saldo a repartir, al centavo.
    assert r2(sum(s.monto_usd for s in tst.reparto)) == A_REPARTIR_TST
    assert r2(sum(s.monto_mxn for s in tst.reparto)) == tst.saldo_a_repartir_mxn
    assert tst.reserva_retenida_usd == RESERVA_TST
    assert (MONTO_A, MONTO_B) == (548.58, 365.72)
    assert dos.saldo_a_repartir_usd == SALDO_DOS and dos.reserva_overhaul_usd == 0


def test_esquema_aditivo_y_helpers() -> None:
    a = RepartoAvion(matricula="X", modelo="M", ingresos_cobrado_usd=0, saldo_usd=10)
    assert a.saldo_a_repartir_usd is None and a.saldo_a_repartir_mxn is None
    assert not rr.reserva_retenida(a)
    assert rr.saldo_a_repartir_usd(a, como_balance=True) == 10
    a.saldo_a_repartir_usd = 0.0
    assert rr.reserva_retenida(a)
    assert rr.saldo_a_repartir_usd(a, como_balance=True) == 0.0
    # Sin la cascada del balance el saldo del API ya trae la reserva: el de siempre.
    assert rr.saldo_a_repartir_usd(a, como_balance=False) == 10
    a.saldo_mxn, a.saldo_a_repartir_mxn = 170.0, None
    assert rr.saldo_a_repartir_mxn(a, como_balance=True) is None
    assert rr.saldo_a_repartir_mxn(a, como_balance=False) == 170.0
    assert rr.reserva_retenida_req(_req())
    assert not rr.reserva_retenida_req(RepartoPdfRequest.model_validate(tcb._payload_todos()))
    vacio = RepartoPdfRequest.model_validate({**_payload(), "aviones": []})
    assert not rr.reserva_retenida_req(vacio)


def test_etiquetas_caben_y_sin_caracteres_que_helvetica_no_dibuja() -> None:
    ancho = 112 * mm - 12
    for t in (rp.FILA_SALDO_ANTES_RESERVA, rp.FILA_RESERVA_RETENIDA, rp.FILA_SALDO_A_REPARTIR):
        assert stringWidth(t, "Helvetica-Bold", 9) <= ancho, t
    for t in (
        rp.FILA_SALDO_ANTES_RESERVA, rp.FILA_RESERVA_RETENIDA, rp.FILA_SALDO_A_REPARTIR,
        rp._PIE_CASCADA_BALANCE_RETENIDA,
        *rp.ENCABEZADOS_SOCIOS_BALANCE_RETENIDA, *rp.ENCABEZADOS_SOCIOS_BALANCE_PESOS_RETENIDA,
        *rp.ENCABEZADOS_RESUMEN_SOCIOS_BALANCE_RETENIDA,
        *rp.ENCABEZADOS_RESUMEN_SOCIOS_BALANCE_PESOS_RETENIDA,
    ):  # fmt: skip
        assert "−" not in t and "→" not in t, t


# ---------------------------------------------------------------------------
# 2) PDF en dólares.
# ---------------------------------------------------------------------------


def test_pdf_usd_cascada_retiene_la_reserva() -> None:
    texto = _texto_pdf(render_reparto_pdf(_req(pesos=None)))
    afac = "(-) Provisión AFAC (100 USD/hr × 3.5 h cobradas)"
    assert (
        f"{afac} $-350.00 {rp.FILA_SALDO_ANTES_RESERVA} {_d(SALDO_TST)} "
        f"{rp.FILA_RESERVA_RETENIDA} {_d(-RESERVA_TST)} {rp.FILA_SALDO_A_REPARTIR} "
        f"{_d(A_REPARTIR_TST)}"
    ) in texto
    # XB-DOS: sin reserva, el saldo a repartir es el saldo.
    assert (
        f"{rp.FILA_PROVISION_AFAC} $-0.00 {rp.FILA_SALDO_ANTES_RESERVA} {_d(SALDO_DOS)} "
        f"{rp.FILA_RESERVA_RETENIDA} $-0.00 {rp.FILA_SALDO_A_REPARTIR} {_d(SALDO_DOS)}"
    ) in texto
    for viejo in (
        "Saldo disponible",
        rp.FILA_RESERVA_INFORMATIVA,
        rp.FILA_RESERVA_DESCONTADA,
        rp.LINEA_RESERVA_RETENIDA.strip(),
        rp.LINEA_RESERVA_RETENIDA_TODOS.strip(),
        "(-) Reserva overhaul",
        "informativa",
        "solo se descuenta a quien lo pidió",
        "(se descuenta)",
        "Reserva a guardar",
    ):
        assert viejo not in texto, viejo


def test_pdf_usd_socios_resumen_global_y_pie() -> None:
    req = _req(pesos=None)
    texto = _texto_pdf(render_reparto_pdf(req))
    assert " ".join(rp.ENCABEZADOS_SOCIOS_BALANCE_RETENIDA) in texto
    assert " ".join(rp.ENCABEZADOS_RESUMEN_SOCIOS_BALANCE_RETENIDA) in texto
    assert f"Socio A 60% {_d(UTIL_A)} {_d(RESERVA_A)} {_d(MONTO_A)}" in texto
    assert f"Socio B 40% {_d(UTIL_B)} {_d(RESERVA_B)} {_d(MONTO_B)}" in texto
    assert f"Socio A 100% {_d(SALDO_DOS)} $0.00 {_d(SALDO_DOS)}" in texto
    assert f"Socio A {_d(RESERVA_A)} {_d(r2(MONTO_A + SALDO_DOS))}" in texto
    assert f"Socio B {_d(RESERVA_B)} {_d(MONTO_B)}" in texto
    # Resumen global: «Saldo a repartir» = Σ saldo a repartir; «Gastos del
    # periodo» mete la reserva para que venta − gastos = saldo a repartir.
    venta = sum(a.ingresos_cobrado_usd for a in req.aviones)
    gastos = sum(
        a.gastos_directos_usd
        + a.gastos_indirectos_usd
        + a.permisos_usd
        + a.otros_usd
        + (a.afac_reservado_usd or 0)
        + a.reserva_overhaul_usd
        + a.comisiones_venta_usd
        for a in req.aviones
    )
    saldo = r2(A_REPARTIR_TST + SALDO_DOS)
    assert r2(venta - gastos) == saldo
    assert f"{_d(venta)} $208.80 {_d(gastos)} {_d(saldo)}" in texto
    assert rp._PIE_CASCADA_BALANCE_RETENIDA.strip() + " Montos en USD." in texto
    assert rp._PIE_CASCADA_BALANCE_TODOS.strip() not in texto
    assert rp._PIE_CASCADA_BALANCE.strip() not in texto


# ---------------------------------------------------------------------------
# 3) PDF en pesos (las dos banderas, como las manda el API).
# ---------------------------------------------------------------------------


def test_pdf_pesos_cascada_retiene_la_reserva() -> None:
    req = _req()
    tst, dos = req.aviones
    texto = _texto_pdf(render_reparto_pdf(req))
    assert (
        f"{rp.FILA_SALDO_ANTES_RESERVA} {_d(tst.saldo_mxn)} {_d(SALDO_TST)} "
        f"{rp.FILA_RESERVA_RETENIDA} {_d(-tst.reserva_overhaul_mxn)} {_d(-RESERVA_TST)} "
        f"en pesos al T.C. oficial del cierre ({K:g}) "
        f"{rp.FILA_SALDO_A_REPARTIR} {_d(tst.saldo_a_repartir_mxn)} {_d(A_REPARTIR_TST)}"
    ) in texto
    assert texto.count("en pesos al T.C. oficial del cierre") == 1  # XB-DOS: reserva 0
    assert " ".join(rp.ENCABEZADOS_SOCIOS_BALANCE_PESOS_RETENIDA) in texto
    assert " ".join(rp.ENCABEZADOS_RESUMEN_SOCIOS_BALANCE_PESOS_RETENIDA) in texto
    assert rp.FILA_RESERVA_DESCONTADA not in texto and "Saldo disponible" not in texto
    # Resumen en pesos: saldo a repartir y gastos con la reserva, cuadrando.
    venta = r2(sum(a.ingresos_cobrado_mxn for a in req.aviones))
    gastos = r2(
        sum(
            a.gastos_directos_mxn
            + a.gastos_indirectos_mxn
            + a.permisos_mxn
            + a.otros_mxn
            + a.afac_reservado_mxn
            + a.reserva_overhaul_mxn
            + a.comisiones_venta_mxn
            for a in req.aviones
        )
    )
    saldo = r2(tst.saldo_a_repartir_mxn + dos.saldo_a_repartir_mxn)
    assert r2(venta - gastos) == saldo
    assert f"{_d(venta)} $208.80 USD {_d(gastos)} {_d(saldo)}" in texto
    assert f"{_d(r2(A_REPARTIR_TST + SALDO_DOS))} USD" in texto
    assert rp._PIE_CASCADA_BALANCE_RETENIDA.strip() in texto


def test_pdf_pesos_sin_pesos_raya_nunca_cero() -> None:
    p = _payload()
    p["aviones"][0]["saldo_a_repartir_mxn"] = None
    texto = _texto_pdf(render_reparto_pdf(RepartoPdfRequest.model_validate(p)))
    assert f"{rp.FILA_SALDO_A_REPARTIR} — {_d(A_REPARTIR_TST)}" in texto
    # El resumen en pesos tampoco inventa el total: «—» (el USD sigue debajo).
    assert " — " in texto and f"{_d(r2(A_REPARTIR_TST + SALDO_DOS))} USD" in texto


# ---------------------------------------------------------------------------
# 4) Sin la llave, o sin la cascada del balance: byte a byte.
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("pesos", [None, True])
def test_sin_la_llave_byte_a_byte(pesos) -> None:
    """Con `saldo_a_repartir_*` en null —o solo el MXN: el discriminador es el
    USD— el PDF y el Excel son los del API 0.0.125 (todos descuentan),
    flujo a flujo y miembro a miembro."""
    base = RepartoPdfRequest.model_validate(tcb._payload_todos(pesos=pesos))
    con_nulos = _payload(pesos=pesos)
    for a in con_nulos["aviones"]:
        a["saldo_a_repartir_usd"] = a["saldo_a_repartir_mxn"] = None
    solo_mxn = _payload(pesos=pesos)
    for a in solo_mxn["aviones"]:
        a["saldo_a_repartir_usd"] = None
    for p in (con_nulos, solo_mxn):
        req = RepartoPdfRequest.model_validate(p)
        assert _flujos(render_reparto_pdf(req)) == _flujos(render_reparto_pdf(base))
        assert _miembros(render_reparto_xlsx(req)) == _miembros(render_reparto_xlsx(base))


@pytest.mark.parametrize("pesos", [None, True])
def test_sin_la_cascada_del_balance_la_llave_se_ignora(pesos) -> None:
    """Sin `reparto_como_balance` el `saldo_*` del API ya trae la reserva
    restada (la cascada de siempre): la llave no cambia nada."""
    p, base = _payload(pesos=pesos), tcb._payload_todos(pesos=pesos)
    for x in (p, base):
        x.pop("reparto_como_balance")
    req, req_base = RepartoPdfRequest.model_validate(p), RepartoPdfRequest.model_validate(base)
    assert _flujos(render_reparto_pdf(req)) == _flujos(render_reparto_pdf(req_base))
    assert _miembros(render_reparto_xlsx(req)) == _miembros(render_reparto_xlsx(req_base))
    texto = _texto_pdf(render_reparto_pdf(req))
    assert "(-) Reserva overhaul" in texto and rp.FILA_SALDO_ANTES_RESERVA not in texto


# ---------------------------------------------------------------------------
# 5) Excel en dólares.
# ---------------------------------------------------------------------------


def test_xlsx_usd_columna_saldo_a_repartir_tras_el_saldo() -> None:
    ws = _hoja(_payload(pesos=None))
    assert [ws.cell(row=4, column=c).value for c in range(11, 17)] == [
        rx.ENCABEZADO_RESERVA_RETENIDA,
        rx.ENCABEZADO_SALDO_ANTES_RESERVA,
        rx.ENCABEZADO_SALDO_A_REPARTIR,
        "Otros ingr. VuelaTour (informativo)",
        "Deuda total del cliente (informativo)",
        rx.ENCABEZADO_PROVISION_AFAC,
    ]
    assert ws["P4"].comment.text == rx.NOTA_ENCABEZADO_AFAC
    assert (ws["K5"].value, ws["L5"].value, ws["M5"].value) == (
        RESERVA_TST,
        SALDO_TST,
        A_REPARTIR_TST,
    )
    assert (ws["K6"].value, ws["L6"].value, ws["M6"].value) == (0.0, SALDO_DOS, SALDO_DOS)
    assert ws["A7"].value == "TOTAL"
    assert (ws["K7"].value, ws["L7"].value, ws["M7"].value) == (
        RESERVA_TST,
        r2(SALDO_TST + SALDO_DOS),
        r2(A_REPARTIR_TST + SALDO_DOS),
    )
    assert ws["M5"].number_format == '"$"#,##0.00' and ws["M7"].font.b
    # La reserva ya no es gris/cursiva: un número como los demás.
    assert not ws["K5"].font.i and not ws["K7"].font.i and ws["K7"].font.b
    assert (ws["N5"].value, ws["O5"].value, ws["P5"].value, ws["P7"].value) == (
        208.8,
        None,
        350.0,
        350.0,
    )
    assert ws["N5"].font.i
    assert ws.column_dimensions["M"].width == 16 and ws.column_dimensions["P"].width == 18
    assert any(str(rg) == "A1:P1" for rg in ws.merged_cells.ranges)
    textos = _textos_col_a(ws)
    assert rx.NOTA_CASCADA_RETENIDA in textos
    assert rx.NOTA_REPARTO_SOCIO_BALANCE_RETENIDA in textos
    for viejo in (
        rx.NOTA_CASCADA_BALANCE,
        rx.NOTA_CASCADA_BALANCE_TODOS,
        rx.NOTA_REPARTO_SOCIO_BALANCE,
        rx.NOTA_REPARTO_SOCIO_BALANCE_TODOS,
    ):
        assert viejo not in textos
    assert not any("informativa" in t or "no resta" in t or "no se resta" in t for t in textos)


def test_xlsx_usd_reparto_por_socio_reserva_retenida() -> None:
    ws = _hoja(_payload(pesos=None))
    t = _fila(ws, "Reparto por socio")
    assert [ws.cell(row=t + 1, column=c).value for c in range(1, 8)] == [
        "Avión", "Socio", "Porcentaje", "Monto USD (a recibir)",
        "Utilidad USD", "Reserva retenida USD", "¿Se descuenta?",
    ]  # fmt: skip
    filas = [[ws.cell(row=t + 1 + i, column=c).value for c in range(1, 8)] for i in (1, 2, 3)]
    assert filas == [
        ["XB-TST", "Socio A", 0.6, MONTO_A, UTIL_A, RESERVA_A, "Sí"],
        ["XB-TST", "Socio B", 0.4, MONTO_B, UTIL_B, RESERVA_B, "Sí"],
        ["XB-DOS", "Socio A", 1.0, SALDO_DOS, SALDO_DOS, 0.0, "Sí"],
    ]
    assert ws.cell(row=t + 5, column=1).value == rx.NOTA_REPARTO_SOCIO_BALANCE_RETENIDA


def test_xlsx_usd_con_combustible_y_otros_ingresos_corre_junto() -> None:
    """Con el combustible separado (`oc` = 3, provisión AFAC en el bloque) y
    otros ingresos del avión, «Saldo a repartir» sigue justo tras el saldo y
    las notas dicen que la reserva se retiene."""
    a = _n621tx_retenido(pesos=False, otros_ingresos=100.0)["aviones"][0]
    assert a["saldo_a_repartir_usd"] == r2(tcs.SALDO + 100.0 - tcs.RESERVA)
    ws = _hoja(_n621tx_retenido(pesos=False, otros_ingresos=100.0))
    assert [ws.cell(row=4, column=c).value for c in range(7, 20)] == [
        rc.ENCABEZADO_GASTOS_DIRECTOS_SIN_COMBUSTIBLE, rx.ENCABEZADO_PROVISION_AFAC,
        rc.ENCABEZADO_UTILIDAD_ANTES_GASTOS, rc.ENCABEZADO_COMBUSTIBLE,
        "Gastos indirectos", "Permisos", "Fijos (prorrateo + reparto)",
        rx.ENCABEZADO_RESERVA_RETENIDA, rx.ENCABEZADO_SALDO_ANTES_RESERVA,
        rx.ENCABEZADO_SALDO_A_REPARTIR, "Otros ingr. VuelaTour (informativo)",
        "Deuda total del cliente (informativo)", rx.ENCABEZADO_OTROS_INGRESOS_AVION,
    ]  # fmt: skip
    assert ws.cell(row=4, column=20).value is None
    assert (ws["N5"].value, ws["O5"].value, ws["P5"].value, ws["S5"].value) == (
        tcs.RESERVA, r2(tcs.SALDO + 100.0), r2(tcs.SALDO + 100.0 - tcs.RESERVA), 100.0,
    )  # fmt: skip
    assert ws["P6"].value == ws["P5"].value and ws["A6"].value == "TOTAL"
    assert not ws["N5"].font.i
    assert ws.column_dimensions["P"].width == 16 and ws.column_dimensions["S"].width == 18
    assert any(str(rg) == "A1:S1" for rg in ws.merged_cells.ranges)
    textos = _textos_col_a(ws)
    assert rx.NOTA_OTROS_INGRESOS_AVION_RETENIDA in textos
    assert rc.NOTA_COMBUSTIBLE_SEPARADO_RETENIDA in textos
    assert rc.NOTA_COMBUSTIBLE_SEPARADO not in textos
    assert rx.NOTA_OTROS_INGRESOS_AVION_BALANCE_TODOS not in textos
    assert "Saldo disponible" not in rc.NOTA_COMBUSTIBLE_SEPARADO_RETENIDA


# ---------------------------------------------------------------------------
# 6) Excel en pesos.
# ---------------------------------------------------------------------------


def test_xlsx_pesos_columnas_totales_y_tabla_usd_informativa() -> None:
    p = _payload()
    ws = _hoja(p)
    enc = rx.encabezados_pesos(False, True, todos_descuentan=True, retenida=True)
    assert [ws.cell(row=4, column=c).value for c in range(1, 17)] == enc
    assert enc[10:13] == [
        rx.ENCABEZADO_RESERVA_RETENIDA_PESOS,
        rx.ENCABEZADO_SALDO_ANTES_RESERVA_PESOS,
        rx.ENCABEZADO_SALDO_A_REPARTIR_PESOS,
    ]
    assert enc[13:] == [
        "Otros ingr. VuelaTour (USD, informativo)",
        "Deuda total del cliente (USD, informativo)",
        rx.ENCABEZADO_PROVISION_AFAC_PESOS,
    ]
    tst, dos = p["aviones"]
    assert (ws["K5"].value, ws["L5"].value, ws["M5"].value) == (
        tst["reserva_overhaul_mxn"], tst["saldo_mxn"], tst["saldo_a_repartir_mxn"],
    )  # fmt: skip
    assert ws["M5"].font.b and ws["L5"].font.b and not ws["K5"].font.i
    assert ws["K5"].comment.text.endswith(f"{K:g}.")
    assert ws["M7"].value == r2(tst["saldo_a_repartir_mxn"] + dos["saldo_a_repartir_mxn"])
    assert ws["M7"].font.b and not ws["K7"].font.i and ws["K7"].font.b
    assert (ws["P5"].value, ws["P7"].value) == (6037.5, 6037.5)
    assert ws["N4"].font.i and ws["N5"].font.i
    assert ws.column_dimensions["M"].width == 18 and ws.column_dimensions["P"].width == 20
    assert any(str(rg) == "A1:P1" for rg in ws.merged_cells.ranges)
    # La tabla informativa en dólares lleva la misma columna.
    t = _fila(ws, rx.TITULO_USD_INFORMATIVO)
    assert [ws.cell(row=t + 1, column=c).value for c in (11, 12, 13, 16)] == [
        rx.ENCABEZADO_RESERVA_RETENIDA,
        rx.ENCABEZADO_SALDO_ANTES_RESERVA,
        rx.ENCABEZADO_SALDO_A_REPARTIR,
        rx.ENCABEZADO_PROVISION_AFAC,
    ]
    assert (ws.cell(row=t + 2, column=13).value, ws.cell(row=t + 4, column=13).value) == (
        A_REPARTIR_TST,
        r2(A_REPARTIR_TST + SALDO_DOS),
    )
    textos = _textos_col_a(ws)
    assert rx.NOTA_CASCADA_RETENIDA in textos and rx.NOTA_REPARTO_SOCIO_BALANCE_RETENIDA in textos
    assert any(t.startswith("Reserva overhaul pasada a pesos al T.C.") for t in textos)
    t = _fila(ws, "Reparto por socio (en pesos)")
    assert [ws.cell(row=t + 1, column=c).value for c in (8, 9)] == [
        "Reserva retenida MXN",
        "Reserva retenida USD (informativo)",
    ]
    a = tst["reparto"][0]
    assert [ws.cell(row=t + 2, column=c).value for c in (4, 5, 8, 10)] == [
        a["monto_mxn"], MONTO_A, a["reserva_mxn"], "Sí",
    ]  # fmt: skip
    # Sin pesos en el saldo a repartir: «—», jamás un 0 (ni en el total).
    p["aviones"][0]["saldo_a_repartir_mxn"] = None
    ws = _hoja(p)
    assert ws["M5"].value == rx.SIN_PESOS and ws["M7"].value == rx.SIN_PESOS


def test_xlsx_pesos_con_combustible_separado() -> None:
    ws = _hoja(_n621tx_retenido(pesos=True))
    enc = rx.encabezados_pesos(False, True, True, todos_descuentan=True, retenida=True)
    assert [ws.cell(row=4, column=c).value for c in range(1, 19)] == enc
    assert len(enc) == 18 and ws.cell(row=4, column=19).value is None
    assert enc[13:16] == [
        rx.ENCABEZADO_RESERVA_RETENIDA_PESOS,
        rx.ENCABEZADO_SALDO_ANTES_RESERVA_PESOS,
        rx.ENCABEZADO_SALDO_A_REPARTIR_PESOS,
    ]
    a = _n621tx_retenido(pesos=True)["aviones"][0]
    assert (ws["O5"].value, ws["P5"].value) == (a["saldo_mxn"], a["saldo_a_repartir_mxn"])
    assert ws["P6"].value == a["saldo_a_repartir_mxn"]
    assert ws["Q4"].value == "Otros ingr. VuelaTour (USD, informativo)" and ws["Q4"].font.i
    assert rc.NOTA_COMBUSTIBLE_SEPARADO_RETENIDA in _textos_col_a(ws)


# ---------------------------------------------------------------------------
# Rutas: el JSON del API 0.0.128 entra sin 422.
# ---------------------------------------------------------------------------

TOKEN = "tok-reserva-retenida-reparto"
client = TestClient(app)


def test_rutas_aceptan_el_payload_0128(monkeypatch, request) -> None:
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    request.addfinalizer(get_settings.cache_clear)
    headers = {"X-Internal-Token": TOKEN}
    for p in (_payload(), _payload(pesos=None)):
        res = client.post("/pdf/reparto", json=p, headers=headers)
        assert res.status_code == 200, res.text
        texto = _texto_pdf(res.content)
        assert rp.FILA_RESERVA_RETENIDA in texto and rp.FILA_SALDO_ANTES_RESERVA in texto
        res = client.post("/pdf/reparto-xlsx", json=p, headers=headers)
        assert res.status_code == 200, res.text
        ws = load_workbook(BytesIO(res.content)).active
        assert ws["M4"].value in (
            rx.ENCABEZADO_SALDO_A_REPARTIR, rx.ENCABEZADO_SALDO_A_REPARTIR_PESOS,
        )  # fmt: skip
