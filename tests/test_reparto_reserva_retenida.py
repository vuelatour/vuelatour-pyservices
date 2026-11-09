"""La RESERVA DE OVERHAUL RESTA EN LA CASCADA del reparto a socios, PDF y Excel
(9-nov-2026, API 0.0.129, invariante 125 del API).

Decisión del cliente: «Necesito resolver ese tema del Overhaul, que quede ANTES
del total», como en su Excel manual. Con la llave booleana
`reserva_overhaul_en_cascada: true` (el DISCRIMINADOR único) y la cascada del
balance (`reparto_como_balance`), `saldo_*` YA RESTA la reserva y los
documentos la pintan como UNA RESTA MÁS: «(-) Reserva de overhaul (se retiene
para el overhaul)» justo después de «(-) Fijos» y antes de «(+) Otros ingresos
del avión», y la fila resaltada es «Saldo a repartir» (= `saldo_*`); sin «Saldo
antes de la reserva» ni filas grises informativas; socios y resumen «Reserva
retenida»; el resumen global suma `saldo_*` y sus gastos incluyen la reserva
(venta + otros ingresos − gastos = saldo a repartir). En el Excel la columna
de la reserva se queda en su lugar en número normal y la del saldo se rotula
«Saldo a repartir»: ninguna columna nueva (el 0.0.128 metía una; el cliente lo
rechazó). `reparto[].monto_*` llega ya neto (Σ = `saldo_*` al centavo). Sin la
llave (API ≤ 0.0.127), con las llaves DEPRECADAS `saldo_a_repartir_*` sin
ella, o sin la cascada del balance: PDF y Excel BYTE A BYTE los de siempre
(`test_reparto_como_balance`, bloque «RESERVA PARA TODOS»).
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
LLAVE = "reserva_overhaul_en_cascada"
SALDO_ANTES_TST, SALDO_DOS = tcb.SALDO_TST, tcb.SALDO_DOS
RESERVA_TST = 140.0
SALDO_TST = 914.30  # 1,054.30 − 140.00: el `saldo_usd` que manda el API 0.0.129
UTIL_A, RESERVA_A, MONTO_A = tcb.UTIL_A, tcb.RESERVA_A, tcb.MONTO_A
UTIL_B, RESERVA_B = tcb.UTIL_B, tcb.RESERVA_B
MONTO_B = r2(UTIL_B - RESERVA_B)
_flujos, _miembros, _fila, _textos_col_a = tcb._flujos, tcb._miembros, tcb._fila, tcb._textos_col_a
FIJOS = "(-) Fijos (prorrateo + reparto manual)"
AFAC_TST = "(-) Provisión AFAC (100 USD/hr × 3.5 h cobradas)"
# Textos EXACTOS del contrato (los necesita el panel).
FILA_RESERVA = "(-) Reserva de overhaul (se retiene para el overhaul)"
FILA_SALDO = "Saldo a repartir"


# ---------------------------------------------------------------------------
# Payloads: el reparto donde todos descuentan (API 0.0.125) con la llave del
# 0.0.129, espejo de la aritmética del API: `saldo_*` ya neto de la reserva;
# los socios ya venían netos (utilidad − reserva, todos descuentan).
# ---------------------------------------------------------------------------


def con_cascada(p: dict) -> dict:
    p = copy.deepcopy(p)
    for a in p["aviones"]:
        a[LLAVE] = True
        a["saldo_usd"] = r2(a["saldo_usd"] - a["reserva_overhaul_usd"])
        a["saldo_mxn"] = (
            None
            if a.get("saldo_mxn") is None or a.get("reserva_overhaul_mxn") is None
            else r2(a["saldo_mxn"] - a["reserva_overhaul_mxn"])
        )
    return p


def _payload(*, pesos: bool | None = True) -> dict:
    return con_cascada(tcb._payload_todos(pesos=pesos))


def _req(**kw) -> RepartoPdfRequest:
    return RepartoPdfRequest.model_validate(_payload(**kw))


def _hoja(p: dict):
    return load_workbook(BytesIO(render_reparto_xlsx(RepartoPdfRequest.model_validate(p)))).active


def _n621tx(*, pesos: bool, otros_ingresos: float = 0.0) -> dict:
    """El caso real N621TX del combustible separado con TODOS descontando y
    la reserva en la cascada (saldo = saldo de antes − 15,000)."""
    p = tcs._payload(pesos=pesos, otros_ingresos=otros_ingresos)
    for a in p["aviones"]:
        for s in a["reparto"]:
            s["descuenta_reserva"] = True
            s["monto_usd"] = r2(s["utilidad_usd"] - s["reserva_usd"])
            s["monto_mxn"] = r2(s["utilidad_mxn"] - s["reserva_mxn"])
        a["reserva_retenida_usd"] = a["reserva_overhaul_usd"]
        a["reserva_retenida_mxn"] = a["reserva_overhaul_mxn"]
    return con_cascada(p)


# ---------------------------------------------------------------------------
# 1) Números, esquema y helpers.
# ---------------------------------------------------------------------------


def test_los_numeros_del_payload_cuadran() -> None:
    req = _req()
    tst, dos = req.aviones
    assert tst.saldo_usd == SALDO_TST == r2(SALDO_ANTES_TST - RESERVA_TST)
    antes = RepartoPdfRequest.model_validate(tcb._payload_todos()).aviones[0]
    assert tst.saldo_mxn == r2(antes.saldo_mxn - antes.reserva_overhaul_mxn)
    # Σ de lo que reciben los socios = el saldo del API, al centavo.
    assert r2(sum(s.monto_usd for s in tst.reparto)) == SALDO_TST
    assert r2(sum(s.monto_mxn for s in tst.reparto)) == tst.saldo_mxn
    assert tst.reserva_retenida_usd == RESERVA_TST == tst.reserva_overhaul_usd
    assert (MONTO_A, MONTO_B) == (548.58, 365.72)
    assert dos.saldo_usd == SALDO_DOS and dos.reserva_overhaul_usd == 0


def test_esquema_aditivo_validador_liberal_y_helpers() -> None:
    a = RepartoAvion(matricula="X", modelo="M", ingresos_cobrado_usd=0, saldo_usd=10)
    assert a.reserva_overhaul_en_cascada is None and not rr.reserva_en_cascada(a)
    for v in (None, False, "true", 1, 0, {}):
        b = RepartoAvion.model_validate(
            {"matricula": "X", "modelo": "M", "ingresos_cobrado_usd": 0, "saldo_usd": 10, LLAVE: v}
        )
        assert b.reserva_overhaul_en_cascada is None and not rr.reserva_en_cascada(b), v
    a.reserva_overhaul_en_cascada = True
    assert rr.reserva_en_cascada(a)
    # Las llaves DEPRECADAS del 0.0.128 entran sin 422 y no encienden nada.
    viejo = RepartoAvion.model_validate(
        {
            "matricula": "X", "modelo": "M", "ingresos_cobrado_usd": 0, "saldo_usd": 10,
            "saldo_a_repartir_usd": 8.0, "saldo_a_repartir_mxn": 140.0,
        }
    )  # fmt: skip
    assert viejo.saldo_a_repartir_usd == 8.0 and not rr.reserva_en_cascada(viejo)
    assert rr.reserva_en_cascada_req(_req())
    assert not rr.reserva_en_cascada_req(RepartoPdfRequest.model_validate(tcb._payload_todos()))
    vacio = RepartoPdfRequest.model_validate({**_payload(), "aviones": []})
    assert not rr.reserva_en_cascada_req(vacio)


def test_etiquetas_caben_y_sin_caracteres_que_helvetica_no_dibuja() -> None:
    assert rp.FILA_RESERVA_EN_CASCADA == FILA_RESERVA and rp.FILA_SALDO_A_REPARTIR == FILA_SALDO
    ancho = 112 * mm - 12
    for t in (rp.FILA_RESERVA_EN_CASCADA, rp.FILA_SALDO_A_REPARTIR):
        assert stringWidth(t, "Helvetica-Bold", 9) <= ancho, t
    for t in (
        rp.FILA_RESERVA_EN_CASCADA, rp.FILA_SALDO_A_REPARTIR,
        rp._PIE_CASCADA_BALANCE_RESERVA_EN_CASCADA,
        *rp.ENCABEZADOS_SOCIOS_BALANCE_RETENIDA, *rp.ENCABEZADOS_SOCIOS_BALANCE_PESOS_RETENIDA,
        *rp.ENCABEZADOS_RESUMEN_SOCIOS_BALANCE_RETENIDA,
        *rp.ENCABEZADOS_RESUMEN_SOCIOS_BALANCE_PESOS_RETENIDA,
    ):  # fmt: skip
        assert "−" not in t and "→" not in t, t
    assert "- reserva de overhaul + otros ingresos" in rp._PIE_CASCADA_BALANCE_RESERVA_EN_CASCADA
    assert rx.ENCABEZADO_SALDO_A_REPARTIR == "Saldo a repartir"
    assert rx.ENCABEZADO_SALDO_A_REPARTIR_PESOS == "Saldo a repartir (MXN)"
    assert rx.ENCABEZADO_RESERVA_RETENIDA == "Reserva overhaul (se retiene)"
    assert rx.ENCABEZADO_RESERVA_RETENIDA_PESOS == "Reserva overhaul (MXN, se retiene)"
    for t in (
        rx.NOTA_CASCADA_RETENIDA,
        rx.NOTA_OTROS_INGRESOS_AVION_RETENIDA,
        rc.NOTA_COMBUSTIBLE_SEPARADO_RESERVA_EN_CASCADA,
    ):
        assert "− reserva de overhaul + otros ingresos del avión" in t, t
        assert "Saldo antes" not in t and "informativa" not in t


# ---------------------------------------------------------------------------
# 2) PDF en dólares.
# ---------------------------------------------------------------------------


def test_pdf_usd_la_reserva_resta_tras_los_fijos_y_el_saldo_es_a_repartir() -> None:
    texto = _texto_pdf(render_reparto_pdf(_req(pesos=None)))
    # XB-TST: fijos → reserva → provisión AFAC → Saldo a repartir (= saldo del API).
    assert (
        f"{FIJOS} $-0.00 {FILA_RESERVA} {_d(-RESERVA_TST)} {AFAC_TST} $-350.00 "
        f"{FILA_SALDO} {_d(SALDO_TST)}"
    ) in texto
    # XB-DOS: sin reserva, la fila dice $-0.00 y el saldo es el de siempre.
    assert (
        f"{FIJOS} $-0.00 {FILA_RESERVA} $-0.00 {rp.FILA_PROVISION_AFAC} $-0.00 "
        f"{FILA_SALDO} {_d(SALDO_DOS)}"
    ) in texto
    for viejo in (
        "Saldo disponible",
        "Saldo antes de la reserva",
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
    # Resumen global: «Saldo a repartir» = Σ saldo del API; «Gastos del
    # periodo» incluye la reserva para que venta − gastos = saldo a repartir.
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
    saldo = r2(SALDO_TST + SALDO_DOS)
    assert r2(venta - gastos) == saldo == r2(sum(a.saldo_usd for a in req.aviones))
    assert f"{_d(venta)} $208.80 {_d(gastos)} {_d(saldo)}" in texto
    assert rp._PIE_CASCADA_BALANCE_RESERVA_EN_CASCADA.strip() + " Montos en USD." in texto
    assert rp._PIE_CASCADA_BALANCE_TODOS.strip() not in texto
    assert rp._PIE_CASCADA_BALANCE.strip() not in texto


def test_pdf_usd_con_combustible_separado_la_reserva_va_antes_de_otros_ingresos() -> None:
    """N621TX (orden del Excel del cliente, combustible aparte): fijos →
    reserva → (+) otros ingresos del avión → saldo a repartir."""
    p = _n621tx(pesos=False, otros_ingresos=100.0)
    a = p["aviones"][0]
    saldo = r2(tcs.SALDO + 100.0 - tcs.RESERVA)
    assert a["saldo_usd"] == saldo
    texto = _texto_pdf(render_reparto_pdf(RepartoPdfRequest.model_validate(p)))
    assert (
        f"{tcs.FIJOS_PDF} {FILA_RESERVA} {_d(-tcs.RESERVA)} {rp.FILA_OTROS_INGRESOS_AVION} $100.00"
    ) in texto
    assert f"{FILA_SALDO} {_d(saldo)}" in texto
    assert "Saldo disponible" not in texto and "Saldo antes" not in texto
    # La reserva no se cuela antes de la utilidad antes de gastos.
    assert texto.index(rc.FILA_UTILIDAD_ANTES_GASTOS) < texto.index(FILA_RESERVA)


# ---------------------------------------------------------------------------
# 3) PDF en pesos (las dos banderas, como las manda el API).
# ---------------------------------------------------------------------------


def test_pdf_pesos_la_reserva_resta_con_su_tc_y_el_saldo_es_a_repartir() -> None:
    req = _req()
    tst, dos = req.aviones
    texto = _texto_pdf(render_reparto_pdf(req))
    assert (
        f"{FIJOS} $-0.00 $-0.00 {FILA_RESERVA} {_d(-tst.reserva_overhaul_mxn)} "
        f"{_d(-RESERVA_TST)} en pesos al T.C. oficial del cierre ({K:g}) {AFAC_TST} "
        f"{_d(-tst.afac_reservado_mxn)} $-350.00 {FILA_SALDO} {_d(tst.saldo_mxn)} {_d(SALDO_TST)}"
    ) in texto
    assert texto.count("en pesos al T.C. oficial del cierre") == 1  # XB-DOS: reserva 0
    assert f"{FILA_RESERVA} $-0.00 $-0.00" in texto  # XB-DOS
    assert " ".join(rp.ENCABEZADOS_SOCIOS_BALANCE_PESOS_RETENIDA) in texto
    assert " ".join(rp.ENCABEZADOS_RESUMEN_SOCIOS_BALANCE_PESOS_RETENIDA) in texto
    for viejo in (rp.FILA_RESERVA_DESCONTADA, "Saldo disponible", "Saldo antes"):
        assert viejo not in texto, viejo
    # Resumen en pesos: Σ saldo del API y gastos con la reserva, cuadrando.
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
    saldo = r2(tst.saldo_mxn + dos.saldo_mxn)
    assert r2(venta - gastos) == saldo
    assert f"{_d(venta)} $208.80 USD {_d(gastos)} {_d(saldo)}" in texto
    assert f"{_d(r2(SALDO_TST + SALDO_DOS))} USD" in texto
    assert rp._PIE_CASCADA_BALANCE_RESERVA_EN_CASCADA.strip() in texto


def test_pdf_pesos_sin_pesos_raya_nunca_cero() -> None:
    p = _payload()
    p["aviones"][0]["saldo_mxn"] = None
    texto = _texto_pdf(render_reparto_pdf(RepartoPdfRequest.model_validate(p)))
    assert f"{FILA_SALDO} — {_d(SALDO_TST)}" in texto
    # El resumen en pesos tampoco inventa el total: «—» (el USD sigue debajo).
    assert " — " in texto and f"{_d(r2(SALDO_TST + SALDO_DOS))} USD" in texto


# ---------------------------------------------------------------------------
# 4) Sin la llave, con las deprecadas del 0.0.128, o sin la cascada del
#    balance: byte a byte.
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("pesos", [None, True])
@pytest.mark.parametrize("forma", ["null", "false", "texto", "deprecadas-0128"])
def test_sin_la_llave_byte_a_byte(pesos, forma) -> None:
    """Con la llave en null/false/texto —o con las llaves DEPRECADAS
    `saldo_a_repartir_*` del 0.0.128 y sin ella— el PDF y el Excel son los del
    API 0.0.125 (todos descuentan), flujo a flujo y miembro a miembro."""
    base = RepartoPdfRequest.model_validate(tcb._payload_todos(pesos=pesos))
    p = tcb._payload_todos(pesos=pesos)
    for a in p["aviones"]:
        if forma == "deprecadas-0128":
            a["saldo_a_repartir_usd"] = r2(a["saldo_usd"] - a["reserva_overhaul_usd"])
            a["saldo_a_repartir_mxn"] = 123.45
        else:
            a[LLAVE] = {"null": None, "false": False, "texto": "true"}[forma]
    req = RepartoPdfRequest.model_validate(p)
    assert _flujos(render_reparto_pdf(req)) == _flujos(render_reparto_pdf(base))
    assert _miembros(render_reparto_xlsx(req)) == _miembros(render_reparto_xlsx(base))


@pytest.mark.parametrize("pesos", [None, True])
def test_sin_la_cascada_del_balance_la_llave_se_ignora(pesos) -> None:
    """Sin `reparto_como_balance` el `saldo_*` del API siempre ha restado la
    reserva y la cascada de siempre ya la pinta como «(-) Reserva overhaul»:
    la llave no cambia nada."""
    p = _payload(pesos=pesos)
    base = copy.deepcopy(p)
    for x in (p, base):
        x.pop("reparto_como_balance")
    for a in base["aviones"]:
        a[LLAVE] = None
    req, req_base = RepartoPdfRequest.model_validate(p), RepartoPdfRequest.model_validate(base)
    assert _flujos(render_reparto_pdf(req)) == _flujos(render_reparto_pdf(req_base))
    assert _miembros(render_reparto_xlsx(req)) == _miembros(render_reparto_xlsx(req_base))
    texto = _texto_pdf(render_reparto_pdf(req))
    assert "(-) Reserva overhaul" in texto and FILA_RESERVA not in texto


# ---------------------------------------------------------------------------
# 5) Excel en dólares.
# ---------------------------------------------------------------------------


def test_xlsx_usd_mismas_columnas_reserva_normal_y_saldo_a_repartir() -> None:
    ws = _hoja(_payload(pesos=None))
    assert [ws.cell(row=4, column=c).value for c in range(11, 17)] == [
        rx.ENCABEZADO_RESERVA_RETENIDA,
        rx.ENCABEZADO_SALDO_A_REPARTIR,
        "Otros ingr. VuelaTour (informativo)",
        "Deuda total del cliente (informativo)",
        rx.ENCABEZADO_PROVISION_AFAC,
        None,  # ninguna columna nueva: 15 columnas, como el 0.0.125
    ]
    assert ws["O4"].comment.text == rx.NOTA_ENCABEZADO_AFAC
    assert (ws["K5"].value, ws["L5"].value) == (RESERVA_TST, SALDO_TST)
    assert (ws["K6"].value, ws["L6"].value) == (0.0, SALDO_DOS)
    assert ws["A7"].value == "TOTAL"
    assert (ws["K7"].value, ws["L7"].value) == (RESERVA_TST, r2(SALDO_TST + SALDO_DOS))
    assert ws["L5"].number_format == '"$"#,##0.00' and ws["L7"].font.b
    # La reserva ya no es gris/cursiva: un número como los demás, con total.
    assert not ws["K5"].font.i and not ws["K7"].font.i and ws["K7"].font.b
    assert (ws["M5"].value, ws["N5"].value, ws["O5"].value, ws["O7"].value) == (
        208.8,
        None,
        350.0,
        350.0,
    )
    assert ws["M5"].font.i
    assert ws.column_dimensions["L"].width == 16 and ws.column_dimensions["O"].width == 18
    assert any(str(rg) == "A1:O1" for rg in ws.merged_cells.ranges)
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
    assert not any("Saldo antes" in t for t in textos)


def test_xlsx_usd_layout_identico_al_0125_salvo_rotulos_notas_y_fuente() -> None:
    """Mismas columnas, anchos, merges y números que el libro donde todos
    descuentan (0.0.125): solo cambian el saldo (ya neto), los rótulos de K y
    L, las notas y la fuente de la reserva."""
    con = _hoja(_payload(pesos=None))
    sin = _hoja(tcb._payload_todos(pesos=None))
    assert con.max_column == sin.max_column and con.max_row == sin.max_row
    assert [c.width for c in con.column_dimensions.values()] == [
        c.width for c in sin.column_dimensions.values()
    ]
    assert {str(r) for r in con.merged_cells.ranges} == {str(r) for r in sin.merged_cells.ranges}
    for fila in (5, 6, 7):
        for col in range(1, 16):
            if col == 12:  # el saldo: ahora neto de la reserva
                continue
            assert con.cell(row=fila, column=col).value == sin.cell(row=fila, column=col).value


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


def test_xlsx_usd_con_combustible_y_otros_ingresos_sin_corrimiento_extra() -> None:
    """Con el combustible separado (`oc` = 3, provisión AFAC en el bloque) y
    otros ingresos del avión, la reserva y el saldo quedan en N y O (solo
    cambia el rótulo) y las notas dicen que la reserva resta en la cascada."""
    p = _n621tx(pesos=False, otros_ingresos=100.0)
    a = p["aviones"][0]
    assert a["saldo_usd"] == r2(tcs.SALDO + 100.0 - tcs.RESERVA)
    ws = _hoja(p)
    assert [ws.cell(row=4, column=c).value for c in range(7, 19)] == [
        rc.ENCABEZADO_GASTOS_DIRECTOS_SIN_COMBUSTIBLE, rx.ENCABEZADO_PROVISION_AFAC,
        rc.ENCABEZADO_UTILIDAD_ANTES_GASTOS, rc.ENCABEZADO_COMBUSTIBLE,
        "Gastos indirectos", "Permisos", "Fijos (prorrateo + reparto)",
        rx.ENCABEZADO_RESERVA_RETENIDA, rx.ENCABEZADO_SALDO_A_REPARTIR,
        "Otros ingr. VuelaTour (informativo)", "Deuda total del cliente (informativo)",
        rx.ENCABEZADO_OTROS_INGRESOS_AVION,
    ]  # fmt: skip
    assert ws.cell(row=4, column=19).value is None
    assert (ws["N5"].value, ws["O5"].value, ws["R5"].value) == (
        tcs.RESERVA, r2(tcs.SALDO + 100.0 - tcs.RESERVA), 100.0,
    )  # fmt: skip
    assert ws["O6"].value == ws["O5"].value and ws["A6"].value == "TOTAL"
    assert not ws["N5"].font.i and ws["O5"].number_format == '"$"#,##0.00'
    assert any(str(rg) == "A1:R1" for rg in ws.merged_cells.ranges)
    textos = _textos_col_a(ws)
    assert rx.NOTA_OTROS_INGRESOS_AVION_RETENIDA in textos
    assert rc.NOTA_COMBUSTIBLE_SEPARADO_RESERVA_EN_CASCADA in textos
    assert rc.NOTA_COMBUSTIBLE_SEPARADO not in textos
    assert rx.NOTA_OTROS_INGRESOS_AVION_BALANCE_TODOS not in textos
    assert "Saldo disponible" not in rc.NOTA_COMBUSTIBLE_SEPARADO_RESERVA_EN_CASCADA


# ---------------------------------------------------------------------------
# 6) Excel en pesos.
# ---------------------------------------------------------------------------


def test_xlsx_pesos_columnas_totales_y_tabla_usd_informativa() -> None:
    p = _payload()
    ws = _hoja(p)
    enc = rx.encabezados_pesos(False, True, todos_descuentan=True, retenida=True)
    assert len(enc) == 15
    assert [ws.cell(row=4, column=c).value for c in range(1, 16)] == enc
    assert ws.cell(row=4, column=16).value is None
    assert enc[10:] == [
        rx.ENCABEZADO_RESERVA_RETENIDA_PESOS,
        rx.ENCABEZADO_SALDO_A_REPARTIR_PESOS,
        "Otros ingr. VuelaTour (USD, informativo)",
        "Deuda total del cliente (USD, informativo)",
        rx.ENCABEZADO_PROVISION_AFAC_PESOS,
    ]
    tst, dos = p["aviones"]
    assert (ws["K5"].value, ws["L5"].value) == (tst["reserva_overhaul_mxn"], tst["saldo_mxn"])
    assert ws["L5"].font.b and not ws["K5"].font.i
    assert ws["K5"].comment.text.endswith(f"{K:g}.")
    assert ws["L7"].value == r2(tst["saldo_mxn"] + dos["saldo_mxn"])
    assert ws["L7"].font.b and not ws["K7"].font.i and ws["K7"].font.b
    assert (ws["O5"].value, ws["O7"].value) == (6037.5, 6037.5)
    assert ws["M4"].font.i and ws["M5"].font.i
    assert ws.column_dimensions["L"].width == 18 and ws.column_dimensions["O"].width == 20
    assert any(str(rg) == "A1:O1" for rg in ws.merged_cells.ranges)
    # La tabla informativa en dólares lleva los mismos rótulos.
    t = _fila(ws, rx.TITULO_USD_INFORMATIVO)
    assert [ws.cell(row=t + 1, column=c).value for c in (11, 12, 15)] == [
        rx.ENCABEZADO_RESERVA_RETENIDA,
        rx.ENCABEZADO_SALDO_A_REPARTIR,
        rx.ENCABEZADO_PROVISION_AFAC,
    ]
    assert (ws.cell(row=t + 2, column=12).value, ws.cell(row=t + 4, column=12).value) == (
        SALDO_TST,
        r2(SALDO_TST + SALDO_DOS),
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
    # Sin pesos en el saldo: «—», jamás un 0 (ni en el total).
    p["aviones"][0]["saldo_mxn"] = None
    ws = _hoja(p)
    assert ws["L5"].value == rx.SIN_PESOS and ws["L7"].value == rx.SIN_PESOS


def test_xlsx_pesos_con_combustible_separado() -> None:
    ws = _hoja(_n621tx(pesos=True))
    enc = rx.encabezados_pesos(False, True, True, todos_descuentan=True, retenida=True)
    assert [ws.cell(row=4, column=c).value for c in range(1, 18)] == enc
    assert len(enc) == 17 and ws.cell(row=4, column=18).value is None
    assert enc[13:15] == [
        rx.ENCABEZADO_RESERVA_RETENIDA_PESOS,
        rx.ENCABEZADO_SALDO_A_REPARTIR_PESOS,
    ]
    a = _n621tx(pesos=True)["aviones"][0]
    assert (ws["N5"].value, ws["O5"].value) == (a["reserva_overhaul_mxn"], a["saldo_mxn"])
    assert ws["O6"].value == a["saldo_mxn"] and ws["O5"].font.b
    assert ws["P4"].value == "Otros ingr. VuelaTour (USD, informativo)" and ws["P4"].font.i
    assert rc.NOTA_COMBUSTIBLE_SEPARADO_RESERVA_EN_CASCADA in _textos_col_a(ws)


# ---------------------------------------------------------------------------
# Rutas: el JSON del API 0.0.129 entra sin 422.
# ---------------------------------------------------------------------------

TOKEN = "tok-reserva-en-cascada-reparto"
client = TestClient(app)


def test_rutas_aceptan_el_payload_0129(monkeypatch, request) -> None:
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    request.addfinalizer(get_settings.cache_clear)
    headers = {"X-Internal-Token": TOKEN}
    for p in (_payload(), _payload(pesos=None)):
        res = client.post("/pdf/reparto", json=p, headers=headers)
        assert res.status_code == 200, res.text
        texto = _texto_pdf(res.content)
        assert FILA_RESERVA in texto and FILA_SALDO in texto
        res = client.post("/pdf/reparto-xlsx", json=p, headers=headers)
        assert res.status_code == 200, res.text
        ws = load_workbook(BytesIO(res.content)).active
        assert ws["L4"].value in (
            rx.ENCABEZADO_SALDO_A_REPARTIR, rx.ENCABEZADO_SALDO_A_REPARTIR_PESOS,
        )  # fmt: skip
