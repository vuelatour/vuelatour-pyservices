"""Cierre de mes POR AVIÓN en los documentos (23/10/2026, API 0.0.104,
invariante 105 del API).

Pedido del cliente: «poner un botón de cierre de mes individual en cada
avión». El API cierra cada avión por separado y la fila global del mes es
DERIVADA; hacia aquí llegan llaves ADITIVAS que solo se PINTAN:

- `cierre.estado` puede ser 'PARCIAL' (algún avión cerrado, la flota activa
  incompleta) con `cierre.aviones_cerrados` (matrículas) y, si el API lo
  manda, `cierre.aviones_activos` («2 de 7 aviones»).
- Reparto: `aviones[].cierre_avion` (estado, sello y `coincide_con_snapshot`
  de ESE avión) ⇒ sello bajo el título de cada avión y aviso de snapshot por
  avión; el sello global dice el estado de la flota (PARCIAL incluido).
- Balance por avión: `cierre.estado` = el del AVIÓN y `cierre.estado_flota`
  = el de la flota ⇒ «CIERRE DEL MES — XA-TST · Mes cerrado el …» + línea
  «Flota: …». Balance general y Libro Dinero: la flota puede estar PARCIAL.
- SIN las llaves nuevas (API ≤ 0.0.103) o con ellas en neutro: TODO byte a
  byte contra las huellas calculadas con el código del HEAD 29107e6 sobre
  los MISMOS payloads de `tests/_cierre_mes_payloads.py` (los que traen
  `cierre` de forma vieja, que `test_cierre_mes.py` no congela).
"""

from __future__ import annotations

import copy
from io import BytesIO

import pytest
from fastapi.testclient import TestClient
from openpyxl import load_workbook

from app.config import get_settings
from app.main import app
from app.schemas.reparto import RepartoCierre, RepartoCierreAvion, RepartoPdfRequest
from app.schemas.reportes import (
    BalanceAvionRequest,
    BalanceCierre,
    BalanceGeneralRequest,
    DineroXlsxRequest,
)
from app.services import cierre_mes
from app.services.balance_avion_xlsx import (
    libro_balance_avion,
    libro_balance_general,
    render_balance_avion_xlsx,
    render_balance_general_xlsx,
)
from app.services.dinero_xlsx import render_dinero_xlsx
from app.services.reparto_pdf import render_reparto_pdf
from app.services.reparto_xlsx import ENCABEZADOS_CIERRE, render_reparto_xlsx
from tests import _cierre_mes_payloads as pay
from tests._evaluador_formulas import verificar_libro
from tests.test_cierre_mes import (
    _HUELLA_REPARTO_PDF_HEAD,
    _HUELLA_REPARTO_XLSX_HEAD,
    _HUELLAS_HEAD,
    _cierre_351,
    _fila_de,
    _huella_json,
    _huella_pdf,
    _huella_xlsx,
    _reparto_351,
    _texto_pdf,
)

MAESTRA_TST = "reporte horas XA-TST"
LLAVES_NUEVAS_BALANCE = {"estado_flota", "aviones_cerrados", "aviones_activos"}

# ---------------------------------------------------------------------------
# Payloads: el reparto con cierre de `_cierre_mes_payloads` (#316 sale, #318
# viene) donde XB-TST está CERRADO y XB-DOS ABIERTO ⇒ la flota queda PARCIAL
# («1 de 2 aviones»). Los balances y el Dinero, con las llaves de la flota.
# ---------------------------------------------------------------------------

CERRADO_TST = {
    "estado": "CERRADO", "cerrado_at": "2026-10-16T19:02:11.123Z", "cerrado_por_nombre": "Ale",
    "reabierto_at": None, "version": 1, "coincide_con_snapshot": True,
}  # fmt: skip
ABIERTO = {
    "estado": "ABIERTO", "cerrado_at": None, "cerrado_por_nombre": None, "reabierto_at": None,
    "version": None, "coincide_con_snapshot": None,
}  # fmt: skip
SELLO_TST = "XB-TST · Mes cerrado el 16/10/2026 por Ale"
SELLO_DOS = "XB-DOS · Mes abierto"
PARCIAL_REPARTO = "Mes cerrado parcialmente: 1 de 2 aviones (XB-TST)"


def reparto_por_avion(*, coincide_tst: bool | None = True, **extra) -> dict:
    p = pay.reparto_con_cierre(coincide=coincide_tst)
    p["cierre"].update(estado="PARCIAL", aviones_cerrados=["XB-TST"])
    p["aviones"][0]["cierre_avion"] = {**CERRADO_TST, "coincide_con_snapshot": coincide_tst}
    p["aviones"][1]["cierre_avion"] = dict(ABIERTO)
    p.update(extra)
    return p


def reparto_parcial_sin_listas() -> dict:
    """La flota PARCIAL sin ningún vuelo diferido: solo sellos."""
    p = pay.reparto_base()
    p["cierre"] = {
        **pay.reparto_cierre_abierto_vacio(), "estado": "PARCIAL", "aviones_cerrados": ["XB-TST"],
    }  # fmt: skip
    p["aviones"][0]["cierre_avion"] = dict(CERRADO_TST)
    p["aviones"][1]["cierre_avion"] = dict(ABIERTO)
    return p


def individual_por_avion() -> dict:
    """XA-TST CERRADO (sus listas de siempre) con la flota PARCIAL (1 de 7)."""
    p = pay.individual_con_cierre()
    p["cierre"].update(estado_flota="PARCIAL", aviones_cerrados=["XA-TST"], aviones_activos=7)
    return p


def _xlsx_reparto(p: dict):
    return load_workbook(BytesIO(render_reparto_xlsx(RepartoPdfRequest.model_validate(p)))).active


def _col_a(ws) -> list[str]:
    return [c.value for c in ws["A"] if isinstance(c.value, str)]


def _dinero_vlos(p: dict):
    data = render_dinero_xlsx(DineroXlsxRequest.model_validate(p))
    return load_workbook(BytesIO(data))["dinero-vlos"]


# ---------------------------------------------------------------------------
# 1) Esquemas: aceptan las llaves y son liberales (jamás un 422).
# ---------------------------------------------------------------------------


def test_esquemas_aceptan_las_llaves_y_son_liberales() -> None:
    r = RepartoCierre.model_validate(
        {"estado": "PARCIAL", "aviones_cerrados": ["XB-TST", 7, None, " "], "aviones_activos": "3"}
    )
    assert (r.estado, r.aviones_cerrados, r.aviones_activos) == ("PARCIAL", ["XB-TST", "7"], None)
    assert r.tiene_algo()
    texto_en_vez_de_lista = RepartoCierre.model_validate({"aviones_cerrados": "x"})
    assert texto_en_vez_de_lista.aviones_cerrados is None
    assert RepartoCierre().aviones_cerrados is None and RepartoCierre().aviones_activos is None
    # Un avión: el bloque por avión, liberal como el global.
    a = RepartoCierreAvion.model_validate(
        {"estado": " cerrado ", "cerrado_at": "2026-10-16T19:02:11.123Z", "cerrado_por_nombre": 12,
         "version": True, "coincide_con_snapshot": "no"}
    )  # fmt: skip
    assert (a.estado, a.cerrado_por_nombre, a.version, a.coincide_con_snapshot) == (
        "cerrado", "12", None, None,
    )  # fmt: skip
    assert RepartoCierreAvion().estado is None
    req = RepartoPdfRequest.model_validate(reparto_por_avion())
    assert req.aviones[0].cierre_avion.estado == "CERRADO"
    assert req.aviones[1].cierre_avion.estado == "ABIERTO"
    basura = pay.reparto_base()
    basura["aviones"][0]["cierre_avion"] = "x"
    basura["aviones"][1]["cierre_avion"] = ["CERRADO"]
    assert all(a.cierre_avion is None for a in RepartoPdfRequest.model_validate(basura).aviones)
    # Balances y Dinero: la flota en el bloque `cierre`.
    b = BalanceCierre.model_validate(
        {"estado": "ABIERTO", "estado_flota": " parcial ", "aviones_cerrados": ["XB-DOS"],
         "aviones_activos": 7.0}
    )  # fmt: skip
    assert (b.estado_flota, b.aviones_cerrados, b.aviones_activos) == ("parcial", ["XB-DOS"], None)
    assert b.tiene_algo()  # la flota ya no está abierta: hay sello que pintar
    assert not BalanceCierre(estado="ABIERTO", estado_flota="ABIERTO").tiene_algo()
    assert BalanceCierre(estado="PARCIAL").tiene_algo()
    dinero = DineroXlsxRequest.model_validate({"cierre": {"estado": "PARCIAL"}})
    assert dinero.cierre.estado == "PARCIAL"


# ---------------------------------------------------------------------------
# 2) Textos (fuente única `cierre_mes.py`).
# ---------------------------------------------------------------------------


def test_texto_global_parcial() -> None:
    texto = cierre_mes.texto_estado_cierre
    flota = RepartoCierre(estado="PARCIAL", aviones_cerrados=["N621TX", "N58BT"], aviones_activos=7)
    assert texto(flota) == "Mes cerrado parcialmente: 2 de 7 aviones (N621TX, N58BT)"
    # Sin `aviones_activos`: el total del documento (solo si son MÁS que los
    # cerrados: PARCIAL quiere decir que no están todos).
    sin_total = RepartoCierre(estado="PARCIAL", aviones_cerrados=["N621TX", "N58BT"])
    assert texto(sin_total, 7) == "Mes cerrado parcialmente: 2 de 7 aviones (N621TX, N58BT)"
    assert texto(sin_total, 2) == "Mes cerrado parcialmente: 2 aviones (N621TX, N58BT)"
    assert texto(sin_total) == "Mes cerrado parcialmente: 2 aviones (N621TX, N58BT)"
    assert texto(RepartoCierre(estado="PARCIAL", aviones_cerrados=["N621TX"])) == (
        "Mes cerrado parcialmente: 1 avión (N621TX)"
    )
    # `aviones_activos` que no alcanza para los cerrados ⇒ el respaldo.
    raro = RepartoCierre(estado="PARCIAL", aviones_cerrados=["A", "B"], aviones_activos=1)
    assert texto(raro, 7) == "Mes cerrado parcialmente: 2 de 7 aviones (A, B)"
    assert texto(RepartoCierre(estado="PARCIAL")) == "Mes cerrado parcialmente"
    assert texto(BalanceCierre(estado="parcial", aviones_cerrados=[])) == "Mes cerrado parcialmente"
    # Los de siempre no cambian (ni con las llaves nuevas al lado).
    cerrado = BalanceCierre.model_validate({**pay.cierre_cerrado(), "aviones_cerrados": ["X"]})
    assert texto(cerrado) == "Mes cerrado el 16/10/2026 por Ale"
    assert texto(BalanceCierre(estado="ABIERTO", aviones_cerrados=["X"])) is None


def test_texto_por_avion_y_flota() -> None:
    por_avion = cierre_mes.texto_estado_cierre_avion
    cerrado = RepartoCierreAvion.model_validate(CERRADO_TST)
    assert por_avion("N621TX", cerrado) == "N621TX · Mes cerrado el 16/10/2026 por Ale"
    reabierto = RepartoCierreAvion(
        estado="REABIERTO", cerrado_at="2026-10-16T19:00:00Z", cerrado_por_nombre="Ale",
        reabierto_at="2026-10-17T15:00:00Z", reabierto_por_nombre="Dani",
    )  # fmt: skip
    assert por_avion("N621TX", reabierto) == "N621TX · Mes reabierto el 17/10/2026 por Dani"
    assert por_avion("N621TX", RepartoCierreAvion(estado="ABIERTO")) == "N621TX · Mes abierto"
    assert por_avion(None, RepartoCierreAvion(estado="ABIERTO")) == "Mes abierto"
    assert por_avion("N621TX", None) is None
    assert por_avion("N621TX", RepartoCierreAvion(estado="RARO")) is None
    assert por_avion("N621TX", RepartoCierreAvion()) is None
    # Línea «Flota: …» del balance por avión: solo si difiere del avión.
    flota = cierre_mes.texto_estado_flota
    parcial = BalanceCierre(
        estado="CERRADO", estado_flota="PARCIAL", aviones_cerrados=["XA-TST"], aviones_activos=7
    )
    assert flota(parcial) == "Flota: cerrada parcialmente (1 de 7 aviones: XA-TST)"
    sin_total = BalanceCierre(estado="ABIERTO", estado_flota="PARCIAL", aviones_cerrados=["A", "B"])
    assert flota(sin_total, 7) == "Flota: cerrada parcialmente (2 de 7 aviones: A, B)"
    assert flota(sin_total) == "Flota: cerrada parcialmente (2 aviones: A, B)"
    sin_lista = BalanceCierre(estado="ABIERTO", estado_flota="PARCIAL")
    assert flota(sin_lista) == "Flota: cerrada parcialmente"
    assert flota(BalanceCierre(estado="REABIERTO", estado_flota="CERRADO")) == "Flota: cerrada"
    assert flota(BalanceCierre(estado="CERRADO", estado_flota="REABIERTO")) == "Flota: reabierta"
    assert flota(BalanceCierre(estado="CERRADO", estado_flota="ABIERTO")) == "Flota: abierta"
    assert flota(BalanceCierre(estado="CERRADO", estado_flota="cerrado")) is None
    assert flota(BalanceCierre(estado="CERRADO", estado_flota="RARO")) is None
    assert flota(BalanceCierre(estado="CERRADO")) is None and flota(None) is None


def test_aviso_de_snapshot_por_avion() -> None:
    uno = cierre_mes.aviso_snapshot(["XB-TST"])
    assert uno == (
        "Los números de hoy de XB-TST no coinciden con los de su cierre (cambió un porcentaje "
        "de socio o un tipo de cambio después de cerrar): reabre ese avión y ciérralo de nuevo "
        "antes de dispersar."
    )
    assert cierre_mes.aviso_snapshot(["XB-TST", "XB-DOS", "N1"]).startswith(
        "Los números de hoy de XB-TST, XB-DOS y N1 no coinciden"
    )
    assert cierre_mes.aviso_snapshot(["XB-TST", "XB-DOS"]).endswith(
        "reabre esos aviones y ciérralos de nuevo antes de dispersar."
    )
    assert cierre_mes.aviso_snapshot([]) == cierre_mes.AVISO_SNAPSHOT
    req = RepartoPdfRequest.model_validate(reparto_por_avion(coincide_tst=False))
    assert cierre_mes.aviones_sin_snapshot(req.aviones) == ["XB-TST"]
    assert cierre_mes.hay_cierre_por_avion(req.aviones)
    abierto = RepartoPdfRequest.model_validate(pay.reparto_base())
    assert cierre_mes.aviones_sin_snapshot(abierto.aviones) == []
    assert not cierre_mes.hay_cierre_por_avion(abierto.aviones)


# ---------------------------------------------------------------------------
# 3) Reparto PDF: sello por avión, aviso por avión, flota PARCIAL.
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("pesos", [False, True])
def test_reparto_pdf_sellos_por_avion_y_flota_parcial(pesos: bool) -> None:
    p = reparto_por_avion(reparto_en_pesos=pesos)
    texto = _texto_pdf(render_reparto_pdf(RepartoPdfRequest.model_validate(p)))
    assert PARCIAL_REPARTO in texto
    assert SELLO_TST in texto and SELLO_DOS in texto
    # El sello de siempre ya no va en la portada (la flota está PARCIAL): la
    # única vez que aparece es en el bloque de XB-TST.
    assert texto.count("Mes cerrado el 16/10/2026 por Ale") == 1
    assert "AVISO:" not in texto
    # Las listas por avión de siempre siguen ahí.
    assert "Pasan a octubre 2026 (pendientes de cobro, no se reparten): #316" in texto


@pytest.mark.parametrize("pesos", [False, True])
def test_reparto_pdf_aviso_de_snapshot_por_avion(pesos: bool) -> None:
    p = reparto_por_avion(coincide_tst=False, reparto_en_pesos=pesos)
    texto = _texto_pdf(render_reparto_pdf(RepartoPdfRequest.model_validate(p)))
    aviso = cierre_mes.aviso_snapshot(["XB-TST"])
    # Arriba (nombrando al avión) y en el bloque de XB-TST; nunca el global
    # de siempre («reabre el mes …»), que ya no dice qué reabrir.
    assert texto.count(f"AVISO: {aviso}") == 2
    assert cierre_mes.AVISO_SNAPSHOT not in texto
    assert SELLO_TST in texto and SELLO_DOS in texto


def test_reparto_pdf_reabierto_por_avion_y_sin_matricula_en_sin_snapshot() -> None:
    p = reparto_por_avion()
    p["aviones"][0]["cierre_avion"].update(
        estado="REABIERTO", reabierto_at="2026-10-17T15:00:00Z", reabierto_por_nombre="Dani"
    )
    p["cierre"].update(estado="REABIERTO", aviones_cerrados=[])
    texto = _texto_pdf(render_reparto_pdf(RepartoPdfRequest.model_validate(p)))
    assert "XB-TST · Mes reabierto el 17/10/2026 por Dani" in texto
    assert SELLO_DOS in texto
    # El global REABIERTO de siempre (es el `cierre` raíz) en la portada.
    assert "Mes reabierto" in texto


def test_reparto_mes_abierto_con_cierre_avion_abierto_es_byte_a_byte() -> None:
    """Lo que manda el API 0.0.104 en un mes abierto: `cierre_avion` ABIERTO
    en todos, `aviones_cerrados: []` y `aviones_activos`: el PDF y el Excel
    de siempre."""
    p = pay.reparto_base()
    p["cierre"] = {
        **pay.reparto_cierre_abierto_vacio(), "aviones_cerrados": [], "aviones_activos": 2,
    }  # fmt: skip
    for a in p["aviones"]:
        a["cierre_avion"] = dict(ABIERTO)
    req = RepartoPdfRequest.model_validate(p)
    assert _huella_pdf(render_reparto_pdf(req)) == _HUELLA_REPARTO_PDF_HEAD
    assert _huella_xlsx(render_reparto_xlsx(req)) == _HUELLA_REPARTO_XLSX_HEAD


# ---------------------------------------------------------------------------
# 4) Reparto XLSX: subtítulo, aviso y la sección «Cierre del mes» por avión.
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("pesos", [False, True])
def test_reparto_xlsx_seccion_cierre_por_avion(pesos: bool) -> None:
    ws = _xlsx_reparto(reparto_por_avion(coincide_tst=False, reparto_en_pesos=pesos))
    assert ws["A2"].value.endswith(f"  ·  {PARCIAL_REPARTO}")
    aviso = cierre_mes.aviso_snapshot(["XB-TST"])
    textos = _col_a(ws)
    assert f"AVISO: {aviso}" in textos
    assert f"AVISO: {cierre_mes.AVISO_SNAPSHOT}" not in textos
    tit = _fila_de(ws, "Cierre del mes")
    assert ws.cell(row=tit, column=1).value == f"Cierre del mes — {PARCIAL_REPARTO}"
    # Una línea por avión (matrícula en A, sello en B:H) y el aviso del
    # avión cuyos números ya no son los de su cierre; luego las listas.
    assert [ws.cell(row=tit + 1, column=c).value for c in (1, 2)] == [
        "XB-TST", "Mes cerrado el 16/10/2026 por Ale",
    ]  # fmt: skip
    assert ws.cell(row=tit + 2, column=1).value == f"AVISO XB-TST: {aviso}"
    assert ws.cell(row=tit + 2, column=1).font.color.rgb.endswith("B91C1C")
    assert [ws.cell(row=tit + 3, column=c).value for c in (1, 2)] == ["XB-DOS", "Mes abierto"]
    assert f"B{tit + 3}:H{tit + 3}" in {str(r) for r in ws.merged_cells.ranges}
    assert [ws.cell(row=tit + 4, column=c).value for c in range(1, 9)] == ENCABEZADOS_CIERRE
    salen = _fila_de(ws, "Pasan a octubre 2026")
    assert salen > tit + 4 and ws.cell(row=salen + 1, column=2).value == "#316"


def test_reparto_xlsx_parcial_sin_listas_solo_sellos() -> None:
    ws = _xlsx_reparto(reparto_parcial_sin_listas())
    assert ws["A2"].value.endswith(f"  ·  {PARCIAL_REPARTO}")
    tit = _fila_de(ws, "Cierre del mes")
    assert [ws.cell(row=tit + 1, column=c).value for c in (1, 2)] == [
        "XB-TST", "Mes cerrado el 16/10/2026 por Ale",
    ]  # fmt: skip
    assert [ws.cell(row=tit + 2, column=c).value for c in (1, 2)] == ["XB-DOS", "Mes abierto"]
    # Sin vuelos diferidos no hay encabezado de listas ni listas.
    assert ws.cell(row=tit + 3, column=1).value is None
    assert not any(t.startswith("Pasan a") for t in _col_a(ws))
    assert "AVISO" not in " ".join(_col_a(ws))


# ---------------------------------------------------------------------------
# 5) Balance por avión (estado del avión + flota) y general (flota PARCIAL).
# ---------------------------------------------------------------------------


def test_balance_avion_estado_del_avion_y_linea_de_flota() -> None:
    p = individual_por_avion()
    data = render_balance_avion_xlsx(BalanceAvionRequest.model_validate(p))
    assert verificar_libro(data)
    ws = load_workbook(BytesIO(data))[MAESTRA_TST]
    tit = _fila_de(ws, "CIERRE DEL MES")
    assert ws.cell(row=tit, column=1).value == (
        "CIERRE DEL MES — XA-TST · Mes cerrado el 16/10/2026 por Ale"
    )
    assert ws.cell(row=tit + 1, column=1).value == "Mes: septiembre 2026"
    assert ws.cell(row=tit + 2, column=1).value == (
        "Flota: cerrada parcialmente (1 de 7 aviones: XA-TST)"
    )
    assert ws.cell(row=tit + 3, column=1).value.startswith("Pasan a octubre 2026")
    vista = libro_balance_avion(BalanceAvionRequest.model_validate({**p, "vinculos_gasto": True}))
    assert (vista["cierre"]["estado"], vista["cierre"]["estado_flota"]) == ("CERRADO", "PARCIAL")
    assert vista["cierre"]["aviones_cerrados"] == ["XA-TST"]
    assert vista["cierre"]["aviones_activos"] == 7


def test_balance_avion_abierto_con_flota_parcial_pinta_el_bloque() -> None:
    p = pay.individual_base()
    p["cierre"] = {**pay.cierre_abierto_vacio(), "estado_flota": "PARCIAL",
                   "aviones_cerrados": ["XB-DOS"], "aviones_activos": 7}  # fmt: skip
    data = render_balance_avion_xlsx(BalanceAvionRequest.model_validate(p))
    assert verificar_libro(data)
    ws = load_workbook(BytesIO(data))[MAESTRA_TST]
    tit = _fila_de(ws, "CIERRE DEL MES")
    assert ws.cell(row=tit, column=1).value == "CIERRE DEL MES — XA-TST · Mes abierto"
    assert ws.cell(row=tit + 1, column=1).value == "Mes: septiembre 2026"
    assert ws.cell(row=tit + 2, column=1).value == (
        "Flota: cerrada parcialmente (1 de 7 aviones: XB-DOS)"
    )
    # Con la flota también ABIERTA: el libro de siempre (y la vista solo gana
    # las llaves que el API mandó).
    p["cierre"].update(estado_flota="ABIERTO", aviones_cerrados=[], aviones_activos=7)
    data = render_balance_avion_xlsx(BalanceAvionRequest.model_validate(p))
    assert _huella_xlsx(data) == _HUELLAS_HEAD["individual"][0]
    vista = libro_balance_avion(BalanceAvionRequest.model_validate({**p, "vinculos_gasto": True}))
    bloque = vista.pop("cierre")
    assert (bloque["estado_flota"], bloque["aviones_cerrados"], bloque["aviones_activos"]) == (
        "ABIERTO", [], 7,
    )  # fmt: skip
    assert _huella_json(vista) == _HUELLAS_HEAD["individual"][1]


def test_vista_sin_las_llaves_no_las_inventa() -> None:
    """Un API ≤ 0.0.103: el bloque `cierre` de la vista no gana llaves en null."""
    vista = libro_balance_avion(
        BalanceAvionRequest.model_validate({**pay.individual_con_cierre(), "vinculos_gasto": True})
    )
    assert not set(vista["cierre"]) & LLAVES_NUEVAS_BALANCE
    # Mandadas en null explícito, tampoco.
    p = pay.individual_con_cierre()
    p["cierre"].update(estado_flota=None, aviones_cerrados=None, aviones_activos=None)
    vista = libro_balance_avion(BalanceAvionRequest.model_validate({**p, "vinculos_gasto": True}))
    assert not set(vista["cierre"]) & LLAVES_NUEVAS_BALANCE


@pytest.mark.parametrize("variante", ["mensual", "general"])
def test_balance_general_flota_parcial(variante: str) -> None:
    p = pay.general_con_cierre(variante)
    p["cierre"].update(estado="PARCIAL", aviones_cerrados=["XA-TST"])
    assert len(p["aviones"]) == 2  # el «de 2» de respaldo: los aviones del libro
    data = render_balance_general_xlsx(BalanceGeneralRequest.model_validate(p))
    assert verificar_libro(data)
    rs = load_workbook(BytesIO(data))["RESUMEN flota"]
    tit = _fila_de(rs, "CIERRE DEL MES")
    assert rs.cell(row=tit, column=1).value == (
        "CIERRE DEL MES — Mes cerrado parcialmente: 1 de 2 aviones (XA-TST)"
    )
    assert rs.cell(row=_fila_de(rs, "Pasan a octubre 2026") + 2, column=4).value == "XA-TST"
    vista = libro_balance_general(
        BalanceGeneralRequest.model_validate({**p, "vinculos_gasto": True})
    )
    assert (vista["cierre"]["estado"], vista["cierre"]["aviones_cerrados"]) == (
        "PARCIAL", ["XA-TST"],
    )  # fmt: skip
    # Con `aviones_activos` del API manda sobre los aviones del libro.
    p["cierre"]["aviones_activos"] = 7
    data = render_balance_general_xlsx(BalanceGeneralRequest.model_validate(p))
    rs = load_workbook(BytesIO(data))["RESUMEN flota"]
    assert rs.cell(row=_fila_de(rs, "CIERRE DEL MES"), column=1).value.endswith(
        "1 de 7 aviones (XA-TST)"
    )


# ---------------------------------------------------------------------------
# 6) Libro Dinero: la flota PARCIAL con sus aviones cerrados.
# ---------------------------------------------------------------------------


def test_dinero_flota_parcial() -> None:
    p = pay.dinero_con_cierre()
    p["cierre"].update(estado="PARCIAL", aviones_cerrados=["N990GG"])
    ws = _dinero_vlos(p)
    tit = _fila_de(ws, "CIERRE DEL MES")
    # El «de 2»: los aviones de la leyenda de colores del libro.
    assert ws.cell(row=tit, column=1).value == (
        "CIERRE DEL MES — Mes cerrado parcialmente: 1 de 2 aviones (N990GG)"
    )
    assert ws.cell(row=tit + 1, column=1).value.startswith("Pasan a octubre 2026")
    # PARCIAL sin vuelos diferidos: solo el sello.
    p = pay.dinero_base()
    p["cierre"] = {**pay.cierre_abierto_vacio(), "estado": "PARCIAL",
                   "aviones_cerrados": ["N990GG", "N4142R"], "aviones_activos": 7}  # fmt: skip
    ws = _dinero_vlos(p)
    tit = _fila_de(ws, "CIERRE DEL MES")
    assert ws.cell(row=tit, column=1).value == (
        "CIERRE DEL MES — Mes cerrado parcialmente: 2 de 7 aviones (N990GG, N4142R)"
    )
    assert ws.cell(row=tit + 1, column=1).value is None


# ---------------------------------------------------------------------------
# 7) BYTE A BYTE: los payloads de `_cierre_mes_payloads` que traen `cierre`
# de forma vieja (sin las llaves nuevas) producen EXACTAMENTE los bytes del
# HEAD 29107e6 (antes del cierre por avión). Mismas huellas que
# `test_cierre_mes`: .xlsx sin docProps/core.xml, flujos de las páginas del
# PDF y JSON de la vista en pantalla.
# ---------------------------------------------------------------------------

# Huellas regeneradas el 23/10/2026 (API 0.0.105, apodos fuera de los documentos; pedido
# del cliente: «dejar solo las columnas de Num de cotización, nombre completo, fecha, ruta y
# consecutivo del vuelo»): entran CLIENTE y Nº VUELO CLIENTE entre CLAVE y FECHA en la hoja
# maestra, la cobranza y 'dinero-vlos', y cliente / ruta / nº vuelo cliente en 'otros
# movimientos' y 'Otros ingresos'; todas las demás columnas corren. Mismos números y las
# mismas fórmulas (citadas por llave, no por letra): solo cambia el layout.
# Huellas regeneradas el 26/10/2026 (notas de celda auditadas; pedido del cliente con la
# captura de la hoja «balance»: «esa nota que está en la celda B6 y B7 y todas están
# actualizadas?, porque unas cosas nos causan ruido»): cada nota de celda queda en una
# idea de 1–2 líneas, sin fechas ni versiones; los rótulos de la cascada de 'balance'
# pasan a «UTILIDAD ANTES DE COMISIONES» y «UTILIDAD DESPUÉS DE COMISIONES» (esta con su
# nota) y el pie de la reserva de overhaul ya no dice que el reparto la descuenta. Cambian
# solo textos (notas, VML y rótulos de la columna A): ningún número ni fórmula se movió.
# Huellas regeneradas el 05/11/2026 (vuelo CANCELADO en la hoja de vuelos: RUTA en rojo y
# TIEMPO / TACÓMETRO en rojo suave, pedido de la oficina; ver test_balance_cancelado_rojo):
# cambian solo la fuente y el relleno de la fila cancelada (y styles.xml renumera los
# estilos de todas las hojas); la vista JSON hereda el color. Ningún número se movió.
# Huellas regeneradas el 06/11/2026 («VENDIDOS CANT» en lugar de «SALIDAS CANT» en la hoja
# «inventario» del general, pedido del dueño; ver test_balance_inventario_xlsx): cambia solo
# ese texto de encabezado (y la hoja/valores que lo contienen). Ningún número se movió.
# Huellas regeneradas el 06/11/2026 (fechas DD/MM/AAAA en todos lados, pedido del dueño: el
# sello «Mes cerrado el 16/10/2026 por Ale» y los cortes «AL 30/09/2026» / «cobrado el
# 02/10/2026» del bloque CIERRE DEL MES cambian los bytes de los libros con cierre).
_HUELLAS_HEAD_29107E6 = {
    "individual": (
        "f2227d173ad589d58918f6d9c239b4cc2e959c009a99e5dab5eae018f0de6c2e",
        "bed974482dbf2d4a99b8bfbbe082863aa42f0e9133c016a885cdab6dd61c5346",
    ),
    "individual-351": (
        "ea1905c4f6cf006f233f6ca33fb700384fd329a6acb4705d3d48b69e4dba5ad1",
        "22bd3c2d954400246d867e067793608cdf6178cbd7c8a0ab4b273619e610f42a",
    ),
    "general-mensual": (
        "21ddf96ec28c941f18fbd113f30c6c6c79cd77de2750d651f447df76ab44662a",
        "d2dde3b282ae495a4884cdaa77cef5387364151044b93d0dc1d0392192a7d5cd",
    ),
    "general-general": (
        "f62e523a775b6f63d22ceb3cf5395eeb0ac5c1f788757b23df510cc13d989519",
        "466e7f1bffa0e1f1d4733c9f3f941e8fa0b65dd40b460d8d3f4b69026239803c",
    ),
    "general-consolidado": (
        "21ddf96ec28c941f18fbd113f30c6c6c79cd77de2750d651f447df76ab44662a",
        "0afcd37d2a705c35daf889cf4184a73d8cd84c65522d060762e5bdd61b46cba8",
    ),
    "reparto": (
        "d1f936c3122d9ecba89f14249e73c6c46f343f15dceb40a5c44e39ffba2d2aca",
        "d551c567b99d7fb916928645bbc79b8bc0728f9355ede230b72fc697809fe5a4",
    ),
    "reparto-sin-snapshot": (
        "23dfd1e26ec113d2720b54cc8e22c0f50eb082e091938ba568e80b83b31010ed",
        "a9cf556b810fff72625e13de7d2028c63df8ccd39d8aa4c3352677214f80d7f9",
    ),
    "reparto-reabierto": (
        "99cbaf0a274c26bf85e4ef469e8cecd8f9859cc0aab8eed3169375cb06841ac6",
        "b066c0786ec8ee606ed255ac8895f1d6683e2d2967d9cecc6cd5c67f90540eb7",
    ),
    "reparto-351": (
        "053ab72315de05d90417c649ceb6777e4f2ec093ffbfa8d1a722ed7c3056ab00",
        "d42260cabfd6a75a8faab2b7cd9c160f65f14f1519773ffd7aeb26de20daba6f",
    ),
    "reparto-pesos-sin-snapshot": (
        "2e86eebc95dc6314f2de5139ba0bcf3a52fe8eff8c8b65bac4aaaaaa1963a626",
        "3226f372370af0296984e1076011e1f75e463d7302f16b0412d7e5935ad31056",
    ),
    "reparto-balance-sin-snapshot": (
        "6138e38aee9ca6c358e4d64bbce9fbb7288ab2c27a36e7cc5df27f019a383a02",
        "f7afae803ba4944ef07e16f917a4d24d05cc4b86d4733dd60faf0a390169f25d",
    ),
    "dinero": ("e811aad47cd97014e4ba76b7576ff818be6e29c2f6d9b4a3a6f96f5c0617697a",),
    "dinero-351": ("f3d11251fdad20ae2c968e93f231aa9e23911eee0edf5a06e0410a1a284a2830",),
}


def _balance_individual(p: dict) -> tuple[str, str]:
    return (
        _huella_xlsx(render_balance_avion_xlsx(BalanceAvionRequest.model_validate(p))),
        _huella_json(
            libro_balance_avion(BalanceAvionRequest.model_validate({**p, "vinculos_gasto": True}))
        ),
    )


def _balance_general(p: dict, *, vinculos: bool = True) -> tuple[str, str]:
    vista = {**p, "vinculos_gasto": True} if vinculos else p
    return (
        _huella_xlsx(render_balance_general_xlsx(BalanceGeneralRequest.model_validate(p))),
        _huella_json(libro_balance_general(BalanceGeneralRequest.model_validate(vista))),
    )


def _reparto(p: dict) -> tuple[str, str]:
    req = RepartoPdfRequest.model_validate(p)
    return (_huella_pdf(render_reparto_pdf(req)), _huella_xlsx(render_reparto_xlsx(req)))


def _dinero(p: dict) -> tuple[str]:
    return (_huella_xlsx(render_dinero_xlsx(DineroXlsxRequest.model_validate(p))),)


def _general_consolidado() -> dict:
    p = pay.general_con_cierre("mensual")
    p["consolidado"]["cierre"] = p.pop("cierre")
    return p


_CASOS_HEAD = {
    "individual": lambda: _balance_individual(pay.individual_con_cierre()),
    "individual-351": lambda: _balance_individual(_cierre_351(pay.individual_con_cierre())),
    "general-mensual": lambda: _balance_general(pay.general_con_cierre("mensual")),
    "general-general": lambda: _balance_general(pay.general_con_cierre("general")),
    "general-consolidado": lambda: _balance_general(_general_consolidado(), vinculos=False),
    "reparto": lambda: _reparto(pay.reparto_con_cierre()),
    "reparto-sin-snapshot": lambda: _reparto(pay.reparto_con_cierre(coincide=False)),
    "reparto-reabierto": lambda: _reparto(pay.reparto_con_cierre(estado="REABIERTO")),
    "reparto-351": lambda: _reparto(_reparto_351()),
    "reparto-pesos-sin-snapshot": lambda: _reparto(
        {**pay.reparto_con_cierre(coincide=False), "reparto_en_pesos": True}
    ),
    "reparto-balance-sin-snapshot": lambda: _reparto(
        {**pay.reparto_con_cierre(coincide=False), "reparto_en_pesos": True,
         "reparto_como_balance": True}
    ),  # fmt: skip
    "dinero": lambda: _dinero(pay.dinero_con_cierre()),
    "dinero-351": lambda: _dinero(_cierre_351(pay.dinero_con_cierre())),
}


@pytest.mark.parametrize("caso", list(_HUELLAS_HEAD_29107E6))
def test_sin_las_llaves_nuevas_byte_a_byte(caso: str) -> None:
    assert _CASOS_HEAD[caso]() == _HUELLAS_HEAD_29107E6[caso]


def test_con_las_llaves_nuevas_en_neutro_byte_a_byte() -> None:
    """Las llaves presentes pero sin nada que decir: `cierre_avion` ABIERTO
    en todos los aviones (sobre el `cierre` CERRADO de siempre) y las de la
    flota en null ⇒ los mismos bytes."""
    p = pay.reparto_con_cierre(coincide=False)
    p["cierre"].update(aviones_cerrados=None, aviones_activos=None)
    for a in p["aviones"]:
        a["cierre_avion"] = dict(ABIERTO)
    assert _reparto(p) == _HUELLAS_HEAD_29107E6["reparto-sin-snapshot"]
    p = pay.individual_con_cierre()
    p["cierre"].update(estado_flota=None, aviones_cerrados=None, aviones_activos=None)
    assert _balance_individual(p) == _HUELLAS_HEAD_29107E6["individual"]
    p = pay.dinero_con_cierre()
    p["cierre"].update(estado_flota=None, aviones_cerrados=None, aviones_activos=None)
    assert _dinero(p) == _HUELLAS_HEAD_29107E6["dinero"]


# ---------------------------------------------------------------------------
# 8) Rutas: los payloads del API 0.0.104 entran (jamás un 422).
# ---------------------------------------------------------------------------

TOKEN = "tok-cierre-por-avion"
client = TestClient(app)


def test_rutas_aceptan_los_payloads_por_avion(monkeypatch, request) -> None:
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    request.addfinalizer(get_settings.cache_clear)
    h = {"X-Internal-Token": TOKEN}
    individual = individual_por_avion()
    general = pay.general_con_cierre("mensual")
    general["cierre"].update(estado="PARCIAL", aviones_cerrados=["XA-TST"], aviones_activos=7)
    reparto = reparto_por_avion(coincide_tst=False)
    dinero = pay.dinero_con_cierre()
    dinero["cierre"].update(estado="PARCIAL", aviones_cerrados=["N990GG"])
    for ruta, payload in (
        ("/pdf/balance-avion-xlsx", individual),
        ("/pdf/balance-general-xlsx", general),
        ("/pdf/balance-avion-libro", individual),
        ("/pdf/balance-general-libro", general),
        ("/pdf/reparto", reparto),
        ("/pdf/reparto-xlsx", reparto),
        ("/pdf/dinero-xlsx", dinero),
    ):
        res = client.post(ruta, json=copy.deepcopy(payload), headers=h)
        assert res.status_code == 200, (ruta, res.text[:300])
    res = client.post("/pdf/balance-avion-libro", json=individual, headers=h)
    assert (res.json()["cierre"]["estado"], res.json()["cierre"]["estado_flota"]) == (
        "CERRADO", "PARCIAL",
    )  # fmt: skip
