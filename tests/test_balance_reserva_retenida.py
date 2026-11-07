"""La RESERVA DE OVERHAUL SE RETIENE en la hoja 'balance' (7-nov-2026, API
0.0.128).

Decisión del cliente: «aún no se está reteniendo de las utilidades para cada
socio en cada avión, aún dice informativa no se resta; debe verse bien y
retenerse bien en el balance por avión y general». El API manda, junto con
las llaves de la reserva del 0.0.98, `utilidad_a_repartir_usd` (el
DISCRIMINADOR: round2(utilidad cobrada − reserva); 0.0 también cuenta) y
`utilidad_a_repartir_mxn` (None si falta alguna de las dos en pesos), y
`socios[].monto_*` llega ya NETO (round2(% × cobrada) − round2(% × reserva)).
Aquí: dos filas más justo después de UTILIDAD COBRADA («(−) RESERVA DE
OVERHAUL» como valor y «UTILIDAD A REPARTIR» = ROUND(cobrada − reserva, 2)),
cada socio = ROUND(% × cobrada, 2) − ROUND(% × reserva, 2), el título del
reparto y los pies sin «informativa» ni «no se resta», y el bloque ámbar del
final diciendo que se retiene (su «Cómo se calcula» se conserva). Sin la
llave (API ≤ 0.0.127) el libro es BYTE A BYTE el de
`test_balance_reserva_overhaul` (bloque informativo incluido).
"""

from __future__ import annotations

import copy
import re
import zipfile
from io import BytesIO

import pytest
from fastapi.testclient import TestClient
from openpyxl import load_workbook

from app.config import get_settings
from app.main import app
from app.schemas.reportes import (
    BalanceAvionBalanceBloque,
    BalanceAvionRequest,
    BalanceGeneralRequest,
)
from app.services import xlsx_formulas
from app.services.balance_avion_xlsx import (
    _EMPRESA_PARTICIPACION,
    _FILA_A_REPARTIR,
    _FILA_A_REPARTIR_PESOS,
    _FILA_RESERVA_CASCADA,
    _FILA_RESERVA_CASCADA_PESOS,
    _NOTA_BALANCE_PESOS,
    _NOTA_BALANCE_PESOS_RESERVA,
    _NOTA_BALANCE_PESOS_RETENIDA,
    _NOTA_CASCADA_RESERVA,
    _NOTA_CASCADA_RESERVA_PESOS,
    _NOTA_RESERVA_OVERHAUL,
    _NOTA_RESERVA_OVERHAUL_RETENIDA,
    _NOTA_SIN_PESOS_DERIVADA,
    _NOTA_SIN_PESOS_RESERVA,
    _NOTA_SIN_PESOS_SOCIO_A_REPARTIR,
    _TITULO_RESERVA_OVERHAUL,
    _TITULO_RESERVA_OVERHAUL_RETENIDA,
    _TITULO_SOCIOS,
    _TITULO_SOCIOS_A_REPARTIR,
    _TITULO_SOCIOS_PESOS,
    _TITULO_SOCIOS_PESOS_A_REPARTIR,
    GREEN,
    MONEY,
    RED,
    _formula_socio,
    _reserva_se_retiene,
    libro_balance_avion,
    render_balance_avion_xlsx,
    render_balance_general_xlsx,
)
from tests import _pesos_payloads as pp
from tests import test_balance_empresa_vuelatour as t_empresa
from tests import test_balance_en_pesos as tp
from tests import test_balance_reserva_overhaul as t_res
from tests import test_libro_json as tlj
from tests import test_notas_celda as t_notas
from tests._evaluador_formulas import verificar_libro

r2 = pp.r2
RESERVA_USD, RESERVA_MXN = t_res.RESERVA_USD, t_res.RESERVA_MXN
# Lo que ya no puede decir ningún texto de la hoja con la llave.
_SUPERADOS = (
    "informativa (no se resta",
    "no se resta en este balance",
    "bloque informativo",
    "solo se descuenta a quien lo pidió",
    "utilidad COBRADA × % vigente",
    "utilidad COBRADA en pesos × % vigente",
)


# ---------------------------------------------------------------------------
# Payloads: los de la reserva del 0.0.98 (`test_balance_reserva_overhaul`) con
# las llaves del 0.0.128, espejo de la aritmética del API.
# ---------------------------------------------------------------------------


def con_retencion(libro: dict, *, sin_tc: bool = False) -> dict:
    """El libro de un avión con la utilidad A REPARTIR y los socios NETOS
    encima de las llaves de la reserva. `sin_tc`: sin T.C. oficial del
    cierre ⇒ nada en pesos (reserva, utilidad a repartir y socios en None)."""
    libro = t_res.con_reserva(libro, sin_tc=sin_tc)
    b = libro["balance"]
    cobrada, reserva = b["utilidad_cobrada_usd"], b["reserva_overhaul_usd"]
    cobrada_mxn, reserva_mxn = b.get("utilidad_cobrada_mxn"), b["reserva_overhaul_mxn"]
    sin_pesos = cobrada_mxn is None or reserva_mxn is None
    b["utilidad_a_repartir_usd"] = r2(cobrada - reserva)
    b["utilidad_a_repartir_mxn"] = None if sin_pesos else r2(cobrada_mxn - reserva_mxn)
    for s in b["socios"]:
        pct = s["porcentaje"]
        s["monto_usd"] = r2(r2(pct / 100 * cobrada) - r2(pct / 100 * reserva))
        s["monto_mxn"] = (
            None
            if sin_pesos or s.get("monto_mxn") is None
            else r2(r2(pct / 100 * cobrada_mxn) - r2(pct / 100 * reserva_mxn))
        )
    return libro


def individual(*, bandera: bool | None = True, **kw) -> dict:
    return con_retencion(pp.individual(bandera=bandera), **kw)


def _empresa_neta(p: dict) -> None:
    """El bloque VUELATOUR del general con la participación de la empresa ya
    NETA (espejo del API: Σ de los `monto_*` de sus socios `es_empresa`) y
    el resultado que sale de ella."""
    emp = p["empresa"]
    socios = {
        a["matricula"]: s
        for a in p["aviones"]
        for s in a["balance"]["socios"]
        if s.get("es_empresa") is True
    }
    for part in emp["participaciones"]:
        s = socios[part["matricula"]]
        part["monto_usd"], part["monto_mxn"] = s["monto_usd"], s.get("monto_mxn")
    emp["participacion_usd"] = r2(sum(x["monto_usd"] for x in emp["participaciones"]))
    partes = [x["monto_mxn"] for x in emp["participaciones"]]
    emp["participacion_mxn"] = None if any(x is None for x in partes) else r2(sum(partes))
    for m in ("usd", "mxn"):
        comp = (
            emp[f"participacion_{m}"],
            emp[f"ingresos_propios_{m}"],
            emp[f"pagos_vendedor_{m}"],
            emp[f"otros_gastos_empresa_{m}"],
        )
        emp[f"resultado_{m}"] = (
            None
            if any(x is None for x in comp)
            else r2(comp[0] + comp[1] - comp[2] - comp[3] + (emp.get(f"tienda_utilidad_{m}") or 0))
        )


def general(variante: str = "mensual", *, bandera: bool | None = True, **kw) -> dict:
    p = pp.general(variante, bandera=bandera)
    p["aviones"] = [con_retencion(a, **kw) for a in p["aviones"]]
    _empresa_neta(p)
    return p


def _miembros(data: bytes) -> dict[str, bytes]:
    z = zipfile.ZipFile(BytesIO(data))
    return {n: z.read(n) for n in sorted(z.namelist()) if n != "docProps/core.xml"}


def _textos_a(ws) -> list[str]:
    return [c.value for c in ws["A"] if isinstance(c.value, str)]


@pytest.fixture
def registros(monkeypatch) -> list[xlsx_formulas.RegistroFormulas]:
    """Registro de fórmulas de cada libro que se cierra (cuáles volvieron a
    valor): con las llaves nuevas ninguna debe degradarse."""
    capturados: list[xlsx_formulas.RegistroFormulas] = []
    original = xlsx_formulas.finalizar

    def espia(wb):
        cache = original(wb)
        capturados.append(xlsx_formulas.registro(wb))
        return cache

    monkeypatch.setattr(xlsx_formulas, "finalizar", espia)
    return capturados


def _sin_textos_superados(ws) -> None:
    """Ningún texto de la hoja (columna A ni nota de celda) dice que la
    reserva es informativa o que no se resta."""
    textos = _textos_a(ws) + [
        c.comment.text for fila in ws.iter_rows() for c in fila if c.comment is not None
    ]
    for t in textos:
        for frase in _SUPERADOS:
            assert frase not in t, (frase, t)


# ---------------------------------------------------------------------------
# 1) Esquema, discriminador y la fórmula del socio.
# ---------------------------------------------------------------------------


def test_esquema_aditivo_y_discriminador() -> None:
    b = BalanceAvionBalanceBloque()
    assert b.utilidad_a_repartir_usd is None and b.utilidad_a_repartir_mxn is None
    assert not _reserva_se_retiene(b)
    assert _reserva_se_retiene(BalanceAvionBalanceBloque(utilidad_a_repartir_usd=0.0))
    assert _reserva_se_retiene(BalanceAvionBalanceBloque(utilidad_a_repartir_usd=-12.5))
    # Solo el MXN no enciende nada: el discriminador es el USD.
    assert not _reserva_se_retiene(BalanceAvionBalanceBloque(utilidad_a_repartir_mxn=10.0))
    b = BalanceAvionRequest.model_validate(individual()).balance
    assert b.utilidad_a_repartir_usd == r2(b.utilidad_cobrada_usd - RESERVA_USD)
    assert b.utilidad_a_repartir_mxn == r2(b.utilidad_cobrada_mxn - RESERVA_MXN)
    for s in b.socios:
        pct = s.porcentaje / 100
        assert s.monto_usd == r2(r2(pct * b.utilidad_cobrada_usd) - r2(pct * RESERVA_USD))


def test_formula_del_socio_es_la_aritmetica_del_api() -> None:
    assert _formula_socio("B", 18, 12, None) == "ROUND(B18/100*$B$12,2)"
    assert _formula_socio("B", 18, 12, 13) == "ROUND(B18/100*$B$12,2)-ROUND(B18/100*$B$13,2)"
    assert _formula_socio("C", 19, 13, 14) == "ROUND(B19/100*$C$13,2)-ROUND(B19/100*$C$14,2)"


# ---------------------------------------------------------------------------
# 2) Libro INDIVIDUAL en dólares.
# ---------------------------------------------------------------------------


def test_individual_usd_resta_la_reserva_y_los_socios_citan_las_dos_filas(registros) -> None:
    p = individual(bandera=None)
    b = p["balance"]
    data = render_balance_avion_xlsx(BalanceAvionRequest.model_validate(p))
    assert tp._degradadas_balance(registros[-1]) == []
    verificar_libro(data)
    wb, wv = tp._libros(data)
    ws, wsv = wb["balance"], wv["balance"]
    # Hasta UTILIDAD COBRADA, celda por celda lo de siempre (el libro del
    # 0.0.98 con la reserva solo informativa).
    sin = load_workbook(
        BytesIO(
            render_balance_avion_xlsx(
                BalanceAvionRequest.model_validate(t_res.individual(bandera=None))
            )
        )
    )["balance"]
    cob = tp._fila_de(ws, "UTILIDAD COBRADA USD")
    assert cob == tp._fila_de(sin, "UTILIDAD COBRADA USD")
    assert t_res._celdas_hasta(ws, cob) == t_res._celdas_hasta(sin, cob)
    # Las dos filas nuevas: la reserva como valor (con nota) y la utilidad a
    # repartir como fórmula en negritas, verde.
    res, rep = cob + 1, cob + 2
    assert ws.cell(row=res, column=1).value == _FILA_RESERVA_CASCADA
    rc = ws.cell(row=res, column=2)
    assert rc.value == RESERVA_USD and rc.number_format == MONEY and not rc.font.b
    assert rc.comment.text == _NOTA_CASCADA_RESERVA
    assert ws.cell(row=rep, column=1).value == _FILA_A_REPARTIR
    ac = ws.cell(row=rep, column=2)
    assert ac.value == f"=ROUND(B{cob}-B{res},2)" and ac.font.b
    assert ac.font.color.rgb.endswith(GREEN)
    assert wsv.cell(row=rep, column=2).value == b["utilidad_a_repartir_usd"]
    assert b["utilidad_a_repartir_usd"] == r2(b["utilidad_cobrada_usd"] - RESERVA_USD)
    # Título y socios dos renglones más abajo, con la resta como fórmula.
    assert ws.cell(row=rep + 2, column=1).value == _TITULO_SOCIOS_A_REPARTIR
    assert _TITULO_SOCIOS not in _textos_a(ws)
    for i, s in enumerate(b["socios"]):
        r = rep + 4 + i
        assert ws.cell(row=r, column=1).value == s["nombre"]
        assert ws.cell(row=r, column=3).value == (
            f"=ROUND(B{r}/100*$B${cob},2)-ROUND(B{r}/100*$B${res},2)"
        )
        assert wsv.cell(row=r, column=3).value == s["monto_usd"]
    # El bloque ámbar sigue tras los socios, ahora diciendo que se retiene y
    # con su «Cómo se calcula» intacto.
    fin = rep + 3 + len(b["socios"])
    assert ws.cell(row=fin + 2, column=1).value == _TITULO_RESERVA_OVERHAUL_RETENIDA
    assert ws.cell(row=fin + 2, column=1).font.b
    assert ws.cell(row=fin + 3, column=1).value == t_res.FILA_USD
    assert ws.cell(row=fin + 3, column=2).value == RESERVA_USD
    assert ws.cell(row=fin + 4, column=1).value == t_res.PREFIJO_COMO + t_res.COMO_PESOS
    assert ws.cell(row=fin + 5, column=1).value == _NOTA_RESERVA_OVERHAUL_RETENIDA
    textos = _textos_a(ws)
    assert _TITULO_RESERVA_OVERHAUL not in textos and _NOTA_RESERVA_OVERHAUL not in textos
    _sin_textos_superados(ws)


def test_individual_usd_utilidad_a_repartir_negativa_en_rojo(registros) -> None:
    """Una reserva mayor que la utilidad cobrada deja la utilidad a repartir
    en rojo (y los socios en negativo), con las fórmulas cuadrando."""
    p = individual(bandera=None)
    b = p["balance"]
    reserva = r2(b["utilidad_cobrada_usd"] + 100)
    b["reserva_overhaul_usd"] = reserva
    b["utilidad_a_repartir_usd"] = -100.0
    for s in b["socios"]:
        pct = s["porcentaje"]
        s["monto_usd"] = r2(r2(pct / 100 * b["utilidad_cobrada_usd"]) - r2(pct / 100 * reserva))
    data = render_balance_avion_xlsx(BalanceAvionRequest.model_validate(p))
    assert tp._degradadas_balance(registros[-1]) == []
    verificar_libro(data)
    wb, wv = tp._libros(data)
    ws = wb["balance"]
    rep = tp._fila_de(ws, _FILA_A_REPARTIR)
    assert ws.cell(row=rep, column=2).font.color.rgb.endswith(RED)
    assert wv["balance"].cell(row=rep, column=2).value == -100.0
    assert wv["balance"].cell(row=rep + 4, column=3).value == b["socios"][0]["monto_usd"] < 0


# ---------------------------------------------------------------------------
# 3) Libro INDIVIDUAL en pesos (bandera `balance_en_pesos`).
# ---------------------------------------------------------------------------


def test_individual_pesos_resta_la_reserva_en_las_dos_columnas(registros) -> None:
    p = individual()
    b = p["balance"]
    data = render_balance_avion_xlsx(BalanceAvionRequest.model_validate(p))
    assert tp._degradadas_balance(registros[-1]) == []
    verificar_libro(data)
    wb, wv = tp._libros(data)
    ws, wsv = wb["balance"], wv["balance"]
    sin = load_workbook(
        BytesIO(render_balance_avion_xlsx(BalanceAvionRequest.model_validate(t_res.individual())))
    )["balance"]
    cob = tp._fila_de(ws, "UTILIDAD COBRADA (MXN)")
    assert t_res._celdas_hasta(ws, cob) == t_res._celdas_hasta(sin, cob)
    res, rep = cob + 1, cob + 2
    assert ws.cell(row=res, column=1).value == _FILA_RESERVA_CASCADA_PESOS
    rc = ws.cell(row=res, column=2)
    assert rc.value == RESERVA_MXN and rc.comment.text == _NOTA_CASCADA_RESERVA_PESOS
    usd = ws.cell(row=res, column=3)
    assert usd.value == RESERVA_USD and usd.font.i and usd.font.size == 9
    assert ws.cell(row=rep, column=1).value == _FILA_A_REPARTIR_PESOS
    ac = ws.cell(row=rep, column=2)
    assert ac.value == f"=ROUND(B{cob}-B{res},2)" and ac.font.b
    assert ac.font.color.rgb.endswith(GREEN)
    assert ws.cell(row=rep, column=3).value == f"=ROUND(C{cob}-C{res},2)"
    assert wsv.cell(row=rep, column=2).value == b["utilidad_a_repartir_mxn"]
    assert b["utilidad_a_repartir_mxn"] == r2(b["utilidad_cobrada_mxn"] - RESERVA_MXN)
    assert wsv.cell(row=rep, column=3).value == b["utilidad_a_repartir_usd"]
    assert ws.cell(row=rep + 2, column=1).value == _TITULO_SOCIOS_PESOS_A_REPARTIR
    assert _TITULO_SOCIOS_PESOS not in _textos_a(ws)
    for i, s in enumerate(b["socios"]):
        r = rep + 4 + i
        assert ws.cell(row=r, column=3).value == (
            f"=ROUND(B{r}/100*$B${cob},2)-ROUND(B{r}/100*$B${res},2)"
        )
        assert ws.cell(row=r, column=4).value == (
            f"=ROUND(B{r}/100*$C${cob},2)-ROUND(B{r}/100*$C${res},2)"
        )
        assert wsv.cell(row=r, column=3).value == s["monto_mxn"]
        assert wsv.cell(row=r, column=4).value == s["monto_usd"]
    textos = _textos_a(ws)
    assert textos[-1] == _NOTA_BALANCE_PESOS_RETENIDA
    assert _NOTA_BALANCE_PESOS_RESERVA not in textos and _NOTA_BALANCE_PESOS not in textos
    fin = rep + 3 + len(b["socios"])
    assert ws.cell(row=fin + 2, column=1).value == _TITULO_RESERVA_OVERHAUL_RETENIDA
    assert ws.cell(row=fin + 3, column=1).value == t_res.FILA_MXN
    assert ws.cell(row=fin + 3, column=2).value == RESERVA_MXN
    assert ws.cell(row=fin + 4, column=1).value == t_res.PREFIJO_COMO + t_res.COMO_PESOS
    assert ws.cell(row=fin + 5, column=1).value == _NOTA_RESERVA_OVERHAUL_RETENIDA
    _sin_textos_superados(ws)


def test_individual_pesos_sin_tc_raya_con_nota_y_nunca_cero() -> None:
    """Sin T.C. oficial del cierre no hay reserva en pesos: la reserva, la
    utilidad a repartir y cada socio quedan «—» con su nota en pesos, y la
    columna USD sigue con sus fórmulas."""
    p = individual(sin_tc=True)
    b = p["balance"]
    assert b["utilidad_a_repartir_mxn"] is None
    assert all(s["monto_mxn"] is None for s in b["socios"])
    data = render_balance_avion_xlsx(BalanceAvionRequest.model_validate(p))
    verificar_libro(data)
    wb, wv = tp._libros(data)
    ws, wsv = wb["balance"], wv["balance"]
    cob = tp._fila_de(ws, "UTILIDAD COBRADA (MXN)")
    res, rep = cob + 1, cob + 2
    assert ws.cell(row=res, column=2).value == "—"
    assert ws.cell(row=res, column=2).comment.text == _NOTA_SIN_PESOS_RESERVA
    assert wsv.cell(row=res, column=3).value == RESERVA_USD
    assert ws.cell(row=rep, column=2).value == "—"
    assert ws.cell(row=rep, column=2).comment.text == _NOTA_SIN_PESOS_DERIVADA
    assert ws.cell(row=rep, column=3).value == f"=ROUND(C{cob}-C{res},2)"
    assert wsv.cell(row=rep, column=3).value == b["utilidad_a_repartir_usd"]
    r = rep + 4
    assert ws.cell(row=r, column=3).value == "—"
    assert ws.cell(row=r, column=3).comment.text == _NOTA_SIN_PESOS_SOCIO_A_REPARTIR
    assert ws.cell(row=r, column=4).value == (
        f"=ROUND(B{r}/100*$C${cob},2)-ROUND(B{r}/100*$C${res},2)"
    )
    assert wsv.cell(row=r, column=4).value == b["socios"][0]["monto_usd"]
    # La cascada en pesos de arriba no se entera.
    assert wsv.cell(row=cob, column=2).value == b["utilidad_cobrada_mxn"]


# ---------------------------------------------------------------------------
# 4) Balance GENERAL: cada bloque resta su reserva; el bloque VUELATOUR cita
#    los montos NETOS de los socios empresa (por coordenada, nunca por fila).
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("variante", ["mensual", "general"])
@pytest.mark.parametrize("bandera", [None, True], ids=["usd", "pesos"])
def test_general_cada_bloque_resta_y_vuelatour_cita_los_montos_netos(
    registros, variante, bandera
) -> None:
    p = general(variante, bandera=bandera)
    data = render_balance_general_xlsx(BalanceGeneralRequest.model_validate(p))
    assert tp._degradadas_balance(registros[-1]) == []
    verificar_libro(data)
    wb, wv = tp._libros(data)
    ws, wsv = wb["balance"], wv["balance"]
    pesos = bool(bandera)
    cobradas = tp._filas_de(ws, "UTILIDAD COBRADA (MXN)" if pesos else "UTILIDAD COBRADA USD")
    assert len(cobradas) == len(p["aviones"]) == 2
    titulos = tp._filas_de(ws, _TITULO_RESERVA_OVERHAUL_RETENIDA)
    assert len(titulos) == 2 and not tp._filas_de(ws, _TITULO_RESERVA_OVERHAUL)
    for cob, titulo, avion in zip(cobradas, titulos, p["aviones"], strict=True):
        b = avion["balance"]
        res, rep = cob + 1, cob + 2
        assert ws.cell(row=res, column=1).value == (
            _FILA_RESERVA_CASCADA_PESOS if pesos else _FILA_RESERVA_CASCADA
        )
        assert ws.cell(row=rep, column=1).value == (
            _FILA_A_REPARTIR_PESOS if pesos else _FILA_A_REPARTIR
        )
        assert ws.cell(row=rep, column=2).value == f"=ROUND(B{cob}-B{res},2)"
        assert wsv.cell(row=rep, column=2).value == (
            b["utilidad_a_repartir_mxn"] if pesos else b["utilidad_a_repartir_usd"]
        )
        fin = rep + 3 + len(b["socios"])
        assert titulo == fin + 2
        for i, s in enumerate(b["socios"]):
            r = rep + 4 + i
            assert ws.cell(row=r, column=1).value == s["nombre"]
            assert ws.cell(row=r, column=3).value == (
                f"=ROUND(B{r}/100*$B${cob},2)-ROUND(B{r}/100*$B${res},2)"
            )
            assert wsv.cell(row=r, column=3).value == (
                s["monto_mxn"] if pesos else s["monto_usd"]
            )
    emp = p["empresa"]
    part = tp._fila_de(
        ws,
        "(+) PARTICIPACIÓN COMO SOCIO EN LOS AVIONES (MXN)" if pesos else _EMPRESA_PARTICIPACION,
    )
    citadas = [int(x) for x in re.findall(r"C(\d+)", ws.cell(row=part, column=2).value)]
    assert len(citadas) == 2
    assert all(ws.cell(row=c, column=1).value == t_empresa._EMPRESA for c in citadas)
    assert all(ws.cell(row=c, column=3).value.count("ROUND(") == 2 for c in citadas)
    assert wsv.cell(row=part, column=2).value == (
        emp["participacion_mxn"] if pesos else emp["participacion_usd"]
    )
    resultado = tp._fila_de(ws, "RESULTADO VUELATOUR (MXN)" if pesos else "RESULTADO VUELATOUR USD")
    assert wsv.cell(row=resultado, column=2).value == (
        emp["resultado_mxn"] if pesos else emp["resultado_usd"]
    )
    if pesos:
        assert _textos_a(ws)[-1] == _NOTA_BALANCE_PESOS_RETENIDA
    _sin_textos_superados(ws)


def test_general_mixto_un_avion_retiene_y_el_otro_solo_informa() -> None:
    """Un API que mandara la llave en un avión y no en otro: cada bloque va
    por la suya (la decisión es por bloque, no por libro)."""
    p = pp.general("mensual", bandera=None)
    p["aviones"][0] = con_retencion(p["aviones"][0])
    p["aviones"][1] = t_res.con_reserva(p["aviones"][1])
    _empresa_neta(p)
    data = render_balance_general_xlsx(BalanceGeneralRequest.model_validate(p))
    verificar_libro(data)
    ws = load_workbook(BytesIO(data))["balance"]
    assert len(tp._filas_de(ws, _FILA_A_REPARTIR)) == 1
    assert len(tp._filas_de(ws, _TITULO_RESERVA_OVERHAUL_RETENIDA)) == 1
    assert len(tp._filas_de(ws, _TITULO_RESERVA_OVERHAUL)) == 1
    cob1, cob2 = tp._filas_de(ws, "UTILIDAD COBRADA USD")
    assert ws.cell(row=cob1 + 1, column=1).value == _FILA_RESERVA_CASCADA
    assert ws.cell(row=cob1 + 4, column=1).value == _TITULO_SOCIOS_A_REPARTIR
    assert ws.cell(row=cob2 + 1, column=1).value is None
    assert ws.cell(row=cob2 + 2, column=1).value == _TITULO_SOCIOS


# ---------------------------------------------------------------------------
# 5) Sin la llave (o solo el MXN): el libro del 0.0.98 BYTE A BYTE, bloque
#    informativo incluido (.xlsx miembro a miembro y vista en pantalla).
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("forma", ["null", "solo-mxn"])
@pytest.mark.parametrize(
    ("libro", "bandera"),
    [
        ("individual", None),
        ("individual", True),
        ("general-mensual", None),
        ("general-general", True),
    ],
)
def test_sin_la_llave_el_libro_del_0098_byte_a_byte(libro, bandera, forma) -> None:
    if libro == "individual":
        base = t_res.individual(bandera=bandera)
    else:
        base = t_res.general(libro.removeprefix("general-"), bandera=bandera)
    p = copy.deepcopy(base)
    for b in t_res._bloques(p):
        b["utilidad_a_repartir_usd"] = None
        b["utilidad_a_repartir_mxn"] = None if forma == "null" else 123.45
    data_base, vista_base = tp._render(libro, base)
    data, vista = tp._render(libro, p)
    assert _miembros(data) == _miembros(data_base)
    assert vista == vista_base


# ---------------------------------------------------------------------------
# 6) Vista en pantalla (el MISMO libro) y notas de celda limpias.
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("bandera", [None, True], ids=["usd", "pesos"])
def test_libro_json_exporta_las_filas_y_las_formulas(bandera) -> None:
    req = BalanceAvionRequest.model_validate(individual(bandera=bandera))
    libro = libro_balance_avion(req)
    assert tlj._comparar_con_xlsx(libro, render_balance_avion_xlsx(req)) > 0
    hoja = next(h for h in libro["hojas"] if h["nombre"] == "balance")
    celdas = {(c["r"], c["c"]): c for c in hoja["celdas"]}
    fila_rep = _FILA_A_REPARTIR_PESOS if bandera else _FILA_A_REPARTIR
    rep = next(r for (r, c), x in celdas.items() if c == 1 and x["v"] == fila_rep)
    cob, res = rep - 2, rep - 1
    assert celdas[(res, 1)]["v"] == (
        _FILA_RESERVA_CASCADA_PESOS if bandera else _FILA_RESERVA_CASCADA
    )
    reserva = celdas[(res, 2)]
    assert reserva.get("f") is None  # valor, no fórmula
    assert reserva["v"] == (RESERVA_MXN if bandera else RESERVA_USD)
    assert reserva["n"] == (_NOTA_CASCADA_RESERVA_PESOS if bandera else _NOTA_CASCADA_RESERVA)
    b = req.balance
    a_repartir = celdas[(rep, 2)]
    assert a_repartir["f"] == f"=ROUND(B{cob}-B{res},2)"
    assert a_repartir["v"] == (b.utilidad_a_repartir_mxn if bandera else b.utilidad_a_repartir_usd)
    socio = celdas[(rep + 4, 3)]
    assert socio["f"] == f"=ROUND(B{rep + 4}/100*$B${cob},2)-ROUND(B{rep + 4}/100*$B${res},2)"
    assert socio["v"] == (b.socios[0].monto_mxn if bandera else b.socios[0].monto_usd)
    if bandera:
        assert celdas[(rep, 3)]["f"] == f"=ROUND(C{cob}-C{res},2)"
    textos = [x["v"] for (_, c), x in celdas.items() if c == 1 and isinstance(x["v"], str)]
    assert _TITULO_RESERVA_OVERHAUL_RETENIDA in textos
    assert _TITULO_RESERVA_OVERHAUL not in textos


@pytest.mark.parametrize("bandera", [None, True], ids=["usd", "pesos"])
def test_notas_de_celda_limpias(bandera) -> None:
    """Las notas nuevas cumplen la regla de las notas de celda (una idea, sin
    fechas ni versiones, sin repetir el rótulo), como las de siempre."""
    data = render_balance_avion_xlsx(
        BalanceAvionRequest.model_validate(individual(bandera=bandera))
    )
    fallas: list[str] = []
    n = 0
    for hoja, celda, nota, rotulo, encabezado in t_notas._notas(data):
        n += 1
        fallas += [f"{hoja}!{celda}: {que}" for que in t_notas._desactualizado(nota)]
        fallas += [f"{hoja}!{celda}: {que}" for que in t_notas._ruido(nota, rotulo, encabezado)]
    assert n > 0 and not fallas, "\n".join(fallas)


# ---------------------------------------------------------------------------
# 7) Rutas: el JSON del API 0.0.128 entra sin 422.
# ---------------------------------------------------------------------------

TOKEN = "tok-reserva-retenida"
client = TestClient(app)


def test_rutas_aceptan_el_payload_0128(monkeypatch, request) -> None:
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    request.addfinalizer(get_settings.cache_clear)
    headers = {"X-Internal-Token": TOKEN}
    res = client.post("/pdf/balance-avion-xlsx", json=individual(), headers=headers)
    assert res.status_code == 200, res.text
    ws = load_workbook(BytesIO(res.content))["balance"]
    assert tp._filas_de(ws, _FILA_A_REPARTIR_PESOS)
    assert tp._filas_de(ws, _TITULO_RESERVA_OVERHAUL_RETENIDA)
    res = client.post(
        "/pdf/balance-general-xlsx", json=general("general", bandera=None), headers=headers
    )
    assert res.status_code == 200, res.text
    ws = load_workbook(BytesIO(res.content))["balance"]
    assert len(tp._filas_de(ws, _FILA_A_REPARTIR)) == 2
    res = client.post("/pdf/balance-general-libro", json=general("mensual"), headers=headers)
    assert res.status_code == 200, res.text
    hoja = next(h for h in res.json()["hojas"] if h["nombre"] == "balance")
    assert sum(1 for c in hoja["celdas"] if c.get("t") == _FILA_A_REPARTIR_PESOS) == 2
