"""La RESERVA DE OVERHAUL RESTA EN LA CASCADA de la hoja 'balance' (9-nov-2026,
API 0.0.129, invariante 125 del API) y se DESCUENTA SOCIO POR SOCIO (API
0.0.130).

Decisión del cliente: «Necesito resolver ese tema del Overhaul, que quede
ANTES del total». Su Excel manual y el Diseño Funcional §4.8 tienen la reserva
como UNA RESTA MÁS de la cascada: INGRESOS − GASTOS DIRECTOS − GASTOS
INDIRECTOS − OTROS GASTOS − PERMISOS − RESERVA OVERHAUL − PENDIENTE DE COBRO =
SALDO DISPONIBLE. El API manda, junto con las llaves de la reserva del 0.0.98,
la llave booleana `reserva_overhaul_en_cascada: true` (el DISCRIMINADOR único)
y, con ella, `utilidad_despues_*` y `utilidad_cobrada_*` YA RESTAN la reserva.
Aquí: la fila «(−) RESERVA DE OVERHAUL» (valor del API) justo después de «(−)
PERMISOS» y antes de «(+) OTROS INGRESOS» / «UTILIDAD DESPUÉS DE GASTOS», cuya
fórmula la resta; todo lo de abajo corre una fila.

Y con el 0.0.129 desplegado el cliente siguió sin verlo: «aún lo veo como
antes, que solo es informativo; necesito que sea claro en cada total lo que se
les está descontando del total de la utilidad» (API 0.0.130): la tabla
«Reparto a socios» pasa a SOCIO · PORCENTAJE · UTILIDAD (antes de la reserva) ·
(−) RESERVA DE OVERHAUL · MONTO A RECIBIR [· USD informativo en pesos] [·
ENTREGAS], con fila TOTAL y nota al pie; `socios[].utilidad_*` = round2(% ×
(cobrada + reserva)), `reserva_*` = round2(% × reserva), `monto_*` = utilidad −
reserva (exacto), y el bloque ámbar del final dice que se RETIENE a los
socios. Sin la llave (API ≤ 0.0.127), con las llaves DEPRECADAS del 0.0.128
(`utilidad_a_repartir_*`) o con las de socio del 0.0.101 (`utilidad_*`,
`reserva_*`, `descuenta_reserva`, que viajan desde el 20-oct) pero SIN la
llave, el libro es BYTE A BYTE el de `test_balance_reserva_overhaul` (bloque
informativo incluido).
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
    BalanceAvionSocio,
    BalanceGeneralRequest,
)
from app.services import xlsx_formulas
from app.services.balance_avion_xlsx import (
    _ANCHO_COL_ENTREGAS,
    _ANCHOS_BALANCE_PESOS_EN_CASCADA,
    _ANCHOS_BALANCE_USD_EN_CASCADA,
    _COL_ENTREGADO,
    _COL_SOCIO_MONTO,
    _COL_SOCIO_MONTO_PESOS,
    _COL_SOCIO_RESERVA,
    _COL_SOCIO_RESERVA_PESOS,
    _COL_SOCIO_UTILIDAD,
    _COL_SOCIO_UTILIDAD_PESOS,
    _EMPRESA_PARTICIPACION,
    _FILA_ANTES,
    _FILA_ANTES_PESOS,
    _FILA_RESERVA_CASCADA,
    _FILA_RESERVA_CASCADA_PESOS,
    _FILA_RESERVA_EN_CASCADA,
    _FILA_RESERVA_EN_CASCADA_PESOS,
    _FILA_TOTAL_ENTREGAS,
    _FILA_TOTAL_SOCIOS,
    _NOTA_BALANCE_PESOS,
    _NOTA_BALANCE_PESOS_EN_CASCADA,
    _NOTA_BALANCE_PESOS_RESERVA,
    _NOTA_CASCADA_RESERVA,
    _NOTA_CASCADA_RESERVA_PESOS,
    _NOTA_ENTREGAS_SOCIOS,
    _NOTA_RESERVA_OVERHAUL_EN_CASCADA,
    _NOTA_SIN_PESOS_DERIVADA,
    _NOTA_SIN_PESOS_RESERVA,
    _NOTA_SIN_PESOS_SOCIO,
    _NOTA_SOCIOS_EN_CASCADA,
    _TITULO_RESERVA_OVERHAUL,
    _TITULO_RESERVA_OVERHAUL_EN_CASCADA,
    _TITULO_SOCIOS,
    _TITULO_SOCIOS_EN_CASCADA,
    _TITULO_SOCIOS_PESOS,
    GREEN,
    LIGHT,
    MONEY,
    RED,
    _formula_socio,
    _formulas_cascada,
    _reserva_en_cascada,
    libro_balance_avion,
    render_balance_avion_xlsx,
    render_balance_general_xlsx,
)
from tests import _pesos_payloads as pp
from tests import test_balance_empresa_vuelatour as t_empresa
from tests import test_balance_en_pesos as tp
from tests import test_balance_entregas_socios as t_ent
from tests import test_balance_reserva_overhaul as t_res
from tests import test_libro_json as tlj
from tests import test_notas_celda as t_notas
from tests._evaluador_formulas import verificar_libro

r2 = pp.r2
RESERVA_USD, RESERVA_MXN = t_res.RESERVA_USD, t_res.RESERVA_MXN
LLAVE = "reserva_overhaul_en_cascada"
LLAVES_SOCIO_0101 = ("utilidad_usd", "utilidad_mxn", "reserva_usd", "reserva_mxn")
# Textos EXACTOS del contrato (los necesita el panel).
TITULO_SOCIOS = (
    "Reparto a socios: su % de la utilidad ANTES de la reserva − su parte de la reserva de "
    "overhaul = monto a recibir"
)
COLS_USD = [
    "SOCIO",
    "PORCENTAJE",
    "UTILIDAD USD (antes de la reserva)",
    "(−) RESERVA DE OVERHAUL USD",
    "MONTO A RECIBIR USD",
]
COLS_MXN = [
    "SOCIO",
    "PORCENTAJE",
    "UTILIDAD (MXN, antes de la reserva)",
    "(−) RESERVA DE OVERHAUL (MXN)",
    "MONTO A RECIBIR (MXN)",
    "USD (informativo)",
]
NOTA_SOCIOS = (
    "A cada socio se le descuenta su parte de la reserva de overhaul (en proporción a su %); "
    "VuelaTour la guarda para el overhaul de los motores."
)
TITULO_EN_CASCADA = (
    "RESERVA DE OVERHAUL RETENIDA A LOS SOCIOS · ya restada en la cascada (fila «(−) RESERVA "
    "DE OVERHAUL»)"
)
FILA_EN_CASCADA_USD = "Retenido a los socios en el periodo USD"
FILA_EN_CASCADA_MXN = "Retenido a los socios en el periodo (MXN)"
NOTA_BLOQUE = (
    "Cada socio aporta su parte (columna «(−) RESERVA DE OVERHAUL» de la tabla de arriba) y "
    "VuelaTour la guarda para el overhaul de los motores. No es dinero depositado en una "
    "cuenta aparte ni un acumulado."
)
# Lo que ya no puede decir ningún texto de la hoja con la llave (ni lo del
# 0.0.98 «informativo», ni lo del 0.0.128 «a repartir»).
_SUPERADOS = (
    "informativa (no se resta",
    "no se resta en este balance",
    "bloque informativo",
    "solo se descuenta a quien lo pidió",
    "UTILIDAD A REPARTIR",
    "A REPARTIR",
    "debe aportar",
    "Lo que aporta el avión",
    "utilidad COBRADA × % vigente",
    "utilidad COBRADA en pesos × % vigente",
)


# ---------------------------------------------------------------------------
# Payloads: los de la reserva del 0.0.98 (`test_balance_reserva_overhaul`) con
# la llave del 0.0.129, espejo de la aritmética del API: la reserva restada en
# UTILIDAD DESPUÉS DE GASTOS y UTILIDAD COBRADA, y los socios del 0.0.130.
# ---------------------------------------------------------------------------


def _socios_0130(b: dict) -> None:
    """`socios[]` como los arma el API 0.0.130 a partir del bloque: utilidad =
    su % de (cobrada + reserva), reserva = su % de la reserva, monto =
    utilidad − reserva (exacto por renglón), `descuenta_reserva` true; sin
    pesos, las tres en None."""
    reserva, reserva_mxn = b["reserva_overhaul_usd"], b["reserva_overhaul_mxn"]
    cob, cob_mxn = b["utilidad_cobrada_usd"], b["utilidad_cobrada_mxn"]
    for s in b["socios"]:
        pct = s["porcentaje"]
        s["utilidad_usd"] = r2(pct / 100 * (cob + reserva))
        s["reserva_usd"] = r2(pct / 100 * reserva)
        s["monto_usd"] = r2(s["utilidad_usd"] - s["reserva_usd"])
        s["descuenta_reserva"] = True
        if cob_mxn is None or reserva_mxn is None:
            s["utilidad_mxn"] = s["reserva_mxn"] = s["monto_mxn"] = None
        else:
            s["utilidad_mxn"] = r2(pct / 100 * (cob_mxn + reserva_mxn))
            s["reserva_mxn"] = r2(pct / 100 * reserva_mxn)
            s["monto_mxn"] = r2(s["utilidad_mxn"] - s["reserva_mxn"])


def con_cascada(libro: dict, *, sin_tc: bool = False) -> dict:
    """El libro de un avión con la reserva EN la cascada (0.0.129) y los
    socios del 0.0.130. `sin_tc`: sin T.C. oficial del cierre ⇒ nada en pesos
    (reserva, utilidad después, utilidad cobrada y socios en None), como lo
    manda el API."""
    libro = t_res.con_reserva(libro, sin_tc=sin_tc)
    b = libro["balance"]
    b[LLAVE] = True
    reserva, reserva_mxn = b["reserva_overhaul_usd"], b["reserva_overhaul_mxn"]
    b["utilidad_despues_usd"] = r2(b["utilidad_despues_usd"] - reserva)
    b["utilidad_cobrada_usd"] = r2(b["utilidad_cobrada_usd"] - reserva)
    if reserva_mxn is None or b.get("utilidad_despues_mxn") is None:
        b["utilidad_despues_mxn"] = None
        b["utilidad_cobrada_mxn"] = None
    else:
        b["utilidad_despues_mxn"] = r2(b["utilidad_despues_mxn"] - reserva_mxn)
        cobrada_mxn = b.get("utilidad_cobrada_mxn")
        b["utilidad_cobrada_mxn"] = None if cobrada_mxn is None else r2(cobrada_mxn - reserva_mxn)
    _socios_0130(b)
    return libro


def con_socios_0101(libro: dict) -> dict:
    """El libro del 0.0.98 (reserva informativa, SIN la llave) con las llaves
    de socio que el API manda desde el 20-oct: utilidad = monto (su % de la
    cobrada), reserva = su % de la reserva, nadie descuenta."""
    libro = copy.deepcopy(libro)
    b = libro["balance"]
    for s in b["socios"]:
        pct = s["porcentaje"]
        s["utilidad_usd"], s["utilidad_mxn"] = s["monto_usd"], s.get("monto_mxn")
        s["reserva_usd"] = r2(pct / 100 * b["reserva_overhaul_usd"])
        s["reserva_mxn"] = (
            None if b["reserva_overhaul_mxn"] is None else r2(pct / 100 * b["reserva_overhaul_mxn"])
        )
        s["descuenta_reserva"] = False
    return libro


def individual(*, bandera: bool | None = True, **kw) -> dict:
    return con_cascada(pp.individual(bandera=bandera), **kw)


def _empresa_desde_socios(p: dict) -> None:
    """El bloque VUELATOUR del general con la participación de la empresa
    (espejo del API: Σ de los `monto_*` de sus socios `es_empresa`) y el
    resultado que sale de ella."""
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
    p["aviones"] = [con_cascada(a, **kw) for a in p["aviones"]]
    _empresa_desde_socios(p)
    return p


def _miembros(data: bytes) -> dict[str, bytes]:
    z = zipfile.ZipFile(BytesIO(data))
    return {n: z.read(n) for n in sorted(z.namelist()) if n != "docProps/core.xml"}


def _textos_a(ws) -> list[str]:
    return [c.value for c in ws["A"] if isinstance(c.value, str)]


def _sin_llave(libro_0098: dict):
    """La hoja 'balance' del libro del 0.0.98 (reserva solo informativa)."""
    if "aviones" in libro_0098:
        data = render_balance_general_xlsx(BalanceGeneralRequest.model_validate(libro_0098))
    else:
        data = render_balance_avion_xlsx(BalanceAvionRequest.model_validate(libro_0098))
    return load_workbook(BytesIO(data))["balance"]


@pytest.fixture
def registros(monkeypatch) -> list[xlsx_formulas.RegistroFormulas]:
    """Registro de fórmulas de cada libro que se cierra (cuáles volvieron a
    valor): con la llave ninguna debe degradarse."""
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
    reserva es informativa, que no se resta o que hay una utilidad «a
    repartir» aparte."""
    textos = _textos_a(ws) + [
        c.comment.text for fila in ws.iter_rows() for c in fila if c.comment is not None
    ]
    for t in textos:
        for frase in _SUPERADOS:
            assert frase not in t, (frase, t)


_REF = re.compile(r"\$?([A-Z]{1,3})\$?(\d+)")


def _ninguna_formula_cita_texto(ws) -> None:
    """Toda referencia de una fórmula a ESTA hoja cae en una celda que no es
    texto (una «—» citada daría #¡VALOR! en Excel)."""
    for fila in ws.iter_rows():
        for c in fila:
            if not (isinstance(c.value, str) and c.value.startswith("=")):
                continue
            propia = re.sub(r"'[^']*'!\$?[A-Z]{1,3}\$?\d+", "", c.value)
            for col, row in _REF.findall(propia):
                citada = ws[f"{col}{row}"].value
                es_texto = isinstance(citada, str) and not citada.startswith("=")
                assert not es_texto, (c.coordinate, c.value, citada)


def _filas_cascada_usd(ws, b: dict) -> dict[str, int]:
    """Las filas de la cascada en dólares que importan aquí."""
    k = tp._fila_de(ws, _FILA_ANTES)
    perm = tp._fila_de(ws, "(−) PERMISOS USD")
    assert perm == k + 4
    return {
        "k": k,
        "reserva": perm + 1,
        "despues": tp._fila_de(ws, "UTILIDAD DESPUÉS DE GASTOS USD"),
        "cobrada": tp._fila_de(ws, "UTILIDAD COBRADA USD"),
    }


def _formula_despues(letra: str, k: int, *, con_oi: bool) -> str:
    restas = "".join(f"-{letra}{k + i}" for i in range(1, 6))  # 4 gastos + reserva
    suma = f"+{letra}{k + 6}" if con_oi else ""
    return f"=ROUND({letra}{k}{restas}{suma},2)"


def _formulas_socio(r: int, cob: int, res: int, *, usd_info: bool = False) -> dict[str, str]:
    """Las fórmulas del renglón de un socio de la tabla del 0.0.130 (C, D, E
    y, en la hoja en pesos, la F informativa en dólares)."""
    f = {
        "C": f"=ROUND(B{r}/100*($B${cob}+$B${res}),2)",
        "D": f"=ROUND(B{r}/100*$B${res},2)",
        "E": f"=ROUND(C{r}-D{r},2)",
    }
    if usd_info:
        f["F"] = f"=ROUND(ROUND(B{r}/100*($C${cob}+$C${res}),2)-ROUND(B{r}/100*$C${res},2),2)"
    return f


def _tabla(ws, wsv, cob: int, b: dict, *, pesos: bool, res: int) -> int:
    """La tabla «Reparto a socios» del 0.0.130 que empieza en `cob + 2`
    (título): encabezados, renglones con sus fórmulas y números (citando la
    fila `res` de la reserva de SU cascada), TOTAL, y devuelve la fila de la
    nota al pie de la tabla (sin entregas)."""
    socios = b["socios"]
    titulo = ws.cell(row=cob + 2, column=1)
    assert titulo.value == TITULO_SOCIOS and titulo.font.b
    n_base = 6 if pesos else 5
    cols = COLS_MXN if pesos else COLS_USD
    assert [ws.cell(row=cob + 3, column=c).value for c in range(1, n_base + 1)] == cols
    m = "mxn" if pesos else "usd"
    for i, s in enumerate(socios):
        r = cob + 4 + i
        esperadas = _formulas_socio(r, cob, res, usd_info=pesos)
        assert ws.cell(row=r, column=1).value == s["nombre"]
        for letra, col in (("C", 3), ("D", 4), ("E", 5)):
            assert ws.cell(row=r, column=col).value == esperadas[letra], (r, letra)
        assert wsv.cell(row=r, column=3).value == s[f"utilidad_{m}"]
        assert wsv.cell(row=r, column=4).value == s[f"reserva_{m}"]
        assert wsv.cell(row=r, column=5).value == s[f"monto_{m}"]
        assert ws.cell(row=r, column=5).font.b and not ws.cell(row=r, column=3).font.b
        assert s[f"monto_{m}"] == r2(s[f"utilidad_{m}"] - s[f"reserva_{m}"])
        if pesos:
            assert ws.cell(row=r, column=6).value == esperadas["F"]
            assert wsv.cell(row=r, column=6).value == s["monto_usd"]
            assert ws.cell(row=r, column=6).font.i and ws.cell(row=r, column=6).font.size == 9
    ini, fin = cob + 4, cob + 3 + len(socios)
    total = fin + 1
    assert ws.cell(row=total, column=1).value == _FILA_TOTAL_SOCIOS == "TOTAL"
    assert ws.cell(row=total, column=1).fill.fgColor.rgb.endswith(LIGHT)
    assert f"A{total}:B{total}" in {str(x) for x in ws.merged_cells.ranges}
    for col, llave in ((3, "utilidad"), (4, "reserva"), (5, "monto")):
        letra = "CDE"[col - 3]
        assert ws.cell(row=total, column=col).value == f"=ROUND(SUM({letra}{ini}:{letra}{fin}),2)"
        assert ws.cell(row=total, column=col).font.b
        assert wsv.cell(row=total, column=col).value == r2(sum(s[f"{llave}_{m}"] for s in socios))
    # Σ de la columna de la reserva = la reserva del avión (lo que se retiene),
    # salvo el centavo de redondeo por renglón.
    reserva_avion = RESERVA_MXN if pesos else RESERVA_USD
    assert abs(wsv.cell(row=total, column=4).value - reserva_avion) <= 0.01 * len(socios)
    if pesos:
        assert ws.cell(row=total, column=6).value == f"=ROUND(SUM(F{ini}:F{fin}),2)"
        assert wsv.cell(row=total, column=6).value == r2(sum(s["monto_usd"] for s in socios))
    return total + 1


def _bloque_retenida(ws, inicio: int, b: dict, *, pesos: bool, como: str) -> None:
    """El bloque ámbar «RETENIDA A LOS SOCIOS» que empieza en `inicio` (la
    fila en blanco): título, monto retenido, «Cómo se calcula» y nota."""
    assert ws.cell(row=inicio, column=1).value is None
    titulo = ws.cell(row=inicio + 1, column=1)
    assert titulo.value == TITULO_EN_CASCADA and titulo.font.b
    if not pesos:
        # 94 caracteres en 76 de ancho: envuelve (en pesos, 107, cabe).
        assert ws.row_dimensions[inicio + 1].height > 15
    assert ws.cell(row=inicio + 2, column=1).value == (
        FILA_EN_CASCADA_MXN if pesos else FILA_EN_CASCADA_USD
    )
    if pesos:
        assert ws.cell(row=inicio + 2, column=2).value == (
            "—" if b["reserva_overhaul_mxn"] is None else b["reserva_overhaul_mxn"]
        )
        assert ws.cell(row=inicio + 2, column=3).value == b["reserva_overhaul_usd"]
    else:
        assert ws.cell(row=inicio + 2, column=2).value == b["reserva_overhaul_usd"]
    assert ws.cell(row=inicio + 3, column=1).value == t_res.PREFIJO_COMO + como
    assert ws.cell(row=inicio + 4, column=1).value == NOTA_BLOQUE


# ---------------------------------------------------------------------------
# 1) Esquema, discriminador y fórmulas puras.
# ---------------------------------------------------------------------------


def test_esquema_aditivo_discriminador_y_validador_liberal() -> None:
    b = BalanceAvionBalanceBloque()
    assert b.reserva_overhaul_en_cascada is None
    assert not _reserva_en_cascada(b)
    # Solo un `true` JSON cuenta; y siempre junto con la reserva USD.
    for v in (None, False, "true", 1, 0, {}):
        bloque = BalanceAvionBalanceBloque.model_validate({LLAVE: v, "reserva_overhaul_usd": 1.0})
        assert bloque.reserva_overhaul_en_cascada is None, v
        assert not _reserva_en_cascada(bloque), v
    con = BalanceAvionBalanceBloque.model_validate({LLAVE: True, "reserva_overhaul_usd": 0.0})
    assert _reserva_en_cascada(con)
    assert not _reserva_en_cascada(BalanceAvionBalanceBloque(**{LLAVE: True}))
    # Las llaves DEPRECADAS del 0.0.128 entran sin 422 y no encienden nada.
    viejo = BalanceAvionBalanceBloque.model_validate(
        {
            "utilidad_a_repartir_usd": 12.0,
            "utilidad_a_repartir_mxn": 200.0,
            "reserva_overhaul_usd": 1.0,
        }
    )
    assert viejo.utilidad_a_repartir_usd == 12.0 and not _reserva_en_cascada(viejo)
    # Las llaves de socio del 0.0.130: aditivas, None por default y liberales.
    s = BalanceAvionSocio()
    assert all(getattr(s, k) is None for k in (*LLAVES_SOCIO_0101, "descuenta_reserva"))
    s = BalanceAvionSocio.model_validate(
        {
            "nombre": "X", "utilidad_usd": 12.5, "reserva_usd": "x", "utilidad_mxn": None,
            "reserva_mxn": float("nan"), "descuenta_reserva": "si",
        }
    )  # fmt: skip
    assert s.utilidad_usd == 12.5 and s.reserva_usd is None and s.reserva_mxn is None
    assert s.utilidad_mxn is None and s.descuenta_reserva is None
    assert BalanceAvionSocio.model_validate({"descuenta_reserva": True}).descuenta_reserva is True
    # El payload del API 0.0.130: la cascada ya neta y los socios
    # utilidad − reserva, exacto por renglón.
    b = BalanceAvionRequest.model_validate(individual()).balance
    assert _reserva_en_cascada(b)
    sin = BalanceAvionRequest.model_validate(t_res.individual()).balance
    assert b.utilidad_despues_usd == r2(sin.utilidad_despues_usd - RESERVA_USD)
    assert b.utilidad_cobrada_usd == r2(sin.utilidad_cobrada_usd - RESERVA_USD)
    assert b.utilidad_cobrada_mxn == r2(sin.utilidad_cobrada_mxn - RESERVA_MXN)
    for s in b.socios:
        pct = s.porcentaje / 100
        assert s.utilidad_usd == r2(pct * (b.utilidad_cobrada_usd + RESERVA_USD))
        assert s.reserva_usd == r2(pct * RESERVA_USD)
        assert s.monto_usd == r2(s.utilidad_usd - s.reserva_usd)
        assert s.monto_mxn == r2(s.utilidad_mxn - s.reserva_mxn)
        assert s.descuenta_reserva is True
    assert r2(sum(s.reserva_usd for s in b.socios)) == RESERVA_USD


@pytest.mark.parametrize(
    ("con_oi", "esperado"),
    [
        (False, ("ROUND(B4-B5-B6-B7-B8-B9,2)", "ROUND(B10-B11,2)", 10, 12)),
        (True, ("ROUND(B4-B5-B6-B7-B8-B9+B10,2)", "ROUND(B11-B12,2)", 11, 13)),
    ],
    ids=["reserva", "reserva+oi"],
)
def test_formulas_cascada_con_la_reserva(con_oi, esperado) -> None:
    assert _formulas_cascada("B", 4, con_oi=con_oi, con_reserva=True) == esperado
    assert _formulas_cascada("C", 4, con_oi=con_oi, con_reserva=True)[0] == (
        esperado[0].replace("B", "C")
    )
    # Sin la reserva, las de siempre (`test_balance_reserva_overhaul` las fija).
    assert _formulas_cascada("B", 4, con_oi=False) == _formulas_cascada(
        "B", 4, con_oi=False, con_reserva=False
    )


def test_formula_del_socio_de_siempre_es_solo_el_respaldo() -> None:
    assert _formula_socio("B", 18, 12) == "ROUND(B18/100*$B$12,2)"
    assert _formula_socio("C", 19, 13) == "ROUND(B19/100*$C$13,2)"


def test_textos_exactos_del_contrato() -> None:
    assert _FILA_RESERVA_CASCADA == "(−) RESERVA DE OVERHAUL USD"
    assert _FILA_RESERVA_CASCADA_PESOS == "(−) RESERVA DE OVERHAUL (MXN)"
    assert _TITULO_SOCIOS_EN_CASCADA == TITULO_SOCIOS
    assert COLS_USD[2:] == [_COL_SOCIO_UTILIDAD, _COL_SOCIO_RESERVA, _COL_SOCIO_MONTO]
    assert COLS_MXN[2:5] == [
        _COL_SOCIO_UTILIDAD_PESOS,
        _COL_SOCIO_RESERVA_PESOS,
        _COL_SOCIO_MONTO_PESOS,
    ]
    assert _NOTA_SOCIOS_EN_CASCADA == NOTA_SOCIOS
    assert _TITULO_RESERVA_OVERHAUL_EN_CASCADA == TITULO_EN_CASCADA
    assert _FILA_RESERVA_EN_CASCADA == FILA_EN_CASCADA_USD
    assert _FILA_RESERVA_EN_CASCADA_PESOS == FILA_EN_CASCADA_MXN
    assert _NOTA_RESERVA_OVERHAUL_EN_CASCADA == NOTA_BLOQUE
    # Los de siempre no cambian (sin la llave se siguen pintando).
    assert _TITULO_SOCIOS == "Reparto a socios (utilidad COBRADA × % vigente)"
    assert _TITULO_SOCIOS_PESOS == "Reparto a socios (utilidad COBRADA en pesos × % vigente)"
    nuevos = (TITULO_SOCIOS, NOTA_SOCIOS, TITULO_EN_CASCADA, NOTA_BLOQUE)
    for t in (*nuevos, _NOTA_BALANCE_PESOS_EN_CASCADA):
        for frase in _SUPERADOS:
            assert frase not in t, (frase, t)
    assert "antes de UTILIDAD DESPUÉS DE GASTOS" in _NOTA_BALANCE_PESOS_EN_CASCADA
    assert "menos su parte de la reserva" in _NOTA_BALANCE_PESOS_EN_CASCADA


# ---------------------------------------------------------------------------
# 2) Libro INDIVIDUAL en dólares.
# ---------------------------------------------------------------------------


def test_individual_usd_resta_la_reserva_antes_de_la_utilidad_despues(registros) -> None:
    p = individual(bandera=None)
    b = p["balance"]
    data = render_balance_avion_xlsx(BalanceAvionRequest.model_validate(p))
    assert tp._degradadas_balance(registros[-1]) == []
    verificar_libro(data)
    wb, wv = tp._libros(data)
    ws, wsv = wb["balance"], wv["balance"]
    sin = _sin_llave(t_res.individual(bandera=None))
    f = _filas_cascada_usd(ws, b)
    k, res = f["k"], f["reserva"]
    # Hasta «(−) PERMISOS», celda por celda lo de siempre.
    assert t_res._celdas_hasta(ws, res - 1) == t_res._celdas_hasta(sin, res - 1)
    # La fila nueva: la reserva como VALOR (sin fórmula), gris, con su nota.
    assert ws.cell(row=res, column=1).value == _FILA_RESERVA_CASCADA
    rc = ws.cell(row=res, column=2)
    assert rc.value == RESERVA_USD and rc.number_format == MONEY and not rc.font.b
    assert rc.comment.text == _NOTA_CASCADA_RESERVA
    # Después de ella, (+) OTROS INGRESOS si el libro trae la hoja y UTILIDAD
    # DESPUÉS DE GASTOS = ROUND(antes − 4 gastos − reserva [+ otros], 2).
    con_oi = "otros_ingresos_avion" in p
    assert f["despues"] == res + (2 if con_oi else 1)
    dc = ws.cell(row=f["despues"], column=2)
    assert dc.value == _formula_despues("B", k, con_oi=con_oi) and dc.font.b
    assert wsv.cell(row=f["despues"], column=2).value == b["utilidad_despues_usd"]
    assert ws.cell(row=f["cobrada"], column=2).value == (
        f"=ROUND(B{f['despues']}-B{f['despues'] + 1},2)"
    )
    cob = f["cobrada"]
    assert wsv.cell(row=cob, column=2).value == b["utilidad_cobrada_usd"]
    assert ws.cell(row=cob, column=2).font.color.rgb.endswith(GREEN)
    # La cascada de siempre corrida UNA fila: los mismos rótulos hasta
    # UTILIDAD COBRADA.
    viejos = [c.value for c in sin["A"] if c.value is not None]
    i_perm, i_cob = viejos.index("(−) PERMISOS USD"), viejos.index("UTILIDAD COBRADA USD")
    nuevos = [c.value for c in ws["A"] if c.value is not None]
    assert nuevos[: i_cob + 2] == [
        *viejos[: i_perm + 1], _FILA_RESERVA_CASCADA, *viejos[i_perm + 1 : i_cob + 1]
    ]  # fmt: skip
    # La tabla socio por socio (0.0.130), su nota y el bloque «retenida».
    fila_nota = _tabla(ws, wsv, cob, b, pesos=False, res=res)
    assert ws.cell(row=fila_nota, column=1).value == NOTA_SOCIOS
    _bloque_retenida(ws, fila_nota + 1, b, pesos=False, como=t_res.COMO_PESOS)
    assert ws.cell(row=fila_nota + 6, column=1).value is None  # ahí termina la hoja
    # Título largo de la tabla: envuelve en dos líneas (negritas, 114 de ancho).
    assert ws.row_dimensions[cob + 2].height > 16
    assert [ws.column_dimensions[c].width for c in "ABCDE"] == _ANCHOS_BALANCE_USD_EN_CASCADA
    assert _TITULO_SOCIOS not in _textos_a(ws)
    _sin_textos_superados(ws)
    _ninguna_formula_cita_texto(ws)


def test_individual_usd_reserva_mayor_que_la_utilidad_deja_cobrada_en_rojo(registros) -> None:
    """Una reserva mayor que la utilidad deja UTILIDAD DESPUÉS y COBRADA en
    negativo (COBRADA en rojo) y a los socios con MONTO negativo (en rojo por
    formato condicional), cuadrando: utilidad − reserva."""
    p = individual(bandera=None)
    b = p["balance"]
    extra = r2(b["utilidad_cobrada_usd"] + 100)  # la reserva se sube `extra`
    b["reserva_overhaul_usd"] = r2(RESERVA_USD + extra)
    b["utilidad_despues_usd"] = r2(b["utilidad_despues_usd"] - extra)
    b["utilidad_cobrada_usd"] = -100.0
    _socios_0130(b)
    data = render_balance_avion_xlsx(BalanceAvionRequest.model_validate(p))
    assert tp._degradadas_balance(registros[-1]) == []
    verificar_libro(data)
    wb, wv = tp._libros(data)
    ws = wb["balance"]
    cob = tp._fila_de(ws, "UTILIDAD COBRADA USD")
    assert ws.cell(row=cob, column=2).font.color.rgb.endswith(RED)
    assert wv["balance"].cell(row=cob, column=2).value == -100.0
    s0 = b["socios"][0]
    assert wv["balance"].cell(row=cob + 4, column=5).value == s0["monto_usd"] < 0
    assert wv["balance"].cell(row=cob + 4, column=3).value == s0["utilidad_usd"] > 0
    total = cob + 4 + len(b["socios"])
    suma_montos = r2(sum(s["monto_usd"] for s in b["socios"]))
    assert wv["balance"].cell(row=total, column=5).value == suma_montos < 0
    rangos = {str(r) for cf in ws.conditional_formatting for r in cf.sqref.ranges}
    assert f"E{cob + 4}" in rangos and f"E{total}" in rangos


def test_individual_usd_skew_api_0129_sin_llaves_de_socio() -> None:
    """Un API 0.0.129 (la llave del bloque, pero los socios con `monto_*` =
    % × cobrada y sin utilidad/reserva): la tabla nueva deja UTILIDAD y
    RESERVA vacías y el MONTO con la fórmula de siempre; nada revienta."""
    p = individual(bandera=None)
    b = p["balance"]
    for s in b["socios"]:
        for k in (*LLAVES_SOCIO_0101, "descuenta_reserva"):
            s.pop(k)
        s["monto_usd"] = r2(s["porcentaje"] / 100 * b["utilidad_cobrada_usd"])
    data = render_balance_avion_xlsx(BalanceAvionRequest.model_validate(p))
    verificar_libro(data)
    wb, wv = tp._libros(data)
    ws, wsv = wb["balance"], wv["balance"]
    cob = tp._fila_de(ws, "UTILIDAD COBRADA USD")
    assert [ws.cell(row=cob + 3, column=c).value for c in range(1, 6)] == COLS_USD
    r = cob + 4
    assert ws.cell(row=r, column=3).value is None and ws.cell(row=r, column=4).value is None
    assert ws.cell(row=r, column=5).value == f"=ROUND(B{r}/100*$B${cob},2)"
    assert wsv.cell(row=r, column=5).value == b["socios"][0]["monto_usd"]
    total = cob + 4 + len(b["socios"])
    assert ws.cell(row=total, column=3).value is None and ws.cell(row=total, column=4).value is None
    assert wsv.cell(row=total, column=5).value == r2(sum(s["monto_usd"] for s in b["socios"]))
    _ninguna_formula_cita_texto(ws)


def test_individual_usd_con_entregas_las_tres_columnas_despues_del_monto(registros) -> None:
    """Con las ENTREGAS del API (invariante 100) las tres columnas van
    DESPUÉS de MONTO A RECIBIR (F..H), su fila de totales sigue a la fila
    TOTAL y las dos notas al pie se apilan."""
    p = t_ent.con_entregas(individual(bandera=None))
    b = p["balance"]
    data = render_balance_avion_xlsx(BalanceAvionRequest.model_validate(p))
    assert tp._degradadas_balance(registros[-1]) == []
    verificar_libro(data)
    wb, wv = tp._libros(data)
    ws, wsv = wb["balance"], wv["balance"]
    cob = tp._fila_de(ws, "UTILIDAD COBRADA USD")
    assert [ws.cell(row=cob + 3, column=c).value for c in range(1, 9)] == [
        *COLS_USD, _COL_ENTREGADO, "POR ENTREGAR (MXN)", "ENTREGADO USD (informativo)"
    ]  # fmt: skip
    res = tp._fila_de(ws, _FILA_RESERVA_CASCADA)
    fila_total = _tabla(ws, wsv, cob, b, pesos=False, res=res) - 1
    for i, s in enumerate(b["socios"]):
        r = cob + 4 + i
        assert wsv.cell(row=r, column=6).value == s["entregado_mxn"]
        assert wsv.cell(row=r, column=7).value == s["por_entregar_mxn"]
    pie = fila_total + 1
    assert ws.cell(row=pie, column=1).value == _FILA_TOTAL_ENTREGAS
    assert f"A{pie}:E{pie}" in {str(x) for x in ws.merged_cells.ranges}
    assert wsv.cell(row=pie, column=6).value == b["entregado_socios_mxn"]
    assert ws.cell(row=pie + 1, column=1).value == _NOTA_ENTREGAS_SOCIOS
    assert ws.cell(row=pie + 2, column=1).value == NOTA_SOCIOS
    _bloque_retenida(ws, pie + 3, b, pesos=False, como=t_res.COMO_PESOS)
    assert [ws.column_dimensions[c].width for c in "ABCDEFGH"] == [
        *_ANCHOS_BALANCE_USD_EN_CASCADA, *([_ANCHO_COL_ENTREGAS] * 3)
    ]
    _ninguna_formula_cita_texto(ws)


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
    sin = _sin_llave(t_res.individual())
    k = tp._fila_de(ws, _FILA_ANTES_PESOS)
    res = tp._fila_de(ws, "(−) PERMISOS (MXN)") + 1
    assert t_res._celdas_hasta(ws, res - 1) == t_res._celdas_hasta(sin, res - 1)
    assert ws.cell(row=res, column=1).value == _FILA_RESERVA_CASCADA_PESOS
    rc = ws.cell(row=res, column=2)
    assert rc.value == RESERVA_MXN and rc.comment.text == _NOTA_CASCADA_RESERVA_PESOS
    usd = ws.cell(row=res, column=3)
    assert usd.value == RESERVA_USD and usd.font.i and usd.font.size == 9
    con_oi = "otros_ingresos_avion" in p
    desp = tp._fila_de(ws, "UTILIDAD DESPUÉS DE GASTOS (MXN)")
    assert desp == res + (2 if con_oi else 1)
    assert ws.cell(row=desp, column=2).value == _formula_despues("B", k, con_oi=con_oi)
    assert ws.cell(row=desp, column=3).value == _formula_despues("C", k, con_oi=con_oi)
    assert wsv.cell(row=desp, column=2).value == b["utilidad_despues_mxn"]
    assert wsv.cell(row=desp, column=3).value == b["utilidad_despues_usd"]
    cob = tp._fila_de(ws, "UTILIDAD COBRADA (MXN)")
    assert cob == desp + 2
    assert wsv.cell(row=cob, column=2).value == b["utilidad_cobrada_mxn"]
    assert ws.cell(row=cob, column=2).font.color.rgb.endswith(GREEN)
    fila_nota = _tabla(ws, wsv, cob, b, pesos=True, res=res)
    assert ws.cell(row=fila_nota, column=1).value == NOTA_SOCIOS
    _bloque_retenida(ws, fila_nota + 1, b, pesos=True, como=t_res.COMO_PESOS)
    textos = _textos_a(ws)
    assert textos[-1] == _NOTA_BALANCE_PESOS_EN_CASCADA
    assert _NOTA_BALANCE_PESOS_RESERVA not in textos and _NOTA_BALANCE_PESOS not in textos
    assert _TITULO_SOCIOS_PESOS not in textos
    assert [ws.column_dimensions[c].width for c in "ABCDEF"] == _ANCHOS_BALANCE_PESOS_EN_CASCADA
    _sin_textos_superados(ws)
    _ninguna_formula_cita_texto(ws)


def test_individual_pesos_sin_tc_raya_sin_formulas_sobre_texto() -> None:
    """Sin T.C. oficial del cierre no hay reserva en pesos: la reserva,
    UTILIDAD DESPUÉS, UTILIDAD COBRADA y, en la tabla, UTILIDAD / RESERVA /
    MONTO de cada socio y su TOTAL quedan «—» con su nota en pesos (sin
    fórmula: ninguna cita una celda de texto); la columna USD informativa
    sigue con sus fórmulas y números."""
    p = individual(sin_tc=True)
    b = p["balance"]
    assert b["utilidad_despues_mxn"] is None and b["utilidad_cobrada_mxn"] is None
    assert all(s["monto_mxn"] is None and s["utilidad_mxn"] is None for s in b["socios"])
    data = render_balance_avion_xlsx(BalanceAvionRequest.model_validate(p))
    verificar_libro(data)
    wb, wv = tp._libros(data)
    ws, wsv = wb["balance"], wv["balance"]
    k = tp._fila_de(ws, _FILA_ANTES_PESOS)
    res = tp._fila_de(ws, "(−) PERMISOS (MXN)") + 1
    assert ws.cell(row=res, column=1).value == _FILA_RESERVA_CASCADA_PESOS
    assert ws.cell(row=res, column=2).value == "—"
    assert ws.cell(row=res, column=2).comment.text == _NOTA_SIN_PESOS_RESERVA
    assert wsv.cell(row=res, column=3).value == RESERVA_USD
    con_oi = "otros_ingresos_avion" in p
    desp = tp._fila_de(ws, "UTILIDAD DESPUÉS DE GASTOS (MXN)")
    assert ws.cell(row=desp, column=2).value == "—"
    assert ws.cell(row=desp, column=2).comment.text == _NOTA_SIN_PESOS_DERIVADA
    assert ws.cell(row=desp, column=3).value == _formula_despues("C", k, con_oi=con_oi)
    assert wsv.cell(row=desp, column=3).value == b["utilidad_despues_usd"]
    cob = tp._fila_de(ws, "UTILIDAD COBRADA (MXN)")
    assert ws.cell(row=cob, column=2).value == "—"
    assert wsv.cell(row=cob, column=3).value == b["utilidad_cobrada_usd"]
    r = cob + 4
    notas = {3: _NOTA_SIN_PESOS_SOCIO, 4: _NOTA_SIN_PESOS_RESERVA, 5: _NOTA_SIN_PESOS_SOCIO}
    for col, nota in notas.items():
        assert ws.cell(row=r, column=col).value == "—"
        assert ws.cell(row=r, column=col).comment.text == nota
    esperadas = _formulas_socio(r, cob, res, usd_info=True)
    assert ws.cell(row=r, column=6).value == esperadas["F"]
    assert wsv.cell(row=r, column=6).value == b["socios"][0]["monto_usd"]
    total = cob + 4 + len(b["socios"])
    assert ws.cell(row=total, column=1).value == "TOTAL"
    for col in (3, 4, 5):
        assert ws.cell(row=total, column=col).value == "—"
    assert wsv.cell(row=total, column=6).value == r2(sum(s["monto_usd"] for s in b["socios"]))
    # La cascada en pesos de arriba no se entera; el bloque dice «—» en pesos.
    assert wsv.cell(row=k, column=2).value == b["utilidad_antes_mxn"]
    _bloque_retenida(ws, total + 2, b, pesos=True, como=t_res.COMO)
    _ninguna_formula_cita_texto(ws)


def test_individual_pesos_sin_tc_con_despues_en_numero_va_como_valor() -> None:
    """Si un API mandara UTILIDAD DESPUÉS (MXN) con número y la reserva en
    pesos en None (no debería), la celda sale como VALOR: jamás una fórmula
    que cite la «—»."""
    p = individual(sin_tc=True)
    b = p["balance"]
    sin = t_res.individual(sin_tc=True)["balance"]
    b["utilidad_despues_mxn"] = sin["utilidad_despues_mxn"]
    data = render_balance_avion_xlsx(BalanceAvionRequest.model_validate(p))
    ws = load_workbook(BytesIO(data))["balance"]
    desp = tp._fila_de(ws, "UTILIDAD DESPUÉS DE GASTOS (MXN)")
    assert ws.cell(row=desp, column=2).value == sin["utilidad_despues_mxn"]
    _ninguna_formula_cita_texto(ws)


def test_individual_pesos_con_entregas_por_entregar_cita_el_monto_nuevo(registros) -> None:
    """Con las ENTREGAS, en pesos: las tres columnas en G..I y POR ENTREGAR =
    ROUND(E − G, 2), con E el MONTO A RECIBIR (MXN) nuevo."""
    p = t_ent.con_entregas(individual())
    b = p["balance"]
    data = render_balance_avion_xlsx(BalanceAvionRequest.model_validate(p))
    assert tp._degradadas_balance(registros[-1]) == []
    verificar_libro(data)
    wb, wv = tp._libros(data)
    ws, wsv = wb["balance"], wv["balance"]
    cob = tp._fila_de(ws, "UTILIDAD COBRADA (MXN)")
    assert [ws.cell(row=cob + 3, column=c).value for c in range(1, 10)] == [
        *COLS_MXN, _COL_ENTREGADO, "POR ENTREGAR (MXN)", "ENTREGADO USD (informativo)"
    ]  # fmt: skip
    res = tp._fila_de(ws, _FILA_RESERVA_CASCADA_PESOS)
    fila_total = _tabla(ws, wsv, cob, b, pesos=True, res=res) - 1
    for i, s in enumerate(b["socios"]):
        r = cob + 4 + i
        assert ws.cell(row=r, column=8).value == f"=ROUND(E{r}-G{r},2)"
        assert wsv.cell(row=r, column=8).value == s["por_entregar_mxn"]
    pie = fila_total + 1
    assert ws.cell(row=pie, column=1).value == _FILA_TOTAL_ENTREGAS
    assert f"A{pie}:F{pie}" in {str(x) for x in ws.merged_cells.ranges}
    assert ws.cell(row=pie + 2, column=1).value == NOTA_SOCIOS
    _ninguna_formula_cita_texto(ws)


# ---------------------------------------------------------------------------
# 4) Balance GENERAL: cada bloque resta su reserva; el bloque VUELATOUR cita
#    los MONTOS (columna E; en pesos también la F) de los socios empresa.
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("variante", ["mensual", "general"])
@pytest.mark.parametrize("bandera", [None, True], ids=["usd", "pesos"])
def test_general_cada_bloque_resta_y_vuelatour_cita_los_montos(
    registros, variante, bandera
) -> None:
    p = general(variante, bandera=bandera)
    data = render_balance_general_xlsx(BalanceGeneralRequest.model_validate(p))
    assert tp._degradadas_balance(registros[-1]) == []
    verificar_libro(data)
    wb, wv = tp._libros(data)
    ws, wsv = wb["balance"], wv["balance"]
    pesos = bool(bandera)
    sufijo = " (MXN)" if pesos else " USD"
    permisos = tp._filas_de(ws, f"(−) PERMISOS{sufijo}")
    despues = tp._filas_de(ws, f"UTILIDAD DESPUÉS DE GASTOS{sufijo}")
    cobradas = tp._filas_de(ws, f"UTILIDAD COBRADA{sufijo}")
    assert len(permisos) == len(despues) == len(cobradas) == len(p["aviones"]) == 2
    titulos = tp._filas_de(ws, TITULO_EN_CASCADA)
    assert len(titulos) == 2 and not tp._filas_de(ws, _TITULO_RESERVA_OVERHAUL)
    assert len(tp._filas_de(ws, TITULO_SOCIOS)) == 2 and len(tp._filas_de(ws, "TOTAL")) == 2
    for perm, desp, cob, titulo, avion in zip(
        permisos, despues, cobradas, titulos, p["aviones"], strict=True
    ):
        b = avion["balance"]
        res = perm + 1
        assert ws.cell(row=res, column=1).value == (
            _FILA_RESERVA_CASCADA_PESOS if pesos else _FILA_RESERVA_CASCADA
        )
        assert wsv.cell(row=res, column=2).value == (RESERVA_MXN if pesos else RESERVA_USD)
        formula = ws.cell(row=desp, column=2).value
        assert formula.startswith("=ROUND(B") and f"-B{res}" in formula
        assert wsv.cell(row=desp, column=2).value == (
            b["utilidad_despues_mxn"] if pesos else b["utilidad_despues_usd"]
        )
        fila_nota = _tabla(ws, wsv, cob, b, pesos=pesos, res=res)
        assert ws.cell(row=fila_nota, column=1).value == NOTA_SOCIOS
        assert titulo == fila_nota + 2
    emp = p["empresa"]
    part = tp._fila_de(
        ws,
        "(+) PARTICIPACIÓN COMO SOCIO EN LOS AVIONES (MXN)" if pesos else _EMPRESA_PARTICIPACION,
    )
    citadas = [int(x) for x in re.findall(r"E(\d+)", ws.cell(row=part, column=2).value)]
    assert len(citadas) == 2 and "C" not in ws.cell(row=part, column=2).value.replace("ROUND", "")
    assert all(ws.cell(row=c, column=1).value == t_empresa._EMPRESA for c in citadas)
    assert all(ws.cell(row=c, column=5).value == f"=ROUND(C{c}-D{c},2)" for c in citadas)
    assert wsv.cell(row=part, column=2).value == (
        emp["participacion_mxn"] if pesos else emp["participacion_usd"]
    )
    if pesos:
        citadas_usd = [int(x) for x in re.findall(r"F(\d+)", ws.cell(row=part, column=3).value)]
        assert citadas_usd == citadas
        assert wsv.cell(row=part, column=3).value == emp["participacion_usd"]
    resultado = tp._fila_de(ws, "RESULTADO VUELATOUR (MXN)" if pesos else "RESULTADO VUELATOUR USD")
    assert wsv.cell(row=resultado, column=2).value == (
        emp["resultado_mxn"] if pesos else emp["resultado_usd"]
    )
    if pesos:
        assert _textos_a(ws)[-1] == _NOTA_BALANCE_PESOS_EN_CASCADA
    anchos = _ANCHOS_BALANCE_PESOS_EN_CASCADA if pesos else _ANCHOS_BALANCE_USD_EN_CASCADA
    assert [ws.column_dimensions[c].width for c in "ABCDEF"[: len(anchos)]] == anchos
    _sin_textos_superados(ws)
    _ninguna_formula_cita_texto(ws)


def test_general_mixto_un_avion_en_cascada_y_el_otro_solo_informa() -> None:
    """Un API que mandara la llave en un avión y no en otro: cada bloque va
    por la suya (la decisión es por bloque, no por libro)."""
    p = pp.general("mensual", bandera=None)
    p["aviones"][0] = con_cascada(p["aviones"][0])
    p["aviones"][1] = t_res.con_reserva(p["aviones"][1])
    _empresa_desde_socios(p)
    data = render_balance_general_xlsx(BalanceGeneralRequest.model_validate(p))
    verificar_libro(data)
    ws = load_workbook(BytesIO(data))["balance"]
    assert len(tp._filas_de(ws, _FILA_RESERVA_CASCADA)) == 1
    assert len(tp._filas_de(ws, TITULO_EN_CASCADA)) == 1
    assert len(tp._filas_de(ws, _TITULO_RESERVA_OVERHAUL)) == 1
    assert len(tp._filas_de(ws, TITULO_SOCIOS)) == 1 and len(tp._filas_de(ws, _TITULO_SOCIOS)) == 1
    perm1, perm2 = tp._filas_de(ws, "(−) PERMISOS USD")
    assert ws.cell(row=perm1 + 1, column=1).value == _FILA_RESERVA_CASCADA
    assert ws.cell(row=perm2 + 1, column=1).value != _FILA_RESERVA_CASCADA
    # El segundo bloque conserva el bloque informativo de siempre.
    assert tp._filas_de(ws, t_res.FILA_USD) and tp._filas_de(ws, FILA_EN_CASCADA_USD)


# ---------------------------------------------------------------------------
# 5) Sin la llave (null, false, texto, solo las deprecadas del 0.0.128, o las
#    de socio del 0.0.101): el libro del 0.0.98 BYTE A BYTE, bloque
#    informativo incluido (.xlsx miembro a miembro y vista en pantalla).
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("forma", ["null", "false", "texto", "deprecadas-0128", "socios-0101"])
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
    if forma == "socios-0101":
        if libro == "individual":
            p = con_socios_0101(p)
        else:
            p["aviones"] = [con_socios_0101(a) for a in p["aviones"]]
    for b in t_res._bloques(p):
        if forma == "deprecadas-0128":
            b["utilidad_a_repartir_usd"] = r2(b["utilidad_cobrada_usd"] - RESERVA_USD)
            b["utilidad_a_repartir_mxn"] = 123.45
        elif forma != "socios-0101":
            b[LLAVE] = {"null": None, "false": False, "texto": "true"}[forma]
    data_base, vista_base = tp._render(libro, base)
    data, vista = tp._render(libro, p)
    assert _miembros(data) == _miembros(data_base)
    assert vista == vista_base


# ---------------------------------------------------------------------------
# 6) Vista en pantalla (el MISMO libro) y notas de celda limpias.
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("bandera", [None, True], ids=["usd", "pesos"])
def test_libro_json_exporta_la_fila_la_tabla_y_las_formulas(bandera) -> None:
    req = BalanceAvionRequest.model_validate(individual(bandera=bandera))
    libro = libro_balance_avion(req)
    assert tlj._comparar_con_xlsx(libro, render_balance_avion_xlsx(req)) > 0
    hoja = next(h for h in libro["hojas"] if h["nombre"] == "balance")
    celdas = {(c["r"], c["c"]): c for c in hoja["celdas"]}
    fila_res = _FILA_RESERVA_CASCADA_PESOS if bandera else _FILA_RESERVA_CASCADA
    res = next(r for (r, c), x in celdas.items() if c == 1 and x["v"] == fila_res)
    reserva = celdas[(res, 2)]
    assert reserva.get("f") is None  # valor, no fórmula
    assert reserva["v"] == (RESERVA_MXN if bandera else RESERVA_USD)
    assert reserva["n"] == (_NOTA_CASCADA_RESERVA_PESOS if bandera else _NOTA_CASCADA_RESERVA)
    b = req.balance
    sufijo = " (MXN)" if bandera else " USD"

    def fila_de(rotulo: str) -> int:
        return next(r for (r, c), x in celdas.items() if c == 1 and x["v"] == rotulo)

    desp = fila_de(f"UTILIDAD DESPUÉS DE GASTOS{sufijo}")
    despues = celdas[(desp, 2)]
    assert f"-B{res}" in despues["f"]
    assert despues["v"] == (b.utilidad_despues_mxn if bandera else b.utilidad_despues_usd)
    cob = fila_de(f"UTILIDAD COBRADA{sufijo}")
    cols = COLS_MXN if bandera else COLS_USD
    assert [celdas[(cob + 3, c)]["v"] for c in range(1, len(cols) + 1)] == cols
    r = cob + 4
    esperadas = _formulas_socio(r, cob, res, usd_info=bool(bandera))
    assert celdas[(r, 3)]["f"] == esperadas["C"] and celdas[(r, 4)]["f"] == esperadas["D"]
    assert celdas[(r, 5)]["f"] == esperadas["E"]
    s0 = b.socios[0]
    assert celdas[(r, 5)]["v"] == (s0.monto_mxn if bandera else s0.monto_usd)
    assert celdas[(r, 4)]["v"] == (s0.reserva_mxn if bandera else s0.reserva_usd)
    if bandera:
        assert celdas[(r, 6)]["f"] == esperadas["F"] and celdas[(r, 6)]["v"] == s0.monto_usd
    total = cob + 4 + len(b.socios)
    assert celdas[(total, 1)]["v"] == "TOTAL"
    reserva_m = "reserva_mxn" if bandera else "reserva_usd"
    assert celdas[(total, 4)]["v"] == r2(sum(getattr(s, reserva_m) for s in b.socios))
    textos = [x["v"] for (_, c), x in celdas.items() if c == 1 and isinstance(x["v"], str)]
    assert TITULO_SOCIOS in textos and NOTA_SOCIOS in textos
    assert TITULO_EN_CASCADA in textos and _TITULO_RESERVA_OVERHAUL not in textos


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
# 7) Rutas: el JSON del API 0.0.130 entra sin 422.
# ---------------------------------------------------------------------------

TOKEN = "tok-reserva-en-cascada"
client = TestClient(app)


def test_rutas_aceptan_el_payload_0130(monkeypatch, request) -> None:
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    request.addfinalizer(get_settings.cache_clear)
    headers = {"X-Internal-Token": TOKEN}
    res = client.post("/pdf/balance-avion-xlsx", json=individual(), headers=headers)
    assert res.status_code == 200, res.text
    ws = load_workbook(BytesIO(res.content))["balance"]
    assert tp._filas_de(ws, _FILA_RESERVA_CASCADA_PESOS)
    assert tp._filas_de(ws, TITULO_EN_CASCADA) and tp._filas_de(ws, TITULO_SOCIOS)
    res = client.post(
        "/pdf/balance-general-xlsx", json=general("general", bandera=None), headers=headers
    )
    assert res.status_code == 200, res.text
    ws = load_workbook(BytesIO(res.content))["balance"]
    assert len(tp._filas_de(ws, _FILA_RESERVA_CASCADA)) == 2
    res = client.post("/pdf/balance-general-libro", json=general("mensual"), headers=headers)
    assert res.status_code == 200, res.text
    hoja = next(h for h in res.json()["hojas"] if h["nombre"] == "balance")
    # Solo la columna A: el encabezado «(−) RESERVA DE OVERHAUL (MXN)» de la
    # tabla de socios (columna D) repite el rótulo de la fila de la cascada.
    rotulos = [c.get("v") for c in hoja["celdas"] if c.get("c") == 1]
    assert rotulos.count(_FILA_RESERVA_CASCADA_PESOS) == 2
    assert rotulos.count("TOTAL") == 2
