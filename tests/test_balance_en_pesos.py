"""Balance por avión y balance general EN PESOS (18-oct-2026, API 0.0.93).

Pedido del cliente: «En los balances (balance general y por avión) debe venir
todo en pesos porque casi siempre se les transfiere a los socios en pesos
mexicanos, y que los dólares solo sean informativos pero no los que llamen
primero la atención.» El API manda los `*_mxn` ADITIVOS del bloque `balance`
(y `monto_mxn` de cada socio, y el bloque `empresa` en pesos) y la bandera
raíz `balance_en_pesos`; con ella la hoja 'balance' pinta la cascada y el
reparto en PESOS (lo principal) y los dólares en una columna INFORMATIVA
(gris, cursiva, letra 9). Aquí no se convierte nada: un `*_mxn` en None se
pinta «—» con una nota. Sin la bandera, el libro es BYTE A BYTE el del HEAD
3b893f1 (huellas `_HUELLAS_HEAD`, calculadas con el código de ese HEAD sobre
los MISMOS constructores de `tests/_pesos_payloads.py`).
"""

from __future__ import annotations

import copy
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
    BalanceAvionBalanceBloque,
    BalanceAvionRequest,
    BalanceAvionSocio,
    BalanceEmpresaBloque,
    BalanceEmpresaParticipacion,
    BalanceGeneralRequest,
)
from app.services import xlsx_formulas
from app.services.balance_avion_xlsx import (
    _ENCABEZADO_PESOS,
    _ENCABEZADO_USD_INFO,
    _FILA_ANTES_COMISIONES_PESOS,
    _FILA_ANTES_PESOS,
    _FILA_COBRADA_PESOS,
    _FILA_COMBUSTIBLE_PESOS,
    _FILA_COMISIONES_PESOS,
    _FILA_DESPUES_PESOS,
    _FILA_INDIRECTOS_PESOS,
    _FILA_OTROS_INGRESOS_PESOS,
    _FILA_PERMISOS_PESOS,
    _FILA_POR_COBRAR_PESOS,
    _FILA_REFACCIONES_PESOS,
    _NOTA_BALANCE_PESOS,
    _NOTA_SIN_PESOS,
    _NOTA_SIN_PESOS_DERIVADA,
    _NOTA_SIN_PESOS_SOCIO,
    _PREFIJO_NOTA_EMPRESA_USD,
    _TITULO_SOCIOS_PESOS,
    _USD_INFO_COLOR,
    libro_balance_avion,
    libro_balance_general,
    render_balance_avion_xlsx,
    render_balance_general_xlsx,
)
from tests import _pesos_payloads as pp
from tests._evaluador_formulas import verificar_libro

MAESTRA_TST = "reporte horas XA-TST"

# Huellas «sin la bandera»: calculadas con el código del HEAD 3b893f1 (antes
# del balance en pesos) corriendo `tests/_pesos_payloads.py` — .xlsx: todos
# los miembros salvo docProps/core.xml; vista en pantalla: el JSON con
# `vinculos_gasto: true`. Coinciden con las de `test_cierre_mes` porque las
# llaves `*_mxn` sin la bandera NO mueven nada. Si una se mueve, un API que no
# manda la bandera (≤ 0.0.92) ya no ve su libro de siempre.
# Huellas regeneradas el 23-oct-2026 (API 0.0.105, apodos fuera de los documentos; pedido
# del cliente: «dejar solo las columnas de Num de cotización, nombre completo, fecha, ruta y
# consecutivo del vuelo»): entran CLIENTE y Nº VUELO CLIENTE entre CLAVE y FECHA en la hoja
# maestra, la cobranza y 'dinero-vlos', y cliente / ruta / nº vuelo cliente en 'otros
# movimientos' y 'Otros ingresos'; todas las demás columnas corren. Mismos números y las
# mismas fórmulas (citadas por llave, no por letra): solo cambia el layout.
_HUELLAS_HEAD = {
    "individual": (
        "3a088800bfd81c515b73bc5b2b0cf3f8522fa7afebf5ed5013c5aa299d02714e",
        "11cef9feb0e6a43564704490ea3a552aad8e546f8d2a5956733c50bdb4036096",
    ),
    "general-mensual": (
        "22d118a1a3df1b26d2bbca1827bd396ef1474c2c84db3af549035f8b2a1945bb",
        "61247cf3aa3d6752721a24a5295a17bfbc72b2a2cd50deb716259942b669280a",
    ),
    "general-general": (
        "34afb591acef40817ae27b1e0ee2aa753fafa401674d55701b1da7c56df3a20e",
        "9817ca0aa214074160e32e5a1ed1f8e83732dbb70aa14ac9a985613fa4132e25",
    ),
}

# Rótulos de la cascada en pesos, en orden (libro individual con la hoja
# «Otros ingresos»).
_CASCADA_INDIVIDUAL = [
    _FILA_ANTES_PESOS,
    _FILA_COMBUSTIBLE_PESOS,
    _FILA_INDIRECTOS_PESOS,
    _FILA_REFACCIONES_PESOS,
    _FILA_PERMISOS_PESOS,
    _FILA_OTROS_INGRESOS_PESOS,
    _FILA_DESPUES_PESOS,
    _FILA_POR_COBRAR_PESOS,
    _FILA_COBRADA_PESOS,
]


def _huella_xlsx(data: bytes) -> str:
    z = zipfile.ZipFile(BytesIO(data))
    h = hashlib.sha256()
    for nombre in sorted(z.namelist()):
        if nombre != "docProps/core.xml":
            h.update(nombre.encode())
            h.update(z.read(nombre))
    return h.hexdigest()


def _huella_json(libro: dict) -> str:
    return hashlib.sha256(
        json.dumps(libro, sort_keys=True, ensure_ascii=False).encode()
    ).hexdigest()


def _payload(libro: str, bandera=True) -> dict:
    if libro == "individual":
        return pp.individual(bandera=bandera)
    return pp.general(libro.removeprefix("general-"), bandera=bandera)


def _render(libro: str, p: dict) -> tuple[bytes, dict]:
    if libro == "individual":
        return (
            render_balance_avion_xlsx(BalanceAvionRequest.model_validate(p)),
            libro_balance_avion(BalanceAvionRequest.model_validate({**p, "vinculos_gasto": True})),
        )
    return (
        render_balance_general_xlsx(BalanceGeneralRequest.model_validate(p)),
        libro_balance_general(BalanceGeneralRequest.model_validate({**p, "vinculos_gasto": True})),
    )


def _libros(data: bytes):
    return load_workbook(BytesIO(data)), load_workbook(BytesIO(data), data_only=True)


def _filas_de(ws, etiqueta: str) -> list[int]:
    return [c.row for c in ws["A"] if c.value == etiqueta]


def _fila_de(ws, etiqueta: str) -> int:
    filas = _filas_de(ws, etiqueta)
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


def _degradadas_balance(reg: xlsx_formulas.RegistroFormulas) -> list[tuple[str, str]]:
    return [d for d in reg.degradadas() if d[0] == "balance"]


# ---------------------------------------------------------------------------
# 1) Esquemas: aditivos, con default, bandera liberal.
# ---------------------------------------------------------------------------


def test_esquemas_aditivos_con_default() -> None:
    b = BalanceAvionBalanceBloque()
    for llave in (
        "utilidad_antes_mxn",
        "combustible_mxn",
        "gastos_indirectos_mxn",
        "refacciones_mxn",
        "otros_mxn",
        "permisos_mxn",
        "otros_ingresos_avion_mxn",
        "utilidad_despues_mxn",
        "por_cobrar_mxn",
        "utilidad_cobrada_mxn",
        "comisiones_mxn",
        "utilidad_antes_comisiones_mxn",
    ):
        assert getattr(b, llave) is None
    assert BalanceAvionSocio().monto_mxn is None
    assert BalanceEmpresaParticipacion().monto_mxn is None
    emp = BalanceEmpresaBloque()
    for llave in (
        "participacion_mxn",
        "ingresos_propios_mxn",
        "pagos_vendedor_mxn",
        "otros_gastos_empresa_mxn",
        "tienda_utilidad_mxn",
        "resultado_mxn",
        "nota_pesos",
    ):
        assert getattr(emp, llave) is None
    assert BalanceAvionRequest().balance_en_pesos is False
    assert BalanceGeneralRequest().balance_en_pesos is False


@pytest.mark.parametrize("valor", [None, False, "true", "si", 1, 0, {}])
def test_bandera_liberal_solo_true_la_enciende(valor) -> None:
    assert BalanceAvionRequest.model_validate({"balance_en_pesos": valor}).balance_en_pesos is False
    assert (
        BalanceGeneralRequest.model_validate({"balance_en_pesos": valor}).balance_en_pesos is False
    )
    assert BalanceAvionRequest.model_validate({"balance_en_pesos": True}).balance_en_pesos is True


def test_nota_pesos_liberal() -> None:
    assert BalanceEmpresaBloque.model_validate({"nota_pesos": "   "}).nota_pesos is None
    assert BalanceEmpresaBloque.model_validate({"nota_pesos": 12}).nota_pesos is None
    assert BalanceEmpresaBloque.model_validate({"nota_pesos": " Base "}).nota_pesos == "Base"


# ---------------------------------------------------------------------------
# 2) Sin la bandera: BYTE A BYTE (.xlsx y vista en pantalla).
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("libro", list(_HUELLAS_HEAD))
@pytest.mark.parametrize("bandera", [None, False, "si"])
def test_sin_bandera_byte_a_byte(libro: str, bandera) -> None:
    """Con las llaves `*_mxn` presentes y la bandera ausente, en false o con
    un valor que no es `true`, el libro es el del HEAD 3b893f1."""
    data, vista = _render(libro, _payload(libro, bandera))
    assert (_huella_xlsx(data), _huella_json(vista)) == _HUELLAS_HEAD[libro]


def test_bandera_del_avion_no_cambia_el_general() -> None:
    """En el general manda la bandera RAÍZ: un `aviones[]` con la bandera y la
    raíz sin ella ⇒ el libro de siempre."""
    p = pp.general("mensual", bandera=None)
    for a in p["aviones"]:
        a["balance_en_pesos"] = True
    data, vista = _render("general-mensual", p)
    assert (_huella_xlsx(data), _huella_json(vista)) == _HUELLAS_HEAD["general-mensual"]


# ---------------------------------------------------------------------------
# 3) Libro INDIVIDUAL en pesos.
# ---------------------------------------------------------------------------


def test_individual_cascada_en_pesos(registros) -> None:
    p = pp.individual()
    b = p["balance"]
    data = render_balance_avion_xlsx(BalanceAvionRequest.model_validate(p))
    assert _degradadas_balance(registros[-1]) == []
    verificar_libro(data)
    wb, wv = _libros(data)
    ws, wsv = wb["balance"], wv["balance"]
    assert "(en pesos MXN; los dólares solo informativos)" in ws["A2"].value
    # Encabezado de la cascada: pesos lo principal; la columna USD, gris.
    k = _fila_de(ws, _FILA_ANTES_PESOS)
    assert ws.cell(row=k - 1, column=1).value == "CONCEPTO"
    assert ws.cell(row=k - 1, column=2).value == _ENCABEZADO_PESOS
    info = ws.cell(row=k - 1, column=3)
    assert info.value == _ENCABEZADO_USD_INFO
    assert info.font.italic and info.font.size == 9
    # Rótulos EXACTOS, en orden y contiguos.
    assert [ws.cell(row=k + i, column=1).value for i in range(9)] == _CASCADA_INDIVIDUAL
    # Ningún rótulo de la cascada dice «USD».
    assert not any("USD" in (c.value or "") for c in ws["A"][k - 1 : k + 9])
    # Fórmulas: la misma cascada citando el TOTAL MXN de cada hoja.
    formulas = [ws.cell(row=k + i, column=2).value for i in range(9)]
    assert formulas[0] == f"='{MAESTRA_TST}'!$AF$9"  # GANANCIA MXN de TOTALES
    assert formulas[1:6] == [
        "='combustible'!$A$5",
        "='Gastos Indirectos'!$A$5",
        "='refacciones'!$A$5",
        "='permisos'!$A$5",
        "='Otros ingresos'!$A$5",
    ]
    assert formulas[6] == f"=ROUND(B{k}-B{k + 1}-B{k + 2}-B{k + 3}-B{k + 4}+B{k + 5},2)"
    assert formulas[7] == f"='{MAESTRA_TST}'!$AU$9"  # POR COBRAR MXN de TOTALES
    assert formulas[8] == f"=ROUND(B{k + 6}-B{k + 7},2)"
    # Valores = los `*_mxn` del API (la caché).
    esperados_mxn = [
        b["utilidad_antes_mxn"],
        b["combustible_mxn"],
        round(b["gastos_indirectos_mxn"] + b["otros_mxn"], 2),
        b["refacciones_mxn"],
        b["permisos_mxn"],
        b["otros_ingresos_avion_mxn"],
        b["utilidad_despues_mxn"],
        b["por_cobrar_mxn"],
        b["utilidad_cobrada_mxn"],
    ]
    assert [wsv.cell(row=k + i, column=2).value for i in range(9)] == esperados_mxn
    # USD informativo: los números de siempre, en gris, cursiva y letra 9.
    esperados_usd = [
        b["utilidad_antes_usd"],
        b["combustible_usd"],
        round(b["gastos_indirectos_usd"] + b["otros_usd"], 2),
        b["refacciones_usd"],
        b["permisos_usd"],
        b["otros_ingresos_avion_usd"],
        b["utilidad_despues_usd"],
        b["por_cobrar_usd"],
        b["utilidad_cobrada_usd"],
    ]
    assert [wsv.cell(row=k + i, column=3).value for i in range(9)] == esperados_usd
    for i in range(9):
        f = ws.cell(row=k + i, column=3).font
        assert f.italic and f.size == 9 and f.color.rgb.endswith(_USD_INFO_COLOR)
        assert ws.cell(row=k + i, column=2).font.size != 9
    assert ws.cell(row=k + 6, column=3).value == (
        f"=ROUND(C{k}-C{k + 1}-C{k + 2}-C{k + 3}-C{k + 4}+C{k + 5},2)"
    )
    # El pie lo explica.
    assert _NOTA_BALANCE_PESOS in [c.value for c in ws["A"]]


def test_individual_socios_en_pesos(registros) -> None:
    p = pp.individual()
    data = render_balance_avion_xlsx(BalanceAvionRequest.model_validate(p))
    assert _degradadas_balance(registros[-1]) == []
    wb, wv = _libros(data)
    ws, wsv = wb["balance"], wv["balance"]
    cobrada = _fila_de(ws, _FILA_COBRADA_PESOS)
    t = _fila_de(ws, _TITULO_SOCIOS_PESOS)
    assert [ws.cell(row=t + 1, column=c).value for c in range(1, 5)] == [
        "SOCIO",
        "PORCENTAJE",
        "MONTO MXN",
        _ENCABEZADO_USD_INFO,
    ]
    for i, s in enumerate(p["balance"]["socios"]):
        r = t + 2 + i
        assert ws.cell(row=r, column=1).value == s["nombre"]
        assert ws.cell(row=r, column=3).value == f"=ROUND(B{r}/100*$B${cobrada},2)"
        assert ws.cell(row=r, column=4).value == f"=ROUND(B{r}/100*$C${cobrada},2)"
        assert wsv.cell(row=r, column=3).value == s["monto_mxn"]
        assert wsv.cell(row=r, column=4).value == s["monto_usd"]
        assert ws.cell(row=r, column=3).font.bold
        assert ws.cell(row=r, column=4).font.italic


def test_individual_sin_pesos_pinta_raya_con_nota() -> None:
    """Un `*_mxn` en None (una hoja con una fila sin pesos): «—» con nota,
    jamás un 0; el null se propaga a DESPUÉS, COBRADA y los socios."""
    p = pp.individual()
    b = p["balance"]
    b["combustible_mxn"] = None
    b["utilidad_despues_mxn"] = None
    b["utilidad_cobrada_mxn"] = None
    for s in b["socios"]:
        s["monto_mxn"] = None
    data = render_balance_avion_xlsx(BalanceAvionRequest.model_validate(p))
    wb, wv = _libros(data)
    ws = wb["balance"]
    comb = _fila_de(ws, _FILA_COMBUSTIBLE_PESOS)
    despues = _fila_de(ws, _FILA_DESPUES_PESOS)
    cobrada = _fila_de(ws, _FILA_COBRADA_PESOS)
    assert ws.cell(row=comb, column=2).value == "—"
    assert ws.cell(row=comb, column=2).comment.text == _NOTA_SIN_PESOS
    for r in (despues, cobrada):
        assert ws.cell(row=r, column=2).value == "—"
        assert ws.cell(row=r, column=2).comment.text == _NOTA_SIN_PESOS_DERIVADA
    t = _fila_de(ws, _TITULO_SOCIOS_PESOS)
    for i in range(len(b["socios"])):
        celda = ws.cell(row=t + 2 + i, column=3)
        assert celda.value == "—"
        assert celda.comment.text == _NOTA_SIN_PESOS_SOCIO
        # El USD informativo sigue ahí.
        assert wv["balance"].cell(row=t + 2 + i, column=4).value == b["socios"][i]["monto_usd"]
    # Nada en la hoja dice 0 donde faltan pesos.
    assert wv["balance"].cell(row=comb, column=2).value == "—"


def test_individual_formula_que_no_cuadra_vuelve_valor(registros) -> None:
    """Si el TOTAL MXN de una hoja no reproduce el `*_mxn` del API, la celda
    se queda con el número del API (regla de siempre de `finalizar`)."""
    p = pp.individual()
    p["balance"]["combustible_mxn"] = round(p["balance"]["combustible_mxn"] + 10, 2)
    data = render_balance_avion_xlsx(BalanceAvionRequest.model_validate(p))
    comb = _fila_de(load_workbook(BytesIO(data))["balance"], _FILA_COMBUSTIBLE_PESOS)
    assert ("balance", f"B{comb}") in _degradadas_balance(registros[-1])
    wsv = load_workbook(BytesIO(data), data_only=True)["balance"]
    assert wsv.cell(row=comb, column=2).value == p["balance"]["combustible_mxn"]


def _contenido(data: bytes) -> dict[str, list]:
    """Valor, fórmula, formato de número y combinadas de cada hoja (lo que se
    ve; los índices de estilo del XML se corren al agregar estilos nuevos)."""
    wb = load_workbook(BytesIO(data))
    return {
        ws.title: [
            sorted(str(m) for m in ws.merged_cells.ranges),
            [
                (
                    c.coordinate,
                    c.value,
                    c.number_format,
                    c.font.b,
                    c.font.color and c.font.color.rgb,
                )
                for fila in ws.iter_rows()
                for c in fila
                if c.value is not None
            ],
        ]
        for ws in wb.worksheets
    }


def _hojas_distintas(con: bytes, sin: bytes) -> set[str]:
    a, b = _contenido(con), _contenido(sin)
    assert list(a) == list(b)
    return {h for h in a if a[h] != b[h]}


def test_individual_solo_cambia_la_hoja_balance() -> None:
    """Con la bandera, las demás pestañas son las de siempre."""
    con = render_balance_avion_xlsx(BalanceAvionRequest.model_validate(pp.individual()))
    sin = render_balance_avion_xlsx(BalanceAvionRequest.model_validate(pp.individual(bandera=None)))
    assert _hojas_distintas(con, sin) == {"balance"}


# ---------------------------------------------------------------------------
# 4) Balance GENERAL en pesos (bloques por avión + VUELATOUR).
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("variante", ["mensual", "general"])
def test_general_bloques_en_pesos(registros, variante: str) -> None:
    p = pp.general(variante)
    data = render_balance_general_xlsx(BalanceGeneralRequest.model_validate(p))
    assert _degradadas_balance(registros[-1]) == []
    verificar_libro(data)
    wb, wv = _libros(data)
    ws, wsv = wb["balance"], wv["balance"]
    assert "(en pesos MXN; los dólares solo informativos)" in ws["A2"].value
    # Un bloque por avión, con las dos previas de comisiones.
    previas = _filas_de(ws, _FILA_ANTES_COMISIONES_PESOS)
    assert len(previas) == len(p["aviones"])
    for fila, avion in zip(previas, p["aviones"], strict=True):
        b = avion["balance"]
        assert ws.cell(row=fila - 2, column=1).value == avion["matricula"]
        assert ws.cell(row=fila + 1, column=1).value == _FILA_COMISIONES_PESOS
        assert ws.cell(row=fila + 2, column=1).value == _FILA_ANTES_PESOS
        assert ws.cell(row=fila + 2, column=2).value == f"=ROUND(B{fila}-B{fila + 1},2)"
        assert wsv.cell(row=fila, column=2).value == b["utilidad_antes_comisiones_mxn"]
        assert wsv.cell(row=fila + 1, column=2).value == b["comisiones_mxn"]
        assert wsv.cell(row=fila + 2, column=2).value == b["utilidad_antes_mxn"]
        assert wsv.cell(row=fila, column=3).value == b["utilidad_antes_comisiones_usd"]
        cobrada = _filas_de(ws, _FILA_COBRADA_PESOS)[previas.index(fila)]
        assert wsv.cell(row=cobrada, column=2).value == b["utilidad_cobrada_mxn"]
        assert wsv.cell(row=cobrada, column=3).value == b["utilidad_cobrada_usd"]


def test_general_bloque_vuelatour_en_pesos(registros) -> None:
    p = pp.general("mensual")
    emp = p["empresa"]
    data = render_balance_general_xlsx(BalanceGeneralRequest.model_validate(p))
    assert _degradadas_balance(registros[-1]) == []
    wb, wv = _libros(data)
    ws, wsv = wb["balance"], wv["balance"]
    part = _fila_de(ws, "(+) PARTICIPACIÓN COMO SOCIO EN LOS AVIONES (MXN)")
    res = _fila_de(ws, "RESULTADO VUELATOUR (MXN)")
    otros = _fila_de(ws, "(−) OTROS GASTOS DE LA EMPRESA (MXN)")
    tienda = _fila_de(ws, "(+) UTILIDAD TIENDA (INVENTARIO) (MXN)")
    ing = next(c.row for c in ws["A"] if (c.value or "").startswith("(+) INGRESOS PROPIOS MXN ("))
    pag = next(c.row for c in ws["A"] if (c.value or "").startswith("(−) EGRESOS PROPIOS MXN ("))
    # PARTICIPACIÓN cita las celdas MONTO MXN de los socios empresa (col C);
    # la columna USD cita las de dólares (col D).
    formula = ws.cell(row=part, column=2).value
    assert formula.startswith("=ROUND(SUM(C") and ",C" in formula
    assert ws.cell(row=part, column=3).value.startswith("=ROUND(SUM(D")
    assert wsv.cell(row=part, column=2).value == emp["participacion_mxn"]
    assert wsv.cell(row=part, column=3).value == emp["participacion_usd"]
    assert ws.cell(row=otros, column=2).value == "='otros gastos'!$A$5"
    assert ws.cell(row=otros, column=3).value == "='otros gastos'!$C$5"
    assert wsv.cell(row=otros, column=2).value == emp["otros_gastos_empresa_mxn"]
    assert wsv.cell(row=ing, column=2).value == emp["ingresos_propios_mxn"]
    assert wsv.cell(row=pag, column=2).value == emp["pagos_vendedor_mxn"]
    assert wsv.cell(row=tienda, column=2).value == emp["tienda_utilidad_mxn"]
    assert ws.cell(row=res, column=2).value == (
        f"=ROUND(B{part}+B{ing}-B{pag}-B{otros}+B{tienda},2)"
    )
    assert wsv.cell(row=res, column=2).value == emp["resultado_mxn"]
    assert wsv.cell(row=res, column=3).value == emp["resultado_usd"]
    # Pie: la nota en pesos del API, la `nota` de siempre como USD
    # informativo y el pie general.
    textos = [c.value for c in ws["A"] if isinstance(c.value, str)]
    assert emp["nota_pesos"] in textos
    assert _PREFIJO_NOTA_EMPRESA_USD + emp["nota"] in textos
    assert textos[-1] == _NOTA_BALANCE_PESOS


def test_general_tienda_sin_inventario_va_vacia_no_raya() -> None:
    """Sin hoja 'inventario' el API manda la tienda en null y la cuenta como
    0: celda vacía (como en dólares), no «—», y el RESULTADO cuadra."""
    p = pp.general("mensual")
    p["inventario"] = None
    p = pp.con_empresa_mxn(p)
    assert p["empresa"]["tienda_utilidad_mxn"] is None
    data = render_balance_general_xlsx(BalanceGeneralRequest.model_validate(p))
    wb, wv = _libros(data)
    ws = wb["balance"]
    tienda = _fila_de(ws, "(+) UTILIDAD TIENDA (INVENTARIO) (MXN)")
    res = _fila_de(ws, "RESULTADO VUELATOUR (MXN)")
    assert ws.cell(row=tienda, column=2).value is None
    assert ws.cell(row=tienda, column=2).comment is None
    assert wv["balance"].cell(row=res, column=2).value == p["empresa"]["resultado_mxn"]
    assert ws.cell(row=res, column=2).value.startswith("=ROUND(")


def test_general_empresa_sin_pesos_raya() -> None:
    p = pp.general("mensual")
    p["empresa"]["ingresos_propios_mxn"] = None
    p["empresa"]["resultado_mxn"] = None
    data = render_balance_general_xlsx(BalanceGeneralRequest.model_validate(p))
    ws = load_workbook(BytesIO(data))["balance"]
    ing = next(c.row for c in ws["A"] if (c.value or "").startswith("(+) INGRESOS PROPIOS MXN ("))
    res = _fila_de(ws, "RESULTADO VUELATOUR (MXN)")
    assert ws.cell(row=ing, column=2).value == "—"
    assert ws.cell(row=res, column=2).value == "—"
    assert ws.cell(row=res, column=2).comment.text == _NOTA_SIN_PESOS_DERIVADA


def test_general_solo_cambia_la_hoja_balance() -> None:
    p = pp.general("mensual")
    sin = copy.deepcopy(p)
    sin.pop("balance_en_pesos")
    con_data = render_balance_general_xlsx(BalanceGeneralRequest.model_validate(p))
    sin_data = render_balance_general_xlsx(BalanceGeneralRequest.model_validate(sin))
    assert _hojas_distintas(con_data, sin_data) == {"balance"}


# ---------------------------------------------------------------------------
# 5) Vista en pantalla (el MISMO libro) y rutas.
# ---------------------------------------------------------------------------


def _hoja_json(libro: dict, nombre: str) -> dict:
    return next(h for h in libro["hojas"] if h["nombre"] == nombre)


def _textos_json(hoja: dict) -> set[str]:
    return {c.get("t") for c in hoja["celdas"] if c.get("t")}


def test_vista_en_pantalla_hereda_los_pesos() -> None:
    p = pp.individual()
    vista = libro_balance_avion(BalanceAvionRequest.model_validate(p))
    textos = _textos_json(_hoja_json(vista, "balance"))
    assert _FILA_COBRADA_PESOS in textos
    assert _ENCABEZADO_USD_INFO in textos
    assert f"{p['balance']['utilidad_cobrada_mxn']:,.2f}" in textos


TOKEN = "tok-balance-pesos"
client = TestClient(app)


def test_rutas_aceptan_el_payload_0093(monkeypatch, request) -> None:
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    request.addfinalizer(get_settings.cache_clear)
    headers = {"X-Internal-Token": TOKEN}
    res = client.post("/pdf/balance-avion-xlsx", json=pp.individual(), headers=headers)
    assert res.status_code == 200, res.text
    ws = load_workbook(BytesIO(res.content))["balance"]
    assert _filas_de(ws, _FILA_COBRADA_PESOS)
    res = client.post("/pdf/balance-general-xlsx", json=pp.general("general"), headers=headers)
    assert res.status_code == 200, res.text
    assert _filas_de(load_workbook(BytesIO(res.content))["balance"], _FILA_COBRADA_PESOS)
    res = client.post("/pdf/balance-general-libro", json=pp.general("mensual"), headers=headers)
    assert res.status_code == 200, res.text
    assert _FILA_COBRADA_PESOS in _textos_json(_hoja_json(res.json(), "balance"))
