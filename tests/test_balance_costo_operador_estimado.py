"""Costo del operador externo ESTIMADO (5-nov-2026, API 0.0.121, invariante
117 del API). Pedido de la oficina: «cuando aún no se le paga al operador
externo ponerlo en amarillo y que la nota diga que es el precio estimado; ya
después se pone el que sí se pagó cuando se suba el gasto».

El API manda `costo_operador_estimado: true` en la fila cuyo OPERACIONES
carga el precio cotizado (ningún gasto lo cubre) y la nota ya lo dice. Aquí:
la celda en ámbar —OPERACIONES en el Balance mensual (individual y flota),
TOTAL PARA PROVEEDOR (PESOS) en el general— con la nota de siempre, y la
leyenda al pie solo cuando alguna fila lo trae. Sin la bandera, nada cambia
(los libros de siempre siguen congelados por bytes en `test_balance_formulas`
y `test_balance_general_costo_hora`).
"""

from __future__ import annotations

from io import BytesIO

import pytest
from openpyxl import load_workbook

from app.schemas.reportes import BalanceAvionVuelo
from app.services.balance_avion_xlsx import (
    _CPH_NOTA_COSTO_OPERADOR_ESTIMADO,
    _NOTA_COSTO_OPERADOR_ESTIMADO,
    AMBER,
    render_balance_avion_xlsx,
    render_balance_general_xlsx,
)
from tests import test_balance_formulas as tf
from tests import test_balance_general_costo_hora as tg

NOTA = (
    "Costo operador externo (Simon Herrera) — $15,804.43 — ESTIMADO: es el precio "
    "cotizado al operador; el pago aún no está en Gastos (al registrarlo marcado "
    "como pago al operador, el balance toma lo pagado)"
)
AMBAR = "00" + AMBER
OPERACIONES = "OPERACIONES"
PROVEEDOR_PESOS = "TOTAL PARA\nPROVEEDOR (PESOS)"


def _estimado(v: BalanceAvionVuelo) -> BalanceAvionVuelo:
    """La fila como la manda el API: externo, OPERACIONES con el costo
    cotizado, su nota y la bandera."""
    return v.model_copy(
        update={
            "es_externo": True,
            "op_mxn": 15804.43,
            "op_detalle": [NOTA],
            "costo_operador_estimado": True,
        }
    )


def _textos_pie(ws) -> list[str]:
    return [c.value for c in ws["A"] if isinstance(c.value, str)]


def test_sin_la_llave_la_bandera_es_false() -> None:
    """API previo: la fila no trae la llave y nada se pinta."""
    assert tf._individual_tst().vuelos[0].costo_operador_estimado is False
    datos = tf._individual_tst().vuelos[0].model_dump()
    datos.pop("costo_operador_estimado")
    assert BalanceAvionVuelo.model_validate(datos).costo_operador_estimado is False


def test_balance_mensual_operaciones_en_ambar_con_nota_y_leyenda() -> None:
    req = tf._individual_tst()
    con = req.model_copy(update={"vuelos": [_estimado(req.vuelos[0]), *req.vuelos[1:]]})
    ws = load_workbook(BytesIO(render_balance_avion_xlsx(con)))["reporte horas XA-TST"]
    celda = tg._celda(ws, 3, OPERACIONES)
    assert tg._relleno(celda) == AMBAR
    assert celda.comment is not None and celda.comment.text == NOTA
    assert celda.value == 15804.43
    # Las demás filas, como siempre; la leyenda, una vez, al pie.
    for fila in range(4, 3 + len(req.vuelos)):
        assert tg._relleno(tg._celda(ws, fila, OPERACIONES)) != AMBAR, fila
    assert _textos_pie(ws).count(_NOTA_COSTO_OPERADOR_ESTIMADO) == 1
    assert _CPH_NOTA_COSTO_OPERADOR_ESTIMADO not in _textos_pie(ws)


def test_sin_bandera_ni_ambar_ni_leyenda() -> None:
    req = tf._individual_tst()
    ws = load_workbook(BytesIO(render_balance_avion_xlsx(req)))["reporte horas XA-TST"]
    for fila in range(3, 3 + len(req.vuelos)):
        assert tg._relleno(tg._celda(ws, fila, OPERACIONES)) != AMBAR, fila
    textos = _textos_pie(ws)
    assert _NOTA_COSTO_OPERADOR_ESTIMADO not in textos
    assert _CPH_NOTA_COSTO_OPERADOR_ESTIMADO not in textos


LEYENDA_MENSUAL = _NOTA_COSTO_OPERADOR_ESTIMADO
LEYENDA_GENERAL = _CPH_NOTA_COSTO_OPERADOR_ESTIMADO


@pytest.mark.parametrize(
    ("variante", "encabezado", "leyenda", "ajena"),
    [
        ("general", PROVEEDOR_PESOS, LEYENDA_GENERAL, LEYENDA_MENSUAL),
        ("mensual", OPERACIONES, LEYENDA_MENSUAL, LEYENDA_GENERAL),
    ],
)
def test_balance_general_pinta_la_celda_que_carga_el_costo(
    variante: str, encabezado: str, leyenda: str, ajena: str
) -> None:
    """Variante general: la hoja de la flota no lleva OPERACIONES; el costo
    vive en TOTAL PARA PROVEEDOR (PESOS) y su nota de desglose trae la línea
    ESTIMADO. Variante mensual del general: OPERACIONES, como el individual."""
    req = tg._general(variante)
    cons = req.consolidado
    vuelos = [_estimado(cons.vuelos[0]), *cons.vuelos[1:]]
    req = req.model_copy(update={"consolidado": cons.model_copy(update={"vuelos": vuelos})})
    ws = tg._libros(render_balance_general_xlsx(req))[0][tg.MAESTRA]
    celda = tg._celda(ws, 3, encabezado)
    assert tg._relleno(celda) == AMBAR
    assert celda.comment is not None and NOTA in celda.comment.text
    for fila in range(4, 3 + len(vuelos)):
        assert tg._relleno(tg._celda(ws, fila, encabezado)) != AMBAR, fila
    textos = _textos_pie(ws)
    assert textos.count(leyenda) == 1
    assert ajena not in textos
