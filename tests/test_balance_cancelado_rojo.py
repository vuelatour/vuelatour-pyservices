"""Vuelo CANCELADO en la hoja de vuelos (5-nov-2026, pedido de la oficina con
su captura del Balance de octubre): «en la ruta cuando el vuelo sea cancelado
aparezca en rojito… al igual que en los tacómetros se ponga en ese colorcito».

RUTA en rojo (como el ESTADO) y el bloque TIEMPO / TACÓMETRO con relleno rojo
suave (`FILL_CANCELADO`: no voló, sin horas ni tacos); en la variante general,
TIEMPO CALZOS/HOBS (HR), que repite TIEMPO VUELO HR. Las demás filas, como
siempre. Las huellas de bytes de `test_balance_general_costo_hora` se
regeneraron a propósito (cambian solo las hojas de vuelos y styles.xml).
"""

from __future__ import annotations

from io import BytesIO

import pytest
from openpyxl import load_workbook

from app.services.balance_avion_xlsx import (
    FILL_CANCELADO,
    RED,
    render_balance_avion_xlsx,
    render_balance_general_xlsx,
)
from tests import test_balance_formulas as tf
from tests import test_balance_general_costo_hora as tg

ROJO_SUAVE = "00" + FILL_CANCELADO
ROJO = "00" + RED
TACOS = ("TIEMPO\nVUELO HR", "TACO\nINICIO", "TACO\nFINAL")


def _color_fuente(celda) -> str | None:
    color = celda.font.color if celda.font is not None else None
    return color.rgb if color is not None and isinstance(color.rgb, str) else None


def _comprobar(ws, vuelos, tacos: tuple[str, ...]) -> None:
    assert any(v.estado == "CANCELADO" for v in vuelos)
    for fila, v in enumerate(vuelos, start=3):
        ruta = tg._celda(ws, fila, "RUTA")
        estado = tg._celda(ws, fila, "ESTADO")
        if v.estado == "CANCELADO":
            assert ruta.value == v.ruta
            assert _color_fuente(ruta) == ROJO, fila
            assert _color_fuente(estado) == ROJO, fila
            for enc in tacos:
                assert tg._relleno(tg._celda(ws, fila, enc)) == ROJO_SUAVE, (fila, enc)
        else:
            assert _color_fuente(ruta) != ROJO, fila
            for enc in tacos:
                assert tg._relleno(tg._celda(ws, fila, enc)) != ROJO_SUAVE, (fila, enc)


def test_individual_cancelado_ruta_roja_y_tacos_en_rojo_suave() -> None:
    req = tf._individual_tst()
    ws = load_workbook(BytesIO(render_balance_avion_xlsx(req)))["reporte horas XA-TST"]
    _comprobar(ws, req.vuelos, TACOS)


@pytest.mark.parametrize(
    ("variante", "tacos"),
    [("mensual", TACOS), ("general", (tg.COSTO_POR_HORA[1],))],
)
def test_general_cancelado_en_las_dos_variantes(variante: str, tacos: tuple[str, ...]) -> None:
    req = tg._general(variante)
    ws = tg._libros(render_balance_general_xlsx(req))[0][tg.MAESTRA]
    _comprobar(ws, req.consolidado.vuelos, tacos)


def test_las_senales_de_siempre_ganan_al_rojo_suave() -> None:
    """Un cancelado con salto de taco interno (raro, pero posible): el ámbar
    de la señal sigue mandando sobre el rojo suave en TIEMPO VUELO HR."""
    req = tf._individual_tst()
    vuelos = [
        v.model_copy(update={"salto_taco_interno": True, "salto_taco_interno_detalle": "x"})
        if v.estado == "CANCELADO"
        else v
        for v in req.vuelos
    ]
    req = req.model_copy(update={"vuelos": vuelos})
    ws = load_workbook(BytesIO(render_balance_avion_xlsx(req)))["reporte horas XA-TST"]
    fila = next(i for i, v in enumerate(vuelos, start=3) if v.estado == "CANCELADO")
    assert tg._relleno(tg._celda(ws, fila, TACOS[0])) == "00" + tg.AMBER
    assert tg._relleno(tg._celda(ws, fila, TACOS[1])) == ROJO_SUAVE
