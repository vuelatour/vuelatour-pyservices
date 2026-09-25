"""Tests del cardex formato LIBRO: encabezado doble, filas por lado y totales.

25-sep-2026 (API 0.0.36): costo = último precio de compra y montos en pesos
con el T.C. oficial de su día; el API lo explica en `nota`, que va en el
subtítulo tras «Montos en MXN». Sin `nota` el libro sale idéntico."""

import hashlib
import json
from io import BytesIO

from openpyxl import load_workbook

from app.schemas.cardex_libro import CardexLibroRequest
from app.services.cardex_libro_xlsx import render_cardex_libro_xlsx


def _req() -> CardexLibroRequest:
    return CardexLibroRequest(
        titulo="Cardex — Filtro 108-1",
        item_nombre="Filtro 108-1",
        numero_parte="108-1",
        unidad="pieza",
        generado="2026-08-29",
        moneda="MXN",
        entradas=[
            {
                "fecha": "2026-08-01",
                "cantidad": 10,
                "descripcion": "Filtro 108-1 · Aircraft Spruce",
                "valor_compra_unitario": 500.0,
                "valor_compra_total": 5000.0,
                "stock_despues": 10,
            },
        ],
        salidas=[
            {
                "fecha": "2026-08-10",
                "cantidad": 2,
                "descripcion": "Filtro 108-1",
                "venta_unitaria": 800.0,
                "venta_total": 1600.0,
                "remanente": 8,
                "ganancia": 600.0,
                "vendido_a": "XB-ABC",
            },
            {
                "fecha": "2026-08-12",
                "cantidad": 1,
                "descripcion": "Filtro 108-1 · a costo FIFO",
                "venta_unitaria": 500.0,
                "venta_total": 500.0,
                "remanente": 7,
                "ganancia": 0.0,
                "vendido_a": "FLOTA",
            },
        ],
        total_compra=5000.0,
        total_venta=2100.0,
        total_ganancia=600.0,
    )


def _hoja(xlsx: bytes):
    return load_workbook(BytesIO(xlsx)).active


def test_encabezado_doble_y_columnas() -> None:
    ws = _hoja(render_cardex_libro_xlsx(_req()))
    # Fila 4: bloques; fila 5: columnas de cada lado.
    assert ws.cell(row=4, column=1).value == "ENTRADAS"
    assert ws.cell(row=4, column=7).value == "SALIDAS"
    assert ws.cell(row=5, column=1).value == "Fecha entrada"
    assert ws.cell(row=5, column=6).value == "Cantidad en stock"
    assert ws.cell(row=5, column=7).value == "Fecha de salida"
    assert ws.cell(row=5, column=10).value == "Valor unitario al que se vendió"
    assert ws.cell(row=5, column=13).value == "Ganancia"
    assert ws.cell(row=5, column=14).value == "A quién se le vendió"


def test_filas_por_lado_y_totales() -> None:
    ws = _hoja(render_cardex_libro_xlsx(_req()))
    # Fila 6: entrada 1 a la izquierda, salida 1 a la derecha.
    assert ws.cell(row=6, column=1).value == "2026-08-01"
    assert ws.cell(row=6, column=5).value == 5000.0
    assert ws.cell(row=6, column=6).value == 10
    assert ws.cell(row=6, column=7).value == "2026-08-10"
    assert ws.cell(row=6, column=10).value == 800.0
    assert ws.cell(row=6, column=13).value == 600.0
    assert ws.cell(row=6, column=14).value == "XB-ABC"
    # Fila 7: sin entrada (lado izquierdo vacío), salida 2 a la derecha.
    assert ws.cell(row=7, column=1).value is None
    assert ws.cell(row=7, column=14).value == "FLOTA"
    # Fila 8: totales (compra | venta | ganancia), los mandó el API.
    assert ws.cell(row=8, column=5).value == 5000.0
    assert ws.cell(row=8, column=11).value == 2100.0
    assert ws.cell(row=8, column=13).value == 600.0


# ---------------------------------------------------------------------------
# Regla ÚLTIMO PRECIO + T.C. OFICIAL DEL DÍA (API 0.0.36): `nota` ADITIVA.
# ---------------------------------------------------------------------------

_NOTA_LIBRO = ("Compras al T.C. oficial del día de la compra; ventas al del "
               "día de la venta. Costo = último precio de compra vigente ese día.")


def _firma(xlsx: bytes, *, sin: tuple[str, ...] = ()) -> str:
    """Huella de TODO lo visible de la hoja (valores, formatos, fuentes,
    rellenos, alineación, merges, anchos y panel fijo), salvo las celdas
    `sin`."""
    ws = _hoja(xlsx)
    partes: list = []
    for fila in ws.iter_rows():
        for c in fila:
            if c.coordinate in sin:
                continue
            f = c.font
            partes.append([
                c.coordinate, repr(c.value), c.number_format, bool(f.b),
                bool(f.i), f.sz, f.color.rgb if f.color is not None else None,
                c.fill.fgColor.rgb if c.fill.fill_type else None,
                c.alignment.horizontal, bool(c.alignment.wrap_text),
                c.border.right.style,
            ])
    partes.append(sorted(str(r) for r in ws.merged_cells.ranges))
    partes.append(sorted((k, d.width) for k, d in ws.column_dimensions.items()))
    partes.append(ws.freeze_panes)
    return hashlib.sha256(json.dumps(partes, default=str).encode()).hexdigest()


def test_sin_nota_el_subtitulo_es_el_de_siempre() -> None:
    ws = _hoja(render_cardex_libro_xlsx(_req()))
    assert ws.cell(row=2, column=1).value == (
        "Parte 108-1 · Unidad: pieza · Montos en MXN · Generado 2026-08-29")
    # null, vacía o solo espacios = sin nota (NestJS puede mandar cualquiera).
    antes = _firma(render_cardex_libro_xlsx(_req()))
    for vacia in (None, "", "   "):
        req = _req().model_copy(update={"nota": vacia})
        assert _firma(render_cardex_libro_xlsx(req)) == antes


def test_nota_del_api_va_en_el_subtitulo_tras_montos_en_mxn() -> None:
    req = CardexLibroRequest.model_validate({**_req().model_dump(), "nota": _NOTA_LIBRO})
    xlsx = render_cardex_libro_xlsx(req)
    ws = _hoja(xlsx)
    assert ws.cell(row=2, column=1).value == (
        "Parte 108-1 · Unidad: pieza · Montos en MXN · Compras al T.C. oficial "
        "del día de la compra; ventas al del día de la venta. Costo = último "
        "precio de compra vigente ese día · Generado 2026-08-29")
    # Solo cambia el subtítulo: bloques, filas, totales y estilos intactos.
    assert _firma(xlsx, sin=("A2",)) == _firma(
        render_cardex_libro_xlsx(_req()), sin=("A2",))


def test_libro_real_aceite_sae_50_compra_en_dolares_ya_en_pesos() -> None:
    """Caso de la captura del cliente (25-sep-2026): «Aceite mineral
    aeronáutico SAE 50», 9 qt × 21.25 USD comprados el 29-ago. Con el API
    0.0.36 la compra llega en PESOS al T.C. oficial de ese día (17.0115):
    191.25 USD → $3,253.45 MXN (antes: «$0.00 MXN», sin T.C.)."""
    req = CardexLibroRequest(
        titulo="Cardex — Aceite mineral aeronáutico SAE 50",
        item_nombre="Aceite mineral aeronáutico SAE 50",
        unidad="cuarto (qt)",
        generado="2026-09-25",
        moneda="MXN",
        nota=_NOTA_LIBRO,
        entradas=[{
            "fecha": "2026-08-29", "cantidad": 9,
            "descripcion": "Aceite mineral aeronáutico SAE 50 · carga VTF-INV-001",
            "valor_compra_unitario": 361.49, "valor_compra_total": 3253.45,
            "stock_despues": 9,
        }],
        salidas=[],
        total_compra=3253.45,
        total_venta=0,
        total_ganancia=0,
    )
    ws = _hoja(render_cardex_libro_xlsx(req))
    assert ws.cell(row=2, column=1).value.startswith(
        "Unidad: cuarto (qt) · Montos en MXN · Compras al T.C. oficial")
    assert [ws.cell(row=6, column=c).value for c in (1, 2, 4, 5, 6)] == [
        "2026-08-29", 9, 361.49, 3253.45, 9]
    # Sin ventas: el lado derecho vacío y la fila de totales justo abajo.
    assert ws.cell(row=6, column=7).value is None
    assert [ws.cell(row=7, column=c).value for c in (3, 5, 11, 13)] == [
        "TOTAL COMPRA", 3253.45, 0, 0]
    assert ws.cell(row=7, column=5).number_format == '"$"#,##0.00'
    textos = [c.value for fila in ws.iter_rows() for c in fila
              if isinstance(c.value, str)]
    assert not any("FIFO" in t for t in textos)
