"""Reporte por vuelo (2-sep-2026): el combustible del AVIÓN se rotula
'Gasavión / Turbosina' (ya no 'Gasolina', que ahora nombra la categoría de
vehículos) y la columna Categoría de los gastos imprime la etiqueta amable
cuando el API la manda, o el código del enum si no. El PDF se valida por su
HTML (WeasyPrint no hace falta)."""

import re
from io import BytesIO

from openpyxl import load_workbook

from app.schemas.reportes import ReporteVueloRequest
from app.services.reporte_vuelo_pdf import _build_html
from app.services.reporte_vuelo_xlsx import render_reporte_vuelo_xlsx


def _req() -> ReporteVueloRequest:
    return ReporteVueloRequest(
        generado="2026-09-02",
        folio="131",
        cliente="Cliente Demo",
        total_usd=1160.0,
        iva_usd=160.0,
        tc_usd_mxn=18.0,
        combustible=[
            {"fecha": "2026-08-20", "detalle": "CUN · 100 L", "moneda": "MXN",
             "monto": 3000.0, "litros": 100.0},
        ],
        combustible_total_usd=166.67,
        gastos=[
            {"fecha": "2026-08-20", "concepto": "OTRO",
             "etiqueta": "Otros gastos VuelaTour", "moneda": "MXN", "monto": 500.0},
            # Sin etiqueta (API viejo): se imprime el código.
            {"fecha": "2026-08-20", "concepto": "TUAS", "moneda": "MXN", "monto": 800.0},
        ],
        gastos_total_usd=72.22,
        remanente_usd=761.11,
        ganancia_final_usd=761.11,
    )


def _valores(xlsx: bytes) -> list[str]:
    ws = load_workbook(BytesIO(xlsx)).active
    return [str(c.value) for fila in ws.iter_rows() for c in fila if c.value is not None]


def test_xlsx_rotula_gasavion_turbosina_y_etiquetas() -> None:
    valores = _valores(render_reporte_vuelo_xlsx(_req()))
    assert "Gasavión / Turbosina" in valores
    assert "Total gasavión / turbosina USD" in valores
    assert "(−) Gasavión / Turbosina" in valores
    assert not any("gasolina" in v.lower() for v in valores)
    assert "Otros gastos VuelaTour" in valores
    assert "OTRO" not in valores
    assert "TUAS" in valores


def test_pdf_html_rotula_gasavion_turbosina_y_etiquetas() -> None:
    html = _build_html(_req())
    assert "Total gasavión / turbosina:" in html
    assert "(−) Gasavión / Turbosina" in html
    assert "gasolina" not in html.lower()
    assert "Otros gastos VuelaTour" in html
    assert "<td>OTRO</td>" not in html
    assert "<td>TUAS</td>" in html


# ---------------------------------------------------------------------------
# Pago al VENDEDOR con el API 0.0.75 (invariante 73 del API; revisión del lote
# 2, 14-oct-2026). Con la bandera `comision_vendedor_solo_gasto_real`,
# `pago_vendedor_usd` es el pago REAL (0 sin gasto capturado) y
# `pago_vendedor_cotizado_usd` la cotizada con su IVA. La sección
# «Cotización» pinta la COTIZADA y su neto (lo mismo que el panel) y la
# cascada rotula «(−) Pago real al vendedor» sin «(c/IVA)». Sin la bandera,
# el reporte de siempre.
# ---------------------------------------------------------------------------


def _req_vendedor(**extra) -> ReporteVueloRequest:
    base = dict(
        generado="2026-10-14",
        folio="233",
        cliente="Cliente Demo",
        total_usd=4640.0,
        iva_usd=640.0,
        venta_sin_iva_usd=4000.0,
        remanente_usd=4000.0,
        ganancia_final_usd=4000.0,
    )
    base.update(extra)
    return ReporteVueloRequest(**base)


def _celdas(xlsx: bytes) -> dict[str, object]:
    ws = load_workbook(BytesIO(xlsx)).active
    out: dict[str, object] = {}
    for fila in ws.iter_rows():
        for c in fila:
            if c.value is not None and isinstance(c.value, str):
                # Valor de la columna B en la misma fila.
                out[c.value] = ws.cell(row=c.row, column=2).value
    return out


def test_esquema_acepta_los_dos_campos_aditivos_del_api_0075() -> None:
    r = _req_vendedor(comision_vendedor_solo_gasto_real=True, pago_vendedor_cotizado_usd=580.0)
    assert r.comision_vendedor_solo_gasto_real is True
    assert r.pago_vendedor_cotizado_usd == 580.0
    viejo = _req_vendedor()
    assert viejo.comision_vendedor_solo_gasto_real is None
    assert viejo.pago_vendedor_cotizado_usd is None


def test_0075_sin_pago_real_la_cotizacion_pinta_la_cotizada_y_su_neto() -> None:
    """#233-like: comisión cotizada $500 ($580 con IVA) y ningún gasto real."""
    r = _req_vendedor(
        comision_vendedor_usd=500.0,
        comision_vendedor_nombre="Juan",
        pago_vendedor_usd=0.0,
        neto_vuelatour_usd=4640.0,
        comision_vendedor_solo_gasto_real=True,
        pago_vendedor_cotizado_usd=580.0,
    )
    html = _build_html(r)
    assert "Pago comisión vendedor (Juan) (comisión + IVA)" in html
    assert "&minus;$580.00 USD" in html
    assert "<b>$4,060.00 USD</b>" in html  # Neto VuelaTour = total − cotizada
    assert "&minus;$0.00 USD" not in html
    # Sin pago real la cascada no lo resta (ni lo rotula).
    assert "Pago real al vendedor" not in html
    assert not re.search(r"vendedor[^<]*\(c/IVA\)", html)
    celdas = _celdas(render_reporte_vuelo_xlsx(r))
    assert celdas["Pago comisión vendedor (Juan) (comisión + IVA)"] == -580
    assert celdas["Neto VuelaTour USD"] == 4060
    assert not any("Pago real al vendedor" in k for k in celdas)


def test_0075_pago_real_mayor_que_la_comision_sin_iva_no_dice_c_iva() -> None:
    """Pago real de $200 contra una comisión de $140 sin IVA (pago en pesos
    convertido): la cotización dice $140 y la cascada «Pago real»."""
    r = _req_vendedor(
        comision_vendedor_usd=140.0,
        comision_vendedor_nombre="Pablo Canales",
        pago_vendedor_usd=200.0,
        neto_vuelatour_usd=4440.0,
        ganancia_final_usd=3800.0,
        comision_vendedor_solo_gasto_real=True,
        pago_vendedor_cotizado_usd=140.0,
    )
    html = _build_html(r)
    assert "Pago comisión vendedor (Pablo Canales)</td>" in html
    assert "(comisión + IVA)" not in html
    assert "&minus;$140.00 USD" in html
    assert "<b>$4,500.00 USD</b>" in html  # 4,640 − 140
    assert "(−) Pago real al vendedor (Pablo Canales)" in html
    assert not re.search(r"vendedor[^<]*\(c/IVA\)", html)
    assert "&minus;$200.00 USD" in html
    celdas = _celdas(render_reporte_vuelo_xlsx(r))
    assert celdas["Pago comisión vendedor (Pablo Canales)"] == -140
    assert celdas["Neto VuelaTour USD"] == 4500
    assert celdas["(−) Pago real al vendedor (Pablo Canales)"] == -200
    assert not any(
        "vendedor" in k and ("(c/IVA)" in k or "(comisión + IVA)" in k) for k in celdas
    )


def test_0075_pago_real_sin_comision_cotizada_caso_332() -> None:
    """#332: sin comisión cotizada y $1,305 MXN pagados (= $70.54 USD)."""
    r = _req_vendedor(
        pago_vendedor_usd=70.54,
        neto_vuelatour_usd=4569.46,
        ganancia_final_usd=3929.46,
        comision_vendedor_solo_gasto_real=True,
        pago_vendedor_cotizado_usd=None,
    )
    html = _build_html(r)
    assert "Pago comisión vendedor" not in html  # la cotización no la tiene
    assert "Neto VuelaTour" not in html
    assert "(−) Pago real al vendedor</td>" in html
    assert not re.search(r"vendedor[^<]*\(c/IVA\)", html)
    assert "&minus;$70.54 USD" in html
    celdas = _celdas(render_reporte_vuelo_xlsx(r))
    assert celdas["(−) Pago real al vendedor"] == -70.54
    assert "Neto VuelaTour USD" not in celdas


def test_sin_bandera_el_reporte_de_siempre_api_0074() -> None:
    """API ≤ 0.0.74: `pago_vendedor_usd` es el COTIZADO con IVA."""
    r = _req_vendedor(
        comision_vendedor_usd=500.0,
        comision_vendedor_nombre="Juan",
        pago_vendedor_usd=580.0,
        neto_vuelatour_usd=4060.0,
        ganancia_final_usd=3420.0,
    )
    html = _build_html(r)
    assert "Pago comisión vendedor (Juan) (comisión + IVA)" in html
    assert "<b>$4,060.00 USD</b>" in html
    assert "(−) Pago comisión vendedor (Juan) (c/IVA)" in html
    assert "Pago real al vendedor" not in html
    celdas = _celdas(render_reporte_vuelo_xlsx(r))
    assert celdas["(−) Pago comisión vendedor (Juan) (c/IVA)"] == -580
    # Un payload del 0.0.75 con la bandera en False tampoco cambia nada.
    falso = _req_vendedor(
        comision_vendedor_usd=500.0,
        comision_vendedor_nombre="Juan",
        pago_vendedor_usd=580.0,
        neto_vuelatour_usd=4060.0,
        ganancia_final_usd=3420.0,
        comision_vendedor_solo_gasto_real=False,
        pago_vendedor_cotizado_usd=580.0,
    )
    assert _build_html(falso) == html
