"""Comisión del vendedor como GASTO (28-sep-2026, API 0.0.39, invariante 31).

Pedido del cliente: «¿cómo registro un gasto para que aparezca en la hoja de
otros movimientos? Como pagarle una comisión a Saab… si lo agrego como "Otros
gastos VuelaTour" me lo manda a la hoja de Otros gastos: queda duplicado.»

El API agrega la categoría «Comisión del vendedor» y el gasto real REEMPLAZA a
la PROVISIÓN en su fila de «otros movimientos» (Balance general) y de «Otros
ingresos» (Libro Dinero). Aquí solo se pinta: los conceptos y las notas de
celda llegan hechos. Lo único que cambia en pyservices son 5 leyendas, y SOLO
con las banderas que el API manda cuando hay pagos reales:

- Balance general: `otros_movimientos.hay_pago_vendedor_real` (fila 2 de
  'otros movimientos', nota del RESUMEN y nota de la hoja maestra del general).
- Libro Dinero: `utilidades_comision_vendedor_pagada_mxn` (fila 2 de 'Otros
  ingresos' y la frase nueva de la fila 9 de 'utilidades').

Sin bandera el Excel es IDÉNTICO al de antes (las leyendas de hoy siguen
siendo verdaderas mientras no haya gasto real): aquí se compara contra el
texto viejo LITERAL. El libro INDIVIDUAL nunca cambia.

6-oct-2026 (API 0.0.65, comisiones a cargo del avión): las notas del RESUMEN
y de la hoja maestra hablan ahora de la columna COMISIONES (antes «COMISIÓN
VENDEDOR MXN va vacía a propósito…»); la bandera sigue cambiando SOLO el
fragmento del pago al vendedor.
"""

from io import BytesIO

import pytest
from fastapi.testclient import TestClient
from openpyxl import load_workbook

from app.config import get_settings
from app.main import app
from app.schemas.reportes import (
    BalanceAvionRequest,
    BalanceGeneralRequest,
    BalanceHojaOtrosMovimientos,
    DineroXlsxRequest,
)
from app.services.balance_avion_xlsx import (
    render_balance_avion_xlsx,
    render_balance_general_xlsx,
)
from app.services.dinero_xlsx import render_dinero_xlsx

# ---------------------------------------------------------------------------
# Textos de HOY (copiados literal del código anterior al 28-sep-2026).
# ---------------------------------------------------------------------------

OM_FILA2_VIEJA = (
    "Ingreso de VuelaTour (no del avión): TUAs, extras, viáticos de "
    "pernocta y comisión del vendedor cobrados al cliente (con su IVA) vs "
    "lo pagado — el pago de la comisión al vendedor va apareado en la "
    "misma fila como PROVISIÓN a la fecha del vuelo mientras no exista el "
    "gasto real (lo dice la nota de la celda). UNA FILA POR "
    "VUELO: los ingresos van sumados en una celda y los egresos en otra; "
    "el desglose concepto por concepto está en el COMENTARIO de cada "
    "celda (pasa el cursor sobre el triángulo rojo). Una fila puede traer "
    "solo ingreso o solo egreso. Todo en MXN. Incluye todos "
    "los estados del periodo (igual que la hoja maestra); los cancelados "
    "se marcan (clave · CANCELADO en rojo) y los demás estados no "
    "normales llevan su estado junto a la clave. Los MOVIMIENTOS SIN "
    "AVIÓN / SIN VUELO de abajo son 'gas sin avión', 'tuas sin vuelo' y "
    "los OTROS INGRESOS registrados en Ingresos (clave ING-n: intereses, "
    "reembolsos, ventas de activos…); los anticipos y las aportaciones NO "
    "aparecen: no son resultado (29-ago: los gastos de empresa viven en "
    "la hoja 'otros gastos' — antes llamada 'gastos VuelaTour')."
)
# Desde el 6-oct-2026 (API 0.0.65) la base de estas dos notas es la de la
# columna COMISIONES; el fragmento final (pago apareado) es el de siempre.
RESUMEN_NOTA_VIEJA = (
    "COMISIONES = las de la columna COMISIONES de la hoja de vuelos de cada "
    "avión (regla de septiembre 2026, por fecha del vuelo: su parte de la "
    "comisión bancaria de los cobros y la comisión del vendedor como "
    "PROVISIÓN; en vuelos anteriores no hay). Lo cobrado al cliente por la "
    "comisión del vendedor sigue siendo INGRESO de VuelaTour y su pago al "
    "vendedor sale de VuelaTour; los dos viven en 'otros movimientos' "
    "(ingreso cobrado y pago apareado como PROVISIÓN a la fecha del vuelo)."
)
MAESTRA_NOTA_VIEJA = (
    "COMISIONES MXN (regla de septiembre 2026, por fecha del vuelo) = lo que "
    "absorbe el avión: su parte de la comisión bancaria de los cobros y la "
    "comisión del vendedor como PROVISIÓN (lo cobrado al cliente por ese "
    "concepto, comisión + IVA); la nota de cada celda dice cuáles son y en los "
    "vuelos anteriores va vacía. Lo cobrado al cliente por la comisión del "
    "vendedor sigue siendo ingreso de VuelaTour y su pago al vendedor sale de "
    "VuelaTour; ambos viven en 'otros movimientos' (el pago apareado como "
    "PROVISIÓN a la fecha del vuelo). GANANCIA de la fila = REMANENTE (VENTA − "
    "COSTO TOTAL) − COMISIONES."
)
# Prefijos con que se encuentran en la columna A.
RESUMEN_PREFIJO = "COMISIONES = las de la columna COMISIONES"
MAESTRA_PREFIJO = "COMISIONES MXN (regla de septiembre 2026"
DINERO_FILA2_VIEJA = (
    "Ingreso de VuelaTour (no del avión): TUAs, extras, pernocta y "
    "comisión del vendedor cobrados (con su IVA) vs lo pagado. El pago de "
    "la comisión al vendedor va apareado en su fila como PROVISIÓN a la "
    "fecha del vuelo mientras no exista el gasto real (regla 28-ago-2026). "
    "Al final, los otros ingresos que no son de un vuelo (clave ING-n, "
    "registrados en Ingresos). Los anticipos de clientes no están aquí: "
    "cuentan como cobro del vuelo al aplicarse."
)
UTILIDADES_FILA9_BASE = (
    "OTROS INGRESOS = ingreso de VuelaTour (TUAs/extras/pernocta/"
    "comisión del vendedor + su IVA), no del avión (regla 28-ago-2026), "
    "NETO de la provisión del pago al vendedor (comisión + su IVA), más "
    "los ingresos sin vuelo registrados en Ingresos (ING-n), netos de su "
    "comisión bancaria: ver hoja 'Otros ingresos'."
)

# ---------------------------------------------------------------------------
# Textos NUEVOS (contrato del 28-sep-2026, sección 5) — fragmento viejo ⇒ nuevo.
# ---------------------------------------------------------------------------

OM_FRAG_VIEJO = (
    "el pago de la comisión al vendedor va apareado en la misma fila como "
    "PROVISIÓN a la fecha del vuelo mientras no exista el gasto real (lo dice "
    "la nota de la celda)."
)
OM_FRAG_NUEVO = (
    "el pago de la comisión al vendedor va apareado en la misma fila: el GASTO "
    "REAL cuando se captura en Gastos con la categoría «Comisión del vendedor» "
    "ligado al vuelo (con «faltan $…» o «excede $…» si no cuadra con lo "
    "cobrado) o, mientras no se capture, una PROVISIÓN por el mismo monto a la "
    "fecha del vuelo (lo dice la nota de la celda). Ese pago NO se captura como "
    "«Otros gastos VuelaTour»: quedaría duplicado en la hoja 'otros gastos'."
)
RESUMEN_FRAG_VIEJO = "(ingreso cobrado y pago apareado como PROVISIÓN a la fecha del vuelo)."
RESUMEN_FRAG_NUEVO = (
    "(ingreso cobrado y pago apareado: el gasto real «Comisión del vendedor» o, "
    "mientras no se capture, la PROVISIÓN a la fecha del vuelo)."
)
MAESTRA_FRAG_VIEJO = "(el pago apareado como PROVISIÓN a la fecha del vuelo)."
MAESTRA_FRAG_NUEVO = (
    "(el pago apareado: el gasto real «Comisión del vendedor» o, mientras no se "
    "capture, la PROVISIÓN a la fecha del vuelo)."
)
DINERO_FRAG_VIEJO = (
    "El pago de la comisión al vendedor va apareado en su fila como PROVISIÓN a "
    "la fecha del vuelo mientras no exista el gasto real (regla 28-ago-2026)."
)
DINERO_FRAG_NUEVO = (
    "El pago de la comisión al vendedor va apareado en su fila: el gasto real "
    "(categoría «Comisión del vendedor» ligada al vuelo) o, mientras no se "
    "capture, una PROVISIÓN por el mismo monto a la fecha del vuelo (regla "
    "28-ago-2026)."
)


def _cambia(viejo: str, frag_viejo: str, frag_nuevo: str) -> str:
    """El texto nuevo = el viejo con SOLO el fragmento sustituido."""
    assert viejo.count(frag_viejo) == 1
    return viejo.replace(frag_viejo, frag_nuevo)


# ---------------------------------------------------------------------------
# Payloads (como los arma el API: conceptos y notas YA hechos).
# ---------------------------------------------------------------------------

CONCEPTO_PARCIAL = "pago comisión vendedor (Alex Saab) · gasto real · parcial: faltan $400.00 MXN"
NOTA_PAGO_REAL = (
    "GASTO REAL: pago al vendedor capturado como «Comisión del vendedor» ligado "
    "a este vuelo; reemplaza la provisión. En la hoja utilidades ya está "
    "descontado de \"otros ingresos\"."
)

_FILA_PROVISION = {
    "clave": "N4142R-317",
    "avion_color": "#FF0000",
    "estado": "COMPLETADO",
    "fecha_vuelo": "2026-09-20",
    "concepto_egreso": (
        "pago comisión vendedor (Alex Saab) · PROVISIÓN (mismo monto que lo "
        "cobrado: comisión + IVA; sin gasto real capturado)"
    ),
    "egreso_mxn": 2030.0,
    "fecha_egreso": "2026-09-20",
    "concepto_ingreso": "comisión vendedor (Alex Saab)",
    "ingreso_mxn": 2030.0,
    "fecha_ingreso": "2026-09-20",
    "remanente_mxn": 0.0,
}
_FILA_PAGO_REAL = {
    "clave": "N4142R-260",
    "avion_color": "#FF0000",
    "estado": "COMPLETADO",
    "fecha_vuelo": "2026-09-10",
    "concepto_egreso": CONCEPTO_PARCIAL,
    "egreso_mxn": 1200.0,
    "fecha_egreso": "2026-09-25",
    "concepto_ingreso": "comisión vendedor (Alex Saab)",
    "ingreso_mxn": 1600.0,
    "fecha_ingreso": "2026-09-10",
    "remanente_mxn": 400.0,
    "nota_egreso": NOTA_PAGO_REAL,
}


def _general(bandera=..., filas=None) -> BalanceGeneralRequest:
    """`bandera=...` ⇒ la clave NO viaja (así la manda el API sin pagos)."""
    om: dict = {"filas": filas if filas is not None else [_FILA_PROVISION]}
    if bandera is not ...:
        om["hay_pago_vendedor_real"] = bandera
    return BalanceGeneralRequest(
        periodo_desde="2026-09-01",
        periodo_hasta="2026-09-30",
        resumen=[{"matricula": "N4142R", "color": "#FF0000", "vuelos": 2}],
        consolidado=BalanceAvionRequest(
            matricula="FLOTA",
            periodo_desde="2026-09-01",
            periodo_hasta="2026-09-30",
            otros_movimientos=om,
        ),
    )


def _dinero(pagada=..., provision: float | None = 2030.0,
            filas=None) -> DineroXlsxRequest:
    datos: dict = {
        "periodo_desde": "2026-09-01",
        "periodo_hasta": "2026-09-30",
        "otros_ingresos": filas if filas is not None else [_FILA_PROVISION],
        "utilidades_otros_ingresos_mxn": 5000.0,
        "utilidades_comision_vendedor_provisionada_mxn": provision,
    }
    if pagada is not ...:
        datos["utilidades_comision_vendedor_pagada_mxn"] = pagada
    return DineroXlsxRequest(**datos)


def _libro(xlsx: bytes):
    return load_workbook(BytesIO(xlsx), data_only=True)


def _celda_que_empieza(ws, prefijo: str) -> str:
    """Valor de la ÚNICA celda de la columna A que empieza con `prefijo`."""
    hallados = [
        c.value
        for c in ws["A"]
        if isinstance(c.value, str) and c.value.startswith(prefijo)
    ]
    assert len(hallados) == 1, hallados
    return hallados[0]


def _leyendas_general(xlsx: bytes) -> tuple[str, str, str]:
    wb = _libro(xlsx)
    om = wb["otros movimientos"]["A2"].value
    resumen = _celda_que_empieza(wb["RESUMEN flota"], RESUMEN_PREFIJO)
    maestra = _celda_que_empieza(wb["reporte horas FLOTA"], MAESTRA_PREFIJO)
    return om, resumen, maestra


def _leyendas_dinero(xlsx: bytes) -> tuple[str, str]:
    wb = _libro(xlsx)
    return wb["Otros ingresos"]["A2"].value, wb["utilidades"]["A9"].value


# ---------------------------------------------------------------------------
# Esquemas: campos ADITIVOS (skew tolerante en ambos sentidos).
# ---------------------------------------------------------------------------


def test_esquemas_aceptan_los_campos_nuevos_con_default_none() -> None:
    assert BalanceHojaOtrosMovimientos().hay_pago_vendedor_real is None
    assert BalanceHojaOtrosMovimientos(hay_pago_vendedor_real=True).hay_pago_vendedor_real is True
    assert DineroXlsxRequest().utilidades_comision_vendedor_pagada_mxn is None
    assert (
        DineroXlsxRequest(utilidades_comision_vendedor_pagada_mxn=1600).utilidades_comision_vendedor_pagada_mxn
        == 1600.0
    )


# ---------------------------------------------------------------------------
# Balance general: SIN bandera ⇒ las 3 leyendas son las de HOY (literal).
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("bandera", [..., None, False], ids=["ausente", "null", "false"])
def test_general_sin_bandera_leyendas_identicas_a_hoy(bandera) -> None:
    om, resumen, maestra = _leyendas_general(render_balance_general_xlsx(_general(bandera)))
    assert om == OM_FILA2_VIEJA
    assert resumen == RESUMEN_NOTA_VIEJA
    assert maestra == MAESTRA_NOTA_VIEJA


def test_general_sin_otros_movimientos_resumen_y_maestra_de_hoy() -> None:
    """API viejo sin la pestaña: nada que marcar, textos de siempre."""
    req = _general()
    req.consolidado.otros_movimientos = None
    wb = _libro(render_balance_general_xlsx(req))
    assert "otros movimientos" not in wb.sheetnames
    assert _celda_que_empieza(wb["RESUMEN flota"], RESUMEN_PREFIJO) == RESUMEN_NOTA_VIEJA
    assert _celda_que_empieza(wb["reporte horas FLOTA"], MAESTRA_PREFIJO) == MAESTRA_NOTA_VIEJA


# ---------------------------------------------------------------------------
# Balance general: CON bandera ⇒ los textos nuevos tal cual.
# ---------------------------------------------------------------------------


def test_general_con_bandera_pinta_las_tres_leyendas_nuevas() -> None:
    om, resumen, maestra = _leyendas_general(
        render_balance_general_xlsx(_general(True, filas=[_FILA_PROVISION, _FILA_PAGO_REAL]))
    )
    assert om == _cambia(OM_FILA2_VIEJA, OM_FRAG_VIEJO, OM_FRAG_NUEVO)
    assert resumen == _cambia(RESUMEN_NOTA_VIEJA, RESUMEN_FRAG_VIEJO, RESUMEN_FRAG_NUEVO)
    assert maestra == _cambia(MAESTRA_NOTA_VIEJA, MAESTRA_FRAG_VIEJO, MAESTRA_FRAG_NUEVO)
    # El camino que confundió al cliente queda escrito en la hoja.
    assert "«Comisión del vendedor»" in om
    assert "quedaría duplicado en la hoja 'otros gastos'" in om
    assert "PROVISIÓN" in om  # la provisión sigue existiendo sin gasto real


def test_individual_nunca_cambia_aunque_traiga_la_bandera() -> None:
    """El libro INDIVIDUAL no pinta 'otros movimientos'; su nota de la hoja
    maestra es la de siempre pase lo que pase."""
    req = BalanceAvionRequest(
        matricula="N4142R",
        periodo_desde="2026-09-01",
        periodo_hasta="2026-09-30",
        otros_movimientos={"filas": [_FILA_PAGO_REAL], "hay_pago_vendedor_real": True},
    )
    wb = _libro(render_balance_avion_xlsx(req))
    assert "otros movimientos" not in wb.sheetnames
    maestra = _celda_que_empieza(wb["reporte horas N4142R"], MAESTRA_PREFIJO)
    assert maestra == MAESTRA_NOTA_VIEJA


def test_otros_movimientos_pinta_el_concepto_del_gasto_real_tal_cual() -> None:
    wb = _libro(render_balance_general_xlsx(_general(True, filas=[_FILA_PAGO_REAL])))
    ws = wb["otros movimientos"]
    assert ws["A4"].value == "N4142R-260"
    assert ws["C4"].value == CONCEPTO_PARCIAL
    assert ws["D4"].value == 1200.0
    assert ws["E4"].value == "25/09/2026"
    assert ws["G4"].value == 1600.0
    assert ws["I4"].value == 400.0
    assert ws["D4"].comment.text == NOTA_PAGO_REAL


# ---------------------------------------------------------------------------
# Libro Dinero.
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("provision", [2030.0, 0.0, None])
def test_dinero_sin_campo_leyendas_identicas_a_hoy(provision) -> None:
    fila2, fila9 = _leyendas_dinero(render_dinero_xlsx(_dinero(provision=provision)))
    assert fila2 == DINERO_FILA2_VIEJA
    nota = " Provisión restada en este periodo: $2,030.00 MXN." if provision else ""
    assert fila9 == UTILIDADES_FILA9_BASE + nota


def test_dinero_con_pagado_en_cero_fila9_de_hoy_y_leyenda_nueva() -> None:
    """El campo presente (aun en 0) dice que hubo gasto real ⇒ la leyenda de
    la hoja cambia; la frase de la fila 9 solo aparece con monto > 0."""
    fila2, fila9 = _leyendas_dinero(render_dinero_xlsx(_dinero(pagada=0.0)))
    assert fila2 == _cambia(DINERO_FILA2_VIEJA, DINERO_FRAG_VIEJO, DINERO_FRAG_NUEVO)
    assert fila9 == UTILIDADES_FILA9_BASE + " Provisión restada en este periodo: $2,030.00 MXN."


def test_dinero_con_pagado_agrega_la_frase_despues_de_la_provision() -> None:
    fila2, fila9 = _leyendas_dinero(
        render_dinero_xlsx(_dinero(pagada=1600.0, filas=[_FILA_PROVISION, _FILA_PAGO_REAL]))
    )
    assert fila2 == _cambia(DINERO_FILA2_VIEJA, DINERO_FRAG_VIEJO, DINERO_FRAG_NUEVO)
    assert fila9 == (
        UTILIDADES_FILA9_BASE
        + " Provisión restada en este periodo: $2,030.00 MXN."
        + " Pagado al vendedor con gasto real en este periodo: $1,600.00 MXN."
    )


def test_dinero_solo_pagos_reales_sin_provision_viva() -> None:
    """Todos los vuelos del periodo con gasto real: provisión 0 ⇒ solo la
    frase de lo pagado (dinero con 2 decimales y moneda)."""
    _, fila9 = _leyendas_dinero(
        render_dinero_xlsx(_dinero(pagada=12345.5, provision=0.0, filas=[_FILA_PAGO_REAL]))
    )
    assert fila9 == (
        UTILIDADES_FILA9_BASE
        + " Pagado al vendedor con gasto real en este periodo: $12,345.50 MXN."
    )


def test_dinero_pinta_el_concepto_y_la_nota_del_gasto_real_tal_cual() -> None:
    wb = _libro(render_dinero_xlsx(_dinero(pagada=1200.0, filas=[_FILA_PAGO_REAL])))
    ws = wb["Otros ingresos"]
    assert ws["C4"].value == CONCEPTO_PARCIAL
    assert ws["D4"].value == 1200.0
    assert ws["I4"].value == 400.0
    assert ws["D4"].comment.text == NOTA_PAGO_REAL


# ---------------------------------------------------------------------------
# Rutas: el payload del API 0.0.39 entra sin 422 y respeta las banderas.
# ---------------------------------------------------------------------------

TOKEN = "secreto-de-prueba"
client = TestClient(app)


def test_rutas_aceptan_el_payload_nuevo(monkeypatch) -> None:
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    headers = {"X-Internal-Token": TOKEN}

    res = client.post(
        "/pdf/balance-general-xlsx",
        json=_general(True, filas=[_FILA_PAGO_REAL]).model_dump(mode="json"),
        headers=headers,
    )
    assert res.status_code == 200, res.text
    om, _, _ = _leyendas_general(res.content)
    assert om == _cambia(OM_FILA2_VIEJA, OM_FRAG_VIEJO, OM_FRAG_NUEVO)

    res = client.post(
        "/pdf/dinero-xlsx",
        json=_dinero(pagada=1200.0, filas=[_FILA_PAGO_REAL]).model_dump(mode="json"),
        headers=headers,
    )
    assert res.status_code == 200, res.text
    _, fila9 = _leyendas_dinero(res.content)
    assert fila9.endswith(" Pagado al vendedor con gasto real en este periodo: $1,200.00 MXN.")
