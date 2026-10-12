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

7-oct-2026 (API 0.0.66): la hoja de vuelos del GENERAL ya no lleva la columna
COMISIONES; su nota dice dónde están y remite a la leyenda de 'otros
movimientos' para el pago al vendedor, así que la bandera ya no la cambia
(sigue cambiando la fila 2 de 'otros movimientos' y la nota del RESUMEN, que
ahora remite al libro individual). La del libro INDIVIDUAL, intacta.

12-oct-2026 (API 0.0.73, invariante 58): «provisión: los gastos que todavía
no se hacen no deben aparecer» (las filas en amarillo de la captura). El API
ya no asigna NINGÚN egreso al vendedor mientras no exista el gasto real
«Comisión del vendedor» ligado al vuelo (ni concepto, ni monto, ni fecha, ni
marca amarilla: la fila queda con su ingreso cobrado y remanente = ingreso) y
lo avisa con la bandera ADITIVA `sin_provision_vendedor` (en
`otros_movimientos` del general; al nivel raíz del payload individual). Con
ella, 6 textos nuevos EXACTOS del contrato más, de la revisión adversaria del
mismo día, la base del pie del bloque VUELATOUR (empresa) sin «o
provisionado» (`_NOTA_EMPRESA_BASE_SIN_PROVISION`); sin ella, los de hoy
byte a byte (un API previo sigue provisionando). El amarillo
(`_MARCA_PROVISION`) no se toca: con el API nuevo ya no llega la marca y no
se pinta nada.
"""

import json
from io import BytesIO

import pytest
from fastapi.testclient import TestClient
from openpyxl import load_workbook

from app.config import get_settings
from app.main import app
from app.schemas.reportes import (
    REGLA_COMISIONES_RESPALDO,
    BalanceAvionRequest,
    BalanceGeneralRequest,
    BalanceHojaOtrosMovimientos,
    DineroXlsxRequest,
)
from app.services.balance_avion_xlsx import (
    _EMPRESA_REGLA_COMISIONES,
    _EMPRESA_REGLA_COMISIONES_CUBRE_AVION,
    _EMPRESA_REGLA_COMISIONES_CUBRE_AVION_SIN_PROVISION,
    _LEYENDA_PROVISION,
    _MAESTRA_PAGO_VENDEDOR_SIN_PROVISION,
    _NOTA_BLOQUE_EMPRESA,
    _NOTA_EMPRESA_BASE,
    _NOTA_EMPRESA_BASE_SIN_PROVISION,
    _OM_PAGO_VENDEDOR_SIN_PROVISION,
    _OM_REGLA_COMISIONES_CUBRE_AVION,
    _OM_REGLA_COMISIONES_CUBRE_AVION_SIN_PROVISION,
    _RESUMEN_PAGO_VENDEDOR_SIN_PROVISION,
    _RESUMEN_REGLA_CUBRE_AVION,
    _RESUMEN_REGLA_CUBRE_AVION_SIN_PROVISION,
    AMBER,
    render_balance_avion_xlsx,
    render_balance_general_xlsx,
)
from app.services.dinero_xlsx import render_dinero_xlsx
from tests._evaluador_formulas import verificar_libro

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
# 7-oct-2026 (API 0.0.66): la del RESUMEN remite a la columna COMISIONES del
# libro individual (la hoja de vuelos del general ya no la lleva).
RESUMEN_NOTA_VIEJA = (
    "COMISIONES = las de la columna COMISIONES del libro individual de cada "
    "avión (regla de septiembre 2026, por fecha del vuelo: su parte de la "
    "comisión bancaria de los cobros y la comisión del vendedor como "
    "PROVISIÓN; en vuelos anteriores no hay); en este libro no van en la hoja "
    "de vuelos (ahí la GANANCIA va antes de comisiones): restan aquí y en la "
    "cascada de cada avión de la hoja 'balance'. Lo cobrado al cliente por la "
    "comisión del vendedor sigue siendo INGRESO de VuelaTour; ese ingreso y el "
    "pago al vendedor viven en 'otros movimientos' "
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
# Nota de la hoja de vuelos del GENERAL desde el 7-oct-2026: sin la columna
# COMISIONES y sin el fragmento del pago apareado (remite a la leyenda de
# 'otros movimientos'): con y sin bandera, la misma.
MAESTRA_GENERAL_NOTA = (
    "COMISIONES (regla de septiembre 2026, por fecha del vuelo): las que "
    "absorbe cada avión —su parte de la comisión bancaria de sus cobros y la "
    "comisión del vendedor como PROVISIÓN— no van en esta hoja: están en la "
    "columna COMISIONES del libro individual de cada avión (Reportes › Balance "
    "por avión) y ya van restadas en la cascada de cada avión de la hoja "
    "'balance'. Lo cobrado al cliente por la comisión del vendedor es ingreso "
    "de VuelaTour: ese ingreso, el pago al vendedor y la parte de VuelaTour de "
    "la comisión bancaria viven en 'otros movimientos' (su leyenda explica "
    "cada línea). GANANCIA de la fila = REMANENTE (VENTA − COSTO TOTAL), antes "
    "de comisiones."
)
# Prefijos con que se encuentran en la columna A.
RESUMEN_PREFIJO = "COMISIONES = las de la columna COMISIONES"
MAESTRA_PREFIJO = "COMISIONES MXN (regla de septiembre 2026"
MAESTRA_GENERAL_PREFIJO = "COMISIONES (regla de septiembre 2026"
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
# Fragmento de la nota de la hoja maestra del libro INDIVIDUAL (la única que
# lo lleva desde el 7-oct-2026; antes, con la bandera, también la del general).
MAESTRA_FRAG_VIEJO = "(el pago apareado como PROVISIÓN a la fecha del vuelo)."
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


def _general(bandera=..., filas=None, sin_provision=...) -> BalanceGeneralRequest:
    """`bandera=...` ⇒ la clave NO viaja (así la manda el API sin pagos);
    `sin_provision` = `otros_movimientos.sin_provision_vendedor` (API 0.0.73,
    12-oct-2026; `...` = no viaja, como en un API previo)."""
    om: dict = {"filas": filas if filas is not None else [_FILA_PROVISION]}
    if bandera is not ...:
        om["hay_pago_vendedor_real"] = bandera
    if sin_provision is not ...:
        om["sin_provision_vendedor"] = sin_provision
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
    maestra = _celda_que_empieza(wb["reporte horas FLOTA"], MAESTRA_GENERAL_PREFIJO)
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
    assert maestra == MAESTRA_GENERAL_NOTA


def test_general_sin_otros_movimientos_resumen_y_maestra_de_hoy() -> None:
    """API viejo sin la pestaña: nada que marcar, textos de siempre."""
    req = _general()
    req.consolidado.otros_movimientos = None
    wb = _libro(render_balance_general_xlsx(req))
    assert "otros movimientos" not in wb.sheetnames
    assert _celda_que_empieza(wb["RESUMEN flota"], RESUMEN_PREFIJO) == RESUMEN_NOTA_VIEJA
    maestra = _celda_que_empieza(wb["reporte horas FLOTA"], MAESTRA_GENERAL_PREFIJO)
    assert maestra == MAESTRA_GENERAL_NOTA


# ---------------------------------------------------------------------------
# Balance general: CON bandera ⇒ los textos nuevos tal cual.
# ---------------------------------------------------------------------------


def test_general_con_bandera_pinta_las_leyendas_nuevas() -> None:
    om, resumen, maestra = _leyendas_general(
        render_balance_general_xlsx(_general(True, filas=[_FILA_PROVISION, _FILA_PAGO_REAL]))
    )
    assert om == _cambia(OM_FILA2_VIEJA, OM_FRAG_VIEJO, OM_FRAG_NUEVO)
    assert resumen == _cambia(RESUMEN_NOTA_VIEJA, RESUMEN_FRAG_VIEJO, RESUMEN_FRAG_NUEVO)
    # 7-oct-2026: la nota de la hoja de vuelos del general ya no habla del
    # pago al vendedor (remite a la leyenda de arriba): no cambia con la bandera.
    assert maestra == MAESTRA_GENERAL_NOTA
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
    assert maestra.count(MAESTRA_FRAG_VIEJO) == 1


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


# ===========================================================================
# 12-oct-2026 (API 0.0.73, invariante 58): «provisión: los gastos que todavía
# no se hacen no deben aparecer». Con la bandera ADITIVA
# `sin_provision_vendedor` los textos EXACTOS del contrato; sin ella, los de
# hoy byte a byte. El amarillo no se toca.
# ===========================================================================

# Textos del CONTRATO, literales (partidos en otros puntos que en el código:
# lo que se compara es la cadena completa).
OM_FRAG_SIN_PROVISION = (
    "el pago de la comisión al vendedor va apareado en la misma fila SOLO cuando "
    "se captura en Gastos el gasto real «Comisión del vendedor» ligado al vuelo "
    "(con «faltan $…» o «excede $…» si no cuadra con lo cobrado); mientras no se "
    "capture, la fila no lleva egreso ni provisión (pedido del 12-oct-2026: los "
    "gastos que todavía no se hacen no aparecen). Ese pago NO se captura como "
    "«Otros gastos VuelaTour»: quedaría duplicado en la hoja 'otros gastos'."
)
RESUMEN_FRAG_SIN_PROVISION = (
    "(ingreso cobrado; el pago al vendedor se aparea solo cuando se captura el "
    "gasto real «Comisión del vendedor»: sin él no se provisiona nada en esa hoja)."
)
MAESTRA_FRAG_SIN_PROVISION = (
    "(el pago se aparea ahí cuando se captura el gasto real «Comisión del vendedor»)."
)
OM_REGLA_CUBRE_SIN_PROVISION = (
    " Con la {regla} (por fecha del vuelo), en los vuelos COMPLETADOS la comisión "
    "del vendedor la cubre el avión con su PROVISIÓN en su columna COMISIONES; "
    "aquí el pago al vendedor aparece solo cuando se captura el gasto real y resta "
    "únicamente lo que ese pago exceda la provisión del avión; la comisión "
    "bancaria de aquí es solo la parte de VuelaTour. Lo que absorbe cada avión (la "
    "provisión y su parte de la comisión bancaria) resta en la columna COMISIONES "
    "de su libro individual y en la cascada de la hoja 'balance'."
)
EMPRESA_REGLA_CUBRE_SIN_PROVISION = (
    " Con la {regla} (por fecha del vuelo), en los vuelos completados el pago al "
    "vendedor lo cubre cada avión con su provisión (los egresos propios solo traen "
    "lo que un pago real capturado la exceda) y, de la comisión bancaria, los "
    "egresos solo traen la parte de VuelaTour."
)
RESUMEN_REGLA_CUBRE_SIN_PROVISION = (
    " Con la {regla} (por fecha del vuelo), en los vuelos completados ese pago lo "
    "cubre el avión con su provisión: en 'otros movimientos' solo queda lo que un "
    "pago real capturado la exceda."
)
# Base del pie del bloque VUELATOUR (empresa) con la bandera (revisión
# adversaria del 12-oct-2026): la de hoy dice «lo pagado o provisionado» y con
# el API 0.0.73 los egresos propios ya no traen provisión alguna.
EMPRESA_BASE_SIN_PROVISION = (
    "Ingresos y egresos propios = hoja 'otros movimientos': lo cotizado de los "
    "vuelos del periodo (no cancelados) más los otros ingresos registrados, y lo "
    "pagado (el pago al vendedor solo cuando se captura el gasto real «Comisión "
    "del vendedor») — no es lo cobrado: el RESULTADO no es dinero en caja."
)
REGLA = REGLA_COMISIONES_RESPALDO

# Filas como las arma el API 0.0.73. Sin gasto real NO hay egreso (ni
# concepto, ni monto, ni fecha, ni marca); con gasto real, igual que hoy.
_FILA_V73_VIGENCIA_SIN_GASTO = {  # la captura del cliente: $12,148.35 cobrados
    "clave": "#412 · Cliente Doce",
    "avion_color": "#2563EB",
    "estado": "COMPLETADO",
    "fecha_vuelo": "2026-10-03",
    "concepto_egreso": None,
    "egreso_mxn": None,
    "fecha_egreso": None,
    "concepto_ingreso": "comisión vendedor (Pablo Canales)",
    "ingreso_mxn": 12148.35,
    "fecha_ingreso": "2026-10-03",
    "remanente_mxn": 12148.35,
}
_FILA_V73_CONFIRMADO_SIN_GASTO = {  # vuelo no realizado sin gasto: tampoco egreso
    "clave": "#415 · Cliente Quince",
    "avion_color": "#2563EB",
    "estado": "CONFIRMADO",
    "fecha_vuelo": "2026-10-20",
    "concepto_ingreso": "comisión vendedor (Alex Saab)",
    "ingreso_mxn": 2030.0,
    "fecha_ingreso": "2026-10-20",
    "remanente_mxn": 2030.0,
}
_FILA_V73_CUBIERTO_PARCIAL = {  # gasto real MENOR que la provisión: $0.00 sin marca
    "clave": "#409 · Cliente Nueve",
    "avion_color": "#16A34A",
    "estado": "COMPLETADO",
    "fecha_vuelo": "2026-10-01",
    "concepto_egreso": (
        "pago comisión vendedor (Pablo Canales) · cubierto por el avión XA-TST "
        "(provisión $2,360.00 en su balance) · gasto real $2,000.00 · parcial: "
        "faltan $360.00 MXN"
    ),
    "egreso_mxn": 0.0,
    "fecha_egreso": "2026-10-09",
    "concepto_ingreso": "comisión vendedor (Pablo Canales)",
    "ingreso_mxn": 2360.0,
    "fecha_ingreso": "2026-10-01",
    "remanente_mxn": 2360.0,
}
_FILA_V73_CUBIERTO_EXCEDE = {  # gasto real MAYOR que la provisión: solo el exceso
    "clave": "#407 · Cliente Cinco",
    "avion_color": "#16A34A",
    "estado": "COMPLETADO",
    "fecha_vuelo": "2026-09-18",
    "concepto_egreso": (
        "pago comisión vendedor (Alex Saab) · cubierto por el avión XB-DOS "
        "(provisión $1,180.00 en su balance) · gasto real $1,500.00 · excede "
        "$320.00 MXN"
    ),
    "egreso_mxn": 320.0,
    "fecha_egreso": "2026-09-25",
    "concepto_ingreso": "comisión vendedor (Alex Saab)",
    "ingreso_mxn": 1180.0,
    "fecha_ingreso": "2026-09-18",
    "remanente_mxn": 860.0,
}
# Los tres sabores con gasto real que siguen EXACTAMENTE como hoy: gasto real
# pre-vigencia (`_FILA_PAGO_REAL`, faltan $400), cubierto con exceso y
# cubierto con parcial (arriba).
_FILAS_V73 = [
    _FILA_V73_VIGENCIA_SIN_GASTO,
    _FILA_V73_CONFIRMADO_SIN_GASTO,
    _FILA_V73_CUBIERTO_PARCIAL,
    _FILA_V73_CUBIERTO_EXCEDE,
    _FILA_PAGO_REAL,
]


def _tg():
    """`tests.test_balance_comisiones_general` (fixtures del API 0.0.66: vuelos
    de la regla, `comisiones_usd` en los aviones, bloque VUELATOUR), importado
    PEREZOSO: `test_balance_formulas` importa este módulo y aquel importa
    `test_balance_formulas` — a nivel de módulo sería un import circular al
    recolectar la suite completa."""
    from tests import test_balance_comisiones_general as tg

    return tg


def _payload_regla_v73(variante: str = "mensual", *, sin_provision=True, filas=None) -> dict:
    """El general del API 0.0.66 (vuelos de la regla, `comisiones_usd` en los
    aviones, bloque VUELATOUR) con 'otros movimientos' como lo manda el
    0.0.73 (`sin_provision=...` ⇒ la clave no viaja)."""
    p = _tg()._payload_om_0066(variante)
    om = p["consolidado"]["otros_movimientos"]
    om["filas"] = filas if filas is not None else list(_FILAS_V73)
    if sin_provision is not ...:
        om["sin_provision_vendedor"] = sin_provision
    return p


def _rellenos_amber(ws) -> list[str]:
    """Coordenadas de las celdas de DATOS con relleno AMBER (la celda de la
    leyenda «Amarillo = provisión…» también va en ámbar y se verifica aparte
    por su texto)."""
    return [
        c.coordinate
        for fila in ws.iter_rows()
        for c in fila
        if c.value != _LEYENDA_PROVISION
        and c.fill is not None
        and c.fill.fill_type
        and c.fill.fgColor.rgb == "00" + AMBER
    ]


def _textos(ws) -> list[str]:
    return [c.value for fila in ws.iter_rows() for c in fila if isinstance(c.value, str)]


def _suma(filas: list[dict], llave: str) -> float:
    return round(sum(f.get(llave) or 0.0 for f in filas), 2)


def test_esquemas_aceptan_sin_provision_vendedor_con_default_none() -> None:
    assert BalanceHojaOtrosMovimientos().sin_provision_vendedor is None
    assert BalanceHojaOtrosMovimientos(sin_provision_vendedor=True).sin_provision_vendedor is True
    assert BalanceAvionRequest().sin_provision_vendedor is None
    assert BalanceAvionRequest(sin_provision_vendedor=True).sin_provision_vendedor is True
    # El payload del API 0.0.73 entra tal cual: la bandera viaja en la hoja
    # 'otros movimientos' del general y al nivel raíz de cada avión; el
    # consolidado FLOTA (armado a mano en el API) NO la lleva — aquí nadie la
    # lee de ahí (`_sin_provision_vendedor(cons.otros_movimientos)`).
    req = BalanceGeneralRequest.model_validate(
        {
            "consolidado": {
                "otros_movimientos": {"filas": [], "sin_provision_vendedor": True},
            },
            "aviones": [{"matricula": "XA-TST", "sin_provision_vendedor": True}],
        }
    )
    assert req.consolidado.otros_movimientos.sin_provision_vendedor is True
    assert req.consolidado.sin_provision_vendedor is None
    assert req.aviones[0].sin_provision_vendedor is True


def test_textos_sin_provision_son_los_del_contrato() -> None:
    assert _OM_PAGO_VENDEDOR_SIN_PROVISION == OM_FRAG_SIN_PROVISION
    assert _RESUMEN_PAGO_VENDEDOR_SIN_PROVISION == RESUMEN_FRAG_SIN_PROVISION
    assert _MAESTRA_PAGO_VENDEDOR_SIN_PROVISION == MAESTRA_FRAG_SIN_PROVISION
    assert _OM_REGLA_COMISIONES_CUBRE_AVION_SIN_PROVISION == OM_REGLA_CUBRE_SIN_PROVISION
    assert _EMPRESA_REGLA_COMISIONES_CUBRE_AVION_SIN_PROVISION == EMPRESA_REGLA_CUBRE_SIN_PROVISION
    assert _RESUMEN_REGLA_CUBRE_AVION_SIN_PROVISION == RESUMEN_REGLA_CUBRE_SIN_PROVISION
    assert _NOTA_EMPRESA_BASE_SIN_PROVISION == EMPRESA_BASE_SIN_PROVISION
    # Ninguno promete una PROVISIÓN apareada en 'otros movimientos'.
    for t in (OM_FRAG_SIN_PROVISION, RESUMEN_FRAG_SIN_PROVISION, MAESTRA_FRAG_SIN_PROVISION):
        assert "PROVISIÓN" not in t
    assert "provisionado" not in EMPRESA_BASE_SIN_PROVISION
    assert "o provisionado" in _NOTA_EMPRESA_BASE  # la de hoy, intacta


# --- Fila 2 de 'otros movimientos' y nota del RESUMEN (sin vuelos de la regla)


@pytest.mark.parametrize("hay_pago_real", [..., True, False], ids=["sin", "true", "false"])
def test_general_con_bandera_fila2_y_resumen_nuevos(hay_pago_real) -> None:
    """Con `sin_provision_vendedor` UN solo texto para los dos casos (con o sin
    pago real en el periodo): ya no hay PROVISIÓN que explicar."""
    om, resumen, maestra = _leyendas_general(
        render_balance_general_xlsx(
            _general(hay_pago_real, filas=list(_FILAS_V73), sin_provision=True)
        )
    )
    assert om == _cambia(OM_FILA2_VIEJA, OM_FRAG_VIEJO, OM_FRAG_SIN_PROVISION)
    assert resumen == _cambia(RESUMEN_NOTA_VIEJA, RESUMEN_FRAG_VIEJO, RESUMEN_FRAG_SIN_PROVISION)
    assert maestra == MAESTRA_GENERAL_NOTA
    assert "PROVISIÓN" not in om
    assert "los gastos que todavía no se hacen no aparecen" in om


@pytest.mark.parametrize("sin_provision", [..., None, False], ids=["ausente", "null", "false"])
@pytest.mark.parametrize("hay_pago_real", [..., True], ids=["sin pago real", "con pago real"])
def test_general_sin_bandera_fila2_y_resumen_de_hoy(sin_provision, hay_pago_real) -> None:
    """Bandera apagada (API previo que sigue provisionando): los textos de hoy
    byte a byte, elegidos por `hay_pago_vendedor_real` como siempre."""
    om, resumen, maestra = _leyendas_general(
        render_balance_general_xlsx(_general(hay_pago_real, sin_provision=sin_provision))
    )
    if hay_pago_real is True:
        assert om == _cambia(OM_FILA2_VIEJA, OM_FRAG_VIEJO, OM_FRAG_NUEVO)
        assert resumen == _cambia(RESUMEN_NOTA_VIEJA, RESUMEN_FRAG_VIEJO, RESUMEN_FRAG_NUEVO)
    else:
        assert om == OM_FILA2_VIEJA
        assert resumen == RESUMEN_NOTA_VIEJA
    assert maestra == MAESTRA_GENERAL_NOTA
    assert OM_FRAG_SIN_PROVISION not in om


# --- Con vuelos de la regla (API 0.0.66+): «cubierto por el avión» ------------


@pytest.mark.parametrize("variante", ["mensual", "general"])
def test_general_con_bandera_y_regla_render_sin_provision(variante) -> None:
    """Payload del API 0.0.73 (sin marca, con la bandera): fila 2, RESUMEN y
    bloque VUELATOUR con los textos nuevos; ninguna celda amarilla ni la
    leyenda del amarillo; la fila sin gasto real SIN egreso y remanente =
    ingreso; las filas con gasto real igual que hoy; TOTALES = Σ filas."""
    p = _payload_regla_v73(variante)
    data = render_balance_general_xlsx(BalanceGeneralRequest.model_validate(p))
    assert verificar_libro(data)
    wb, wv = load_workbook(BytesIO(data)), _libro(data)
    om, omv = wb["otros movimientos"], wv["otros movimientos"]

    # Fila 2: fragmento nuevo + regla «cubierto por el avión» nueva.
    assert om["A2"].value == _cambia(
        OM_FILA2_VIEJA,
        OM_FRAG_VIEJO,
        OM_FRAG_SIN_PROVISION + OM_REGLA_CUBRE_SIN_PROVISION.format(regla=REGLA),
    )
    assert _OM_REGLA_COMISIONES_CUBRE_AVION.format(regla=REGLA) not in om["A2"].value
    # RESUMEN y bloque VUELATOUR (empresa).
    resumen = _celda_que_empieza(wb["RESUMEN flota"], RESUMEN_PREFIJO)
    assert resumen == (
        _cambia(RESUMEN_NOTA_VIEJA, RESUMEN_FRAG_VIEJO, RESUMEN_FRAG_SIN_PROVISION)
        + RESUMEN_REGLA_CUBRE_SIN_PROVISION.format(regla=REGLA)
    )
    assert _RESUMEN_REGLA_CUBRE_AVION.format(regla=REGLA) not in resumen
    pie = _textos(wb["balance"])
    base = EMPRESA_BASE_SIN_PROVISION + EMPRESA_REGLA_CUBRE_SIN_PROVISION.format(regla=REGLA)
    assert base in pie  # una celda: la base nueva + la regla nueva
    vieja = _NOTA_EMPRESA_BASE + _EMPRESA_REGLA_COMISIONES_CUBRE_AVION.format(regla=REGLA)
    assert not any(vieja in x for x in pie)
    # La base de hoy («lo pagado o provisionado») no se pinta en ninguna celda.
    assert not any(_NOTA_EMPRESA_BASE in x or "o provisionado" in x for x in pie)

    # Sin marca ⇒ sin amarillo, sin leyenda y sin «PROVISIÓN (» en la pestaña.
    assert _rellenos_amber(om) == []
    assert _LEYENDA_PROVISION not in _textos(om)
    assert not any("PROVISIÓN (" in t for t in _textos(om))

    # (a) vigencia, COMPLETADO, sin gasto: SIN concepto/egreso/fecha, ingreso
    # igual, remanente = ingreso (fórmula viva: el egreso vacío cuenta 0).
    assert om["A4"].value == "#412 · Cliente Doce"
    assert (om["C4"].value, om["D4"].value, om["E4"].value) == (None, None, None)
    assert om["G4"].value == 12148.35
    assert om["I4"].value == "=ROUND(G4-D4,2)" and omv["I4"].value == 12148.35
    # (b) CONFIRMADO sin gasto: igual, sin egreso.
    assert (om["C5"].value, om["D5"].value, om["E5"].value) == (None, None, None)
    assert omv["I5"].value == 2030.0
    # (c) con gasto real, idéntico a hoy: cubierto con parcial ($0.00, sin
    # amarillo), cubierto con exceso ($320) y gasto real pre-vigencia (faltan).
    assert om["C6"].value == _FILA_V73_CUBIERTO_PARCIAL["concepto_egreso"]
    assert (om["D6"].value, om["E6"].value, omv["I6"].value) == (0.0, "09/10/2026", 2360.0)
    assert om["C7"].value == _FILA_V73_CUBIERTO_EXCEDE["concepto_egreso"]
    assert (om["D7"].value, omv["I7"].value) == (320.0, 860.0)
    assert om["C8"].value == CONCEPTO_PARCIAL
    assert (om["D8"].value, omv["I8"].value) == (1200.0, 400.0)
    assert om["D8"].comment.text == NOTA_PAGO_REAL
    # (d) TOTALES = Σ de las filas (vuelos + sueltas), remanente = Σ ing − Σ egr.
    todas = p["consolidado"]["otros_movimientos"]["filas"] + p["consolidado"][
        "otros_movimientos"
    ].get("filas_sueltas", [])
    tot = next(c.row for c in om["A"] if c.value == "TOTALES")
    assert omv[f"D{tot}"].value == _suma(todas, "egreso_mxn")
    assert omv[f"G{tot}"].value == _suma(todas, "ingreso_mxn")
    assert omv[f"I{tot}"].value == round(
        _suma(todas, "ingreso_mxn") - _suma(todas, "egreso_mxn"), 2
    )


@pytest.mark.parametrize("sin_provision", [..., None, False], ids=["ausente", "null", "false"])
def test_general_con_regla_sin_bandera_textos_cubre_avion_de_hoy(sin_provision) -> None:
    """API previo (sigue mandando la marca, sin la bandera): las leyendas
    «cubierto por el avión» de hoy y el amarillo de siempre."""
    p = _payload_regla_v73(
        sin_provision=sin_provision, filas=[_tg()._FILA_CUBIERTA, _FILA_V73_CUBIERTO_EXCEDE]
    )
    wb = load_workbook(
        BytesIO(render_balance_general_xlsx(BalanceGeneralRequest.model_validate(p)))
    )
    om = wb["otros movimientos"]
    assert _OM_REGLA_COMISIONES_CUBRE_AVION.format(regla=REGLA) in om["A2"].value
    assert OM_REGLA_CUBRE_SIN_PROVISION.format(regla=REGLA) not in om["A2"].value
    assert OM_FRAG_SIN_PROVISION not in om["A2"].value
    resumen = _celda_que_empieza(wb["RESUMEN flota"], RESUMEN_PREFIJO)
    assert resumen.endswith(RESUMEN_FRAG_VIEJO + _RESUMEN_REGLA_CUBRE_AVION.format(regla=REGLA))
    pie = _textos(wb["balance"])
    vieja = _NOTA_EMPRESA_BASE + _EMPRESA_REGLA_COMISIONES_CUBRE_AVION.format(regla=REGLA)
    assert vieja in pie
    assert not any(EMPRESA_REGLA_CUBRE_SIN_PROVISION.format(regla=REGLA) in x for x in pie)
    assert not any(EMPRESA_BASE_SIN_PROVISION in x for x in pie)
    # La marca sigue pintando (fila 4 = `_FILA_CUBIERTA`), con su leyenda.
    assert _rellenos_amber(om) == ["C4", "D4"]
    assert _LEYENDA_PROVISION in _textos(om)


# --- Base del pie del bloque VUELATOUR (empresa) ------------------------------


def _tev():
    """`tests.test_balance_empresa_vuelatour` (general del API 0.0.59 con el
    bloque VUELATOUR y SIN vuelos de la regla), importado perezoso por la
    misma razón que `_tg`."""
    from tests import test_balance_empresa_vuelatour as tev

    return tev


def _payload_empresa(caso: str, sin_provision) -> dict:
    """Un general con el bloque VUELATOUR en los tres sabores del pie:
    «cubre avión» (API 0.0.66), regla sin «cubre avión» (API 0.0.65) y sin
    vuelos de la regla (API 0.0.59). `sin_provision` = la bandera de la hoja
    (`...` ⇒ no viaja)."""
    if caso == "cubre_avion":
        p = _tg()._payload_om_0066()
    elif caso == "regla_sin_cubre":
        p = _tg()._payload(api_0066=False)
    else:
        p = _tev()._payload()
    om = p["consolidado"].setdefault("otros_movimientos", {"filas": [_FILA_V73_VIGENCIA_SIN_GASTO]})
    om.pop("sin_provision_vendedor", None)
    if sin_provision is not ...:
        om["sin_provision_vendedor"] = sin_provision
    return p


_REGLA_EMPRESA_POR_CASO = {
    "cubre_avion": (
        _EMPRESA_REGLA_COMISIONES_CUBRE_AVION_SIN_PROVISION,
        _EMPRESA_REGLA_COMISIONES_CUBRE_AVION,
    ),
    "regla_sin_cubre": (_EMPRESA_REGLA_COMISIONES, _EMPRESA_REGLA_COMISIONES),
    "sin_regla": ("", ""),
}


def _pie_balance(p: dict) -> list[str]:
    wb = load_workbook(
        BytesIO(render_balance_general_xlsx(BalanceGeneralRequest.model_validate(p)))
    )
    return _textos(wb["balance"])


@pytest.mark.parametrize("caso", ["cubre_avion", "regla_sin_cubre", "sin_regla"])
def test_bloque_empresa_con_bandera_base_sin_provisionado(caso) -> None:
    """Con la bandera, la base del pie del bloque VUELATOUR ya no dice «lo
    pagado o provisionado» (los egresos propios no traen provisión alguna),
    con «cubre avión», con regla sin «cubre avión» y sin regla."""
    con_bandera, _ = _REGLA_EMPRESA_POR_CASO[caso]
    pie = _pie_balance(_payload_empresa(caso, True))
    assert EMPRESA_BASE_SIN_PROVISION + con_bandera.format(regla=REGLA) in pie
    assert not any(_NOTA_EMPRESA_BASE in x or "o provisionado" in x for x in pie)


@pytest.mark.parametrize("sin_provision", [..., None, False], ids=["ausente", "null", "false"])
@pytest.mark.parametrize("caso", ["cubre_avion", "regla_sin_cubre", "sin_regla"])
def test_bloque_empresa_sin_bandera_base_de_hoy(caso, sin_provision) -> None:
    """Sin la bandera (API previo, que sigue provisionando) la base de hoy
    byte a byte, con la regla de hoy."""
    _, sin_bandera = _REGLA_EMPRESA_POR_CASO[caso]
    pie = _pie_balance(_payload_empresa(caso, sin_provision))
    assert _NOTA_EMPRESA_BASE + sin_bandera.format(regla=REGLA) in pie
    assert not any(EMPRESA_BASE_SIN_PROVISION in x for x in pie)


def test_bloque_empresa_sin_nota_del_api_pie_fijo_con_la_base_sin_provision() -> None:
    """Payload sin `nota` (pie fijo de respaldo): la base nueva sustituye a la
    de hoy dentro de `_NOTA_BLOQUE_EMPRESA`."""
    p = _payload_empresa("sin_regla", True)
    p["empresa"]["nota"] = None
    pie = _pie_balance(p)
    assert _NOTA_BLOQUE_EMPRESA.replace(_NOTA_EMPRESA_BASE, EMPRESA_BASE_SIN_PROVISION) in pie
    assert _NOTA_BLOQUE_EMPRESA not in pie
    assert not any("o provisionado" in x for x in pie)


def test_el_amarillo_no_depende_de_la_bandera() -> None:
    """`_MARCA_PROVISION` no se toca: una fila que TRAIGA la marca se pinta
    aunque viaje la bandera (tolerancia; el API 0.0.73 ya no la manda)."""
    p = _payload_regla_v73(filas=[_tg()._FILA_CUBIERTA, _FILA_V73_VIGENCIA_SIN_GASTO])
    om = load_workbook(
        BytesIO(render_balance_general_xlsx(BalanceGeneralRequest.model_validate(p)))
    )["otros movimientos"]
    assert OM_FRAG_SIN_PROVISION in om["A2"].value
    assert _rellenos_amber(om) == ["C4", "D4"]
    assert _LEYENDA_PROVISION in _textos(om)


# --- Libro INDIVIDUAL: la bandera RAÍZ ----------------------------------------


def _individual(**extra) -> BalanceAvionRequest:
    return BalanceAvionRequest(
        matricula="N4142R", periodo_desde="2026-10-01", periodo_hasta="2026-10-31", **extra
    )


def _maestra_individual(req: BalanceAvionRequest) -> str:
    return _celda_que_empieza(
        _libro(render_balance_avion_xlsx(req))["reporte horas N4142R"], MAESTRA_PREFIJO
    )


def test_individual_con_bandera_raiz_cambia_solo_el_fragmento() -> None:
    maestra = _maestra_individual(_individual(sin_provision_vendedor=True))
    assert maestra == _cambia(MAESTRA_NOTA_VIEJA, MAESTRA_FRAG_VIEJO, MAESTRA_FRAG_SIN_PROVISION)
    assert "PROVISIÓN a la fecha del vuelo" not in maestra


@pytest.mark.parametrize("bandera", [..., None, False], ids=["ausente", "null", "false"])
def test_individual_sin_bandera_raiz_nota_de_hoy(bandera) -> None:
    extra = {} if bandera is ... else {"sin_provision_vendedor": bandera}
    assert _maestra_individual(_individual(**extra)) == MAESTRA_NOTA_VIEJA


def test_individual_ignora_la_bandera_de_la_hoja_otros_movimientos() -> None:
    """El libro individual no pinta 'otros movimientos': solo la bandera RAÍZ
    de su payload cambia su nota (la de la hoja es del general)."""
    req = _individual(
        otros_movimientos={"filas": [_FILA_V73_VIGENCIA_SIN_GASTO], "sin_provision_vendedor": True}
    )
    wb = _libro(render_balance_avion_xlsx(req))
    assert "otros movimientos" not in wb.sheetnames
    assert _celda_que_empieza(wb["reporte horas N4142R"], MAESTRA_PREFIJO) == MAESTRA_NOTA_VIEJA


# --- Rutas: el payload del API 0.0.73 entra sin 422 y activa los textos ------


def test_rutas_aceptan_el_payload_0073(monkeypatch) -> None:
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    headers = {"X-Internal-Token": TOKEN}
    p = _payload_regla_v73()

    res = client.post("/pdf/balance-general-xlsx", json=p, headers=headers)
    assert res.status_code == 200, res.text
    om, resumen, _ = _leyendas_general(res.content)
    assert OM_FRAG_SIN_PROVISION in om and OM_REGLA_CUBRE_SIN_PROVISION.format(regla=REGLA) in om
    assert RESUMEN_FRAG_SIN_PROVISION in resumen
    assert _rellenos_amber(load_workbook(BytesIO(res.content))["otros movimientos"]) == []

    # La vista en pantalla es el MISMO libro (libro JSON, 12-oct-2026).
    res = client.post("/pdf/balance-general-libro", json=p, headers=headers)
    assert res.status_code == 200, res.text
    assert OM_FRAG_SIN_PROVISION in json.dumps(res.json(), ensure_ascii=False)

    res = client.post(
        "/pdf/balance-avion-xlsx",
        json=_individual(sin_provision_vendedor=True).model_dump(mode="json"),
        headers=headers,
    )
    assert res.status_code == 200, res.text
    maestra = _celda_que_empieza(_libro(res.content)["reporte horas N4142R"], MAESTRA_PREFIJO)
    assert maestra == _cambia(MAESTRA_NOTA_VIEJA, MAESTRA_FRAG_VIEJO, MAESTRA_FRAG_SIN_PROVISION)
