"""Payloads del cierre de mes (16-oct-2026, API 0.0.88, invariante 89 del
API) para `tests/test_cierre_mes.py`.

Módulo aparte A PROPÓSITO: solo importa lo que ya existía en el HEAD
`f0c8ad1` (los constructores de forma real de los tests de otros ingresos,
del reparto y del Libro Dinero), para que las huellas «de siempre» se
calculen corriendo ESTOS mismos constructores con el código de ese HEAD
(`_HUELLAS_HEAD` del test). Nada de aquí calcula dinero: los montos del
cierre son los del contrato (QA-1, #316 Mike Nelson) y los de las filas son
los del payload de forma real.
"""

from __future__ import annotations

import copy

from tests import test_balance_otros_ingresos_avion as t_oi
from tests import test_dinero_otros_ingresos_avion as t_din
from tests import test_reparto_comisiones as trc

# ---------------------------------------------------------------------------
# Payloads «de siempre» (API ≤ 0.0.87): los de forma real del lote 0.0.77.
# ---------------------------------------------------------------------------


def individual_base() -> dict:
    return t_oi._individual_forma_real(api_0077=True)


def general_base(variante: str) -> dict:
    return t_oi._general_forma_real(variante, api_0077=True)


def reparto_base() -> dict:
    """El reparto de `test_reparto_comisiones` con el detalle de vuelos de
    XB-TST (la tabla «Vuelo / Cliente» del PDF y del Excel)."""
    p = trc._payload()
    p["aviones"][0]["vuelos"] = [
        {"folio": 316, "cliente": "Mike Nelson", "fecha": "2026-09-25",
         "estado": "COMPLETADO", "cobrado_usd": 1200.0, "pendiente_usd": 0.0},
        {"folio": 318, "cliente": "Ambulancia Internacional", "fecha": "2026-09-27",
         "estado": "COMPLETADO", "cobrado_usd": 800.0, "pendiente_usd": 0.0},
    ]  # fmt: skip
    return p


def dinero_base() -> dict:
    return copy.deepcopy(t_din._BASE)


# ---------------------------------------------------------------------------
# Las llaves NUEVAS en su valor neutro (lo que mandaría un API 0.0.88 sobre
# un mes sin cierre ni vuelos diferidos, si las mandara explícitas).
# ---------------------------------------------------------------------------

_FILA_NEUTRA = {"cierre_marca": None, "diferido_a": None, "diferido_de": None}


def con_llaves_neutras(p: dict, libro: str) -> dict:
    p = copy.deepcopy(p)
    if libro == "reparto":
        p["cierre"] = None
        for a in p["aviones"]:
            for v in a["vuelos"]:
                v.update(_FILA_NEUTRA)
        return p
    if libro == "dinero":
        p["cierre"] = None
        for v in p["vuelos"]:
            v.update(_FILA_NEUTRA)
        return p
    p["cierre"] = None
    libros = [p] if libro == "individual" else [p["consolidado"], *p.get("aviones", [])]
    if libro != "individual":
        p["resumen_cierre"] = None
    for lb in libros:
        lb["cierre"] = None
        for v in lb.get("vuelos", []):
            v.update(_FILA_NEUTRA, fila_sombra=False)
    return p


def cierre_abierto_vacio() -> dict:
    """El bloque `cierre` de un mes ABIERTO sin vuelos que pasarían."""
    return {"disponible": True, "periodo": "2026-09", "estado": "ABIERTO",
            "cerrado_at": None, "cerrado_por_nombre": None,
            "pasan": [], "vienen": [], "pasarian": []}  # fmt: skip


def reparto_cierre_abierto_vacio() -> dict:
    return {"disponible": True, "periodo": "2026-09", "estado": "ABIERTO",
            "cerrado_at": None, "cerrado_por_nombre": None, "version": None,
            "coincide_con_snapshot": None, "diferidos_salen": [],
            "diferidos_entran": [], "pendientes_que_pasarian": []}  # fmt: skip


# ---------------------------------------------------------------------------
# Balance CON cierre (septiembre 2026 cerrado el 16-oct por Ale): la forma
# del contrato §4.3. XA-TST: #499 PASA a octubre (fila SOMBRA: horas reales,
# dinero en null) y #403 VIENE de agosto (dinero completo, horas en null).
# Las horas del mes no cambian (0.95 h salen con #403 y 0.95 h entran con la
# sombra), así que los totales del API del payload siguen siendo los suyos.
# ---------------------------------------------------------------------------

MARCA_PASA = "Pasa a octubre 2026 (pendiente de cobro: cobrado $0.00 de $36,632.89 MXN)"
MARCA_VIENE = "Viene de agosto 2026"
LEYENDA_PASA = "Pasa a octubre 2026 · pendiente de cobro: cobrado $0.00 de $2,076.40 USD"

# La fila SOMBRA y la que VIENE, campo por campo como las arma el API
# (`aircraft/cierre-balance.util.ts` de la rama 0.0.88: `aplicarFilaSombra` y
# `aplicarFilaViene`). Si el API cambia esa lista, se cambia aquí: el test
# prueba el libro con la forma REAL de la fila, no con una aproximación.
# Llaves que el API pone en null en la SOMBRA (todo el dinero, los T.C. y
# las horas COBRADAS, que son de la venta).
DINERO_FILA = (
    "horas_cobradas", "tarifa_usd", "iva_hr_usd", "total_usd", "iva_usd", "tc_venta",
    "total_cotizacion_usd", "total_cotizacion_mxn", "otros_ingresos_usd", "tua_pagado_mxn",
    "total_mxn", "iva_mxn", "subtotal_mxn", "gas_mxn", "gas_litros", "gas_precio_litro",
    "op_mxn", "piloto_mxn", "otros_mxn", "permiso_afac_mxn", "permiso_afac_usd_hr",
    "costo_total_mxn", "tc_costos", "costo_usd", "costo_usd_siva", "iva_pagado_usd",
    "iva_pagado_mxn", "remanente_mxn", "dif_iva_mxn", "comision_vendedor_mxn",
    "ganancia_mxn", "ganancia_usd", "costo_hr_usd", "costo_hr_usd_siva", "cobrado_mxn",
    "cobrado_real_mxn", "por_cobrar_mxn", "por_cobrar_usd", "factura_vuelatour",
)  # fmt: skip
# Listas que el API vacía en la SOMBRA.
LISTAS_SOMBRA = (
    "gas_detalle", "op_detalle", "piloto_detalle", "otros_detalle", "op_vinculos",
    "piloto_vinculos", "otros_vinculos", "cobros",
)  # fmt: skip
# Llaves que el API BORRA de la SOMBRA (`delete fila.…`).
BORRADAS_SOMBRA = (
    "extension_pagada_mxn", "comisiones_mxn", "comisiones_detalle",
    "comision_banco_avion_mxn", "comision_vendedor_prov_mxn",
    "comision_vendedor_real_mxn", "comision_vendedor_real_pagos",
)  # fmt: skip


def aplicar_fila_sombra(fila: dict, destino: str, leyenda: str) -> dict:
    """`aplicarFilaSombra` del API: horas, tacos y ruta REALES; el dinero en
    null; ESTADO real; STATUS «Pasa a <mes corto>»."""
    s = copy.deepcopy(fila)
    for k in DINERO_FILA:
        s[k] = None
    for k in LISTAS_SOMBRA:
        s[k] = []
    for k in BORRADAS_SOMBRA:
        s.pop(k, None)
    s.update(
        status_cobro=f"Pasa a {_MES_CORTO[int(destino[5:7]) - 1]}-{destino[:4]}",
        cierre_marca=leyenda,
        fila_sombra=True,
        diferido_a=destino,
        diferido_de=None,
    )
    return s


def aplicar_fila_viene(fila: dict, origen: str) -> None:
    """`aplicarFilaViene` del API: el dinero completo; horas y tacos en null."""
    fila.update(
        tiempo_vuelo=None,
        taco_inicio=None,
        taco_fin=None,
        taco_inicio_obs=[],
        taco_fin_obs=[],
        salto_taco_inicio=False,
        salto_taco_esperado=None,
        salto_taco_interno=False,
        salto_taco_interno_detalle=None,
        cierre_marca=f"Viene de {cierre_mes_etiqueta(origen)}",
        fila_sombra=False,
        diferido_a=None,
        diferido_de=origen,
    )


_MES_CORTO = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
_MES_LARGO = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
    "septiembre", "octubre", "noviembre", "diciembre",
]  # fmt: skip


def cierre_mes_etiqueta(mes: str) -> str:
    return f"{_MES_LARGO[int(mes[5:7]) - 1]} {mes[:4]}"


def _fila_sombra(modelo: dict) -> dict:
    """#499 (voló el 28-sep, PASA a octubre): la fila de forma real del
    modelo con sus horas propias (0.95 h, tacos que empalman)."""
    s = copy.deepcopy(modelo)
    s.update(
        clave="#499 · Cliente Nueve",
        fecha="2026-09-28",
        orden_ts="2026-09-28T12:00:00Z",
        tiempo_vuelo=0.95,
        taco_inicio=1100.0,
        taco_fin=1100.95,
        estado="COMPLETADO",
    )
    return aplicar_fila_sombra(s, "2026-10", MARCA_PASA)


def con_filas_de_cierre(p: dict) -> dict:
    """Las dos filas del cierre en el libro de XA-TST (o del consolidado)."""
    p = copy.deepcopy(p)
    viene = next(v for v in p["vuelos"] if v["clave"].startswith("#403"))
    aplicar_fila_viene(viene, "2026-08")
    assert viene["cierre_marca"] == MARCA_VIENE
    p["vuelos"].append(_fila_sombra(p["vuelos"][0]))
    return p


def cierre_cerrado(*, con_avion: bool = False) -> dict:
    pasa = {"clave": "#499 · Cliente Nueve", "folio": 499, "cliente": "Cliente Nueve",
            "matricula": "XA-TST", "fecha": "2026-09-28", "venta_mxn": 36632.89,
            "venta_usd": 2076.4, "cobrado_mxn": 0, "cobrado_usd": 0,
            "por_cobrar_mxn": 36632.89, "por_cobrar_usd": 2076.4, "destino": "2026-10",
            "leyenda": LEYENDA_PASA}  # fmt: skip
    viene = {"clave": "#403 · Traslado", "folio": 403, "cliente": "Traslado",
             "matricula": "XA-TST", "fecha": "2026-08-30", "venta_mxn": 1000.0,
             "venta_usd": None, "cobrado_mxn": 1000.0, "por_cobrar_mxn": 0,
             "origen": "2026-08", "leyenda": MARCA_VIENE}  # fmt: skip
    if not con_avion:
        pasa.pop("matricula")
        viene.pop("matricula")
    return {"disponible": True, "periodo": "2026-09", "estado": "CERRADO",
            "cerrado_at": "2026-10-16T19:02:11.123Z", "cerrado_por_nombre": "Ale",
            "pasan": [pasa], "vienen": [viene], "pasarian": []}  # fmt: skip


def individual_con_cierre() -> dict:
    p = con_filas_de_cierre(individual_base())
    p["cierre"] = cierre_cerrado()
    return p


def general_con_cierre(variante: str) -> dict:
    """El general de forma real con el libro de XA-TST CON sus filas del
    cierre (el consolidado se arma con el constructor de siempre, que junta
    las filas de los libros) + `cierre` y `resumen_cierre` en la raíz."""
    original = t_oi._individual_forma_real
    try:
        t_oi._individual_forma_real = lambda **kw: con_filas_de_cierre(original(**kw))
        p = general_base(variante)
    finally:
        t_oi._individual_forma_real = original
    p["cierre"] = cierre_cerrado(con_avion=True)
    p["resumen_cierre"] = {
        "pasan": [{"matricula": "XA-TST", "vuelos": 1, "venta_mxn": 36632.89,
                   "por_cobrar_mxn": 36632.89}],
        "vienen": [{"matricula": "XA-TST", "vuelos": 1, "venta_mxn": 1000.0,
                    "por_cobrar_mxn": 0}],
    }  # fmt: skip
    return p


# ---------------------------------------------------------------------------
# Reparto CON cierre (contrato §4.1): #316 de XB-TST SALE a octubre (no está
# en `vuelos`), #318 VIENE de agosto; en XB-DOS un vuelo pasaría (mes
# abierto en la variante) y uno EXTERNO sin avión sale.
# ---------------------------------------------------------------------------


def reparto_con_cierre(*, coincide: bool | None = True, estado: str = "CERRADO") -> dict:
    p = reparto_base()
    p["aviones"][0]["vuelos"] = [
        {"folio": 318, "cliente": "Ambulancia Internacional", "fecha": "2026-08-27",
         "estado": "COMPLETADO", "cobrado_usd": 800.0, "pendiente_usd": 0.0,
         "cierre_marca": "Viene de agosto 2026", "diferido_de": "2026-08"},
    ]  # fmt: skip
    p["cierre"] = {
        "disponible": True, "periodo": "2026-09", "estado": estado,
        "cerrado_at": "2026-10-16T19:02:11.123Z", "cerrado_por_nombre": "Ale",
        "version": 1, "coincide_con_snapshot": coincide,
        "diferidos_salen": [
            {"id": "v-316", "folio": 316, "cliente": "Mike & Nelson <MN>",
             "matricula": "XB-TST", "venta_usd": 2076.4, "cobrado_usd": 0,
             "por_cobrar_usd": 2076.4, "destino": "2026-10",
             "leyenda": "Pasa a octubre 2026 · pendiente de cobro: cobrado $0.00 de $2,076.40 USD"},
            {"id": "v-990", "folio": 990, "cliente": "Operador Externo",
             "matricula": None, "venta_usd": 500.0, "cobrado_usd": 100.0,
             "por_cobrar_usd": 400.0, "destino": "2026-10", "leyenda": None},
        ],
        "diferidos_entran": [
            {"id": "v-318", "folio": 318, "cliente": "Ambulancia Internacional",
             "matricula": "XB-TST", "venta_usd": 800.0, "cobrado_usd": 800.0,
             "por_cobrar_usd": 0, "origen": "2026-08", "leyenda": "Viene de agosto 2026"},
        ],
        "pendientes_que_pasarian": [],
    }  # fmt: skip
    return p


# ---------------------------------------------------------------------------
# Libro Dinero CON cierre (contrato §4.4): el vuelo que VIENE trae su fila
# con la marca; el que PASA no está en `vuelos`, solo en `cierre.pasan`.
# ---------------------------------------------------------------------------


def dinero_con_cierre() -> dict:
    p = dinero_base()
    p["vuelos"][0].update(cierre_marca="Viene de septiembre 2026", diferido_de="2026-09")
    p["cierre"] = cierre_cerrado()
    p["cierre"]["periodo"] = "2026-10"
    return p
