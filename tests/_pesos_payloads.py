"""Payloads del balance y del reparto EN PESOS (18-oct-2026, API 0.0.93) para
`tests/test_balance_en_pesos.py` y `tests/test_reparto_en_pesos.py`.

Módulo aparte A PROPÓSITO (como `_cierre_mes_payloads`): solo importa lo que
ya existía en el HEAD 3b893f1, para que las huellas «sin la bandera» se
calculen corriendo ESTOS mismos constructores con el código de ese HEAD
(`_HUELLAS_HEAD` de los tests). Los `*_mxn` son el ESPEJO de la aritmética del
API (`aircraft-balance.service.ts`, bloque «LA CASCADA EN PESOS», y
`balance-empresa.util.ts`, bloque «EN PESOS»): el render no calcula nada de
esto, solo lo pinta.
"""

from __future__ import annotations

import copy

from tests import _cierre_mes_payloads as pay


def r2(x: float) -> float:
    return round(x + 0.0, 2)


def _mxn_de_hoja(hoja: dict | None) -> float | None:
    """`mxnDeHoja` del API: null si alguna fila no trae pesos; si no, el
    TOTAL MXN de la hoja (sin hoja ⇒ 0: el API la manda siempre)."""
    if hoja is None:
        return 0.0
    if any(f.get("monto_mxn") is None for f in hoja.get("filas") or []):
        return None
    return r2(hoja.get("total_mxn") or 0)


def _mxn_otros_ingresos(hoja: dict | None) -> float | None:
    if hoja is None:
        return 0.0
    if any(f.get("neto_mxn") is None for f in hoja.get("filas") or []):
        return None
    return r2(hoja.get("total_mxn") or 0)


def con_cascada_mxn(libro: dict) -> dict:
    """El libro de un avión con las llaves `*_mxn` del bloque `balance` y
    `monto_mxn` de cada socio, como las manda el API 0.0.93 (SIN la bandera:
    quien llama decide)."""
    libro = copy.deepcopy(libro)
    t, b = libro["totales"], libro["balance"]
    antes = t.get("ganancia_mxn")
    comb = _mxn_de_hoja(libro.get("combustible"))
    ind = _mxn_de_hoja(libro.get("gastos_indirectos"))
    ref = _mxn_de_hoja(libro.get("refacciones"))
    otros = _mxn_de_hoja(libro.get("otros_gastos"))
    perm = _mxn_de_hoja(libro.get("permisos"))
    oi = _mxn_otros_ingresos(libro.get("otros_ingresos_avion"))
    partes = [comb, ind, ref, otros, perm, oi]
    despues = (
        None
        if antes is None or any(x is None for x in partes)
        else r2(antes - comb - ind - ref - otros - perm + oi)
    )
    por_cobrar = t.get("por_cobrar_mxn")
    cobrada = None if despues is None or por_cobrar is None else r2(despues - por_cobrar)
    b.update(
        utilidad_antes_mxn=antes,
        combustible_mxn=comb,
        gastos_indirectos_mxn=ind,
        refacciones_mxn=ref,
        otros_mxn=otros,
        permisos_mxn=perm,
        otros_ingresos_avion_mxn=oi,
        utilidad_despues_mxn=despues,
        por_cobrar_mxn=por_cobrar,
        utilidad_cobrada_mxn=cobrada,
    )
    if "comisiones_usd" in b:
        # Σ COMISIONES MXN que restaron las filas (solo las que tienen
        # GANANCIA MXN), y la utilidad antes de ellas.
        com = r2(
            sum(
                v.get("comisiones_mxn") or 0
                for v in libro.get("vuelos") or []
                if v.get("ganancia_mxn") is not None
            )
        )
        b["comisiones_mxn"] = com
        b["utilidad_antes_comisiones_mxn"] = None if antes is None else r2(antes + com)
    for s in b.get("socios") or []:
        pct = s.get("porcentaje") or 0
        s["monto_mxn"] = None if cobrada is None else r2(pct / 100 * cobrada)
    return libro


def _lado_mxn(filas: list[dict], llave: str) -> float | None:
    if any(f.get("sin_tc") for f in filas):
        return None
    return r2(sum(f.get(llave) or 0 for f in filas))


def con_empresa_mxn(general: dict) -> dict:
    """El bloque `empresa` del general con sus llaves en pesos (espejo de
    `armarBalanceEmpresa` del API 0.0.93)."""
    general = copy.deepcopy(general)
    emp = general.get("empresa")
    if emp is None:
        return general
    socios_empresa = {
        a["matricula"]: s
        for a in general["aviones"]
        for s in a["balance"].get("socios") or []
        if s.get("es_empresa")
    }
    for p in emp.get("participaciones") or []:
        s = socios_empresa.get(p.get("matricula"))
        p["monto_mxn"] = None if s is None else s.get("monto_mxn")
    parts = [p.get("monto_mxn") for p in emp.get("participaciones") or []]
    part = None if any(x is None for x in parts) else r2(sum(parts))
    om = (general.get("consolidado") or {}).get("otros_movimientos") or {}
    filas = [*(om.get("filas") or []), *(om.get("filas_sueltas") or [])]
    ing = _lado_mxn(filas, "ingreso_mxn")
    egr = _lado_mxn(filas, "egreso_mxn")
    otros = _mxn_de_hoja(general.get("gastos_empresa"))
    inventario = general.get("inventario")
    tienda = None if inventario is None else inventario.get("total_utilidad_mxn")
    resultado = (
        None
        if any(x is None for x in (part, ing, egr, otros))
        else r2(part + ing - egr - otros + (tienda or 0))
    )
    emp.update(
        participacion_mxn=part,
        ingresos_propios_mxn=ing,
        pagos_vendedor_mxn=egr,
        otros_gastos_empresa_mxn=otros,
        tienda_utilidad_mxn=tienda,
        resultado_mxn=resultado,
        nota_pesos=(
            "En pesos: participación = utilidad COBRADA en pesos × % de la empresa; "
            "ingresos y egresos = columnas MXN de 'otros movimientos'."
        ),
    )
    return general


def individual(*, bandera: bool | None = True) -> dict:
    """Libro individual de forma real (API 0.0.77+) con la cascada en pesos;
    `bandera` None = sin la llave."""
    p = con_cascada_mxn(pay.individual_base())
    if bandera is not None:
        p["balance_en_pesos"] = bandera
    return p


def general(variante: str = "mensual", *, bandera: bool | None = True) -> dict:
    """Balance general de forma real con los bloques en pesos (bandera en la
    raíz y en cada `aviones[]`, como el API 0.0.93; el consolidado no)."""
    p = pay.general_base(variante)
    p["aviones"] = [con_cascada_mxn(a) for a in p["aviones"]]
    if bandera is not None:
        for a in p["aviones"]:
            a["balance_en_pesos"] = bandera
        p["balance_en_pesos"] = bandera
    return con_empresa_mxn(p)


# Reparto: T.C. de la cotización con que se pasan a pesos los números de
# prueba (los pesos los manda el API; aquí solo se arma un payload coherente).
K_REPARTO = 17.25
_ATTRS_REPARTO = (
    "ingresos_cobrado",
    "comisiones_venta",
    "pendiente_cobro",
    "gastos_directos",
    "gastos_indirectos",
    "permisos",
    "otros",
    "reserva_overhaul",
    "otros_ingresos_avion",
)


def reparto(*, bandera: bool | None = True) -> dict:
    """El reparto de forma real (`_cierre_mes_payloads.reparto_base`) con las
    llaves en pesos del API 0.0.93."""
    p = copy.deepcopy(pay.reparto_base())
    for a in p["aviones"]:
        for k in _ATTRS_REPARTO:
            a[f"{k}_mxn"] = r2((a.get(f"{k}_usd") or 0.0) * K_REPARTO)
        a["saldo_mxn"] = r2(
            a["ingresos_cobrado_mxn"]
            - a["comisiones_venta_mxn"]
            - a["gastos_directos_mxn"]
            - a["gastos_indirectos_mxn"]
            - a["permisos_mxn"]
            - a["otros_mxn"]
            - a["reserva_overhaul_mxn"]
            + a["otros_ingresos_avion_mxn"]
        )
        a["pesos_sin_tc_usd"] = 0
        a["reserva_tc"] = K_REPARTO
        for linea in a.get("reparto") or []:
            linea["monto_mxn"] = r2(a["saldo_mxn"] * (linea.get("porcentaje") or 0) / 100)
        for v in a.get("vuelos") or []:
            v["cobrado_mxn"] = r2((v.get("cobrado_usd") or 0) * K_REPARTO)
            v["pendiente_mxn"] = r2((v.get("pendiente_usd") or 0) * K_REPARTO)
    if bandera is not None:
        p["reparto_en_pesos"] = bandera
    return p
