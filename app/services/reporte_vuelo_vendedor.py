"""Pago al VENDEDOR en el reporte por vuelo (PDF y Excel) — fuente única de
qué se pinta en la sección «Cotización» y en la cascada «Balance del vuelo».

Aquí NO se calcula ningún pago: los dos números llegan del API.

- **API ≤ 0.0.74** (sin la bandera): `pago_vendedor_usd` = el pago COTIZADO
  (comisión + su IVA cuando grava, `pagoVendedorUsd`) y `neto_vuelatour_usd`
  = total − ese pago. Sin `pago_vendedor_usd` (API más viejo), la comisión
  pre-IVA. Las dos secciones pintan ese mismo pago, byte a byte como antes.
- **API 0.0.75+** (bandera `comision_vendedor_solo_gasto_real`, invariante
  73 del API, revisión del lote 2 del 14-oct-2026): `pago_vendedor_usd` =
  el pago REAL al vendedor (los gastos «Comisión del vendedor (VuelaTour)» /
  «(avión)» capturados; 0 si no hay) y es lo que resta la ganancia;
  `pago_vendedor_cotizado_usd` = la comisión COTIZADA con su IVA (dato de la
  cotización). Entonces la sección «Cotización» pinta la COTIZADA y su neto
  (total − cotizada: lo mismo que el panel en «Pago al vendedor (comisión +
  IVA)» / «Neto VuelaTour») y la cascada rotula «(−) Pago real al vendedor»
  con el pago real, sin «(c/IVA)» (un pago real puede no ser comisión + IVA:
  un pago en pesos convertido no lo es). Antes la sección «Cotización»
  decía «Pago comisión vendedor −$0.00» y «Neto VuelaTour» = total cuando no
  había pago capturado, contra el panel.
"""

from __future__ import annotations

from app.schemas.reportes import ReporteVueloRequest

# Margen para decidir «(comisión + IVA)»: el pago supera la comisión pre-IVA.
_EPS = 0.005


def solo_gasto_real(r: ReporteVueloRequest) -> bool:
    """¿El API manda el pago REAL en `pago_vendedor_usd` (0.0.75+)?"""
    return r.comision_vendedor_solo_gasto_real is True


def _quien(r: ReporteVueloRequest) -> str:
    return f" ({r.comision_vendedor_nombre})" if r.comision_vendedor_nombre else ""


def _pago_legado(r: ReporteVueloRequest) -> float:
    """El pago que mandaba el API ≤ 0.0.74 (el cotizado); sin él, la comisión."""
    return (
        r.pago_vendedor_usd
        if r.pago_vendedor_usd is not None
        else r.comision_vendedor_usd
    )


def pago_cotizacion(r: ReporteVueloRequest) -> float:
    """Pago al vendedor de la sección «Cotización»: la COTIZADA con su IVA."""
    if solo_gasto_real(r):
        return (
            r.pago_vendedor_cotizado_usd
            if r.pago_vendedor_cotizado_usd is not None
            else r.comision_vendedor_usd
        )
    return _pago_legado(r)


def neto_cotizacion(r: ReporteVueloRequest, pago: float) -> float:
    """«Neto VuelaTour» de la sección «Cotización» (total − pago cotizado)."""
    if solo_gasto_real(r):
        # `neto_vuelatour_usd` del 0.0.75 resta el pago REAL: aquí se habla de
        # la cotización (como el panel), así que es total − cotizada.
        return round(r.total_usd - pago, 2)
    return (
        r.neto_vuelatour_usd
        if r.neto_vuelatour_usd is not None
        else r.total_usd - pago
    )


def etiqueta_pago_cotizacion(r: ReporteVueloRequest, pago: float) -> str:
    """«Pago comisión vendedor (Juan) (comisión + IVA)» (texto plano)."""
    con_iva = " (comisión + IVA)" if pago > r.comision_vendedor_usd + _EPS else ""
    return f"Pago comisión vendedor{_quien(r)}{con_iva}"


def pago_balance(r: ReporteVueloRequest) -> float:
    """Lo que resta la cascada del balance (= lo que restó la ganancia del API)."""
    if solo_gasto_real(r):
        return r.pago_vendedor_usd or 0.0
    return _pago_legado(r)


def etiqueta_pago_balance(r: ReporteVueloRequest) -> str:
    """«(−) Pago real al vendedor (Juan)» con la bandera; si no, la de siempre."""
    if solo_gasto_real(r):
        return f"(−) Pago real al vendedor{_quien(r)}"
    con_iva = " (c/IVA)" if _pago_legado(r) > r.comision_vendedor_usd + _EPS else ""
    return f"(−) Pago comisión vendedor{_quien(r)}{con_iva}"
