"""La RESERVA DE OVERHAUL SE RETIENE en el reparto a socios (7-nov-2026, API
0.0.128): helpers del PDF y del Excel (fuente única de los dos documentos,
como `reparto_combustible` y `reparto_comisiones`).

Decisión del cliente (7-nov-2026): «aún no se está reteniendo de las
utilidades para cada socio en cada avión, aún dice informativa no se resta;
debe verse bien y retenerse bien en el balance por avión y general, y verse
también en el reparto de utilidades». Desde el API 0.0.101 el reparto llevaba
la cascada del balance con la reserva INFORMATIVA tras el saldo (y desde el
0.0.125 descontada a todos los socios, pero todavía fuera de la cascada). Con
el 0.0.128 la reserva vuelve a ser una RESTA explícita antes de repartir:

    Saldo antes de la reserva de overhaul   (`saldo_*`, el de siempre)
    (-) Reserva de overhaul (se retiene para el overhaul)   (`reserva_overhaul_*`)
    = Saldo a repartir                       (`saldo_a_repartir_*`)

`saldo_a_repartir_usd` (round2(saldo − reserva); 0.0 también cuenta) es el
DISCRIMINADOR; `reparto[].monto_*` ya es neto (utilidad − reserva; Σ monto =
saldo a repartir al centavo) y `reserva_retenida_*` es la reserva completa.
Aquí no se calcula nada: se pinta lo que manda el API. Solo aplica con la
cascada del balance (`reparto_como_balance`, donde `saldo_*` va ANTES de la
reserva); sin la llave (API ≤ 0.0.127) los documentos de siempre, byte a byte.
"""

from app.schemas.reparto import RepartoAvion, RepartoPdfRequest


def reserva_retenida(avion: RepartoAvion) -> bool:
    """¿El API (0.0.128+) manda el saldo A REPARTIR de este avión (saldo −
    reserva de overhaul)? La presencia de `saldo_a_repartir_usd` decide (0.0
    también cuenta); None = API previo ⇒ la reserva solo se informa."""
    return avion.saldo_a_repartir_usd is not None


def reserva_retenida_req(req: RepartoPdfRequest) -> bool:
    """Lo mismo para los textos de TODO el documento (pie del PDF, resumen por
    socio, layout del Excel): algún avión trae la llave (el API 0.0.128 la
    manda en todos; un avión sin ella —no debería darse— se pinta como antes
    en su bloque)."""
    return any(reserva_retenida(a) for a in req.aviones)


def saldo_a_repartir_usd(avion: RepartoAvion, *, como_balance: bool) -> float:
    """El saldo que de verdad se reparte, en dólares: con la reserva retenida
    (y la cascada del balance) `saldo_a_repartir_usd`; si no, el `saldo_usd`
    de siempre."""
    if como_balance and reserva_retenida(avion):
        return avion.saldo_a_repartir_usd
    return avion.saldo_usd


def saldo_a_repartir_mxn(avion: RepartoAvion, *, como_balance: bool) -> float | None:
    """Gemelo en PESOS (None ⇒ «—»: aquí jamás se resta un peso)."""
    if como_balance and reserva_retenida(avion):
        return avion.saldo_a_repartir_mxn
    return avion.saldo_mxn
