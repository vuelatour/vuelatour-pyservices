"""La RESERVA DE OVERHAUL RESTA EN LA CASCADA del reparto a socios (9-nov-2026,
API 0.0.129, invariante 125 del API): helpers del PDF y del Excel (fuente
única de los dos documentos, como `reparto_combustible` y
`reparto_comisiones`).

Decisión del cliente (9-nov-2026): «Necesito resolver ese tema del Overhaul,
que quede ANTES del total». Su Excel manual y el Diseño Funcional §4.8 tienen
la reserva como UNA RESTA MÁS de la cascada, antes del saldo:

    Venta cobrada − Comisiones − Gastos directos − Provisión AFAC
      = Utilidad antes de gastos
    − Combustible − Gastos indirectos − Permisos − Fijos prorrateados
    − Reserva de overhaul + Otros ingresos del avión
      = Saldo a repartir                                     (`saldo_*`)

(El API 0.0.128 del 7-nov la restaba DESPUÉS del saldo —«Saldo antes de la
reserva» − reserva = «Saldo a repartir», llaves `saldo_a_repartir_*`— y el
cliente lo rechazó: esas llaves ya no viajan, el esquema las conserva
deprecadas y aquí se IGNORAN.)

Discriminador ÚNICO, aditivo y booleano: `reserva_overhaul_en_cascada: true`
(`reserva_en_cascada`; solo un `true` JSON cuenta). Con él, `saldo_usd` y
`saldo_mxn` YA RESTAN `reserva_overhaul_usd/_mxn`; `reparto[].utilidad_*` =
su parte del saldo ANTES de la reserva, `reparto[].reserva_*` = su parte de
la reserva, `monto_*` = utilidad − reserva (Σ monto = `saldo_*` al centavo),
`descuenta_reserva` true en todos y `reserva_retenida_*` = Σ reserva de los
socios (= la del avión). Aquí no se calcula nada: se pinta lo que manda el
API. Solo aplica con la cascada del balance (`reparto_como_balance`; sin
ella el `saldo_*` del API siempre ha restado la reserva y los documentos ya
la pintan como una resta); sin la llave (API ≤ 0.0.127) los documentos de
siempre, byte a byte (textos por datos de la invariante 121 incluidos).
"""

from app.schemas.reparto import RepartoAvion, RepartoPdfRequest


def reserva_en_cascada(avion: RepartoAvion) -> bool:
    """¿El API (0.0.129+) restó la reserva de overhaul EN la cascada de este
    avión (`saldo_*` ya neto)? Solo un `true` JSON en
    `reserva_overhaul_en_cascada`; None = API previo ⇒ los documentos de
    siempre."""
    return avion.reserva_overhaul_en_cascada is True


def reserva_en_cascada_req(req: RepartoPdfRequest) -> bool:
    """Lo mismo para los textos de TODO el documento (pie del PDF, resumen por
    socio, rótulos y notas del Excel): algún avión trae la llave (el API
    0.0.129 la manda en todos; un avión sin ella —no debería darse— se pinta
    como antes en su bloque)."""
    return any(reserva_en_cascada(a) for a in req.aviones)
