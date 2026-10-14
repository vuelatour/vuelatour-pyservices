"""Comisiones a cargo del avión en el reparto a socios: textos del PDF y del
Excel (fuente única de los dos documentos).

Regla de septiembre 2026 (6-oct-2026, API 0.0.65; respuestas del cliente: la
comisión del vendedor la absorbe el avión, lo cobrado al cliente por ella
sigue siendo ingreso de VuelaTour, desde septiembre 2026 por fecha del
vuelo). El API suma en `comisiones_venta_usd` —el campo de siempre, que la
cascada ya resta antes del saldo— la parte del avión de la comisión bancaria
de sus cobros y lo del vendedor (fuente única `comisionesDelVuelo`, la misma
del balance): hasta el API 0.0.73 su provisión; desde el 0.0.74 (13-oct-2026,
invariante 61 del API, bandera `comision_vendedor_solo_gasto_real`) el pago
capturado como «Comisión del vendedor (avión)» y ligado al vuelo, en
CUALQUIER fecha (el de «Comisión del vendedor (VuelaTour)» lo paga VuelaTour
y nunca llega aquí). Aquí no se calcula nada: cambian la etiqueta de esa
fila/columna y las frases que decían «su pago no es costo del avión», y SOLO
cuando un avión de verdad absorbe comisiones: sin eso, los textos de
siempre. Con la bandera, las frases nombran las DOS categorías
(`frase_regla`, `linea_comision_vendedor`); sin ella, las de siempre, byte a
byte.

Por qué basta `comisiones_venta_usd` (revisión 6-oct-2026): desde la regla A
del 28-ago-2026 el API lo manda en 0 y solo los vuelos de la regla de
septiembre lo llenan; ≠ 0 en un avión ⇔ su cascada resta comisiones. Un
reparto sin comisiones absorbidas (periodo anterior, o vuelos sin comisión)
no contradice los textos de siempre. Desde el API 0.0.74 un mes ANTERIOR a
la regla también puede llenarlo con un pago «(avión)» (en cualquier fecha):
`absorbe_comisiones` / `comisiones_al_avion` no cambian (≠ 0 sigue siendo
«su cascada resta comisiones») y los textos con la bandera siguen siendo
verdad (la regla gobierna solo la parte bancaria).
"""

from app.schemas.reparto import RepartoAvion, RepartoPdfRequest
from app.schemas.reportes import REGLA_COMISIONES_RESPALDO

# Fila de la cascada (PDF) y columna E (Excel) con `comisiones_venta_usd`
# cuando aplica la regla; antes, «(-) Comisiones de venta» / «Comisiones
# venta» (la comisión del vendedor de antes del 28-ago-2026).
ETIQUETA_FILA = "(-) Comisiones (banco + vendedor)"
ETIQUETA_COLUMNA = "Comisiones (banco + vendedor)"
# Línea gris bajo «Otros ingresos VuelaTour» del PDF cuando el avión absorbe
# comisiones. Corta a propósito: la columna mide 120 mm y a 7.5 pt una línea
# más larga se encima con el monto (la frase completa va al pie).
LINEA_COMISION_VENDEDOR = (
    "   incluye comisión del vendedor cotizada (pre-IVA) — su provisión resta en Comisiones"
)
# La misma línea con la bandera del API 0.0.74 (dos categorías: solo el pago
# «(avión)» resta en Comisiones; misma longitud ±2 —88 caracteres, ~301 pt—:
# cabe en los 120 mm a 7.5 pt; test).
LINEA_COMISION_VENDEDOR_GASTO_REAL = (
    "   incluye comisión del vendedor cotizada (pre-IVA) — pago «(avión)» resta en Comisiones"
)


def solo_gasto_real(req: RepartoPdfRequest) -> bool:
    """¿El payload es del API 0.0.74+ (`comision_vendedor_solo_gasto_real`:
    la comisión del vendedor resta SOLO como gasto real)? Exige `is True`:
    None/False/ausente ⇒ los textos de siempre."""
    return req.comision_vendedor_solo_gasto_real is True


def linea_comision_vendedor(req: RepartoPdfRequest) -> str:
    """Línea gris del avión que absorbe comisiones: «pago «(avión)»» con la
    bandera del API 0.0.74; «provisión» sin ella (byte a byte)."""
    return LINEA_COMISION_VENDEDOR_GASTO_REAL if solo_gasto_real(req) else LINEA_COMISION_VENDEDOR


def absorbe_comisiones(avion: RepartoAvion) -> bool:
    """¿La cascada de este avión resta comisiones? El MISMO criterio con que
    el PDF imprime su fila y el Excel muestra la columna E (monto ≠ 0)."""
    return bool(avion.comisiones_venta_usd)


def comisiones_al_avion(req: RepartoPdfRequest) -> bool:
    """¿Algún avión del reparto absorbe comisiones? Decide los textos
    globales (resumen, pie y notas del Excel)."""
    return any(absorbe_comisiones(a) for a in req.aviones)


def regla(req: RepartoPdfRequest) -> str:
    """«regla sep-2026» del API (`regla_comisiones`) o, sin ella, «regla de
    septiembre 2026»."""
    return req.regla_comisiones or REGLA_COMISIONES_RESPALDO


def frase_regla(req: RepartoPdfRequest, inicio: str = "Con") -> str:
    """«Con la regla de septiembre 2026 (por fecha del vuelo), el avión
    absorbe la provisión de la comisión del vendedor y su parte de la
    comisión bancaria de sus cobros» (sin punto: cada documento la cierra).
    Con la bandera del API 0.0.74 (dos categorías): «… el avión absorbe su
    parte de la comisión bancaria de sus cobros y, en cualquier fecha, el
    pago al vendedor capturado como «Comisión del vendedor (avión)» (el de
    «Comisión del vendedor (VuelaTour)» lo paga VuelaTour)» — la regla
    gobierna solo la parte bancaria."""
    if solo_gasto_real(req):
        return (
            f"{inicio} la {regla(req)} (por fecha del vuelo), el avión absorbe su "
            "parte de la comisión bancaria de sus cobros y, en cualquier fecha, "
            "el pago al vendedor capturado como «Comisión del vendedor (avión)» "
            "(el de «Comisión del vendedor (VuelaTour)» lo paga VuelaTour)"
        )
    return (
        f"{inicio} la {regla(req)} (por fecha del vuelo), el avión absorbe la "
        "provisión de la comisión del vendedor y su parte de la comisión "
        "bancaria de sus cobros"
    )
