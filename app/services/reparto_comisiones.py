"""Comisiones a cargo del avión en el reparto a socios: textos del PDF y del
Excel (fuente única de los dos documentos).

Regla de septiembre 2026 (6-oct-2026, API 0.0.65; respuestas del cliente: la
comisión del vendedor la absorbe el avión, lo cobrado al cliente por ella
sigue siendo ingreso de VuelaTour, desde septiembre 2026 por fecha del
vuelo). El API suma en `comisiones_venta_usd` —el campo de siempre, que la
cascada ya resta antes del saldo— la parte del avión de la comisión bancaria
de sus cobros y la PROVISIÓN de la comisión del vendedor (fuente única
`comisionesDelVuelo`, la misma del balance). Aquí no se calcula nada: cambian
la etiqueta de esa fila/columna y las frases que decían «su pago no es costo
del avión», y SOLO cuando un avión de verdad absorbe comisiones: sin eso,
los textos de siempre.

Por qué basta `comisiones_venta_usd` (revisión 6-oct-2026): desde la regla A
del 28-ago-2026 el API lo manda en 0 y solo los vuelos de la regla de
septiembre lo llenan; ≠ 0 en un avión ⇔ su cascada resta comisiones. Un
reparto sin comisiones absorbidas (periodo anterior, o vuelos sin comisión)
no contradice los textos de siempre.
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
    comisión bancaria de sus cobros» (sin punto: cada documento la cierra)."""
    return (
        f"{inicio} la {regla(req)} (por fecha del vuelo), el avión absorbe la "
        "provisión de la comisión del vendedor y su parte de la comisión "
        "bancaria de sus cobros"
    )
