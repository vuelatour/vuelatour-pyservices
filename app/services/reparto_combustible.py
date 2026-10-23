"""Combustible SEPARADO de los gastos directos en el reparto a socios
(23-oct-2026, API 0.0.103, invariante 104 del API): textos y helpers del PDF
y del Excel (fuente única de los dos documentos, como `reparto_comisiones`).

Pedido del cliente (23-oct-2026): «debemos separar gastos directos del
combustible, porque los está uniendo, porque no nos está coincidiendo con el
excel». El cliente concilia el reparto contra su Excel «Balance» por avión,
que lleva el combustible en su propia línea y los costos directos sin él:

    Venta cobrada
    (-) Comisiones de venta
    (-) Costos directos (OPERACIONES, piloto, otros)   <- SIN combustible
    (-) Provisión AFAC
    = UTILIDAD ANTES DE GASTOS
    (-) Combustible del mes
    (-) Gastos indirectos (incluye refacciones)
    (-) Permisos
    (-) Fijos prorrateados          (línea propia del reparto; casi siempre 0)
    (+) Otros ingresos
    = UTILIDAD DESPUÉS DE GASTOS = saldo a repartir

Hasta el API 0.0.102 «(-) Gastos directos» del reparto UNÍA el combustible
(categoría GAS, Gasavión/Turbosina) con los demás gastos directos del vuelo:
N621TX sep-2026 decía «Gastos directos 187,865.72» = 147,277.58 de
combustible + 40,588.14 de directos, y no cuadraba línea por línea. Con el
0.0.103 `gastos_directos_*` viene SIN combustible y llegan, por avión,
`combustible_*` (el GAS del periodo) y `utilidad_antes_gastos_*` (venta
cobrada − comisiones − gastos directos − provisión AFAC, redondeado a 2
decimales con componentes ya redondeados: la línea «UTILIDAD ANTES DE
GASTOS» del Excel); `saldo_*` = utilidad antes de gastos − combustible −
indirectos − permisos − fijos + otros ingresos del avión. Todo lo calcula el
API: aquí solo se pinta EN EL ORDEN DEL EXCEL. La PRESENCIA de
`combustible_usd` (no None; 0.0 cuenta) enciende ese orden; sin ella (API ≤
0.0.102) los documentos de siempre, byte a byte.
"""

from app.schemas.reparto import RepartoAvion, RepartoPdfRequest

# Filas de la cascada del PDF (Helvetica: sin «−» ni «→», que se dibujan
# como caja; «(-)» como las demás filas). La etiqueta de los directos dice
# «(sin combustible)» para que quien concilia con el Excel («Costos
# directos») no busque ahí el combustible.
FILA_GASTOS_DIRECTOS_SIN_COMBUSTIBLE = "(-) Gastos directos (sin combustible)"
FILA_UTILIDAD_ANTES_GASTOS = "= Utilidad antes de gastos"
FILA_COMBUSTIBLE = "(-) Combustible"
# Cláusula del pie del PDF (va tras la de la cascada del balance y antes de
# «Montos en …»); la provisión AFAC solo se nombra con `reparto_como_balance`.
_PIE_COMBUSTIBLE_INICIO = (
    " El combustible (gasavión/turbosina) va en su propia línea, fuera de los gastos "
    "directos, como en el Balance: Utilidad antes de gastos = venta del avión cobrada - "
    "comisiones - gastos directos"
)

# Encabezados del Excel. En la tabla principal el bloque «Gastos directos
# (sin combustible) · Provisión AFAC · Utilidad antes de gastos · Combustible»
# va donde el Excel del cliente lo lee (tras los directos); sin la llave, la
# columna «Provisión AFAC» sigue AL FINAL como desde el 20-oct-2026.
ENCABEZADO_GASTOS_DIRECTOS_SIN_COMBUSTIBLE = "Gastos directos (sin combustible)"
ENCABEZADO_GASTOS_DIRECTOS_SIN_COMBUSTIBLE_PESOS = "Gastos directos (sin combustible) (MXN)"
ENCABEZADO_UTILIDAD_ANTES_GASTOS = "Utilidad antes de gastos"
ENCABEZADO_UTILIDAD_ANTES_GASTOS_PESOS = "Utilidad antes de gastos (MXN)"
ENCABEZADO_COMBUSTIBLE = "Combustible"
ENCABEZADO_COMBUSTIBLE_PESOS = "Combustible (MXN)"
NOTA_COMBUSTIBLE_SEPARADO = (
    "Gastos directos SIN combustible: el combustible (gasavión/turbosina) va en su propia "
    "columna, como en el Balance. Utilidad antes de gastos = venta cobrada − comisiones − "
    "gastos directos − provisión AFAC; Saldo disponible = utilidad antes de gastos − "
    "combustible − gastos indirectos − permisos − fijos + otros ingresos del avión."
)
NOTA_ENCABEZADO_UTILIDAD_ANTES_GASTOS = (
    "Venta del avión cobrada − comisiones − gastos directos (sin combustible) − provisión "
    "AFAC: la línea «UTILIDAD ANTES DE GASTOS» del Balance."
)


def combustible_separado(avion: RepartoAvion) -> bool:
    """¿El API mandó el combustible APARTE de los gastos directos (0.0.103)?
    La presencia de `combustible_usd` decide (0.0 también cuenta); None =
    API previo ⇒ la cascada de siempre."""
    return avion.combustible_usd is not None


def hay_combustible_separado(req: RepartoPdfRequest) -> bool:
    """¿Algún avión del reparto trae el combustible aparte? Decide el pie del
    PDF y el layout del Excel (una sola tabla para todos los aviones)."""
    return any(combustible_separado(a) for a in req.aviones)


def pie_combustible(req: RepartoPdfRequest) -> str:
    """Cláusula del pie del PDF con el combustible separado; con la cascada
    del balance la utilidad antes de gastos también resta la provisión AFAC."""
    afac = " - provisión AFAC" if req.reparto_como_balance else ""
    return f"{_PIE_COMBUSTIBLE_INICIO}{afac}."


def utilidad_antes_gastos_usd(avion: RepartoAvion, *, como_balance: bool) -> float:
    """La línea «UTILIDAD ANTES DE GASTOS» en dólares: la que manda el API
    (`utilidad_antes_gastos_usd`) y, solo si no viene, venta cobrada −
    comisiones − gastos directos (sin combustible) − provisión AFAC (esta
    última solo con la cascada del balance), redondeada a 2 decimales."""
    if avion.utilidad_antes_gastos_usd is not None:
        return avion.utilidad_antes_gastos_usd
    afac = (avion.afac_reservado_usd or 0.0) if como_balance else 0.0
    return round(
        avion.ingresos_cobrado_usd
        - avion.comisiones_venta_usd
        - avion.gastos_directos_usd
        - afac,
        2,
    )


def utilidad_antes_gastos_mxn(avion: RepartoAvion, *, como_balance: bool) -> float | None:
    """Gemela en PESOS: la del API y, si no viene, la resta de los `*_mxn`
    del avión (las comisiones cuentan con el mismo criterio con que la
    cascada en pesos imprime su fila). Un componente en None ⇒ None («—»):
    aquí jamás se inventa un peso."""
    if avion.utilidad_antes_gastos_mxn is not None:
        return avion.utilidad_antes_gastos_mxn
    comisiones = (
        avion.comisiones_venta_mxn
        if avion.comisiones_venta_usd or avion.comisiones_venta_mxn
        else 0.0
    )
    afac = avion.afac_reservado_mxn if como_balance else 0.0
    partes = (avion.ingresos_cobrado_mxn, comisiones, avion.gastos_directos_mxn, afac)
    if any(x is None for x in partes):
        return None
    venta, comisiones, directos, afac = partes
    return round(venta - comisiones - directos - afac, 2)
