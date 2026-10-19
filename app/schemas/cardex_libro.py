"""Cardex de un ítem de inventario en formato LIBRO (29-ago-2026).

Réplica del cuaderno del cliente: bloque ENTRADAS | bloque SALIDAS lado a
lado, con venta, remanente y ganancia por salida. Todo llega YA calculado
de vuelatour-api (stock corriente, venta en pesos, ganancia): aquí SOLO se
renderiza el libro — a lo sumo se re-suman totales para mostrar, jamás se
recalcula negocio.

Regla de costo (25-sep-2026, API 0.0.36): el costo de una salida es el
ÚLTIMO PRECIO DE COMPRA vigente el día de la salida (congelado en su fila)
y los montos en dólares llegan YA en pesos con el T.C. oficial de su día
(compras: el del día de la compra; ventas: el del día de la venta). Hasta
el API 0.0.35 era costo FIFO. El API lo explica en `nota` (subtítulo).

Esquemas ADITIVOS: todos los campos con default para tolerar skew de deploy.
"""

from pydantic import BaseModel, Field, field_validator


class CardexLibroEntrada(BaseModel):
    """Fila del bloque ENTRADAS (montos en la moneda del libro, hoy MXN)."""

    fecha: str = ""
    cantidad: float = 0
    # Ítem + referencia/proveedor; DEVOLUCION/AJUSTE llegan con su prefijo.
    descripcion: str = ""
    valor_compra_unitario: float | None = None
    valor_compra_total: float | None = None
    # Stock corriente DESPUÉS de este movimiento.
    stock_despues: float = 0


class CardexLibroSalida(BaseModel):
    """Fila del bloque SALIDAS (montos en la moneda del libro, hoy MXN)."""

    fecha: str = ""
    cantidad: float = 0
    descripcion: str = ""
    # Precio al que se vendió (o el costo de la pieza si la salida fue "a
    # costo": último precio de compra desde el API 0.0.36, FIFO antes).
    venta_unitaria: float | None = None
    venta_total: float | None = None
    # Stock corriente DESPUÉS de la salida.
    remanente: float = 0
    # Venta total − costo de la pieza, los dos en pesos con el MISMO T.C.
    # (el del día de la venta). La manda el API.
    ganancia: float | None = None
    # Matrícula del avión; 'FLOTA' en salidas para toda la flota.
    vendido_a: str = ""


class CardexLibroMovimiento(BaseModel):
    """Una fila de la LISTA por fecha (18-oct-2026, API 0.0.92).

    Compras, devoluciones, ajustes y salidas en UN orden (el del cardex:
    fecha, alta). Montos en la moneda del libro (hoy MXN), ya calculados.
    """

    fecha: str = ""
    # YYYY-MM: agrupa la lista por mes.
    mes: str = ""
    # ENTRADA | DEVOLUCION | AJUSTE | SALIDA.
    tipo: str = ""
    # «Aircraft Spruce · ref F-12» / «XA-VGV» / «Toda la flota · a costo».
    detalle: str = ""
    # Unidades que entraron (null en una salida) / que salieron (null si no).
    entra: float | None = None
    sale: float | None = None
    # Stock DESPUÉS del movimiento.
    existencia: float = 0
    # Compra: costo unitario; salida: precio al que se vendió (o el costo).
    precio_unitario: float | None = None
    total: float | None = None
    # Salida: venta − costo (0 a costo); null en entradas o sin T.C.
    ganancia: float | None = None
    a_costo: bool = False


class CardexLibroMes(BaseModel):
    """Un renglón del resumen POR MES (18-oct-2026, API 0.0.92)."""

    mes: str = ""
    # «Septiembre 2026» / «Octubre 2026 (en curso)».
    etiqueta: str = ""
    stock_inicio: float = 0
    # Compras + devoluciones + ajustes.
    entradas: float = 0
    compras: float = 0
    # Todas las salidas (con precio o a costo): lo que baja el stock.
    unidades_vendidas: float = 0
    stock_final: float = 0
    movimientos: int = 0
    vendido: float | None = None
    ganancia: float | None = None
    en_curso: bool = False


class CardexLibroRequest(BaseModel):
    titulo: str = "Cardex"
    item_nombre: str = ""
    numero_parte: str | None = None
    unidad: str | None = None
    generado: str | None = None
    # Moneda de TODOS los montos del libro (hoy siempre 'MXN').
    moneda: str = "MXN"
    # ADITIVO (25-sep-2026, API 0.0.36): cómo se convirtieron los montos
    # («Compras al T.C. oficial del día de la compra; ventas al del día de
    # la venta. Costo = último precio de compra vigente ese día.»). Se pinta
    # TAL CUAL en el subtítulo, tras «Montos en MXN». None = API previo ⇒
    # subtítulo de siempre.
    nota: str | None = None
    entradas: list[CardexLibroEntrada] = Field(default_factory=list)
    salidas: list[CardexLibroSalida] = Field(default_factory=list)
    # ADITIVOS (18-oct-2026, API 0.0.92; pedido del cliente: el cardex «en
    # forma de lista por orden de fecha» con «stock inicio del mes y stock
    # final del mes y cantidad de unidades vendidas»). `movimientos` presente
    # (aunque vacío) ⇒ el libro nuevo (lista + resumen por mes + la hoja
    # ENTRADAS | SALIDAS de siempre al final); None ⇒ el libro de siempre.
    movimientos: list[CardexLibroMovimiento] | None = None
    meses: list[CardexLibroMes] = Field(default_factory=list)
    total_compra: float = 0
    total_venta: float = 0
    total_ganancia: float = 0

    @field_validator("meses", mode="before")
    @classmethod
    def _meses_liberal(cls, v):
        # null del API ⇒ sin resumen (jamás un 422 por un dato de apoyo).
        return [] if v is None else v
