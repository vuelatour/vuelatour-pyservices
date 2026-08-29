"""Cardex de un ítem de inventario en formato LIBRO (29-ago-2026).

Réplica del cuaderno del cliente: bloque ENTRADAS | bloque SALIDAS lado a
lado, con venta, remanente y ganancia FIFO por salida. Todo llega YA
calculado de vuelatour-api (stock corriente, venta en pesos, ganancia):
aquí SOLO se renderiza el libro — a lo sumo se re-suman totales para
mostrar, jamás se recalcula negocio.

Esquemas ADITIVOS: todos los campos con default para tolerar skew de deploy.
"""

from pydantic import BaseModel, Field


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
    # Precio al que se vendió (o el costo FIFO si la salida fue "a costo").
    venta_unitaria: float | None = None
    venta_total: float | None = None
    # Stock corriente DESPUÉS de la salida.
    remanente: float = 0
    # Venta total − costo FIFO de las capas consumidas (la manda el API).
    ganancia: float | None = None
    # Matrícula del avión; 'FLOTA' en salidas para toda la flota.
    vendido_a: str = ""


class CardexLibroRequest(BaseModel):
    titulo: str = "Cardex"
    item_nombre: str = ""
    numero_parte: str | None = None
    unidad: str | None = None
    generado: str | None = None
    # Moneda de TODOS los montos del libro (hoy siempre 'MXN').
    moneda: str = "MXN"
    entradas: list[CardexLibroEntrada] = Field(default_factory=list)
    salidas: list[CardexLibroSalida] = Field(default_factory=list)
    total_compra: float = 0
    total_venta: float = 0
    total_ganancia: float = 0
