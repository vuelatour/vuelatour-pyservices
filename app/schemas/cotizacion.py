"""Esquema del payload que NestJS envia para generar el PDF de cotizacion."""

from pydantic import BaseModel, Field


class LineaCotizacion(BaseModel):
    concepto: str
    monto_usd: float


class CotizacionPdfRequest(BaseModel):
    folio: int
    version: int = 1
    fecha: str
    cliente: str
    aeronave: str
    ruta: str
    tipo_vuelo: str
    pasajeros: int
    lineas: list[LineaCotizacion] = Field(default_factory=list)
    subtotal_usd: float
    tuas_usd: float = 0.0
    iva_usd: float = 0.0
    total_usd: float
    tc_usd_mxn: float | None = None
    notas: str | None = None
