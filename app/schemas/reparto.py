"""Esquema del payload para el PDF de reparto de utilidades (doc 5.9)."""

from pydantic import BaseModel, Field


class RepartoSocioLinea(BaseModel):
    socio_nombre: str
    porcentaje: float
    monto_usd: float


class RepartoAvion(BaseModel):
    matricula: str
    modelo: str
    ingresos_cobrado_usd: float
    pendiente_cobro_usd: float = 0.0
    gastos_directos_usd: float = 0.0
    gastos_indirectos_usd: float = 0.0
    permisos_usd: float = 0.0
    otros_usd: float = 0.0
    reserva_overhaul_usd: float = 0.0
    saldo_usd: float
    reparto: list[RepartoSocioLinea] = Field(default_factory=list)


class RepartoPdfRequest(BaseModel):
    periodo_desde: str
    periodo_hasta: str
    generado: str
    aviones: list[RepartoAvion] = Field(default_factory=list)
