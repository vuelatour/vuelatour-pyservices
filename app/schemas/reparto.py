"""Esquema del payload para el PDF de reparto de utilidades (doc 5.9)."""

from pydantic import BaseModel, Field


class RepartoSocioLinea(BaseModel):
    socio_nombre: str
    porcentaje: float
    monto_usd: float


class RepartoAvion(BaseModel):
    matricula: str
    modelo: str
    # VENTA DEL AVIÓN cobrada (regla 28-ago-2026): tiempo de vuelo + ajuste +
    # IVA proporcional. Los TUAs/extras/pernocta cobrados NO entran aquí.
    ingresos_cobrado_usd: float
    # Otros ingresos de VuelaTour (TUAs/extras/pernocta cobrados, con su
    # IVA): SOLO informativo — no se reparten ni entran al saldo.
    otros_ingresos_vuelatour_usd: float = 0.0
    # Comisiones de venta (Itzy/Pablo/broker): se descuentan del ingreso.
    comisiones_venta_usd: float = 0.0
    # Pendiente de cobro = parte del AVIÓN (venta avión no cobrada).
    pendiente_cobro_usd: float = 0.0
    # Deuda COMPLETA del cliente (con TUAs/extras/pernocta); informativa.
    # 0 = API viejo o sin extras (no se muestra).
    pendiente_bruto_usd: float = 0.0
    horas_voladas_hr: float = 0.0
    gastos_directos_usd: float = 0.0
    gastos_indirectos_usd: float = 0.0
    permisos_usd: float = 0.0
    otros_usd: float = 0.0
    reserva_overhaul_usd: float = 0.0
    saldo_usd: float
    # Advertencias de integridad (montos que no pudieron entrar al balance)
    gastos_sin_tc_mxn: float = 0.0
    cobros_sin_tc_mxn: float = 0.0
    reserva_incompleta: bool = False
    # Suma de porcentajes de socios vigentes en el periodo. ≠ 100 = vigencias
    # traslapadas o incompletas: el reparto impreso estaría doble o corto.
    # Default 100 (payloads viejos no disparan la advertencia).
    reparto_porcentaje_total: float = 100.0
    reparto: list[RepartoSocioLinea] = Field(default_factory=list)


class RepartoPdfRequest(BaseModel):
    periodo_desde: str
    periodo_hasta: str
    generado: str
    aviones: list[RepartoAvion] = Field(default_factory=list)
    # Σ otros ingresos VuelaTour del periodo (informativo, fuera del reparto).
    otros_ingresos_vuelatour_total_usd: float = 0.0
