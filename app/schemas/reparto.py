"""Esquema del payload para el PDF de reparto de utilidades (doc 5.9)."""

from pydantic import BaseModel, Field


class RepartoSocioLinea(BaseModel):
    socio_nombre: str
    porcentaje: float
    monto_usd: float


class RepartoOtrosIngresosDesglose(BaseModel):
    """Composición (cotizada, pre-IVA) del ingreso de VuelaTour: TUAs,
    extras, pernocta y —regla A del 28-ago-2026 tarde— la comisión del
    vendedor; el IVA de todo va aparte. SOLO informativo: nada de esto se
    reparte. Todo opcional (un API viejo no manda el bloque)."""

    tuas_usd: float | None = None
    extras_usd: float | None = None
    pernocta_usd: float | None = None
    # Comisión del vendedor (pre-IVA): ingreso de VuelaTour, NO del avión;
    # su pago al vendedor sale de VuelaTour ('otros movimientos').
    comision_usd: float | None = None
    iva_usd: float | None = None


class RepartoVueloLinea(BaseModel):
    """Detalle de UN vuelo del avión (opcional, aditivo): solo se imprime si
    el API lo manda. `participacion` < 1 = vuelo MULTI-AVIÓN (regla B,
    28-ago-2026): la venta del avión se repartió entre los aviones en partes
    iguales por tramo vendido (ferries/tramos operativos no reparten) y el
    folio se imprime con el sufijo ' · 50 %'."""

    folio: str | int | None = None
    cliente: str | None = None
    fecha: str | None = None
    estado: str | None = None
    # Parte de la venta del avión cobrada / pendiente que toca a ESTE avión.
    cobrado_usd: float | None = None
    pendiente_usd: float | None = None
    participacion: float | None = None
    multi_avion: bool | None = None
    # Tramos de este avión en el vuelo (p. ej. "CUN→MID"); informativo.
    tramos_avion: str | None = None


class RepartoAvion(BaseModel):
    matricula: str
    modelo: str
    # VENTA DEL AVIÓN cobrada (regla 28-ago-2026): tiempo de vuelo + ajuste +
    # IVA proporcional. Los TUAs/extras/pernocta/comisión del vendedor
    # cobrados NO entran aquí. En vuelos MULTI-AVIÓN es la parte de este
    # avión (partes iguales por tramo vendido).
    ingresos_cobrado_usd: float
    # Otros ingresos de VuelaTour (TUAs/extras/pernocta + comisión del
    # vendedor cobrados, con su IVA): SOLO informativo — no se reparten ni
    # entran al saldo.
    otros_ingresos_vuelatour_usd: float = 0.0
    otros_ingresos_vuelatour_desglose: RepartoOtrosIngresosDesglose | None = None
    # Comisiones de venta: HASTA el 28-ago-2026 (mañana) se descontaban del
    # ingreso del avión. Regla A (28-ago tarde): la comisión del vendedor es
    # ingreso de VuelaTour y su pago sale de VuelaTour — ya NO es costo del
    # avión, el API manda 0. Se conserva por compat de shape: un payload
    # viejo con monto la sigue mostrando como deducción.
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
    # Detalle de vuelos del avión (opcional): vacío = no se imprime.
    vuelos: list[RepartoVueloLinea] = Field(default_factory=list)


class RepartoPdfRequest(BaseModel):
    periodo_desde: str
    periodo_hasta: str
    generado: str
    aviones: list[RepartoAvion] = Field(default_factory=list)
    # Σ otros ingresos VuelaTour del periodo (informativo, fuera del reparto).
    otros_ingresos_vuelatour_total_usd: float = 0.0
    # Composición del Σ anterior (opcional; incluye comision_usd).
    otros_ingresos_vuelatour_desglose: RepartoOtrosIngresosDesglose | None = None
