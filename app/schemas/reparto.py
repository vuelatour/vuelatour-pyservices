"""Esquema del payload para el PDF de reparto de utilidades (doc 5.9)."""

from typing import Any

from pydantic import BaseModel, Field, field_validator

from app.schemas.reportes import regla_comisiones_liberal


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
    # Comisión del vendedor COTIZADA (pre-IVA): lo cobrado al cliente por
    # ella es ingreso de VuelaTour, que le paga al vendedor ('otros
    # movimientos'). Desde la regla de septiembre 2026 (API 0.0.65, por fecha
    # del vuelo) el avión absorbe su PROVISIÓN, que viaja dentro de
    # `RepartoAvion.comisiones_venta_usd` — este monto sigue siendo solo
    # informativo.
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


class RepartoTcOficialGastos(BaseModel):
    """Gastos MXN sin tc_gasto convertidos con el TC oficial de referencia
    del día del gasto (regla del cliente, 29-ago-2026). Ya están DENTRO de
    los montos del avión: solo alimentan una nota informativa."""

    count: int = 0
    monto_mxn: float = 0.0


class RepartoTcOficial(BaseModel):
    """Resumen global del TC oficial de respaldo usado en el periodo
    (open.er-api diario / BCE histórico): vuelos con cobros MXN sin TC y
    gastos MXN sin TC convertidos con él. Informativo, opcional."""

    vuelos: int = 0
    gastos: RepartoTcOficialGastos | None = None
    fuentes: list[str] = Field(default_factory=list)
    leyenda: str | None = None


class RepartoGastoCategoriaLinea(BaseModel):
    """Gastos del avión agrupados por categoría (`detalle.gastos_por_categoria`
    del API), ADITIVO (2-sep-2026): solo se imprime si el API lo manda.
    `categoria` es el CÓDIGO del enum (GAS, OTRO, …) y no cambia nunca;
    `etiqueta` es el texto amable homologado (panel/app/API) que se pinta en
    su lugar cuando viene — sin ella se imprime el código tal cual."""

    categoria: str = ""
    etiqueta: str | None = None
    # DIRECTO / INDIRECTO / PERMISO / FIJO / EXCLUIDO (mismo grupo del panel).
    grupo: str | None = None
    count: int = 0
    usd: float = 0.0
    sin_tc_count: int = 0
    sin_tc_mxn: float = 0.0
    tc_oficial_count: int = 0
    tc_oficial_mxn: float = 0.0


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
    # COMISIONES que absorbe el avión: restan antes del saldo (la cascada de
    # siempre). Historia: hasta el 28-ago-2026 (mañana) era la comisión del
    # vendedor; con la regla A (28-ago tarde) el API mandó 0 (la comisión es
    # ingreso de VuelaTour y su pago sale de VuelaTour). Desde la regla de
    # septiembre 2026 (API 0.0.65, vuelos con fecha desde
    # `comisiones_al_avion_desde`) trae la parte del avión de la comisión
    # BANCARIA de sus cobros + la PROVISIÓN de la comisión del vendedor
    # (fuente única `comisionesDelVuelo` del API). ≠ 0 ⇒ el PDF y el Excel la
    # pintan como «Comisiones (banco + vendedor)» con los textos de esa regla
    # (`reparto_comisiones.comisiones_al_avion`); 0 ⇒ los textos de siempre.
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
    # Gastos por categoría (opcional, 2-sep-2026): vacío = no se imprime. La
    # categoría se pinta con su `etiqueta` (o el código si no viene).
    gastos_por_categoria: list[RepartoGastoCategoriaLinea] = Field(default_factory=list)
    # ADITIVOS (29-ago-2026): convertidos con el TC oficial de referencia del
    # día (ya dentro de los montos; solo nota). None/0 = nada que anotar.
    gastos_tc_oficial: RepartoTcOficialGastos | None = None
    cobros_tc_oficial_count: int = 0


class RepartoPdfRequest(BaseModel):
    periodo_desde: str
    periodo_hasta: str
    generado: str
    aviones: list[RepartoAvion] = Field(default_factory=list)
    # Σ otros ingresos VuelaTour del periodo (informativo, fuera del reparto).
    otros_ingresos_vuelatour_total_usd: float = 0.0
    # Composición del Σ anterior (opcional; incluye comision_usd).
    otros_ingresos_vuelatour_desglose: RepartoOtrosIngresosDesglose | None = None
    # Nota global del TC oficial de respaldo (opcional, 29-ago-2026).
    tc_oficial: RepartoTcOficial | None = None
    # Nombre de la regla de comisiones a cargo del avión (6-oct-2026,
    # ADITIVO): la etiqueta que el API arma con la vigencia CONFIGURADA
    # («regla sep-2026», `etiquetaReglaComisiones`). Solo cambia el texto; sin
    # ella, «regla de septiembre 2026». Liberal: blanco o no-texto ⇒ None.
    regla_comisiones: str | None = None

    @field_validator("regla_comisiones", mode="before")
    @classmethod
    def _regla_liberal(cls, v: Any) -> str | None:
        return regla_comisiones_liberal(v)
