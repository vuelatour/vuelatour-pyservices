from pydantic import BaseModel, Field


class EscalaPdf(BaseModel):
    orden: int
    origen: str
    destino: str


class CotizacionPdfRequest(BaseModel):
    folio: str
    fecha: str | None = None  # fecha de la cotización (texto ya formateado o ISO)
    empresa: str = "VuelaTour — Aero Charter Cancún"
    cliente: str
    origen: str
    destino: str
    tipo: str = "REDONDO"
    pasajeros: int = 1
    fecha_traslado_inicial: str | None = None
    fecha_traslado_final: str | None = None
    escalas: list[EscalaPdf] = Field(default_factory=list)
    tiempo_cobrable_hr: float | None = None
    tarifa_hora_usd: float | None = None
    subtotal_usd: float = 0
    tuas_usd: float = 0
    iva_pct: float = 0
    iva_usd: float = 0
    total_usd: float = 0
    moneda: str = "USD"
    notas: str | None = None


# ===== Reporte consolidado de UN vuelo (cotización + ingreso + combustible +
# tacómetro + gastos). Lo arma vuelatour-api y lo renderiza pyservices. =====
class ReporteVueloTramo(BaseModel):
    orden: int
    ruta: str
    pasajeros: int | None = None
    pasajeros_nombres: str | None = None
    taco_salida: float | None = None
    taco_llegada: float | None = None
    horas: float | None = None
    es_ferry: bool = False


class ReporteVueloLinea(BaseModel):
    fecha: str | None = None
    concepto: str = ""
    detalle: str | None = None
    moneda: str | None = None
    monto: float | None = None


class ReporteVueloRequest(BaseModel):
    generado: str
    folio: str
    cliente: str = ""
    aeronave: str | None = None
    piloto: str | None = None
    copiloto: str | None = None
    tipo: str = ""
    estado: str = ""
    ruta: str = ""
    fecha_vuelo: str | None = None
    fecha_traslado_final: str | None = None
    pasajeros: int = 0
    pasajeros_nombres: str | None = None
    # Cotización
    tarifa_tipo: str | None = None
    tarifa_hora_usd: float | None = None
    tiempo_cobrable_hr: float | None = None
    subtotal_usd: float = 0
    tuas_usd: float = 0
    iva_usd: float = 0
    viaticos_pernocta_usd: float = 0
    extras_total_usd: float = 0
    ajuste_final_usd: float = 0
    total_usd: float = 0
    total_mxn: float | None = None
    tc_usd_mxn: float | None = None
    # Comisión del vendedor (interna): el cliente paga el total completo;
    # neto = total − comisión es lo que queda a VuelaTour.
    comision_vendedor_usd: float = 0
    comision_vendedor_nombre: str | None = None
    neto_vuelatour_usd: float | None = None
    metodo_cobro: str | None = None
    # Secciones
    tramos: list[ReporteVueloTramo] = Field(default_factory=list)
    # Comparación horas cotizadas (ruta comercial) vs voladas (ruta operativa):
    # el delta positivo es la utilidad operativa que el cliente busca maximizar.
    horas_cotizadas_hr: float | None = None
    horas_voladas_hr: float | None = None
    horas_delta_hr: float | None = None
    notas_horas: list[str] = Field(default_factory=list)
    cobros: list[ReporteVueloLinea] = Field(default_factory=list)
    total_cobrado_usd: float = 0
    # Comisiones bancarias de los cobros (terminal/transferencia): el banco
    # deposita menos de lo que pagó el cliente. Neto = cobrado − comisiones =
    # lo que realmente entró a la cuenta (pedido del cliente: el reporte no
    # cuadraba con el estado de cuenta).
    comision_banco_usd: float = 0
    total_cobrado_neto_usd: float | None = None
    saldo_usd: float = 0
    combustible: list[ReporteVueloLinea] = Field(default_factory=list)
    gastos: list[ReporteVueloLinea] = Field(default_factory=list)
    notas: str | None = None
