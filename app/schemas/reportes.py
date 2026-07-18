from pydantic import BaseModel, Field


class EscalaPdf(BaseModel):
    orden: int
    origen: str
    destino: str


class ExtraPdf(BaseModel):
    """Extra cobrado al cliente (línea del cotizador). Todo con default:
    los payloads viejos del API no mandaban extras (aditivo)."""

    concepto: str = ""
    monto_usd: float = 0
    # Moneda NATIVA del extra: los pagados en pesos se muestran "· $X MXN"
    # (requisito del cliente: cada concepto ligado a su moneda real).
    moneda: str = "USD"
    monto_nativo: float | None = None
    aplica_iva: bool = True


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
    # Detalle de TUAS por aeropuerto CON su moneda (strings ya formateados del
    # desglose canónico del motor, p. ej. "TUA PCE · $330.60 MXN × 4 pax = …").
    # Vacío en cotizaciones viejas → la plantilla conserva la línea única "TUAS".
    tuas_detalle: list[str] = Field(default_factory=list)
    extras: list[ExtraPdf] = Field(default_factory=list)
    extras_total_usd: float = 0
    viaticos_pernocta_usd: float = 0
    # Descuento YA en positivo (el API manda |ajuste| solo si fue descuento;
    # el redondeo hacia arriba nunca se muestra al cliente).
    descuento_usd: float = 0
    iva_pct: float = 0
    iva_usd: float = 0
    total_usd: float = 0
    # Total MXN EXACTO por composición (USD×TC + nativos MXN tal cual) y el TC
    # congelado de la cotización, para la línea final "Total MXN".
    total_mxn: float | None = None
    tc_usd_mxn: float | None = None
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
    # Litros (solo combustible): el Excel del equipo muestra litros y $/litro.
    litros: float | None = None


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
    # Sub-líneas INFORMATIVAS bajo la fila "TUAS" (una por aeropuerto, con su
    # moneda). La fila numérica tuas_usd sigue siendo la que cuadra la suma.
    tuas_detalle: list[str] = Field(default_factory=list)
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
    # ===== Economía del vuelo (formato de los Excel de control del equipo:
    # "Balance VGV" / "Dinero <mes>"). Todos aditivos: tolera API viejo. =====
    taco_inicio: float | None = None
    taco_fin: float | None = None
    gastos_total_usd: float = 0
    combustible_total_usd: float = 0
    gastos_sin_tc_count: int = 0
    gastos_sin_tc_mxn: float = 0
    venta_sin_iva_usd: float = 0
    remanente_usd: float | None = None
    ganancia_final_usd: float | None = None
    ganancia_x_hr_usd: float | None = None
    ganancia_pct: float | None = None
    notas: str | None = None
