from typing import Any

from pydantic import BaseModel, ConfigDict, Field


class EscalaPdf(BaseModel):
    orden: int
    origen: str
    destino: str
    # Fecha de PARED (YYYY-MM-DD, sin hora ni zona) que verá el cliente en
    # el PDF para este tramo (3-sep-2026): SOLO presentación — no es la
    # salida operativa ni mueve fechas de vuelo. None = tramo sin fecha
    # ('—' si otro tramo sí la trae; sin ninguna, la tabla no lleva la
    # columna). ADITIVO: un API viejo no la manda y el PDF queda igual.
    fecha: str | None = None


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


class MapaPuntoPdf(BaseModel):
    """Tramo con coordenadas para el mapa de ruta del PDF (26-ago)."""

    orden: int
    origen_iata: str
    destino_iata: str
    o_lat: float
    o_lon: float
    d_lat: float
    d_lon: float
    es_ferry: bool = False


class MapaSvgRequest(BaseModel):
    """Body de `POST /reportes/cotizacion/mapa-svg` (8-sep-2026): los mismos
    `mapa_puntos` que viajan en `CotizacionPdfRequest` — el panel pide el
    mapa suelto para pintarlo inline en la hoja que edita."""

    mapa_puntos: list[MapaPuntoPdf] = []


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
    # Ruta VISIBLE ya resuelta por el API ("CUN → AZP → BZE → CZM → CUN"):
    # con tramos ocultos (pdf_oculto) el título NO puede derivarse de las
    # escalas (el walk une los huecos en NestJS, que es quien filtra).
    # ADITIVO: sin el campo (payload viejo) el título se arma de las escalas
    # como siempre — tolera skew de deploy en ambos sentidos.
    ruta: str | None = None
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
    # Presentación configurable por cotización (27-ago): tarifa/hr apagada
    # e itinerario prendido por defecto (aditivo: tolera skew de deploy).
    mostrar_tarifa_hora: bool = False
    mostrar_itinerario: bool = True
    # ===== PDF profesional (26-ago): mapa, matrícula y fotos del avión =====
    matricula: str | None = None
    # Data-URIs (base64) firmadas y descargadas por el API; None = sin foto.
    foto_exterior: str | None = None
    foto_interior: str | None = None
    # ===== Ficha comercial de la aeronave (26-ago v2, mockup del cliente):
    # página 2 = exterior ancho + interior con tarjeta "De un vistazo" +
    # características. Todo opcional (defaults) para tolerar skew de deploy.
    avion_modelo: str | None = None
    # Avión EXTERNO (28-ago, venta broker): ficha capturada a mano, p.ej.
    # "HAWKER 400 A · XA-REG". Cuando viene, se muestra bajo la ruta; la
    # página "La aeronave" no aplica (no hay fotos/ficha de flota).
    avion_externo: str | None = None
    avion_velocidad_kts: float | None = None
    avion_pasajeros: int | None = None
    avion_num_motores: int | None = None
    avion_motor_hp: int | None = None
    avion_caracteristicas: list[str] = Field(default_factory=list)
    # Duración estimada del tramo MÁS LARGO del viaje (horas decimales).
    avion_tiempo_tramo_hr: float | None = None
    mapa_puntos: list[MapaPuntoPdf] = Field(default_factory=list)
    # ===== Avión COTIZADO en la hoja 1 (feedback del cliente 4-sep-2026):
    # el MODELO con el que se pactó (snapshot) — a veces se cotiza en un
    # avión y la ruta operativa va en otro. NUNCA matrícula. Con tramos en
    # aviones distintos el API manda los modelos distintos en orden de
    # tramo (`modelos_cotizados`); si viene vacío se usa
    # `aeronave_cotizada_modelo`. ADITIVO: sin los campos (API viejo) la
    # hoja 1 no pinta la línea.
    aeronave_cotizada_modelo: str | None = None
    modelos_cotizados: list[str] = Field(default_factory=list)



# ===== Cotización de GRUPO (4-sep-2026): varios aviones para un mismo
# cliente, UN documento y UN total. El API (`groups-pdf.service`) manda el
# consolidado YA calculado (Σ de los desgloses canónicos de los hijos vivos,
# comisión del vendedor y redondeo absorbidos en "Servicio aéreo"); aquí
# SOLO se pinta. TODO con default y extra="ignore" (aditivo): tolera skew de
# deploy en ambos sentidos. =====
class CotizacionGrupoTramoPdf(EscalaPdf):
    """Tramo de la plantilla comercial del grupo (solo visibles, renumerados
    1..N por el API). Hereda orden/origen/destino/fecha de `EscalaPdf` para
    reusar la tabla de itinerario del PDF de un avión."""

    es_ferry: bool = False
    requiere_pernocta: bool = False
    tipo_parada: str = "NORMAL"
    servicio_notas: str | None = None


class CotizacionGrupoLineaPdf(BaseModel):
    """Línea del desglose consolidado APTA para el cliente. `cantidad` ×
    `unitario` viajan cuando la línea nació así (tour por persona: "44 ×
    $85.00"); `moneda` es la nativa de la línea (MXN en extras en pesos)."""

    clave: str = ""
    concepto: str = ""
    monto_usd: float = 0
    cantidad: float | None = None
    unitario: float | None = None
    moneda: str | None = None


class CotizacionGrupoExtraPdf(ExtraPdf):
    """Extra del grupo (misma forma que `ExtraPdf` + cantidad × unitario)."""

    cantidad: float | None = None
    unitario: float | None = None


class CotizacionGrupoAvionPdf(BaseModel):
    """Un avión del grupo (hijo vivo). `matricula` SIEMPRE viaja: la plantilla
    aplica la regla vigente de mostrarla o no (oculta salvo VGV).
    `subtotal_usd` (total del hijo) se pinta SOLO con
    `mostrar_subtotal_por_avion`; `tarifa_hora_usd` SOLO con `mostrar_tarifa`.
    Fotos: data URI base64 (solo el PRIMER avión de cada modelo las trae) con
    la URL pública como respaldo."""

    posicion: int = 0
    modelo: str | None = None
    matricula: str | None = None
    asientos: int | None = None
    pasajeros: int = 0
    rotaciones: int = 1
    tiempo_hr: float | None = None
    salida_estimada: str | None = None
    subtotal_usd: float | None = None
    tarifa_hora_usd: float | None = None
    velocidad_kts: float | None = None
    num_motores: int | None = None
    motor_hp: int | None = None
    caracteristicas: list[str] = Field(default_factory=list)
    foto_exterior: str | None = None
    foto_interior: str | None = None
    foto_exterior_url: str | None = None
    foto_interior_url: str | None = None


class CotizacionGrupoPdfRequest(BaseModel):
    """Payload de `POST /reportes/cotizacion-grupo` (contrato
    `CotizacionGrupoPdfRequest` de pyservices.service.ts)."""

    model_config = ConfigDict(extra="ignore")

    folio_grupo: str = ""  # "G-12"
    folio: int = 0
    nombre: str = ""
    empresa: str = "VuelaTour — Aero Charter Cancún"
    cliente: str = "Cliente"
    fecha: str | None = None  # ISO de la salida del grupo
    pasajeros_total: int = 0
    aviones_total: int = 0
    # Ruta VISIBLE ya resuelta por el API ("CUN → CZA → CUN").
    ruta: str | None = None
    itinerario: list[CotizacionGrupoTramoPdf] = Field(default_factory=list)
    mapa_puntos: list[MapaPuntoPdf] = Field(default_factory=list)
    # Desglose consolidado apto para el cliente, en el orden del API. NUNCA
    # trae COMISION_VENDEDOR ni redondeo; la plantilla los filtra igual.
    desglose_consolidado: list[CotizacionGrupoLineaPdf] = Field(default_factory=list)
    servicio_aereo_usd: float = 0
    horas_total_hr: float = 0
    tuas_usd: float = 0
    tuas_detalle: list[str] = Field(default_factory=list)
    extras: list[CotizacionGrupoExtraPdf] = Field(default_factory=list)
    extras_total_usd: float = 0
    viaticos_pernocta_usd: float = 0
    descuento_usd: float = 0
    # = total − IVA (lo manda el API); None (API viejo) → se deriva igual
    # que en el PDF de un avión.
    subtotal_usd: float | None = None
    iva_pct: float = 0
    iva_usd: float = 0
    total_usd: float = 0
    total_mxn: float | None = None
    tc_usd_mxn: float | None = None
    precio_por_persona_usd: float | None = None
    moneda: str = "USD"
    # Toggles de presentación (defaults = los de la cabecera vuelo_grupo).
    mostrar_precio_por_persona: bool = True
    mostrar_tarifa: bool = False
    mostrar_anexo_aviones: bool = True
    mostrar_subtotal_por_avion: bool = False
    mostrar_itinerario: bool = True
    aviones: list[CotizacionGrupoAvionPdf] = Field(default_factory=list)
    notas: str | None = None
    condiciones: str | None = None  # reservado (texto de condiciones)


# ===== Recibo de pago por cobro (documento NO fiscal). El API manda TODO
# calculado (folio, cobrado a la fecha, saldo, liquidado): aquí SOLO se pinta.
# TODO con default (aditivo): tolera skew de deploy en ambos sentidos. =====
class ReciboAbonoPdf(BaseModel):
    """Abono previo del vuelo (historial del recibo). Un monto NEGATIVO es un
    reembolso: se pinta restando, con su etiqueta."""

    fecha: str | None = None
    monto: float = 0
    moneda: str = "USD"
    # "Abono" / "Reembolso" (la pone el API); None = sin etiqueta.
    etiqueta: str | None = None


class ReciboPdfRequest(BaseModel):
    folio_recibo: str = ""
    empresa: str = "VuelaTour — Aero Charter Cancún"
    cliente: str = ""
    vuelo_folio: str = ""
    ruta: str = ""
    fecha_vuelo: str | None = None
    fecha_cobro: str | None = None
    # Monto BRUTO que pagó el cliente, en su moneda (la comisión bancaria
    # NUNCA aparece en el recibo).
    monto: float = 0
    moneda: str = "USD"
    # TC usado y equivalente en USD (solo cobros MXN); None = no se pinta.
    tc_usd_mxn: float | None = None
    equivalente_usd: float | None = None
    # Método YA legible ("Transferencia", "BillPocket"…), armado por el API.
    metodo: str = ""
    cuenta_destino: str | None = None
    referencia: str | None = None
    # Recibo de un SOBRE de grupo (4-sep-2026): el API manda vuelo_folio="G-12"
    # y en `ruta` el concepto ("7 aeronaves · 44 pasajeros · CUN → CZA → CUN").
    # Aditivo: con grupo_folio se etiqueta "Grupo" en vez de "Vuelo".
    grupo_folio: str | None = None
    aviones_n: int | None = None
    pasajeros_total: int | None = None
    # Resumen del vuelo (fuente única cobrosEnUsd, neto de reembolsos).
    total_cotizacion_usd: float = 0
    cobrado_a_la_fecha_usd: float = 0
    saldo_pendiente_usd: float = 0
    liquidado: bool = False
    # Cobros MXN sin TC fuera de la suma (jamás desaparecen en silencio).
    sin_tc_nota: str | None = None
    notas: str | None = None
    cobros_previos: list[ReciboAbonoPdf] = Field(default_factory=list)


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
    # Etiqueta amable de la categoría (2-sep-2026, aditivo): en las líneas de
    # gastos `concepto` trae el CÓDIGO del enum (OTRO, TUAS, …); cuando el
    # API manda `etiqueta` se imprime ella en la columna Categoría.
    etiqueta: str | None = None


class ReporteVueloParticipacion(BaseModel):
    """Parte de la VENTA DEL AVIÓN que toca a una matrícula en un vuelo
    MULTI-AVIÓN (regla B, 28-ago-2026): factores Σ == 1 y venta ya repartida
    por el API (centavos exactos). Todo opcional salvo la matrícula."""

    aeronave_id: str | None = None
    matricula: str = ""
    factor: float | None = None
    # Tramos del avión en el vuelo: el API puede mandar cuenta (int), texto
    # ("CUN→MID") o lista de órdenes — se formatea tolerante.
    tramos: Any = None
    venta_usd: float | None = None


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
    # neto = total − comisión es lo que queda a VuelaTour. Regla A
    # (28-ago-2026 tarde): la comisión es INGRESO de VuelaTour (va dentro de
    # otros_ingresos_vuelatour_usd) y su pago sale de VuelaTour, no del
    # avión; este reporte —economía COMPLETA del vuelo— la sigue restando
    # UNA sola vez en su ganancia.
    comision_vendedor_usd: float = 0
    comision_vendedor_nombre: str | None = None
    neto_vuelatour_usd: float | None = None
    # PAGO al vendedor = comisión + su IVA cuando grava (`pagoVendedorUsd`,
    # fuente única del API): es lo que se resta en "(−) Pago comisión
    # vendedor" y en el neto. None = API viejo → se usa comision_vendedor_usd.
    pago_vendedor_usd: float | None = None
    # Regla 28-ago-2026 (informativos): del total, cuánto es VENTA DEL AVIÓN
    # (tiempo + ajuste + IVA proporcional) y cuánto ingreso de VuelaTour
    # (TUAs/extras/pernocta/comisión del vendedor + su IVA). None = API
    # viejo → no se pintan.
    venta_avion_usd: float | None = None
    otros_ingresos_vuelatour_usd: float | None = None
    # Regla B (28-ago-2026): vuelo MULTI-AVIÓN — cómo se repartió la venta
    # del avión entre las matrículas (por tramo). None o 1 elemento = vuelo
    # de un solo avión (no se pinta).
    participacion_aviones: list[ReporteVueloParticipacion] | None = None
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


# ===== Cotización INTERNA v2 (8-sep-2026): UNA hoja carta para la oficina
# con SOLO lo de la COTIZACIÓN (feedback de administración con la foto de su
# formato): fecha protagonista = día del vuelo, tabla de tramos desglosados
# (RUTA con nombre de ciudad · FECHA · MILLAS · TIEMPO VUELO con calzos ·
# COSTO POR HORA · TOTAL POR TRAMO), desglose canónico, TUAS solo las
# COBRADAS, cobros compactos con comisión bancaria y notas internas. Nada de
# operación (tacos, horas voladas, avión operativo, traslados) ni partición /
# gastos / utilidad / CFDI: eso vive en el reporte del vuelo.
# El API (`quotes-pdf-interno.util` → `PyservicesService.generateCotizacionInternaPdf`)
# manda TODO ya calculado (desglose canónico v1.3 del snapshot, cobrosEnUsd,
# pagoVendedorUsd, `total_usd` por tramo y el ajuste contra la línea
# TIEMPO_VUELO): aquí SOLO se pinta. TODO con default y extra="ignore"
# (aditivo, skew tolerante): los campos de la v1 siguen aceptados y se ignoran. =====
class CotizacionInternaTramoCotizadoPdf(BaseModel):
    """Tramo COTIZADO (`snapshot.tramos[]`, ruta comercial congelada): una
    fila de la tabla de administración. `tiempo_hr` es el tiempo COBRABLE del
    tramo e INCLUYE el calzo (0.15 h por aterrizaje); `total_usd` =
    round2(tiempo_hr × tarifa) lo calcula el API (único número nuevo)."""

    model_config = ConfigDict(extra="ignore")

    orden: int = 0  # 1..N
    ruta: str = ""  # "Cancun-Merida" (ciudad del catálogo → nombre → IATA)
    origen_iata: str = ""
    destino_iata: str = ""
    origen_nombre: str = ""
    destino_nombre: str = ""
    fecha: str | None = None  # YYYY-MM-DD (pared Cancún)
    millas: float | None = None  # millas náuticas del tramo
    tiempo_hr: float = 0  # horas cobrables CON calzo (4 dec.)
    tiempo_hhmm: str | None = None  # "01:18" (si falta se formatea de tiempo_hr)
    tarifa_hora_usd: float | None = None  # única del vuelo salvo tarifa por tramo
    total_usd: float = 0
    pax: int | None = None  # 0 en ferry
    es_ferry: bool = False
    pernocta: bool = False
    pernocta_usd: float = 0
    tuas_usd: float = 0  # prorrateo del motor (informativo)
    consolidado: bool = False  # fila ÚNICA de respaldo (snapshot sin tramos)


class CotizacionInternaTuaCobradaPdf(BaseModel):
    """TUA COBRADA (`snapshot.tuas.filas`: el motor ya excluye exentas/$0).
    `total_usd` es igual a la línea TUAS del desglose canónico."""

    model_config = ConfigDict(extra="ignore")

    iata: str = ""
    pax: int = 0
    unitario: float = 0  # por pax en moneda NATIVA
    moneda: str = "USD"
    total_nativo: float = 0
    tc_aplicado: float | None = None  # None en USD
    total_usd: float = 0


class CotizacionInternaLineaPdf(CotizacionGrupoLineaPdf):
    """Línea del desglose canónico v1.3 TAL CUAL (Σ monto_usd == total_usd).
    Hereda clave/concepto/monto_usd/cantidad/unitario/moneda; agrega el
    nativo MXN y el TC congelado. `exento` es LEGADO v1 (línea sintética de
    TUA exento): la v2 no la pinta — solo viajan TUAS cobradas."""

    model_config = ConfigDict(extra="ignore")

    monto_nativo: float | None = None
    tc_aplicado: float | None = None
    aplica_iva: bool | None = None
    exento: bool = False


class CotizacionInternaCobroPdf(BaseModel):
    """Un cobro del vuelo (partes de sobre de grupo incluidas). `monto` es el
    BRUTO que pagó el cliente (negativo = reembolso); `neto` = monto −
    comisión bancaria; `monto_usd` viene de cobrosEnUsd (None = MXN sin TC)."""

    model_config = ConfigDict(extra="ignore")

    fecha: str | None = None
    metodo: str = ""
    metodo_label: str | None = None
    monto: float = 0
    moneda: str = "USD"
    tc: float | None = None
    monto_usd: float | None = None
    comision_pct: float | None = None
    comision_monto: float | None = None
    neto: float | None = None
    referencia: str | None = None
    cuenta_destino: str | None = None
    notas: str | None = None
    conciliado: bool | None = None
    es_reembolso: bool = False
    sobre_grupo_folio: str | None = None
    sobre_grupo_monto_total: float | None = None
    sobre_grupo_moneda: str | None = None
    grupo_factor: float | None = None
    registrado_por: str | None = None


class CotizacionInternaPdfRequest(BaseModel):
    """Payload de `POST /reportes/cotizacion-interna` (contrato
    `CotizacionInternaPdfRequest` de pyservices.service.ts, v2 del 8-sep).
    Nada se calcula aquí: las operaciones («2.60 h × $900.00», «8.857 % =
    $3,236.36 → neto») solo se PINTAN con los números que manda el API."""

    model_config = ConfigDict(extra="ignore")

    # (1) Cabecera
    folio: str = ""
    version: int | None = None
    estado: str = ""
    estado_label: str | None = None
    tipo: str | None = None
    cliente: str = "Cliente"
    razon_social: str | None = None
    cliente_rfc: str | None = None
    es_broker: bool = False
    fecha_vuelo: str | None = None  # YYYY-MM-DD (pared Cancún) — FECHA PROTAGONISTA
    fecha_vuelo_fin: str | None = None  # YYYY-MM-DD solo si multi-día y difiere
    fecha: str | None = None  # fecha de cotización (ISO) — se pinta en pequeño
    fecha_confirmacion: str | None = None
    tarifa_tipo: str | None = None  # PUBLICO | BROKER
    tarifa_tipo_label: str | None = None
    tarifa_hora_usd: float | None = None
    tarifa_override: bool = False
    tarifa_preferencial: bool = False
    metodo_cobro: str | None = None
    metodo_cobro_detalle: str | None = None
    metodo_cobro_label: str | None = None
    tc_usd_mxn: float | None = None
    vendedor: str | None = None
    cotizado_por: str | None = None
    aeronave_cotizada_modelo: str | None = None
    aeronave_cotizada_matricula: str | None = None  # SIEMPRE visible aquí
    avion_externo: str | None = None
    operador_externo: str | None = None
    piloto: str | None = None
    copiloto: str | None = None
    pasajeros: int = 0
    ruta: str | None = None
    itinerario_operativo: bool = False
    cotizacion_abierta: bool = False
    es_interno: bool = False
    es_externo: bool = False
    grupo_folio: str | None = None
    grupo_posicion: int | None = None
    grupo_total_aviones: int | None = None
    combinado_con_folio: str | None = None
    # (2) Tramos cotizados y horas (snapshot.tramos / snapshot.tiempos)
    tramos_cotizados: list[CotizacionInternaTramoCotizadoPdf] = Field(default_factory=list)
    tramos_tiempo_total_hr: float = 0  # Σ tiempo_hr
    tramos_tiempo_total_hhmm: str | None = None  # "02:36"
    tramos_total_usd: float = 0  # Σ total_usd (fila TOTAL)
    # Línea TIEMPO_VUELO canónica − Σ tramos (0 si cuadra): Σ tramos + ajuste
    # == «Servicio aéreo». Se pinta como fila aparte, nunca se reparte.
    tramos_ajuste_usd: float = 0
    tramos_ajuste_motivo: str | None = None  # "Hora mínima 1.0 h" · "Sobrevuelo 0.5 h" · …
    horas_cotizadas_hr: float | None = None
    vuelo_hr: float | None = None
    calzos_hr: float | None = None
    sobrevuelo_hr: float | None = None
    tiempo_cobrable_hr: float | None = None
    hora_minima_aplicada: bool = False
    cobrable_override: bool = False
    # (3) Desglose interno (líneas canónicas + escalares espejo)
    lineas: list[CotizacionInternaLineaPdf] = Field(default_factory=list)
    tuas_cobradas: list[CotizacionInternaTuaCobradaPdf] = Field(default_factory=list)
    subtotal_vuelo_usd: float = 0
    tuas_usd: float = 0
    extras_total_usd: float = 0
    viaticos_pernocta_usd: float = 0
    comision_vendedor_usd: float = 0
    comision_vendedor_nombre: str | None = None
    comision_vendedor_modo: str | None = None  # FIJA | POR_HORA
    comision_vendedor_tarifa_hr: float | None = None
    iva_comision_vendedor_usd: float = 0
    pago_vendedor_usd: float | None = None
    ajuste_final_usd: float = 0
    descuento_usd: float | None = None
    redondeo_auto_usd: float | None = None
    total_pactado_usd: float | None = None
    comision_billpocket_pct: float | None = None
    subtotal_usd: float | None = None  # total − IVA
    iva_pct: float = 0  # 16, no 0.16
    iva_base_usd: float | None = None
    iva_usd: float = 0
    iva_nota: str | None = None
    total_usd: float = 0
    total_mxn: float | None = None
    mxn_nativos: float | None = None
    version_motor: str | None = None
    calculado_at: str | None = None
    # (4) Cobros
    cobros: list[CotizacionInternaCobroPdf] = Field(default_factory=list)
    total_cobrado_usd: float = 0
    cobros_sin_tc_count: int = 0
    cobros_sin_tc_mxn: float = 0
    comision_banco_usd: float = 0
    total_cobrado_neto_usd: float | None = None
    saldo_usd: float | None = None
    cobrado_flag: bool = False
    semaforo_cobro: str | None = None  # verde | amarillo | rojo | gris
    semaforo_cobro_key: str | None = None
    semaforo_cobro_label: str | None = None
    # (5) Notas
    notas_cliente: str | None = None  # no se pinta en la v2 (solo internas)
    notas_internas: str | None = None
    # (6) Pie
    generado: str | None = None  # ISO instante
    generado_cancun: str | None = None  # "YYYY-MM-DD HH:mm" ya en Cancún
    generado_por: str | None = None
    # ---- LEGADO v1 (API anterior al 8-sep): aceptados y NO pintados. Tipos
    # laxos a propósito: solo sirven para que un payload viejo no truene. ----
    tramos: list[dict[str, Any]] = Field(default_factory=list)
    horas_voladas_hr: float | None = None
    delta_horas_hr: float | None = None
    aeronave_operativa: str | None = None
    apoyos: list[str] = Field(default_factory=list)
    fecha_traslado_inicial: str | None = None
    fecha_traslado_final: str | None = None
    tuas_exentos: list[str] = Field(default_factory=list)
    neto_vuelatour_usd: float | None = None
    venta_avion_usd: float | None = None
    otros_ingresos_vuelatour_usd: float | None = None
    iva_avion_usd: float | None = None
    iva_vuelatour_usd: float | None = None
    particion_fuente: str | None = None
    particion_inconsistente: bool = False
    participacion_aviones: list[dict[str, Any]] = Field(default_factory=list)
    gastos_por_categoria: list[dict[str, Any]] = Field(default_factory=list)
    gastos_total_usd: float | None = None
    gastos_sin_tc_count: int = 0
    gastos_sin_tc_mxn: float = 0
    costo_externo_usd: float | None = None
    utilidad_bruta_usd: float | None = None
    utilidad_base: str | None = None
    facturado: bool = False
    facturas: list[dict[str, Any]] = Field(default_factory=list)
    cfdi_estatus: str | None = None
    cfdi_folio: str | None = None


# ===== Balance por avión (réplica sistematizada del Excel "Balance N990GG").
# El API (NestJS) calcula TODO el dinero; aquí SOLO se pinta el libro. Por eso
# todos los numéricos son opcionales: None = celda vacía (nunca un 0 falso) y
# las listas tienen default (skew tolerante entre deploys). =====
class BalanceAvionCobro(BaseModel):
    """Parcialidad de cobro del vuelo (fecha + monto en MXN)."""

    fecha: str | None = None
    monto_mxn: float | None = None
    # Método del cobro (EFECTIVO/TRANSFERENCIA/...); el API ya lo mandaba.
    metodo: str | None = None
    # Regla 28-ago-2026: comisión que retuvo el banco (MXN) y cuenta que
    # recibió el depósito (Paywise · HSBC Dólares · HSBC Pesos · Scotiabank
    # Dólares · Scotiabank Pesos). Opcionales: un API viejo no los manda.
    comision_mxn: float | None = None
    cuenta: str | None = None


class BalanceAvionVuelo(BaseModel):
    """Una fila de la hoja maestra 'reporte horas <MATRÍCULA>' (1 vuelo)."""

    # Llave interna de orden cronológico (el API ya manda las filas
    # ordenadas); NO se pinta.
    orden_ts: str | None = None
    clave: str | None = None  # folio + clave del cliente (estilo "vt...")
    fecha: str | None = None  # ISO date o texto libre (multi-día: "9-10 sep")
    ruta: str | None = None
    estado: str | None = None  # COMPLETADO / CANCELADO / ... (se pinta tal cual)
    # Balance GENERAL (flota): la fila se tiñe con el color del avión
    # (aeronave.color_calendario — así identifica el equipo en su libro).
    avion_color: str | None = None
    # Vuelo EXTERNO (operador ajeno): en el general aparece como un vuelo
    # más (fila gris, sin color de avión); su costo del operador va en
    # OPERACIONES. None/False = vuelo propio.
    es_externo: bool | None = None
    operador_externo: str | None = None
    # --- Bloque VENTA ---
    horas_cobradas: float | None = None
    tarifa_usd: float | None = None
    iva_hr_usd: float | None = None
    # VENTA DEL AVIÓN (regla 28-ago-2026): tiempo de vuelo (tarifa × horas
    # cobradas) + ajuste/descuento + su IVA proporcional. Los TUAs, extras,
    # viáticos de pernocta y la comisión del vendedor cobrados NO están
    # aquí: son ingreso de VuelaTour (pestaña "otros movimientos" del
    # Balance general). En vuelos MULTI-AVIÓN (regla B) es la parte
    # proporcional de esta matrícula (ver `participacion`).
    total_usd: float | None = None
    iva_usd: float | None = None
    tc_venta: float | None = None
    # true = la cotización no traía TC: se usó el oficial de REFERENCIA del
    # día de la cotización (TipoCambioService.oficialDetallePara) — el Excel
    # pinta las celdas en azul claro y la nota dice fuente y fecha del dato.
    tc_venta_oficial: bool = False
    # Fuente del oficial: 'OPEN_ER_API' (diario) · 'ECB_FRANKFURTER' (fechas
    # pasadas) · 'BANXICO_FIX' (legado). None = API viejo (sin detalle).
    tc_venta_oficial_fuente: str | None = None
    # Fecha (YYYY-MM-DD) del dato del TC usado.
    tc_venta_oficial_fecha: str | None = None
    total_mxn: float | None = None
    iva_mxn: float | None = None
    subtotal_mxn: float | None = None
    # --- Bloque TIEMPO / TACÓMETRO ---
    tiempo_vuelo: float | None = None
    taco_inicio: float | None = None
    taco_fin: float | None = None
    # Salto en la cadena: el taco inicial no empalma con el final de la fila
    # anterior del avión (se pinta en ámbar, igual que el panel).
    salto_taco_inicio: bool = False
    salto_taco_esperado: float | None = None
    # Salto INTERNO: un tramo del vuelo no empalma con el anterior (infla las
    # horas sin romper la cadena entre vuelos) — se pinta en TIEMPO VUELO.
    salto_taco_interno: bool = False
    salto_taco_interno_detalle: str | None = None
    # Observaciones del EQUIPO sobre las lecturas de taco de los tramos del
    # vuelo (capturadas en Tacómetros en vivo): líneas ya formateadas
    # "CUN→PTU salida: <texto> — Nombre, dd-mmm". Celda en ámbar + nota.
    taco_inicio_obs: list[str] = Field(default_factory=list)
    taco_fin_obs: list[str] = Field(default_factory=list)
    # --- Bloque COSTOS DIRECTOS (MXN) ---
    gas_mxn: float | None = None
    gas_litros: float | None = None
    gas_precio_litro: float | None = None
    op_mxn: float | None = None
    piloto_mxn: float | None = None
    otros_mxn: float | None = None
    # Desglose por celda (nota de Excel): una línea por gasto, p. ej.
    # "Comida · Starbucks — $206.00". Vacío = sin nota en la celda.
    gas_detalle: list[str] = Field(default_factory=list)
    op_detalle: list[str] = Field(default_factory=list)
    piloto_detalle: list[str] = Field(default_factory=list)
    otros_detalle: list[str] = Field(default_factory=list)
    permiso_afac_mxn: float | None = None
    costo_total_mxn: float | None = None
    tc_costos: float | None = None
    # --- Bloque INDICADORES USD e IVA ---
    costo_usd: float | None = None
    costo_usd_siva: float | None = None
    iva_pagado_usd: float | None = None
    iva_pagado_mxn: float | None = None
    remanente_mxn: float | None = None
    dif_iva_mxn: float | None = None
    # Regla A (28-ago-2026 tarde): la comisión del vendedor ya NO es venta
    # ni costo del avión (es ingreso y pago de VuelaTour: 'otros
    # movimientos'). El API la manda None; la columna se conserva por layout.
    comision_vendedor_mxn: float | None = None
    ganancia_mxn: float | None = None
    ganancia_usd: float | None = None
    costo_hr_usd: float | None = None
    costo_hr_usd_siva: float | None = None
    # --- Bloque STATUS DE COBROS ---
    status_cobro: str | None = None  # Cobrado / Parcial / Pendiente / —
    cobros: list[BalanceAvionCobro] = Field(default_factory=list)
    # cobrado_mxn / por_cobrar_mxn vienen YA prorrateados al avión (regla
    # 28-ago): parcialidades reales × (venta avión ÷ total cotización).
    cobrado_mxn: float | None = None
    por_cobrar_mxn: float | None = None
    por_cobrar_usd: float | None = None
    # --- Regla 28-ago-2026 (todo opcional: tolera API viejo) ---
    # Total COMPLETO cobrado al cliente (con TUAs/extras/pernocta + su IVA).
    total_cotizacion_usd: float | None = None
    total_cotizacion_mxn: float | None = None
    # Depósitos reales del vuelo (sin prorratear).
    cobrado_real_mxn: float | None = None
    # venta avión ÷ total cotización (factor del prorrateo de los cobros).
    venta_factor: float | None = None
    # Parte VuelaTour del vuelo (TUAs/extras/pernocta/comisión del vendedor
    # + su IVA).
    otros_ingresos_usd: float | None = None
    # TUA pagado del vuelo: SOLO informativo (nota en OPERACIONES); no resta.
    tua_pagado_mxn: float | None = None
    # --- Regla B (28-ago-2026): vuelo MULTI-AVIÓN (todo opcional) ---
    # Fracción de la venta del avión que toca a esta matrícula (Σ == 1 entre
    # los aviones del vuelo; 1 o None = vuelo de un solo avión). La RUTA ya
    # trae el texto "· 50 % …" armado por el API; aquí solo se anota.
    participacion: float | None = None
    multi_avion: bool | None = None
    # Base del reparto: 'tramos' (partes iguales por tramo vendido; los
    # ferries/tramos operativos no reparten) · 'unico' (un solo avión). El
    # API no emite otra fuente (nunca por horas).
    participacion_fuente: str | None = None


class BalanceAvionTotales(BaseModel):
    """Fila TOTALES de la hoja maestra: sumas y promedios YA calculados."""

    horas_cobradas: float | None = None
    tiempo_vuelo: float | None = None
    total_mxn: float | None = None
    iva_mxn: float | None = None
    subtotal_mxn: float | None = None
    gas_mxn: float | None = None
    gas_litros: float | None = None
    op_mxn: float | None = None
    piloto_mxn: float | None = None
    otros_mxn: float | None = None
    permiso_afac_mxn: float | None = None
    costo_total_mxn: float | None = None
    remanente_mxn: float | None = None
    dif_iva_mxn: float | None = None
    # None/0 desde el 28-ago-2026 tarde (regla A): ya no es costo del avión.
    comision_vendedor_mxn: float | None = None
    ganancia_mxn: float | None = None
    ganancia_usd: float | None = None
    cobrado_mxn: float | None = None
    por_cobrar_mxn: float | None = None
    por_cobrar_usd: float | None = None
    tc_promedio: float | None = None  # AVERAGE de los TC de costos no nulos
    costo_hr_prom_usd: float | None = None  # AVERAGE de costo/hr USD no nulos
    # TUAs/extras/pernocta/comisión del vendedor cotizados (+ su IVA)
    # EXCLUIDOS de las filas: ingreso de VuelaTour (pestaña "otros
    # movimientos" del Balance general).
    otros_ingresos_usd: float | None = None
    # --- Regla 28-ago-2026 (opcionales) ---
    cobrado_real_mxn: float | None = None  # depósitos reales del periodo
    total_cotizacion_mxn: float | None = None  # total completo (c/extras)
    tua_pagado_mxn: float | None = None  # solo nota; no resta en el libro
    comision_banco_mxn: float | None = None  # Σ comisiones retenidas por bancos


class BalanceAvionGastoFila(BaseModel):
    """Partida del ledger de gastos (indirectos / refacciones / otros /
    permisos)."""

    fecha: str | None = None
    # Categoría del gasto (etiqueta amable es-MX, 29-ago) — columna
    # CATEGORÍA entre FECHA y DETALLE. None = API viejo (celda vacía).
    categoria: str | None = None
    detalle: str | None = None
    monto_mxn: float | None = None
    moneda_original: str | None = None  # solo si ≠ MXN
    monto_original: float | None = None
    # Litros de la carga (solo hoja "combustible").
    litros: float | None = None
    # Balance GENERAL: fila teñida con el color del avión.
    avion_color: str | None = None
    # Matrícula de la fila (hoja "combustible" del GENERAL, 28-ago): la hoja
    # se agrupa en SECCIONES por matrícula con subtotal. None = libro
    # individual (una sola matrícula: req.matricula).
    matricula: str | None = None
    # Solo hoja "refacciones" del GENERAL (29-ago): costo FIFO de la salida
    # de inventario y VENTA al avión (= monto del gasto). Cuando vienen, la
    # hoja agrega columnas COSTO VUELATOUR / VENTA AL AVIÓN / GANANCIA
    # (ganancia = venta − costo, solo para mostrar; 0 mientras la salida se
    # cargue a costo). None = sin columnas extra.
    costo_mxn: float | None = None
    venta_mxn: float | None = None


class BalanceAvionHojaGastos(BaseModel):
    """Hoja de gastos: ledger + resumen (USD al TC promedio del periodo)."""

    filas: list[BalanceAvionGastoFila] = Field(default_factory=list)
    total_mxn: float | None = None
    usd: float | None = None
    usd_hr: float | None = None


class BalanceAvionHojaCombustible(BalanceAvionHojaGastos):
    """Hoja 'combustible' (26-ago-2026): el gas del avión POR MES, con
    litros y precio por litro promedio — ya no va por vuelo."""

    litros_total: float | None = None
    precio_litro_prom: float | None = None


class BalanceAvionSocio(BaseModel):
    nombre: str = ""
    porcentaje: float | None = None  # % vigente (0–100)
    monto_usd: float | None = None  # utilidad cobrada × pct/100


class BalanceAvionBalanceBloque(BaseModel):
    """Hoja 'balance': bloque A–I del original (todo USD) + socios."""

    utilidad_antes_usd: float | None = None
    # "Gasto de combustible" del mes (hoja combustible al TC promedio).
    combustible_usd: float | None = None
    gastos_indirectos_usd: float | None = None
    # Hoja "refacciones" (29-ago): salidas de inventario, ANTES dentro de
    # gastos indirectos (indirectos + refacciones == lo de antes). None =
    # API viejo (celda vacía; la utilidad ya las traía dentro de indirectos).
    refacciones_usd: float | None = None
    otros_usd: float | None = None
    permisos_usd: float | None = None
    utilidad_despues_usd: float | None = None
    por_cobrar_usd: float | None = None
    utilidad_cobrada_usd: float | None = None
    socios: list[BalanceAvionSocio] = Field(default_factory=list)


class BalanceOtroMovimientoFila(BaseModel):
    """Fila de la pestaña "Otros movimientos" (28-ago, hoja manual del
    cliente): egreso pagado apareado con el ingreso cobrado por concepto y
    su remanente. Cualquier lado puede venir vacío (solo-ingreso o
    solo-egreso); viene YA calculada del API — aquí jamás se recalcula.
    Desde el 28-ago-2026 tarde (regla A) incluye la comisión del vendedor
    como INGRESO de VuelaTour y su pago al vendedor como EGRESO apareado
    (PROVISIÓN a la fecha del vuelo mientras no exista gasto real; la nota
    de la celda lo dice)."""

    clave: str = ""
    avion_color: str | None = None
    # Estado del vuelo (COMPLETADO / CANCELADO / RESERVA…): el API lista
    # TODOS los estados del periodo (igual que la hoja maestra) y el Excel
    # marca los que no son normales. None = API viejo (no se marca).
    estado: str | None = None
    fecha_vuelo: str | None = None
    concepto_egreso: str | None = None
    egreso_mxn: float | None = None
    fecha_egreso: str | None = None
    concepto_ingreso: str | None = None
    ingreso_mxn: float | None = None
    fecha_ingreso: str | None = None
    remanente_mxn: float | None = None
    factura: str | None = None
    # Desglose línea por línea (una fila por vuelo, 28-ago): se pinta como
    # COMENTARIO de la celda de ingreso / egreso. None = API viejo.
    nota_ingreso: str | None = None
    nota_egreso: str | None = None


class BalanceHojaOtrosMovimientos(BaseModel):
    filas: list[BalanceOtroMovimientoFila] = Field(default_factory=list)
    # Movimientos SIN avión y SIN vuelo (dinero de empresa).
    filas_sueltas: list[BalanceOtroMovimientoFila] = Field(default_factory=list)


class BalanceAvionRequest(BaseModel):
    generado: str | None = None
    matricula: str = ""
    modelo: str | None = None
    # Color del avión (aeronave.color_calendario) — lo usa el balance
    # GENERAL para teñir su bloque en la hoja "balance".
    avion_color: str | None = None
    periodo_desde: str | None = None
    periodo_hasta: str | None = None
    vuelos: list[BalanceAvionVuelo] = Field(default_factory=list)
    totales: BalanceAvionTotales = Field(default_factory=BalanceAvionTotales)
    gastos_indirectos: BalanceAvionHojaGastos = Field(default_factory=BalanceAvionHojaGastos)
    # Hoja "refacciones" (29-ago): salidas de inventario (gasto REFACCION
    # medio BODEGA ligado al cardex), separadas de "gastos indirectos".
    # None = API viejo → no se pinta la hoja (skew tolerante).
    refacciones: BalanceAvionHojaGastos | None = None
    otros_gastos: BalanceAvionHojaGastos = Field(default_factory=BalanceAvionHojaGastos)
    permisos: BalanceAvionHojaGastos = Field(default_factory=BalanceAvionHojaGastos)
    # Pestaña mensual de combustible (26-ago-2026); default vacío = skew
    # tolerante con un API viejo que aún no la manda.
    combustible: BalanceAvionHojaCombustible = Field(
        default_factory=BalanceAvionHojaCombustible
    )
    balance: BalanceAvionBalanceBloque = Field(default_factory=BalanceAvionBalanceBloque)
    # Pestaña "Otros movimientos" (28-ago): solo la manda el GENERAL; None =
    # no se pinta (skew tolerante con API viejo).
    otros_movimientos: BalanceHojaOtrosMovimientos | None = None
    pendientes: list[str] = Field(default_factory=list)


class BalanceGeneralResumenFila(BaseModel):
    """Fila del RESUMEN del balance general (viene YA calculada del API —
    aquí jamás se recalcula dinero)."""

    matricula: str = ""
    # Color del avión (aeronave.color_calendario): la celda de la matrícula
    # se tiñe con él — es la LEYENDA de colores del libro general.
    color: str | None = None
    vuelos: int = 0
    horas: float | None = None
    horas_cobradas: float | None = None
    venta_mxn: float | None = None
    costo_mxn: float | None = None
    # "Gasto de combustible" del mes.
    combustible_mxn: float | None = None
    # Comisiones de vendedor: desde el 28-ago-2026 tarde (regla A) llegan
    # None/0 — la comisión es ingreso y pago de VuelaTour, no del avión.
    # La columna se conserva por layout.
    comisiones_mxn: float | None = None
    # VENTA − COSTO TOTAL − COMBUSTIBLE = GANANCIA (leyenda impresa).
    ganancia_mxn: float | None = None
    cobrado_mxn: float | None = None
    por_cobrar_mxn: float | None = None
    pendientes: int = 0


class BalanceInventarioItemFila(BaseModel):
    """Fila del bloque POR ÍTEM de la hoja 'inventario' del Balance GENERAL
    (tiendita, 30-ago-2026). EXISTENCIA y VALOR A COSTO son A HOY (todo el
    cardex FIFO), no una foto al corte; el resto es del periodo. None =
    celda vacía (sin actividad de ese tipo), nunca un 0 falso."""

    nombre: str = ""  # nombre del ítem (+ ' · nº de parte' cuando lo tiene)
    existencia: float | None = None
    valor_costo_mxn: float | None = None
    # ENTRADAs del periodo (compras reales; DEVOLUCION/AJUSTE no son compra).
    compradas_cant: float | None = None
    compradas_costo_mxn: float | None = None
    salidas_cant: float | None = None
    # Σ venta de las salidas CON precio (lo cargado a los aviones) y su
    # utilidad (venta − costo FIFO consumido).
    vendido_mxn: float | None = None
    utilidad_mxn: float | None = None
    # Matrículas a las que se aplicó en el periodo (únicas, ' + ').
    matriculas: str | None = None


class BalanceHojaInventario(BaseModel):
    """Hoja 'inventario' del Balance GENERAL (tiendita): bloque por ítem +
    totales — el bloque 2 (detalle de salidas) sale de `refacciones`."""

    filas: list[BalanceInventarioItemFila] = Field(default_factory=list)
    total_piezas: float | None = None
    total_valor_mxn: float | None = None
    total_compras_mxn: float | None = None
    total_vendido_mxn: float | None = None
    total_utilidad_mxn: float | None = None


class BalanceGeneralRequest(BaseModel):
    """Balance GENERAL de flota: los libros individuales de varios aviones
    (misma estructura de hojas) concatenados en un workbook + hoja RESUMEN."""

    generado: str | None = None
    periodo_desde: str | None = None
    periodo_hasta: str | None = None
    resumen: list[BalanceGeneralResumenFila] = Field(default_factory=list)
    resumen_totales: BalanceGeneralResumenFila | None = None
    # CONSOLIDADO (regla del cliente, 18-ago): UN solo juego de hojas con los
    # datos de TODOS los aviones juntos (filas teñidas con el color de cada
    # avión). `aviones` se usa solo para los bloques de la hoja "balance"
    # (los socios son POR avión).
    consolidado: BalanceAvionRequest | None = None
    aviones: list[BalanceAvionRequest] = Field(default_factory=list)
    # Hoja 'otros gastos' del GENERAL (29-ago; 1-sep-2026: antes se llamaba
    # "gastos VuelaTour" — solo cambió el NOMBRE de la hoja, este campo del
    # contrato NO): gastos de EMPRESA (sin vuelo ni avión; sin PERSONAL_DUENO
    # ni GAS) sin reparto + remanentes de reparto manual — egresos de
    # VuelaTour, fuera de toda cascada por avión. Antes salían como filas
    # sueltas de "Otros movimientos". None = API viejo.
    gastos_empresa: BalanceAvionHojaGastos | None = None
    # Hoja "inventario" (tiendita, 30-ago-2026): resumen POR ÍTEM del
    # periodo. Cuando viene, SUSTITUYE a la hoja 'refacciones' en el render
    # del general (el detalle de salidas de `consolidado.refacciones` pasa a
    # ser su bloque 2); el libro INDIVIDUAL conserva la suya. None = API
    # viejo → se pinta 'refacciones' como antes (skew tolerante).
    inventario: BalanceHojaInventario | None = None


class BitacoraTacoFila(BaseModel):
    """LEGADO (payload plano, anterior a las tiras por componente).

    Una fila = UN vuelo: fecha, tacómetro inicial→final, horas y ruta. Los
    tiempos de hélice solo venían en el formato bimotor (MOTOR_HELICE). Se
    conserva para que un API viejo siga imprimiendo durante el deploy
    (pyservices sale ANTES que el API); el payload nuevo usa ``BitacoraTira``.
    """

    fecha: str  # ISO (date o datetime); se formatea dd-mmm en hora Cancún
    taco_inicial: float
    horas: float
    taco_final: float
    ruta: str  # "cun-pps-cun" (minúsculas, guiones)
    helice_inicial: float | None = None
    helice_final: float | None = None


class BitacoraTiraFila(BaseModel):
    """Una fila = UN vuelo dentro de una tira (planeador, motor o hélice).

    Mismo renglón que el equipo llenaba a mano en la plantilla Excel para
    recortar y pegar en cada libro. ``tiempo_inicial``/``tiempo_final`` son el
    acumulado del componente al inicio y al final del vuelo, YA derivados por
    el API desde el tacómetro y la base capturada en la ficha; ``None`` se
    pinta "—" para llenarlo a mano.
    """

    fecha: str  # ISO (date o datetime); se formatea dd-mmm en hora Cancún
    taco_inicial: float
    horas: float
    taco_final: float
    tiempo_inicial: float | None = None
    tiempo_final: float | None = None
    ruta: str  # "cun-pps-cun" (minúsculas, guiones)


class BitacoraTira(BaseModel):
    """Una bitácora imprimible (una página): planeador, motor o hélice.

    ``con_tiempo`` decide las columnas: True ⇒ 7 (tacómetro + tiempo del
    componente inicial/final); False ⇒ 5 (solo tacómetro, la tira histórica).
    ``etiqueta`` encabeza las columnas de tiempo ("Tiempo planeador", ...).
    """

    tipo: str  # PLANEADOR | MOTOR | HELICE
    titulo: str
    etiqueta: str
    nota: str | None = None
    con_tiempo: bool = True
    filas: list[BitacoraTiraFila] = Field(default_factory=list)


class BitacoraTacoRequest(BaseModel):
    """Bitácoras de vuelo por componente, una tira por página.

    Payload nuevo: ``tiras`` (planeador / motor / hélice), cada una con sus
    filas ya derivadas por el API. Payload LEGADO (skew de deploy): ``formato``
    + ``filas`` planas; el servicio lo convierte a UNA tira equivalente
    (``_tiras_normalizadas``): MOTOR_HELICE ⇒ 7 columnas con los tiempos de
    hélice, PLANEADOR ⇒ 5 columnas (solo tacómetro). Si vienen ambas, mandan
    las ``tiras``.
    """

    matricula: str = ""
    modelo: str | None = None
    formato: str = "PLANEADOR"  # LEGADO: PLANEADOR | MOTOR_HELICE
    desde: str | None = None
    hasta: str | None = None
    generado: str | None = None
    filas: list[BitacoraTacoFila] = Field(default_factory=list)  # LEGADO
    tiras: list[BitacoraTira] = Field(default_factory=list)


class DineroCobroPago(BaseModel):
    """Parcialidad de cobro (o pago a proveedor) de la hoja dinero-vlos."""

    fecha: str | None = None
    monto_mxn: float | None = None


class DineroVueloFila(BaseModel):
    """Una fila de la hoja 'dinero-vlos' (un vuelo del periodo).

    Las columnas SIN regla todavía (costo proveedor, pagos) viajan en None
    y se pintan vacías conservando su columna del libro. Las de COMISIÓN
    también van vacías, pero por regla (28-ago-2026 tarde): la comisión del
    vendedor es ingreso y pago de VuelaTour y vive en la hoja 'Otros
    ingresos', no en la fila del avión.
    """

    clave: str = ""
    matricula: str | None = None  # nota de la celda CLAVE
    color: str | None = None  # color de fila (hex del avión, #RRGGBB)
    fecha: str | None = None
    ruta: str = ""
    tiempo: float | None = None  # horas cobradas (calzos/hobs)
    venta_hr_usd: float | None = None
    venta_hr_mxn: float | None = None
    iva_hr_usd: float | None = None
    venta_hr_masiva_usd: float | None = None
    # VENTA DEL AVIÓN (regla 28-ago-2026): tiempo + ajuste + IVA proporcional
    # — sin TUAs/extras/pernocta/comisión del vendedor (esos van en la hoja
    # "Otros ingresos"). Multi-avión: parte proporcional de esta matrícula.
    total_cobrado_usd: float | None = None
    iva_total_usd: float | None = None
    tc_venta: float | None = None
    total_cobrado_mxn: float | None = None
    iva_total_mxn: float | None = None
    total_siva_mxn: float | None = None
    status_cobro: str | None = None  # COBRADO | PENDIENTE
    cobros: list[DineroCobroPago] = Field(default_factory=list)
    total_cobros_mxn: float | None = None
    me_deben_mxn: float | None = None
    factura_vuelatour: str | None = None
    # Total COMPLETO facturado al cliente (con TUAs/extras/pernocta + su IVA):
    # columna informativa junto al STATUS DE COBROS (regla 28-ago-2026).
    total_cliente_usd: float | None = None
    total_cliente_mxn: float | None = None
    # Regla B (28-ago-2026): vuelo MULTI-AVIÓN — fracción de la venta del
    # avión que toca a esta matrícula (la CLAVE ya trae el sufijo del API).
    participacion: float | None = None
    multi_avion: bool | None = None


class DineroOtroIngresoFila(BaseModel):
    """Fila de la hoja 'Otros ingresos' (TUAs/extras/pernocta y comisión del
    vendedor por vuelo — ingreso de VuelaTour, no del avión; desde 28-ago
    también líneas 'iva de tuas/extras' con el mismo shape y el pago de la
    comisión al vendedor como egreso apareado en PROVISIÓN)."""

    clave: str = ""
    fecha_vuelo: str | None = None
    concepto_egreso: str | None = None
    egreso_mxn: float | None = None
    fecha_egreso: str | None = None
    # Nota de la celda del egreso (p. ej. "PROVISIÓN: pago al vendedor =
    # comisión + IVA…"); se pinta como comentario de Excel.
    nota_egreso: str | None = None
    concepto_ingreso: str | None = None
    ingreso_mxn: float | None = None
    fecha_ingreso: str | None = None
    remanente_mxn: float | None = None
    factura: str | None = None


class DineroOtroGastoFila(BaseModel):
    """Fila de la hoja 'otros gastos' (gastos del mes sin vuelo)."""

    fecha: str | None = None
    concepto: str = ""
    monto_mxn: float | None = None
    acumulado_mxn: float | None = None


class DineroUtilidadAvion(BaseModel):
    """Columna por avión de la hoja 'utilidades' (lo computable hoy)."""

    matricula: str = ""
    gastos_indirectos_mxn: float | None = None
    otros_gastos_mxn: float | None = None
    permisos_mxn: float | None = None
    # "Gasto de combustible" del mes del avión (pestaña Combustible).
    combustible_mxn: float | None = None


class DineroCombustibleFila(BaseModel):
    """Fila de la pestaña 'Combustible' (26-ago-2026): el gas del mes por
    avión, con o sin vuelo — ya no se persigue la asignación por vuelo."""

    fecha: str | None = None
    matricula: str = "—"  # '—' = carga sin avión (pendiente de asignar)
    avion_color: str | None = None
    concepto: str = ""
    litros: float | None = None
    monto_mxn: float | None = None
    acumulado_mxn: float | None = None


class DineroXlsxRequest(BaseModel):
    """Libro 'Dinero <periodo>' (réplica del control manual del equipo)."""

    periodo_desde: str | None = None
    periodo_hasta: str | None = None
    generado: str | None = None
    # Leyenda de colores por avión (matrícula → hex).
    leyenda_colores: list[dict] = Field(default_factory=list)
    vuelos: list[DineroVueloFila] = Field(default_factory=list)
    otros_ingresos: list[DineroOtroIngresoFila] = Field(default_factory=list)
    otros_gastos: list[DineroOtroGastoFila] = Field(default_factory=list)
    # Pestaña "Combustible" (26-ago-2026); defaults = skew tolerante.
    combustible: list[DineroCombustibleFila] = Field(default_factory=list)
    combustible_total_mxn: float | None = None
    combustible_litros: float | None = None
    combustible_precio_litro: float | None = None
    combustible_sin_avion: int = 0
    # "Gasto de combustible" del mes: resta en la hoja utilidades.
    utilidades_combustible_mxn: float | None = None
    utilidades_otros_ingresos_mxn: float | None = None
    # Provisión del pago al vendedor (comisión + IVA) ya restada de
    # utilidades_otros_ingresos_mxn (regla 28-ago); informativa para la nota.
    utilidades_comision_vendedor_provisionada_mxn: float | None = None
    utilidades_otros_gastos_mxn: float | None = None
    utilidades_tc: float | None = None
    utilidades_aviones: list[DineroUtilidadAvion] = Field(default_factory=list)
