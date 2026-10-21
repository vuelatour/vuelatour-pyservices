import math
from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field, field_validator


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
    # BASE GRAVABLE del IVA (22-sep-2026): el renglón que va SOBRE el IVA
    # cuando hay conceptos exentos (pernocta, extras con `aplica_iva=false`),
    # que bajan DEBAJO del IVA. ADITIVO: hoy el API del cliente todavía no lo
    # manda y se deriva como `total − IVA − Σ exentos` (re-suma de columna).
    iva_base_usd: float | None = None
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
    # YA NO SE PINTA (15-sep-2026, pedido del cliente: fuera «Tiempo de vuelo
    # · H:MM h por tramo» de la tarjeta "De un vistazo"). El campo se
    # conserva —el API lo sigue mandando y es ADITIVO— para no romper skew.
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
    $85.00"); `moneda` es la nativa de la línea (MXN en extras en pesos).

    `aplica_iva=False` marca un concepto que NO causa IVA: desde el
    22-sep-2026 esas líneas se pintan DEBAJO del IVA. ADITIVO (None = el API
    no lo dice → la línea se trata como gravable, igual que siempre); la
    PERNOCTA se reconoce además por su clave."""

    clave: str = ""
    concepto: str = ""
    monto_usd: float = 0
    cantidad: float | None = None
    unitario: float | None = None
    moneda: str | None = None
    aplica_iva: bool | None = None


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
    # Base gravable del IVA (22-sep-2026, mismo criterio que el PDF de un
    # avión): ADITIVO; sin él se deriva `total − IVA − Σ exentos`.
    iva_base_usd: float | None = None
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
    # Método del COBRO, ya legible ("Transferencia", "BillPocket", "Paywise"…),
    # armado por el API desde `cobro_vuelo.metodo_cobro` de ESTE cobro —
    # NUNCA el método de cobro de la cotización/vuelo (`quote.metodo_cobro`),
    # que es solo la intención pactada y puede diferir del pago real. Aquí
    # solo se pinta lo que llega en este campo.
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
    # Desde el API 0.0.75 (invariante 73, con la bandera de abajo) es el pago
    # REAL al vendedor (gastos capturados; 0 si no hay) — lo que resta la
    # ganancia — y `neto_vuelatour_usd` = total − ese pago real.
    pago_vendedor_usd: float | None = None
    # ADITIVOS del API 0.0.75 (revisión del lote 2, 14-oct-2026): con la
    # bandera, `pago_vendedor_usd` es el pago REAL y `pago_vendedor_cotizado_usd`
    # la comisión COTIZADA con su IVA (dato de la cotización, jamás se resta).
    # La sección «Cotización» pinta la cotizada y su neto; la cascada rotula
    # «(−) Pago real al vendedor» (`reporte_vuelo_vendedor.py`). Sin la
    # bandera (API ≤ 0.0.74), el reporte de siempre byte a byte.
    comision_vendedor_solo_gasto_real: bool | None = None
    pago_vendedor_cotizado_usd: float | None = None
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
    tiempo_hhmm: str | None = None  # LEGADO "01:18": ya no se pinta (API 0.0.33)
    # «TIEMPO VUELO (HRS)» (API 0.0.33): "1.19", 2 decimales fijos y repartido
    # por residuo mayor para que Σ tramos == `tramos_tiempo_total_horas`.
    # None con el total presente = el tramo no trae tiempo («—»).
    tiempo_horas: str | None = None
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
    # Avión REALMENTE utilizado (11-sep-2026; contrato nuevo del API, en
    # paralelo): segunda línea «Avión utilizado: MATRÍCULA · MODELO» de la
    # cabecera, solo cuando viene. Tipo laxo A PROPÓSITO: acepta el texto ya
    # armado ("N4142R · Piper Seneca V") o {"matricula": …, "modelo": …}; con
    # cualquier otra forma —o sin dato— la línea no se pinta y el PDF sale
    # igual (skew tolerante en ambos sentidos).
    aeronave_utilizada: Any = None
    # ⚠ La cotización se pactó con un avión y hoy vuela OTRO. Lo decide el
    # API comparando por ID (dos aviones pueden compartir modelo, así que el
    # texto no basta) y aquí solo se pinta como marca ámbar junto a la línea
    # «Avión utilizado». ADITIVO: sin el campo la línea sale sin marca.
    aeronave_cotizada_vs_utilizada_difiere: bool = False
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
    tramos_tiempo_total_hhmm: str | None = None  # LEGADO "02:36": ya no se pinta
    # Fila TOTAL de «TIEMPO VUELO (HRS)» (API 0.0.33): "2.38" = Σ tiempo_horas.
    # Ausente (API previo) ⇒ la tabla se reparte aquí con el MISMO criterio.
    tramos_tiempo_total_horas: str | None = None
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


# ===== Nombre de la regla de COMISIONES a cargo del avión (6-oct-2026, API
# 0.0.65). La vigencia es configurable en el API (`comisiones_al_avion_desde`,
# default 2026-09-01) y sus propios textos la nombran con
# `etiquetaReglaComisiones` («regla sep-2026»). Los balances y el reparto la
# reciben en `regla_comisiones` (ADITIVO); sin ella, este respaldo: el texto
# de siempre. =====
REGLA_COMISIONES_RESPALDO = "regla de septiembre 2026"


def regla_comisiones_liberal(v: Any) -> str | None:
    """`regla_comisiones` del API: texto de UNA línea o None. Blanco o algo
    que no es texto ⇒ None (los textos usan `REGLA_COMISIONES_RESPALDO`); sin
    la palabra «regla» al inicio se antepone, para que «con la …» se lea
    bien. Jamás un 422 por un texto."""
    if not isinstance(v, str):
        return None
    texto = " ".join(v.split())
    if not texto:
        return None
    return texto if texto.lower().startswith("regla") else f"regla {texto}"


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
    # 6-oct-2026 (API 0.0.60, pedido del cliente: «cómo se cobró, quién lo
    # cobró y a qué cuenta»): etiqueta legible del método (con el mismo
    # sufijo de «parte de esta fila» en multi-avión), nombre de quien
    # registró el cobro y el texto YA armado por el API
    # («Transferencia → Scotiabank Pesos · Itzi», fuente única
    # `etiquetaCobradoCon`). Solo se pintan como NOTA de la celda COBRO n
    # MXN y en la nota resumen de STATUS. Opcionales: un API ≤ 0.0.59 no
    # los manda y la nota se compone con `metodo`/`cuenta`.
    metodo_etiqueta: str | None = None
    registro: str | None = None
    cobrado_con: str | None = None
    # 6-oct-2026 (API 0.0.65, regla «comisiones a cargo del avión» desde
    # septiembre 2026): lo que ENTRÓ a la cuenta = `monto_mxn` (bruto) −
    # `comision_mxn`. El API lo manda solo en vuelos dentro de la vigencia;
    # con él la celda COBRO n MXN pinta el NETO y su nota dice bruto,
    # comisión y neto. None (API ≤ 0.0.64 o vuelo anterior) ⇒ `monto_mxn`.
    neto_mxn: float | None = None

    @field_validator("metodo_etiqueta", "registro", "cobrado_con", mode="before")
    @classmethod
    def _texto_liberal(cls, v: Any) -> Any:
        """Texto informativo: en blanco o algo que no es texto ⇒ None. Jamás
        un 422 que tumbe el balance entero por una nota."""
        if isinstance(v, str):
            return v.strip() or None
        return None


def clave_cliente_liberal(v: Any) -> str | None:
    """ID de vuelo por cliente (13-oct-2026, API 0.0.75): texto no vacío ⇒
    recortado; vacío o no-texto ⇒ None. Es metadato (solo decide el ancho de
    la columna CLAVE): jamás un 422 por él."""
    if isinstance(v, str) and v.strip():
        return v.strip()
    return None


def _monto_vinculo_liberal(v: Any) -> float | None:
    """Monto de un metadato de vínculo: número finito o texto numérico ⇒
    float; bool, NaN/infinito, texto no numérico u otra cosa ⇒ None. Jamás
    un 422 por un metadato (tumbaría el .xlsx y el libro)."""
    if isinstance(v, bool):
        return None
    if isinstance(v, int | float):
        return float(v) if math.isfinite(v) else None
    if isinstance(v, str):
        try:
            x = float(v.strip())
        except ValueError:
            return None
        return x if math.isfinite(x) else None
    return None


class BalanceVinculoGasto(BaseModel):
    """Un gasto que COMPONE una celda de costos de la fila del vuelo
    (OPERACIONES, PILOTO u OTROS) — 13-oct-2026, API 0.0.75, punto 5 del lote
    («abrir el balance en la celda del gasto» desde Conciliación). Solo se
    transporta a `libro_json.vincular_lista`; aquí no se decide nada.
    `rol`: 'suma' (la celda lo suma), 'nota' (solo va en la nota: TUA,
    extensión de horario) o 'sin_tc' (USD sin T.C. que todavía no suma);
    `monto_mxn` = lo que aporta a la celda; `gasto_mxn` = el total del gasto
    cuando la celda lleva solo una parte."""

    model_config = ConfigDict(extra="ignore")

    id: str = ""
    rol: str = "suma"
    monto_mxn: float | None = None
    gasto_mxn: float | None = None

    @field_validator("monto_mxn", "gasto_mxn", mode="before")
    @classmethod
    def _monto_liberal(cls, v: Any) -> float | None:
        return _monto_vinculo_liberal(v)

    @field_validator("rol", mode="before")
    @classmethod
    def _rol_liberal(cls, v: Any) -> str:
        """Texto ⇒ recortado y en minúsculas; vacío o no-texto ⇒ 'suma'."""
        if isinstance(v, str) and v.strip():
            return v.strip().lower()
        return "suma"


def _vinculos_liberales(v: Any) -> list:
    """Lista de vínculos tolerante: lo que no es lista o tupla ⇒ []; se
    quedan solo los dict (o `BalanceVinculoGasto`) con `id` de texto no
    vacío (recortado). Jamás un 422 por un metadato."""
    if not isinstance(v, list | tuple):
        return []
    out: list = []
    for x in v:
        if isinstance(x, BalanceVinculoGasto):
            if x.id.strip():
                out.append(x)
        elif isinstance(x, dict):
            ident = x.get("id")
            if isinstance(ident, str) and ident.strip():
                out.append({**x, "id": ident.strip()})
    return out


class BalanceAvionVuelo(BaseModel):
    """Una fila de la hoja maestra 'reporte horas <MATRÍCULA>' (1 vuelo)."""

    # Llave interna de orden cronológico (el API ya manda las filas
    # ordenadas); NO se pinta.
    orden_ts: str | None = None
    # CLAVE armada por el API: '#<folio> · <ID del cliente> · <cliente>' (sin
    # ID, '#<folio> · <cliente>'); se pinta tal cual.
    clave: str | None = None
    # ID de vuelo por cliente del API 0.0.75 (decisión 6 del lote del
    # 13-oct-2026): NO se pinta (la CLAVE ya llega armada); solo decide el
    # ancho de la columna A (`_formato.ancho_columna_clave`); None = API previo
    # o migración sin aplicar ⇒ libro byte a byte el de antes.
    clave_cliente: str | None = None
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
    # ADITIVOS (13-oct-2026, API 0.0.75, punto 5 del lote): los gastos que
    # componen OPERACIONES / PILOTO / OTROS de esta fila, empujados por el API
    # en las MISMAS ramas que su línea de `*_detalle`. Solo viajan al libro
    # JSON de la vista en pantalla (`vinculos` de la celda); el .xlsx no
    # cambia. Vacíos (API previo) ⇒ ninguna anotación.
    op_vinculos: list[BalanceVinculoGasto] = Field(default_factory=list)
    piloto_vinculos: list[BalanceVinculoGasto] = Field(default_factory=list)
    otros_vinculos: list[BalanceVinculoGasto] = Field(default_factory=list)
    permiso_afac_mxn: float | None = None
    # ADITIVO (12-oct-2026, API 0.0.69): la «Aportación AFAC USD/hr» del
    # AVIÓN de la fila. En el Balance general (hoja de vuelos de la flota)
    # cada fila trae la tarifa de su avión y PERMISO AFAC (PROVISIÓN) lleva
    # fórmula por fila (pedido de Ale: «una de las columnas de Permisos de
    # AFAC necesitamos que venga con la fórmula en el excel»). None = API
    # previo o avión sin tarifa (la columna va como valor / vacía).
    permiso_afac_usd_hr: float | None = None
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
    # movimientos'). El API la manda None (también el 0.0.65, por
    # compatibilidad): la columna COMISIONES MXN pinta `comisiones_mxn` y
    # solo cae a este campo con un API ≤ 0.0.64.
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
    # --- FACTURA VUELATOUR (30-sep-2026, API 0.0.45; pedido de Marie: «que
    # diga el num de factura que nosotros emitimos del servicio») ---
    # Folio de la factura del SERVICIO emitida al cliente, YA resuelto por la
    # cascada única del API `etiquetasFacturaDeVuelos` (CFDI timbrado vivo →
    # facturas EMITIDAS vigentes «A-0424, A-0430» → `vuelo.factura_folio`
    # tecleado → etiqueta del estatus «Facturado» / «Factura elaborada y
    # enviada» → desde el 13-oct-2026 (API 0.0.74, invariante 60 del API)
    # «Pendiente de factura» si algún cobro exige factura o hay solicitud
    # viva / «No lleva factura» si hay cobros y ninguno exige → None = SIN
    # COBROS). La misma etiqueta que «FACTURA VUELATOUR» del Libro Dinero y
    # «factura vuelatour» de 'otros movimientos'. Es del VUELO: en un
    # multi-avión todas sus filas traen la misma. None/ausente (API viejo) ⇒
    # celda vacía. Los textos nuevos se pintan TAL CUAL (`_factura_liberal`
    # acepta cualquier texto).
    factura_vuelatour: str | None = None

    @field_validator("factura_vuelatour", mode="before")
    @classmethod
    def _factura_liberal(cls, v: Any) -> Any:
        """Liberal con lo que manda NestJS: un folio NUMÉRICO (1234) se
        vuelve texto, un texto en blanco es None y cualquier otra cosa
        (booleano, objeto, NaN) es None — jamás un 422 que tumbe el balance
        entero por una columna informativa."""
        if v is None or isinstance(v, bool):
            return None
        if isinstance(v, int):
            return str(v)
        if isinstance(v, float):
            if not math.isfinite(v):
                return None
            return str(int(v)) if v.is_integer() else str(v)
        if isinstance(v, str):
            return v.strip() or None
        return None

    # --- EXTENSIÓN DE HORARIO pagada (1-oct-2026, API 0.0.47; pedido de Ale
    # sobre el #192 CUN-CTM-CUN: «me se está poniendo la extensión de
    # servicios como Operación y no va en ese apartado») ---
    # «Extensión y/o antelación de horario» pagada al aeropuerto, en MXN:
    # traslado al cliente igual que el TUA, SOLO informativo. La nota de la
    # celda OPERACIONES ya llega armada en `op_detalle` («Extensión de
    # horario (IVA incluido) $X**») y no suma en ninguna columna. El API la
    # manda AL FINAL de la fila y solo cuando es ≠ 0: sin ella (o con un API
    # ≤ 0.0.46) el libro es byte-idéntico al de siempre.
    extension_pagada_mxn: float | None = None

    # --- COMISIONES a cargo del avión (6-oct-2026, API 0.0.65; pedido: «la
    # comisión del banco y vendedor se puede ir a la columna de comisiones
    # del vendedor pero cambiar el nombre a "comisiones" y poner notas el
    # tipo de comisión… para que el monto real total cobrado ya sea después
    # de cualquier comisión»). Vigencia por fecha del vuelo (septiembre 2026,
    # clave `comisiones_al_avion_desde` del API). El API las calcula con su
    # fuente única `comisionesDelVuelo`; aquí solo se pintan. ---
    # Total de la celda COMISIONES MXN = parte del avión de la comisión
    # bancaria de los cobros + comisión del vendedor (hasta el API 0.0.73 como
    # PROVISIÓN; desde el 0.0.74, los pagos capturados como «Comisión del
    # vendedor (avión)», en cualquier fecha). None (vuelo sin nada que
    # absorber o API ≤ 0.0.64) ⇒ celda vacía y GANANCIA = REMANENTE.
    comisiones_mxn: float | None = None
    # Una línea por concepto, YA armada por el API (nota de la celda).
    comisiones_detalle: list[str] = Field(default_factory=list)
    # Las partes del total (informativas: si no llega el detalle, la nota se
    # compone con ellas). `comision_vendedor_prov_mxn` vale 0 desde el API
    # 0.0.74 (ya no hay provisión: compat de forma).
    comision_banco_avion_mxn: float | None = None
    comision_vendedor_prov_mxn: float | None = None
    # ADITIVOS (13-oct-2026, API 0.0.74, invariante 61 del API: comisión del
    # vendedor SOLO como gasto real, en dos categorías): la parte de ESTE
    # avión del Σ de los gastos «Comisión del vendedor (avión)» del vuelo
    # (MXN; 0 si el avión no los absorbe; los de «(VuelaTour)» nunca) y
    # cuántos pagos absorbió (0 = no absorbió). Viajan juntas con las de
    # arriba; aquí no se pintan en celda (solo el respaldo de la nota de
    # COMISIONES, `_nota_comisiones`, y la excepción de la nota de un vuelo
    # multi-avión, `_nota_participacion`). None = API previo.
    comision_vendedor_real_mxn: float | None = None
    comision_vendedor_real_pagos: int | None = None
    # ADITIVO (19-oct-2026, API 0.0.98, regla B): en las filas con
    # `comisiones_detalle`, la ganancia en dólares ANTES de comisiones
    # (= `ganancia_usd` + comisiones USD de la fila, a centavos; null cuando
    # `ganancia_usd` es null). Solo la usa la hoja de vuelos del Balance
    # GENERAL (que va antes de comisiones) con `usd_por_gasto`: ahí se pinta
    # como VALOR en GANANCIA USD; sin ella la celda va vacía. Sin la bandera
    # se ignora (libro byte a byte).
    ganancia_antes_comisiones_usd: float | None = None
    # --- CIERRE DE MES (16-oct-2026, API 0.0.88, invariante 89 del API) ---
    # Vuelo que PASÓ al mes siguiente por cobro pendiente al cerrar, o que
    # VINO de un mes anterior. `cierre_marca` es el texto YA armado por el
    # API que se pinta en ESTADO («Pasa a octubre 2026 (pendiente de cobro:
    # cobrado $0.00 de $36,632.89 MXN)» / «Viene de septiembre 2026» /
    # «Pendiente de cobro: pasará a octubre 2026 al cerrar»); `fila_sombra`
    # = fila del mes de ORIGEN de un diferido: horas reales y TODO el dinero
    # en null (el API lo manda así: ninguna SUM del libro la cuenta y aquí no
    # se vacía nada); `diferido_a` / `diferido_de` = 'YYYY-MM'. None/False
    # (API ≤ 0.0.87 o vuelo normal) ⇒ la fila de siempre, byte a byte.
    cierre_marca: str | None = None
    fila_sombra: bool = False
    diferido_a: str | None = None
    diferido_de: str | None = None

    @field_validator("cierre_marca", "diferido_a", "diferido_de", mode="before")
    @classmethod
    def _texto_cierre_liberal(cls, v: Any) -> str | None:
        return texto_cierre_liberal(v)

    @field_validator("fila_sombra", mode="before")
    @classmethod
    def _sombra_liberal(cls, v: Any) -> bool:
        return v is True

    @field_validator("comisiones_detalle", mode="before")
    @classmethod
    def _detalle_liberal(cls, v: Any) -> Any:
        """Nota informativa: null o algo que no es lista ⇒ []; se quedan solo
        las líneas de texto no vacías. Jamás un 422 por una nota."""
        if not isinstance(v, (list, tuple)):
            return []
        return [x.strip() for x in v if isinstance(x, str) and x.strip()]

    @field_validator("op_vinculos", "piloto_vinculos", "otros_vinculos", mode="before")
    @classmethod
    def _vinculos_liberales(cls, v: Any) -> Any:
        return _vinculos_liberales(v)

    @field_validator("clave_cliente", mode="before")
    @classmethod
    def _clave_cliente_liberal(cls, v: Any) -> str | None:
        return clave_cliente_liberal(v)


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
    # Σ extension_pagada_mxn de las filas (1-oct-2026, API 0.0.47): solo nota,
    # como el TUA. Viaja solo con suma ≠ 0 (libro de cada avión y
    # `consolidado` del general); sin la llave no se pinta nada nuevo.
    extension_pagada_mxn: float | None = None
    # Σ comisiones_mxn de las filas (6-oct-2026, API 0.0.65): TOTALES de la
    # columna COMISIONES MXN. Sin ella (API ≤ 0.0.64) la celda lleva
    # `comision_vendedor_mxn`, como siempre.
    comisiones_mxn: float | None = None
    # Σ de las partes del vendedor de las filas (informativas, no se pintan):
    # la provisión (0 desde el API 0.0.74) y, ADITIVO del 13-oct-2026 (API
    # 0.0.74, invariante 61), los pagos «Comisión del vendedor (avión)» que
    # absorben los aviones.
    comision_vendedor_prov_mxn: float | None = None
    comision_vendedor_real_mxn: float | None = None


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
    # T.C. con el que el API convirtió ESTA fila a pesos (su `tc_gasto` o,
    # sin él, el T.C. promedio de las hojas), EXACTAMENTE el número que
    # multiplicó (13-oct-2026, API 0.0.74). Solo lo manda la hoja 'permisos'
    # junto con `tc_por_fila`; aquí se pinta en la columna T.C. y MONTO MXN
    # = ROUND(MONTO ORIGINAL × T.C., 2). None = gasto en MXN, sin ningún
    # T.C. (`monto_mxn` None) o API previo.
    tc_aplicado: float | None = None
    # Litros de la carga (solo hoja "combustible").
    litros: float | None = None
    # Medio de pago YA legible que arma el API ("Efectivo", "Tarjeta BBVA
    # ···4321", "Transferencia HSBC"…) — columna PAGO al final de la hoja
    # "combustible" (11-sep-2026, pedido del cliente: saber CON QUÉ se pagó
    # cada carga). Aquí solo se pinta. None / API viejo = celda vacía.
    pago: str | None = None
    # Balance GENERAL: fila teñida con el color del avión.
    avion_color: str | None = None
    # Matrícula de la fila (hoja "combustible" del GENERAL, 28-ago): la hoja
    # se agrupa en SECCIONES por matrícula con subtotal. None = libro
    # individual (una sola matrícula: req.matricula).
    matricula: str | None = None
    # Solo hoja "refacciones" del GENERAL (29-ago): costo de la salida de
    # inventario (el que se GUARDÓ en la fila al registrarla: costo FIFO
    # hasta el API 0.0.35, último precio de compra desde el 0.0.36 —
    # 25-sep-2026) y VENTA al avión (= monto del gasto). Cuando vienen, la
    # hoja agrega columnas COSTO VUELATOUR / VENTA AL AVIÓN / GANANCIA
    # (ganancia = venta − costo, solo para mostrar; 0 mientras la salida se
    # cargue a costo). None = sin columnas extra.
    costo_mxn: float | None = None
    venta_mxn: float | None = None
    # ADITIVOS (13-oct-2026, API 0.0.75, punto 5 del lote): el id del gasto de
    # la fila (en un parcial de reparto manual, el del gasto PADRE) y, si la
    # fila es una parte, cuál ('reparto' | 'remanente'). Solo viajan al libro
    # JSON de la vista en pantalla (`vinculos` de la celda MONTO MXN / VENTA
    # AL AVIÓN); el .xlsx no cambia. None = API previo.
    gasto_id: str | None = None
    gasto_parte: str | None = None

    @field_validator("gasto_id", mode="before")
    @classmethod
    def _gasto_id_liberal(cls, v: Any) -> str | None:
        """Texto no vacío ⇒ recortado; lo demás ⇒ None (jamás un 422)."""
        if isinstance(v, str) and v.strip():
            return v.strip()
        return None

    @field_validator("gasto_parte", mode="before")
    @classmethod
    def _gasto_parte_liberal(cls, v: Any) -> str | None:
        """Solo 'reparto' o 'remanente' (sin distinguir mayúsculas ni
        espacios); cualquier otra cosa ⇒ None."""
        if isinstance(v, str) and v.strip().lower() in ("reparto", "remanente"):
            return v.strip().lower()
        return None


class BalanceAvionHojaGastos(BaseModel):
    """Hoja de gastos: ledger + resumen (USD al TC promedio del periodo)."""

    filas: list[BalanceAvionGastoFila] = Field(default_factory=list)
    total_mxn: float | None = None
    usd: float | None = None
    usd_hr: float | None = None
    # True = cada fila trae `tc_aplicado` y la pestaña 'permisos' se pinta
    # con la columna T.C. y MONTO MXN como fórmula (13-oct-2026, API 0.0.74).
    # None/ausente = API ≤ 0.0.73 ⇒ la hoja de siempre, byte a byte.
    tc_por_fila: bool | None = None


class BalanceAvionHojaCombustible(BalanceAvionHojaGastos):
    """Hoja 'combustible' (26-ago-2026): el gas del avión POR MES, con
    litros y precio por litro promedio — ya no va por vuelo."""

    litros_total: float | None = None
    precio_litro_prom: float | None = None


class BalanceAvionSocio(BaseModel):
    nombre: str = ""
    porcentaje: float | None = None  # % vigente (0–100)
    monto_usd: float | None = None  # utilidad cobrada × pct/100
    # ADITIVO (6-oct-2026, API 0.0.59): el socio es la EMPRESA (Aero Charter
    # Cancún, `usuario.es_empresa`). Solo lo usa el bloque «VUELATOUR
    # (empresa)» del Balance general para citar la celda MONTO USD de este
    # socio; la fila del socio se pinta igual. None = API viejo.
    es_empresa: bool | None = None
    # ADITIVO (18-oct-2026, API 0.0.93; balances EN PESOS): lo que le toca
    # en pesos = round2(pct/100 × `utilidad_cobrada_mxn`). Solo se pinta con
    # la bandera `balance_en_pesos`; None = sin pesos (o API viejo).
    monto_mxn: float | None = None
    # ENTREGAS A SOCIOS (21-oct-2026, invariante 100 del API; pedido del
    # cliente: «agregar en el balance general y por avión los pagos que se
    # vayan dando a los socios»). Lo YA entregado a este socio «por este
    # avión» en los meses del periodo —entregas vivas de Pagos a socios—, en
    # PESOS (`entregado_mxn`: lo que cada entrega descontó del saldo), el USD
    # informativo (`entregado_usd`), el conteo (`entregas`) y lo que FALTA
    # (`por_entregar_mxn` = monto a recibir en pesos − entregado; NEGATIVO =
    # se le adelantó). Todo lo calcula el API; aquí solo se pinta, y SOLO
    # cuando el bloque trae `entregado_socios_mxn` (el discriminador). None
    # = sin dato ⇒ celda vacía / «—», jamás un 0 ni una resta aquí. Liberal:
    # jamás un 422 por un metadato.
    entregado_mxn: float | None = None
    entregado_usd: float | None = None
    entregas: int | None = None
    por_entregar_mxn: float | None = None

    @field_validator("entregado_mxn", "entregado_usd", "por_entregar_mxn", mode="before")
    @classmethod
    def _entregas_numeros_liberales(cls, v: Any) -> float | None:
        return numero_cierre_liberal(v)

    @field_validator("entregas", mode="before")
    @classmethod
    def _entregas_conteo_liberal(cls, v: Any) -> int | None:
        n = numero_cierre_liberal(v)
        return int(n) if n is not None and n >= 0 and n.is_integer() else None


def bandera_liberal(v: Any) -> bool:
    """Bandera ADITIVA del API: SOLO un `true` JSON la enciende; null, texto
    o cualquier otra cosa ⇒ False (jamás un 422 por una bandera)."""
    return v is True


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
    # COMISIONES a cargo del avión en la cascada del Balance GENERAL
    # (7-oct-2026, API 0.0.66; pedido: «estas comisiones se están duplicando
    # en la general, ya que las tenemos en "Otros movimientos"… deben
    # aparecer en el individual de los avioncitos»). La hoja de vuelos del
    # general ya no lleva la columna COMISIONES y su GANANCIA va ANTES de
    # ellas; para seguir cuadrando con lo que ve el socio, la cascada de cada
    # avión del general las muestra en dos líneas: `utilidad_antes_comisiones
    # _usd` (= `utilidad_antes_usd` + `comisiones_usd`) y `comisiones_usd`
    # (lo que absorbe el avión: por vuelo, la MISMA cifra que descuenta el
    # reparto a socios; el TOTAL puede diferir del reparto cuando un vuelo que
    # aún no se completa tiene un anticipo con comisión bancaria — el balance
    # cuenta su cobro y el reparto no lee el vuelo, solo COMPLETADOS y
    # CANCELADOS). El API las calcula; aquí solo se pintan. None (API ≤
    # 0.0.65) o comisiones en 0 ⇒ la cascada de siempre. El libro INDIVIDUAL
    # las ignora: ahí ya restan, por vuelo, en su columna COMISIONES.
    comisiones_usd: float | None = None
    utilidad_antes_comisiones_usd: float | None = None
    # ADITIVO (13-oct-2026, API 0.0.75, invariante «INGRESO ASIGNADO A UN
    # AVIÓN» del API): TOTAL USD de la hoja «Otros ingresos» de ESTE
    # avión (= `otros_ingresos_avion.usd`). La cascada lo SUMA en la línea
    # «(+) OTROS INGRESOS USD», entre «(−) PERMISOS USD» y «UTILIDAD DESPUÉS
    # DE GASTOS USD»; `utilidad_despues_usd` del API YA lo incluye. Viaja
    # junto con la hoja: sin `otros_ingresos_avion` (API ≤ 0.0.74) no hay
    # línea. None con la hoja presente = algún ingreso USD sin T.C. ⇒ celda
    # vacía (y la utilidad también, el null se propaga en el API). OJO: NO es
    # `otros_ingresos_usd` de los vuelos/totales (TUAs/extras de VuelaTour).
    otros_ingresos_avion_usd: float | None = None
    # BALANCE EN PESOS (18-oct-2026, API 0.0.93; pedido del cliente: «en los
    # balances … debe venir todo en pesos porque casi siempre se les
    # transfiere a los socios en pesos mexicanos, y que los dólares solo
    # sean informativos»). La MISMA cascada en PESOS REALES, calculada por el
    # API con lo que el libro ya tiene en pesos (GANANCIA MXN de las filas, el
    # TOTAL MXN de cada hoja, el POR COBRAR MXN). ADITIVAS: solo se pintan
    # con la bandera raíz `balance_en_pesos`; None = sin pesos (alguna fila
    # sin monto en pesos) ⇒ «—», jamás un 0 ni una conversión aquí.
    utilidad_antes_mxn: float | None = None
    combustible_mxn: float | None = None
    gastos_indirectos_mxn: float | None = None
    refacciones_mxn: float | None = None
    otros_mxn: float | None = None
    permisos_mxn: float | None = None
    otros_ingresos_avion_mxn: float | None = None
    utilidad_despues_mxn: float | None = None
    por_cobrar_mxn: float | None = None
    utilidad_cobrada_mxn: float | None = None
    # Solo cuando viajan `comisiones_usd` (cascada del GENERAL).
    comisiones_mxn: float | None = None
    utilidad_antes_comisiones_mxn: float | None = None
    # RESERVA DE OVERHAUL MARCADA en el balance, SIN RESTAR (19-oct-2026, API
    # 0.0.98; decisión del cliente: el reparto a socios ya resta la reserva
    # de overhaul —horas voladas del periodo × Σ `monto_por_hora_usd` de los
    # motores— y el balance NO la resta: solo la MARCA, muy clara, para que se
    # entienda qué debe aportar el avión para el overhaul). ADITIVAS:
    # `reserva_overhaul_usd` es el DISCRIMINADOR (número; 0.0 si no hubo horas
    # o no hay tarifa; None = API previo ⇒ ningún bloque y el libro de siempre
    # byte a byte); la cascada (filas, índices y fórmulas) NO cambia con la
    # llave: el bloque informativo va DESPUÉS de los socios.
    # `reserva_overhaul_mxn` = round2(usd × T.C. oficial del cierre del
    # periodo); None con la USD presente = sin T.C. oficial ⇒ «—» en pesos.
    # Lo demás solo alimenta el texto «Cómo se calcula» (liberal: jamás un 422
    # por un metadato): `reserva_overhaul_usd_hr` es la tarifa Σ motores
    # (USD/hr; el texto la muestra por motor: Σ ÷ motores).
    # `reserva_overhaul_incompleta` True = voló horas pero el avión no tiene
    # tarifa configurada ⇒ no se puede calcular lo que debe aportar.
    reserva_overhaul_usd: float | None = None
    reserva_overhaul_mxn: float | None = None
    reserva_overhaul_horas_hr: float | None = None  # horas voladas del periodo
    reserva_overhaul_usd_hr: float | None = None  # tarifa Σ motores, USD/hr
    reserva_overhaul_motores: int | None = None
    reserva_overhaul_tc: float | None = None  # T.C. oficial con que salió el MXN
    reserva_overhaul_tc_dia: str | None = None  # 'YYYY-MM-DD', día Cancún del T.C.
    reserva_overhaul_incompleta: bool | None = None
    # ENTREGAS A SOCIOS (21-oct-2026, invariante 100 del API): Σ de
    # `socios[].entregado_mxn` y Σ de `socios[].por_entregar_mxn` (None si a
    # algún socio le faltan pesos), en pesos. `entregado_socios_mxn` es el
    # DISCRIMINADOR (`_con_entregas_socios` en el render: número, 0.0 también
    # cuenta); None = API previo o cuenta corriente de los socios sin su
    # migración ⇒ ni columnas, ni totales, ni nota: libro byte a byte.
    entregado_socios_mxn: float | None = None
    por_entregar_socios_mxn: float | None = None

    @field_validator(
        "reserva_overhaul_horas_hr",
        "reserva_overhaul_usd_hr",
        "reserva_overhaul_tc",
        "entregado_socios_mxn",
        "por_entregar_socios_mxn",
        mode="before",
    )
    @classmethod
    def _reserva_numeros_liberales(cls, v: Any) -> float | None:
        return numero_cierre_liberal(v)

    @field_validator("reserva_overhaul_motores", mode="before")
    @classmethod
    def _reserva_motores_liberal(cls, v: Any) -> int | None:
        n = numero_cierre_liberal(v)
        return int(n) if n is not None and n >= 0 and n.is_integer() else None

    @field_validator("reserva_overhaul_tc_dia", mode="before")
    @classmethod
    def _reserva_tc_dia_liberal(cls, v: Any) -> str | None:
        return _texto_liberal(v)

    @field_validator("reserva_overhaul_incompleta", mode="before")
    @classmethod
    def _reserva_incompleta_liberal(cls, v: Any) -> bool | None:
        """Solo un `true` JSON la enciende; lo demás (null, false, texto) ⇒
        None = con tarifa."""
        return True if v is True else None


def _texto_liberal(v: Any) -> str | None:
    """Texto no vacío ⇒ recortado; vacío o no-texto ⇒ None (metadato: jamás
    un 422)."""
    if isinstance(v, str) and v.strip():
        return v.strip()
    return None


def _conteo_liberal(v: Any) -> int:
    """Conteo de un resumen (solo alimenta textos): entero ≥ 0 o número
    entero ⇒ int; lo demás (None, bool, texto, negativo) ⇒ 0."""
    if isinstance(v, bool):
        return 0
    if isinstance(v, int):
        return max(v, 0)
    if isinstance(v, float) and math.isfinite(v) and v.is_integer():
        return max(int(v), 0)
    return 0


class BalanceAfacVuelo(BaseModel):
    """AFAC POR VUELO (16-oct-2026, API 0.0.89, invariante 90 del API;
    decisiones del cliente del 16-oct: filas UNA POR VUELO, el costo real es
    EGRESO DE VUELATOUR y el avión sigue pagando lo RESERVADO). Viaja como
    `afac` de cada fila `seccion == 'AFAC_VUELO'` de 'otros movimientos'. Los
    montos de la fila (EGRESO = costo real, INGRESO = reservado, remanente)
    ya vienen en las llaves de siempre de la fila; esto es el METADATO de la
    celda editable del libro JSON (vínculo `afac_real` en la celda EGRESO):
    aquí no se calcula nada. Todo liberal: jamás un 422 por un metadato."""

    model_config = ConfigDict(extra="ignore")

    vuelo_id: str = ""
    matricula: str = ""
    horas_cobradas: float | None = None
    tarifa_usd_hr: float | None = None
    tc_costos: float | None = None
    # `permiso_afac_mxn` de la fila del libro TAL CUAL (null = sin tarifa).
    reservado_mxn: float | None = None
    # null = sin capturar (o fila no «dueña» de un vuelo multi-avión).
    costo_real_mxn: float | None = None
    remanente_mxn: float | None = None
    capturado_por_nombre: str | None = None
    capturado_at: str | None = None
    notas: str | None = None
    mes_contable: str | None = None  # 'AAAA-MM'
    # disponible (migración) && fila dueña && mes contable abierto. El ROL
    # lo decide el panel (el libro no conoce al usuario) y el PUT lo valida.
    editable: bool = False
    motivo_no_editable: str | None = None
    cierre_marca: str | None = None

    @field_validator("vuelo_id", "matricula", mode="before")
    @classmethod
    def _id_liberal(cls, v: Any) -> str:
        return _texto_liberal(v) or ""

    @field_validator(
        "horas_cobradas",
        "tarifa_usd_hr",
        "tc_costos",
        "reservado_mxn",
        "costo_real_mxn",
        "remanente_mxn",
        mode="before",
    )
    @classmethod
    def _numero_liberal(cls, v: Any) -> float | None:
        return _monto_vinculo_liberal(v)

    @field_validator(
        "capturado_por_nombre",
        "capturado_at",
        "notas",
        "mes_contable",
        "motivo_no_editable",
        "cierre_marca",
        mode="before",
    )
    @classmethod
    def _textos_liberales(cls, v: Any) -> str | None:
        return _texto_liberal(v)

    @field_validator("editable", mode="before")
    @classmethod
    def _editable_estricto(cls, v: Any) -> bool:
        """Solo un `true` JSON abre la celda; cualquier otra cosa ⇒ False."""
        return v is True


class BalanceAfacPago(BaseModel):
    """Un gasto marcado «Pago de AFAC» del periodo (lo paga VuelaTour; no
    resta a ningún avión): renglón del bloque «AFAC DEL PERIODO» de 'otros
    movimientos' (16-oct-2026, API 0.0.89). `detalle` llega armado del API
    (« · XA-VGV» / « · vuelo #280»); `monto_mxn` null = USD sin T.C. (no
    suma en la comparación)."""

    model_config = ConfigDict(extra="ignore")

    gasto_id: str | None = None
    fecha: str | None = None
    detalle: str = ""
    matricula: str | None = None
    monto_mxn: float | None = None
    moneda_original: str | None = None
    monto_original: float | None = None

    @field_validator("gasto_id", mode="before")
    @classmethod
    def _gasto_id_liberal(cls, v: Any) -> str | None:
        return _texto_liberal(v)

    @field_validator("detalle", mode="before")
    @classmethod
    def _detalle_liberal(cls, v: Any) -> str:
        return v if isinstance(v, str) else ""


class BalanceAfacResumen(BaseModel):
    """Resumen «AFAC DEL PERIODO» de 'otros movimientos' (16-oct-2026, API
    0.0.89, invariante 90): Σ reservado / Σ costo real / remanente de las
    filas AFAC por vuelo y la comparación contra los gastos marcados «Pago
    de AFAC». Lo calcula el API; aquí se pinta (con fórmulas verificadas
    sobre las filas cuando son contiguas). Montos None ⇒ celda vacía (nunca
    un 0 falso)."""

    model_config = ConfigDict(extra="ignore")

    vuelos: int = 0
    vuelos_con_reservado: int = 0
    vuelos_con_real: int = 0
    vuelos_sin_real: int = 0
    reservado_mxn: float | None = None
    costo_real_mxn: float | None = None
    remanente_mxn: float | None = None
    pagado_mxn: float | None = None
    pagos_sin_tc: int = 0
    diferencia_pagado_mxn: float | None = None
    reales_fuera_de_libro_mxn: float | None = None
    pagos: list[BalanceAfacPago] = Field(default_factory=list)

    @field_validator(
        "vuelos", "vuelos_con_reservado", "vuelos_con_real", "vuelos_sin_real", "pagos_sin_tc",
        mode="before",
    )  # fmt: skip
    @classmethod
    def _conteos(cls, v: Any) -> int:
        return _conteo_liberal(v)


class BalanceOtroMovimientoFila(BaseModel):
    """Fila de la pestaña "Otros movimientos" (28-ago, hoja manual del
    cliente): egreso pagado apareado con el ingreso cobrado por concepto y
    su remanente. Cualquier lado puede venir vacío (solo-ingreso o
    solo-egreso); viene YA calculada del API — aquí jamás se recalcula.
    Desde el 28-ago-2026 tarde (regla A) incluye la comisión del vendedor
    como INGRESO de VuelaTour y su pago al vendedor como EGRESO apareado
    (PROVISIÓN a la fecha del vuelo mientras no exista gasto real; la nota
    de la celda lo dice). Desde el 28-sep-2026 (API 0.0.39, invariante 31)
    el gasto real con categoría «Comisión del vendedor» ligado al vuelo
    REEMPLAZA a la provisión: el concepto llega hecho del API («pago
    comisión vendedor (X) · gasto real», con «· parcial: faltan $… MXN» o
    «· excede $… MXN») y aquí se pinta tal cual."""

    # Con ID de vuelo por cliente (API 0.0.75): '#<folio> · <ID>'; sin él, la
    # clave de siempre ('vt' + primera palabra + folio). ' · ESTADO' lo
    # agrega este módulo en estados no normales.
    clave: str = ""
    # ID de vuelo por cliente del API 0.0.75 (decisión 6 del lote del
    # 13-oct-2026): NO se pinta (la CLAVE ya llega armada); solo decide el
    # ancho de la columna A (`_formato.ancho_columna_clave`); None = API previo
    # o migración sin aplicar ⇒ libro byte a byte el de antes.
    clave_cliente: str | None = None
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
    # «factura vuelatour» (col J): la MISMA etiqueta que `factura_vuelatour`
    # de la hoja maestra; desde el 13-oct-2026 (invariante 60 del API) puede
    # ser «Pendiente de factura» / «No lleva factura» (texto tal cual).
    factura: str | None = None
    # Desglose línea por línea (una fila por vuelo, 28-ago): se pinta como
    # COMENTARIO de la celda de ingreso / egreso. None = API viejo.
    nota_ingreso: str | None = None
    nota_egreso: str | None = None
    # ADITIVO (13-oct-2026, API 0.0.74, invariante 62 del API): 'AFAC' = fila
    # por avión de PERMISOS AFAC (ingreso: provisión del periodo; egreso: los
    # permisos pagados CITADOS con $0.00 «a cargo del avión» —ya restan en la
    # cascada de su avión—), al final de `filas`; con `afac_en_filas` se pinta
    # un encabezado de sección antes de la primera. None = fila de vuelo o
    # suelta, o API previo.
    # Desde el API 0.0.89 (16-oct-2026, invariante 90): 'AFAC_VUELO' = fila
    # AFAC POR VUELO (ingreso = RESERVADO del vuelo, egreso = COSTO REAL que
    # VuelaTour paga a AFAC, null = por capturar), al final de `filas`,
    # contiguas; con `afac_por_vuelo` se pinta la fila separadora y la celda
    # EGRESO lleva el vínculo editable `afac_real` en el libro JSON.
    seccion: str | None = None
    # ADITIVO (16-oct-2026, API 0.0.89, invariante 90): SOLO en las filas
    # `seccion == 'AFAC_VUELO'`, ÚLTIMA llave de la fila. Datos del vuelo para
    # la celda editable (vínculo `afac_real` del libro JSON); los montos de
    # la fila (ingreso/egreso/remanente) ya vienen en sus llaves de siempre.
    afac: BalanceAfacVuelo | None = None

    @field_validator("clave_cliente", mode="before")
    @classmethod
    def _clave_cliente_liberal(cls, v: Any) -> str | None:
        return clave_cliente_liberal(v)


class BalanceAfacProvisionAvion(BaseModel):
    """Provisión AFAC del periodo de UN avión: Σ de su columna PERMISO AFAC
    (PROVISIÓN) — ya cargada a cada vuelo en su libro."""

    matricula: str = ""
    tarifa_usd_hr: float | None = None
    vuelos: int = 0
    monto_mxn: float = 0.0


class BalanceAfacPagado(BaseModel):
    """Permiso AFAC PAGADO en el periodo (gasto categoría PERMISO; la fila
    de la hoja «permisos» del libro individual de ese avión)."""

    matricula: str = ""
    fecha: str | None = None
    detalle: str = ""
    monto_mxn: float | None = None


class BalanceAfacOtrosMovimientos(BaseModel):
    """Bloque «PERMISOS AFAC» de la pestaña «otros movimientos» (12-oct-2026,
    API 0.0.69; pedido de Ale: «en el balance general en la página de otros
    movimientos no se ven los AFAC, aquí deben estar también»). INFORMATIVO:
    la provisión ya resta en cada avión (columna PERMISO AFAC (PROVISIÓN) de
    la hoja de vuelos) y los permisos pagados viven en la hoja «permisos» del
    libro individual; aquí se ven juntos, provisión vs pagado, SIN entrar a
    los TOTALES de la pestaña (si entraran, se contarían dos veces).
    Solo lo manda un API 0.0.69–0.0.73 (desde el 0.0.74 las filas AFAC suman
    en `filas`, invariante 62 del API); se pinta SOLO sin `afac_en_filas`."""

    provision: list[BalanceAfacProvisionAvion] = Field(default_factory=list)
    provision_total_mxn: float = 0.0
    pagados: list[BalanceAfacPagado] = Field(default_factory=list)
    pagados_total_mxn: float = 0.0
    # provisión − pagado: lo que queda del fondo (negativo = se pagó más de
    # lo provisionado en el periodo).
    remanente_mxn: float = 0.0


class BalanceHojaOtrosMovimientos(BaseModel):
    filas: list[BalanceOtroMovimientoFila] = Field(default_factory=list)
    # Movimientos SIN avión y SIN vuelo (dinero de empresa).
    filas_sueltas: list[BalanceOtroMovimientoFila] = Field(default_factory=list)
    # Comisión del vendedor como GASTO (28-sep-2026, API 0.0.39, invariante
    # 31 del API): el API manda `true` SOLO si algún vuelo del periodo tiene
    # un gasto real «Comisión del vendedor» (que reemplaza a su provisión);
    # si no, la clave NO viaja. Con `true` las 3 leyendas del Balance general
    # que hablan de la PROVISIÓN mencionan también el gasto real; sin ella
    # (None/False, o un API viejo) el Excel es byte-idéntico al de antes —
    # esas leyendas siguen siendo verdaderas mientras no haya gasto real.
    hay_pago_vendedor_real: bool | None = None
    # ADITIVO (12-oct-2026, API 0.0.73, invariante 58 del API; pedido del
    # cliente: «provisión: los gastos que todavía no se hacen no deben
    # aparecer»): `true` = el libro viene de un API que YA NO pinta la
    # PROVISIÓN del pago al vendedor en las filas de 'otros movimientos':
    # mientras no exista el gasto real «Comisión del vendedor» ligado al
    # vuelo la fila no lleva egreso (ni concepto, ni monto, ni marca); el
    # ingreso cobrado sigue y el remanente = ingreso. Con gasto real, todo
    # como antes. Con la bandera cambian SOLO las leyendas (fila 2 de la
    # pestaña, nota del RESUMEN y base del bloque VUELATOUR (empresa)); sin
    # ella (None/False, o un API previo que sigue provisionando) los textos
    # de siempre, byte a byte. El amarillo NO depende de esta bandera: lo
    # decide la marca «· PROVISIÓN (» de cada fila (`_es_provision`), que
    # con el API nuevo ya no llega.
    sin_provision_vendedor: bool | None = None
    # ADITIVO (13-oct-2026, API 0.0.74, invariante 61 del API; decisión 3 del
    # lote del cliente: «comisión del vendedor SOLO como gasto real», con su
    # ajuste de DOS categorías): ya no hay provisión en NINGÚN libro; el pago
    # capturado como «Comisión del vendedor (VuelaTour)» egresa en esta
    # pestaña y el capturado como «Comisión del vendedor (avión)» lo carga el
    # avión en su columna COMISIONES (aquí, apareado con egreso $0.00 «· a
    # cargo del avión …»; si ningún avión lo carga, lo paga VuelaTour). Con
    # `true` cambian SOLO leyendas (fila 2, RESUMEN, bloque VUELATOUR) y MANDA
    # sobre `sin_provision_vendedor`; sin ella (None/False, API ≤ 0.0.73) los
    # textos de siempre, byte a byte.
    comision_vendedor_solo_gasto_real: bool | None = None
    # ADITIVO (12-oct-2026, API 0.0.69): bloque informativo «PERMISOS AFAC»
    # (provisión vs pagado). None = API previo o periodo sin provisión ni
    # permisos pagados: la pestaña es la de siempre. Desde el API 0.0.74 ya
    # no viaja (las filas AFAC suman en `filas`); se conserva por skew y se
    # pinta SOLO sin `afac_en_filas`.
    afac: BalanceAfacOtrosMovimientos | None = None
    # ADITIVO (13-oct-2026, API 0.0.74, invariante 62 del API): True = las
    # filas AFAC ya viajan en `filas` (`seccion == 'AFAC'`) y SUMAN en TOTALES
    # y en el bloque VUELATOUR; con ella el bloque `afac` NO se pinta aunque
    # llegara y se activan `_OM_SECCION_AFAC`, `_OM_AFAC_EN_FILAS`,
    # `_EMPRESA_*_AFAC` y `_EMPRESA_AFAC_EN_FILAS`. None/False = API previo ⇒
    # libro de siempre, byte a byte.
    afac_en_filas: bool | None = None
    # ADITIVO (13-oct-2026, API 0.0.75, invariante «INGRESO ASIGNADO A UN
    # AVIÓN» del API): `true` = los otros ingresos registrados en Ingresos y
    # ASIGNADOS a un avión ya no viajan en `filas_sueltas` (cuentan en la
    # hoja «Otros ingresos» del libro de ese avión, en su cascada y en su
    # reparto). Viaja SIEMPRE desde el 0.0.75; aquí SOLO agrega
    # `_OM_OTROS_INGRESOS_POR_AVION` al final de la fila 2. None/False (API ≤
    # 0.0.74) ⇒ la pestaña de siempre, byte a byte.
    otros_ingresos_por_avion: bool | None = None
    # ADITIVO (16-oct-2026, API 0.0.89, invariante 90 del API): `true` = las
    # filas AFAC POR VUELO (`seccion == 'AFAC_VUELO'`, una por vuelo con AFAC:
    # INGRESO = reservado, EGRESO = costo real) viajan contiguas al final de
    # `filas` y SUMAN en TOTALES y en el bloque VUELATOUR como cualquier fila.
    # Con ella: fila separadora antes de la primera, vínculo editable
    # `afac_real` en su celda EGRESO (libro JSON), leyenda de la fila 2 y
    # textos del bloque VUELATOUR; MANDA sobre `afac_en_filas` (el 0.0.89 ya no
    # la manda). Solo con ≥ 1 fila AFAC_VUELO. None/False ⇒ byte a byte.
    afac_por_vuelo: bool | None = None
    # ADITIVO (16-oct-2026, API 0.0.89): bloque «AFAC DEL PERIODO» debajo de
    # TOTALES (reservado vs costo real vs pagado a AFAC con los gastos
    # marcados «Pago de AFAC»). Viaja con filas AFAC_VUELO o con pagos
    # marcados. None ⇒ sin bloque.
    afac_resumen: BalanceAfacResumen | None = None


class BalanceAvionOtroIngresoFila(BaseModel):
    """Fila de la hoja «Otros ingresos» de un avión (13-oct-2026, API 0.0.75,
    invariante «INGRESO ASIGNADO A UN AVIÓN» del API): un ingreso registrado
    en Ingresos (clave ING-n) y ASIGNADO a ese avión (reembolso de su
    servicio, venta de una pieza suya, intereses…). Todo llega YA calculado
    de la fuente única del API (`filaLibroDeIngreso`): bruto, comisión y
    neto en pesos con el T.C. PROPIO del ingreso; aquí solo se pinta
    (fórmulas que reproducen esos números). None = celda vacía, nunca un 0
    falso."""

    # Id del ingreso: viaja para el futuro vínculo celda → registro (no se
    # pinta todavía).
    ingreso_id: str | None = None
    clave: str = ""  # 'ING-4'
    fecha: str | None = None  # 'YYYY-MM-DD' (día Cancún)
    # Etiqueta amable de la categoría ('Reembolsos y devoluciones recibidos').
    categoria: str | None = None
    detalle: str | None = None  # 'ING-4 · TRANSF INTERBANCARIA SPEI · angel alvarez'
    moneda_original: str | None = None  # 'MXN' | 'USD'
    monto_original: float | None = None  # BRUTO en su moneda
    # USD: el `tc_usd_mxn` EXACTO del ingreso (sin redondear); MXN: None.
    tc_aplicado: float | None = None
    ingreso_mxn: float | None = None
    comision_mxn: float | None = None  # None = sin comisión bancaria
    neto_mxn: float | None = None  # ingreso − comisión
    cuenta: str | None = None  # 'EJE SCOTIA · Scotiabank' | 'Efectivo / caja'
    notas: str | None = None
    # ADITIVOS (14-oct-2026, API 0.0.77; decisión del cliente: «unificar la
    # conversión del ingreso de un avión entre balance y reparto… hazlo, si es
    # lo mejor»): la conversión ÚNICA por ingreso del API (`usdDeIngresoAvion`,
    # la MISMA del reparto a socios y del tablero). `tc_usd` = el T.C. con que
    # el API pasó ESTE ingreso a USD — en pesos, su T.C. propio, el del gasto
    # que reembolsa (reembolso LIGADO) o el oficial del día de su `fecha`,
    # EXACTO (sin redondear: aquí se cita en ROUND(NETO MXN ÷ T.C., 2)); en
    # dólares, None —; `neto_usd` = su NETO en USD (lo que SUMA; None = pesos
    # sin ningún T.C.); `tc_usd_origen` = 'PROPIO' | 'GASTO_LIGADO' |
    # 'OFICIAL_DEL_DIA' | None (solo el oficial pinta la celda de azul) y
    # `tc_usd_nota` = la nota de la celda T.C. A USD (None con T.C. propio o
    # en dólares). Solo se pintan con la bandera de la hoja
    # `BalanceAvionHojaOtrosIngresos.usd_por_ingreso`; sin ella se ignoran.
    tc_usd: float | None = None
    neto_usd: float | None = None
    tc_usd_origen: str | None = None
    tc_usd_nota: str | None = None
    # Solo en el consolidado FLOTA del Balance general (hoja «otros ingresos
    # aviones»): el DETALLE ya llega con «MATRÍCULA · » al frente.
    avion_color: str | None = None
    matricula: str | None = None

    @field_validator("clave", mode="before")
    @classmethod
    def _clave_liberal(cls, v: Any) -> Any:
        """Texto de una celda: null ⇒ '' (jamás un 422 por una etiqueta)."""
        if v is None:
            return ""
        if isinstance(v, int | float) and not isinstance(v, bool):
            return str(v)
        return v

    @field_validator("tc_usd_origen", "tc_usd_nota", mode="before")
    @classmethod
    def _texto_tc_usd_liberal(cls, v: Any) -> str | None:
        """Origen y nota del T.C. A USD: texto recortado; vacío o algo que no
        es texto ⇒ None (sin relleno ni nota). Jamás un 422 por una nota."""
        if isinstance(v, str):
            return v.strip() or None
        return None


class BalanceAvionHojaOtrosIngresos(BaseModel):
    """Hoja «Otros ingresos» del libro de un avión (y, en el Balance general,
    la consolidada «otros ingresos aviones»). `total_mxn` = Σ NETO MXN; `usd`
    = TOTAL USD que SUMA en la cascada (None = algún ingreso USD sin T.C.;
    en el general, None si algún libro trajo None). Con `usd_por_ingreso`
    (API 0.0.77) `usd` = Σ de los `neto_usd` de las filas (None = alguna fila
    sin `neto_usd`: pesos sin ningún T.C.; vacía = 0)."""

    filas: list[BalanceAvionOtroIngresoFila] = Field(default_factory=list)
    total_mxn: float | None = None
    usd: float | None = None
    # ADITIVA (14-oct-2026, API 0.0.77; decisión del cliente: conversión ÚNICA
    # del ingreso de un avión, la del reparto a socios): `true` = cada ingreso
    # viaja con su `tc_usd`/`neto_usd` y `usd` es su suma. El API la manda
    # SIEMPRE como ÚLTIMA llave (libro individual, `aviones[]` y consolidado
    # FLOTA). Con ella la hoja gana las columnas T.C. A USD (J) y NETO USD (K),
    # TOTAL USD = ROUND(SUM(K8:Kn),2) (en el general también) y las notas
    # `*_POR_INGRESO`. None/False (API ≤ 0.0.76) ⇒ la hoja y las notas de
    # siempre, byte a byte (TOTAL USD = TOTAL NETO MXN ÷ TC PROMEDIO).
    usd_por_ingreso: bool | None = None

    @field_validator("filas", mode="before")
    @classmethod
    def _filas_liberal(cls, v: Any) -> Any:
        """NestJS puede mandar null en una lista vacía: jamás un 422."""
        return [] if v is None else v


def texto_cierre_liberal(v: Any) -> str | None:
    """Texto informativo del cierre de mes (16-oct-2026; balances y reparto,
    ÚNICA regla): número finito ⇒ texto (el API manda el folio como número),
    blanco o cualquier otra cosa ⇒ None (jamás un 422 por una etiqueta)."""
    if isinstance(v, bool):
        return None
    if isinstance(v, int | float):
        return str(v) if math.isfinite(v) else None
    if isinstance(v, str):
        return v.strip() or None
    return None


def numero_cierre_liberal(v: Any) -> float | None:
    """Monto ADITIVO del cierre (16-oct-2026): número finito ⇒ float; lo
    demás (null, texto, NaN, booleano) ⇒ None. Jamás un 422 por un dato
    informativo nuevo."""
    if isinstance(v, bool):
        return None
    if isinstance(v, int | float):
        return float(v) if math.isfinite(v) else None
    return None


def lista_cierre_liberal(v: Any) -> list:
    """NestJS puede mandar null en una lista vacía: jamás un 422."""
    return v if isinstance(v, list) else []


class BalanceCierreVuelo(BaseModel):
    """Un vuelo de las listas del cierre (16-oct-2026, API 0.0.88): el que
    PASA al mes siguiente (`destino`), VIENE de un mes anterior (`origen`) o
    PASARÍA al cerrar un mes abierto. Montos informativos YA calculados por el
    API: aquí se pintan como valor, fuera de toda suma. Todo opcional."""

    clave: str | None = None
    folio: str | None = None
    cliente: str | None = None
    matricula: str | None = None
    fecha: str | None = None
    venta_mxn: float | None = None
    venta_usd: float | None = None
    cobrado_mxn: float | None = None
    cobrado_usd: float | None = None
    por_cobrar_mxn: float | None = None
    por_cobrar_usd: float | None = None
    destino: str | None = None  # 'YYYY-MM'
    origen: str | None = None  # 'YYYY-MM'
    leyenda: str | None = None
    # ADITIVOS (16-oct-2026, «aparece que aún no se cobran pero sí están
    # cobrados», caso #351): por qué pasa (`SIN_COBRO` | `COBRO_PARCIAL` |
    # `COBRO_POSTERIOR`), el día del corte con el que se midió «cobrado»
    # ('YYYY-MM-DD') y, si el API lo sabe, lo cobrado DESPUÉS del corte y el
    # día del último cobro. Sin ellos (API ≤ 0.0.90) el bloque es el de
    # siempre, byte a byte.
    motivo: str | None = None
    corte: str | None = None
    cobrado_despues_usd: float | None = None
    cobrado_despues_mxn: float | None = None
    fecha_cobro_despues: str | None = None

    @field_validator(
        "clave", "folio", "cliente", "matricula", "fecha", "destino", "origen", "leyenda",
        "motivo", "corte", "fecha_cobro_despues",
        mode="before",
    )
    @classmethod
    def _textos(cls, v: Any) -> str | None:
        return texto_cierre_liberal(v)

    @field_validator("cobrado_despues_usd", "cobrado_despues_mxn", mode="before")
    @classmethod
    def _montos_aditivos(cls, v: Any) -> float | None:
        return numero_cierre_liberal(v)


class BalanceCierreResumenFila(BaseModel):
    """Renglón por avión de `resumen_cierre` del Balance general (cuántos
    vuelos pasan/vienen y sus montos; el API los suma, aquí solo se pintan)."""

    matricula: str | None = None
    vuelos: int = 0
    venta_mxn: float | None = None
    por_cobrar_mxn: float | None = None

    @field_validator("matricula", mode="before")
    @classmethod
    def _matricula(cls, v: Any) -> str | None:
        return texto_cierre_liberal(v)

    @field_validator("vuelos", mode="before")
    @classmethod
    def _vuelos(cls, v: Any) -> Any:
        return 0 if v is None else v


class BalanceCierreResumen(BaseModel):
    pasan: list[BalanceCierreResumenFila] = Field(default_factory=list)
    vienen: list[BalanceCierreResumenFila] = Field(default_factory=list)

    @field_validator("pasan", "vienen", mode="before")
    @classmethod
    def _listas(cls, v: Any) -> list:
        return lista_cierre_liberal(v)


class BalanceCierre(BaseModel):
    """Bloque raíz `cierre` de los balances (16-oct-2026, API 0.0.88):
    estado del mes («ABIERTO» / «CERRADO» / «REABIERTO»), quién y cuándo lo
    cerró y las tres listas de vuelos: `pasan` (salen de este mes al cerrar:
    NO suman), `vienen` (entraron de un mes anterior: suman completos) y
    `pasarian` (mes abierto: hoy suman, al cerrar pasan). Con `cierre` ausente
    (API ≤ 0.0.87) o con estado ABIERTO y listas vacías no se pinta nada:
    libro byte a byte."""

    disponible: bool | None = None
    periodo: str | None = None
    estado: str | None = None
    cerrado_at: str | None = None
    cerrado_por_nombre: str | None = None
    reabierto_at: str | None = None
    reabierto_por_nombre: str | None = None
    pasan: list[BalanceCierreVuelo] = Field(default_factory=list)
    vienen: list[BalanceCierreVuelo] = Field(default_factory=list)
    pasarian: list[BalanceCierreVuelo] = Field(default_factory=list)

    @field_validator(
        "periodo", "estado", "cerrado_at", "cerrado_por_nombre", "reabierto_at",
        "reabierto_por_nombre", mode="before",
    )
    @classmethod
    def _textos(cls, v: Any) -> str | None:
        return texto_cierre_liberal(v)

    @field_validator("pasan", "vienen", "pasarian", mode="before")
    @classmethod
    def _listas(cls, v: Any) -> list:
        return lista_cierre_liberal(v)

    def tiene_algo(self) -> bool:
        """¿Hay algo que pintar? Listas con filas o un mes CERRADO/REABIERTO
        (su sello se imprime aunque ningún vuelo haya pasado)."""
        estado = (self.estado or "").strip().upper()
        return bool(self.pasan or self.vienen or self.pasarian) or estado in (
            "CERRADO",
            "REABIERTO",
        )


class BalanceAvionRequest(BaseModel):
    generado: str | None = None
    matricula: str = ""
    modelo: str | None = None
    # Color del avión (aeronave.color_calendario) — lo usa el balance
    # GENERAL para teñir su bloque en la hoja "balance".
    avion_color: str | None = None
    periodo_desde: str | None = None
    periodo_hasta: str | None = None
    # Aportación AFAC USD/hr de la ficha del avión (5-oct-2026, ADITIVO: el
    # API ya la mandaba y aquí se ignoraba). Es la constante del encabezado
    # de la hoja maestra con la que la columna PERMISO AFAC (PROVISIÓN) va
    # como FÓRMULA visible (tarifa × TC costos × horas cobradas). None = el
    # avión no la tiene configurada (o el consolidado del general, donde la
    # tarifa es por avión): la columna queda como valor.
    permiso_afac_usd_hr: float | None = None
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
    combustible: BalanceAvionHojaCombustible = Field(default_factory=BalanceAvionHojaCombustible)
    balance: BalanceAvionBalanceBloque = Field(default_factory=BalanceAvionBalanceBloque)
    # Pestaña "Otros movimientos" (28-ago): solo la manda el GENERAL; None =
    # no se pinta (skew tolerante con API viejo).
    otros_movimientos: BalanceHojaOtrosMovimientos | None = None
    pendientes: list[str] = Field(default_factory=list)
    # Nombre de la regla de COMISIONES a cargo del avión (6-oct-2026,
    # ADITIVO): la etiqueta que el API arma con la vigencia CONFIGURADA
    # («regla sep-2026», `etiquetaReglaComisiones`). Con ella las notas de
    # COMISIONES, COBRADO REAL y 'cobranza' nombran esa regla; sin ella (API
    # 0.0.65) dicen «septiembre 2026», como siempre.
    regla_comisiones: str | None = None
    # ADITIVO (12-oct-2026, API 0.0.73, invariante 58 del API): el payload
    # INDIVIDUAL de cada avión trae `sin_provision_vendedor: true` al nivel
    # raíz cuando viene de un API que ya no provisiona el pago al vendedor
    # en 'otros movimientos' del Balance general (la misma señal que
    # `BalanceHojaOtrosMovimientos.sin_provision_vendedor`). Aquí solo
    # cambia el fragmento del pago al vendedor de la nota de COMISIONES de
    # la hoja de vuelos del libro individual
    # (`_MAESTRA_PAGO_VENDEDOR_SIN_PROVISION`); sin ella (None/False, API
    # previo) el libro es el de siempre, byte a byte.
    sin_provision_vendedor: bool | None = None
    # ADITIVO (13-oct-2026, API 0.0.74, invariante 60 del API): la columna
    # FACTURA VUELATOUR trae «Pendiente de factura» / «No lleva factura» /
    # vacío = sin cobros según los cobros y la solicitud del vuelo. Aquí SOLO
    # cambia la nota al pie de la hoja maestra (`_nota_factura_vuelatour`):
    # en el libro individual (payload raíz) y en «reporte horas FLOTA» del
    # general (que la recibe del `consolidado`). None/False/ausente (API
    # previo) ⇒ la nota de hoy, libro byte a byte.
    factura_vuelatour_segun_cobros: bool | None = None
    # ADITIVO (14-oct-2026, API 0.0.77, invariante 60 del API; decisión del
    # cliente «¿Un anticipo devuelto lleva CFDI? eso no»): el dinero devuelto
    # COMPLETO no exige factura salga por donde salga; lo decide el API y la
    # columna FACTURA VUELATOUR llega hecha. Aquí SOLO cambia la nota al pie de
    # la hoja maestra (`_NOTA_FACTURA_VUELATOUR_DEVOLUCION`, que manda sobre
    # `factura_vuelatour_segun_cobros`): libro individual, cada `aviones[]` y
    # el `consolidado`. None/False/ausente (API ≤ 0.0.76) ⇒ las dos notas de
    # siempre, byte a byte.
    factura_vuelatour_devolucion: bool | None = None
    # ADITIVO (13-oct-2026, API 0.0.74, invariante 61 del API): la comisión
    # del vendedor resta SOLO como gasto real, en dos categorías (COMISIONES
    # = parte del avión de la comisión bancaria + los pagos «Comisión del
    # vendedor (avión)»; la «(VuelaTour)» nunca; sin pago, solo la
    # bancaria). Viaja en el libro individual, en cada `aviones[]` del
    # general Y en el `consolidado` FLOTA (a diferencia de
    # `sin_provision_vendedor`; como `factura_vuelatour_segun_cobros`): las
    # notas de COMISIONES de la hoja de vuelos del general salen del
    # consolidado y las de la cascada, de cada avión. Con `true` cambian SOLO
    # textos (encabezado y pie de COMISIONES, notas de la hoja de vuelos del
    # general, cascada); manda sobre `sin_provision_vendedor`. None/False
    # (API ≤ 0.0.73) ⇒ libro byte a byte.
    comision_vendedor_solo_gasto_real: bool | None = None
    # ADITIVO (13-oct-2026, API 0.0.75, invariante «INGRESO ASIGNADO A UN
    # AVIÓN» del API): la hoja «Otros ingresos» del avión — ingresos de
    # RESULTADO registrados en Ingresos con este avión, en el mes de su
    # `fecha`. Su PRESENCIA manda: el libro individual gana la pestaña «Otros
    # ingresos» (entre 'permisos' y 'balance') y la cascada de 'balance' la
    # línea «(+) OTROS INGRESOS USD» (también en cada bloque del Balance
    # general). El API la manda SIEMPRE en los libros de avión (vacía:
    # `{filas: [], total_mxn: 0, usd: 0}`; desde el 0.0.77 con
    # `usd_por_ingreso: true` al final) y en el `consolidado` FLOTA SOLO con
    # filas (hoja «otros ingresos aviones»). None (API ≤ 0.0.74) ⇒ ni
    # pestaña ni línea: libro byte a byte.
    otros_ingresos_avion: BalanceAvionHojaOtrosIngresos | None = None
    # ADITIVO (13-oct-2026, API 0.0.75, punto 5 del lote): `true` SOLO en la
    # petición del libro de la vista en pantalla: el JSON lleva
    # `vinculos_gasto: true` en la raíz (el libro marca las celdas de los
    # gastos). Cualquier otro valor ⇒ None (la raíz de siempre).
    vinculos_gasto: bool | None = None
    # ADITIVO (16-oct-2026, API 0.0.88, invariante 89 del API): el cierre del
    # mes del libro (ver `BalanceCierre`). None (API previo) ⇒ sin bloque
    # «CIERRE DEL MES» bajo TOTALES: libro byte a byte.
    cierre: BalanceCierre | None = None
    # ADITIVO (16-oct-2026, API 0.0.89, invariante 90 del API): los gastos
    # PERMISO de ESTE avión marcados «Pago de AFAC» (los paga VuelaTour: ya NO
    # están en `permisos` ni restan en la cascada). ÚLTIMA llave del libro y
    # SOLO con ≥ 1. El individual pinta una línea en cursiva bajo el título de
    # 'permisos' que lo dice (con el detalle en su nota); el general no la
    # usa (los pagos van en `afac_resumen` de 'otros movimientos'). None ⇒
    # libro byte a byte.
    pagos_afac: BalanceAvionHojaGastos | None = None
    # ADITIVO (18-oct-2026, API 0.0.93): la hoja 'balance' de ESTE libro se
    # pinta EN PESOS (cascada y socios con los `*_mxn` del bloque `balance`;
    # los dólares como columna informativa). Libro individual y cada
    # `aviones[]` del general (el consolidado no la lleva). Sin ella
    # (False/null/ausente) ⇒ libro byte a byte.
    balance_en_pesos: bool = False
    # ADITIVO (19-oct-2026, API 0.0.98, regla B del API): con `true` el API
    # calculó TODO dólar GASTO POR GASTO (cada gasto a USD con su T.C. propio
    # o, sin él, el oficial del día en que se pagó: el criterio del reparto a
    # socios) y ya no es «MXN ÷ T.C. promedio». Aquí ninguna celda deriva
    # entonces un USD como fórmula de pesos ÷ T.C.: el TOTAL USD de las hojas
    # de gastos y combustible y la cadena USD/IVA de la hoja de vuelos van
    # como VALOR del API con nota en el encabezado. Raíz del libro
    # individual y de cada `aviones[]` del general (el consolidado FLOTA
    # puede traerla o no). Solo un `true` JSON la enciende; lo demás ⇒ None
    # (API previo: libro byte a byte).
    usd_por_gasto: bool | None = None

    @field_validator("usd_por_gasto", mode="before")
    @classmethod
    def _usd_por_gasto_liberal(cls, v: Any) -> bool | None:
        return True if v is True else None

    @field_validator("regla_comisiones", mode="before")
    @classmethod
    def _regla_liberal(cls, v: Any) -> str | None:
        return regla_comisiones_liberal(v)

    @field_validator("vinculos_gasto", mode="before")
    @classmethod
    def _vinculos_gasto_liberal(cls, v: Any) -> bool | None:
        return True if v is True else None

    @field_validator("balance_en_pesos", mode="before")
    @classmethod
    def _balance_en_pesos_liberal(cls, v: Any) -> bool:
        return bandera_liberal(v)


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
    # Comisiones del avión (columna COMISIONES MXN, 6-oct-2026): desde la
    # regla de septiembre 2026 (API 0.0.65) = Σ COMISIONES de la hoja de
    # vuelos del libro INDIVIDUAL del avión (parte del avión de la comisión
    # bancaria + provisión de la comisión del vendedor) y GANANCIA ya va
    # después de ellas — el RESUMEN conserva la columna aunque la hoja de
    # vuelos del general ya no la lleve (7-oct-2026). Un API ≤ 0.0.64 manda
    # None/0 (regla A del 28-ago-2026: la comisión del vendedor era solo de
    # VuelaTour).
    comisiones_mxn: float | None = None
    # VENTA − COSTO TOTAL − COMBUSTIBLE = GANANCIA (leyenda impresa).
    ganancia_mxn: float | None = None
    cobrado_mxn: float | None = None
    por_cobrar_mxn: float | None = None
    pendientes: int = 0


class BalanceInventarioItemFila(BaseModel):
    """Fila del bloque POR ÍTEM de la hoja 'inventario' del Balance GENERAL
    (tiendita, 30-ago-2026). EXISTENCIA y VALOR A COSTO son A HOY (todo el
    cardex), no una foto al corte; el resto es del periodo. None = celda
    vacía (sin actividad de ese tipo), nunca un 0 falso.

    Regla de costo (25-sep-2026, API 0.0.36, `regla_costo='ULTIMO_PRECIO'`
    en la hoja): VALOR A COSTO = existencia × ÚLTIMO PRECIO DE COMPRA al
    T.C. oficial de HOY; las compras y ventas en dólares llegan YA en pesos
    con el T.C. oficial de su día. Hasta el 0.0.35 era el costo FIFO de las
    capas vivas. Aquí no cambia nada del render por eso: los números
    llegan hechos."""

    nombre: str = ""  # nombre del ítem (+ ' · nº de parte' cuando lo tiene)
    existencia: float | None = None
    valor_costo_mxn: float | None = None
    # Valor a costo de las capas vivas compradas en USD SIN tipo de cambio
    # (22-sep-2026): son DÓLARES y se muestran como dólares en su propia
    # columna — `valor_costo_mxn` suma únicamente pesos reales. Jamás se
    # convierte aquí (no hay T.C. que usar) ni se suma al total en pesos.
    # None = ese ítem no tiene capas en USD sin T.C. (y un API viejo, que
    # no manda el campo, deja la columna fuera de la hoja).
    valor_costo_usd: float | None = None
    # true cuando alguna capa viva del ítem está en USD sin T.C.
    sin_tc: bool = False
    # ADITIVOS (13-oct-2026, API 0.0.74): COSTO UNITARIO = último precio de
    # compra del ítem. `costo_unitario_mxn` en PESOS (compra en pesos tal
    # cual; compra en dólares al T.C. oficial de hoy, el mismo `tc_hoy` del
    # valorizado); `costo_unitario_usd` SOLO cuando el último precio es en
    # dólares y no hay T.C. (el mismo caso de `valor_costo_usd`/`sin_tc`).
    # Jamás los dos a la vez; aquí jamás se convierte. None/ausente en todas
    # las filas (API viejo) ⇒ la hoja no gana columna.
    costo_unitario_mxn: float | None = None
    costo_unitario_usd: float | None = None
    # ENTRADAs del periodo (compras reales; DEVOLUCION/AJUSTE no son compra).
    compradas_cant: float | None = None
    compradas_costo_mxn: float | None = None
    salidas_cant: float | None = None
    # Σ venta de las salidas CON precio (lo cargado a los aviones) y su
    # utilidad (venta − costo de la pieza). Desde el API 0.0.36 incluye las
    # ventas en dólares CONVERTIDAS por el API con el T.C. oficial del día
    # de la venta (venta y costo con el mismo T.C.).
    vendido_mxn: float | None = None
    utilidad_mxn: float | None = None
    # Utilidad de la TIENDA en DÓLARES (25-sep-2026, API 0.0.35): salidas
    # cobradas al avión en USD sobre costo en USD sin T.C. Antes esa
    # utilidad no se podía expresar en pesos y la celda quedaba vacía;
    # ahora viaja en su moneda. Desde el API 0.0.36 SOLO es el RESPALDO de
    # filas que siguen sin T.C. (tras la migración de datos del 25-sep no
    # queda ninguna ⇒ None y las columnas USD se apagan solas). Jamás se
    # convierte ni se suma con `vendido_mxn`/`utilidad_mxn`: una salida
    # cuenta en UNA sola moneda (`ventaDeSalida` del API). None = sin
    # ventas en dólares sin T.C.
    vendido_usd: float | None = None
    utilidad_usd: float | None = None
    # ADITIVOS (API 0.0.36): el USD ORIGINAL de las ventas dólar-sobre-dólar
    # que ya cuentan en pesos arriba. Dato secundario del panel; esta hoja
    # NO los pinta (sumarlos a VENDIDO/UTILIDAD USD contaría doble).
    vendido_usd_original: float | None = None
    utilidad_usd_original: float | None = None
    # Salidas CON venta cuya utilidad no se puede expresar en ninguna moneda
    # (venta en pesos sobre costo en USD sin T.C.). 0 = ninguna.
    ventas_sin_utilidad: int = 0
    # Matrículas a las que se aplicó en el periodo (únicas, ' + ').
    matriculas: str | None = None

    @field_validator("sin_tc", mode="before")
    @classmethod
    def _sin_tc_bool(cls, v: Any) -> Any:
        """Liberal con lo que manda NestJS: null ⇒ False (nunca un 422 que
        tumbe el libro entero por una bandera ausente)."""
        return False if v is None else v

    @field_validator("ventas_sin_utilidad", mode="before")
    @classmethod
    def _ventas_sin_utilidad_num(cls, v: Any) -> Any:
        return 0 if v is None else v


class TcHoy(BaseModel):
    """T.C. oficial de HOY con el que el API valorizó la bodega (25-sep-2026,
    API 0.0.36): la MISMA fuente que las cotizaciones
    (`TipoCambioService.oficialDetallePara`). Solo se REDACTA en la nota de
    la hoja: aquí jamás se multiplica. `tc` None = sin dato (liberal: un
    null no debe tumbar el libro con un 422)."""

    model_config = ConfigDict(extra="ignore")

    tc: float | None = None
    fecha_dato: str | None = None
    fuente: str | None = None

    @field_validator("tc", mode="before")
    @classmethod
    def _tc_liberal(cls, v: Any) -> Any:
        """Un T.C. ilegible («N/D», booleano, NaN) ⇒ None: la nota no cita
        número. Solo es texto de una nota — jamás un 422 que tumbe el libro
        (revisión adversaria 25-sep-2026)."""
        if v is None or isinstance(v, bool):
            return None
        try:
            tc = float(v)
        except (TypeError, ValueError):
            return None
        return tc if math.isfinite(tc) else None


# Valor de `regla_costo` del API 0.0.36 (25-sep-2026): costo = último precio
# de compra (congelado en cada salida al registrarla) y montos en pesos con
# el T.C. oficial del día. Ausente = API previo (costo FIFO) ⇒ textos de
# siempre, hoja idéntica.
REGLA_COSTO_ULTIMO_PRECIO = "ULTIMO_PRECIO"


class BalanceHojaInventario(BaseModel):
    """Hoja 'inventario' del Balance GENERAL (tiendita): bloque por ítem +
    totales — el bloque 2 (detalle de salidas) sale de `refacciones`."""

    filas: list[BalanceInventarioItemFila] = Field(default_factory=list)
    total_piezas: float | None = None
    total_valor_mxn: float | None = None
    # Σ de `valor_costo_usd` de las filas (22-sep-2026): total APARTE, en
    # dólares. None = API viejo → la hoja sale como siempre, sin la columna.
    total_valor_usd: float | None = None
    total_compras_mxn: float | None = None
    total_vendido_mxn: float | None = None
    total_utilidad_mxn: float | None = None
    # Cuántas filas traen costo en USD sin T.C. (para la nota al pie de la
    # tabla). 0 / ausente = no se pinta la nota.
    filas_sin_tc: int = 0
    # Utilidad de la tienda en DÓLARES (25-sep-2026): Σ de `vendido_usd` /
    # `utilidad_usd` de las filas, APARTE de los totales en pesos. None = API
    # viejo o ninguna venta en dólares ⇒ la hoja sale sin esas columnas.
    total_vendido_usd: float | None = None
    total_utilidad_usd: float | None = None
    # Cuántas filas tienen ventas cuya utilidad no se puede calcular (venta
    # en pesos sobre costo en USD sin T.C.): nota roja bajo la tabla.
    filas_utilidad_incompleta: int = 0
    # Margen de la tienda VIGENTE (% sobre el costo de la pieza —costo FIFO
    # hasta el API 0.0.35, último precio de compra desde el 0.0.36—; config
    # `inventario_margen_venta_pct`, 25 por default). Solo texto de la nota.
    margen_venta_pct: float | None = None
    # ADITIVOS (25-sep-2026, API 0.0.36). `regla_costo` = 'ULTIMO_PRECIO'
    # cambia SOLO los textos (notas, portada): ninguna nota dice «FIFO».
    # `tc_hoy` = T.C. oficial con el que se valorizó (se cita en la nota).
    # None = API previo ⇒ hoja idéntica a la de siempre.
    regla_costo: str | None = None
    tc_hoy: TcHoy | None = None

    @field_validator("filas_sin_tc", "filas_utilidad_incompleta", mode="before")
    @classmethod
    def _conteos_num(cls, v: Any) -> Any:
        return 0 if v is None else v

    @field_validator("tc_hoy", mode="before")
    @classmethod
    def _tc_hoy_liberal(cls, v: Any) -> Any:
        """Liberal con lo que manda NestJS (revisión adversaria 25-sep-2026):
        el API usa el MISMO nombre `tc_hoy` como NÚMERO suelto en otra
        respuesta (PATCH de costo de una entrada: `stats.tc_hoy`) y como
        objeto aquí. Un número ⇒ `{tc}`; algo que no es objeto ⇒ None. Es
        solo el texto de una nota: jamás un 422 que tumbe el Balance general
        entero."""
        if v is None or isinstance(v, dict | TcHoy):
            return v
        if isinstance(v, bool):
            return None
        if isinstance(v, int | float | str):
            return {"tc": v}
        return None


class BalanceEmpresaParticipacion(BaseModel):
    """Participación de la EMPRESA como socio de UN avión (su `monto_usd` es
    el mismo que el del socio `es_empresa` del bloque de ese avión)."""

    matricula: str | None = None
    porcentaje: float | None = None
    monto_usd: float | None = None
    # ADITIVO (18-oct-2026, API 0.0.93): el `monto_mxn` del socio `es_empresa`
    # de ese avión (balances en pesos). None = sin pesos.
    monto_mxn: float | None = None


class BalanceEmpresaBloque(BaseModel):
    """Bloque «VUELATOUR (empresa)» al final de la hoja 'balance' del Balance
    GENERAL (6-oct-2026, API 0.0.59). Todo en USD y YA calculado por el API
    (`balance-empresa.util.ts`): aquí no se recalcula nada, solo se pinta
    (las sumas que están en el libro van como fórmula verificada).
    Todos opcionales: un API que mande el bloque a medias no tumba el
    balance (celda vacía, nunca un 0 falso)."""

    # Socios `es_empresa` de cada avión (monto = utilidad COBRADA × %).
    participaciones: list[BalanceEmpresaParticipacion] = Field(default_factory=list)
    participacion_usd: float | None = None  # Σ anterior (round2)
    # Σ ingreso_mxn / egreso_mxn de 'otros movimientos' convertidos con el TC
    # de SU vuelo (fila suelta: su TC o el oficial del día). NO es caja: el
    # ingreso es lo COTIZADO de los vuelos del periodo no cancelados y el
    # egreso TODO lo que esa hoja resta (pago o provisión al vendedor, TUAs,
    # extensión de horario, comisión bancaria, sueltos) — no solo el vendedor.
    ingresos_propios_usd: float | None = None
    pagos_vendedor_usd: float | None = None
    # = TOTAL USD de la hoja 'otros gastos' (total MXN / TC promedio flota).
    otros_gastos_empresa_usd: float | None = None
    tc_usado: float | None = None  # TC con el que se convirtió lo anterior
    # Utilidad de la hoja 'inventario' / TC promedio; None = sin inventario.
    tienda_utilidad_usd: float | None = None
    resultado_usd: float | None = None
    tc_promedio: float | None = None  # TC promedio de la flota
    nota: str | None = None  # base de cada línea (texto corto del API)
    # BALANCE EN PESOS (18-oct-2026, API 0.0.93), ADITIVAS: el MISMO bloque en
    # pesos reales (Σ INGRESO / EGRESO MXN de 'otros movimientos', TOTAL MXN
    # de 'otros gastos', utilidad MXN del inventario). Solo se pintan con la
    # bandera `balance_en_pesos` del general; None = sin pesos (null se
    # propaga al resultado). `nota_pesos` = base en pesos (texto corto del
    # API); la `nota` de siempre sigue explicando la columna USD.
    participacion_mxn: float | None = None
    ingresos_propios_mxn: float | None = None
    pagos_vendedor_mxn: float | None = None
    otros_gastos_empresa_mxn: float | None = None
    tienda_utilidad_mxn: float | None = None
    resultado_mxn: float | None = None
    nota_pesos: str | None = None

    @field_validator("participaciones", mode="before")
    @classmethod
    def _participaciones_liberal(cls, v: Any) -> Any:
        """NestJS puede mandar null en una lista vacía: jamás un 422."""
        return [] if v is None else v

    @field_validator("nota_pesos", mode="before")
    @classmethod
    def _nota_pesos_liberal(cls, v: Any) -> str | None:
        """Texto informativo: blanco o no-texto ⇒ None (jamás un 422)."""
        return _texto_liberal(v)


# Variantes del libro de flota (6-oct-2026, API 0.0.64). Pedido del cliente:
# «el que ya tenemos que se renombre a "Balance mensual" y el nuevo botón sea
# el "Balance general"». MISMO payload y mismos números en las dos: solo
# cambia la hoja «reporte horas FLOTA» (ver `balance_avion_xlsx._COLS_GENERAL`).
VARIANTE_BALANCE_MENSUAL = "mensual"
VARIANTE_BALANCE_GENERAL = "general"
VARIANTES_BALANCE_GENERAL = (VARIANTE_BALANCE_MENSUAL, VARIANTE_BALANCE_GENERAL)


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
    # Bloque «VUELATOUR (empresa)» al final de la hoja 'balance' (6-oct-2026,
    # API 0.0.59). None = API viejo ⇒ el bloque NO se pinta (libro
    # byte-idéntico al de antes).
    empresa: BalanceEmpresaBloque | None = None
    # Variante del libro (6-oct-2026, API 0.0.64): 'mensual' = el libro de
    # siempre (también sin la llave: un API ≤ 0.0.63 no la manda); 'general'
    # = la hoja «reporte horas FLOTA» resumida a COSTO TOTAL + las 12
    # columnas del costo por hora de la hoja «utilidades» del cliente. Ningún
    # número del payload depende de la variante. Desde el 7-oct-2026 (API
    # 0.0.66) la hoja de vuelos de las DOS va sin la columna COMISIONES (vive
    # en el libro individual).
    variante: Literal["mensual", "general"] = VARIANTE_BALANCE_MENSUAL
    # Nombre de la regla de COMISIONES (6-oct-2026, ADITIVO; ver
    # `BalanceAvionRequest.regla_comisiones`): nota del RESUMEN y, si el
    # `consolidado` no trae la suya, las hojas de la flota.
    regla_comisiones: str | None = None
    # ADITIVO (13-oct-2026, API 0.0.75, punto 5 del lote): como en
    # `BalanceAvionRequest`: solo el libro de la vista en pantalla.
    vinculos_gasto: bool | None = None
    # ADITIVOS (16-oct-2026, API 0.0.88, invariante 89 del API): el cierre
    # del mes del libro (bloque «CIERRE DEL MES» al pie del RESUMEN) y el
    # resumen por avión de lo que pasa/viene. None ⇒ RESUMEN de siempre.
    cierre: BalanceCierre | None = None
    resumen_cierre: BalanceCierreResumen | None = None
    # ADITIVO (18-oct-2026, API 0.0.93): la hoja 'balance' del general (los
    # bloques de cada avión y el de VUELATOUR) se pinta EN PESOS; dólares
    # informativos. La RAÍZ manda el layout de toda la hoja (un avión sin
    # pesos sale con «—»). Sin ella ⇒ libro byte a byte.
    balance_en_pesos: bool = False

    @field_validator("regla_comisiones", mode="before")
    @classmethod
    def _regla_liberal(cls, v: Any) -> str | None:
        return regla_comisiones_liberal(v)

    @field_validator("vinculos_gasto", mode="before")
    @classmethod
    def _vinculos_gasto_liberal(cls, v: Any) -> bool | None:
        return True if v is True else None

    @field_validator("balance_en_pesos", mode="before")
    @classmethod
    def _balance_en_pesos_liberal(cls, v: Any) -> bool:
        return bandera_liberal(v)

    @field_validator("variante", mode="before")
    @classmethod
    def _variante_liberal(cls, v: Any) -> Any:
        """Liberal con lo que manda NestJS: mayúsculas y espacios se
        normalizan; null o un valor desconocido = el libro de siempre (el API
        valida `modo` con 400) — jamás un 422 que tumbe la descarga."""
        if isinstance(v, str):
            texto = v.strip().lower()
            if texto in VARIANTES_BALANCE_GENERAL:
                return texto
        return VARIANTE_BALANCE_MENSUAL


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

    # Con ID de vuelo por cliente (API 0.0.75): 'beh05-011026' (+ ' (50 %
    # N4142R)' en multi-avión; un vuelo de servicio siempre 'vtservicio');
    # sin él, la clave de siempre ('vtbeh').
    clave: str = ""
    # ID de vuelo por cliente del API 0.0.75 (decisión 6 del lote del
    # 13-oct-2026): NO se pinta (la CLAVE ya llega armada); solo decide el
    # ancho de la columna A (`_formato.ancho_columna_clave`); None = API previo
    # o migración sin aplicar ⇒ libro byte a byte el de antes.
    clave_cliente: str | None = None
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
    # «FACTURA VUELATOUR» del Libro Dinero: la misma etiqueta del API que en
    # los balances; desde el 13-oct-2026 (invariante 60 del API) puede ser
    # «Pendiente de factura» / «No lleva factura» (texto tal cual; None =
    # sin cobros).
    factura_vuelatour: str | None = None
    # Total COMPLETO facturado al cliente (con TUAs/extras/pernocta + su IVA):
    # columna informativa junto al STATUS DE COBROS (regla 28-ago-2026).
    total_cliente_usd: float | None = None
    total_cliente_mxn: float | None = None
    # Regla B (28-ago-2026): vuelo MULTI-AVIÓN — fracción de la venta del
    # avión que toca a esta matrícula (la CLAVE ya trae el sufijo del API).
    participacion: float | None = None
    multi_avion: bool | None = None
    # CIERRE DE MES (16-oct-2026, API 0.0.88, invariante 89 del API): marca
    # YA armada por el API («Viene de septiembre 2026» / «Pendiente de cobro:
    # pasará a octubre 2026 al cerrar»), se pinta junto a la CLAVE. Un vuelo
    # que PASÓ al mes siguiente no viene en `vuelos` del mes de origen (el
    # Libro es de dinero): va en `DineroXlsxRequest.cierre.pasan`. None (API
    # ≤ 0.0.87 o vuelo normal) ⇒ la fila de siempre.
    cierre_marca: str | None = None
    diferido_a: str | None = None
    diferido_de: str | None = None

    @field_validator("clave_cliente", mode="before")
    @classmethod
    def _clave_cliente_liberal(cls, v: Any) -> str | None:
        return clave_cliente_liberal(v)

    @field_validator("cierre_marca", "diferido_a", "diferido_de", mode="before")
    @classmethod
    def _texto_cierre_liberal(cls, v: Any) -> str | None:
        return texto_cierre_liberal(v)


class DineroOtroIngresoFila(BaseModel):
    """Fila de la hoja 'Otros ingresos' (TUAs/extras/pernocta y comisión del
    vendedor por vuelo — ingreso de VuelaTour, no del avión; desde 28-ago
    también líneas 'iva de tuas/extras' con el mismo shape y el pago de la
    comisión al vendedor como egreso apareado: hasta el API 0.0.73 la
    PROVISIÓN mientras no hubiera gasto real; desde el 0.0.74 solo el gasto
    real, sin él la fila no lleva egreso)."""

    # Fila de un vuelo con ID de vuelo por cliente (API 0.0.75): la misma
    # clave de 'dinero-vlos' ('beh05-011026'); ING-n sin cambio.
    clave: str = ""
    # ID de vuelo por cliente del API 0.0.75 (decisión 6 del lote del
    # 13-oct-2026): NO se pinta (la CLAVE ya llega armada); solo decide el
    # ancho de la columna A (`_formato.ancho_columna_clave`); None = API previo
    # o migración sin aplicar ⇒ libro byte a byte el de antes.
    clave_cliente: str | None = None
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
    # «factura vuelatour» de 'Otros ingresos': misma etiqueta que la hoja 1
    # (invariante 60 del API: «Pendiente de factura» / «No lleva factura»).
    factura: str | None = None
    # ADITIVOS (13-oct-2026, API 0.0.75, invariante «INGRESO ASIGNADO A UN
    # AVIÓN» del API): SOLO en las filas de `otros_ingresos_avion` (ingresos
    # ING-n asignados a un avión). La CLAVE se tiñe con `avion_color`; la
    # matrícula ya viene dentro del concepto del API.
    matricula: str | None = None
    avion_color: str | None = None

    @field_validator("clave_cliente", mode="before")
    @classmethod
    def _clave_cliente_liberal(cls, v: Any) -> str | None:
        return clave_cliente_liberal(v)


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
    # ADITIVO (13-oct-2026, API 0.0.75, invariante «INGRESO ASIGNADO A UN
    # AVIÓN» del API): Σ NETO MXN (remanente) de los ingresos ING-n
    # asignados a este avión. El API la manda SOLO con dato; si ALGÚN avión
    # la trae, la hoja 'utilidades' gana la columna «OTROS INGRESOS
    # <matrícula>» por avión (vacía en los que no la traen). None en todos
    # (API ≤ 0.0.74 o sin ingresos de avión) ⇒ la hoja de siempre.
    otros_ingresos_mxn: float | None = None


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
    # ADITIVO (13-oct-2026, API 0.0.75, invariante «INGRESO ASIGNADO A UN
    # AVIÓN» del API): los ingresos ING-n ASIGNADOS a un avión (la misma
    # fila de `filaLibroDeIngreso` + `matricula`/`avion_color`). Cuentan en
    # el balance y el reparto de SU avión: se pintan al final de 'Otros
    # ingresos', en una sección APARTE y FUERA de TOTALES (los TOTALES y
    # `utilidades_otros_ingresos_mxn` ya los excluyen en el API). El API la
    # manda SOLO con filas; vacía ⇒ la hoja de siempre, byte a byte.
    otros_ingresos_avion: list[DineroOtroIngresoFila] = Field(default_factory=list)
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
    # Desde el 28-sep-2026 son SOLO las provisiones VIVAS (vuelos sin gasto
    # real «Comisión del vendedor»); con el API 0.0.74 (invariante 61) viaja
    # SIEMPRE en 0: ya no hay provisión.
    utilidades_comision_vendedor_provisionada_mxn: float | None = None
    # Pagado al vendedor con GASTO REAL (categoría «Comisión del vendedor»
    # ligada al vuelo; 28-sep-2026, API 0.0.39, invariante 31 del API), ya
    # restado de utilidades_otros_ingresos_mxn en lugar de la provisión de
    # esos vuelos (nunca los dos). El API lo manda SOLO si algún vuelo del
    # periodo tiene pago real: None ⇒ leyendas de siempre (byte-idéntico).
    # Desde el API 0.0.74 (dos categorías) es el Σ de las DOS.
    utilidades_comision_vendedor_pagada_mxn: float | None = None
    # ADITIVO (13-oct-2026, API 0.0.74, invariante 61 del API, dos
    # categorías): de lo pagado, la parte capturada como «Comisión del
    # vendedor (avión)» en vuelos que carga un avión (MXN). INFORMATIVA: ya va
    # dentro de lo pagado y de utilidades_otros_ingresos_mxn (este libro es
    # la caja de VuelaTour y resta las dos); es la diferencia esperada contra
    # el bloque VUELATOUR del Balance general. El API la manda SOLO > 0 y con
    # la bandera; la fila 9 la nombra (`_UTILIDADES_NOTA_COMISION_AVION`).
    utilidades_comision_vendedor_avion_mxn: float | None = None
    utilidades_otros_gastos_mxn: float | None = None
    utilidades_tc: float | None = None
    utilidades_aviones: list[DineroUtilidadAvion] = Field(default_factory=list)
    # ADITIVO (13-oct-2026, API 0.0.74, invariante 61 del API): sin pago
    # capturado la fila de comisión de 'Otros ingresos' ya NO lleva egreso
    # (sin provisión) y utilidades solo resta lo pagado de verdad, en las DOS
    # categorías («(VuelaTour)» y «(avión)»: la caja es de VuelaTour). Con
    # `true` cambian SOLO tres leyendas (fila 2 de 'Otros ingresos' y filas
    # 6 y 9 de 'utilidades'), mandando sobre
    # `utilidades_comision_vendedor_pagada_mxn`; None/False (API ≤ 0.0.73) ⇒
    # libro byte a byte.
    comision_vendedor_solo_gasto_real: bool | None = None
    # ADITIVO (16-oct-2026, API 0.0.88, invariante 89 del API): el cierre del
    # mes (mismo bloque que los balances, `BalanceCierre`): líneas «CIERRE
    # DEL MES» al final de 'dinero-vlos'. None (API previo) ⇒ libro byte a
    # byte.
    cierre: BalanceCierre | None = None

    @field_validator("otros_ingresos_avion", mode="before")
    @classmethod
    def _otros_ingresos_avion_liberal(cls, v: Any) -> Any:
        """NestJS puede mandar null en una lista vacía: jamás un 422."""
        return [] if v is None else v


# ================= Caja chica: Excel de la REPOSICIÓN (24-sep-2026) =================
#
# Pedido del cliente: «al momento de reembolsar la caja de cada uno, me puede
# arrojar un Excel descargable con la información de lo que estoy
# reembolsando». El API manda TODO calculado (qué gastos repone la
# reposición, saldos por fila del historial, totales y diferencia salen de
# `caja-chica-saldo.util.ts`); aquí SOLO se pinta. Todo con default y
# extra="ignore": skew tolerante en ambos sentidos.


class CajaChicaDato(BaseModel):
    """Par etiqueta → valor del encabezado o del bloque de totales."""

    model_config = ConfigDict(extra="ignore")

    etiqueta: str = ""
    # float = monto (formato "$"#,##0.00); str = texto tal cual.
    valor: float | str | None = None
    # Texto a la derecha del valor (p. ej. «Quedó pendiente por reponer»).
    nota: str | None = None
    # true = renglón protagonista (negritas + fondo): «POR REPONER HOY».
    destacado: bool = False


class CajaChicaFila(BaseModel):
    """Una entrada del libro dentro del periodo de la reposición."""

    model_config = ConfigDict(extra="ignore")

    # Fecha de pared YYYY-MM-DD (se escribe como FECHA de Excel dd/mm/aaaa).
    fecha: str | None = None
    tipo: str = ""
    categoria: str = ""
    descripcion: str = ""
    vuelo: str = ""
    matricula: str = ""
    # Monto del GASTO (positivo); None en reintegros/ajustes.
    gasto: float | None = None
    # Reintegro / ajuste con signo; None en gastos.
    otro: float | None = None
    comprobante: str = ""
    facturacion: str = ""
    capturo: str = ""
    # Texto ya formateado por el API («2026-09-01 07:00», hora Cancún).
    capturado: str = ""
    saldo: float | None = None
    por_reponer: float | None = None
    # Capturado DESPUÉS de registrar la reposición: fecha y captura en naranja.
    resaltar: bool = False


class CajaChicaReposicionRequest(BaseModel):
    """Excel de una reposición de caja chica (o de lo pendiente por reponer)."""

    model_config = ConfigDict(extra="ignore")

    titulo: str = "Reposición de caja chica"
    subtitulo: str | None = None
    # Nombre de la pestaña (≤ 31, sin \ / * ? : [ ]; el render lo sanea).
    hoja: str | None = None
    encabezado_titulo: str | None = None
    encabezado: list[CajaChicaDato] = Field(default_factory=list)
    filas: list[CajaChicaFila] = Field(default_factory=list)
    # Fila TOTAL de la tabla (Σ columna Gasto y Σ columna Reintegro/ajuste).
    n_gastos: int = 0
    total_gastos: float | None = None
    total_otros: float | None = None
    totales_titulo: str | None = None
    totales: list[CajaChicaDato] = Field(default_factory=list)
    avisos: list[str] = Field(default_factory=list)
    # Texto cuando no hay filas (en vez de una tabla vacía muda).
    sin_filas: str | None = None
