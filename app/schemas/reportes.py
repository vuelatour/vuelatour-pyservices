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
    # Regla 28-ago-2026 (informativos): del total, cuánto es VENTA DEL AVIÓN
    # (tiempo + ajuste + IVA proporcional) y cuánto ingreso de VuelaTour
    # (TUAs/extras/pernocta + su IVA). None = API viejo → no se pintan.
    venta_avion_usd: float | None = None
    otros_ingresos_vuelatour_usd: float | None = None
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
    # cobradas) + ajuste/descuento + su IVA proporcional. Los TUAs, extras y
    # viáticos de pernocta cobrados NO están aquí: son ingreso de VuelaTour
    # (pestaña "otros movimientos" del Balance general).
    total_usd: float | None = None
    iva_usd: float | None = None
    tc_venta: float | None = None
    # true = la cotización no traía TC: se usó el oficial (Banxico FIX/DOF)
    # del día de la cotización — el Excel pinta las celdas en azul claro.
    tc_venta_oficial: bool = False
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
    # Parte VuelaTour del vuelo (TUAs/extras/pernocta + su IVA).
    otros_ingresos_usd: float | None = None
    # TUA pagado del vuelo: SOLO informativo (nota en OPERACIONES); no resta.
    tua_pagado_mxn: float | None = None


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
    comision_vendedor_mxn: float | None = None
    ganancia_mxn: float | None = None
    ganancia_usd: float | None = None
    cobrado_mxn: float | None = None
    por_cobrar_mxn: float | None = None
    por_cobrar_usd: float | None = None
    tc_promedio: float | None = None  # AVERAGE de los TC de costos no nulos
    costo_hr_prom_usd: float | None = None  # AVERAGE de costo/hr USD no nulos
    # TUAs/extras/pernocta cobrados (+ su IVA) EXCLUIDOS de las filas: ingreso
    # de VuelaTour (pestaña "otros movimientos" del Balance general).
    otros_ingresos_usd: float | None = None
    # --- Regla 28-ago-2026 (opcionales) ---
    cobrado_real_mxn: float | None = None  # depósitos reales del periodo
    total_cotizacion_mxn: float | None = None  # total completo (c/extras)
    tua_pagado_mxn: float | None = None  # solo nota; no resta en el libro
    comision_banco_mxn: float | None = None  # Σ comisiones retenidas por bancos


class BalanceAvionGastoFila(BaseModel):
    """Partida del ledger de gastos (indirectos / otros / permisos)."""

    fecha: str | None = None
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
    solo-egreso); viene YA calculada del API — aquí jamás se recalcula."""

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
    # Comisiones de vendedor (la ganancia ya las netea).
    comisiones_mxn: float | None = None
    # VENTA − COSTO − COMBUSTIBLE − COMISIONES = GANANCIA (leyenda impresa).
    ganancia_mxn: float | None = None
    cobrado_mxn: float | None = None
    por_cobrar_mxn: float | None = None
    pendientes: int = 0


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


class BitacoraTacoFila(BaseModel):
    """Una fila = UN vuelo: fecha, tacómetro inicial→final, horas y ruta.

    Mismo renglón que el equipo llenaba a mano en la plantilla Excel
    ("Imprimir planeador" / "MOTOR - HÉLICE") para recortar y pegar en la
    bitácora física. Los tiempos de hélice solo vienen en formato bimotor.
    """

    fecha: str  # ISO (date o datetime); se formatea dd-mmm en hora Cancún
    taco_inicial: float
    horas: float
    taco_final: float
    ruta: str  # "cun-pps-cun" (minúsculas, guiones)
    helice_inicial: float | None = None
    helice_final: float | None = None


class BitacoraTacoRequest(BaseModel):
    """Tira imprimible de bitácora de tacómetros.

    formato PLANEADOR (monomotor): Fecha | Taco inicial | Horas | Taco final
    | Ruta. formato MOTOR_HELICE (bimotor): agrega Tiempo hélice inicial y
    final junto a cada tacómetro (offset que arrastra el equipo en su hoja).
    """

    matricula: str = ""
    modelo: str | None = None
    formato: str = "PLANEADOR"  # PLANEADOR | MOTOR_HELICE
    desde: str | None = None
    hasta: str | None = None
    generado: str | None = None
    filas: list[BitacoraTacoFila] = Field(default_factory=list)


class DineroCobroPago(BaseModel):
    """Parcialidad de cobro (o pago a proveedor) de la hoja dinero-vlos."""

    fecha: str | None = None
    monto_mxn: float | None = None


class DineroVueloFila(BaseModel):
    """Una fila de la hoja 'dinero-vlos' (un vuelo del periodo).

    Las columnas SIN regla todavía (costo proveedor, comisiones, pagos)
    viajan en None y se pintan vacías conservando su columna del libro.
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
    # — sin TUAs/extras/pernocta (esos van en la hoja "Otros ingresos").
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


class DineroOtroIngresoFila(BaseModel):
    """Fila de la hoja 'Otros ingresos' (TUAs/extras/pernocta por vuelo —
    ingreso de VuelaTour, no del avión; desde 28-ago también líneas
    'iva de tuas/extras' con el mismo shape)."""

    clave: str = ""
    fecha_vuelo: str | None = None
    concepto_egreso: str | None = None
    egreso_mxn: float | None = None
    fecha_egreso: str | None = None
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
    utilidades_otros_gastos_mxn: float | None = None
    utilidades_tc: float | None = None
    utilidades_aviones: list[DineroUtilidadAvion] = Field(default_factory=list)
