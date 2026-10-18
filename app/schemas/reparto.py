"""Esquema del payload para el PDF de reparto de utilidades (doc 5.9)."""

from typing import Any

from pydantic import BaseModel, Field, field_validator

from app.schemas.reportes import (
    lista_cierre_liberal,
    numero_cierre_liberal,
    regla_comisiones_liberal,
    texto_cierre_liberal,
)


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
    # del vuelo) el avión absorbe lo que se le paga al vendedor dentro de
    # `RepartoAvion.comisiones_venta_usd`: hasta el API 0.0.73 su PROVISIÓN;
    # con el 0.0.74 (bandera `comision_vendedor_solo_gasto_real`) el pago
    # capturado como «Comisión del vendedor (avión)» (el de «(VuelaTour)» lo
    # paga VuelaTour). Este monto sigue siendo solo informativo.
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
    # CIERRE DE MES (16-oct-2026, API 0.0.88): texto YA armado por el API
    # para un vuelo que VINO de un mes anterior («Viene de septiembre 2026»)
    # o que, en un mes abierto, pasará al cerrar («Pendiente de cobro: pasará
    # a octubre 2026 al cerrar»). Los que SALEN no vienen en `vuelos` (no
    # cuentan): van en `RepartoCierre.diferidos_salen`. Se imprime junto al
    # folio; None ⇒ el folio de siempre.
    cierre_marca: str | None = None
    diferido_de: str | None = None
    diferido_a: str | None = None

    @field_validator("cierre_marca", "diferido_de", "diferido_a", mode="before")
    @classmethod
    def _texto_cierre_liberal(cls, v: Any) -> str | None:
        return texto_cierre_liberal(v)


class RepartoCierreVuelo(BaseModel):
    """Un vuelo de las listas del cierre del reparto (16-oct-2026, API
    0.0.88): SALE del mes (`destino`; pendiente de cobro al cierre: NO se
    reparte), ENTRA de un mes anterior (`origen`) o, en un mes abierto,
    PASARÍA al cerrar. Montos en USD ya calculados por el API; informativos."""

    id: str | None = None
    folio: str | None = None
    cliente: str | None = None
    matricula: str | None = None
    venta_usd: float | None = None
    cobrado_usd: float | None = None
    por_cobrar_usd: float | None = None
    destino: str | None = None
    origen: str | None = None
    leyenda: str | None = None
    # ADITIVOS (16-oct-2026, caso #351): por qué pasa, lo cobrado DESPUÉS del
    # corte (USD) y el día del último cobro. Un vuelo que sale por la
    # bitácora del cierre trae el motivo pero no lo cobrado después (0).
    motivo: str | None = None
    cobrado_posterior_usd: float | None = None
    fecha_cobro_posterior: str | None = None

    @field_validator(
        "id", "folio", "cliente", "matricula", "destino", "origen", "leyenda",
        "motivo", "fecha_cobro_posterior", mode="before",
    )
    @classmethod
    def _textos(cls, v: Any) -> str | None:
        return texto_cierre_liberal(v)

    @field_validator("cobrado_posterior_usd", mode="before")
    @classmethod
    def _monto_posterior(cls, v: Any) -> float | None:
        return numero_cierre_liberal(v)


class RepartoCierre(BaseModel):
    """Bloque raíz `cierre` de `GET /v1/profit-sharing` (16-oct-2026, API
    0.0.88): estado del mes, sello del cierre y las listas de vuelos
    diferidos. `coincide_con_snapshot` False = alguien movió porcentajes o un
    T.C. después de cerrar (el API lo detecta; aquí solo se avisa)."""

    disponible: bool | None = None
    periodo: str | None = None
    estado: str | None = None
    cerrado_at: str | None = None
    cerrado_por_nombre: str | None = None
    reabierto_at: str | None = None
    reabierto_por_nombre: str | None = None
    version: int | None = None
    coincide_con_snapshot: bool | None = None
    diferidos_salen: list[RepartoCierreVuelo] = Field(default_factory=list)
    diferidos_entran: list[RepartoCierreVuelo] = Field(default_factory=list)
    pendientes_que_pasarian: list[RepartoCierreVuelo] = Field(default_factory=list)

    @field_validator(
        "periodo", "estado", "cerrado_at", "cerrado_por_nombre", "reabierto_at",
        "reabierto_por_nombre", mode="before",
    )
    @classmethod
    def _textos(cls, v: Any) -> str | None:
        return texto_cierre_liberal(v)

    @field_validator(
        "diferidos_salen", "diferidos_entran", "pendientes_que_pasarian", mode="before"
    )
    @classmethod
    def _listas(cls, v: Any) -> list:
        return lista_cierre_liberal(v)

    @field_validator("version", mode="before")
    @classmethod
    def _version(cls, v: Any) -> Any:
        return v if isinstance(v, int) and not isinstance(v, bool) else None

    @field_validator("coincide_con_snapshot", mode="before")
    @classmethod
    def _coincide(cls, v: Any) -> bool | None:
        return v if isinstance(v, bool) else None

    def tiene_algo(self) -> bool:
        estado = (self.estado or "").strip().upper()
        return bool(
            self.diferidos_salen or self.diferidos_entran or self.pendientes_que_pasarian
        ) or estado in ("CERRADO", "REABIERTO")


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


class RepartoOtroIngresoLinea(BaseModel):
    """Un ingreso registrado en Ingresos (clave ING-n) y ASIGNADO a este avión
    (13-oct-2026, API 0.0.75, invariante «INGRESO ASIGNADO A UN AVIÓN» del
    API): reembolso de su servicio, venta de una pieza suya, intereses…
    `usd` = lo que SUMÓ al saldo (neto de la comisión bancaria; MXN con su
    T.C. propio o el oficial del día del ingreso); None = MXN sin ningún
    T.C. (no entró al saldo). Solo detalle informativo: el total que suma es
    `RepartoAvion.otros_ingresos_avion_usd`."""

    clave: str = ""
    fecha: str | None = None
    categoria: str | None = None  # etiqueta amable
    concepto: str | None = None  # descripción · cliente|pagador
    monto: float | None = None  # BRUTO en su moneda
    moneda: str | None = None
    comision: float | None = None
    usd: float | None = None

    @field_validator("clave", mode="before")
    @classmethod
    def _clave_liberal(cls, v: Any) -> Any:
        """Texto de una celda: null ⇒ '' (jamás un 422 por una etiqueta)."""
        if v is None:
            return ""
        if isinstance(v, int | float) and not isinstance(v, bool):
            return str(v)
        return v


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
    # BANCARIA de sus cobros + lo del vendedor (fuente única
    # `comisionesDelVuelo` del API): hasta el 0.0.73 su PROVISIÓN; con la
    # bandera `comision_vendedor_solo_gasto_real` (API 0.0.74) los pagos
    # capturados como «Comisión del vendedor (avión)», en CUALQUIER fecha (un
    # mes anterior a la regla también puede traerlos). ≠ 0 ⇒ el PDF y el
    # Excel la pintan como «Comisiones (banco + vendedor)» con los textos de
    # esa regla (`reparto_comisiones.comisiones_al_avion`); 0 ⇒ los textos de
    # siempre.
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
    # ADITIVOS (13-oct-2026, API 0.0.75, invariante «INGRESO ASIGNADO A UN
    # AVIÓN» del API): otros ingresos registrados en Ingresos con ESTE avión,
    # netos de comisión bancaria. `otros_ingresos_avion_usd` SUMA al saldo
    # (`saldo_usd` del API YA lo incluye: aquí jamás se vuelve a sumar);
    # `otros_ingresos_avion_sin_tc_mxn` = los MXN sin ningún T.C. (NO
    # entraron al saldo: se avisa); `otros_ingresos_avion` = el detalle por
    # ingreso. Defaults (API ≤ 0.0.74) ⇒ el PDF y el Excel de siempre.
    otros_ingresos_avion_usd: float = 0.0
    otros_ingresos_avion_sin_tc_mxn: float = 0.0
    otros_ingresos_avion: list[RepartoOtroIngresoLinea] = Field(default_factory=list)

    @field_validator(
        "otros_ingresos_avion_usd", "otros_ingresos_avion_sin_tc_mxn", mode="before"
    )
    @classmethod
    def _monto_o_cero(cls, v: Any) -> Any:
        """null ⇒ 0 (el API manda número; jamás un 422 por un null)."""
        return 0.0 if v is None else v

    @field_validator("otros_ingresos_avion", mode="before")
    @classmethod
    def _lineas_liberal(cls, v: Any) -> Any:
        """NestJS puede mandar null en una lista vacía: jamás un 422."""
        return [] if v is None else v


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
    # ADITIVO (13-oct-2026, API 0.0.74, invariante 61 del API): la comisión
    # del vendedor resta SOLO como gasto real, en dos categorías
    # (`comisiones_venta_usd` = parte del avión de la bancaria + los pagos
    # «Comisión del vendedor (avión)»; los de «(VuelaTour)» los paga
    # VuelaTour). Con `true` cambian SOLO textos (`reparto_comisiones`);
    # None/False (API ≤ 0.0.73) ⇒ los documentos de siempre.
    comision_vendedor_solo_gasto_real: bool | None = None
    # ADITIVO (16-oct-2026, API 0.0.88, invariante 89 del API): el cierre del
    # mes del reparto (ver `RepartoCierre`). None (API previo) ⇒ el PDF y el
    # Excel de siempre, byte a byte.
    cierre: RepartoCierre | None = None

    @field_validator("regla_comisiones", mode="before")
    @classmethod
    def _regla_liberal(cls, v: Any) -> str | None:
        return regla_comisiones_liberal(v)
