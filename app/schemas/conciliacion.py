from typing import Literal

from pydantic import BaseModel, Field, field_validator

from app.schemas.uso_ia import UsoIA


class MapeoColumnasPaywise(BaseModel):
    """Mapeo MANUAL de columnas del estado de cuenta de Paywise (9-sep-2026).

    Respaldo cuando la detección por encabezados no reconoce el archivo:
    cada valor es el NOMBRE de la columna tal como viene en el archivo
    (`ConciliacionParseResponse.columnas` los lista). Con mapeo, el archivo
    se lee como formato `paywise` (abono NETO con bruto y comisión).
    """

    fecha: str = Field(description="Columna de fecha (operación/pago)")
    bruto: str | None = Field(default=None, description="Columna del monto BRUTO")
    comision: str | None = Field(default=None, description="Columna de la comisión")
    neto: str | None = Field(default=None, description="Columna del NETO depositado")
    referencia: str | None = Field(default=None, description="Columna de referencia/ID")
    descripcion: str | None = Field(default=None, description="Columna de descripción")
    estatus: str | None = Field(default=None, description="Columna de estatus")


class ConciliacionParseRequest(BaseModel):
    """Estado de cuenta a parsear. CSV/Excel (preferido) o PDF, en base64."""

    filename: str = Field(description="Nombre del archivo (define el parser por extensión)")
    file_base64: str = Field(description="Contenido en base64 (sin prefijo data:)")
    # ADITIVO (Paywise): fuerza el parser Paywise con estas columnas.
    mapeo: MapeoColumnasPaywise | None = Field(
        default=None, description="Mapeo manual de columnas (Paywise); None = detección"
    )
    # ADITIVOS (15-sep-2026): contexto de la cuenta para el parse del PDF.
    banco: str | None = Field(
        default=None,
        description="Banco de la cuenta (Scotiabank/BBVA/Banorte/Santander…): afina el prompt",
    )
    cuenta_moneda: Literal["MXN", "USD"] | None = Field(
        default=None, description="Moneda de la cuenta (para avisar de cargos en otra moneda)"
    )


class MovimientoParseado(BaseModel):
    fecha: str | None = Field(default=None, description="Fecha YYYY-MM-DD")
    descripcion: str | None = Field(default=None)
    monto: float = Field(description="Monto positivo (pasarela: NETO depositado)")
    tipo: Literal["CARGO", "ABONO"] = Field(description="CARGO = salida, ABONO = entrada")
    referencia: str | None = Field(default=None)
    # ADITIVOS (Paywise, 9-sep-2026): solo el formato `paywise` los llena.
    monto_bruto: float | None = Field(default=None, description="Bruto cobrado al cliente")
    comision: float | None = Field(default=None, description="Comisión retenida")
    estatus: str | None = Field(default=None, description="Estatus del archivo (aprobado…)")
    # ADITIVO (15-sep-2026): saldo corrido impreso en el estado de cuenta. Es
    # el detector barato de líneas omitidas (saldo[i] = saldo[i−1] − cargo +
    # abono); la columna `movimiento_bancario.saldo_posterior` ya existe.
    saldo_posterior: float | None = Field(
        default=None,
        description="Saldo de la cuenta DESPUÉS de este movimiento, si el archivo lo trae",
    )


class ConciliacionParseResponse(BaseModel):
    movimientos: list[MovimientoParseado] = Field(default_factory=list)
    total: int = Field(default=0)
    formato: str = Field(description="csv | excel | pdf | paywise")
    notas: str = Field(default="")
    modelo: str | None = Field(default=None, description="Modelo de Claude si se usó (PDF)")
    uso_ia: UsoIA | None = Field(
        default=None, description="Consumo de tokens (solo PDF; CSV/Excel = None)"
    )
    # ADITIVO: encabezados del archivo tabular, para el mapeo manual del panel.
    columnas: list[str] = Field(default_factory=list)
    # ADITIVO (15-sep-2026): problemas detectados que NO impiden importar pero
    # que el operador debe ver (cadena de saldos rota = faltan movimientos,
    # fechas fuera del periodo, montos en otra moneda…).
    advertencias: list[str] = Field(default_factory=list)


class MovimientoSinConciliar(BaseModel):
    """Movimiento bancario que aún no tiene gasto vinculado."""

    fecha: str | None = Field(default=None, description="Fecha del movimiento YYYY-MM-DD")
    monto: float = Field(description="Monto del movimiento (positivo)")
    descripcion: str | None = Field(default=None, description="Descripción/concepto del banco")
    # ADITIVOS (15-sep-2026): contexto que el modelo NO tenía y que decide el
    # match (la terminación de tarjeta vive en la referencia del banco).
    tipo: Literal["CARGO", "ABONO"] | None = Field(
        default=None, description="CARGO = salida (paga un gasto), ABONO = entrada"
    )
    referencia: str | None = Field(
        default=None, description="Referencia del banco (suele traer la terminación de la tarjeta)"
    )
    cuenta_alias: str | None = Field(default=None, description="Nombre de la cuenta bancaria")
    cuenta_moneda: str | None = Field(default=None, description="Moneda de la cuenta (MXN/USD)")
    terminacion_tarjeta_detectada: str | None = Field(
        default=None,
        description=(
            "Terminación (4 dígitos) que el API dedujo de la referencia, si empató"
            " con una tarjeta"
        ),
    )


class GastoCandidato(BaseModel):
    id: str = Field(description="Identificador del gasto")
    fecha: str | None = Field(default=None, description="Fecha del gasto YYYY-MM-DD")
    monto: float = Field(description="Monto del gasto (positivo)")
    proveedor: str | None = Field(default=None, description="Proveedor/comercio del gasto")
    # ADITIVOS (15-sep-2026): el API ya tiene todo esto; antes se tiraba al
    # serializar y el modelo elegía a ciegas.
    moneda: str | None = Field(default=None, description="Moneda del gasto (MXN/USD)")
    medio_pago: str | None = Field(default=None, description="TARJETA_CORP/TRANSFERENCIA/…")
    tarjeta_terminacion: str | None = Field(
        default=None, description="Terminación de la tarjeta con la que se pagó"
    )
    categoria: str | None = Field(default=None, description="Categoría del gasto (GAS, TUAS…)")
    lugar: str | None = Field(default=None, description="Lugar/aeropuerto del gasto")
    nota: str | None = Field(
        default=None, description="Primera línea de las notas (suele nombrar al aeropuerto)"
    )
    matricula: str | None = Field(default=None, description="Matrícula del avión del gasto")
    # El API manda el folio como NÚMERO (`Number(vuelo.folio)`): pydantic v2
    # NO convierte int → str y la llamada entera reventaba con 422 («Input
    # should be a valid string») en cuanto un candidato tenía vuelo. Se acepta
    # int o str y se normaliza a texto — el prompt solo lo lee.
    vuelo_folio: str | int | None = Field(default=None, description="Folio del vuelo ligado")

    @field_validator("vuelo_folio", mode="before")
    @classmethod
    def _folio_a_texto(cls, v):
        return None if v is None else str(v)
    capturado_por: str | None = Field(default=None, description="Quién capturó el gasto")
    monto_vinculado: float | None = Field(
        default=None, description="Suma ya conciliada de este gasto (pagos parciales)"
    )
    faltante: float | None = Field(
        default=None, description="Lo que falta por cubrir del gasto (monto − vinculado)"
    )
    tc_implicito: float | None = Field(
        default=None, description="TC que resultaría si el cargo MXN pagara este gasto USD"
    )


class ConciliacionSugerirRequest(BaseModel):
    """Movimiento sin conciliar + gastos candidatos cercanos para que Claude proponga el match."""

    movimiento: MovimientoSinConciliar
    candidatos: list[GastoCandidato] = Field(default_factory=list)


class SugerenciaAlternativa(BaseModel):
    """Segunda/tercera opción para que el operador desempate sin volver a buscar."""

    gasto_id: str = Field(description="ID de un gasto de la lista de candidatos")
    confianza: float = Field(ge=0, le=1, default=0.0)
    razon: str = Field(default="")


class ConciliacionSugerirResponse(BaseModel):
    gasto_id_sugerido: str | None = Field(
        default=None, description="ID del gasto candidato más probable, o null si ninguno encaja"
    )
    confianza: float = Field(ge=0, le=1, default=0.0, description="Confianza 0..1 del match")
    razon: str = Field(default="", description="Explicación breve en español del match")
    modelo: str = Field(description="Modelo de Claude usado")
    uso_ia: UsoIA | None = Field(default=None, description="Consumo de tokens (aditivo)")
    # ADITIVOS (15-sep-2026): la IA PROPONE, el humano confirma — y para
    # confirmar hay que ver POR QUÉ.
    evidencias: list[str] = Field(
        default_factory=list,
        description="Hechos que sostienen el match (monto exacto, terminación 0577, ASUR ≈ lugar…)",
    )
    alternativas: list[SugerenciaAlternativa] = Field(
        default_factory=list, description="Hasta 3 segundas opciones, de más a menos probable"
    )
    motivo_sin_match: str | None = Field(
        default=None,
        description="Por qué ningún candidato encaja (solo cuando el sugerido es null)",
    )
