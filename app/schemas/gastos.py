"""Sugerencia IA: ¿a qué vuelo/aeronave pertenece un gasto de la bandeja?

El API arma los candidatos DETERMINISTAS (vuelos del piloto que capturó,
fecha cercana); Claude solo elige entre ellos usando el contexto del gasto
(notas, lugar, categoría). Nunca inventa vuelos.
"""

from pydantic import BaseModel, Field


class GastoParaMatch(BaseModel):
    fecha: str | None = Field(default=None, description="fecha_gasto YYYY-MM-DD")
    monto: float | None = None
    moneda: str | None = None
    categoria: str | None = None
    notas: str | None = Field(
        default=None, description="Notas/desglose del gasto (proveedor, conceptos)"
    )
    lugar: str | None = Field(default=None, description="Lugar/aeropuerto si se capturó")
    piloto_nombre: str | None = None


class VueloCandidato(BaseModel):
    vuelo_id: str
    folio: int | None = None
    fecha_vuelo: str | None = Field(default=None, description="ISO de salida")
    matricula: str | None = None
    ruta: str | None = Field(default=None, description="Ej. CUN → CZM → CUN")


class GastoVueloSugerirRequest(BaseModel):
    gasto: GastoParaMatch
    candidatos: list[VueloCandidato] = Field(default_factory=list)


class GastoVueloSugerirResponse(BaseModel):
    vuelo_id_sugerido: str | None = None
    confianza: float = Field(ge=0, le=1, default=0.0)
    razon: str = ""
    modelo: str


# --- Carga masiva de combustible (plantilla XLSX + parseo estructural) ---


class PlantillaCombustibleRequest(BaseModel):
    """Catálogos vigentes para los dropdowns de la plantilla (los manda el API)."""

    matriculas: list[str] = Field(default_factory=list)
    proveedores: list[str] = Field(default_factory=list)
    medios_pago: list[str] = Field(default_factory=list)
    monedas: list[str] = Field(default_factory=list)
    tipos_combustible: list[str] = Field(default_factory=list)


class ParseCombustibleRequest(BaseModel):
    archivo_base64: str
    filename: str = ""


class FilaCombustible(BaseModel):
    """Valores CRUDOS de una fila de la plantilla, en tipos básicos.

    Aquí NO se valida negocio (matrículas reales, folios, monedas): eso lo hace
    el API. litros/monto/tipo_cambio llegan como float si la celda fue numérica;
    si el texto no se pudo convertir se regresa la cadena cruda para que el API
    la reporte con claridad.
    """

    fila: int = Field(description="Número de fila real en el Excel/CSV (1-based)")
    matricula: str | None = None
    fecha: str | None = Field(
        default=None, description="'YYYY-MM-DD' si la celda era fecha de Excel; texto crudo si no"
    )
    hora: str | None = Field(default=None, description="'HH:MM' si la celda era hora de Excel")
    litros: float | str | None = None
    monto: float | str | None = None
    moneda: str | None = None
    tipo_cambio: float | str | None = None
    tipo_combustible: str | None = None
    lugar: str | None = None
    proveedor: str | None = None
    medio_pago: str | None = None
    folio_vuelo: str | None = None
    comprobante: str | None = None
    notas: str | None = None


class ParseCombustibleResponse(BaseModel):
    filas: list[FilaCombustible] = Field(default_factory=list)
