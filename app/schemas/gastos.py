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
    notas: str | None = Field(default=None, description="Notas/desglose del gasto (proveedor, conceptos)")
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
