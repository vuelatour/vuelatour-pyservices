from typing import Literal

from pydantic import BaseModel, Field

from app.schemas.uso_ia import UsoIA


class ConciliacionParseRequest(BaseModel):
    """Estado de cuenta a parsear. CSV/Excel (preferido) o PDF, en base64."""

    filename: str = Field(description="Nombre del archivo (define el parser por extensión)")
    file_base64: str = Field(description="Contenido en base64 (sin prefijo data:)")


class MovimientoParseado(BaseModel):
    fecha: str | None = Field(default=None, description="Fecha YYYY-MM-DD")
    descripcion: str | None = Field(default=None)
    monto: float = Field(description="Monto positivo")
    tipo: Literal["CARGO", "ABONO"] = Field(description="CARGO = salida, ABONO = entrada")
    referencia: str | None = Field(default=None)


class ConciliacionParseResponse(BaseModel):
    movimientos: list[MovimientoParseado] = Field(default_factory=list)
    total: int = Field(default=0)
    formato: str = Field(description="csv | excel | pdf")
    notas: str = Field(default="")
    modelo: str | None = Field(default=None, description="Modelo de Claude si se usó (PDF)")
    uso_ia: UsoIA | None = Field(
        default=None, description="Consumo de tokens (solo PDF; CSV/Excel = None)"
    )


class MovimientoSinConciliar(BaseModel):
    """Movimiento bancario que aún no tiene gasto vinculado."""

    fecha: str | None = Field(default=None, description="Fecha del movimiento YYYY-MM-DD")
    monto: float = Field(description="Monto del movimiento (positivo)")
    descripcion: str | None = Field(default=None, description="Descripción/concepto del banco")


class GastoCandidato(BaseModel):
    id: str = Field(description="Identificador del gasto")
    fecha: str | None = Field(default=None, description="Fecha del gasto YYYY-MM-DD")
    monto: float = Field(description="Monto del gasto (positivo)")
    proveedor: str | None = Field(default=None, description="Proveedor/comercio del gasto")


class ConciliacionSugerirRequest(BaseModel):
    """Movimiento sin conciliar + gastos candidatos cercanos para que Claude proponga el match."""

    movimiento: MovimientoSinConciliar
    candidatos: list[GastoCandidato] = Field(default_factory=list)


class ConciliacionSugerirResponse(BaseModel):
    gasto_id_sugerido: str | None = Field(
        default=None, description="ID del gasto candidato más probable, o null si ninguno encaja"
    )
    confianza: float = Field(ge=0, le=1, default=0.0, description="Confianza 0..1 del match")
    razon: str = Field(default="", description="Explicación breve en español del match")
    modelo: str = Field(description="Modelo de Claude usado")
    uso_ia: UsoIA | None = Field(default=None, description="Consumo de tokens (aditivo)")
