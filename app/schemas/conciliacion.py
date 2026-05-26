from typing import Literal

from pydantic import BaseModel, Field


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
