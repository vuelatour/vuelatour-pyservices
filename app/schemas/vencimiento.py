from typing import Literal

from pydantic import BaseModel, Field, model_validator

from app.schemas.uso_ia import UsoIA

MediaType = Literal["image/jpeg", "image/png", "image/webp", "image/gif"]


class VencimientoExtraerRequest(BaseModel):
    """Documento de vencimiento renovado (póliza/seguro/tarjeta de circulación).

    Fuente: PDF (pdf_base64) o imagen (image_base64 + media_type). Envía una.
    """

    pdf_base64: str | None = Field(default=None, description="PDF en base64 (sin prefijo data:)")
    image_base64: str | None = Field(
        default=None, description="Imagen en base64 (sin prefijo data:)"
    )
    media_type: MediaType | None = Field(
        default=None, description="Requerido si se envía image_base64"
    )

    # ADITIVOS (15-sep-2026): catálogos vivos del API para validar la lectura.
    matriculas_flota: list[str] | None = Field(
        default=None,
        max_length=80,
        description="Matrículas/placas vigentes (aeronaves y vehículos)",
    )
    tipos_documento: list[str] | None = Field(
        default=None, max_length=60, description="Catálogo de tipos de documento del sistema"
    )

    @model_validator(mode="after")
    def _check_source(self) -> "VencimientoExtraerRequest":
        if not self.pdf_base64 and not self.image_base64:
            raise ValueError("Debes enviar pdf_base64 o image_base64")
        if self.image_base64 and not self.media_type:
            raise ValueError("media_type es requerido cuando se envía image_base64")
        return self


class VencimientoExtraerResponse(BaseModel):
    matricula: str | None = Field(
        default=None, description="Matrícula/placa de la aeronave o vehículo, o null"
    )
    tipo_documento: str | None = Field(
        default=None, description="Tipo de documento (póliza, seguro, tarjeta de circulación...)"
    )
    fecha_vigencia: str | None = Field(
        default=None, description="Inicio de vigencia YYYY-MM-DD, o null"
    )
    fecha_vencimiento: str | None = Field(
        default=None, description="Fecha de vencimiento/expiración YYYY-MM-DD, o null"
    )
    emisor: str | None = Field(
        default=None, description="Aseguradora/entidad emisora del documento, o null"
    )
    # ADITIVOS (15-sep-2026): una póliza de FLOTA cubre varias aeronaves (antes
    # se perdían todas menos la primera) y el folio evita duplicar la renovación.
    matriculas: list[str] = Field(
        default_factory=list,
        description="Todas las matrículas que cubre el documento (la primera va en `matricula`)",
    )
    numero_poliza: str | None = Field(
        default=None, description="Número de póliza/folio del documento, o null"
    )
    confianza: float = Field(ge=0, le=1, default=0.0, description="Confianza 0..1 de la extracción")
    notas: str = Field(default="", description="Observaciones breves en español")
    advertencias: list[str] = Field(
        default_factory=list, description="Validaciones que no pasaron (fechas, matrícula, tipo)"
    )
    modelo: str = Field(description="Modelo de Claude usado")
    uso_ia: UsoIA | None = Field(default=None, description="Consumo de tokens (aditivo)")
