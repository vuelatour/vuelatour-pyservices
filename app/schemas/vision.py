from typing import Literal

from pydantic import BaseModel, Field, model_validator

MediaType = Literal["image/jpeg", "image/png", "image/webp", "image/gif"]


class TacometroRequest(BaseModel):
    """Una de dos fuentes de imagen: base64 (preferido) o URL pública/firmada."""

    image_base64: str | None = Field(default=None, description="Imagen en base64 (sin prefijo data:)")
    media_type: MediaType | None = Field(default=None, description="Requerido si se usa image_base64")
    image_url: str | None = Field(default=None, description="URL pública o firmada de la imagen")

    @model_validator(mode="after")
    def _check_source(self) -> "TacometroRequest":
        if not self.image_base64 and not self.image_url:
            raise ValueError("Debes enviar image_base64 o image_url")
        if self.image_base64 and not self.media_type:
            raise ValueError("media_type es requerido cuando se envía image_base64")
        return self


class TacometroResponse(BaseModel):
    lectura: float | None = Field(description="Lectura del tacómetro en horas (HOBBS), o null si ilegible")
    confianza: float = Field(ge=0, le=1, description="Confianza 0..1 de la lectura")
    legible: bool = Field(description="true si el display se pudo leer")
    notas: str = Field(default="", description="Observaciones (borrosa, reflejo, dígitos parciales, etc.)")
    modelo: str = Field(description="Modelo de Claude usado")


class GastoTicketRequest(BaseModel):
    """Foto de un ticket/recibo de gasto. base64 (preferido) o URL firmada."""

    image_base64: str | None = Field(default=None, description="Imagen en base64 (sin prefijo data:)")
    media_type: MediaType | None = Field(default=None, description="Requerido si se usa image_base64")
    image_url: str | None = Field(default=None, description="URL pública o firmada de la imagen")

    @model_validator(mode="after")
    def _check_source(self) -> "GastoTicketRequest":
        if not self.image_base64 and not self.image_url:
            raise ValueError("Debes enviar image_base64 o image_url")
        if self.image_base64 and not self.media_type:
            raise ValueError("media_type es requerido cuando se envía image_base64")
        return self


CategoriaGasto = Literal[
    "GAS",
    "ATERRIZAJE",
    "TUAS",
    "FBO",
    "COMIDA",
    "HOTEL",
    "TAXI",
    "REFACCION",
    "PERMISO",
    "FIJO",
    "OTRO",
]


class GastoTicketResponse(BaseModel):
    monto: float | None = Field(default=None, description="Total del ticket, o null si ilegible")
    moneda: Literal["MXN", "USD"] | None = Field(default=None, description="Moneda detectada")
    fecha: str | None = Field(default=None, description="Fecha del ticket YYYY-MM-DD, o null")
    proveedor: str | None = Field(default=None, description="Nombre del comercio/proveedor")
    concepto: str | None = Field(default=None, description="Descripción breve de lo comprado")
    categoria_sugerida: CategoriaGasto | None = Field(
        default=None, description="Categoría sugerida del catálogo"
    )
    medio_pago: Literal["EFECTIVO", "TARJETA_CORP", "TRANSFERENCIA"] | None = Field(
        default=None, description="Medio de pago detectado en el ticket"
    )
    tarjeta_terminacion: str | None = Field(
        default=None, description="Ultimos 4 digitos de la tarjeta, si aparecen"
    )
    confianza: float = Field(ge=0, le=1, default=0.0, description="Confianza 0..1 de la extracción")
    legible: bool = Field(default=False, description="true si el ticket se pudo leer")
    notas: str = Field(default="", description="Observaciones")
    modelo: str = Field(description="Modelo de Claude usado")


class CombustibleTicketResponse(BaseModel):
    """Datos extraídos de un ticket de carga de combustible (turbosina/avgas)."""

    litros: float | None = Field(default=None, description="Litros cargados (galones→L si aplica), o null")
    precio_litro: float | None = Field(default=None, description="Precio por litro, o null")
    total: float | None = Field(default=None, description="Total pagado, o null")
    moneda: Literal["MXN", "USD"] | None = Field(default=None, description="Moneda detectada")
    aeropuerto: str | None = Field(default=None, description="Código/nombre del aeropuerto o FBO, o null")
    tipo_combustible: Literal["TURBOSINA", "AVGAS"] | None = Field(
        default=None, description="Turbosina (Jet A) o avgas (100LL)"
    )
    fecha: str | None = Field(default=None, description="Fecha del ticket YYYY-MM-DD, o null")
    hora: str | None = Field(default=None, description="Hora de la carga HH:MM (24h), o null")
    proveedor: str | None = Field(default=None, description="Proveedor/FBO, o null")
    tarjeta_terminacion: str | None = Field(
        default=None, description="Ultimos 4 digitos de la tarjeta usada, si aparecen en el ticket"
    )
    medio_pago: Literal["EFECTIVO", "TARJETA_CORP", "TRANSFERENCIA"] | None = Field(
        default=None, description="Medio de pago detectado en el ticket"
    )
    confianza: float = Field(ge=0, le=1, default=0.0, description="Confianza 0..1")
    legible: bool = Field(default=False, description="true si el ticket se pudo leer")
    notas: str = Field(default="", description="Observaciones")
    modelo: str = Field(description="Modelo de Claude usado")
