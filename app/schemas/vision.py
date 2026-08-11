from typing import Literal

from pydantic import BaseModel, Field, model_validator

MediaType = Literal["image/jpeg", "image/png", "image/webp", "image/gif"]


class TacometroRequest(BaseModel):
    """Una de dos fuentes de imagen: base64 (preferido) o URL pública/firmada."""

    image_base64: str | None = Field(default=None, description="Imagen en base64 (sin prefijo data:)")
    media_type: MediaType | None = Field(default=None, description="Requerido si se usa image_base64")
    image_url: str | None = Field(default=None, description="URL pública o firmada de la imagen")
    ultimo: float | None = Field(
        default=None,
        description="Última lectura conocida de la aeronave (ancla de magnitud). La nueva debe ser similar y nunca menor.",
    )

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
    calidad_foto: str = Field(
        default="MEDIA",
        description=(
            "Calidad de la FOTO para leer el instrumento: ALTA (dígitos nítidos), "
            "MEDIA (se lee pero con reflejo/ángulo/algo de desenfoque) o BAJA "
            "(borrosa/oscura: la lectura puede tener dígitos equivocados)."
        ),
    )
    modelo: str = Field(description="Modelo de Claude usado")


class ImagenFuente(BaseModel):
    """Una imagen: base64 (preferido) o URL pública/firmada."""

    image_base64: str | None = Field(default=None, description="Imagen en base64 (sin prefijo data:)")
    media_type: MediaType | None = Field(default=None, description="Requerido si se usa image_base64")
    image_url: str | None = Field(default=None, description="URL pública o firmada de la imagen")

    @model_validator(mode="after")
    def _check_source(self) -> "ImagenFuente":
        if not self.image_base64 and not self.image_url:
            raise ValueError("Debes enviar image_base64 o image_url")
        if self.image_base64 and not self.media_type:
            raise ValueError("media_type es requerido cuando se envía image_base64")
        return self


class GastoTicketRequest(BaseModel):
    """Ticket/factura de gasto. Fuentes (exactamente una):
    - image_base64+media_type o image_url (una foto, contrato original),
    - images: varias fotos = HOJAS del MISMO documento (facturas multi-página),
    - pdf_base64: la factura en PDF (visión nativa de documento).
    Campos nuevos ADITIVOS: el API viejo sigue mandando una sola imagen."""

    image_base64: str | None = Field(default=None, description="Imagen en base64 (sin prefijo data:)")
    media_type: MediaType | None = Field(default=None, description="Requerido si se usa image_base64")
    image_url: str | None = Field(default=None, description="URL pública o firmada de la imagen")
    images: list[ImagenFuente] | None = Field(
        default=None,
        max_length=8,
        description="Varias fotos del MISMO documento (hojas de una factura), máx 8",
    )
    pdf_base64: str | None = Field(default=None, description="Factura en PDF (base64, sin prefijo data:)")
    excel_base64: str | None = Field(
        default=None,
        description="Factura en Excel (.xlsx) o CSV en base64: se convierte a texto y la IA extrae los datos",
    )
    excel_filename: str | None = Field(
        default=None, description="Nombre del archivo Excel/CSV (decide el parser por extensión)"
    )

    @model_validator(mode="after")
    def _check_source(self) -> "GastoTicketRequest":
        fuentes = sum(
            1
            for presente in (
                bool(self.image_base64 or self.image_url),
                bool(self.images),
                bool(self.pdf_base64),
                bool(self.excel_base64),
            )
            if presente
        )
        if fuentes != 1:
            raise ValueError(
                "Debes enviar exactamente una fuente: image_base64/image_url, images[], pdf_base64 o excel_base64"
            )
        if self.image_base64 and not self.media_type:
            raise ValueError("media_type es requerido cuando se envía image_base64")
        return self


CategoriaGasto = Literal[
    "GAS",
    # Categoría unificada de la app (aterrizaje + FBO + servicios de
    # aeropuerto). Faltaba aquí y pydantic RECHAZABA toda la lectura de
    # facturas de FBO ("No se pudo interpretar el ticket").
    "OPERACIONES",
    "ATERRIZAJE",
    "TUAS",
    "FBO",
    "COMIDA",
    "HOTEL",
    "TAXI",
    "REFACCION",
    "PERMISO",
    # Honorario del piloto externo (freelance): lo captura oficina; aditivo.
    "PILOTO_EXTERNO",
    "FIJO",
    # Gasto indirecto de la operación (sin vuelo). La IA no lo sugiere (no
    # está en el prompt); aditivo para tolerar el enum del API.
    "INDIRECTO",
    "OTRO",
]


class ConceptoTicket(BaseModel):
    concepto: str = Field(description="Nombre del renglón (ej. Aterrizaje, Handling)")
    monto: float = Field(description="Importe del renglón")


class GastoTicketResponse(BaseModel):
    monto: float | None = Field(default=None, description="Total del ticket, o null si ilegible")
    propina: float | None = Field(
        default=None,
        description="Propina/tip si el ticket la muestra como línea (null si no aparece)",
    )
    folio: str | None = Field(
        default=None,
        description="Folio/remisión impreso en el ticket (llave anti-duplicados), o null",
    )
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
    litros: float | None = Field(
        default=None,
        description="Litros cargados si es ticket de combustible (galones→L), o null",
    )
    conceptos: list["ConceptoTicket"] = Field(
        default_factory=list,
        description="Renglones del ticket incl. IVA como renglón si viene aparte (suma = total); máx 8",
    )
    matricula: str | None = Field(
        default=None, description="Matrícula de la aeronave si aparece (XA-ABC/N123XX)"
    )
    confianza: float = Field(ge=0, le=1, default=0.0, description="Confianza 0..1 de la extracción")
    legible: bool = Field(default=False, description="true si el ticket se pudo leer")
    notas: str = Field(default="", description="Observaciones")
    modelo: str = Field(description="Modelo de Claude usado")


class ConstanciaFiscalRequest(BaseModel):
    """Constancia de Situación Fiscal del SAT. Exactamente UNA fuente:
    pdf_base64 (documento nativo, mismo patrón que gasto-ticket) o
    image_base64+media_type (foto/captura de la constancia)."""

    pdf_base64: str | None = Field(
        default=None, description="Constancia en PDF (base64, sin prefijo data:)"
    )
    image_base64: str | None = Field(
        default=None, description="Imagen en base64 (sin prefijo data:)"
    )
    media_type: MediaType | None = Field(
        default=None, description="Requerido si se usa image_base64"
    )

    @model_validator(mode="after")
    def _check_source(self) -> "ConstanciaFiscalRequest":
        if bool(self.pdf_base64) == bool(self.image_base64):
            raise ValueError("Debes enviar exactamente una fuente: pdf_base64 o image_base64")
        if self.image_base64 and not self.media_type:
            raise ValueError("media_type es requerido cuando se envía image_base64")
        return self


class ConstanciaFiscalResponse(BaseModel):
    """Datos fiscales extraídos de la constancia. Contrato con el API:
    los fallos de IA degradan a disponible=false/legible=false, nunca 500."""

    disponible: bool = Field(
        default=False, description="false si la IA no pudo procesar (Claude caído)"
    )
    legible: bool = Field(default=False, description="true si la constancia se pudo leer")
    rfc: str | None = Field(default=None, description="RFC del contribuyente (12-13 caracteres)")
    razon_social: str | None = Field(
        default=None,
        description="Denominación (PM, SIN régimen societario) o nombre completo (PF)",
    )
    regimen_fiscal: str | None = Field(
        default=None, description="Código c_RegimenFiscal (3 dígitos) del régimen vigente"
    )
    regimen_descripcion: str | None = Field(
        default=None, description="Nombre del régimen tal cual la constancia"
    )
    cp: str | None = Field(
        default=None, description="Código postal del domicilio fiscal (5 dígitos)"
    )
    domicilio: str | None = Field(
        default=None, description="Domicilio fiscal completo en una línea"
    )
    confianza: float = Field(ge=0, le=1, default=0.0, description="Confianza 0..1 de la extracción")
    motivo: str | None = Field(
        default=None, description="Por qué no se pudo leer / observaciones de la extracción"
    )


class CombustibleTicketResponse(BaseModel):
    """Datos extraídos de un ticket de carga de combustible (turbosina/avgas)."""

    litros: float | None = Field(default=None, description="Litros cargados (galones→L si aplica), o null")
    precio_litro: float | None = Field(default=None, description="Precio por litro, o null")
    total: float | None = Field(default=None, description="Total pagado, o null")
    folio: str | None = Field(
        default=None,
        description="Folio/remisión impreso en el ticket (ej. remisión ASA), o null",
    )
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
