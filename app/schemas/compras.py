from pydantic import BaseModel, Field


class CompraExtraerRequest(BaseModel):
    """PDF de una factura/orden de compra (ej. Aircraft Spruce) en base64."""

    pdf_base64: str = Field(description="PDF en base64 (sin prefijo data:)")


class CompraLinea(BaseModel):
    nombre: str = Field(description="Descripción del producto")
    numero_parte: str | None = Field(default=None, description="Part number si aparece")
    cantidad: float = Field(default=1, description="Cantidad")
    precio_unitario_usd: float | None = Field(default=None, description="Precio unitario")
    total_usd: float | None = Field(default=None, description="Total de la línea")


class CompraExtraerResponse(BaseModel):
    proveedor: str | None = Field(default=None, description="Proveedor/comercio")
    fecha: str | None = Field(default=None, description="Fecha de la orden YYYY-MM-DD")
    moneda: str = Field(default="USD", description="Moneda detectada")
    lineas: list[CompraLinea] = Field(default_factory=list)
    subtotal_usd: float | None = None
    shipping_usd: float | None = None
    impuestos_usd: float | None = None
    total_usd: float | None = None
    confianza: float = Field(ge=0, le=1, default=0.0)
    notas: str = Field(default="")
    modelo: str = Field(description="Modelo de Claude usado")
