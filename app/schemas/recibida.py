"""Schema para parsear un CFDI recibido (factura de proveedor)."""

from __future__ import annotations

from pydantic import BaseModel, Field


class ParseRecibidaRequest(BaseModel):
    # XML del CFDI en base64.
    xml_b64: str
    # ADITIVO (15-sep-2026): RFC de las razones sociales PROPIAS. Si llegan, se
    # valida que el receptor sea una de ellas (una factura ajena no es una
    # factura recibida por la empresa). Sin la lista no se valida nada.
    rfcs_propios: list[str] = Field(
        default_factory=list, description="RFC de las razones sociales de la empresa"
    )


class FacturaRecibidaParsed(BaseModel):
    uuid_fiscal: str | None = None
    emisor_rfc: str | None = None
    emisor_nombre: str | None = None
    receptor_rfc: str | None = None
    receptor_nombre: str | None = None
    tipo_comprobante: str | None = None
    subtotal: float | None = None
    total: float | None = None
    moneda: str | None = None
    fecha_emision: str | None = None
    conceptos_resumen: str | None = None
    # ADITIVOS (15-sep-2026). `valido` sale true salvo que falte el timbre o el
    # receptor no sea de la empresa: un API que no lo lea se comporta igual
    # que antes.
    version_cfdi: str | None = Field(default=None, description="Version del CFDI (3.3 / 4.0)")
    conceptos_n: int | None = Field(default=None, description="Cuantos conceptos trae el CFDI")
    valido: bool = Field(default=True, description="false si el CFDI no se puede dar por bueno")
    motivo: str | None = Field(default=None, description="Por que no es valido, si aplica")
    advertencias: list[str] = Field(default_factory=list)
