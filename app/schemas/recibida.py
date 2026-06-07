"""Schema para parsear un CFDI recibido (factura de proveedor)."""

from __future__ import annotations

from pydantic import BaseModel


class ParseRecibidaRequest(BaseModel):
    # XML del CFDI en base64.
    xml_b64: str


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
