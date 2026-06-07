"""Schema para ensamblar un .zip a partir de archivos en base64."""

from __future__ import annotations

from pydantic import BaseModel, Field


class ArchivoZip(BaseModel):
    # Ruta/nombre dentro del zip (puede incluir carpetas, ej. "facturas/AC1.xml").
    nombre: str
    contenido_b64: str


class ZipRequest(BaseModel):
    archivos: list[ArchivoZip] = Field(default_factory=list)
