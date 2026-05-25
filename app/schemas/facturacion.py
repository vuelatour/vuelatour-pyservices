from pydantic import BaseModel, Field


class Emisor(BaseModel):
    rfc: str
    nombre: str
    regimen_fiscal: str  # c_RegimenFiscal (ej. 601)


class Receptor(BaseModel):
    rfc: str
    nombre: str
    domicilio_fiscal: str  # CP (DomicilioFiscalReceptor)
    regimen_fiscal: str  # c_RegimenFiscal del receptor
    uso_cfdi: str  # c_UsoCFDI (ej. G03)


class Concepto(BaseModel):
    clave_prod_serv: str = "78101800"  # transporte aéreo (default razonable)
    clave_unidad: str = "E48"  # unidad de servicio
    cantidad: float = 1
    descripcion: str
    valor_unitario: float
    objeto_imp: str = "02"  # sí objeto de impuesto
    tasa_iva: float = 0.16


class TimbrarRequest(BaseModel):
    referencia: str = Field(min_length=4, description="Única por CFDI (idempotencia FEL)")
    serie: str | None = None
    folio: str | None = None
    moneda: str = "MXN"
    forma_pago: str = "03"  # transferencia (c_FormaPago)
    metodo_pago: str = "PUE"  # pago en una exhibición
    lugar_expedicion: str  # CP del emisor
    emisor: Emisor
    receptor: Receptor
    conceptos: list[Concepto]
    # CSD del emisor (base64) + contraseña de la llave.
    csd_cer_b64: str
    csd_key_b64: str
    csd_password: str


class TimbrarResponse(BaseModel):
    ok: bool
    uuid: str | None = None
    fecha_timbrado: str | None = None
    xml_b64: str | None = None
    pdf_b64: str | None = None
    error: str | None = None
