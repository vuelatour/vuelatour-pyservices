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
    # Tipo de comprobante: I=Ingreso (default), E=Egreso (nota de crédito).
    tipo_comprobante: str = "I"
    # Relación a otro CFDI (p. ej. nota de crédito 01 → UUID de la factura original).
    cfdi_relacionado_uuid: str | None = None
    tipo_relacion: str | None = None


class TimbrarResponse(BaseModel):
    ok: bool
    uuid: str | None = None
    fecha_timbrado: str | None = None
    xml_b64: str | None = None
    pdf_b64: str | None = None
    error: str | None = None


class CancelarRequest(BaseModel):
    uuid: str = Field(
        min_length=36, max_length=36, description="UUID (folio fiscal) del CFDI a cancelar"
    )
    rfc_emisor: str
    # c_MotivoCancelacion: 01 con sustitución, 02 errores sin relación,
    # 03 no se llevó a cabo, 04 operación nominativa global.
    motivo: str = "02"
    folio_sustitucion: str | None = None  # UUID sustituto, requerido si motivo=01
    # FEL exige RFC receptor y total del CFDI en el detalle de cancelación.
    rfc_receptor: str | None = None
    total: float | None = None


class CancelarResponse(BaseModel):
    ok: bool
    estatus: str | None = None
    acuse_xml: str | None = None
    error: str | None = None
