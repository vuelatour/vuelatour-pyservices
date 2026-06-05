"""Generación y timbrado de CFDI 4.0 con FEL.

El armado y sellado del CFDI usa la librería `satcfdi` (Anexo 20 + cadena
original + sello con CSD). El timbrado se hace contra el Web Service SOAP de FEL
(método TimbrarCFDI). Todo el flujo es defensivo: ante cualquier falta de
configuración o error, devuelve un TimbrarResponse(ok=False, error=...) en vez
de tumbar el servicio.

NOTA: el armado exacto del CFDI (claves de catálogos, impuestos) debe validarse
contra el ambiente de PRUEBAS de FEL antes de producción.
"""

import base64
import logging
from decimal import Decimal

from app.config import get_settings
from app.schemas.facturacion import (
    CancelarRequest,
    CancelarResponse,
    TimbrarRequest,
    TimbrarResponse,
)

logger = logging.getLogger("facturacion")


def _construir_y_sellar(req: TimbrarRequest) -> bytes:
    """Arma el CFDI 4.0 y lo sella con el CSD. Devuelve el XML en bytes."""
    from satcfdi.create.cfd import cfdi40
    from satcfdi.models import Signer

    signer = Signer.load(
        certificate=base64.b64decode(req.csd_cer_b64),
        key=base64.b64decode(req.csd_key_b64),
        password=req.csd_password,
    )

    conceptos = [
        cfdi40.Concepto(
            clave_prod_serv=c.clave_prod_serv,
            cantidad=Decimal(str(c.cantidad)),
            clave_unidad=c.clave_unidad,
            descripcion=c.descripcion,
            valor_unitario=Decimal(str(c.valor_unitario)),
            objeto_imp=c.objeto_imp,
            impuestos=cfdi40.Impuestos(
                traslados=[
                    cfdi40.Traslado(
                        impuesto="002",  # IVA
                        tipo_factor="Tasa",
                        tasa_o_cuota=Decimal(str(c.tasa_iva)),
                    )
                ],
            ),
        )
        for c in req.conceptos
    ]

    # Relación a otro CFDI (nota de crédito → factura original).
    relacionados = None
    if req.cfdi_relacionado_uuid:
        relacionados = cfdi40.CfdiRelacionados(
            tipo_relacion=req.tipo_relacion or "01",
            cfdi_relacionado=[
                cfdi40.CfdiRelacionado(uuid=req.cfdi_relacionado_uuid)
            ],
        )

    comprobante = cfdi40.Comprobante(
        emisor=cfdi40.Emisor(
            rfc=req.emisor.rfc,
            nombre=req.emisor.nombre,
            regimen_fiscal=req.emisor.regimen_fiscal,
        ),
        lugar_expedicion=req.lugar_expedicion,
        tipo_de_comprobante=req.tipo_comprobante or "I",
        cfdi_relacionados=relacionados,
        receptor=cfdi40.Receptor(
            rfc=req.receptor.rfc,
            nombre=req.receptor.nombre,
            domicilio_fiscal_receptor=req.receptor.domicilio_fiscal,
            regimen_fiscal_receptor=req.receptor.regimen_fiscal,
            uso_cfdi=req.receptor.uso_cfdi,
        ),
        conceptos=conceptos,
        moneda=req.moneda,
        forma_pago=req.forma_pago,
        metodo_pago=req.metodo_pago,
        serie=req.serie,
        folio=req.folio,
    )
    comprobante.sign(signer)
    return comprobante.xml_bytes()


def _timbrar_fel(xml_bytes: bytes, referencia: str) -> dict:
    """Llama al Web Service SOAP de FEL (TimbrarCFDI)."""
    from zeep import Client, Settings as ZeepSettings

    s = get_settings()
    client = Client(s.fel_wsdl_url, settings=ZeepSettings(strict=False, xml_huge_tree=True))
    resp = client.service.TimbrarCFDI(
        usuario=s.fel_usuario,
        password=s.fel_password,
        cadenaXML=xml_bytes.decode("utf-8"),
        referencia=referencia,
    )
    return {
        "ok": bool(getattr(resp, "OperacionExitosa", False)),
        "xml": getattr(resp, "XMLResultado", None),
        "pdf": getattr(resp, "PDFResultado", None),
        "uuid": getattr(getattr(resp, "Timbre", None), "UUID", None),
        "fecha": getattr(getattr(resp, "Timbre", None), "FechaTimbrado", None),
        "error": getattr(resp, "MensajeError", None)
        or getattr(resp, "MensajeErrorDetallado", None),
    }


def timbrar(req: TimbrarRequest) -> TimbrarResponse:
    s = get_settings()
    if not s.fel_configurado:
        return TimbrarResponse(ok=False, error="FEL no configurado (faltan credenciales de timbrado).")
    try:
        xml_bytes = _construir_y_sellar(req)
    except Exception as e:  # noqa: BLE001
        logger.exception("Error armando/sellando CFDI")
        return TimbrarResponse(ok=False, error=f"No se pudo armar/sellar el CFDI: {e}")

    try:
        r = _timbrar_fel(xml_bytes, req.referencia)
    except Exception as e:  # noqa: BLE001
        logger.exception("Error timbrando con FEL")
        return TimbrarResponse(ok=False, error=f"FEL no disponible: {e}")

    if not r["ok"]:
        return TimbrarResponse(ok=False, error=r.get("error") or "FEL rechazó el comprobante")

    xml_final = r["xml"] or xml_bytes.decode("utf-8")
    return TimbrarResponse(
        ok=True,
        uuid=r["uuid"],
        fecha_timbrado=str(r["fecha"]) if r["fecha"] else None,
        xml_b64=base64.b64encode(xml_final.encode("utf-8")).decode("ascii"),
        pdf_b64=r["pdf"] or None,
    )


def cancelar(req: CancelarRequest) -> CancelarResponse:
    """Cancela un CFDI ya timbrado vía FEL (CancelarCFDI).

    NOTA: confirmar la firma exacta del método `CancelarCFDI` contra el WSDL de
    FEL/Blikon al contratar (nombres de parámetros pueden variar). El CFDI 4.0
    requiere indicar el motivo (c_MotivoCancelacion) y, si es 01, el UUID que lo
    sustituye. Validar contra el ambiente de PRUEBAS antes de producción.
    """
    s = get_settings()
    if not s.fel_configurado:
        return CancelarResponse(
            ok=False, error="FEL no configurado (faltan credenciales de timbrado)."
        )
    try:
        from zeep import Client, Settings as ZeepSettings

        client = Client(
            s.fel_wsdl_url, settings=ZeepSettings(strict=False, xml_huge_tree=True)
        )
        resp = client.service.CancelarCFDI(
            usuario=s.fel_usuario,
            password=s.fel_password,
            uuid=req.uuid,
            rfcEmisor=req.rfc_emisor,
            motivo=req.motivo,
            folioSustitucion=req.folio_sustitucion or "",
        )
    except Exception as e:  # noqa: BLE001
        logger.exception("Error cancelando con FEL")
        return CancelarResponse(ok=False, error=f"FEL no disponible: {e}")

    ok = bool(getattr(resp, "OperacionExitosa", False))
    return CancelarResponse(
        ok=ok,
        estatus=getattr(resp, "Estatus", None)
        or getattr(resp, "EstatusUUID", None),
        acuse_xml=getattr(resp, "Acuse", None)
        or getattr(resp, "XMLResultado", None),
        error=None
        if ok
        else (
            getattr(resp, "MensajeError", None)
            or getattr(resp, "MensajeErrorDetallado", None)
            or "FEL rechazó la cancelación"
        ),
    )
