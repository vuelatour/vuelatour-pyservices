"""Timbrado vía Facturama (API Multiemisor / api-lite).

Facturama NO usa API key: autentica con Basic Auth (usuario:contraseña de la
cuenta, el mismo login del portal). El flujo multiemisor:
  1. Cargar una vez el CSD de cada RFC emisor (POST /api-lite/csds).
  2. Emitir el CFDI como JSON (POST /api-lite/3/cfdis): Facturama sella con el
     CSD guardado y timbra — sin armar/firmar XML de nuestro lado (a
     diferencia de FEL) y SIN PFX para cancelar.
  3. Descargar XML/PDF: GET /Cfdi/{formato}/issuedLite/{id}.
  4. Cancelar: DELETE /api-lite/cfdis/{id}?motive=..&uuidReplacement=..

Sandbox: https://apisandbox.facturama.mx (cuenta gratuita propia, separada de
producción). Producción: https://api.facturama.mx.
"""

from __future__ import annotations

import base64
import json
import logging
import urllib.error
import urllib.parse
import urllib.request

from app.config import get_settings
from app.schemas.facturacion import (
    CancelarRequest,
    CancelarResponse,
    TimbrarRequest,
    TimbrarResponse,
)

logger = logging.getLogger(__name__)

_TIMEOUT_S = 60


def _base_url() -> str:
    s = get_settings()
    return (
        "https://api.facturama.mx"
        if s.facturama_modo == "prod"
        else "https://apisandbox.facturama.mx"
    )


def _auth_header() -> str:
    s = get_settings()
    tok = base64.b64encode(
        f"{s.facturama_user}:{s.facturama_password}".encode()
    ).decode()
    return f"Basic {tok}"


def _request(method: str, path: str, body: dict | None = None) -> tuple[int, dict | list | None]:
    url = f"{_base_url()}{path}"
    req = urllib.request.Request(
        url,
        data=json.dumps(body).encode() if body is not None else None,
        headers={
            "Authorization": _auth_header(),
            "Content-Type": "application/json",
        },
        method=method,
    )
    try:
        with urllib.request.urlopen(req, timeout=_TIMEOUT_S) as resp:
            raw = resp.read()
            return resp.status, json.loads(raw) if raw else None
    except urllib.error.HTTPError as e:
        raw = e.read()
        try:
            return e.code, json.loads(raw) if raw else None
        except json.JSONDecodeError:
            return e.code, {"Message": raw.decode(errors="replace")[:500]}


def _error_text(payload: dict | list | None) -> str:
    """Aplana el error de Facturama (Message + ModelState) a una línea legible."""
    if not isinstance(payload, dict):
        return str(payload)[:300]
    partes: list[str] = []
    if payload.get("Message"):
        partes.append(str(payload["Message"]))
    model_state = payload.get("ModelState")
    if isinstance(model_state, dict):
        for k, v in model_state.items():
            detalle = "; ".join(str(x) for x in v) if isinstance(v, list) else str(v)
            partes.append(f"{k}: {detalle}")
    return " | ".join(partes)[:500] or "Error de Facturama sin detalle"


def _ensure_csd(req: TimbrarRequest) -> str | None:
    """Garantiza que el CSD del RFC emisor esté cargado. Devuelve error o None."""
    rfc = req.emisor.rfc
    status, _ = _request("GET", f"/api-lite/csds/{rfc}")
    if status == 200:
        return None
    payload = {
        "Rfc": rfc,
        "Certificate": req.csd_cer_b64,
        "PrivateKey": req.csd_key_b64,
        "PrivateKeyPassword": req.csd_password,
    }
    status, body = _request("POST", "/api-lite/csds", payload)
    if status not in (200, 201):
        return f"No se pudo cargar el CSD de {rfc} en Facturama: {_error_text(body)}"
    logger.info("CSD de %s cargado en Facturama", rfc)
    return None


def _descargar(formato: str, cfdi_id: str) -> str | None:
    """Descarga XML/PDF del CFDI emitido (viene como base64 en Content)."""
    status, body = _request("GET", f"/Cfdi/{formato}/issuedLite/{cfdi_id}")
    if status == 200 and isinstance(body, dict) and body.get("Content"):
        return str(body["Content"])
    logger.warning("Facturama: no se pudo descargar %s del CFDI %s (%s)", formato, cfdi_id, status)
    return None


def _to_cfdi_json(req: TimbrarRequest) -> dict:
    items = []
    for c in req.conceptos:
        subtotal = round(c.cantidad * c.valor_unitario, 2)
        iva = round(subtotal * c.tasa_iva, 2)
        items.append(
            {
                "ProductCode": c.clave_prod_serv,
                "UnitCode": c.clave_unidad,
                "Description": c.descripcion,
                "Quantity": c.cantidad,
                "UnitPrice": c.valor_unitario,
                "Subtotal": subtotal,
                "TaxObject": "02",
                "Taxes": [
                    {
                        "Name": "IVA",
                        "Base": subtotal,
                        "Rate": c.tasa_iva,
                        "Total": iva,
                        "IsRetention": False,
                    }
                ],
                "Total": round(subtotal + iva, 2),
            }
        )

    cfdi: dict = {
        "CfdiType": req.tipo_comprobante or "I",
        "Currency": req.moneda,
        "PaymentForm": req.forma_pago,
        "PaymentMethod": req.metodo_pago,
        "ExpeditionPlace": req.lugar_expedicion,
        "Issuer": {
            "Rfc": req.emisor.rfc,
            "Name": req.emisor.nombre,
            "FiscalRegime": req.emisor.regimen_fiscal,
        },
        "Receiver": {
            "Rfc": req.receptor.rfc,
            "Name": req.receptor.nombre,
            "FiscalRegime": req.receptor.regimen_fiscal,
            "CfdiUse": req.receptor.uso_cfdi,
            "TaxZipCode": req.receptor.domicilio_fiscal,
        },
        "Items": items,
    }
    if req.folio:
        cfdi["Folio"] = req.folio
    if req.serie:
        cfdi["Serie"] = req.serie
    # PÚBLICO EN GENERAL (XAXX010101000): nodo obligatorio en CFDI 4.0.
    if req.informacion_global is not None:
        cfdi["GlobalInformation"] = {
            "Periodicity": req.informacion_global.periodicidad,
            "Months": req.informacion_global.meses,
            "Year": req.informacion_global.anio,
        }
    # Nota de crédito (Egreso relacionado a la factura original).
    if req.cfdi_relacionado_uuid:
        cfdi["Relations"] = {
            "Type": req.tipo_relacion or "01",
            "Cfdis": [{"Uuid": req.cfdi_relacionado_uuid}],
        }
    return cfdi


def timbrar_facturama(req: TimbrarRequest) -> TimbrarResponse:
    error_csd = _ensure_csd(req)
    if error_csd:
        return TimbrarResponse(ok=False, error=error_csd)

    status, body = _request("POST", "/api-lite/3/cfdis", _to_cfdi_json(req))
    if status not in (200, 201) or not isinstance(body, dict):
        return TimbrarResponse(ok=False, error=_error_text(body))

    cfdi_id = str(body.get("Id") or "")
    complement = body.get("Complement") or {}
    tax_stamp = complement.get("TaxStamp") or {} if isinstance(complement, dict) else {}
    uuid = tax_stamp.get("Uuid") or body.get("Uuid")
    fecha = tax_stamp.get("Date") or body.get("Date")

    return TimbrarResponse(
        ok=True,
        uuid=str(uuid) if uuid else None,
        fecha_timbrado=str(fecha) if fecha else None,
        xml_b64=_descargar("xml", cfdi_id) if cfdi_id else None,
        pdf_b64=_descargar("pdf", cfdi_id) if cfdi_id else None,
        pac_id=cfdi_id or None,
    )


def cancelar_facturama(req: CancelarRequest) -> CancelarResponse:
    if not req.pac_id:
        return CancelarResponse(
            ok=False,
            error="Falta pac_id (Id de Facturama) de la factura: solo las timbradas con Facturama se cancelan por esta vía.",
        )
    qs = urllib.parse.urlencode(
        {
            "motive": req.motivo,
            **({"uuidReplacement": req.folio_sustitucion} if req.folio_sustitucion else {}),
        }
    )
    status, body = _request("DELETE", f"/api-lite/cfdis/{req.pac_id}?{qs}")
    if status not in (200, 201, 204):
        return CancelarResponse(ok=False, error=_error_text(body))
    estatus = None
    if isinstance(body, dict):
        estatus = str(body.get("Status") or body.get("CancelationStatus") or "cancelado")
    return CancelarResponse(ok=True, estatus=estatus or "cancelado")
