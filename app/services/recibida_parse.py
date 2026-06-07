"""Parseo determinista de un CFDI recibido (xml.etree, stdlib).

Soporta CFDI 3.3 y 4.0 (deriva el namespace cfdi del root). Extrae emisor,
receptor, totales, moneda, fecha y el UUID del Timbre Fiscal Digital.
"""

from __future__ import annotations

import base64
import xml.etree.ElementTree as ET

from app.schemas.recibida import FacturaRecibidaParsed, ParseRecibidaRequest

TFD_NS = "http://www.sat.gob.mx/TimbreFiscalDigital"


def _f(v: str | None) -> float | None:
    if v is None:
        return None
    try:
        return float(v)
    except ValueError:
        return None


def _cfdi_ns(root: ET.Element) -> str:
    if root.tag.startswith("{"):
        return root.tag[1:].split("}", 1)[0]
    return "http://www.sat.gob.mx/cfd/4"


def parse_cfdi(req: ParseRecibidaRequest) -> FacturaRecibidaParsed:
    raw = base64.b64decode(req.xml_b64)
    root = ET.fromstring(raw)  # Comprobante
    ns = {"cfdi": _cfdi_ns(root), "tfd": TFD_NS}

    emisor = root.find("cfdi:Emisor", ns)
    receptor = root.find("cfdi:Receptor", ns)
    tfd = root.find(".//tfd:TimbreFiscalDigital", ns)
    conceptos = root.findall(".//cfdi:Concepto", ns)
    descripciones = [c.get("Descripcion", "") for c in conceptos[:5]]
    resumen = " · ".join(d for d in descripciones if d)[:500] or None

    return FacturaRecibidaParsed(
        uuid_fiscal=tfd.get("UUID") if tfd is not None else None,
        emisor_rfc=emisor.get("Rfc") if emisor is not None else None,
        emisor_nombre=emisor.get("Nombre") if emisor is not None else None,
        receptor_rfc=receptor.get("Rfc") if receptor is not None else None,
        receptor_nombre=receptor.get("Nombre") if receptor is not None else None,
        tipo_comprobante=root.get("TipoDeComprobante"),
        subtotal=_f(root.get("SubTotal")),
        total=_f(root.get("Total")),
        moneda=root.get("Moneda"),
        fecha_emision=root.get("Fecha"),
        conceptos_resumen=resumen,
    )
