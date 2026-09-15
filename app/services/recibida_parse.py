"""Parseo determinista de un CFDI recibido (sin IA: un CFDI es estructurado).

Soporta CFDI 3.3 y 4.0 (deriva el namespace cfdi del root). Extrae emisor,
receptor, totales, moneda, fecha y el UUID del Timbre Fiscal Digital.

SEGURIDAD (15-sep-2026): el XML llega de un tercero (lo sube el proveedor o
el operador). Se parsea con `defusedxml` y, además, se rechaza cualquier
documento con DTD/ENTITY antes de tocar el parser — así el archivo no puede
usar XXE (leer archivos del servidor) ni «billion laughs» (agotar memoria).
Ese doble candado funciona aunque `defusedxml` no esté instalado.
"""

from __future__ import annotations

import base64
import re
import xml.etree.ElementTree as ET  # noqa: N817 — solo para tipos y fallback

from app.schemas.recibida import FacturaRecibidaParsed, ParseRecibidaRequest

TFD_NS = "http://www.sat.gob.mx/TimbreFiscalDigital"
VERSIONES_SOPORTADAS = ("3.3", "4.0")

try:  # defusedxml es la defensa preferida
    from defusedxml.ElementTree import fromstring as _fromstring_seguro

    DEFUSED_DISPONIBLE = True
except ImportError:  # pragma: no cover - depende del entorno
    _fromstring_seguro = ET.fromstring
    DEFUSED_DISPONIBLE = False

_DTD = re.compile(rb"<!\s*(DOCTYPE|ENTITY)", re.IGNORECASE)


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


def _rfc_norm(rfc: str | None) -> str:
    return (rfc or "").strip().upper().replace("-", "").replace(" ", "")


def parsear_xml_seguro(raw: bytes) -> ET.Element:
    """Root del XML, rechazando DTD/entidades antes de parsear."""
    if _DTD.search(raw[:4096]):
        raise ValueError(
            "El XML declara un DOCTYPE/ENTITY: no se procesa por seguridad. "
            "Sube el CFDI original que emitió el proveedor."
        )
    return _fromstring_seguro(raw)


def parse_cfdi(req: ParseRecibidaRequest) -> FacturaRecibidaParsed:
    raw = base64.b64decode(req.xml_b64)
    root = parsear_xml_seguro(raw)  # Comprobante
    ns = {"cfdi": _cfdi_ns(root), "tfd": TFD_NS}

    emisor = root.find("cfdi:Emisor", ns)
    receptor = root.find("cfdi:Receptor", ns)
    tfd = root.find(".//tfd:TimbreFiscalDigital", ns)
    conceptos = root.findall(".//cfdi:Concepto", ns)
    descripciones = [c.get("Descripcion", "") for c in conceptos[:5]]
    resumen = " · ".join(d for d in descripciones if d)[:500] or None

    # `Version` en 4.0/3.3; los CFDI 3.2 usaban `version` en minúscula.
    version = root.get("Version") or root.get("version")
    uuid = tfd.get("UUID") if tfd is not None else None
    receptor_rfc = receptor.get("Rfc") if receptor is not None else None

    advertencias: list[str] = []
    motivo: str | None = None
    valido = True

    if uuid is None:
        valido = False
        motivo = (
            "El XML no trae Timbre Fiscal Digital (UUID): no es un CFDI timbrado "
            "válido, es un borrador o un archivo incompleto."
        )
    if version and version not in VERSIONES_SOPORTADAS:
        advertencias.append(
            f"Versión de CFDI {version} (soportadas: {', '.join(VERSIONES_SOPORTADAS)})."
        )
    if len(conceptos) > 5:
        advertencias.append(
            f"La factura tiene {len(conceptos)} conceptos; el resumen muestra los primeros 5."
        )
    propios = {_rfc_norm(r) for r in (req.rfcs_propios or []) if _rfc_norm(r)}
    if propios and _rfc_norm(receptor_rfc) not in propios:
        # Una factura ajena (otro receptor) NO es una factura recibida por la
        # empresa: registrarla desharía el control de gastos deducibles.
        valido = False
        motivo = (
            f"El receptor del CFDI ({receptor_rfc or 'sin RFC'}) no es ninguna de las "
            "razones sociales de la empresa: esta factura es de otro contribuyente."
        )
    if not DEFUSED_DISPONIBLE:
        advertencias.append(
            "defusedxml no está instalado en el servidor: el XML se validó con el "
            "candado propio (sin DTD)."
        )

    return FacturaRecibidaParsed(
        uuid_fiscal=uuid,
        emisor_rfc=emisor.get("Rfc") if emisor is not None else None,
        emisor_nombre=emisor.get("Nombre") if emisor is not None else None,
        receptor_rfc=receptor_rfc,
        receptor_nombre=receptor.get("Nombre") if receptor is not None else None,
        tipo_comprobante=root.get("TipoDeComprobante"),
        subtotal=_f(root.get("SubTotal")),
        total=_f(root.get("Total")),
        moneda=root.get("Moneda"),
        fecha_emision=root.get("Fecha"),
        conceptos_resumen=resumen,
        version_cfdi=version,
        conceptos_n=len(conceptos),
        valido=valido,
        motivo=motivo,
        advertencias=advertencias,
    )
