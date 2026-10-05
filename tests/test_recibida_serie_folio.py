"""Serie y folio del CFDI recibido (5-oct-2026).

Pedido del cliente: «al momento de la conciliación me apoyan a poner el número
de la factura con la que se enlaza el movimiento… en notas». El número de
factura del proveedor son los atributos `Serie`/`Folio` del
`cfdi:Comprobante`; aquí solo se leen (recortados, vacío ⇒ None). El rótulo
(«A-0411», respaldo al UUID) lo arma el API.
"""

from __future__ import annotations

import base64

from app.schemas.recibida import FacturaRecibidaParsed, ParseRecibidaRequest
from app.services.recibida_parse import parse_cfdi

UUID = "11111111-2222-3333-4444-555555555555"


def _cfdi(
    attrs_extra: str = "", ns: str = "http://www.sat.gob.mx/cfd/4", version: str = "4.0"
) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<cfdi:Comprobante xmlns:cfdi="{ns}"
  xmlns:tfd="http://www.sat.gob.mx/TimbreFiscalDigital" Version="{version}"
  Fecha="2026-09-07T10:00:00" SubTotal="1000.00" Total="1160.00" Moneda="MXN"
  TipoDeComprobante="I" {attrs_extra}>
  <cfdi:Emisor Rfc="ASU970101AAA" Nombre="ASUR"/>
  <cfdi:Receptor Rfc="TME840315KT6" Nombre="Aero Charter Cancun"/>
  <cfdi:Conceptos><cfdi:Concepto Descripcion="Estacionamiento"/></cfdi:Conceptos>
  <cfdi:Complemento>
    <tfd:TimbreFiscalDigital UUID="{UUID}"/>
  </cfdi:Complemento>
</cfdi:Comprobante>"""


def _parse(xml: str) -> FacturaRecibidaParsed:
    b64 = base64.b64encode(xml.encode("utf-8")).decode("ascii")
    return parse_cfdi(ParseRecibidaRequest(xml_b64=b64, rfcs_propios=["TME840315KT6"]))


def test_cfdi_con_serie_y_folio() -> None:
    res = _parse(_cfdi('Serie="FEACZM" Folio="72128"'))
    assert res.serie == "FEACZM"
    assert res.folio == "72128"
    # Lo demás no cambia.
    assert res.valido is True
    assert res.uuid_fiscal == UUID


def test_cfdi_33_con_serie_y_folio() -> None:
    res = _parse(_cfdi('Serie="A" Folio="0411"', ns="http://www.sat.gob.mx/cfd/3", version="3.3"))
    assert res.serie == "A"
    # El folio es texto: los ceros a la izquierda se conservan.
    assert res.folio == "0411"


def test_cfdi_sin_serie_ni_folio() -> None:
    res = _parse(_cfdi())
    assert res.serie is None
    assert res.folio is None
    assert res.uuid_fiscal == UUID


def test_cfdi_solo_folio_sin_serie() -> None:
    res = _parse(_cfdi('Folio="AB1144717"'))
    assert res.serie is None
    assert res.folio == "AB1144717"


def test_cfdi_serie_y_folio_con_espacios_se_recortan() -> None:
    res = _parse(_cfdi('Serie="  FEACZM " Folio=" 72128  "'))
    assert res.serie == "FEACZM"
    assert res.folio == "72128"


def test_cfdi_serie_y_folio_vacios_o_en_blanco_son_none() -> None:
    res = _parse(_cfdi('Serie="" Folio="   "'))
    assert res.serie is None
    assert res.folio is None


def test_cfdi_32_con_serie_y_folio_en_minuscula() -> None:
    res = _parse(_cfdi('serie="B" folio="15"', ns="http://www.sat.gob.mx/cfd/3", version="3.3"))
    assert res.serie == "B"
    assert res.folio == "15"


def test_schema_aditivo_sin_serie_ni_folio() -> None:
    # Un payload previo (sin los campos) sigue validando: defaults None.
    res = FacturaRecibidaParsed(uuid_fiscal=UUID)
    assert res.serie is None
    assert res.folio is None
    dumped = res.model_dump()
    assert dumped["serie"] is None and dumped["folio"] is None
