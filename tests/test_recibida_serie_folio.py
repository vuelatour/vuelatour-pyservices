"""Serie y folio del CFDI recibido (5-oct-2026).

Pedido del cliente: «al momento de la conciliación me apoyan a poner el número
de la factura con la que se enlaza el movimiento… en notas». El número de
factura del proveedor son los atributos `Serie`/`Folio` del
`cfdi:Comprobante`; aquí solo se leen (recortados, vacío ⇒ None). El rótulo
(«A-0411», respaldo al UUID) lo arma el API.

Contrato con el API (revisión 5-oct-2026): las llaves `serie`/`folio` SIEMPRE
viajan en la respuesta HTTP de `/facturacion/parse-recibida`, aunque valgan
null. El cron de relectura del API (`recibidas-releer-folio`) usa su PRESENCIA
para saber que este pyservices ya las lee y sellar `folio_releido_at`; sin
ellas, cada CFDI sin Serie ni Folio se volvería a descargar y parsear cada
10 minutos sin fin. Por eso se prueba la ruta real y no solo el modelo.
"""

from __future__ import annotations

import base64
from collections.abc import Iterator

import pytest
from fastapi import FastAPI
from fastapi.testclient import TestClient

from app.config import get_settings
from app.routers.facturacion import router as facturacion_router
from app.schemas.recibida import FacturaRecibidaParsed, ParseRecibidaRequest
from app.services.recibida_parse import parse_cfdi

UUID = "11111111-2222-3333-4444-555555555555"
TOKEN = "secreto-de-prueba"


def _cfdi(
    attrs_extra: str = "",
    ns: str = "http://www.sat.gob.mx/cfd/4",
    version: str = "4.0",
    version_attr: str = "Version",
) -> str:
    # `version_attr`: los CFDI 3.2 escribían `version` en minúscula (igual que
    # `serie`/`folio`); 3.3 y 4.0, con mayúscula.
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<cfdi:Comprobante xmlns:cfdi="{ns}"
  xmlns:tfd="http://www.sat.gob.mx/TimbreFiscalDigital" {version_attr}="{version}"
  Fecha="2026-09-07T10:00:00" SubTotal="1000.00" Total="1160.00" Moneda="MXN"
  TipoDeComprobante="I" {attrs_extra}>
  <cfdi:Emisor Rfc="ASU970101AAA" Nombre="ASUR"/>
  <cfdi:Receptor Rfc="TME840315KT6" Nombre="Aero Charter Cancun"/>
  <cfdi:Conceptos><cfdi:Concepto Descripcion="Estacionamiento"/></cfdi:Conceptos>
  <cfdi:Complemento>
    <tfd:TimbreFiscalDigital UUID="{UUID}"/>
  </cfdi:Complemento>
</cfdi:Comprobante>"""


def _b64(xml: str) -> str:
    return base64.b64encode(xml.encode("utf-8")).decode("ascii")


def _parse(xml: str) -> FacturaRecibidaParsed:
    return parse_cfdi(ParseRecibidaRequest(xml_b64=_b64(xml), rfcs_propios=["TME840315KT6"]))


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
    # Un 3.2 real: `version`, `serie` y `folio` en minúscula.
    res = _parse(
        _cfdi(
            'serie="B" folio="15"',
            ns="http://www.sat.gob.mx/cfd/3",
            version="3.2",
            version_attr="version",
        )
    )
    assert res.serie == "B"
    assert res.folio == "15"
    assert res.version_cfdi == "3.2"
    # 3.2 ya no está soportada: se avisa, pero serie/folio se leen igual.
    assert any("Versión de CFDI 3.2" in a for a in res.advertencias)


def test_schema_aditivo_sin_serie_ni_folio() -> None:
    # Un payload previo (sin los campos) sigue validando: defaults None.
    res = FacturaRecibidaParsed(uuid_fiscal=UUID)
    assert res.serie is None
    assert res.folio is None
    dumped = res.model_dump()
    assert dumped["serie"] is None and dumped["folio"] is None


# ---------------------------------------------------------------- ruta HTTP


@pytest.fixture
def cliente(monkeypatch: pytest.MonkeyPatch) -> Iterator[TestClient]:
    """TestClient sobre el router REAL de facturación (sin `app.main`, que en
    esta Mac arrastra WeasyPrint)."""
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    app = FastAPI()
    app.include_router(facturacion_router)
    yield TestClient(app)
    get_settings.cache_clear()


def _post(cliente: TestClient, xml: str):
    return cliente.post(
        "/facturacion/parse-recibida",
        json={"xml_b64": _b64(xml), "rfcs_propios": ["TME840315KT6"]},
        headers={"X-Internal-Token": TOKEN},
    )


def test_ruta_sin_serie_ni_folio_emite_las_llaves_en_null(cliente: TestClient) -> None:
    # Candado del contrato con el API: sin Serie ni Folio las llaves VIAJAN
    # (null). Un `response_model_exclude_none=True` en la ruta las quitaría y
    # el cron del API nunca sellaría `folio_releido_at`.
    res = _post(cliente, _cfdi())
    assert res.status_code == 200
    cuerpo = res.json()
    assert "serie" in cuerpo and cuerpo["serie"] is None
    assert "folio" in cuerpo and cuerpo["folio"] is None
    assert cuerpo["uuid_fiscal"] == UUID


def test_ruta_con_serie_y_folio(cliente: TestClient) -> None:
    res = _post(cliente, _cfdi('Serie="FEACZM" Folio="72128"'))
    assert res.status_code == 200
    cuerpo = res.json()
    assert cuerpo["serie"] == "FEACZM"
    assert cuerpo["folio"] == "72128"
