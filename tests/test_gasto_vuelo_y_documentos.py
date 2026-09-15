"""Sugerir vuelo de un gasto, compras, vencimientos y CFDI recibido.

Todo determinista: el cliente de Claude está simulado y el CFDI no usa IA.
"""

import base64

import pytest

from app.schemas.compras import CompraExtraerRequest
from app.schemas.gastos import GastoParaMatch, GastoVueloSugerirRequest, VueloCandidato
from app.schemas.recibida import ParseRecibidaRequest
from app.schemas.vencimiento import VencimientoExtraerRequest
from app.services import compras_extract, gasto_vuelo, vencimiento_extract
from app.services.compras_extract import extraer_compra
from app.services.gasto_vuelo import sugerir_vuelo_para_gasto
from app.services.recibida_parse import parse_cfdi
from app.services.vencimiento_extract import extraer_vencimiento

# --- Gasto → vuelo ---------------------------------------------------------

GASTO = GastoParaMatch(
    fecha="2026-09-07",
    monto=125.82,
    categoria="ATERRIZAJE",
    notas="Aeropuerto de Cozumel · plataforma",
    lugar="Cozumel",
)
VUELO_CZM = VueloCandidato(
    vuelo_id="v-1",
    folio=1234,
    # 8-sep 01:00 UTC = 7-sep 20:00 en Cancún: el mismo día del gasto.
    fecha_vuelo="2026-09-08T01:00:00Z",
    matricula="XB-PEV",
    ruta="CUN → CZM → CUN",
    estado="COMPLETADO",
)
VUELO_MID = VueloCandidato(
    vuelo_id="v-2",
    folio=1235,
    fecha_vuelo="2026-09-07T15:00:00-05:00",
    matricula="N990GG",
    ruta="CUN → MID → CUN",
    estado="COMPLETADO",
)


def test_vuelo_del_mismo_dia_cancun_y_aeropuerto_de_la_ruta(claude_fake) -> None:
    cliente = claude_fake(
        gasto_vuelo,
        {"vuelo_id_sugerido": "v-1", "confianza": 0.95, "razon": "voló a CZM"},
    )
    res = sugerir_vuelo_para_gasto(
        GastoVueloSugerirRequest(gasto=GASTO, candidatos=[VUELO_CZM, VUELO_MID])
    )
    assert res.vuelo_id_sugerido == "v-1"
    assert res.confianza == 0.95
    assert any("mismo día" in e for e in res.evidencias)
    assert any("CZM" in e for e in res.evidencias)
    assert "CONTEXTO DE DOMINIO" in cliente.ultima["system"][0]["text"]


def test_aeropuerto_que_la_ruta_no_toca_baja_la_confianza(claude_fake) -> None:
    claude_fake(
        gasto_vuelo,
        {"vuelo_id_sugerido": "v-2", "confianza": 0.9, "razon": "mismo día"},
    )
    res = sugerir_vuelo_para_gasto(
        GastoVueloSugerirRequest(gasto=GASTO, candidatos=[VUELO_CZM, VUELO_MID])
    )
    assert res.vuelo_id_sugerido == "v-2"
    assert res.confianza <= 0.4
    assert any("no lo toca" in e for e in res.evidencias)


def test_vuelo_id_inventado_se_descarta(claude_fake) -> None:
    claude_fake(gasto_vuelo, {"vuelo_id_sugerido": "v-999", "confianza": 0.9, "razon": "x"})
    res = sugerir_vuelo_para_gasto(
        GastoVueloSugerirRequest(gasto=GASTO, candidatos=[VUELO_CZM])
    )
    assert res.vuelo_id_sugerido is None
    assert res.motivo_sin_match is not None


def test_sin_candidatos_no_llama_al_modelo() -> None:
    res = sugerir_vuelo_para_gasto(GastoVueloSugerirRequest(gasto=GASTO, candidatos=[]))
    assert res.vuelo_id_sugerido is None
    assert res.motivo_sin_match is not None


# --- Compras ---------------------------------------------------------------

COMPRA_OK = {
    "proveedor": "Aircraft Spruce",
    "fecha": "2026-09-01",
    "numero_orden": "1234567",
    "tracking": "1Z999",
    "moneda": "USD",
    "lineas": [
        {
            "nombre": "MS20470AD4-4 rivet",
            "numero_parte": "MS20470AD4-4",
            "cantidad": 2,
            "precio_unitario_usd": 10.0,
            "total_usd": 20.0,
        }
    ],
    "subtotal_usd": 20.0,
    "shipping_usd": 15.0,
    "impuestos_usd": 5.0,
    "total_usd": 40.0,
    "confianza": 0.9,
}


def test_compra_valida_conserva_confianza_y_numero_de_orden(claude_fake) -> None:
    claude_fake(compras_extract, COMPRA_OK)
    res = extraer_compra(CompraExtraerRequest(pdf_base64="UERG"))
    assert res.numero_orden == "1234567"
    assert res.tracking == "1Z999"
    assert res.advertencias == []
    assert res.confianza == 0.9


def test_compra_que_no_suma_baja_la_confianza(claude_fake) -> None:
    claude_fake(compras_extract, {**COMPRA_OK, "total_usd": 99.0})
    res = extraer_compra(CompraExtraerRequest(pdf_base64="UERG"))
    assert res.confianza <= 0.3
    assert any("total" in a for a in res.advertencias)


def test_compra_con_linea_incoherente_avisa(claude_fake) -> None:
    payload = {
        **COMPRA_OK,
        "lineas": [{**COMPRA_OK["lineas"][0], "total_usd": 50.0}],
        "subtotal_usd": 50.0,
        "total_usd": 70.0,
    }
    claude_fake(compras_extract, payload)
    res = extraer_compra(CompraExtraerRequest(pdf_base64="UERG"))
    assert any("rivet" in a for a in res.advertencias)


# --- Vencimientos ----------------------------------------------------------

VENC_OK = {
    "matriculas": ["XB-PEV", "N990GG"],
    "matricula": "XB-PEV",
    "tipo_documento": "Póliza de seguro",
    "numero_poliza": "POL-998877",
    "fecha_vigencia": "2026-01-01",
    "fecha_vencimiento": "2027-01-01",
    "emisor": "Aseguradora",
    "confianza": 0.9,
}


def test_vencimiento_poliza_de_flota_conserva_todas_las_matriculas(claude_fake) -> None:
    cliente = claude_fake(vencimiento_extract, VENC_OK)
    res = extraer_vencimiento(
        VencimientoExtraerRequest(
            pdf_base64="UERG",
            matriculas_flota=["XB-PEV", "N990GG"],
            tipos_documento=["Póliza de seguro", "Certificado"],
        )
    )
    assert res.matricula == "XB-PEV"
    assert res.matriculas == ["XB-PEV", "N990GG"]
    assert res.numero_poliza == "POL-998877"
    assert any("varias matrículas" in a for a in res.advertencias)
    assert res.confianza == 0.9
    assert "XB-PEV" in cliente.texto_usuario()


def test_vencimiento_con_fechas_invertidas_baja_la_confianza(claude_fake) -> None:
    claude_fake(
        vencimiento_extract,
        {**VENC_OK, "fecha_vigencia": "2027-01-01", "fecha_vencimiento": "2026-01-01"},
    )
    res = extraer_vencimiento(VencimientoExtraerRequest(pdf_base64="UERG"))
    assert res.confianza <= 0.3
    assert any("no es posterior" in a for a in res.advertencias)


# --- CFDI recibido (sin IA) ------------------------------------------------

CFDI = """<?xml version="1.0" encoding="UTF-8"?>
<cfdi:Comprobante xmlns:cfdi="http://www.sat.gob.mx/cfd/4"
  xmlns:tfd="http://www.sat.gob.mx/TimbreFiscalDigital" Version="4.0"
  Fecha="2026-09-07T10:00:00" SubTotal="1000.00" Total="1160.00" Moneda="MXN"
  TipoDeComprobante="I">
  <cfdi:Emisor Rfc="GODE561231GR8" Nombre="ASA"/>
  <cfdi:Receptor Rfc="TME840315KT6" Nombre="Aero Charter Cancun"/>
  <cfdi:Conceptos><cfdi:Concepto Descripcion="Turbosina"/></cfdi:Conceptos>
  <cfdi:Complemento>
    <tfd:TimbreFiscalDigital UUID="11111111-2222-3333-4444-555555555555"/>
  </cfdi:Complemento>
</cfdi:Comprobante>"""

BOMBA = """<?xml version="1.0"?>
<!DOCTYPE lolz [<!ENTITY lol "lol"><!ENTITY lol2 "&lol;&lol;&lol;">]>
<cfdi:Comprobante xmlns:cfdi="http://www.sat.gob.mx/cfd/4">&lol2;</cfdi:Comprobante>"""


def _b64(texto: str) -> str:
    return base64.b64encode(texto.encode("utf-8")).decode("ascii")


def test_cfdi_valido() -> None:
    res = parse_cfdi(ParseRecibidaRequest(xml_b64=_b64(CFDI), rfcs_propios=["TME840315KT6"]))
    assert res.valido is True
    assert res.uuid_fiscal == "11111111-2222-3333-4444-555555555555"
    assert res.emisor_rfc == "GODE561231GR8"
    assert res.version_cfdi == "4.0"
    assert res.conceptos_n == 1
    assert res.total == 1160.0


def test_cfdi_con_entidades_se_rechaza() -> None:
    with pytest.raises(ValueError, match="DOCTYPE"):
        parse_cfdi(ParseRecibidaRequest(xml_b64=_b64(BOMBA)))


def test_cfdi_de_otro_receptor_no_es_factura_recibida() -> None:
    res = parse_cfdi(ParseRecibidaRequest(xml_b64=_b64(CFDI), rfcs_propios=["GODE561231GR8"]))
    assert res.valido is False
    assert res.motivo and "receptor" in res.motivo


def test_cfdi_sin_timbre_no_es_valido() -> None:
    sin_tfd = CFDI.replace(
        '<tfd:TimbreFiscalDigital UUID="11111111-2222-3333-4444-555555555555"/>', ""
    )
    res = parse_cfdi(ParseRecibidaRequest(xml_b64=_b64(sin_tfd)))
    assert res.valido is False
    assert res.uuid_fiscal is None
    assert res.motivo and "UUID" in res.motivo


def test_cfdi_sin_rfcs_propios_no_valida_receptor() -> None:
    res = parse_cfdi(ParseRecibidaRequest(xml_b64=_b64(CFDI)))
    assert res.valido is True
