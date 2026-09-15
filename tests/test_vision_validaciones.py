"""Validaciones deterministas de los puntos de visión (15-sep-2026).

Todo con cliente de Claude simulado: aquí se prueba QUÉ hace el servicio con
la respuesta del modelo, no el modelo.
"""

from app.schemas.vision import (
    ConstanciaFiscalRequest,
    GastoTicketRequest,
    TacometroRequest,
)
from app.services import anthropic_vision
from app.services.anthropic_vision import (
    leer_constancia_fiscal,
    leer_tacometro,
    leer_ticket_combustible,
    leer_ticket_gasto,
)

IMG = {"image_base64": "AAAA", "media_type": "image/jpeg"}


def _ticket(**extra) -> GastoTicketRequest:
    return GastoTicketRequest(**IMG, **extra)


# --- Ticket de gasto -------------------------------------------------------

TICKET_OK = {
    "monto": 939.60,
    "moneda": "MXN",
    "fecha": "2026-09-07",
    "proveedor": "Aeropuertos del Sureste",
    "folio": "FEDCUN 193422",
    "concepto": "TUA",
    "categoria_sugerida": "TUAS",
    "medio_pago": "TARJETA_CORP",
    "tarjeta_terminacion": "TARJETA 6256 Debito",
    "conceptos": [
        {"concepto": "TUA NACIONAL", "monto": 810.0},
        {"concepto": "IVA 16%", "monto": 129.6},
    ],
    "matricula": "XB-PEV",
    "rfc_emisor": "GODE561231GR8",
    "iva_monto": 129.6,
    "lugar": "Aeropuerto de Cozumel",
    "confianza": 0.93,
    "legible": True,
    "notas": "",
}


def test_ticket_copia_folio_matricula_y_datos_nuevos(claude_fake) -> None:
    cliente = claude_fake(anthropic_vision, TICKET_OK)
    res = leer_ticket_gasto(_ticket())

    # El folio es la llave anti-duplicados: antes se pedía y nunca se copiaba.
    assert res.folio == "FEDCUN 193422"
    # Las matrículas mexicanas no llevan dígitos: antes se descartaban.
    assert res.matricula == "XB-PEV"
    assert res.rfc_emisor == "GODE561231GR8"
    assert res.iva_monto == 129.6
    assert res.lugar == "Aeropuerto de Cozumel"
    assert res.tarjeta_terminacion == "6256"
    assert res.conceptos and len(res.conceptos) == 2
    assert res.confianza == 0.93
    assert res.advertencias == []
    assert "CONTEXTO DE DOMINIO" in cliente.ultima["system"][0]["text"]


def test_ticket_desglose_que_no_suma_se_descarta_con_aviso(claude_fake) -> None:
    payload = {**TICKET_OK, "conceptos": [{"concepto": "TUA", "monto": 500.0}]}
    claude_fake(anthropic_vision, payload)
    res = leer_ticket_gasto(_ticket())
    assert res.conceptos == []
    assert any("desglose" in a for a in res.advertencias)
    # El TOTAL es el dato principal y sigue siendo confiable.
    assert res.monto == 939.60
    assert res.confianza == 0.93


def test_ticket_monto_como_texto_con_comas(claude_fake) -> None:
    claude_fake(anthropic_vision, {**TICKET_OK, "monto": "1,234.56", "conceptos": []})
    res = leer_ticket_gasto(_ticket())
    assert res.monto == 1234.56


def test_ticket_matricula_y_tarjeta_fuera_de_catalogo_avisan(claude_fake) -> None:
    payload = {**TICKET_OK, "matricula": "XA-ZZZ", "tarjeta_terminacion": "1111"}
    claude_fake(anthropic_vision, payload)
    res = leer_ticket_gasto(
        _ticket(matriculas_flota=["XB-PEV", "N990GG"], terminaciones_validas=["6256", "0577"])
    )
    assert res.matricula == "XA-ZZZ"  # se conserva: la flota puede haber crecido
    assert any("XA-ZZZ" in a for a in res.advertencias)
    assert any("1111" in a for a in res.advertencias)


def test_ticket_catalogos_vivos_viajan_en_el_mensaje(claude_fake) -> None:
    cliente = claude_fake(anthropic_vision, TICKET_OK)
    leer_ticket_gasto(_ticket(matriculas_flota=["XB-PEV"], terminaciones_validas=["6256"]))
    texto = cliente.texto_usuario()
    assert "XB-PEV" in texto and "6256" in texto


def test_ticket_rfc_invalido_no_se_entrega(claude_fake) -> None:
    claude_fake(anthropic_vision, {**TICKET_OK, "rfc_emisor": "GODE561231GR1"})
    res = leer_ticket_gasto(_ticket())
    assert res.rfc_emisor is None
    assert any("verificador" in a for a in res.advertencias)


# --- Combustible -----------------------------------------------------------

COMBUSTIBLE_OK = {
    "litros": 200.0,
    "precio_litro": 25.0,
    "total": 5000.0,
    "moneda": "MXN",
    "tipo_combustible": "AVGAS",
    "fecha": "2026-09-07",
    "proveedor": "ASA",
    "folio": "2622242310",
    "matricula": "XB-ANU",
    "confianza": 0.9,
    "legible": True,
}


def test_combustible_folio_matricula_y_cuadre(claude_fake) -> None:
    claude_fake(anthropic_vision, COMBUSTIBLE_OK)
    res = leer_ticket_combustible(_ticket())
    assert res.folio == "2622242310"
    assert res.matricula == "XB-ANU"
    assert res.advertencias == []
    assert res.confianza == 0.9


def test_combustible_que_no_cuadra_baja_la_confianza(claude_fake) -> None:
    claude_fake(anthropic_vision, {**COMBUSTIBLE_OK, "total": 900.0})
    res = leer_ticket_combustible(_ticket())
    assert res.confianza <= 0.3
    assert any("precio por litro" in a for a in res.advertencias)


def test_combustible_recalcula_los_galones(claude_fake) -> None:
    payload = {
        **COMBUSTIBLE_OK,
        "galones_origen": 50.0,
        "litros": 50.0,  # el modelo olvidó convertir
        "precio_litro": 26.42,
        "total": 5000.0,
    }
    claude_fake(anthropic_vision, payload)
    res = leer_ticket_combustible(_ticket())
    assert res.litros == 189.27
    assert res.galones_origen == 50.0
    assert any("galones" in a for a in res.advertencias)


# --- Tacómetro -------------------------------------------------------------


def test_taco_salto_grande_pero_posible_solo_avisa(claude_fake) -> None:
    """+20 hrs sin captura intermedia es POSIBLE: aviso y MEDIA, nunca BAJA."""
    claude_fake(
        anthropic_vision,
        {"lectura": 1577.3, "confianza": 0.97, "legible": True, "calidad_foto": "ALTA"},
    )
    res = leer_tacometro(TacometroRequest(**IMG, ultimo=1555.8))
    assert res.lectura == 1577.3
    assert res.calidad_foto == "MEDIA"
    assert res.confianza == 0.7
    assert any("captura" in a for a in res.advertencias)


def test_taco_lectura_menor_que_el_ultimo_no_sale_como_segura(claude_fake) -> None:
    claude_fake(
        anthropic_vision,
        {"lectura": 1200.4, "confianza": 0.97, "legible": True, "calidad_foto": "ALTA"},
    )
    res = leer_tacometro(TacometroRequest(**IMG, ultimo=1555.8))
    assert res.confianza <= 0.3
    assert res.calidad_foto == "BAJA"
    assert any("MENOR" in a for a in res.advertencias)


def test_taco_lectura_disparada_avisa(claude_fake) -> None:
    claude_fake(
        anthropic_vision,
        {"lectura": 15558.0, "confianza": 0.99, "legible": True, "calidad_foto": "ALTA"},
    )
    res = leer_tacometro(TacometroRequest(**IMG, ultimo=1555.8))
    assert res.confianza <= 0.3
    assert any("décima" in a for a in res.advertencias)


def test_taco_lectura_normal_no_se_castiga(claude_fake) -> None:
    cliente = claude_fake(
        anthropic_vision,
        {"lectura": 1557.9, "confianza": 0.96, "legible": True, "calidad_foto": "ALTA"},
    )
    res = leer_tacometro(TacometroRequest(**IMG, ultimo=1555.8))
    assert res.confianza == 0.96
    assert res.calidad_foto == "ALTA"
    assert res.advertencias == []
    assert "ANCLA DE MAGNITUD" in cliente.ultima["system"][1]["text"]


# --- Constancia fiscal -----------------------------------------------------


def test_constancia_rfc_con_digito_malo_no_se_entrega(claude_fake) -> None:
    claude_fake(
        anthropic_vision,
        {
            "rfc": "TME840315KT1",
            "razon_social": "TELEFONOS DE MEXICO",
            "regimen_fiscal": "601",
            "cp": "06500",
            "confianza": 0.95,
            "legible": True,
        },
    )
    res = leer_constancia_fiscal(ConstanciaFiscalRequest(**IMG))
    assert res.rfc is None
    assert res.legible is True  # el documento se leyó; solo el RFC no valida
    assert res.razon_social == "TELEFONOS DE MEXICO"
    assert res.confianza <= 0.3
    assert res.motivo and "verificador" in res.motivo
    assert any("verificador" in a for a in res.advertencias)


def test_constancia_valida_pasa_completa(claude_fake) -> None:
    claude_fake(
        anthropic_vision,
        {
            "rfc": "TME840315KT6",
            "razon_social": "TELEFONOS DE MEXICO",
            "regimen_fiscal": "601",
            "cp": "6500",
            "confianza": 0.95,
            "legible": True,
        },
    )
    res = leer_constancia_fiscal(ConstanciaFiscalRequest(**IMG))
    assert res.rfc == "TME840315KT6"
    assert res.cp == "06500"
    assert res.confianza == 0.95
    assert res.advertencias == []
