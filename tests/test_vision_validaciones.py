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


def test_ticket_nunca_sugiere_una_comision_del_vendedor(claude_fake) -> None:
    """Regresión del 13-oct-2026 (dos categorías de la comisión del vendedor,
    invariante 61 del API): ni «Comisión del vendedor (VuelaTour)» ni
    «Comisión del vendedor (avión)» se sugieren desde un ticket — `valid_cats`
    no las incluye; si el modelo las devuelve, la sugerencia se descarta y el
    resto de la lectura queda igual (el pago al vendedor lo clasifica oficina)."""
    for categoria in ("COMISION_VENDEDOR_AVION", "COMISION_VENDEDOR"):
        claude_fake(anthropic_vision, {**TICKET_OK, "categoria_sugerida": categoria})
        res = leer_ticket_gasto(_ticket())
        assert res.categoria_sugerida is None, categoria
        assert res.monto == 939.60 and res.folio == "FEDCUN 193422"


def test_ticket_comisariato_se_entrega(claude_fake) -> None:
    """Comisariato (13-oct-2026, pedido del cliente): los snacks/bebidas de
    los PASAJEROS tienen categoría propia COMISARIATO (gasto directo del
    vuelo, columna OTROS del balance del avión). `valid_cats` y el Literal
    del esquema la dejan pasar; antes de este cambio la sugerencia salía
    None. Una categoría inventada sigue descartándose."""
    payload = {
        **TICKET_OK,
        "concepto": "Botanas y agua para pasajeros",
        "categoria_sugerida": "COMISARIATO",
        "conceptos": [],
        "monto": 176.5,
        "proveedor": "OXXO",
    }
    claude_fake(anthropic_vision, payload)
    res = leer_ticket_gasto(_ticket())
    assert res.categoria_sugerida == "COMISARIATO"
    assert res.monto == 176.5 and res.proveedor == "OXXO"
    # Control: una categoría que no existe en el catálogo ⇒ None.
    claude_fake(anthropic_vision, {**payload, "categoria_sugerida": "PISTA_VIP"})
    assert leer_ticket_gasto(_ticket()).categoria_sugerida is None


def test_ticket_prompt_pide_comisariato_para_pasajeros(claude_fake) -> None:
    """El prompt del ticket (system) lista COMISARIATO, lo define para los
    PASAJEROS y el desempate manda las señales de pasajeros a COMISARIATO
    (ya no a OTRO). La frase vieja que mandaba el comisariato a OTRO no
    vuelve. El desempate «sin señas ⇒ tripulación» se conserva."""
    cliente = claude_fake(anthropic_vision, TICKET_OK)
    leer_ticket_gasto(_ticket())
    texto = cliente.texto_system()
    assert "COMIDA, COMISARIATO, HOTEL" in texto
    assert "comisariato para los PASAJEROS" in texto
    assert "es gasto directo del vuelo, NO es OTRO" in texto
    assert "señales de PASAJEROS mandan a COMISARIATO" in texto
    assert "asume TRIPULACIÓN (COMIDA/TAXI/HOTEL)" in texto
    assert "comisariato (snacks/bebidas para los PASAJEROS), mensajería" not in texto
    assert "señales de PASAJEROS mandan a OTRO" not in texto


def test_ticket_paqueteria_se_entrega(claude_fake) -> None:
    """Paquetería (20-oct-2026, pedido del cliente): los envíos de mensajería
    (DHL, FedEx, Estafeta) tienen categoría propia PAQUETERIA (gasto de
    VuelaTour). `valid_cats` y el Literal la dejan pasar; antes caían en
    OTRO. OTROS_DIRECTOS_VUELO no se sugiere pero se tolera (reanálisis)."""
    payload = {
        **TICKET_OK,
        "concepto": "Envío de documentos a Mérida",
        "categoria_sugerida": "PAQUETERIA",
        "conceptos": [],
        "monto": 389.0,
        "proveedor": "DHL Express",
    }
    claude_fake(anthropic_vision, payload)
    res = leer_ticket_gasto(_ticket())
    assert res.categoria_sugerida == "PAQUETERIA"
    assert res.monto == 389.0 and res.proveedor == "DHL Express"
    claude_fake(
        anthropic_vision, {**payload, "categoria_sugerida": "OTROS_DIRECTOS_VUELO"}
    )
    assert leer_ticket_gasto(_ticket()).categoria_sugerida == "OTROS_DIRECTOS_VUELO"


def test_ticket_prompt_pide_paqueteria_para_envios(claude_fake) -> None:
    """El prompt lista PAQUETERIA, la define para mensajería/envíos y OTRO ya
    no menciona la paquetería. OTROS_DIRECTOS_VUELO no se ofrece."""
    cliente = claude_fake(anthropic_vision, TICKET_OK)
    leer_ticket_gasto(_ticket())
    texto = cliente.texto_system()
    assert "PERMISO, FIJO, PAQUETERIA, OTRO" in texto
    assert "PAQUETERIA = envíos de mensajería o paquetería" in texto
    assert "mensajería/paquetería, papelería" not in texto
    assert "OTROS_DIRECTOS_VUELO" not in texto


# --- TUA sistémica (13-oct-2026, API 0.0.75): `tua_monto` leído aparte ------


def test_ticket_tua_monto_se_copia(claude_fake) -> None:
    claude_fake(anthropic_vision, {**TICKET_OK, "tua_monto": 939.6})
    res = leer_ticket_gasto(_ticket())
    assert res.tua_monto == 939.6
    assert res.advertencias == []
    assert res.confianza == 0.93


def test_ticket_tua_monto_mayor_que_el_total_se_descarta(claude_fake) -> None:
    """Dato SECUNDARIO: se limpia con aviso; el total y su confianza intactos."""
    claude_fake(anthropic_vision, {**TICKET_OK, "tua_monto": 1000})
    res = leer_ticket_gasto(_ticket())
    assert res.tua_monto is None
    assert any("TUA leída" in a for a in res.advertencias)
    assert res.advertencias == [
        "La TUA leída (1,000.00) es mayor que el total del documento (939.60): se descartó."
    ]
    assert res.monto == 939.60
    assert res.confianza == 0.93


def test_ticket_sin_tua_monto_sigue_igual(claude_fake) -> None:
    """Un modelo que no manda la llave (o la manda en 0/null) ⇒ None sin aviso;
    el resto de la lectura es la de siempre."""
    claude_fake(anthropic_vision, TICKET_OK)
    base = leer_ticket_gasto(_ticket())
    assert base.tua_monto is None
    assert base.advertencias == []
    for valor in (None, 0, -5, "abc"):
        claude_fake(anthropic_vision, {**TICKET_OK, "tua_monto": valor})
        res = leer_ticket_gasto(_ticket())
        assert res.tua_monto is None, valor
        assert res.model_dump(exclude={"uso_ia"}) == base.model_dump(exclude={"uso_ia"}), valor


def test_ticket_caso_383_tabla_resumen_corrida(claude_fake) -> None:
    """#383 (gasto 8c60e428): la IA llamó «Autobuses» a la Tarifa TUA de la
    tabla resumen de ASUR. Los renglones cuadran y se entregan TAL CUAL (aquí
    no se re-etiqueta nada) y `tua_monto` viaja para que el API separe."""
    payload = {
        **TICKET_OK,
        "monto": 1925.27,
        "categoria_sugerida": "OPERACIONES",
        "conceptos": [
            {"concepto": "Operaciones", "monto": 1365.64},
            {"concepto": "Autobuses", "monto": 559.63},
        ],
        "iva_monto": None,
        "tua_monto": 559.63,
    }
    claude_fake(anthropic_vision, payload)
    res = leer_ticket_gasto(_ticket())
    assert [(c.concepto, c.monto) for c in res.conceptos] == [
        ("Operaciones", 1365.64),
        ("Autobuses", 559.63),
    ]
    assert res.tua_monto == 559.63
    assert res.advertencias == []


def test_ticket_tua_monto_sobrevive_si_se_descartan_conceptos(claude_fake) -> None:
    """Renglones que no suman el total ⇒ se vacían (regla del 26-ago), pero la
    TUA leída aparte se conserva: el panel la enseña como sugerencia."""
    payload = {
        **TICKET_OK,
        "monto": 1925.27,
        "conceptos": [
            {"concepto": "Operaciones", "monto": 1000.0},
            {"concepto": "Tarifa TUA", "monto": 559.63},
        ],
        "iva_monto": None,
        "tua_monto": 559.63,
    }
    claude_fake(anthropic_vision, payload)
    res = leer_ticket_gasto(_ticket())
    assert res.conceptos == []
    assert any("desglose" in a for a in res.advertencias)
    assert res.tua_monto == 559.63


def test_prompt_tabla_resumen_y_tua_monto() -> None:
    texto = anthropic_vision._TICKET_SYSTEM
    assert "MISMA FILA" in texto
    assert '"Tarifa TUA"' in texto
    assert "Migración" in texto
    assert "la etiqueta está corrida: corrígela" in texto
    assert '"tua_monto"' in texto
    assert "con duda, null" in texto
    # La clave nueva va entre "conceptos" y "matricula".
    assert texto.index('"conceptos"') < texto.index('"tua_monto"') < texto.index('"matricula"')


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
