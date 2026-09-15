"""Validaciones deterministas de las respuestas de IA (sin red, sin Claude)."""

from datetime import date

from app.services._dominio import (
    BLOQUE_DOMINIO,
    MATRICULAS_FLOTA,
    bloque_sistema_dominio,
    sistema_con_dominio,
)
from app.services.validaciones_ia import (
    confianza_calibrada,
    confianza_de,
    dia_cancun,
    dias_entre_cancun,
    matricula_conocida,
    normalizar_matricula,
    normalizar_referencia,
    normalizar_texto_banco,
    num,
    rfc_digito_verificador,
    terminacion_4,
    terminacion_de_referencia,
    validar_fecha_documento,
    validar_identidad_combustible,
    validar_rfc,
    validar_suma_conceptos,
)

# --- Contexto de dominio ---------------------------------------------------


def test_bloque_dominio_trae_flota_y_reglas() -> None:
    for matricula in MATRICULAS_FLOTA:
        assert matricula in BLOQUE_DOMINIO
    assert "DD/MM/AAAA" in BLOQUE_DOMINIO
    assert "UTC−5" in BLOQUE_DOMINIO
    assert "16 %" in BLOQUE_DOMINIO
    assert "SEL TRASPASO ENTRE CUENTAS" in BLOQUE_DOMINIO


def test_sistema_con_dominio_pone_el_dominio_primero_y_cachea() -> None:
    bloques = sistema_con_dominio("PROMPT DEL ENDPOINT")
    assert len(bloques) == 2
    assert bloques[0] == bloque_sistema_dominio()
    assert bloques[0]["cache_control"] == {"type": "ephemeral"}
    assert bloques[1]["text"] == "PROMPT DEL ENDPOINT"
    assert bloques[1]["cache_control"] == {"type": "ephemeral"}


# --- Texto -----------------------------------------------------------------


def test_normalizar_texto_banco_quita_agregador_acentos_y_folios() -> None:
    assert normalizar_texto_banco("MERPAGO*TAQUERÍA 1234567") == "TAQUERIA"
    assert normalizar_texto_banco("SEL TRASPASO ENTRE CUENTAS") == "TRASPASO ENTRE CUENTAS"
    assert normalizar_texto_banco("Aeropuerto de Cozumel") == "AEROPUERTO DE COZUMEL"


def test_normalizar_referencia_solo_alfanumerico() -> None:
    assert normalizar_referencia("00-2583/0577 ") == "0025830577"


# --- Terminación de tarjeta ------------------------------------------------


def test_terminacion_se_extrae_solo_si_empata_con_una_tarjeta_real() -> None:
    validas = ["0577", "6256", "0585"]
    assert terminacion_de_referencia("0025830577", validas) == "0577"
    assert terminacion_de_referencia("9155656256", validas) == "6256"
    # Referencia corta del archivo del 15-sep: no empata ⇒ jamás se inventa.
    assert terminacion_de_referencia("174465", validas) is None
    assert terminacion_de_referencia(None, validas) is None


def test_terminacion_sin_catalogo_solo_de_referencias_largas() -> None:
    assert terminacion_de_referencia("0025830577") == "0577"
    assert terminacion_de_referencia("174465") is None


def test_terminacion_4() -> None:
    assert terminacion_4("**** 4242") == "4242"
    assert terminacion_4(6250) == "6250"
    assert terminacion_4("12") is None


# --- Números ---------------------------------------------------------------


def test_num_acepta_texto_con_formato_de_dinero() -> None:
    assert num("1,234.56") == 1234.56
    assert num("$ 99.90") == 99.90
    assert num(True) is None
    assert num("abc") is None


def test_validar_suma_conceptos() -> None:
    conceptos = [{"concepto": "TUA", "monto": 810.0}, {"concepto": "IVA 16%", "monto": 129.6}]
    assert validar_suma_conceptos(conceptos, 939.60) is None
    aviso = validar_suma_conceptos(conceptos, 1000.0)
    assert aviso is not None and "939.60" in aviso
    assert validar_suma_conceptos([], 939.60) is None


def test_identidad_combustible() -> None:
    # Cuadra.
    assert validar_identidad_combustible(100, 25.0, 2500.0) == (None, False)
    # Precio SIN IVA contra total CON IVA: se informa pero no es error.
    aviso, roto = validar_identidad_combustible(100, 25.0, 2900.0)
    assert roto is False and aviso is not None and "IVA" in aviso
    # No cuadra de ninguna forma: dato principal sospechoso.
    aviso, roto = validar_identidad_combustible(100, 25.0, 900.0)
    assert roto is True and aviso is not None


# --- Fechas ----------------------------------------------------------------


def test_validar_fecha_documento() -> None:
    hoy = date(2026, 9, 15)
    assert validar_fecha_documento("2026-09-07", hoy) == ("2026-09-07", None)
    fecha, aviso = validar_fecha_documento("2027-01-01", hoy)
    assert fecha == "2027-01-01" and aviso is not None and "futuro" in aviso
    fecha, aviso = validar_fecha_documento("2016-09-07", hoy)
    assert aviso is not None and "año" in aviso
    fecha, aviso = validar_fecha_documento("07/09/2026", hoy)
    assert fecha is None and aviso is not None


def test_dia_cancun_no_corre_la_fecha() -> None:
    # 8-sep 01:00 UTC = 7-sep 20:00 en Cancún.
    assert dia_cancun("2026-09-08T01:00:00Z") == date(2026, 9, 7)
    assert dia_cancun("2026-09-07T20:00:00") == date(2026, 9, 7)
    assert dia_cancun("2026-09-07") == date(2026, 9, 7)
    assert dias_entre_cancun("2026-09-08T01:00:00Z", "2026-09-07") == 0


# --- Catálogos -------------------------------------------------------------


def test_matriculas() -> None:
    assert normalizar_matricula("xb-pev") == "XB-PEV"
    assert normalizar_matricula("XBPEV") == "XB-PEV"
    assert normalizar_matricula("N4142R") == "N4142R"
    assert normalizar_matricula("HOLA") is None
    assert matricula_conocida("XBPEV", MATRICULAS_FLOTA) is True
    assert matricula_conocida("XA-ZZZ", MATRICULAS_FLOTA) is False


def test_rfc_digito_verificador_y_validacion() -> None:
    assert rfc_digito_verificador("GODE561231GR8") == "8"  # persona física
    assert rfc_digito_verificador("TME840315KT6") == "6"  # persona moral
    assert validar_rfc("gode561231gr8") == ("GODE561231GR8", None)
    rfc, aviso = validar_rfc("GODE561231GR1")
    assert rfc is None and aviso is not None and "verificador" in aviso
    rfc, aviso = validar_rfc("NO-ES-RFC")
    assert rfc is None and aviso is not None
    assert validar_rfc(None) == (None, None)


# --- Confianza -------------------------------------------------------------


def test_confianza() -> None:
    assert confianza_de(None) == 0.0
    assert confianza_de("0.8") == 0.8
    assert confianza_de(1.7) == 1.0
    assert confianza_calibrada(0.99, True) == 0.3
    assert confianza_calibrada(0.99, False) == 0.99
