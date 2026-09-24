"""Lectura del PDF de una factura EMITIDA (24-sep-2026).

`POST /facturacion/leer-pdf-emitida`: Mari suelta el PDF de la factura que
emitió y el registro se prellena. Determinista (pypdf + regex, sin IA). Los
PDFs de estas pruebas son SINTÉTICOS (ReportLab): el PDF real del ejemplo
(Seguros Inbursa → MAQAR MACHINERY) NO se commitea; aquí se reproduce su
LAYOUT —etiquetas en una línea y valores en la siguiente, la basura binaria
del QR entre líneas, la segunda línea «Emisor: 26300 Póliza…», el RFC del PAC
en la cadena original— para congelar las trampas que tiene.
"""

from __future__ import annotations

import base64
import hashlib
import threading
from io import BytesIO

import pytest
from fastapi.testclient import TestClient
from reportlab.lib import pdfencrypt
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

from app.config import get_settings
from app.main import app
from app.services import emitida_pdf
from app.services.emitida_pdf import (
    AVISO_ILEGIBLE,
    AVISO_PAGINAS,
    AVISO_PROTEGIDO,
    AVISO_SIN_TEXTO,
    AVISO_TIEMPO,
    AVISO_VARIOS_UUID,
    PdfNoValidoError,
    extraer_campos,
    leer_pdf_emitida,
)

TOKEN = "secreto-de-prueba"
client = TestClient(app)


class Crudo(str):
    """Línea que va TAL CUAL al content stream (con escapes octales PDF):
    así se meten bytes de control como los que pypdf saca del QR real."""


def _pdf(
    lineas: list[str], *, paginas: int = 1, cifrado: pdfencrypt.StandardEncryption | None = None
) -> bytes:
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=letter, encrypt=cifrado)
    for _ in range(paginas):
        c.setFont("Helvetica", 7)
        y = 770
        for linea in lineas:
            if isinstance(linea, Crudo):
                c._code.append(f"BT /F1 7 Tf 1 0 0 1 30 {y} Tm ({linea}) Tj ET")
            else:
                c.drawString(30, y, linea)
            y -= 11
        c.showPage()
    c.save()
    return buf.getvalue()


def _b64(raw: bytes) -> str:
    return base64.b64encode(raw).decode("ascii")


def _sello(semilla: str, renglones: int = 5) -> list[str]:
    """Base64 con pinta de sello digital (sin guiones: jamás parece UUID)."""
    datos = b"".join(hashlib.sha512(f"{semilla}{i}".encode()).digest() for i in range(6))
    s = base64.b64encode(datos).decode()
    return [s[i * 72 : (i + 1) * 72] for i in range(renglones)]


@pytest.fixture
def token(monkeypatch):
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    yield {"X-Internal-Token": TOKEN}
    get_settings.cache_clear()


# ---------------------------------------------------------------------------
# (1) Formato del ejemplo real (representación impresa Seguros Inbursa)
# ---------------------------------------------------------------------------

UUID_EJEMPLO = "D08B6837-A3B5-45AF-96E1-36F07FBA8FAF"
SELLO = _sello("inbursa")

EJEMPLO = [
    "CFDI",
    "Seguros Inbursa, S.A., Grupo Financiero Inbursa, Avenida Insurgentes Sur 3500, "
    "Col. Peña Pobre, Tlalpan, C.P.14060, Ciudad de México",
    "Serie: AAI Folio:26062737 Tipo Comprobante: INGRESO Folio fiscal: "
    "d08b6837-a3b5-45af-96e1-36f07fba8faf",
    # Basura binaria del QR tal como la entrega pypdf (bytes de control).
    Crudo(r"\200B^^^B\200\001 \023L5Y:uPm"),
    Crudo(r"/&\021\006x-\006mXc^]\0261$;|M4EV'p1"),
    Crudo(r"y\011iii\011y\001iqq\0319aqQA\)11\011I\)YiQ!iA\)Y\0211A!qII1Y"),
    "Emisor: RFC emisor: Régimen fiscal emisor:",
    "SEGUROS INBURSA, S.A., GRUPO FINANCIERO INBURSA SIN9408027L7 601",
    "Receptor: RFC receptor: Régimen fiscal receptor:",
    "MAQAR MACHINERY MMA150622P83 601",
    "Versión: 4.0",
    "Clave de",
    "producto",
    "Clave",
    "unidad Cantidad Unidad Descripción Objeto",
    "Impuesto Importe unitario Importe",
    "total",
    "84131500 IP 1 SERVICIO PRIMA NETA SEGURO",
    "DE CASCOS",
    "02 6,286.80 6,286.80",
    "84131500 IP 1 SERVICIO GASTOS DE",
    "EXPEDICION SEGURO",
    "DE CASCOS",
    "02 50.00 50.00",
    "Subtotal: 6,336.80",
    "IVA 16%: 1,013.89",
    "Prima total: 7,350.69",
    "Importe total con letra: Siete mil trescientos cincuenta dolares 69/100 USD",
    # TRAMPA: otra línea «Emisor:» sin RFC — no pisa al emisor.
    "Emisor: 26300 Póliza: 30359175 CIS: 36893692",
    "Asesor:",
    "Periodo cubierto: Desde el 28/08/2026 hasta el 28/08/2027",
    "No. de serie del certificado digital: 00001000000711043467 "
    "Fecha y hora de expedición: 2026-08-31 17:19:29",
    "No. de Serie del certificado del SAT:00001000000710611637 "
    "Fecha y hora de certificación del CFDI:2026-08-31 17:19:30",
    "Sello digital:",
    *SELLO,
    "Cadena original del complemento de certificación del SAT:",
    # TRAMPA: el RFC del PAC (AUR100128NN3) dentro de la cadena original.
    f"||1.1|d08b6837-a3b5-45af-96e1-36f07fba8faf|2026-08-31T17:19:30|AUR100128NN3|{SELLO[0][:8]}",
    *SELLO[1:3],
    "Régimen Fiscal Receptor: 601",
    "USO CFDI: G03",
    "Código Postal Receptor: 77539",
    "Método De Pago: PPD",
    "Forma de pago: 99",
    "Moneda: USD",
    "Lugar Expedición: 14060",
    "RFC proveedor de certificación: AUR100128NN3",
    "Tipo de exportación: 01",
    "Este documento es una representación impresa de un CFDI",
    "Página 1 de 1",
]


def test_formato_del_ejemplo_real() -> None:
    r = leer_pdf_emitida(_b64(_pdf(EJEMPLO)))
    assert r.texto_extraido is True
    assert r.paginas == 1
    assert r.serie == "AAI"
    assert r.folio == "26062737"
    assert r.uuid == UUID_EJEMPLO  # siempre en MAYÚSCULAS
    assert r.fecha_emision == "2026-08-31"  # expedición, no certificación
    assert r.emisor_rfc == "SIN9408027L7"
    assert r.emisor_nombre == "SEGUROS INBURSA, S.A., GRUPO FINANCIERO INBURSA"
    assert r.receptor_rfc == "MMA150622P83"
    assert r.receptor_nombre == "MAQAR MACHINERY"
    assert r.subtotal == 6336.80
    assert r.iva == 1013.89
    assert r.total == 7350.69  # «Prima total», NO el «Subtotal» ni el importe con letra
    assert r.moneda == "USD"
    assert r.metodo_pago == "PPD"
    assert r.forma_pago == "99"
    assert r.avisos == []


def test_ejemplo_no_toma_el_rfc_del_pac_ni_la_linea_poliza() -> None:
    r = leer_pdf_emitida(_b64(_pdf(EJEMPLO)))
    assert "AUR100128NN3" not in (r.emisor_rfc, r.receptor_rfc)
    assert r.emisor_nombre and "26300" not in r.emisor_nombre
    assert r.total != r.subtotal


def test_linea_emisor_sin_rfc_no_aporta_nada() -> None:
    """La trampa sola: «Emisor: 26300 Póliza: …» no inventa un emisor."""
    campos, _ = extraer_campos(
        "Emisor: 26300 Póliza: 30359175 CIS: 36893692\nTotal: $1,160.00\nMoneda: MXN"
    )
    assert campos["emisor_rfc"] is None
    assert campos["emisor_nombre"] is None
    assert campos["total"] == 1160.0


# ---------------------------------------------------------------------------
# (2) Formato típico de otro PAC
# ---------------------------------------------------------------------------

TIPICO = [
    "FACTURA",
    "Emisor: AERO CHARTER CANCUN SA DE CV RFC: ACC0501017X9",
    "Receptor: MAQAR MACHINERY RFC: MMA150622P83",
    "Serie y Folio: A-123",
    "Folio Fiscal (UUID): 5f2b1c3a-9d4e-4b7a-8c1d-0e2f3a4b5c6d",
    "Fecha de emisión: 24/09/2026 10:15",
    "Cantidad Descripción Valor unitario Importe",
    "1 Vuelo privado CUN-MID-CUN $1,000.00 $1,000.00",
    "Subtotal: $1,000.00",
    "IVA 16%: $160.00",
    "Total: $1,160.00",
    "Importe con letra: UN MIL CIENTO SESENTA PESOS 00/100 M.N.",
    "Moneda: MXN - Peso Mexicano",
    "Método de pago: PUE - Pago en una sola exhibición",
    "Forma de pago: 03 - Transferencia electrónica de fondos",
    "RFC proveedor de certificación: SAT970701NN3",
    "||1.1|5F2B1C3A-9D4E-4B7A-8C1D-0E2F3A4B5C6D|2026-09-24T10:16:00|SAT970701NN3|abc|",
]


def test_formato_tipico() -> None:
    r = leer_pdf_emitida(_b64(_pdf(TIPICO)))
    assert r.texto_extraido is True
    assert (r.serie, r.folio) == ("A", "123")
    assert r.uuid == "5F2B1C3A-9D4E-4B7A-8C1D-0E2F3A4B5C6D"
    assert r.fecha_emision == "2026-09-24"
    assert (r.emisor_rfc, r.emisor_nombre) == ("ACC0501017X9", "AERO CHARTER CANCUN SA DE CV")
    assert (r.receptor_rfc, r.receptor_nombre) == ("MMA150622P83", "MAQAR MACHINERY")
    assert (r.subtotal, r.iva, r.total) == (1000.0, 160.0, 1160.0)
    assert r.moneda == "MXN"
    assert r.metodo_pago == "PUE"
    assert r.forma_pago == "03"
    assert r.avisos == []


BLOQUES = [
    "AERO CHARTER CANCUN SA DE CV",
    "RFC: ACC0501017X9",
    "Régimen fiscal: 601 General de Ley Personas Morales",
    "FACTURA",
    "Serie: F",
    "Folio: 00045",
    "Receptor",
    "Nombre: MAQAR MACHINERY",
    "RFC: MMA150622P83",
    "Uso CFDI: G03",
    "UUID: 7A1B2C3D-4E5F-4A6B-8C7D-9E0F1A2B3C4D",
    "Fecha: 2026-09-20T09:30:00",
    # pypdf separa en renglones los objetos de texto de una tabla: etiquetas
    # apiladas y montos apilados debajo.
    "Subtotal:",
    "IVA 16%:",
    "Total:",
    "$1,000.00",
    "$160.00",
    "$1,160.00",
    "Moneda: USD - Dólar americano",
    "Método de pago: PPD - Pago en parcialidades o diferido",
    "Forma de pago: Por definir",
]


def test_formato_en_bloques_y_totales_en_columna() -> None:
    r = leer_pdf_emitida(_b64(_pdf(BLOQUES)))
    assert (r.serie, r.folio) == ("F", "00045")
    assert r.uuid == "7A1B2C3D-4E5F-4A6B-8C7D-9E0F1A2B3C4D"
    assert r.fecha_emision == "2026-09-20"
    assert (r.receptor_rfc, r.receptor_nombre) == ("MMA150622P83", "MAQAR MACHINERY")
    # El «RFC:» suelto del encabezado es el emisor (distinto del receptor).
    assert r.emisor_rfc == "ACC0501017X9"
    assert (r.subtotal, r.iva, r.total) == (1000.0, 160.0, 1160.0)
    assert (r.moneda, r.metodo_pago, r.forma_pago) == ("USD", "PPD", "99")


# ---------------------------------------------------------------------------
# (3) PDF sin texto (escaneado)
# ---------------------------------------------------------------------------


def test_pdf_sin_texto_es_escaneado() -> None:
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=letter)
    c.rect(50, 50, 400, 600, fill=1)  # «la foto» de la factura: sin capa de texto
    c.showPage()
    c.save()
    r = leer_pdf_emitida(_b64(buf.getvalue()))
    assert r.texto_extraido is False
    assert r.paginas == 1
    assert AVISO_SIN_TEXTO in r.avisos
    assert all(getattr(r, k) is None for k in emitida_pdf.CAMPOS)


# ---------------------------------------------------------------------------
# (4) No es un PDF ⇒ 422
# ---------------------------------------------------------------------------


def test_no_pdf_es_422(token) -> None:
    res = client.post(
        "/facturacion/leer-pdf-emitida",
        json={"pdf_b64": _b64(b"<html>no soy un pdf</html>")},
        headers=token,
    )
    assert res.status_code == 422
    assert res.json()["detail"] == "El archivo no es un PDF"


def test_base64_invalido_es_422(token) -> None:
    res = client.post(
        "/facturacion/leer-pdf-emitida", json={"pdf_b64": "esto%no#es$base64"}, headers=token
    )
    assert res.status_code == 422


def test_pdf_demasiado_grande_es_422(monkeypatch) -> None:
    monkeypatch.setattr(emitida_pdf, "MAX_BYTES", 1000)
    with pytest.raises(PdfNoValidoError, match="más de 10 MB"):
        leer_pdf_emitida(_b64(_pdf(TIPICO)))


def test_tolera_prefijo_data_url_y_saltos_de_linea() -> None:
    b64 = _b64(_pdf(TIPICO))
    envuelto = "data:application/pdf;base64," + "\n".join(
        b64[i : i + 76] for i in range(0, len(b64), 76)
    )
    assert leer_pdf_emitida(envuelto).folio == "123"


# ---------------------------------------------------------------------------
# (5) Dos UUID distintos sin etiqueta
# ---------------------------------------------------------------------------


def test_dos_uuid_sin_etiqueta_no_adivina() -> None:
    lineas = [
        "Factura de servicios aéreos",
        "Referencia: 11111111-2222-4333-8444-555555555555",
        "Referencia: AAAAAAAA-BBBB-4CCC-8DDD-EEEEEEEEEEEE",
        "Total: $1,160.00",
    ]
    r = leer_pdf_emitida(_b64(_pdf(lineas)))
    assert r.uuid is None
    assert AVISO_VARIOS_UUID in r.avisos
    assert r.total == 1160.0


def test_uuid_etiquetado_gana_y_el_relacionado_no_cuenta() -> None:
    campos, avisos = extraer_campos(
        "CFDI relacionado (04 sustitución) UUID: 11111111-2222-4333-8444-555555555555\n"
        "Folio fiscal: aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee\n"
    )
    assert campos["uuid"] == "AAAAAAAA-BBBB-4CCC-8DDD-EEEEEEEEEEEE"
    assert avisos == []


def test_bloque_de_cfdi_relacionados_no_se_confunde_con_el_propio() -> None:
    """Re-emisión (04 sustitución): el título del bloque va ARRIBA y el UUID de
    la factura anterior abajo; el propio solo viene en la cadena del timbre."""
    campos, avisos = extraer_campos(
        "CFDI relacionados · Tipo de relación: 04\n"
        "UUID: 11111111-2222-4333-8444-555555555555\n"
        "Total: $1,160.00\n"
        "||1.1|aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee|2026-09-24T10:16:00|SAT970701NN3|x|\n"
    )
    assert campos["uuid"] == "AAAAAAAA-BBBB-4CCC-8DDD-EEEEEEEEEEEE"
    assert avisos == []


def test_uuid_cortado_en_el_guion_por_el_renglon() -> None:
    campos, _ = extraer_campos("Folio fiscal: D08B6837-A3B5-45AF-\n96E1-36F07FBA8FAF")
    assert campos["uuid"] == UUID_EJEMPLO


# ---------------------------------------------------------------------------
# (6) Token
# ---------------------------------------------------------------------------


def test_endpoint_sin_token_es_401(token) -> None:
    res = client.post("/facturacion/leer-pdf-emitida", json={"pdf_b64": _b64(_pdf(TIPICO))})
    assert res.status_code == 401


def test_endpoint_con_token_devuelve_el_contrato(token) -> None:
    res = client.post(
        "/facturacion/leer-pdf-emitida", json={"pdf_b64": _b64(_pdf(EJEMPLO))}, headers=token
    )
    assert res.status_code == 200
    cuerpo = res.json()
    assert set(cuerpo) == {*emitida_pdf.CAMPOS, "texto_extraido", "paginas", "avisos"}
    assert cuerpo["uuid"] == UUID_EJEMPLO
    assert cuerpo["total"] == 7350.69
    assert cuerpo["texto_extraido"] is True


# ---------------------------------------------------------------------------
# (7) pypdf falla o se pasa del tope de tiempo ⇒ 200 + aviso, nunca 500
# ---------------------------------------------------------------------------


def test_pypdf_lanza_devuelve_aviso(monkeypatch, token) -> None:
    def revienta(_raw: bytes):
        raise RuntimeError("xref roto")

    monkeypatch.setattr(emitida_pdf, "_extraer_texto", revienta)
    res = client.post(
        "/facturacion/leer-pdf-emitida", json={"pdf_b64": _b64(_pdf(TIPICO))}, headers=token
    )
    assert res.status_code == 200
    assert res.json()["texto_extraido"] is False
    assert res.json()["avisos"] == [AVISO_ILEGIBLE]


def test_pypdf_excede_el_tope_de_tiempo(monkeypatch, token) -> None:
    liberar = threading.Event()

    def atorado(_raw: bytes):
        liberar.wait(5)
        raise AssertionError("no debió esperarse")

    monkeypatch.setattr(emitida_pdf, "_extraer_texto", atorado)
    monkeypatch.setattr(emitida_pdf, "TOPE_SEGUNDOS", 0.05)
    try:
        res = client.post(
            "/facturacion/leer-pdf-emitida", json={"pdf_b64": _b64(_pdf(TIPICO))}, headers=token
        )
    finally:
        liberar.set()
    assert res.status_code == 200
    assert res.json()["texto_extraido"] is False
    assert res.json()["avisos"] == [AVISO_TIEMPO]


def test_interpretar_el_texto_tambien_tiene_tope(monkeypatch) -> None:
    """Un texto patológico (miles de etiquetas apiladas) no deja al endpoint
    colgado: la interpretación corre dentro del MISMO tope que pypdf."""
    liberar = threading.Event()

    def atorado(_texto: str):
        liberar.wait(5)
        raise AssertionError("no debió esperarse")

    monkeypatch.setattr(emitida_pdf, "extraer_campos", atorado)
    monkeypatch.setattr(emitida_pdf, "TOPE_SEGUNDOS", 0.05)
    try:
        r = leer_pdf_emitida(_b64(_pdf(TIPICO)))
    finally:
        liberar.set()
    assert r.texto_extraido is True
    assert r.total is None
    assert r.avisos == [AVISO_TIEMPO]


def test_pdf_malformado_no_es_500() -> None:
    r = leer_pdf_emitida(_b64(b"%PDF-1.7\n1 0 obj << /Type /Catalog >> basura sin xref"))
    assert r.texto_extraido is False
    assert r.avisos  # siempre explica por qué


def test_error_al_interpretar_el_texto_no_es_500(monkeypatch) -> None:
    def revienta(_texto: str):
        raise ValueError("regex rara")

    monkeypatch.setattr(emitida_pdf, "extraer_campos", revienta)
    r = leer_pdf_emitida(_b64(_pdf(TIPICO)))
    assert r.texto_extraido is True
    assert r.total is None
    assert r.avisos == [emitida_pdf.AVISO_CAMPOS]


# ---------------------------------------------------------------------------
# Cifrado y páginas
# ---------------------------------------------------------------------------


def test_pdf_con_contrasena_avisa() -> None:
    cifrado = pdfencrypt.StandardEncryption("clave-del-cliente", ownerPassword="dueno")
    r = leer_pdf_emitida(_b64(_pdf(TIPICO, cifrado=cifrado)))
    assert r.texto_extraido is False
    assert r.avisos == [AVISO_PROTEGIDO]


def test_pdf_cifrado_sin_contrasena_de_usuario_se_lee() -> None:
    """Muchos PAC cifran solo para impedir copiar/imprimir (usuario vacío)."""
    cifrado = pdfencrypt.StandardEncryption("", ownerPassword="dueno", canCopy=0)
    r = leer_pdf_emitida(_b64(_pdf(TIPICO, cifrado=cifrado)))
    assert r.texto_extraido is True
    assert r.folio == "123"


def test_solo_lee_las_primeras_cinco_paginas() -> None:
    r = leer_pdf_emitida(_b64(_pdf(TIPICO, paginas=7)))
    assert r.paginas == 7
    assert AVISO_PAGINAS in r.avisos
    assert r.uuid == "5F2B1C3A-9D4E-4B7A-8C1D-0E2F3A4B5C6D"  # el mismo UUID en todas


# ---------------------------------------------------------------------------
# Reglas finas del extractor (texto directo)
# ---------------------------------------------------------------------------


def test_folio_fiscal_no_es_el_folio_ni_el_certificado_la_serie() -> None:
    campos, _ = extraer_campos(
        "Folio fiscal: D08B6837-A3B5-45AF-96E1-36F07FBA8FAF\n"
        "No. de Serie: 00001000000711043467\n"
        "Folio: B-0045\n"
    )
    assert campos["uuid"] == UUID_EJEMPLO
    assert (campos["serie"], campos["folio"]) == ("B", "0045")


def test_montos_en_una_fila_de_etiquetas() -> None:
    campos, avisos = extraer_campos("Subtotal: IVA 16%: Total:\n2,500.00 400.00 2,900.00")
    assert (campos["subtotal"], campos["iva"], campos["total"]) == (2500.0, 400.0, 2900.0)
    assert avisos == []


def test_monto_en_la_linea_siguiente_solo_si_es_uno() -> None:
    campos, _ = extraer_campos("Uso CFDI: G03 Total:\n$1,160.00 MXN")
    assert campos["total"] == 1160.0
    # Con dos montos debajo no se sabe cuál es el del total: mejor vacío.
    campos, _ = extraer_campos("Total:\n1,000.00 1,160.00")
    assert campos["total"] is None


def test_etiquetas_apiladas_sin_sus_montos_no_adivinan() -> None:
    """Tres etiquetas apiladas y solo dos montos: ninguna se lleva el de otra."""
    campos, _ = extraer_campos("Subtotal:\nIVA 16%:\nTotal:\n1,000.00\n160.00\nGracias")
    assert (campos["subtotal"], campos["iva"], campos["total"]) == (None, None, None)


def test_tasa_y_porcentajes_no_son_montos() -> None:
    campos, _ = extraer_campos(
        "002 IVA Tasa 0.160000 160.00\nIVA Retenido 4%: 40.00\nIVA (16%): $160.00\n"
        "Subtotal: $1,000.00\nTotal impuestos trasladados: $160.00\nTotal: $1,120.00"
    )
    assert campos["iva"] == 160.0
    assert campos["total"] == 1120.0


def test_montos_que_no_cuadran_avisan() -> None:
    campos, avisos = extraer_campos("Subtotal: 1,000.00\nIVA 16%: 160.00\nTotal: 1,610.00")
    assert campos["total"] == 1610.0
    assert any("no da el total impreso" in a and "$1,610" in a for a in avisos)


def test_retenciones_impresas_explican_el_descuadre() -> None:
    _, avisos = extraer_campos(
        "Subtotal: 1,000.00\nIVA 16%: 160.00\nIVA retenido: 106.67\nTotal: 1,053.33"
    )
    assert avisos == []


def test_moneda_no_soportada_avisa() -> None:
    campos, avisos = extraer_campos("Moneda: EUR - Euro\nTotal: 100.00")
    assert campos["moneda"] is None
    assert "Moneda EUR: el registro solo maneja MXN y USD." in avisos


def test_moneda_xxx_y_por_nombre() -> None:
    assert extraer_campos("Moneda: XXX")[0]["moneda"] == "MXN"
    assert extraer_campos("Moneda: Peso Mexicano")[0]["moneda"] == "MXN"
    assert extraer_campos("Tipo de moneda: Dólar americano")[0]["moneda"] == "USD"


def test_moneda_nacional_del_importe_con_letra_no_confunde() -> None:
    campos, avisos = extraer_campos("(MIL PESOS 00/100 MONEDA NACIONAL)\nMoneda: USD")
    assert campos["moneda"] == "USD"
    assert avisos == []


def test_fechas_en_otros_formatos() -> None:
    assert extraer_campos("Fecha de emisión: 24/09/2026 10:15")[0]["fecha_emision"] == (
        "2026-09-24"
    )
    assert (
        extraer_campos("Fecha de expedición: 5 de septiembre de 2026")[0]["fecha_emision"]
        == "2026-09-05"
    )
    assert extraer_campos("Fecha: 2026-09-20T09:30:00")[0]["fecha_emision"] == "2026-09-20"


def test_forma_y_metodo_tolerantes() -> None:
    campos, _ = extraer_campos(
        "Método de pago: Pago en una sola exhibición Forma de pago: Transferencia (03)"
    )
    assert campos["metodo_pago"] == "PUE"
    assert campos["forma_pago"] == "03"


def test_egreso_avisa_que_no_es_factura_de_ingreso() -> None:
    _, avisos = extraer_campos("Tipo de comprobante: E - Egreso\nTotal: 100.00")
    assert any("tipo Egreso" in a for a in avisos)
    _, avisos = extraer_campos("Tipo Comprobante: INGRESO\nTotal: 100.00")
    assert avisos == []


def test_rfc_de_etiquetas_en_la_misma_linea() -> None:
    campos, _ = extraer_campos(
        "RFC emisor: ACC0501017X9 Nombre emisor: AERO CHARTER CANCUN SA DE CV\n"
        "RFC receptor: XAXX010101000 Nombre receptor: PUBLICO EN GENERAL\n"
    )
    assert (campos["emisor_rfc"], campos["emisor_nombre"]) == (
        "ACC0501017X9",
        "AERO CHARTER CANCUN SA DE CV",
    )
    assert (campos["receptor_rfc"], campos["receptor_nombre"]) == (
        "XAXX010101000",
        "PUBLICO EN GENERAL",
    )


FACTURA_SAT = """RFC emisor: ACC0501017X9
Nombre emisor: AERO CHARTER CANCUN
Folio: 123
RFC receptor: MMA150622P83
Nombre receptor: MAQAR MACHINERY
Código postal del receptor: 77539 Régimen fiscal receptor: General de Ley Personas Morales
Folio fiscal: 5F2B1C3A-9D4E-4B7A-8C1D-0E2F3A4B5C6D
No. de serie del CSD: 00001000000504465028
Serie: A
Código postal, fecha y hora de emisión: 77500 2026-09-24 10:15:00
Efecto de comprobante: Ingreso
Impuesto Tipo Base Tipo Factor Tasa o Cuota Importe
IVA Traslado 1,000.00 Tasa 16.00% 160.00
Moneda: Peso Mexicano Subtotal $1,000.00
Forma de pago: Transferencia electrónica de fondos Impuestos trasladados IVA 16.00% $160.00
Método de pago: Pago en una sola exhibición Total $1,160.00
||1.1|5F2B1C3A-9D4E-4B7A-8C1D-0E2F3A4B5C6D|2026-09-24T10:16:00|SPR190613I52|abc==|0001||
RFC del proveedor de certificación: SPR190613I52
"""


def test_formato_de_la_factura_del_sat() -> None:
    """Generador gratuito del SAT: etiquetas y montos mezclados en la misma
    línea y «Efecto de comprobante» en vez de «Tipo»."""
    campos, avisos = extraer_campos(FACTURA_SAT)
    assert campos == {
        "serie": "A",
        "folio": "123",
        "uuid": "5F2B1C3A-9D4E-4B7A-8C1D-0E2F3A4B5C6D",
        "fecha_emision": "2026-09-24",
        "emisor_rfc": "ACC0501017X9",
        "emisor_nombre": "AERO CHARTER CANCUN",
        "receptor_rfc": "MMA150622P83",
        "receptor_nombre": "MAQAR MACHINERY",
        "subtotal": 1000.0,
        "iva": 160.0,  # no la BASE 1,000.00 del renglón «IVA Traslado 1,000.00 …»
        "total": 1160.0,
        "moneda": "MXN",
        "metodo_pago": "PUE",
        "forma_pago": "03",
    }
    assert avisos == []
    egreso = FACTURA_SAT.replace("comprobante: Ingreso", "comprobante: Egreso")
    _, avisos = extraer_campos(egreso)
    assert any("tipo Egreso" in a for a in avisos)


def test_bloque_cliente_con_nombre_en_la_linea_siguiente() -> None:
    campos, _ = extraer_campos(
        "AERO CHARTER CANCUN SA DE CV\nR.F.C. ACC0501017X9\nCliente\nMAQAR MACHINERY\n"
        "RFC: MMA150622P83\nUso del CFDI: G03\nTotal $2,320.00"
    )
    assert (campos["receptor_rfc"], campos["receptor_nombre"]) == (
        "MMA150622P83",
        "MAQAR MACHINERY",
    )
    assert campos["emisor_rfc"] == "ACC0501017X9"
    assert campos["total"] == 2320.0


def test_texto_con_basura_de_control_se_normaliza() -> None:
    campos, _ = extraer_campos("Serie:\x01 AAI\x80 Folio:\x13 26062737\r\nMoneda:\tUSD")
    assert (campos["serie"], campos["folio"], campos["moneda"]) == ("AAI", "26062737", "USD")


# ---------------------------------------------------------------------------
# Revisión adversaria (24-sep-2026): trampas que sí mordían
# ---------------------------------------------------------------------------


def test_encabezado_moneda_iva_no_corta_la_busqueda_de_moneda() -> None:
    """«CANTIDAD UNIDAD MONEDA IVA IMPORTE» (encabezado de conceptos) daba
    «Moneda IVA: el registro solo maneja…» y jamás llegaba a «MONEDA: MXN»."""
    campos, avisos = extraer_campos(
        "CANTIDAD UNIDAD MONEDA IVA IMPORTE\nTOTAL: $1,160.00\nMONEDA: MXN PESO MEXICANO"
    )
    assert campos["moneda"] == "MXN"
    assert avisos == []


def test_dolares_abreviados_son_usd() -> None:
    assert extraer_campos("Moneda: DLS")[0]["moneda"] == "USD"
    assert extraer_campos("Moneda: DLLS")[0]["moneda"] == "USD"
    campos, _ = extraer_campos("Total: US$ 1,160.00\nMoneda: USD")
    assert campos["total"] == 1160.0


def test_retencion_de_iva_no_es_el_iva_y_cuadra() -> None:
    campos, avisos = extraer_campos(
        "Subtotal: 1,000.00\nIVA 16%: 160.00\nRetención IVA: 106.67\n"
        "Retención ISR: 100.00\nTotal: 953.33"
    )
    assert (campos["iva"], campos["total"]) == (160.0, 953.33)
    assert avisos == []  # 1,000 + 160 − 106.67 − 100 = 953.33


def test_retencion_isr_sola_explica_el_descuadre() -> None:
    _, avisos = extraer_campos(
        "Subtotal: 1,000.00\nIVA 16%: 160.00\nRetención ISR (10%): 100.00\nTotal: 1,060.00"
    )
    assert avisos == []


def test_importe_total_de_un_concepto_no_le_gana_al_total_que_cuadra() -> None:
    campos, avisos = extraer_campos(
        "Cantidad Descripción Importe total: 1,000.00\n"
        "Subtotal: 1,000.00\nIVA: 160.00\nTotal: 1,160.00"
    )
    assert campos["total"] == 1160.0
    assert avisos == []


def test_total_que_no_cuadra_con_nada_se_conserva_y_avisa() -> None:
    campos, avisos = extraer_campos("Subtotal: 1,000.00\nIVA: 160.00\nImporte total: 1,500.00")
    assert campos["total"] == 1500.0
    assert any("no da el total impreso" in a for a in avisos)


def test_folio_con_dos_puntos_gana_a_folio_interno() -> None:
    campos, _ = extraer_campos("Folio interno 555\nSerie: Z\nFolio: 777")
    assert (campos["serie"], campos["folio"]) == ("Z", "777")


def test_serie_vacia_no_toma_la_palabra_de_la_linea_siguiente() -> None:
    campos, _ = extraer_campos("Serie:\nFecha 2026-09-24\nFolio: 457")
    assert (campos["serie"], campos["folio"]) == (None, "457")
    # …pero si la línea siguiente es SOLO la serie, sí.
    campos, _ = extraer_campos("Serie:\nAAI\nFolio: 457")
    assert (campos["serie"], campos["folio"]) == ("AAI", "457")


def test_fecha_con_diagonales_anio_primero() -> None:
    assert extraer_campos("Fecha de emisión: 2026/09/24 10:15")[0]["fecha_emision"] == (
        "2026-09-24"
    )
