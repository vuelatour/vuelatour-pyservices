"""Bitácoras de vuelo por componente (planeador / motor / hélice), SIN
WeasyPrint: en el entorno local no renderiza (faltan libs nativas), así que
se prueban ``_tiras_normalizadas`` y ``_build_html`` directo, y el endpoint
con el render sustituido por un stub.

Contrato (2-sep-2026): el API manda ``tiras`` ya derivadas (una por libro,
cada una en su propia página); el payload LEGADO (``formato`` + ``filas``
planas) sigue aceptándose porque pyservices se despliega ANTES que el API.
"""

import re

from fastapi.testclient import TestClient

from app.config import get_settings
from app.main import app
from app.routers import pdf as pdf_router
from app.schemas.reportes import BitacoraTacoRequest
from app.services.bitacora_taco_pdf import _build_html, _tiras_normalizadas

client = TestClient(app)
TOKEN = "secreto-de-prueba"

_TH = re.compile(r"<th[ >]")
_TABLE = re.compile(r"<table[ >]")


def _fila(
    taco_inicial: float,
    taco_final: float,
    tiempo_inicial: float | None = None,
    tiempo_final: float | None = None,
) -> dict:
    return {
        "fecha": "2026-08-14T15:00:00Z",
        "taco_inicial": taco_inicial,
        "horas": round(taco_final - taco_inicial, 1),
        "taco_final": taco_final,
        "tiempo_inicial": tiempo_inicial,
        "tiempo_final": tiempo_final,
        "ruta": "cun-hol-cun",
    }


def _payload_tiras() -> dict:
    """Caso PEV (monomotor): planeador con base, motor sin ficha, hélice sin
    horas capturadas."""
    return {
        "matricula": "XB-PEV",
        "modelo": "Cessna 206",
        "desde": "2026-08-01",
        "hasta": "2026-08-31",
        "generado": "2026-09-02T14:00:00Z",
        "tiras": [
            {
                "tipo": "PLANEADOR",
                "titulo": "Bitácora de planeador",
                "etiqueta": "Tiempo planeador",
                "nota": "Base del planeador: 5,226.1 h cuando el tacómetro marcaba 151.9",
                "con_tiempo": True,
                "filas": [
                    _fila(343.0, 345.2, 5417.2, 5419.4),
                    _fila(345.2, 346.0, 5419.4, 5420.2),
                ],
            },
            {
                "tipo": "MOTOR",
                "titulo": "Bitácora de motor",
                "etiqueta": "Tiempo motor",
                "nota": (
                    "Tiempo del motor = lectura del tacómetro "
                    "(sin horas del motor capturadas en su ficha)"
                ),
                "con_tiempo": False,
                "filas": [_fila(343.0, 345.2), _fila(345.2, 346.0)],
            },
            {
                "tipo": "HELICE",
                "titulo": "Bitácora de hélice",
                "etiqueta": "Tiempo hélice",
                "nota": (
                    "Sin horas de hélice capturadas: llena las columnas a mano "
                    "(o captúralas en Componentes → hélice)"
                ),
                "con_tiempo": True,
                "filas": [_fila(343.0, 345.2), _fila(345.2, 346.0)],
            },
        ],
    }


def _payload_legado(formato: str | None) -> dict:
    base: dict = {
        "matricula": "XB-ABC",
        "desde": "2026-07-01",
        "hasta": "2026-07-31",
        "generado": "2026-08-01T12:00:00Z",
        "filas": [
            {
                "fecha": "2026-07-10",
                "taco_inicial": 1000.0,
                "horas": 2.2,
                "taco_final": 1002.2,
                "ruta": "cun-mid-cun",
                "helice_inicial": 1234.5,
                "helice_final": 1236.7,
            }
        ],
    }
    if formato is not None:
        base["formato"] = formato
    return base


def _secciones(html: str) -> list[str]:
    """Una entrada por tira (bloques ``<section class="tira …">``)."""
    partes = html.split("<section ")[1:]
    return ["<section " + p for p in partes]


# ---------------------------------------------------------------- payload nuevo


def test_tres_tiras_pintan_tres_tablas_una_por_pagina() -> None:
    req = BitacoraTacoRequest.model_validate(_payload_tiras())
    assert [t.tipo for t in _tiras_normalizadas(req)] == ["PLANEADOR", "MOTOR", "HELICE"]

    html = _build_html(req)
    secciones = _secciones(html)
    assert len(_TABLE.findall(html)) == 3
    assert len(secciones) == 3

    # Título = tira.titulo · matrícula · modelo; rango a la derecha.
    assert "Bitácora de planeador · XB-PEV · Cessna 206" in secciones[0]
    assert "Bitácora de motor · XB-PEV · Cessna 206" in secciones[1]
    assert "Bitácora de hélice · XB-PEV · Cessna 206" in secciones[2]
    assert html.count("01/08/2026 — 31/08/2026") == 3

    # Salto de página en todas menos la primera.
    assert "page-break-before: always" in html
    assert "salto" not in secciones[0].split(">", 1)[0]
    assert 'class="tira ancho-5 salto"' in secciones[1]
    assert 'class="tira ancho-7 salto"' in secciones[2]

    # 7 columnas con tiempo (etiqueta propia), 5 sin tiempo.
    assert len(_TH.findall(secciones[0])) == 7
    assert len(_TH.findall(secciones[1])) == 5
    assert len(_TH.findall(secciones[2])) == 7
    assert "Tiempo planeador<br/>inicial" in secciones[0]
    assert "Tiempo planeador<br/>final" in secciones[0]
    assert "Tiempo hélice<br/>inicial" in secciones[2]
    assert "Tiempo planeador" not in secciones[1]
    assert "Tiempo motor" not in secciones[1]  # sin con_tiempo no hay columnas de tiempo

    # Año del encabezado y pie en cada tira (se recortan por separado).
    assert html.count("Fecha<br/>2026") == 3
    assert html.count("Generado 02/09/2026") == 3


def test_notas_y_tiempos_derivados_presentes() -> None:
    html = _build_html(BitacoraTacoRequest.model_validate(_payload_tiras()))
    secciones = _secciones(html)
    assert "Base del planeador: 5,226.1 h cuando el tacómetro marcaba 151.9" in secciones[0]
    assert "Tiempo del motor = lectura del tacómetro" in secciones[1]
    assert "Sin horas de hélice capturadas" in secciones[2]
    # Tiempos del planeador con separador de miles y 1 decimal.
    assert "<td class='n'>5,417.2</td>" in secciones[0]
    assert "<td class='n'>5,419.4</td>" in secciones[0]
    assert "<td class='n'>5,420.2</td>" in secciones[0]
    assert "5,419.4" not in secciones[1]
    assert "5,419.4" not in secciones[2]


def test_tiempo_none_pinta_guion_para_llenar_a_mano() -> None:
    html = _build_html(BitacoraTacoRequest.model_validate(_payload_tiras()))
    secciones = _secciones(html)
    # Hélice sin base: 2 filas × (inicial + final) = 4 guiones; el tacómetro
    # y las horas siguen pintados.
    assert secciones[2].count("<td class='n'>—</td>") == 4
    assert "<td class='n'>343.0</td>" in secciones[2]
    assert "<td class='n'>2.2</td>" in secciones[2]
    # Planeador con base y motor sin columnas de tiempo: cero guiones.
    assert "<td class='n'>—</td>" not in secciones[0]
    assert "<td class='n'>—</td>" not in secciones[1]


def test_tira_sin_nota_no_pinta_bloque() -> None:
    payload = _payload_tiras()
    payload["tiras"] = [dict(payload["tiras"][0], nota=None)]
    html = _build_html(BitacoraTacoRequest.model_validate(payload))
    assert "class='nota'" not in html


def test_tira_sin_filas_muestra_mensaje() -> None:
    payload = _payload_tiras()
    for t in payload["tiras"]:
        t["filas"] = []
    html = _build_html(BitacoraTacoRequest.model_validate(payload))
    assert html.count("Sin vuelos con tacómetro en el periodo") == 3
    assert "colspan='7'" in _secciones(html)[0]
    assert "colspan='5'" in _secciones(html)[1]
    assert len(_TABLE.findall(html)) == 3


def test_tiras_mandan_sobre_filas_legadas() -> None:
    payload = _payload_tiras()
    payload["formato"] = "MOTOR_HELICE"
    payload["filas"] = _payload_legado("MOTOR_HELICE")["filas"]
    tiras = _tiras_normalizadas(BitacoraTacoRequest.model_validate(payload))
    assert [t.titulo for t in tiras] == [
        "Bitácora de planeador",
        "Bitácora de motor",
        "Bitácora de hélice",
    ]
    assert "motor–hélice" not in _build_html(BitacoraTacoRequest.model_validate(payload))


# --------------------------------------------------------------- payload LEGADO


def test_legado_motor_helice_es_una_tira_de_siete_columnas() -> None:
    req = BitacoraTacoRequest.model_validate(_payload_legado("MOTOR_HELICE"))
    tiras = _tiras_normalizadas(req)
    assert len(tiras) == 1
    t = tiras[0]
    assert t.con_tiempo is True
    assert t.titulo == "Bitácora motor–hélice"
    assert t.etiqueta == "Tiempo hélice"
    assert t.nota is None
    assert (t.filas[0].tiempo_inicial, t.filas[0].tiempo_final) == (1234.5, 1236.7)
    assert t.filas[0].ruta == "cun-mid-cun"

    html = _build_html(req)
    assert len(_TABLE.findall(html)) == 1
    assert len(_TH.findall(html)) == 7
    assert "Bitácora motor–hélice · XB-ABC" in html
    assert "Tiempo hélice<br/>inicial" in html
    assert "<td class='n'>1,234.5</td>" in html
    assert "<td class='n'>1,236.7</td>" in html
    assert "salto" not in html.split("<section ", 1)[1].split(">", 1)[0]


def test_legado_planeador_es_la_tira_de_tacometro_de_cinco_columnas() -> None:
    req = BitacoraTacoRequest.model_validate(_payload_legado("PLANEADOR"))
    tiras = _tiras_normalizadas(req)
    assert len(tiras) == 1
    assert tiras[0].con_tiempo is False
    assert tiras[0].titulo == "Bitácora de tacómetro"
    assert tiras[0].filas[0].tiempo_inicial is None

    html = _build_html(req)
    assert len(_TABLE.findall(html)) == 1
    assert len(_TH.findall(html)) == 5
    assert "Bitácora de tacómetro · XB-ABC" in html
    assert "<td class='n'>1,000.0</td>" in html
    assert "<td class='n'>1,002.2</td>" in html
    # Los tiempos de hélice del payload viejo NO se pintan en formato planeador.
    assert "1,234.5" not in html
    assert "—" not in html.split("<tbody>", 1)[1]


def test_legado_sin_formato_ni_filas_sigue_pintando_la_tira_vacia() -> None:
    # API viejo con periodo sin vuelos: mismo PDF que antes (mensaje), no en blanco.
    req = BitacoraTacoRequest.model_validate({"matricula": "XB-ABC"})
    tiras = _tiras_normalizadas(req)
    assert len(tiras) == 1 and tiras[0].con_tiempo is False
    html = _build_html(req)
    assert "Sin vuelos con tacómetro en el periodo" in html
    assert "colspan='5'" in html
    assert "Bitácora de tacómetro · XB-ABC" in html


# ------------------------------------------------------------------- endpoint


def test_endpoint_sin_token_rechazado(monkeypatch) -> None:
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    res = client.post("/pdf/bitacora-taco", json=_payload_tiras())
    assert res.status_code == 401


def test_endpoint_acepta_payload_nuevo_y_legado(monkeypatch) -> None:
    """El router valida el esquema y entrega el PDF (render sustituido:
    WeasyPrint no corre en local)."""
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    recibidos: list[BitacoraTacoRequest] = []

    def _stub(payload: BitacoraTacoRequest) -> bytes:
        recibidos.append(payload)
        return b"%PDF-stub"

    monkeypatch.setattr(pdf_router, "render_bitacora_taco_pdf", _stub)
    headers = {"X-Internal-Token": TOKEN}

    res = client.post("/pdf/bitacora-taco", json=_payload_tiras(), headers=headers)
    assert res.status_code == 200
    assert res.headers["content-type"] == "application/pdf"
    assert 'filename="bitacora-XB-PEV.pdf"' in res.headers["content-disposition"]
    assert res.content == b"%PDF-stub"
    assert len(recibidos[-1].tiras) == 3

    res = client.post("/pdf/bitacora-taco", json=_payload_legado("MOTOR_HELICE"), headers=headers)
    assert res.status_code == 200
    assert recibidos[-1].tiras == []
    assert recibidos[-1].formato == "MOTOR_HELICE"
    assert len(_tiras_normalizadas(recibidos[-1])) == 1


def test_endpoint_rechaza_tira_incompleta(monkeypatch) -> None:
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    payload = _payload_tiras()
    del payload["tiras"][0]["titulo"]
    res = client.post("/pdf/bitacora-taco", json=payload, headers={"X-Internal-Token": TOKEN})
    assert res.status_code == 422
