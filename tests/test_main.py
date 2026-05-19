from fastapi.testclient import TestClient

from app.config import get_settings
from app.main import app

client = TestClient(app)

TOKEN = "secreto-de-prueba"


def _payload() -> dict:
    return {
        "folio": 1042,
        "version": 2,
        "fecha": "2026-05-16",
        "cliente": "Cliente Demo S.A.",
        "aeronave": "XB-PEV — Cessna 206",
        "ruta": "CUN → CZM (redondo)",
        "tipo_vuelo": "REDONDO",
        "pasajeros": 3,
        "lineas": [
            {"concepto": "Tiempo de vuelo (1.20 hrs × $750 USD)", "monto_usd": 900.0},
        ],
        "subtotal_usd": 900.0,
        "tuas_usd": 75.0,
        "iva_usd": 156.0,
        "total_usd": 1131.0,
        "tc_usd_mxn": 18.5,
        "notas": "Cotización de prueba.",
    }


def test_health() -> None:
    res = client.get("/health")
    assert res.status_code == 200
    assert res.json() == {"status": "healthy"}


def test_root() -> None:
    res = client.get("/")
    assert res.status_code == 200
    assert res.json()["service"] == "vuelatour-pyservices"


def test_pdf_sin_token_rechazado(monkeypatch) -> None:
    monkeypatch.setenv("SERVICE_TOKEN", TOKEN)
    get_settings.cache_clear()
    res = client.post("/pdf/cotizacion", json=_payload())
    assert res.status_code == 401


def test_pdf_token_invalido_rechazado(monkeypatch) -> None:
    monkeypatch.setenv("SERVICE_TOKEN", TOKEN)
    get_settings.cache_clear()
    res = client.post(
        "/pdf/cotizacion",
        json=_payload(),
        headers={"X-Service-Token": "incorrecto"},
    )
    assert res.status_code == 401


def test_pdf_genera_documento(monkeypatch) -> None:
    monkeypatch.setenv("SERVICE_TOKEN", TOKEN)
    get_settings.cache_clear()
    res = client.post(
        "/pdf/cotizacion",
        json=_payload(),
        headers={"X-Service-Token": TOKEN},
    )
    assert res.status_code == 200
    assert res.headers["content-type"] == "application/pdf"
    assert res.content[:4] == b"%PDF"
    assert len(res.content) > 1000


def _reparto_payload() -> dict:
    return {
        "periodo_desde": "2026-05-01",
        "periodo_hasta": "2026-05-16",
        "generado": "2026-05-16",
        "aviones": [
            {
                "matricula": "XB-PEV",
                "modelo": "Cessna 206",
                "ingresos_cobrado_usd": 12000.0,
                "pendiente_cobro_usd": 1500.0,
                "gastos_directos_usd": 3000.0,
                "gastos_indirectos_usd": 800.0,
                "permisos_usd": 200.0,
                "otros_usd": 500.0,
                "reserva_overhaul_usd": 400.0,
                "saldo_usd": 7100.0,
                "reparto": [
                    {
                        "socio_nombre": "Aerocharter",
                        "porcentaje": 100.0,
                        "monto_usd": 7100.0,
                    }
                ],
            }
        ],
    }


def test_pdf_reparto_genera_documento(monkeypatch) -> None:
    monkeypatch.setenv("SERVICE_TOKEN", TOKEN)
    get_settings.cache_clear()
    res = client.post(
        "/pdf/reparto",
        json=_reparto_payload(),
        headers={"X-Service-Token": TOKEN},
    )
    assert res.status_code == 200
    assert res.headers["content-type"] == "application/pdf"
    assert res.content[:4] == b"%PDF"
    assert len(res.content) > 1000


def test_pdf_reparto_sin_token_rechazado() -> None:
    res = client.post("/pdf/reparto", json=_reparto_payload())
    assert res.status_code == 401
