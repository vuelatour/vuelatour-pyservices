from fastapi.testclient import TestClient

from app.config import get_settings
from app.main import app

client = TestClient(app)

TOKEN = "secreto-de-prueba"


def _payload() -> dict:
    return {
        "folio": "COT-1042",
        "fecha": "2026-05-16",
        "cliente": "Cliente Demo S.A.",
        "origen": "CUN",
        "destino": "CZM",
        "tipo": "REDONDO",
        "pasajeros": 3,
        "escalas": [
            {"orden": 1, "origen": "CUN", "destino": "CZM"},
        ],
        "tiempo_cobrable_hr": 1.2,
        "tarifa_hora_usd": 750.0,
        "subtotal_usd": 900.0,
        "tuas_usd": 75.0,
        "iva_pct": 16.0,
        "iva_usd": 156.0,
        "total_usd": 1131.0,
        "moneda": "USD",
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


def test_cotizacion_sin_token_rechazado(monkeypatch) -> None:
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    res = client.post("/reportes/cotizacion", json=_payload())
    assert res.status_code == 401


def test_cotizacion_token_invalido_rechazado(monkeypatch) -> None:
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    res = client.post(
        "/reportes/cotizacion",
        json=_payload(),
        headers={"X-Internal-Token": "incorrecto"},
    )
    assert res.status_code == 401


def test_cotizacion_genera_documento(monkeypatch) -> None:
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    res = client.post(
        "/reportes/cotizacion",
        json=_payload(),
        headers={"X-Internal-Token": TOKEN},
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
