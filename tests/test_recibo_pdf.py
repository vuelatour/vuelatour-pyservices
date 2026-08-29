from fastapi.testclient import TestClient

from app.config import get_settings
from app.main import app

client = TestClient(app)

TOKEN = "secreto-de-prueba"


def _payload() -> dict:
    return {
        "folio_recibo": "REC-131-2",
        "cliente": "Cliente Demo S.A.",
        "vuelo_folio": "131",
        "ruta": "CUN → CZM → CUN",
        "fecha_vuelo": "2026-08-20",
        "fecha_cobro": "2026-08-21T15:30:00Z",
        "monto": 25000.0,
        "moneda": "MXN",
        "tc_usd_mxn": 17.5,
        "equivalente_usd": 1428.57,
        "metodo": "Transferencia",
        "cuenta_destino": "HSBC Pesos",
        "referencia": "REF-778899",
        "total_cotizacion_usd": 3596.0,
        "cobrado_a_la_fecha_usd": 2428.57,
        "saldo_pendiente_usd": 1167.43,
        "liquidado": False,
        "notas": "Anticipo acordado con el cliente.",
        "cobros_previos": [
            {"fecha": "2026-08-15", "monto": 1500.0, "moneda": "USD", "etiqueta": "Abono"},
            {"fecha": "2026-08-16", "monto": -500.0, "moneda": "USD", "etiqueta": "Reembolso"},
        ],
    }


def test_recibo_sin_token_rechazado(monkeypatch) -> None:
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    res = client.post("/pdf/recibo", json=_payload())
    assert res.status_code == 401


def test_recibo_genera_documento(monkeypatch) -> None:
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    res = client.post(
        "/pdf/recibo",
        json=_payload(),
        headers={"X-Internal-Token": TOKEN},
    )
    assert res.status_code == 200
    assert res.headers["content-type"] == "application/pdf"
    assert "recibo-REC-131-2.pdf" in res.headers["content-disposition"]
    assert res.content[:4] == b"%PDF"
    assert len(res.content) > 1000


def test_recibo_liquidado_payload_minimo(monkeypatch) -> None:
    """Aditivo: un payload casi vacío (defaults) también renderiza, y el
    sello LIQUIDADO no truena sin historial ni notas."""
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    res = client.post(
        "/pdf/recibo",
        json={"liquidado": True, "monto": 3500.0, "moneda": "USD"},
        headers={"X-Internal-Token": TOKEN},
    )
    assert res.status_code == 200
    assert res.content[:4] == b"%PDF"
