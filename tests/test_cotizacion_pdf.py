"""Partes testeables del PDF de cotización SIN WeasyPrint (import perezoso):
título de ruta, tabla de itinerario y numeración del mapa.

Regla 31-ago (tramos ocultos): el API filtra pdf_oculto, RENUMERA 1..N y
manda la ruta visible resuelta en `ruta`; aquí solo se pinta lo que llega —
la numeración del mapa sale de `orden` del payload, nunca de índices propios.
"""

from app.schemas.reportes import CotizacionPdfRequest, MapaPuntoPdf
from app.services.cotizacion_pdf import _build_html, _mapa_svg


def _req(**extra) -> CotizacionPdfRequest:
    base: dict = {
        "folio": "COT-1042",
        "cliente": "Cliente Demo S.A.",
        "origen": "CUN",
        "destino": "CUN",
        # Payload YA renumerado por el API (los tramos ocultos no llegan):
        # visibles 1, 4 y 5 del viaje real → orden 1, 2 y 3.
        "escalas": [
            {"orden": 1, "origen": "CUN", "destino": "AZP"},
            {"orden": 2, "origen": "BZE", "destino": "CZM"},
            {"orden": 3, "origen": "CZM", "destino": "CUN"},
        ],
    }
    base.update(extra)
    return CotizacionPdfRequest(**base)


def test_titulo_usa_la_ruta_visible_del_api() -> None:
    html = _build_html(_req(ruta="CUN → AZP → BZE → CZM → CUN"))
    assert "CUN → AZP → BZE → CZM → CUN" in html


def test_payload_viejo_sin_ruta_conserva_el_walk_local() -> None:
    # Skew tolerante: sin `ruta` (API viejo) el título se arma de las escalas
    # como siempre (primer origen + destinos).
    html = _build_html(_req())
    assert "CUN → AZP → CZM → CUN" in html


def test_tabla_pinta_el_orden_renumerado_del_payload() -> None:
    html = _build_html(_req(ruta="CUN → AZP → BZE → CZM → CUN"))
    assert "<td>1</td><td>CUN → AZP</td>" in html
    assert "<td>2</td><td>BZE → CZM</td>" in html
    assert "<td>3</td><td>CZM → CUN</td>" in html
    # Jamás la posición original de un tramo visible (delataría los ocultos).
    assert "<td>4</td>" not in html
    assert "<td>5</td>" not in html


def test_sin_escalas_el_titulo_cae_a_origen_destino() -> None:
    # Todos los tramos ocultos: el API manda escalas=[] y ruta=None → sin
    # tabla ni mapa; el título degrada al origen→destino del vuelo.
    html = _build_html(_req(escalas=[], ruta=None))
    assert "CUN → CUN" in html
    assert "Itinerario" not in html


def test_mapa_numera_con_el_orden_del_payload() -> None:
    puntos = [
        MapaPuntoPdf(
            orden=2,
            origen_iata="BZE",
            destino_iata="CZM",
            o_lat=17.53,
            o_lon=-88.30,
            d_lat=20.52,
            d_lon=-86.93,
        ),
        MapaPuntoPdf(
            orden=1,
            origen_iata="CUN",
            destino_iata="AZP",
            o_lat=21.04,
            o_lon=-86.87,
            d_lat=19.71,
            d_lon=-90.50,
        ),
    ]
    svg = _mapa_svg(puntos)
    # Badges con el orden RENUMERADO que mandó el API (no índice propio).
    assert ">1</text>" in svg
    assert ">2</text>" in svg
    assert ">3</text>" not in svg
    # Solo aeropuertos visibles en los marcadores.
    for iata in ("CUN", "AZP", "BZE", "CZM"):
        assert f">{iata}</text>" in svg
