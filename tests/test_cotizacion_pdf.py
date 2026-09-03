"""Partes testeables del PDF de cotización SIN WeasyPrint (import perezoso):
título de ruta, tabla de itinerario y numeración del mapa.

Regla 31-ago (tramos ocultos): el API filtra pdf_oculto, RENUMERA 1..N y
manda la ruta visible resuelta en `ruta`; aquí solo se pinta lo que llega —
la numeración del mapa sale de `orden` del payload, nunca de índices propios.

Fecha por tramo (3-sep-2026): `EscalaPdf.fecha` es un DÍA de pared
(YYYY-MM-DD) SOLO para el PDF del cliente; sin hora, sin zona, sin fallback.
"""

import re

from app.schemas.reportes import CotizacionPdfRequest, MapaPuntoPdf
from app.services.cotizacion_pdf import (
    _build_html,
    _fecha_dia,
    _mapa_svg,
    _peninsula_paths,
    _xy,
)


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


# ===== Mapa auto-ajustable (1-sep): modo LOCAL vs AMPLIO =====
# Coordenadas reales (redondeadas) de los aeropuertos usados abajo.
_CUN = (21.0365, -86.8771)
_CZM = (20.5224, -86.9256)
_MID = (20.9370, -89.6577)
_BJX = (20.9935, -101.4808)  # León/Bajío — el caso real del cliente
_AZP = (19.5748, -99.2886)  # Atizapán (Cd. de México)


def _tramo(orden: int, o_iata: str, o: tuple[float, float], d_iata: str, d: tuple[float, float]):
    return MapaPuntoPdf(
        orden=orden,
        origen_iata=o_iata,
        destino_iata=d_iata,
        o_lat=o[0],
        o_lon=o[1],
        d_lat=d[0],
        d_lon=d[1],
    )


def _viewbox(svg: str) -> tuple[float, float, float, float]:
    m = re.search(r'viewBox="([^"]+)"', svg)
    assert m, "el SVG debe traer viewBox"
    a, b, c, d = (float(t) for t in m.group(1).split())
    return a, b, c, d


def _n_paths_tierra(svg: str) -> int:
    """Paths del fondo geográfico (península y, en modo amplio, México)."""
    return svg.count('fill="#e8eef5"')


def _radio_badge(svg: str) -> float:
    m = re.search(r'r="([0-9.]+)" fill="#dc2626"', svg)
    assert m, "debe haber al menos un badge numerado"
    return float(m.group(1))


_N_PENINSULA = len(_peninsula_paths()[0])


def test_mapa_local_ruta_peninsular_intacto() -> None:
    # CUN→MID→CUN cabe en el lienzo peninsular: modo local con zoom a la
    # ruta, SIN el contorno de México y sin salirse de la vista completa.
    svg = _mapa_svg(
        [_tramo(1, "CUN", _CUN, "MID", _MID), _tramo(2, "MID", _MID, "CUN", _CUN)]
    )
    bx0, by0, bw, bh = _viewbox(svg)
    assert _n_paths_tierra(svg) == _N_PENINSULA
    assert bx0 >= -0.01 and by0 >= -0.01
    assert bx0 + bw <= 600.01 and by0 + bh <= 420.01
    # Los puntos quedan dentro del viewBox.
    for lat, lon in (_CUN, _MID):
        x, y = _xy(lon, lat)
        assert bx0 <= x <= bx0 + bw and by0 <= y <= by0 + bh


def test_mapa_local_ruta_ancha_conserva_vista_completa() -> None:
    # Regresión del tope de lejanía: una ruta peninsular ancha (llega hasta
    # Belice) sigue cayendo a la vista completa EXACTA de siempre.
    svg = _mapa_svg(
        [
            _tramo(1, "CUN", _CUN, "AZP", (19.71, -90.50)),
            _tramo(2, "BZE", (17.53, -88.30), "CZM", _CZM),
        ]
    )
    assert _viewbox(svg) == (0.0, 0.0, 600.0, 420.0)
    assert _n_paths_tierra(svg) == _N_PENINSULA


def test_mapa_amplio_contiene_los_puntos_lejanos() -> None:
    # Caso real del cliente: escala en BJX (León) y AZP — antes se salían
    # del cuadro fijo y la ruta se cortaba.
    svg = _mapa_svg(
        [
            _tramo(1, "CUN", _CUN, "BJX", _BJX),
            _tramo(2, "BJX", _BJX, "AZP", _AZP),
            _tramo(3, "AZP", _AZP, "CUN", _CUN),
        ]
    )
    bx0, by0, bw, bh = _viewbox(svg)
    # Vista más amplia que el lienzo peninsular + contorno de México de fondo.
    assert bw > 600
    assert _n_paths_tierra(svg) == _N_PENINSULA + 1
    # TODOS los puntos dentro del viewBox (con aire, no en el borde).
    for lat, lon in (_CUN, _BJX, _AZP):
        x, y = _xy(lon, lat)
        assert bx0 < x < bx0 + bw and by0 < y < by0 + bh
    # Letterbox: se conserva la proporción del lienzo (600:420).
    assert abs(bw / bh - 600 / 420) < 0.01
    # Badges numerados y etiquetas IATA presentes.
    for n in (1, 2, 3):
        assert f">{n}</text>" in svg
    for iata in ("CUN", "BJX", "AZP"):
        assert f">{iata}</text>" in svg


def test_badges_se_ven_igual_de_grandes_en_ambos_modos() -> None:
    # El radio del badge vive en unidades del viewBox: su proporción contra
    # el ancho de la vista debe ser la misma (9/600) en local y en amplio.
    svg_local = _mapa_svg([_tramo(1, "CUN", _CUN, "MID", _MID)])
    svg_amplio = _mapa_svg([_tramo(1, "CUN", _CUN, "BJX", _BJX)])
    for svg in (svg_local, svg_amplio):
        _, _, bw, _ = _viewbox(svg)
        assert abs(_radio_badge(svg) / bw - 9 / 600) < 0.001


def test_bbox_degenerado_lejano_cae_a_modo_local() -> None:
    # Todos los puntos casi en el mismo lugar (aunque sea lejos de la
    # península): modo local — centra y acerca sin viewBox minúsculo.
    cerca = (_BJX[0] + 0.01, _BJX[1] + 0.01)
    svg = _mapa_svg([_tramo(1, "BJX", _BJX, "BJX", cerca)])
    bx0, by0, bw, bh = _viewbox(svg)
    assert _n_paths_tierra(svg) == _N_PENINSULA  # sin contorno de México
    assert bw < 600
    for lat, lon in (_BJX, cerca):
        x, y = _xy(lon, lat)
        assert bx0 <= x <= bx0 + bw and by0 <= y <= by0 + bh


# ===== Fecha por tramo SOLO para el PDF del cliente (3-sep-2026) =====


def test_itinerario_con_fechas_agrega_columna_fecha() -> None:
    # `fecha` es un día de PARED (YYYY-MM-DD) que el API toma de la escala
    # viva (pdf_fecha) — sin hora ni zona. El tramo 3 trae un datetime con
    # zona (defensivo): se imprime SOLO el día tal cual, sin convertir a
    # UTC/Cancún (movería el día) y sin hora.
    html = _build_html(
        _req(
            ruta="CUN → AZP → BZE → CZM → CUN",
            escalas=[
                {"orden": 1, "origen": "CUN", "destino": "AZP", "fecha": "2026-09-03"},
                {"orden": 2, "origen": "BZE", "destino": "CZM"},
                {
                    "orden": 3,
                    "origen": "CZM",
                    "destino": "CUN",
                    "fecha": "2026-09-02T19:00:00-05:00",
                },
            ],
        )
    )
    assert "<thead><tr><th>#</th><th>Tramo</th><th>Fecha</th></tr></thead>" in html
    assert '<td>1</td><td>CUN → AZP</td><td class="fecha">3 sep 2026</td></tr>' in html
    # Tramo sin fecha: guion, nunca un fallback a otra fecha del payload.
    assert '<td>2</td><td>BZE → CZM</td><td class="fecha">—</td></tr>' in html
    assert '<td>3</td><td>CZM → CUN</td><td class="fecha">2 sep 2026</td></tr>' in html
    # Jamás hora ni el formato dd/mm/aaaa de _fecha_legible para el tramo.
    assert "19:00" not in html
    assert "02/09/2026" not in html
    assert "03/09/2026" not in html


def test_itinerario_sin_fechas_no_agrega_columna() -> None:
    # Sin fecha en ningún tramo (API viejo o sin captura): la tabla queda
    # IDÉNTICA a la de siempre — ni encabezado ni celdas de fecha.
    html = _build_html(_req(ruta="CUN → AZP → BZE → CZM → CUN"))
    assert "<thead><tr><th>#</th><th>Tramo</th></tr></thead>" in html
    assert "<td>1</td><td>CUN → AZP</td></tr>" in html
    assert "<th>Fecha</th>" not in html
    assert 'class="fecha"' not in html


def test_fecha_dia_es_mx_sin_zona() -> None:
    assert _fecha_dia("2026-01-01") == "1 ene 2026"
    assert _fecha_dia("2026-12-31") == "31 dic 2026"
    assert _fecha_dia(None) == ""
    assert _fecha_dia("") == ""
    # Texto no parseable: tal cual (escapado), nunca una excepción.
    assert _fecha_dia("basura") == "basura"
    assert _fecha_dia("<x>") == "&lt;x&gt;"
