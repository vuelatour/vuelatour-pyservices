"""Partes testeables del PDF de cotización SIN WeasyPrint (import perezoso):
título de ruta, tabla de itinerario y numeración del mapa.

Regla 31-ago (tramos ocultos): el API filtra pdf_oculto, RENUMERA 1..N y
manda la ruta visible resuelta en `ruta`; aquí solo se pinta lo que llega —
la numeración del mapa sale de `orden` del payload, nunca de índices propios.

Fecha por tramo (3-sep-2026): `EscalaPdf.fecha` es un DÍA de pared
(YYYY-MM-DD) SOLO para el PDF del cliente; sin hora, sin zona, sin fallback.
"""

import base64
import hashlib
import re
import sys

from fastapi.testclient import TestClient

from app.config import get_settings
from app.main import app
from app.routers import reportes as reportes_router
from app.schemas.reportes import CotizacionPdfRequest, MapaPuntoPdf
from app.services.cotizacion_pdf import (
    _STATIC,
    CLASE_RAIZ,
    PREVIEW_ANCHO_PX,
    TZ_NOTA,
    _build_html,
    _estilos_base,
    _estilos_cuerpo,
    _estilos_fuente,
    _estilos_hoja,
    _estilos_page,
    _fecha_dia,
    _mapa_svg,
    _mapa_svg_elemento,
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


# ===== Modelo del avión COTIZADO en la hoja 1 (feedback del cliente 4-sep-2026) =====
# El cliente quiere ver el TIPO de avión con el que se cotizó (a veces la
# ruta operativa va en otro), NUNCA la matrícula. Sin los campos (API viejo)
# la hoja 1 queda igual que siempre.


def _meta(html: str) -> str:
    """Bloque `.meta` (folio/cliente/fecha/tipo) de la hoja 1."""
    return html[html.index('class="meta"') : html.index('class="route"')]


def test_hoja1_pinta_el_modelo_cotizado_y_nunca_la_matricula() -> None:
    html = _build_html(_req(aeronave_cotizada_modelo="Piper Seneca V", matricula="XB-ABC"))
    meta = _meta(html)
    assert "<strong>Aeronave cotizada:</strong> Piper Seneca V" in meta
    # Junto a fecha/tipo, en la columna derecha (después de "Tipo:").
    assert meta.index("<strong>Tipo:</strong>") < meta.index("Aeronave cotizada:")
    # La matrícula (no VGV) no sale en NINGÚN lado de la hoja.
    assert "XB-ABC" not in html
    assert "Aeronaves cotizadas" not in html


def test_hoja1_varios_modelos_en_orden_de_tramo() -> None:
    html = _build_html(
        _req(
            aeronave_cotizada_modelo="Kodiak 100",
            modelos_cotizados=["Kodiak 100", "Cessna 206"],
            matricula="XB-ABC",
        )
    )
    assert "<strong>Aeronaves cotizadas:</strong> Kodiak 100 · Cessna 206" in _meta(html)
    assert "Aeronave cotizada:" not in html
    assert "XB-ABC" not in html


def test_hoja1_modelos_cotizados_gana_sobre_el_modelo_unico_y_se_depura() -> None:
    # Repetidos, espacios y (defensivo) una matrícula colada entre los modelos.
    html = _build_html(
        _req(
            aeronave_cotizada_modelo="Otro",
            modelos_cotizados=["Kodiak 100", " Kodiak 100 ", "", "XB-ABC"],
            matricula="XB-ABC",
        )
    )
    assert "<strong>Aeronave cotizada:</strong> Kodiak 100" in _meta(html)
    assert "Otro" not in _meta(html)
    assert "XB-ABC" not in html


def test_hoja1_sin_modelo_no_pinta_la_linea() -> None:
    # API viejo: sin ninguno de los campos nuevos. `avion_modelo` (ficha de la
    # hoja 2) NO alimenta la línea: el modelo cotizado es otro contrato.
    for req in (_req(), _req(avion_modelo="Piper Seneca V"), _req(modelos_cotizados=[])):
        html = _build_html(req)
        assert "Aeronave cotizada" not in html
        assert "Aeronaves cotizadas" not in html
    # Modelo en blanco tampoco.
    assert "Aeronave cotizada" not in _build_html(_req(aeronave_cotizada_modelo="  "))


def test_hoja1_vgv_conserva_su_regla_pero_la_linea_nueva_va_sin_matricula() -> None:
    html = _build_html(_req(aeronave_cotizada_modelo="Kodiak 100", matricula="XA-VGV"))
    meta = _meta(html)
    assert "<strong>Aeronave cotizada:</strong> Kodiak 100" in meta
    assert "VGV" not in meta
    # La sublínea de la ruta sigue mostrando el VGV (regla 26-ago intacta).
    assert "· XA-VGV" in html


def test_hoja1_avion_externo_no_duplica_el_modelo() -> None:
    # El externo ya se presenta bajo la ruta como "MODELO · MATRÍCULA" (§9.1):
    # la línea de la cabecera no se repite.
    html = _build_html(
        _req(
            aeronave_cotizada_modelo="HAWKER 400 A",
            avion_externo="HAWKER 400 A · XA-REG",
        )
    )
    assert "Aeronave cotizada" not in html
    assert html.count("HAWKER 400 A") == 1
    assert "HAWKER 400 A · XA-REG" in html


def test_modelo_cotizado_se_escapa() -> None:
    html = _build_html(_req(aeronave_cotizada_modelo="Cessna <206>"))
    assert "Cessna &lt;206&gt;" in _meta(html)
    assert "<206>" not in html


# ===== Vista previa de la hoja 1 en pantalla (rediseño del cotizador, 8-sep-2026) =====
# La preview del panel es el MISMO `_build_html` con el MISMO payload del PDF:
# solo hoja 1, sin fotos ni hoja "La aeronave", con CSS de pantalla. El HTML
# del PDF (default) queda byte-idéntico al de siempre: si un cambio en la
# hoja 1 no sale en ambos, la preview miente al operador.

FOTO_EXT = "data:image/jpeg;base64,RVhU"
FOTO_INT = "data:image/jpeg;base64,SU5U"


def _req_completo(**extra) -> CotizacionPdfRequest:
    """Payload con TODO lo que pinta el PDF: hoja 1 (VGV, modelo cotizado,
    traslados, itinerario con mapa, TUAS, extras, pernocta, descuento, IVA,
    MXN, notas) y hoja 2 (fotos + tarjeta "De un vistazo" + características)."""
    base = dict(
        ruta="CUN → AZP → BZE → CZM → CUN",
        fecha="2026-09-08T14:00:00Z",
        fecha_traslado_inicial="2026-09-12T13:00:00Z",
        fecha_traslado_final="2026-09-12T23:00:00Z",
        pasajeros=4,
        tiempo_cobrable_hr=2.4,
        tarifa_hora_usd=1650,
        subtotal_usd=4110,
        tuas_usd=100,
        tuas_detalle=["TUA CUN · $25.00 USD × 4 pax = $100.00"],
        extras=[{"concepto": "Catering", "monto_usd": 170}],
        viaticos_pernocta_usd=150,
        descuento_usd=20,
        iva_pct=16,
        iva_usd=734.4,
        total_usd=5324.4,
        total_mxn=96371.64,
        tc_usd_mxn=18.1,
        notas="Sujeto a slot en CUN",
        mostrar_tarifa_hora=True,
        matricula="XA-VGV",
        aeronave_cotizada_modelo="Piper Seneca V",
        avion_modelo="Piper Seneca V",
        foto_exterior=FOTO_EXT,
        foto_interior=FOTO_INT,
        avion_velocidad_kts=180,
        avion_pasajeros=6,
        avion_num_motores=2,
        avion_motor_hp=220,
        avion_caracteristicas=["Aire acondicionado"],
        avion_tiempo_tramo_hr=1.3,
        mapa_puntos=[
            _tramo(1, "CUN", _CUN, "AZP", (19.71, -90.50)),
            _tramo(2, "BZE", (17.53, -88.30), "CZM", _CZM),
            _tramo(3, "CZM", _CZM, "CUN", _CUN),
        ],
    )
    base.update(extra)
    return _req(**base)


def test_preview_es_solo_la_hoja_1_sin_fotos_ni_ficha() -> None:
    req = _req_completo()
    html = _build_html(req, solo_hoja_1=True)
    # La hoja 1 COMPLETA, con las mismas reglas del PDF (VGV visible, modelo
    # cotizado, tarifa/hr por toggle, TUAS por aeropuerto, MXN, notas).
    for esperado in (
        "#COT-1042",
        "Cliente Demo S.A.",
        "CUN → AZP → BZE → CZM → CUN",
        "· XA-VGV",
        "<strong>Aeronave cotizada:</strong> Piper Seneca V",
        "<h2>Traslados</h2>",
        "<h2>Itinerario</h2>",
        '<div class="mapa">',
        "<h2>Desglose</h2>",
        "Servicio aéreo (2.4 h × $1,650.00/hr)",
        "TUA CUN · $25.00 USD × 4 pax = $100.00",
        "Catering",
        "Viáticos por pernocta",
        "&minus;$20.00",
        "IVA (16%)",
        "$5,324.40",
        "Total MXN (T.C. 18.1)",
        "Sujeto a slot en CUN",
    ):
        assert esperado in html, esperado
    # SIN hoja 2 en el MARCADO: ni contenedor, ni ficha, ni tarjeta, ni fotos,
    # ni características. (El CSS del cuerpo es el compartido y conserva sus
    # selectores `.av-*`/`.foto-ancha` sin usar: fuente única con el PDF.)
    cuerpo = html[html.index(f'<body class="{CLASE_RAIZ}">') :]
    for prohibido in (
        'class="detalles"',
        "av-titulo",
        "De un vistazo",
        "foto-ancha",
        "km/h",
        "Aire acondicionado",
        FOTO_EXT,
        FOTO_INT,
        "data:image/jpeg",
    ):
        assert prohibido not in cuerpo, prohibido
    # Las únicas imágenes son los logos (membrete + marca de agua), en PNG.
    srcs = re.findall(r'<img[^>]+src="([^"]+)"', html)
    assert srcs and all(src.startswith("data:image/png;base64,") for src in srcs)


def test_preview_css_de_pantalla_sin_reglas_de_pagina() -> None:
    html = _build_html(_req_completo(), solo_hoja_1=True)
    assert "@page" not in html
    assert _estilos_page() not in html
    # Hoja blanca de ancho FIJO que el panel escala (contrato con el panel).
    assert PREVIEW_ANCHO_PX == 794
    assert f"width: {PREVIEW_ANCHO_PX}px" in html
    assert "background: #fff" in html
    # Marca de agua sutil IDÉNTICA (misma opacidad) pero anclada a la hoja:
    # la regla de pantalla va después de la del cuerpo y gana la cascada.
    assert 'class="marca"' in html and "opacity: 0.05" in html
    assert html.rindex(".marca { position: absolute; }") > html.index(".marca { position: fixed")
    # El pie que en el PDF vive en `@page :first` se ve como bloque, DESPUÉS
    # de las notas (cierra la hoja) y con la misma leyenda.
    pie = (
        f'<div class="pie-pantalla">{TZ_NOTA}<br>'
        "Gracias por volar con VuelaTour, Aero Charter Cancún.<br>www.vuelatour.com</div>"
    )
    assert pie in html
    assert html.index("Sujeto a slot en CUN") < html.index('class="pie-pantalla"')


def test_preview_hoja_1_es_byte_identica_a_la_del_pdf() -> None:
    # Mismo payload → el <body> de la hoja 1 (membrete, meta, ruta, traslados,
    # itinerario+mapa, desglose, notas) es el MISMO texto; solo cambia la cola
    # (hoja 2 vs pie de pantalla) y el CSS de página.
    req = _req_completo()
    pdf = _build_html(req)
    prev = _build_html(req, solo_hoja_1=True)
    body = f'<body class="{CLASE_RAIZ}">'
    hoja1_pdf = pdf[pdf.index(body) : pdf.index('<div class="detalles">')]
    hoja1_prev = prev[prev.index(body) : prev.index('<div class="pie-pantalla">')]
    assert hoja1_pdf == hoja1_prev
    assert _estilos_cuerpo() in pdf and _estilos_cuerpo() in prev


def test_pdf_html_no_cambia_con_la_vista_previa() -> None:
    # El default sigue siendo el documento de SIEMPRE (2 hojas, @page, fotos).
    req = _req_completo()
    pdf = _build_html(req)
    assert pdf == _build_html(req, solo_hoja_1=False)
    assert _estilos_base() in pdf
    assert _estilos_base() == _estilos_page() + _estilos_hoja()
    assert "@page :first" in pdf and "position: fixed" in pdf
    assert '<div class="detalles">' in pdf and FOTO_EXT in pdf and FOTO_INT in pdf
    assert "De un vistazo" in pdf and "Aire acondicionado" in pdf
    assert "pie-pantalla" not in pdf and f"{PREVIEW_ANCHO_PX}px" not in pdf


# Cinturón del refactor (8-sep-2026): sha256 del HTML del PDF para 3 payloads
# (los logos data-URI se normalizan a "data:LOGO" y las fuentes woff2 a
# "data:FUENTE" para no depender de los binarios). Refrescados el 8-sep-2026
# al mover el CSS del cuerpo a `app/static/cotizacion-hoja.css` (selectores
# acotados a `.cot-hoja`, <body class="cot-hoja">) e incrustar Arimo — cambio
# INTENCIONAL y verificado con `test_preview_hoja_1_es_byte_identica_a_la_del_pdf`.
# Refrescados otra vez el 8-sep-2026 (revisión de fidelidad de la hoja
# editable): `cotizacion-hoja.css` neutraliza dos reglas más del preflight de
# Tailwind — `h2 { font-weight: 700 }` (el preflight pone `inherit`; en
# WeasyPrint h2 ya es bold) y `td, th { padding: 1px }` (el preflight pone 0;
# es el valor del agente de usuario) — sin efecto en el PDF.
# Si un cambio INTENCIONAL de la hoja del cliente mueve estos hashes, se
# refrescan con el valor que imprime el assert — pero antes hay que
# preguntarse si la vista previa del panel y la hoja del panel (mismo CSS)
# siguen mostrando lo mismo (`_build_html` + el .css son la fuente única).
_SNAPSHOTS: dict[str, tuple[dict, str]] = {
    "completo": (
        dict(
            folio="COT-1042",
            fecha="2026-09-08T14:00:00Z",
            cliente="Punta Pájaros S.A.",
            origen="CUN",
            destino="CUN",
            tipo="MULTIESCALA",
            pasajeros=4,
            fecha_traslado_inicial="2026-09-12T13:00:00Z",
            fecha_traslado_final="2026-09-12T23:00:00Z",
            escalas=[
                {"orden": 1, "origen": "CUN", "destino": "HOL", "fecha": "2026-09-12"},
                {"orden": 2, "origen": "HOL", "destino": "CUN"},
            ],
            ruta="CUN → HOL → CUN",
            tiempo_cobrable_hr=2.4,
            tarifa_hora_usd=1650,
            subtotal_usd=4110,
            tuas_usd=100,
            tuas_detalle=["TUA CUN · $25.00 USD × 4 pax = $100.00"],
            extras=[
                {"concepto": "Catering", "monto_usd": 170, "moneda": "USD"},
                {"concepto": "Handler", "monto_usd": 80, "moneda": "MXN", "monto_nativo": 1450},
            ],
            extras_total_usd=250,
            viaticos_pernocta_usd=150,
            descuento_usd=20,
            iva_pct=16,
            iva_usd=734.4,
            total_usd=5324.4,
            total_mxn=96371.64,
            tc_usd_mxn=18.1,
            notas="Sujeto a slot en CUN <ojo>",
            mostrar_tarifa_hora=True,
            mostrar_itinerario=True,
            matricula="XA-VGV",
            foto_exterior=FOTO_EXT,
            foto_interior=FOTO_INT,
            avion_modelo="Piper Seneca V",
            avion_velocidad_kts=180,
            avion_pasajeros=6,
            avion_num_motores=2,
            avion_motor_hp=220,
            avion_caracteristicas=["Aire acondicionado", "Baño"],
            avion_tiempo_tramo_hr=1.3,
            mapa_puntos=[
                {
                    "orden": 1,
                    "origen_iata": "CUN",
                    "destino_iata": "HOL",
                    "o_lat": 21.0365,
                    "o_lon": -86.8771,
                    "d_lat": 21.1,
                    "d_lon": -86.9,
                },
                {
                    "orden": 2,
                    "origen_iata": "HOL",
                    "destino_iata": "CUN",
                    "o_lat": 21.1,
                    "o_lon": -86.9,
                    "d_lat": 21.0365,
                    "d_lon": -86.8771,
                    "es_ferry": True,
                },
            ],
            aeronave_cotizada_modelo="Piper Seneca V",
            modelos_cotizados=["Piper Seneca V", "Cessna 206"],
        ),
        "190d3002fd0833f0792d073aa9cf54c73ca5c083f03619fb0e4769a659a318eb",
    ),
    "minimo": (
        dict(folio="COT-1", cliente="Cliente", origen="CUN", destino="MID"),
        "780afab7bfdc9f9c76fb5c085c73381073eaa49ae1325ce3ffd95e2d81ce81d4",
    ),
    "externo_sin_itinerario": (
        dict(
            folio="COT-77",
            cliente="Broker X",
            origen="CUN",
            destino="BJX",
            pasajeros=1,
            escalas=[
                {"orden": 1, "origen": "CUN", "destino": "BJX"},
                {"orden": 2, "origen": "BJX", "destino": "CUN"},
            ],
            ruta="CUN → BJX → CUN",
            subtotal_usd=9000,
            tuas_usd=0,
            iva_pct=0.16,
            iva_usd=1440,
            total_usd=10440,
            mostrar_itinerario=False,
            avion_externo="HAWKER 400 A · XA-REG",
            aeronave_cotizada_modelo="HAWKER 400 A",
            mapa_puntos=[
                {
                    "orden": 1,
                    "origen_iata": "CUN",
                    "destino_iata": "BJX",
                    "o_lat": 21.0365,
                    "o_lon": -86.8771,
                    "d_lat": 20.9935,
                    "d_lon": -101.4808,
                },
                {
                    "orden": 2,
                    "origen_iata": "BJX",
                    "destino_iata": "CUN",
                    "o_lat": 20.9935,
                    "o_lon": -101.4808,
                    "d_lat": 21.0365,
                    "d_lon": -86.8771,
                },
            ],
            avion_modelo=None,
            foto_exterior=None,
        ),
        "c0052706556f72e3793ce973de2348468ffd9b4bc2eaa7ab14ab9c4d048ce526",
    ),
}


def _sha_normalizado(html: str) -> str:
    sin_logo = re.sub(r"data:image/png;base64,[A-Za-z0-9+/=]+", "data:LOGO", html)
    sin_fuente = re.sub(r"data:font/woff2;base64,[A-Za-z0-9+/=]+", "data:FUENTE", sin_logo)
    return hashlib.sha256(sin_fuente.encode("utf-8")).hexdigest()


def test_pdf_html_snapshot_previo_al_refactor_de_vista_previa() -> None:
    for nombre, (payload, esperado) in _SNAPSHOTS.items():
        actual = _sha_normalizado(_build_html(CotizacionPdfRequest(**payload)))
        assert actual == esperado, f"HTML del PDF cambió para '{nombre}': sha256={actual}"


# ===== Router: POST /reportes/cotizacion/preview-html =====

TOKEN = "secreto-de-prueba"
client = TestClient(app)


def _payload_preview() -> dict:
    return _req_completo().model_dump(mode="json")


def test_preview_router_sin_token_rechazado(monkeypatch) -> None:
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    res = client.post("/reportes/cotizacion/preview-html", json=_payload_preview())
    assert res.status_code == 401


def test_preview_router_devuelve_html_sin_weasyprint(monkeypatch) -> None:
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    # Cinturón: si algo intentara `from weasyprint import HTML` reventaría
    # con ImportError → la vista previa NO depende de WeasyPrint.
    monkeypatch.setitem(sys.modules, "weasyprint", None)
    res = client.post(
        "/reportes/cotizacion/preview-html",
        json=_payload_preview(),
        headers={"X-Internal-Token": TOKEN},
    )
    assert res.status_code == 200
    assert res.headers["content-type"] == "text/html; charset=utf-8"
    assert res.headers["cache-control"] == "no-store"
    assert "#COT-1042" in res.text and "Total MXN (T.C. 18.1)" in res.text
    assert 'class="detalles"' not in res.text and "@page" not in res.text
    assert FOTO_EXT not in res.text


def test_preview_router_error_es_500_con_detalle(monkeypatch) -> None:
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()

    def _roto(req: CotizacionPdfRequest) -> str:
        raise ValueError("boom en la plantilla")

    monkeypatch.setattr(reportes_router, "render_cotizacion_preview_html", _roto)
    res = client.post(
        "/reportes/cotizacion/preview-html",
        json=_payload_preview(),
        headers={"X-Internal-Token": TOKEN},
    )
    assert res.status_code == 500
    assert "boom en la plantilla" in res.json()["detail"]


# ===== CSS de la hoja en archivo estático + fuente incrustada (form-as-document, 8-sep-2026) =====
# La hoja que edita el operador en el panel y el PDF comparten ARCHIVO
# (`app/static/cotizacion-hoja.css`), FUENTE (`cotizacion-fuente.css`, Arimo
# woff2 base64) y MAPA (`_mapa_svg_elemento`). Python solo lee los archivos.

# Clases del marcado de la hoja 1 y de la hoja "La aeronave" que el CSS DEBE
# cubrir (contrato con el panel: si una desaparece del .css, la hoja del
# panel deja de parecerse al PDF).
_CLASES_MARCADO = (
    "marca", "header", "logo", "titulos", "meta", "route", "grid", "fecha",
    "itin-row", "itin-tabla", "itin-mapa", "mapa-solo", "mapa", "totales",
    "lbl", "val", "sub-row", "total-row", "total-mxn", "notas", "detalles",
    "av-titulo", "av-linea", "foto-ancha", "av-row", "av-foto", "vistazo",
    "vz-titulo", "vz-lbl", "vz-val", "caracts",
)


def _selectores(css: str) -> list[str]:
    sin_comentarios = re.sub(r"/\*.*?\*/", "", css, flags=re.S)
    return [
        s.strip()
        for bloque in re.findall(r"([^{}]+)\{", sin_comentarios)
        for s in bloque.split(",")
        if s.strip()
    ]


def test_css_del_cuerpo_y_la_fuente_son_los_archivos_estaticos() -> None:
    assert _estilos_cuerpo() == (_STATIC / "cotizacion-hoja.css").read_text(encoding="utf-8")
    assert _estilos_fuente() == (_STATIC / "cotizacion-fuente.css").read_text(encoding="utf-8")
    # Lo que recibe el panel = fuente + cuerpo, contiguo.
    assert _estilos_hoja() == _estilos_fuente() + _estilos_cuerpo()
    assert (_STATIC / "Arimo-OFL.txt").read_text(encoding="utf-8").startswith("Copyright")


def test_todo_selector_del_cuerpo_cuelga_de_la_raiz_cot_hoja() -> None:
    # Ningún selector suelto (body, table, h2, *): en el panel el CSS convive
    # con Tailwind y solo puede tocar lo que está dentro de `.cot-hoja`.
    assert CLASE_RAIZ == "cot-hoja"
    selectores = _selectores(_estilos_cuerpo())
    assert selectores
    sueltos = [s for s in selectores if not s.startswith(f".{CLASE_RAIZ}")]
    assert sueltos == [], sueltos
    # Todas las clases del marcado tienen regla.
    css = _estilos_cuerpo()
    faltan = [c for c in _CLASES_MARCADO if f".{c}" not in css]
    assert faltan == [], faltan
    # La tipografía de la hoja es la fuente incrustada, con fallback.
    assert (
        f".{CLASE_RAIZ}, .{CLASE_RAIZ} * {{ font-family: 'Arimo', 'Liberation Sans', "
        "'Helvetica Neue', Arial, sans-serif; color: #1d1d1d; }"
    ) in css
    # El pie de página del PDF (@page) usa la misma fuente.
    assert _estilos_page().count("font-family: 'Arimo'") == 3


def test_fuente_incrustada_regular_y_bold_con_los_glifos_de_la_hoja() -> None:
    from io import BytesIO

    from fontTools.ttLib import TTFont  # dependencia de WeasyPrint

    f = _estilos_fuente()
    assert f.count("@font-face") == 2
    assert f.count("font-family: 'Arimo';") == 2
    assert "font-weight: 400;" in f and "font-weight: 700;" in f
    blobs = re.findall(r"data:font/woff2;base64,([A-Za-z0-9+/=]+)\) format\('woff2'\)", f)
    assert len(blobs) == 2
    pesos = []
    for b in blobs:
        fuente = TTFont(BytesIO(base64.b64decode(b)))
        cmap = fuente.getBestCmap()
        # Lo que imprime la hoja: flecha de la ruta, menos del descuento,
        # × de las TUAS, · separador, acentos/ñ del español.
        for ch in "→−×·áéíóúñÁÉÍÓÚÑ¿¡$…–—“”":
            assert ord(ch) in cmap, ch
        pesos.append(fuente["OS/2"].usWeightClass)
    assert sorted(pesos) == [400, 700]


def test_pdf_preview_y_grupo_llevan_la_hoja_css_completa_y_la_raiz() -> None:
    from app.schemas.reportes import CotizacionGrupoPdfRequest
    from app.services.cotizacion_grupo_pdf import _build_html as build_grupo
    from tests.test_cotizacion_grupo_pdf import _payload as payload_grupo

    req = _req_completo()
    documentos = (
        _build_html(req),
        _build_html(req, solo_hoja_1=True),
        build_grupo(CotizacionGrupoPdfRequest(**payload_grupo())),
    )
    for html in documentos:
        assert _estilos_hoja() in html
        assert html.count("@font-face") == 2
        assert html.count("<body") == 1 and f'<body class="{CLASE_RAIZ}">' in html


def test_mapa_svg_elemento_es_exactamente_lo_que_embebe_la_hoja() -> None:
    local = [_tramo(1, "CUN", _CUN, "MID", _MID), _tramo(2, "MID", _MID, "CUN", _CUN)]
    amplio = [_tramo(1, "CUN", _CUN, "BJX", _BJX), _tramo(2, "BJX", _BJX, "CUN", _CUN)]
    for puntos in (local, amplio):
        svg = _mapa_svg_elemento(puntos)
        assert svg.startswith("<svg viewBox=") and svg.endswith("</svg>")
        assert _mapa_svg(puntos) == f'<div class="mapa">{svg}</div>'
        assert f'<div class="mapa">{svg}</div>' in _build_html(_req_completo(mapa_puntos=puntos))
    assert _mapa_svg_elemento([]) == "" and _mapa_svg([]) == ""


# ===== Router: GET /reportes/cotizacion/hoja.css y POST /reportes/cotizacion/mapa-svg =====


def test_hoja_css_router_sin_token_rechazado(monkeypatch) -> None:
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    assert client.get("/reportes/cotizacion/hoja.css").status_code == 401
    res = client.get("/reportes/cotizacion/hoja.css", headers={"X-Internal-Token": "malo"})
    assert res.status_code == 401


def test_hoja_css_router_devuelve_fuente_y_cuerpo_cacheables(monkeypatch) -> None:
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    monkeypatch.setitem(sys.modules, "weasyprint", None)
    res = client.get("/reportes/cotizacion/hoja.css", headers={"X-Internal-Token": TOKEN})
    assert res.status_code == 200
    assert res.headers["content-type"] == "text/css; charset=utf-8"
    assert res.headers["cache-control"] == "public, max-age=3600"
    # EXACTAMENTE lo que llevan el PDF y la vista previa.
    assert res.text == _estilos_hoja()
    assert res.text in _build_html(_req_completo())
    assert res.text in _build_html(_req_completo(), solo_hoja_1=True)


def test_mapa_svg_router_sin_token_rechazado(monkeypatch) -> None:
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    body = {"mapa_puntos": [_tramo(1, "CUN", _CUN, "MID", _MID).model_dump()]}
    assert client.post("/reportes/cotizacion/mapa-svg", json=body).status_code == 401


def test_mapa_svg_router_devuelve_el_mismo_mapa_del_pdf(monkeypatch) -> None:
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    monkeypatch.setitem(sys.modules, "weasyprint", None)
    casos = {
        "local": [_tramo(1, "CUN", _CUN, "MID", _MID), _tramo(2, "MID", _MID, "CUN", _CUN)],
        "amplio": [
            _tramo(1, "CUN", _CUN, "BJX", _BJX),
            _tramo(2, "BJX", _BJX, "AZP", _AZP),
            _tramo(3, "AZP", _AZP, "CUN", _CUN),
        ],
    }
    for nombre, puntos in casos.items():
        res = client.post(
            "/reportes/cotizacion/mapa-svg",
            json={"mapa_puntos": [p.model_dump() for p in puntos]},
            headers={"X-Internal-Token": TOKEN},
        )
        assert res.status_code == 200, nombre
        assert res.headers["content-type"] == "image/svg+xml"
        assert res.headers["cache-control"] == "no-store"
        assert res.text == _mapa_svg_elemento(puntos)
        hoja = _build_html(_req_completo(mapa_puntos=puntos))
        assert f'<div class="mapa">{res.text}</div>' in hoja
    # Mismo modo local/amplio que el PDF: el amplio abre la vista más allá del
    # lienzo peninsular y pinta el contorno de México de fondo.
    _, _, bw_local, _ = _viewbox(_mapa_svg_elemento(casos["local"]))
    _, _, bw_amplio, _ = _viewbox(_mapa_svg_elemento(casos["amplio"]))
    assert bw_local <= 600 < bw_amplio


def test_mapa_svg_router_sin_puntos_es_204(monkeypatch) -> None:
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    for body in ({"mapa_puntos": []}, {}):
        res = client.post(
            "/reportes/cotizacion/mapa-svg", json=body, headers={"X-Internal-Token": TOKEN}
        )
        assert res.status_code == 204
        assert res.content == b""
