"""PDF de la cotización de GRUPO (4-sep-2026), probado sobre el HTML (sin
WeasyPrint, import perezoso): hoja 1 (folio, cliente, ruta, grupo,
itinerario, desglose consolidado, totales, precio por persona), hoja
"Flota asignada" (toggles) y fichas por modelo.

Reglas cubiertas: el consolidado se pinta TAL CUAL viene (nunca se
recalcula), COMISION_VENDEDOR/redondeo jamás aparecen, matrícula oculta
salvo VGV (misma regla que el PDF de un avión), columna Fecha solo si algún
tramo la trae, y un payload mínimo (skew) también renderiza.
"""

from fastapi.testclient import TestClient

from app.config import get_settings
from app.main import app
from app.routers import reportes as reportes_router
from app.schemas.reportes import CotizacionGrupoPdfRequest
from app.services import cotizacion_grupo_pdf, cotizacion_pdf
from app.services.cotizacion_grupo_pdf import _build_html

FOTO_EXT = "data:image/jpeg;base64,RVhU"
FOTO_INT = "data:image/jpeg;base64,SU5U"


def _avion(posicion: int, modelo: str, matricula: str, **extra) -> dict:
    base = {
        "posicion": posicion,
        "modelo": modelo,
        "matricula": matricula,
        "asientos": 5,
        "pasajeros": 5,
        "rotaciones": 1,
        "tiempo_hr": 1.5,
        "salida_estimada": "2026-10-12T13:00:00Z",
        "subtotal_usd": 1500.0,
        "tarifa_hora_usd": 650.0,
        "velocidad_kts": 120,
        "num_motores": 1,
        "motor_hp": 300,
        "caracteristicas": ["Aire acondicionado"],
        "foto_exterior": None,
        "foto_interior": None,
        "foto_exterior_url": None,
        "foto_interior_url": None,
    }
    base.update(extra)
    return base


def _payload(**extra) -> dict:
    """Payload COMPLETO con la forma exacta de groups-pdf.service.ts."""
    base: dict = {
        "folio_grupo": "G-12",
        "folio": 12,
        "nombre": "Tour Chichén Itzá",
        "cliente": "Agencia Demo S.A.",
        "fecha": "2026-10-12T13:00:00Z",
        "pasajeros_total": 44,
        "aviones_total": 3,
        "ruta": "CUN → CZA → CUN",
        "itinerario": [
            {
                "orden": 1,
                "origen": "CUN",
                "destino": "CZA",
                "es_ferry": False,
                "requiere_pernocta": False,
                "tipo_parada": "NORMAL",
                "servicio_notas": None,
                "fecha": "2026-10-12",
            },
            {
                "orden": 2,
                "origen": "CZA",
                "destino": "CUN",
                "es_ferry": False,
                "requiere_pernocta": False,
                "tipo_parada": "NORMAL",
                "servicio_notas": None,
                "fecha": None,
            },
        ],
        "mapa_puntos": [
            {
                "orden": 1,
                "origen_iata": "CUN",
                "destino_iata": "CZA",
                "o_lat": 21.0365,
                "o_lon": -86.8771,
                "d_lat": 20.6413,
                "d_lon": -88.4461,
                "es_ferry": False,
            },
            {
                "orden": 2,
                "origen_iata": "CZA",
                "destino_iata": "CUN",
                "o_lat": 20.6413,
                "o_lon": -88.4461,
                "d_lat": 21.0365,
                "d_lon": -86.8771,
                "es_ferry": False,
            },
        ],
        "desglose_consolidado": [
            {
                "clave": "TIEMPO_VUELO",
                "concepto": "Servicio aéreo · 3 aeronaves",
                "monto_usd": 13965.0,
            },
            {"clave": "TUAS", "concepto": "TUA CZA · 44 pax", "monto_usd": 792.0},
            {"clave": "TUAS", "concepto": "TUA CUN · 5 pax", "monto_usd": 125.0},
            {
                "clave": "EXTRA",
                "concepto": "Tour Chichén Itzá · 44 × $85.00",
                "monto_usd": 3740.0,
                "cantidad": 44,
                "unitario": 85,
                "moneda": "USD",
            },
            {
                "clave": "EXTRA",
                "concepto": "Camionetas",
                "monto_usd": 500.0,
                "cantidad": 2,
                "unitario": 250,
                "moneda": "USD",
            },
            {"clave": "AJUSTE", "concepto": "Descuento", "monto_usd": -100.0},
            {"clave": "IVA", "concepto": "IVA 16%", "monto_usd": 3028.32},
            {"clave": "PERNOCTA", "concepto": "Viáticos por pernocta", "monto_usd": 150.0},
        ],
        "servicio_aereo_usd": 13965.0,
        "horas_total_hr": 6.5,
        "tuas_usd": 917.0,
        "tuas_detalle": ["TUA CZA · 44 pax", "TUA CUN · 5 pax"],
        "extras": [
            {
                "concepto": "Tour Chichén Itzá · 44 × $85.00",
                "monto_usd": 3740.0,
                "cantidad": 44,
                "unitario": 85,
                "moneda": "USD",
                "aplica_iva": True,
            }
        ],
        "extras_total_usd": 4240.0,
        "viaticos_pernocta_usd": 150.0,
        "descuento_usd": 100.0,
        "subtotal_usd": 19072.0,
        "iva_pct": 16,
        "iva_usd": 3028.32,
        "total_usd": 22100.32,
        "total_mxn": 397805.76,
        "tc_usd_mxn": 18.0,
        "precio_por_persona_usd": 502.28,
        "moneda": "USD",
        "mostrar_precio_por_persona": True,
        "mostrar_tarifa": False,
        "mostrar_anexo_aviones": True,
        "mostrar_subtotal_por_avion": False,
        "mostrar_itinerario": True,
        "aviones": [
            _avion(
                1,
                "Kodiak 100",
                "XA-VGV",
                asientos=9,
                pasajeros=9,
                subtotal_usd=6000.0,
                tarifa_hora_usd=1750.0,
                velocidad_kts=150,
                motor_hp=750,
                foto_exterior=FOTO_EXT,
                foto_interior=FOTO_INT,
            ),
            _avion(2, "Cessna 206", "XB-ANU", rotaciones=2, pasajeros=10),
            _avion(3, "Cessna 206", "XB-RTO", pasajeros=5, foto_exterior=FOTO_EXT),
        ],
        "notas": "Incluye guía certificado <en sitio>.",
        "condiciones": None,
    }
    base.update(extra)
    return base


def _html(**extra) -> str:
    return _build_html(CotizacionGrupoPdfRequest(**_payload(**extra)))


# ===== Hoja 1 =====


def test_hoja1_folio_cliente_fecha_y_grupo() -> None:
    html = _html()
    assert "Cotización de grupo" in html
    assert "<strong>Folio:</strong> G-12" in html
    assert "Agencia Demo S.A." in html
    # Salida en hora Cancún (UTC−5): 13:00Z → 08:00.
    assert "12/10/2026 08:00" in html
    assert "CUN → CZA → CUN" in html
    assert "Grupo de 44 pasajeros · 3 aeronaves" in html
    assert "Tour Chichén Itzá" in html
    # Notas escapadas.
    assert "Incluye guía certificado &lt;en sitio&gt;." in html


def test_reusa_estilos_mapa_y_branding_del_pdf_de_un_avion() -> None:
    # Importa, no copia: mismo mapa (modo local/amplio) y mismos estilos.
    assert cotizacion_grupo_pdf._mapa_svg is cotizacion_pdf._mapa_svg
    assert cotizacion_grupo_pdf._itinerario_html is cotizacion_pdf._itinerario_html
    assert cotizacion_grupo_pdf._ficha_aeronave_html is cotizacion_pdf._ficha_aeronave_html
    html = _html()
    assert cotizacion_pdf._estilos_base() in html
    assert "#dc2626" in html and "#102a43" in html
    assert 'viewBox="' in html  # mapa SVG presente
    assert ">CZA</text>" in html


def test_itinerario_con_fecha_agrega_columna() -> None:
    html = _html()
    assert "<thead><tr><th>#</th><th>Tramo</th><th>Fecha</th></tr></thead>" in html
    assert '<td>1</td><td>CUN → CZA</td><td class="fecha">12 oct 2026</td></tr>' in html
    assert '<td>2</td><td>CZA → CUN</td><td class="fecha">—</td></tr>' in html


def test_itinerario_sin_fecha_no_agrega_columna() -> None:
    itin = _payload()["itinerario"]
    for t in itin:
        t["fecha"] = None
    html = _html(itinerario=itin)
    assert "<thead><tr><th>#</th><th>Tramo</th></tr></thead>" in html
    assert "<th>Fecha</th>" not in html
    assert 'class="fecha"' not in html


def test_itinerario_oculto_deja_solo_el_mapa() -> None:
    html = _html(mostrar_itinerario=False)
    assert "<h2>Itinerario</h2>" not in html
    assert "<h2>La ruta</h2>" in html
    assert 'class="mapa-solo"' in html


def test_desglose_pinta_lineas_tal_cual_y_totales_del_api() -> None:
    html = _html()
    assert "Servicio aéreo · 3 aeronaves" in html and "$13,965.00" in html
    assert "TUA CZA · 44 pax" in html and "$792.00" in html
    assert "TUA CUN · 5 pax" in html and "$125.00" in html
    # El concepto ya trae "44 × $85.00": no se duplica.
    assert "Tour Chichén Itzá · 44 × $85.00" in html
    assert "44 × $85.00 · 44 × $85.00" not in html
    # Línea con cantidad/unitario pero sin "×" en el concepto: se agrega.
    assert "Camionetas · 2 × $250.00" in html
    # Descuento en negativo tipográfico.
    assert "Descuento" in html and "&minus;$100.00" in html
    assert "Viáticos por pernocta" in html and "$150.00" in html
    # Totales del API, sin recalcular.
    assert "Subtotal (sin IVA)" in html and "$19,072.00" in html
    assert "IVA (16%)" in html and "$3,028.32" in html
    assert "Total (USD)" in html and "$22,100.32" in html
    assert "Total MXN (T.C. 18)" in html and "$397,805.76 MXN" in html
    # El IVA NO sale como línea del cuerpo (solo en los totales).
    assert html.count("$3,028.32") == 1
    # Sin horas ni tarifa por hora en la hoja 1.
    assert "6.5 h" not in html and "/hr" not in html


def test_nunca_pinta_comision_del_vendedor_ni_redondeo() -> None:
    lineas = _payload()["desglose_consolidado"] + [
        {"clave": "COMISION_VENDEDOR", "concepto": "Comisión del vendedor", "monto_usd": 300.0},
        {"clave": "REDONDEO", "concepto": "Redondeo", "monto_usd": 12.0},
        {"clave": "AJUSTE", "concepto": "Redondeo", "monto_usd": 8.0},
    ]
    html = _html(desglose_consolidado=lineas)
    assert "Comisi" not in html
    assert "Redondeo" not in html
    assert "$300.00" not in html and "$12.00" not in html and "$8.00" not in html


def test_sin_total_mxn_no_pinta_linea_mxn() -> None:
    html = _html(total_mxn=None, tc_usd_mxn=None)
    assert "Total MXN" not in html


def test_precio_por_persona_toggle() -> None:
    html = _html()
    assert "Precio por persona (44 pasajeros)" in html and "$502.28 USD" in html
    assert "Precio por persona" not in _html(mostrar_precio_por_persona=False)
    # Toggle prendido pero el API no lo mandó: nunca se divide aquí.
    assert "Precio por persona" not in _html(precio_por_persona_usd=None)
    assert "$502.28" not in _html(mostrar_precio_por_persona=False)


# ===== Hoja "Flota asignada" =====


def test_anexo_flota_toggle_y_columnas_base() -> None:
    html = _html()
    assert "Flota asignada" in html
    assert "3 aeronaves para 44 pasajeros." in html
    assert "<th>#</th><th>Aeronave</th>" in html
    assert '<th class="num">Asientos</th><th class="num">Pasajeros</th><th>Salidas</th>' in html
    assert "<td>1 vuelta</td>" in html
    assert "<td>2 vueltas</td>" in html
    # Pie con el total de pasajeros del API.
    assert '<td colspan="3">Total del grupo</td><td class="num">44</td>' in html
    # Sin subtotal ni tarifa por default.
    assert "Subtotal (USD)" not in html and "$6,000.00" not in html
    assert "<th class=\"num\">Tarifa</th>" not in html and "$1,750.00" not in html
    # Apagado: no hay hoja.
    assert "Flota asignada" not in _html(mostrar_anexo_aviones=False)


def test_anexo_subtotal_por_avion_toggle() -> None:
    html = _html(mostrar_subtotal_por_avion=True)
    assert '<th class="num">Subtotal (USD)</th>' in html
    assert "$6,000.00" in html and "$1,500.00" in html
    # El pie pinta el total del API (no una suma propia).
    assert '<td class="num">$22,100.32</td>' in html
    assert "Subtotales por aeronave con IVA incluido." in html


def test_anexo_tarifa_toggle() -> None:
    html = _html(mostrar_tarifa=True)
    assert '<th class="num">Horas</th><th class="num">Tarifa</th>' in html
    assert "$1,750.00/hr" in html and "$650.00/hr" in html
    assert "1.5 h" in html
    assert "6.5 h" in html  # total de horas del API en el pie


def test_matricula_oculta_salvo_vgv() -> None:
    html = _html()
    # Hoja de flota: el VGV se comercializa por matrícula; los demás no.
    assert "<td>Kodiak 100 · XA-VGV</td>" in html
    assert "XB-ANU" not in html
    assert "XB-RTO" not in html
    assert "<td>Cessna 206</td>" in html
    # Regla compartida con el PDF de un avión (fuente única).
    assert cotizacion_grupo_pdf._mostrar_matricula is cotizacion_pdf._mostrar_matricula


# ===== Hojas "Las aeronaves" =====


def test_fichas_una_por_modelo_con_foto_del_primero_que_la_trae() -> None:
    html = _html()
    # Dos modelos distintos → dos fichas (aunque haya tres aviones).
    assert html.count('class="detalles"') == 2
    assert '<div class="av-titulo">Kodiak 100</div>' in html
    assert html.count('<div class="av-titulo">Cessna 206</div>') == 1
    # Kodiak: exterior ancho + interior en la fila con la tarjeta.
    assert f'<img class="foto-ancha" src="{FOTO_EXT}"' in html
    assert f'<td class="av-foto"><img src="{FOTO_INT}"' in html
    # Cessna: la foto sale del avión 3 (el 2 no trae).
    assert html.count(f'src="{FOTO_EXT}"') == 2
    # Ficha "De un vistazo" con velocidad/motores + conteo del modelo.
    assert "De un vistazo" in html
    assert "150 kt / 278 km/h" in html
    assert "1 × 750 HP" in html
    assert "2 aeronaves · 15 pasajeros" in html
    assert "1 aeronave · 9 pasajeros" in html
    assert "Aire acondicionado" in html
    # La hoja de la aeronave NUNCA lleva matrícula (ni el VGV).
    hoja_av = html[html.index('class="detalles"') :]
    assert "VGV" not in hoja_av


def test_ficha_usa_url_publica_como_respaldo_y_omite_modelo_vacio() -> None:
    aviones = [
        _avion(
            1,
            "Kodiak 100",
            "XA-VGV",
            foto_exterior=None,
            foto_exterior_url='https://cdn.example.com/kodiak.jpg?x="1"',
        ),
        # Modelo sin foto y sin datos de ficha: no genera hoja.
        _avion(
            2,
            "Misterioso",
            "XB-XXX",
            asientos=None,
            velocidad_kts=None,
            num_motores=None,
            motor_hp=None,
            caracteristicas=[],
        ),
    ]
    html = _html(aviones=aviones)
    assert 'src="https://cdn.example.com/kodiak.jpg?x=&quot;1&quot;"' in html
    assert html.count('class="detalles"') == 1
    assert "Misterioso" in html  # sí aparece en la hoja de flota
    assert '<div class="av-titulo">Misterioso</div>' not in html


# ===== Skew / payload mínimo =====


def test_payload_minimo_renderiza_y_campos_extra_se_ignoran() -> None:
    req = CotizacionGrupoPdfRequest(campo_nuevo_del_futuro=1)
    html = _build_html(req)
    assert "Cotización de grupo" in html
    assert "<strong>Folio:</strong> G-s/n" in html
    assert "Por confirmar" in html
    # Sin desglose consolidado (API viejo): cuerpo con los escalares.
    assert "Servicio aéreo" in html and "TUAS" in html
    assert "Total (USD)" in html
    assert "Flota asignada" not in html
    assert 'class="detalles"' not in html
    assert "Precio por persona" not in html


def test_payload_sin_desglose_usa_escalares_y_subtotal_derivado() -> None:
    html = _html(
        desglose_consolidado=[],
        subtotal_usd=None,
        folio_grupo="",
        folio=7,
        ruta=None,
    )
    assert "<strong>Folio:</strong> G-7" in html
    # Título armado de la plantilla (skew sin `ruta`).
    assert "CUN → CZA → CUN" in html
    assert "Servicio aéreo · 3 aeronaves" in html
    assert "TUA CZA · 44 pax" in html and "TUAS (total)" in html
    assert "Tour Chichén Itzá · 44 × $85.00" in html
    assert "Viáticos por pernocta" in html and "&minus;$100.00" in html
    # subtotal = total − IVA solo cuando el API no lo manda.
    assert "$19,072.00" in html


# ===== Router =====

TOKEN = "secreto-de-prueba"
client = TestClient(app)


def test_router_sin_token_rechazado(monkeypatch) -> None:
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    res = client.post("/reportes/cotizacion-grupo", json=_payload())
    assert res.status_code == 401


def test_router_devuelve_pdf_con_nombre_de_archivo(monkeypatch) -> None:
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    capturado: dict = {}

    def _render_falso(req: CotizacionGrupoPdfRequest) -> bytes:
        capturado["req"] = req
        return b"%PDF-1.4 fake"

    # WeasyPrint puede no tener libs de sistema en local: se sustituye el
    # render y se prueba el contrato HTTP (token, tipo, nombre de archivo).
    monkeypatch.setattr(reportes_router, "render_cotizacion_grupo_pdf", _render_falso)
    res = client.post(
        "/reportes/cotizacion-grupo",
        json=_payload(),
        headers={"X-Internal-Token": TOKEN},
    )
    assert res.status_code == 200
    assert res.headers["content-type"] == "application/pdf"
    assert 'filename="cotizacion-grupo-G-12.pdf"' in res.headers["content-disposition"]
    assert res.content == b"%PDF-1.4 fake"
    assert capturado["req"].folio_grupo == "G-12"
    assert len(capturado["req"].aviones) == 3


def test_router_error_de_render_es_500_con_detalle(monkeypatch) -> None:
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()

    def _render_roto(req: CotizacionGrupoPdfRequest) -> bytes:
        raise OSError("cannot load library 'libgobject-2.0-0'")

    monkeypatch.setattr(reportes_router, "render_cotizacion_grupo_pdf", _render_roto)
    res = client.post(
        "/reportes/cotizacion-grupo",
        json=_payload(),
        headers={"X-Internal-Token": TOKEN},
    )
    assert res.status_code == 500
    assert "libgobject" in res.json()["detail"]


# ===== TUAS con la operación visible (feedback del cliente 4-sep-2026) =====
# Con cantidad (pax gravados) + unitario nativo la línea se lee
# «TUA CZA · 44 pax × $20.85»; sin ellos, como siempre. El monto es del API.


def test_tuas_con_unitario_pinta_pax_por_tarifa_sin_tocar_el_monto() -> None:
    lineas = [
        ln
        for ln in _payload()["desglose_consolidado"]
        if ln["clave"] != "TUAS"
    ]
    lineas[1:1] = [
        # Contrato 5-sep: concepto "TUA CZA" + cantidad/unitario/moneda.
        {
            "clave": "TUAS",
            "concepto": "TUA CZA",
            "monto_usd": 917.4,
            "cantidad": 44,
            "unitario": 20.85,
            "moneda": "USD",
        },
        # Unitario nativo en pesos: el monto sigue en USD, la operación en MXN.
        {
            "clave": "TUAS",
            "concepto": "TUA PCE",
            "monto_usd": 73.47,
            "cantidad": 4,
            "unitario": 330.6,
            "moneda": "MXN",
        },
        # Legado "TUA CUN · 5 pax" + números: no se duplica el pax.
        {
            "clave": "TUAS",
            "concepto": "TUA CUN · 5 pax",
            "monto_usd": 125.0,
            "cantidad": 5,
            "unitario": 25,
        },
    ]
    html = _html(desglose_consolidado=lineas)
    assert '<td class="lbl">TUA CZA · 44 pax × $20.85</td><td class="val">$917.40</td>' in html
    assert '<td class="lbl">TUA PCE · 4 pax × $330.60 MXN</td><td class="val">$73.47</td>' in html
    assert '<td class="lbl">TUA CUN · 5 pax × $25.00</td><td class="val">$125.00</td>' in html
    assert "5 pax · 5 pax" not in html
    assert "TUA CZA · 44 × $20.85" not in html  # los pasajeros se leen como pax
    # Los extras conservan su forma "n × $u" (sin "pax").
    assert "Camionetas · 2 × $250.00" in html
    assert "Tour Chichén Itzá · 44 × $85.00" in html


def test_tuas_sin_unitario_queda_como_hoy() -> None:
    html = _html()
    assert '<td class="lbl">TUA CZA · 44 pax</td><td class="val">$792.00</td>' in html
    assert '<td class="lbl">TUA CUN · 5 pax</td><td class="val">$125.00</td>' in html
    assert "pax ×" not in html
    # Solo cantidad (sin unitario) o solo unitario: tampoco se inventa la operación.
    lineas = _payload()["desglose_consolidado"]
    lineas[1] = {**lineas[1], "cantidad": 44}
    lineas[2] = {**lineas[2], "unitario": 25}
    html = _html(desglose_consolidado=lineas)
    assert "TUA CZA · 44 pax</td>" in html and "TUA CUN · 5 pax</td>" in html
    assert "pax ×" not in html
