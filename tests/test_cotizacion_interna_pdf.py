"""PDF de la COTIZACIÓN INTERNA v2 (8-sep-2026), probado sobre el HTML (sin
WeasyPrint, import perezoso): banda «interna», FECHA DEL VUELO protagonista,
matrícula SIEMPRE visible, tabla de tramos como la hoja de administración
(RUTA «Cancun-Merida» · FECHA «26-jun» · MILLAS · TIEMPO «01:18» con calzos ·
COSTO POR HORA · TOTAL POR TRAMO + TOTAL), fila de ajuste SOLO si Σ tramos
no cuadra con el servicio aéreo, TUAS solo cobradas, desglose con comisión
del vendedor, cobros compactos con comisión bancaria (caso real de la foto:
Paywise 8.857 % = $3,236.36 → neto $33,303.64), notas internas truncadas y
pie «Documento interno · generado … · usuario».

Reglas: nada se recalcula (todo viene del API); jamás fotos ni ficha del
avión; nada de operación (tacos, horas voladas, avión operativo, traslados)
ni partición/gastos/utilidad/CFDI aunque un API viejo los mande; un payload
mínimo (skew) y campos extra también renderizan.
"""

from fastapi.testclient import TestClient

from app.config import get_settings
from app.main import app
from app.routers import reportes as reportes_router
from app.schemas.reportes import CotizacionInternaPdfRequest
from app.services import cotizacion_grupo_pdf, cotizacion_interna_pdf, cotizacion_pdf
from app.services.cotizacion_interna_pdf import (
    BANDA_INTERNA,
    NOTA_TRAMOS,
    _build_html,
    _dia_mes,
    _hhmm,
    _millas,
    _truncar,
)


def _tramo(orden: int, origen: str, destino: str, **extra) -> dict:
    nombres = {"CUN": "Cancun", "MID": "Merida", "HOL": "Isla Holbox", "CZM": "Cozumel"}
    base = {
        "orden": orden,
        "ruta": f"{nombres[origen]}-{nombres[destino]}",
        "origen_iata": origen,
        "destino_iata": destino,
        "origen_nombre": nombres[origen],
        "destino_nombre": nombres[destino],
        "fecha": "2026-06-26",
        "millas": 157,
        "tiempo_hr": 1.3,  # 1.15 h de vuelo + 0.15 h de calzo
        "tiempo_hhmm": "01:18",
        "tarifa_hora_usd": 900.0,
        "total_usd": 1170.0,
        "pax": 4,
        "es_ferry": False,
        "pernocta": False,
        "pernocta_usd": 0,
        "tuas_usd": 0,
        "consolidado": False,
    }
    base.update(extra)
    return base


def _payload(**extra) -> dict:
    """Payload COMPLETO con la forma del contrato v2 de pyservices.service.ts
    (números coherentes con la hoja de administración: 2 tramos de 01:18 a
    $900/hr = $2,340.00; TUA CUN 4 × $20.85; extra; comisión; descuento;
    IVA 16 % de $2,720.00 = $435.20 → total $3,155.20; cobro Paywise MXN)."""
    base: dict = {
        "folio": "1042",
        "version": 3,
        "estado": "CONFIRMADO",
        "estado_label": "Confirmado",
        "tipo": "REDONDO",
        "cliente": "Cliente Demo S.A.",
        "razon_social": "Cliente Demo, S.A. de C.V.",
        "cliente_rfc": "CDE010101AAA",
        "es_broker": False,
        "fecha_vuelo": "2026-06-26",
        "fecha_vuelo_fin": None,
        "fecha": "2026-06-01T15:12:00Z",
        "fecha_confirmacion": None,
        "tarifa_tipo": "PUBLICO",
        "tarifa_tipo_label": "Público",
        "tarifa_hora_usd": 900.0,
        "tarifa_override": False,
        "tarifa_preferencial": False,
        "metodo_cobro": "PAYWISE",
        "metodo_cobro_detalle": None,
        "metodo_cobro_label": "Paywise",
        "tc_usd_mxn": 18.27,
        "vendedor": "Saab",
        "cotizado_por": "Itzi",
        "aeronave_cotizada_modelo": "Piper Seneca V",
        "aeronave_cotizada_matricula": "N4142R",
        "avion_externo": None,
        "operador_externo": None,
        "piloto": "Luis Piloto",
        "copiloto": "Ana Copiloto",
        "pasajeros": 4,
        "ruta": "CUN → MID → CUN",
        "itinerario_operativo": False,
        "cotizacion_abierta": False,
        "es_interno": False,
        "es_externo": False,
        "grupo_folio": None,
        "grupo_posicion": None,
        "grupo_total_aviones": None,
        "combinado_con_folio": None,
        "tramos_cotizados": [_tramo(1, "CUN", "MID"), _tramo(2, "MID", "CUN")],
        "tramos_tiempo_total_hr": 2.6,
        "tramos_tiempo_total_hhmm": "02:36",
        "tramos_total_usd": 2340.0,
        "tramos_ajuste_usd": 0,
        "tramos_ajuste_motivo": None,
        "horas_cotizadas_hr": 2.6,
        "vuelo_hr": 2.3,
        "calzos_hr": 0.3,
        "sobrevuelo_hr": 0,
        "tiempo_cobrable_hr": 2.6,
        "hora_minima_aplicada": False,
        "cobrable_override": False,
        "lineas": [
            {
                "clave": "TIEMPO_VUELO",
                "concepto": "Tiempo de vuelo · 2.6 hr × $900/hr",
                "monto_usd": 2340.0,
                "cantidad": 2.6,
                "unitario": 900.0,
                "moneda": "USD",
            },
            {
                "clave": "TUAS",
                "concepto": "TUA CUN",
                "monto_usd": 83.4,
                "cantidad": 4,
                "unitario": 20.85,
                "moneda": "USD",
            },
            {
                "clave": "EXTRA",
                "concepto": "Hieleras",
                "monto_usd": 50.0,
                "cantidad": 2,
                "unitario": 25.0,
                "moneda": "USD",
                "aplica_iva": True,
            },
            {
                "clave": "COMISION_VENDEDOR",
                "concepto": "Comisión del vendedor",
                "monto_usd": 280.0,
            },
            {"clave": "AJUSTE", "concepto": "Descuento", "monto_usd": -33.4},
            {"clave": "IVA", "concepto": "IVA 16%", "monto_usd": 435.2},
        ],
        "tuas_cobradas": [
            {
                "iata": "CUN",
                "pax": 4,
                "unitario": 20.85,
                "moneda": "USD",
                "total_nativo": 83.4,
                "tc_aplicado": None,
                "total_usd": 83.4,
            }
        ],
        "subtotal_vuelo_usd": 2340.0,
        "tuas_usd": 83.4,
        "extras_total_usd": 50.0,
        "viaticos_pernocta_usd": 0,
        "comision_vendedor_usd": 280.0,
        "comision_vendedor_nombre": "Saab",
        "comision_vendedor_modo": "FIJA",
        "comision_vendedor_tarifa_hr": None,
        "iva_comision_vendedor_usd": 44.8,
        "pago_vendedor_usd": 324.8,
        "ajuste_final_usd": -33.4,
        "descuento_usd": 33.4,
        "redondeo_auto_usd": None,
        "total_pactado_usd": None,
        "comision_billpocket_pct": None,
        "subtotal_usd": 2720.0,
        "iva_pct": 16,
        "iva_base_usd": 2720.0,
        "iva_usd": 435.2,
        "iva_nota": None,
        "total_usd": 3155.2,
        "total_mxn": 57645.5,
        "mxn_nativos": 0,
        "version_motor": "v1.3",
        "calculado_at": "2026-06-01T15:12:00Z",
        "cobros": [
            {
                "fecha": "2026-06-03",
                "metodo": "PAYWISE",
                "metodo_label": "Paywise",
                "monto": 36540.0,
                "moneda": "MXN",
                "tc": 18.27,
                "monto_usd": 2000.0,
                "comision_pct": 8.857,
                "comision_monto": 3236.36,
                "neto": 33303.64,
                "referencia": "PW-77812",
                "cuenta_destino": "Paywise",
                "notas": None,
                "conciliado": True,
                "es_reembolso": False,
                "sobre_grupo_folio": None,
                "sobre_grupo_monto_total": None,
                "sobre_grupo_moneda": None,
                "grupo_factor": None,
                "registrado_por": "Itzi",
            }
        ],
        "total_cobrado_usd": 2000.0,
        "cobros_sin_tc_count": 0,
        "cobros_sin_tc_mxn": 0,
        "comision_banco_usd": 177.14,
        "total_cobrado_neto_usd": 1822.86,
        "saldo_usd": 1155.2,
        "cobrado_flag": False,
        "semaforo_cobro": "amarillo",
        "semaforo_cobro_key": "PARCIAL",
        "semaforo_cobro_label": "Parcial",
        "notas_cliente": "Cliente pide <hielo> extra.",
        "notas_internas": "Confirmar pernocta con el hotel.",
        "generado": "2026-09-08T15:12:00Z",
        "generado_cancun": "2026-09-08 10:12",
        "generado_por": "Itzi",
    }
    base.update(extra)
    return base


def _html(**extra) -> str:
    return _build_html(CotizacionInternaPdfRequest(**_payload(**extra)))


def _fila_tramo(html: str, ruta: str) -> str:
    """La fila <tr>…</tr> de la tabla de tramos cuya ruta es `ruta`."""
    i = html.index(f'<td class="ruta">{ruta}')
    ini = html.rindex("<tr>", 0, i)
    return html[ini : html.index("</tr>", i) + 5]


# ===== Identidad del documento =====


def test_banda_interna_cabecera_y_pie() -> None:
    html = _html()
    assert BANDA_INTERNA in html
    assert "uso exclusivo de oficina" in BANDA_INTERNA
    assert 'class="banda"' in html
    assert "Cotización interna" in html
    assert "Folio #1042 · v3" in html
    assert "Confirmado · REDONDO" in html
    # Pie en el margen de página (hora Cancún: 15:12Z → 10:12).
    assert "Documento interno · generado 8 sep 2026 10:12 (hora Cancún) · Itzi" in html
    assert 'Página " counter(page) " de " counter(pages)' in html


def test_reusa_branding_y_helpers_sin_la_hoja_del_cliente() -> None:
    # Importa, no copia: branding y formato del PDF del cliente.
    assert cotizacion_interna_pdf._money is cotizacion_pdf._money
    assert cotizacion_interna_pdf._logo_data_uri is cotizacion_pdf._logo_data_uri
    assert cotizacion_interna_pdf._monto is cotizacion_grupo_pdf._monto
    html = _html()
    assert "#dc2626" in html and "#102a43" in html
    # Pero NUNCA la hoja del cliente: sin leyenda de gracias, sin fotos ni
    # ficha, sin mapa, sin marca de agua.
    assert "Gracias por volar" not in html
    assert "www.vuelatour.com" not in html
    assert "foto-ancha" not in html and 'class="detalles"' not in html
    assert "De un vistazo" not in html
    assert 'viewBox="' not in html and 'class="marca"' not in html
    # El único <img> es el logo del membrete.
    assert html.count("<img") == 1 and 'class="logo"' in html


# ===== Cabecera =====


def test_fecha_del_vuelo_protagonista_y_cotizacion_en_pequeno() -> None:
    html = _html()
    assert 'class="lbl">Fecha del vuelo</div>' in html
    assert '<div class="big">26 jun 2026</div>' in html
    assert "al " not in html[html.index('class="big"') : html.index("Avión cotizado")]
    # La fecha de cotización sigue, pero en pequeño y al final de la ficha.
    assert '<tr class="small"><td class="k">Cotizada</td><td class="v">01/06/2026 10:12' in html
    assert html.index('class="big">26 jun 2026') < html.index(">Cotizada<")
    assert "Fecha de cotización" not in html
    html = _html(fecha_vuelo_fin="2026-06-28", fecha_confirmacion="2026-06-02T14:00:00Z")
    assert '<div class="sub">al 28 jun 2026</div>' in html
    assert "01/06/2026 10:12 · confirmada 02/06/2026 09:00" in html
    # Sin fecha (cotización abierta): «Por definir», nunca una fecha inventada.
    html = _html(fecha_vuelo=None, cotizacion_abierta=True)
    assert '<div class="big">Por definir</div>' in html and "Cotización abierta" in html
    # Un instante ISO (skew) se lleva a su día en Cancún, no en UTC.
    assert '<div class="big">26 jun 2026</div>' in _html(fecha_vuelo="2026-06-27T03:30:00Z")


def test_ficha_cliente_condiciones_avion_con_matricula_y_tripulacion() -> None:
    html = _html()
    assert "Cliente Demo S.A." in html and "RFC CDE010101AAA" in html
    assert "Cliente Demo, S.A. de C.V." in html
    assert "Público · $900.00/hr" in html
    assert "Paywise" in html
    assert "18.27" in html
    assert "Saab" in html and "cotizó Itzi" in html
    # Matrícula SIEMPRE visible (no aplica la regla VGV del cliente).
    assert '<div class="med">Piper Seneca V · N4142R</div>' in html
    assert "Luis Piloto · copiloto Ana Copiloto" in html
    assert "CUN → MID → CUN" in html and "· 4 pasajeros" in html
    # Sin tripulación asignada la fila no aparece (cotización pura).
    assert ">Piloto<" not in _html(piloto=None, copiloto=None)


def test_marcas_broker_interno_grupo_y_externo() -> None:
    html = _html(
        es_broker=True,
        es_interno=True,
        grupo_folio="G-12",
        grupo_posicion=2,
        grupo_total_aviones=3,
        combinado_con_folio="1040",
        itinerario_operativo=True,
        tarifa_preferencial=True,
    )
    assert ">Broker<" in html and ">Cliente interno<" in html
    assert "Grupo G-12 · avión 2 de 3" in html
    assert "Combinado con #1040" in html
    assert "Itinerario operativo" in html
    assert "tarifa preferencial del cliente" in html
    html = _html(
        es_externo=True,
        avion_externo="HAWKER 400 A · XA-REG",
        operador_externo="Aerolíneas Ejecutivas",
        aeronave_cotizada_modelo=None,
        aeronave_cotizada_matricula=None,
    )
    assert "HAWKER 400 A · XA-REG" in html and "operador Aerolíneas Ejecutivas" in html
    assert ">Externo<" in html


# ===== Tabla de tramos (formato de administración) =====


def test_tabla_de_tramos_con_las_seis_columnas_y_total() -> None:
    html = _html()
    assert (
        "<th>Ruta</th><th>Fecha</th>"
        '<th class="num">Distancia millas</th>\n'
        '    <th class="num">Tiempo vuelo</th><th class="num">Costo por hora vuelo</th>\n'
        '    <th class="num">Total por tramo</th>' in html
    )
    fila = _fila_tramo(html, "Cancun-Merida")
    assert '<span class="muted">CUN–MID</span>' in fila
    assert (
        "<td>26-jun</td>"
        '<td class="num">157</td>'
        '<td class="num">01:18</td>'
        '<td class="num">$900.00</td>'
        '<td class="num">$1,170.00</td></tr>' in fila
    )
    assert '<td class="ruta">Merida-Cancun' in html
    # Fila TOTAL en USD con Σ tiempo (del API) y Σ millas (columna informativa).
    assert (
        '<tr class="total"><td>TOTAL</td><td></td><td class="num">314</td>'
        '<td class="num">02:36</td><td></td><td class="num">$2,340.00 USD</td></tr>' in html
    )
    assert f"{NOTA_TRAMOS} (0.3 h en total) · distancia en millas náuticas." in html
    # Sin ajuste: ni fila de ajuste ni «Servicio aéreo» repetido en la tabla.
    assert 'class="ajuste"' not in html
    assert html.count("Servicio aéreo") == 1  # solo en el desglose


def test_fila_de_ajuste_solo_si_no_cuadra_con_el_servicio_aereo() -> None:
    # Hora mínima: 1 tramo de 00:24 ($360) pero se cobra 1.0 h ($900).
    lineas = _payload()["lineas"]
    lineas[0].update(monto_usd=900.0, cantidad=1.0, unitario=900.0)
    html = _html(
        tramos_cotizados=[
            _tramo(1, "CUN", "HOL", tiempo_hr=0.4, tiempo_hhmm="00:24", total_usd=360.0)
        ],
        tramos_tiempo_total_hr=0.4,
        tramos_tiempo_total_hhmm="00:24",
        tramos_total_usd=360.0,
        tramos_ajuste_usd=540.0,
        tramos_ajuste_motivo="Hora mínima 1.0 h",
        tiempo_cobrable_hr=1.0,
        hora_minima_aplicada=True,
        lineas=lineas,
    )
    assert '<td class="num">00:24</td>' in html and "$360.00 USD" in html
    assert (
        '<tr class="ajuste"><td colspan="5">Hora mínima 1.0 h'
        ' <span class="op">servicio aéreo cotizado − Σ tramos</span></td>'
        '<td class="num">$540.00</td></tr>'
        '<tr class="total"><td colspan="5">Servicio aéreo</td>'
        '<td class="num">$900.00 USD</td></tr>' in html
    )
    # El desglose enlaza con la tabla: Σ tramos + ajuste, y marca la hora mínima.
    assert (
        'Servicio aéreo <span class="op">1.00 h × $900.00/hr · hora mínima · '
        "Σ tramos $360.00 + $540.00</span></td>"
        '<td class="val">$900.00' in html
    )
    assert "Cobrables</td>" in html and "hora mínima aplicada" in html
    # Ajuste negativo (horas pactadas por debajo) y motivo ausente → etiqueta genérica.
    html = _html(tramos_ajuste_usd=-90.0, tramos_ajuste_motivo=None, cobrable_override=True)
    assert 'Ajuste de horas <span class="op">' in html and "&minus;$90.00" in html
    assert "Σ tramos $2,340.00 − $90.00" in html and "cobrable manual" in html
    # Diferencias de menos de un centavo no generan fila.
    assert 'class="ajuste"' not in _html(tramos_ajuste_usd=0.004)


def test_tramos_marcas_ferry_pernocta_hhmm_de_respaldo_y_consolidado() -> None:
    tramos = [
        _tramo(1, "CUN", "HOL", es_ferry=True, pax=0, tiempo_hhmm=None, tiempo_hr=0.4),
        _tramo(2, "HOL", "MID", pernocta=True, pernocta_usd=150, fecha="2026-06-27", millas=98.6),
        _tramo(3, "MID", "CUN", fecha=None, millas=None, tarifa_hora_usd=None, total_usd=0),
    ]
    html = _html(tramos_cotizados=tramos, tramos_tiempo_total_hhmm=None, tramos_tiempo_total_hr=3.0)
    f1 = _fila_tramo(html, "Cancun-Isla Holbox")
    assert "CUN–HOL · ferry" in f1 and '<td class="num">00:24</td>' in f1  # hh:mm de tiempo_hr
    f2 = _fila_tramo(html, "Isla Holbox-Merida")
    assert "pernocta $150.00" in f2 and "<td>27-jun</td>" in f2
    assert '<td class="num">98.6</td>' in f2
    f3 = _fila_tramo(html, "Merida-Cancun")
    assert "<td>—</td>" in f3 and f3.count('<td class="num">—</td>') == 2 and "$0.00" in f3
    # Σ millas solo cuando todos los tramos las traen; Σ tiempo de respaldo.
    assert '<td>TOTAL</td><td></td><td class="num"></td><td class="num">03:00</td>' in html
    # Fila ÚNICA de respaldo (snapshot sin tramos): ruta completa y aviso.
    html = _html(
        tramos_cotizados=[
            _tramo(1, "CUN", "CZM", ruta="Cancun-Cozumel-Cancun", consolidado=True, millas=120)
        ]
    )
    assert (
        '<td class="ruta">Cancun-Cozumel-Cancun<br><span class="muted">tramos consolidados' in html
    )
    assert "cotización sin desglose por tramo" in html
    assert "Sin tramos cotizados." in _html(tramos_cotizados=[])


def test_formatos_hhmm_dia_mes_y_millas() -> None:
    assert _hhmm(1.3) == "01:18" and _hhmm(0.4) == "00:24" and _hhmm(2.6) == "02:36"
    assert _hhmm(0) == "00:00" and _hhmm(None) == "—"
    assert _dia_mes("2026-06-26") == "26-jun" and _dia_mes("2026-12-03") == "3-dic"
    assert _dia_mes("2026-06-27T03:00:00Z") == "26-jun"  # instante → día Cancún
    assert _dia_mes(None) == "—" and _dia_mes("") == "—"
    assert _millas(157) == "157" and _millas(157.34) == "157.3" and _millas(1000.0) == "1,000"
    assert _millas(None) == "—"


# ===== Desglose =====


def test_desglose_con_operaciones_tuas_cobradas_y_comision() -> None:
    html = _html()
    assert (
        'Servicio aéreo <span class="op">2.60 h × $900.00/hr</span></td>'
        '<td class="val">$2,340.00' in html
    )
    assert 'TUA CUN <span class="op">4 pax × $20.85</span></td><td class="val">$83.40' in html
    assert 'Hieleras <span class="op">2 × $25.00</span></td><td class="val">$50.00' in html
    # Comisión del vendedor SÍ se ve, con nombre y pago con IVA (pagoVendedorUsd).
    assert (
        'Comisión del vendedor · Saab <span class="op">fija · pago al vendedor c/IVA $324.80</span>'
        '</td><td class="val">$280.00' in html
    )
    assert 'Descuento <span class="op">descuento</span></td><td class="val">&minus;$33.40' in html
    assert "Subtotal (sin IVA)" in html and "$2,720.00" in html
    assert 'IVA 16 % <span class="op">16 % de $2,720.00</span></td><td class="val">$435.20' in html
    assert "<td>Total USD</td>" in html and "$3,155.20" in html
    assert 'Total MXN <span class="op">T.C. 18.27</span></td><td class="val">$57,645.50 MXN' in html
    assert "motor v1.3 · calculado 01/06/2026 10:12" in html
    # El IVA NO sale duplicado como línea del cuerpo.
    assert html.count("$435.20") == 1
    # Panel de horas del snapshot junto al desglose.
    assert "Horas cotizadas" in html
    assert 'Vuelo</td><td class="v">2.30 h' in html
    assert 'Calzos</td><td class="v">0.30 h' in html
    assert "<b>2.60 h</b>" in html and "Tarifa</td><td class=\"v\">$900.00/hr" in html


def test_tuas_exentas_no_se_muestran() -> None:
    # API viejo: línea sintética `exento` y lista `tuas_exentos` → se ignoran.
    lineas = _payload()["lineas"]
    lineas.insert(
        2,
        {
            "clave": "TUAS",
            "concepto": "TUA MID",
            "monto_usd": 0,
            "cantidad": 4,
            "unitario": 0,
            "exento": True,
        },
    )
    html = _html(lineas=lineas, tuas_exentos=["MID"])
    assert "exento" not in html.lower()
    assert "TUA MID" not in html
    assert html.count("TUA CUN") == 1
    # TUAS cobradas en MXN: pax × unitario nativo, total nativo y T.C. congelado.
    html = _html(
        tuas_cobradas=[
            {
                "iata": "CZM",
                "pax": 2,
                "unitario": 380.0,
                "moneda": "MXN",
                "total_nativo": 760.0,
                "tc_aplicado": 18.27,
                "total_usd": 41.6,
            }
        ]
    )
    assert 'TUA CZM <span class="op">2 pax × $380.00 MXN = $760.00 MXN · T.C. 18.27</span>' in html
    assert '<td class="val">$41.60' in html
    # Sin `tuas_cobradas` (skew) las líneas TUAS canónicas con monto se pintan igual.
    html = _html(tuas_cobradas=[])
    assert 'TUA CUN <span class="op">4 pax × $20.85</span>' in html


def test_comision_por_hora_y_ajuste_pactado() -> None:
    lineas = _payload()["lineas"]
    for ln in lineas:
        if ln["clave"] == "COMISION_VENDEDOR":
            ln.update(cantidad=2.6, unitario=175.0)
        if ln["clave"] == "AJUSTE":
            ln.update(concepto="Redondeo", monto_usd=12.0)
    html = _html(lineas=lineas, comision_vendedor_modo="POR_HORA", total_pactado_usd=3200.0)
    assert "2.60 h × $175.00/hr · pago al vendedor c/IVA $324.80" in html
    assert (
        'Redondeo <span class="op">a precio pactado $3,200.00</span></td>'
        '<td class="val">$12.00' in html
    )


# ===== Cobros =====


def test_cobros_compactos_paywise_con_comision_bancaria_neto_y_semaforo() -> None:
    html = _html()
    # Caso real de la foto: bruto 36,540 MXN, 8.857 % = 3,236.36 → neto 33,303.64.
    assert (
        "<th>Fecha</th><th>Método</th>"
        '<th class="num">Bruto</th>\n    <th class="num">Comisión banco</th>'
        '<th class="num">Neto</th><th class="num">Equiv. USD</th><th>Conc.</th>' in html
    )
    assert "<th>Referencia" not in html and '<th class="num">T.C.</th>' not in html
    assert (
        '<td>03/06/2026</td><td>Paywise<br><span class="muted">PW-77812 · Paywise</span></td>'
        in html
    )
    assert '$36,540.00 <span class="muted">MXN</span>' in html
    assert "8.8570 % = $3,236.36" in html
    assert '$33,303.64 <span class="muted">MXN</span>' in html
    # Equivalente USD (cobrosEnUsd) + T.C. del cobro, porque hay cobros en MXN.
    assert '<td class="num">$2,000.00 <span class="muted">T.C. 18.27</span></td>' in html
    assert "registró" not in html  # quién lo capturó no es dato del documento
    assert '<span class="verde">Sí</span>' in html
    assert "Cobrado $2,000.00 USD · comisiones banco &minus;$177.14 · neto $1,822.86" in html
    assert "Saldo $1,155.20" in html
    assert '<span class="sem sem-amarillo"></span>Parcial' in html


def test_cobros_vacios_y_sin_tc() -> None:
    html = _html(
        cobros=[],
        total_cobrado_usd=0,
        saldo_usd=3155.2,
        comision_banco_usd=0,
        total_cobrado_neto_usd=None,
        semaforo_cobro="rojo",
        semaforo_cobro_label="Sin cobro",
        cobros_sin_tc_count=1,
        cobros_sin_tc_mxn=5000,
    )
    assert "Sin cobros registrados." in html
    assert "Equiv. USD" not in html
    assert "Cobrado $0.00 USD · Saldo $3,155.20" in html
    assert '<span class="sem sem-rojo"></span>Sin cobro' in html
    assert "1 cobro en MXN por $5,000.00 SIN tipo de cambio" in html
    # 2.9 → «2.90 %» (dos decimales cuando no hay más precisión).
    cobros = [dict(_payload()["cobros"][0], comision_pct=2.9, comision_monto=1059.66)]
    assert "2.90 % = $1,059.66" in _html(cobros=cobros)
    # Sin subtotal/saldo del API la plantilla NO los calcula: deja «—».
    html = _html(subtotal_usd=None, saldo_usd=None)
    assert 'Subtotal (sin IVA)</td><td class="val">—</td>' in html
    assert "Saldo —" in html


def test_cobro_reembolso_y_sobre_de_grupo() -> None:
    cobros = _payload()["cobros"] + [
        {
            "fecha": "2026-06-05T18:00:00Z",
            "metodo": "TRANSFERENCIA",
            "metodo_label": "Transferencia",
            "monto": -100.0,
            "moneda": "USD",
            "neto": -100.0,
            "es_reembolso": True,
            "conciliado": False,
            "sobre_grupo_folio": "G-12",
            "sobre_grupo_monto_total": 5000.0,
            "sobre_grupo_moneda": "USD",
            "grupo_factor": 0.25,
        }
    ]
    html = _html(cobros=cobros)
    assert "<td>05/06/2026 13:00</td><td>Transferencia<br>" in html
    assert "Reembolso · Sobre G-12 $5,000.00 USD × 25 %" in html
    assert "&minus;$100.00" in html and "<td>No</td>" in html
    # Reembolso en USD sin monto_usd explícito: vale su bruto (cobrosEnUsd),
    # NO se marca "sin T.C." (eso es solo para MXN sin tipo de cambio).
    assert '<td class="num">&minus;$100.00</td><td class="num">&minus;$100.00</td>' in html
    assert "sin T.C." not in html
    cobros[1].update(moneda="MXN", monto=-1800.0, neto=-1800.0, tc=None, monto_usd=None)
    html = _html(cobros=cobros)
    assert '<span class="rojo">sin T.C.</span>' in html


# ===== Notas =====


def test_notas_internas_escapadas_truncadas_y_sin_las_del_cliente() -> None:
    larga = " ".join(f"palabra{i} <x>" for i in range(120))
    html = _html(notas_internas=larga)
    assert "<h2>Notas internas</h2>" in html
    assert "&lt;x&gt;" in html and "<x>" not in html
    assert "…</div></div>" in html and "palabra119" not in html
    # Las notas del cliente NO son de este documento.
    assert "Notas del cliente" not in html and "hielo" not in html
    assert len(_truncar(larga)) <= 281
    assert _truncar("a\nb\nc\nd\ne\nf\ng").endswith("…")
    assert _truncar("corta") == "corta"
    assert "Notas internas" not in _html(notas_internas=None)


# ===== Bloques que administración pidió QUITAR =====


def test_no_pinta_operacion_particion_gastos_ni_cfdi_aunque_el_api_los_mande() -> None:
    html = _html(
        # Campos legado v1 (API anterior): aceptados por el schema, jamás pintados.
        tramos=[{"orden": 1, "origen": "CUN", "destino": "MID", "taco_salida": 1234.0}],
        horas_voladas_hr=2.9,
        delta_horas_hr=0.3,
        aeronave_operativa="Cessna 206 · XB-RTO",
        apoyos=["Pepe Apoyo"],
        fecha_traslado_inicial="2026-06-26T12:00:00Z",
        fecha_traslado_final="2026-06-26T20:00:00Z",
        venta_avion_usd=2500.0,
        otros_ingresos_vuelatour_usd=655.2,
        neto_vuelatour_usd=2830.4,
        particion_fuente="desglose",
        participacion_aviones=[{"matricula": "N4142R", "factor": 1}],
        gastos_por_categoria=[{"categoria": "GAS", "etiqueta": "Gasavión", "total_usd": 310.5}],
        gastos_total_usd=310.5,
        utilidad_bruta_usd=1689.5,
        utilidad_base="cobrado",
        facturado=True,
        facturas=[{"serie": "A", "folio": "123"}],
        cfdi_estatus="TIMBRADA",
        cfdi_folio="A-123",
    )
    for ausente in (
        "Taco",
        "Voladas",
        "Δ",
        "Avión operativo",
        "XB-RTO",
        "Pepe Apoyo",
        "Traslado",
        "Itinerario</h2>",
        "Partición",
        "Venta del avión",
        "Neto VuelaTour",
        "Gastos del vuelo",
        "Gasavión",
        "Utilidad",
        "Facturación",
        "CFDI",
        "A-123",
        "TIMBRADA",
    ):
        assert ausente not in html, ausente


# ===== Skew / payload mínimo =====


def test_payload_minimo_renderiza_y_campos_extra_se_ignoran() -> None:
    req = CotizacionInternaPdfRequest(campo_nuevo_del_futuro=1)
    html = _build_html(req)
    assert BANDA_INTERNA in html
    assert "Folio s/n" in html
    assert "Por definir" in html
    assert "Sin tramos cotizados." in html
    assert "Sin desglose (cotización sin precio)." in html
    assert "Sin cobros registrados." in html
    assert "Horas cotizadas" not in html
    assert "Documento interno · generado" in html


def test_sin_lineas_usa_escalares_espejo() -> None:
    html = _html(
        lineas=[],
        subtotal_usd=None,
        iva_base_usd=None,
        version_motor=None,
        calculado_at=None,
    )
    assert 'Servicio aéreo <span class="op">2.60 h × $900.00/hr</span>' in html
    assert 'TUA CUN <span class="op">4 pax × $20.85</span>' in html and "$83.40" in html
    assert "Extras" in html and "$50.00" in html
    assert "Comisión del vendedor · Saab" in html and "$280.00" in html
    assert "&minus;$33.40" in html
    # Sin subtotal del API la plantilla NO lo deriva (cero recálculo): «—».
    assert 'Subtotal (sin IVA)</td><td class="val">—</td>' in html
    assert 'IVA 16 %</td><td class="val">$435.20' in html


# ===== Router =====

TOKEN = "secreto-de-prueba"
client = TestClient(app)


def test_router_sin_token_rechazado(monkeypatch) -> None:
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    res = client.post("/reportes/cotizacion-interna", json=_payload())
    assert res.status_code == 401


def test_router_devuelve_pdf_con_nombre_de_archivo(monkeypatch) -> None:
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    capturado: dict = {}

    def _render_falso(req: CotizacionInternaPdfRequest) -> bytes:
        capturado["req"] = req
        return b"%PDF-1.4 fake"

    monkeypatch.setattr(reportes_router, "render_cotizacion_interna_pdf", _render_falso)
    res = client.post(
        "/reportes/cotizacion-interna",
        json=_payload(folio="COT/1042"),
        headers={"X-Internal-Token": TOKEN},
    )
    assert res.status_code == 200
    assert res.headers["content-type"] == "application/pdf"
    assert 'filename="cotizacion-interna-COT1042.pdf"' in res.headers["content-disposition"]
    assert res.content == b"%PDF-1.4 fake"
    assert capturado["req"].aeronave_cotizada_matricula == "N4142R"
    assert len(capturado["req"].tramos_cotizados) == 2
    assert capturado["req"].tramos_cotizados[0].tiempo_hhmm == "01:18"
    assert len(capturado["req"].cobros) == 1


def test_router_error_de_render_es_500_con_detalle(monkeypatch) -> None:
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()

    def _render_roto(req: CotizacionInternaPdfRequest) -> bytes:
        raise OSError("cannot load library 'libgobject-2.0-0'")

    monkeypatch.setattr(reportes_router, "render_cotizacion_interna_pdf", _render_roto)
    res = client.post(
        "/reportes/cotizacion-interna",
        json=_payload(),
        headers={"X-Internal-Token": TOKEN},
    )
    assert res.status_code == 500
    assert "libgobject" in res.json()["detail"]


# ===== Una hoja (solo si WeasyPrint y sus libs de sistema están) =====


def test_cabe_en_una_hoja_carta() -> None:
    import pytest

    try:
        from weasyprint import HTML
    except Exception as e:  # noqa: BLE001 — ImportError u OSError (pango/cairo)
        pytest.skip(f"WeasyPrint no disponible: {e}")
    tramos = [
        _tramo(i, "CUN" if i % 2 else "MID", "MID" if i % 2 else "CUN", pernocta=i == 4)
        for i in range(1, 9)
    ]
    cobros = _payload()["cobros"] * 4
    html = _html(
        tramos_cotizados=tramos,
        cobros=cobros,
        tramos_ajuste_usd=540.0,
        tramos_ajuste_motivo="Hora mínima 1.0 h",
        notas_internas="\n".join(f"Nota {i}" for i in range(5)),
    )
    assert len(HTML(string=html).render().pages) == 1
