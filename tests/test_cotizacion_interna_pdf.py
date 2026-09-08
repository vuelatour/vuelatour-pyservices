"""PDF de la COTIZACIÓN INTERNA (8-sep-2026), probado sobre el HTML (sin
WeasyPrint, import perezoso): banda «interna», matrícula SIEMPRE visible,
itinerario con tacómetros, desglose completo con comisión del vendedor y
operaciones en gris, partición, cobros con comisión bancaria (caso real de
la foto: Paywise 8.857 % = $3,236.36 → neto $33,303.64), gastos, notas
truncadas y pie «Documento interno · generado … · usuario».

Reglas: nada se recalcula (todo viene del API); jamás fotos ni ficha del
avión; un payload mínimo (skew) y campos extra también renderizan.
"""

from fastapi.testclient import TestClient

from app.config import get_settings
from app.main import app
from app.routers import reportes as reportes_router
from app.schemas.reportes import CotizacionInternaPdfRequest
from app.services import cotizacion_grupo_pdf, cotizacion_interna_pdf, cotizacion_pdf
from app.services.cotizacion_interna_pdf import BANDA_INTERNA, _build_html, _truncar


def _tramo(orden: int, origen: str, destino: str, **extra) -> dict:
    base = {
        "orden": orden,
        "orden_real": orden,
        "origen": origen,
        "destino": destino,
        "pasajeros": 4,
        "fecha_plan": f"2026-09-1{orden}T14:30:00Z",  # 09:30 Cancún
        "hora_salida": None,
        "hora_llegada": None,
        "matricula": "N4142R",
        "piloto": "Luis Piloto",
        "taco_salida": 1234.0 + (orden - 1) * 0.4,
        "taco_llegada": 1234.4 + (orden - 1) * 0.4,
        "taco_salida_origen": "PILOTO",
        "taco_llegada_origen": "PILOTO",
        "horas_taco": 0.4,
        "horas_cotizadas": 0.4,
        "es_ferry": False,
        "solo_operativa": False,
        "es_sobrevuelo": False,
        "requiere_pernocta": False,
        "pernocta_usd": 0,
        "cancelado": False,
        "cancelada_motivo": None,
        "revision_requerida": False,
    }
    base.update(extra)
    return base


def _payload(**extra) -> dict:
    """Payload COMPLETO con la forma del contrato de pyservices.service.ts
    (números coherentes: 1.6 h × $950 + TUA + extra + comisión − descuento,
    IVA 16 % = $2,204.00; cobro Paywise en MXN parcial)."""
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
        "fecha": "2026-09-01T15:12:00Z",
        "fecha_confirmacion": None,
        "tarifa_tipo": "PUBLICO",
        "tarifa_tipo_label": "Público",
        "tarifa_hora_usd": 950.0,
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
        "aeronave_operativa": None,
        "avion_externo": None,
        "operador_externo": None,
        "piloto": "Luis Piloto",
        "copiloto": "Ana Copiloto",
        "apoyos": [],
        "fecha_traslado_inicial": "2026-09-11T14:30:00Z",
        "fecha_traslado_final": "2026-09-14T20:00:00Z",
        "pasajeros": 4,
        "ruta": "CUN → HOL → CUN → HOL → CUN",
        "itinerario_operativo": False,
        "cotizacion_abierta": False,
        "es_interno": False,
        "es_externo": False,
        "grupo_folio": None,
        "grupo_posicion": None,
        "grupo_total_aviones": None,
        "combinado_con_folio": None,
        "tramos": [
            _tramo(1, "CUN", "HOL"),
            _tramo(2, "HOL", "CUN"),
            _tramo(3, "CUN", "HOL"),
            _tramo(4, "HOL", "CUN"),
        ],
        "horas_cotizadas_hr": 1.6,
        "vuelo_hr": 1.6,
        "calzos_hr": 0,
        "sobrevuelo_hr": 0,
        "horas_voladas_hr": 1.6,
        "delta_horas_hr": 0.0,
        "tiempo_cobrable_hr": 1.6,
        "hora_minima_aplicada": False,
        "cobrable_override": False,
        "lineas": [
            {
                "clave": "TIEMPO_VUELO",
                "concepto": "Servicio aéreo",
                "monto_usd": 1520.0,
                "cantidad": 1.6,
                "unitario": 950.0,
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
                "clave": "TUAS",
                "concepto": "TUA HOL",
                "monto_usd": 0,
                "cantidad": 4,
                "unitario": 0,
                "exento": True,
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
            {"clave": "IVA", "concepto": "IVA 16%", "monto_usd": 304.0},
        ],
        "tuas_exentos": ["HOL"],
        "subtotal_vuelo_usd": 1520.0,
        "tuas_usd": 83.4,
        "extras_total_usd": 50.0,
        "viaticos_pernocta_usd": 0,
        "comision_vendedor_usd": 280.0,
        "comision_vendedor_nombre": "Saab",
        "comision_vendedor_modo": "FIJA",
        "comision_vendedor_tarifa_hr": None,
        "iva_comision_vendedor_usd": 44.8,
        "pago_vendedor_usd": 324.8,
        "neto_vuelatour_usd": 1879.2,
        "ajuste_final_usd": -33.4,
        "descuento_usd": 33.4,
        "redondeo_auto_usd": None,
        "total_pactado_usd": None,
        "comision_billpocket_pct": None,
        "subtotal_usd": 1900.0,
        "iva_pct": 16,
        "iva_base_usd": 1900.0,
        "iva_usd": 304.0,
        "iva_nota": None,
        "total_usd": 2204.0,
        "total_mxn": 40267.08,
        "mxn_nativos": 0,
        "version_motor": "v1.3",
        "calculado_at": "2026-09-01T15:12:00Z",
        "venta_avion_usd": 1724.46,
        "otros_ingresos_vuelatour_usd": 479.54,
        "iva_avion_usd": 237.86,
        "iva_vuelatour_usd": 66.14,
        "particion_fuente": "desglose",
        "particion_inconsistente": False,
        "participacion_aviones": [],
        "cobros": [
            {
                "fecha": "2026-09-03",
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
        "saldo_usd": 204.0,
        "cobrado_flag": False,
        "semaforo_cobro": "amarillo",
        "semaforo_cobro_key": "PARCIAL",
        "semaforo_cobro_label": "Parcial",
        "gastos_por_categoria": [
            {"categoria": "GAS", "etiqueta": "Gasavión / Turbosina", "total_usd": 310.5, "n": 2},
            {"categoria": "PILOTO", "etiqueta": "Viáticos piloto", "total_usd": 60.0, "n": 1},
        ],
        "gastos_total_usd": 370.5,
        "gastos_sin_tc_count": 0,
        "gastos_sin_tc_mxn": 0,
        "costo_externo_usd": None,
        "utilidad_bruta_usd": 1629.5,
        "utilidad_base": "cobrado",
        "notas_cliente": "Cliente pide <hielo> extra.",
        "notas_internas": "Confirmar pernocta con el hotel.",
        "facturado": False,
        "facturas": [],
        "cfdi_estatus": None,
        "cfdi_folio": None,
        "generado": "2026-09-08T15:12:00Z",
        "generado_cancun": "2026-09-08 10:12",
        "generado_por": "Itzi",
    }
    base.update(extra)
    return base


def _html(**extra) -> str:
    return _build_html(CotizacionInternaPdfRequest(**_payload(**extra)))


# ===== Identidad del documento =====


def test_banda_interna_cabecera_y_pie() -> None:
    html = _html()
    assert BANDA_INTERNA in html
    assert "uso exclusivo de oficina" in BANDA_INTERNA
    assert "Cotización interna" in html
    assert "Folio #1042 · v3" in html
    assert "Confirmado · REDONDO" in html
    # Pie en el margen de página (hora Cancún: 15:12Z → 10:12).
    assert "Documento interno · generado 8 sep 2026 10:12 (hora Cancún) · Itzi" in html
    assert "Página \" counter(page) \" de \" counter(pages)" in html


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


def test_ficha_cliente_condiciones_y_avion_con_matricula_siempre() -> None:
    html = _html()
    assert "Cliente Demo S.A." in html and "RFC CDE010101AAA" in html
    assert "Cliente Demo, S.A. de C.V." in html
    assert "01/09/2026 10:12" in html  # fecha de cotización en Cancún
    assert "Público · $950.00/hr" in html
    assert "Paywise" in html
    assert "18.27" in html
    assert "Saab" in html and "cotizó Itzi" in html
    # Matrícula SIEMPRE visible (no aplica la regla VGV del cliente).
    assert "Piper Seneca V · N4142R" in html
    assert "Avión operativo" not in html
    assert "Luis Piloto · copiloto Ana Copiloto" in html
    assert "11/09/2026 09:30" in html and "14/09/2026 15:00" in html
    assert "CUN → HOL → CUN → HOL → CUN · 4 pasajeros" in html


def test_avion_operativo_solo_si_difiere_y_marcas() -> None:
    html = _html(
        aeronave_operativa="Cessna 206 · XB-RTO",
        cotizacion_abierta=True,
        grupo_folio="G-12",
        grupo_posicion=2,
        grupo_total_aviones=3,
        combinado_con_folio="1040",
        es_broker=True,
    )
    assert "Avión operativo" in html and "Cessna 206 · XB-RTO" in html
    assert "difiere del cotizado" in html
    assert "Cotización abierta" in html
    assert "Grupo G-12 · avión 2 de 3" in html
    assert "Combinado con #1040" in html
    assert ">Broker<" in html


def test_avion_externo() -> None:
    html = _html(
        es_externo=True,
        avion_externo="HAWKER 400 A · XA-REG",
        operador_externo="Aerolíneas Ejecutivas",
        aeronave_cotizada_modelo=None,
        aeronave_cotizada_matricula=None,
    )
    assert "HAWKER 400 A · XA-REG" in html and "operador Aerolíneas Ejecutivas" in html
    assert ">Externo<" in html


# ===== Itinerario =====


def test_itinerario_tacos_por_tramo_y_totales() -> None:
    html = _html()
    assert "<th>Matrícula</th>" in html
    assert (
        '<td>1</td><td class="tramo">CUN → HOL</td><td class="num">4</td><td>11/09 09:30</td>'
        '<td>N4142R</td><td class="num">1234.0</td><td class="num">1234.4</td>'
        '<td class="num">0.4</td><td class="num">0.40</td>' in html
    )
    assert html.count("<td>N4142R</td>") == 4
    assert "Cotizadas 1.60 h · Voladas (tacos) 1.6 h · Δ +0.0 h · Cobrables 1.60 h" in html


def test_itinerario_marcas_cancelado_ferry_pernocta_y_origen_taco() -> None:
    tramos = [
        _tramo(1, "CUN", "HOL", es_ferry=True, pasajeros=0),
        _tramo(2, "HOL", "MID", requiere_pernocta=True, pernocta_usd=150, taco_salida_origen="IA"),
        _tramo(
            3,
            "MID",
            "CUN",
            cancelado=True,
            cancelada_motivo="clima",
            taco_llegada_origen="DEDUCIDO",
        ),
    ]
    html = _html(
        tramos=tramos,
        horas_voladas_hr=0.8,
        horas_cotizadas_hr=1.2,
        delta_horas_hr=-0.4,
        hora_minima_aplicada=True,
    )
    assert "Ferry" in html
    assert "Pernocta $150.00" in html
    assert '<tr class="cancelado">' in html and "Cancelado: clima" in html
    assert '1234.4 <span class="muted">IA</span>' in html
    assert '<span class="muted">ded.</span>' in html
    assert "Voladas (tacos) 0.8 h · Δ -0.4 h" in html
    assert "Cobrables 1.60 h (hora mínima aplicada)" in html


def test_itinerario_sin_tacos_deja_celdas_vacias() -> None:
    tramos = [
        _tramo(1, "CUN", "HOL", taco_salida=None, taco_llegada=None, horas_taco=None),
        _tramo(2, "HOL", "CUN", taco_salida=None, taco_llegada=None, horas_taco=None),
    ]
    html = _html(tramos=tramos, horas_voladas_hr=None)
    assert (
        '<td class="num"></td><td class="num"></td><td class="num"></td><td class="num">0.40</td>'
        in html
    )
    assert "Voladas (tacos) —" in html


# ===== Desglose interno =====


def test_desglose_completo_con_operaciones_y_comision() -> None:
    html = _html()
    assert (
        'Servicio aéreo <span class="op">1.60 h × $950.00/hr</span></td>'
        '<td class="val">$1,520.00' in html
    )
    assert 'TUA CUN <span class="op">4 pax × $20.85</span></td><td class="val">$83.40' in html
    # TUA exento como línea sintética (sin monto).
    assert 'TUA HOL <span class="op">4 pax · exento</span></td><td class="val">—' in html
    assert 'Hieleras <span class="op">2 × $25.00</span></td><td class="val">$50.00' in html
    # Comisión del vendedor SÍ se ve, con nombre y pago con IVA (pagoVendedorUsd).
    assert (
        'Comisión del vendedor · Saab <span class="op">fija · pago al vendedor c/IVA $324.80</span>'
        '</td><td class="val">$280.00' in html
    )
    assert 'Descuento <span class="op">descuento</span></td><td class="val">&minus;$33.40' in html
    assert "Subtotal (sin IVA)" in html and "$1,900.00" in html
    assert 'IVA 16 % <span class="op">sobre $1,900.00</span></td><td class="val">$304.00' in html
    assert "<td>Total USD</td>" in html and "$2,204.00" in html
    assert 'Total MXN <span class="op">T.C. 18.27</span></td><td class="val">$40,267.08 MXN' in html
    assert "motor v1.3 · calculado 01/09/2026 10:12" in html
    # El IVA NO sale duplicado como línea del cuerpo.
    assert html.count("$304.00") == 1


def test_particion_del_ingreso_y_neto() -> None:
    html = _html()
    assert "Venta del avión" in html and "$1,724.46" in html and "incluye IVA $237.86" in html
    assert "Ingreso VuelaTour" in html and "$479.54" in html
    assert "Pago al vendedor" in html and "&minus;$324.80" in html
    assert "comisión $280.00 + IVA $44.80 · Saab" in html
    assert "Neto VuelaTour" in html and "<b>$1,879.20</b>" in html


def test_particion_multi_avion_por_matricula() -> None:
    html = _html(
        participacion_aviones=[
            {"matricula": "N4142R", "factor": 0.5, "tramos": 2, "venta_usd": 862.23},
            {"matricula": "XB-RTO", "factor": 0.5, "tramos": 2, "venta_usd": 862.23},
        ],
        particion_inconsistente=True,
    )
    assert "· N4142R" in html and "· XB-RTO" in html
    assert "50 % (2 tramos) = $862.23" in html
    assert "Partición inconsistente" in html


def test_comision_por_hora_y_ajuste_pactado() -> None:
    lineas = _payload()["lineas"]
    for ln in lineas:
        if ln["clave"] == "COMISION_VENDEDOR":
            ln.update(cantidad=1.6, unitario=175.0)
        if ln["clave"] == "AJUSTE":
            ln.update(concepto="Redondeo", monto_usd=12.0)
    html = _html(lineas=lineas, comision_vendedor_modo="POR_HORA", total_pactado_usd=2250.0)
    assert "1.60 h × $175.00/hr · pago al vendedor c/IVA $324.80" in html
    assert (
        'Redondeo <span class="op">a precio pactado $2,250.00</span></td>'
        '<td class="val">$12.00' in html
    )


# ===== Cobros =====


def test_cobros_paywise_con_comision_bancaria_neto_y_semaforo() -> None:
    html = _html()
    # Caso real de la foto: bruto 36,540 MXN, 8.857 % = 3,236.36 → neto 33,303.64.
    assert "<td>03/09/2026</td><td>Paywise</td>" in html
    assert '$36,540.00 <span class="muted">MXN</span>' in html
    assert "8.8570 % = $3,236.36" in html
    assert '$33,303.64 <span class="muted">MXN</span>' in html
    # Equivalente USD (cobrosEnUsd) porque hay cobros en MXN.
    assert '<th class="num">Equiv. USD</th>' in html and "$2,000.00" in html
    assert "PW-77812 · Paywise" in html
    assert "registró" not in html  # quién lo capturó no es dato del documento
    assert '<span class="verde">Sí</span>' in html
    assert "Cobrado $2,000.00 USD · comisiones banco &minus;$177.14 · neto $1,822.86" in html
    assert "Saldo $204.00" in html
    assert '<span class="sem sem-amarillo"></span>Parcial' in html


def test_cobros_vacios_y_sin_tc() -> None:
    html = _html(
        cobros=[],
        total_cobrado_usd=0,
        saldo_usd=2204.0,
        comision_banco_usd=0,
        total_cobrado_neto_usd=None,
        semaforo_cobro="rojo",
        semaforo_cobro_label="Sin cobro",
        cobros_sin_tc_count=1,
        cobros_sin_tc_mxn=5000,
    )
    assert "Sin cobros registrados." in html
    assert "Equiv. USD" not in html
    assert "Cobrado $0.00 USD · Saldo $2,204.00" in html
    assert '<span class="sem sem-rojo"></span>Sin cobro' in html
    assert "1 cobro en MXN por $5,000.00 SIN tipo de cambio" in html


def test_cobro_reembolso_y_sobre_de_grupo() -> None:
    cobros = _payload()["cobros"] + [
        {
            "fecha": "2026-09-05T18:00:00Z",
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
    assert "<td>05/09/2026 13:00</td><td>Transferencia<br>" in html
    assert "Reembolso" in html and "&minus;$100.00" in html
    assert "Sobre G-12 $5,000.00 USD × 25 %" in html
    assert "<td>No</td>" in html
    # Reembolso en USD sin monto_usd explícito: vale su bruto (cobrosEnUsd),
    # NO se marca "sin T.C." (eso es solo para MXN sin tipo de cambio).
    assert '<td class="num">&minus;$100.00</td><td class="num">&minus;$100.00</td>' in html
    assert "sin T.C." not in html
    cobros[1].update(moneda="MXN", monto=-1800.0, neto=-1800.0, tc=None, monto_usd=None)
    html = _html(cobros=cobros)
    assert '<span class="rojo">sin T.C.</span>' in html


# ===== Gastos, facturación y notas =====


def test_gastos_por_categoria_total_y_utilidad_de_referencia() -> None:
    html = _html()
    assert "Gasavión / Turbosina (2)" in html and "$310.50" in html
    assert "Viáticos piloto (1)" in html and "$60.00" in html
    assert "Total gastos" in html and "<b>$370.50</b>" in html
    assert "Utilidad bruta (referencia)" in html
    assert '<b class="verde">$1,629.50</b>' in html
    assert "cobrado $2,000.00 − gastos $370.50 · no es el reparto" in html


def test_gastos_pinta_todas_las_categorias_sin_sumar_y_avisa_sin_tc() -> None:
    cats = [
        {"categoria": f"C{i}", "etiqueta": f"Cat {i}", "total_usd": 10.0 * (10 - i), "n": 1}
        for i in range(9)
    ]
    html = _html(gastos_por_categoria=cats, gastos_sin_tc_count=2, gastos_sin_tc_mxn=1200)
    # Todas las categorías tal cual llegan (el API ya agregó); la plantilla no
    # agrupa ni re-suma nada.
    for i in range(9):
        assert f"Cat {i} (1)" in html
    assert "Otras categorías" not in html
    assert "2 gastos en MXN por $1,200.00 sin tipo de cambio" in html


def test_comision_bancaria_pct_con_2_a_4_decimales_y_delta_horas_del_api() -> None:
    # 8.857 → «8.8570 %» (como en la foto anotada a mano); 2.9 → «2.90 %».
    cobros = [dict(_payload()["cobros"][0], comision_pct=8.857)]
    html = _html(cobros=cobros, delta_horas_hr=-0.2)
    assert "8.8570 % = $3,236.36" in html
    assert "Δ -0.2 h" in html  # viene del API, no se resta aquí
    cobros = [dict(_payload()["cobros"][0], comision_pct=2.9, comision_monto=1059.66)]
    html = _html(cobros=cobros, delta_horas_hr=None)
    assert "2.90 % = $1,059.66" in html
    assert "Δ" not in html
    # Sin subtotal/saldo del API la plantilla NO los calcula: deja «—».
    html = _html(subtotal_usd=None, saldo_usd=None)
    assert "Subtotal (sin IVA)</td><td class=\"val\">—</td>" in html
    assert "Saldo —" in html


def test_sin_gastos_y_sin_cfdi() -> None:
    html = _html(gastos_por_categoria=[], gastos_total_usd=None, utilidad_bruta_usd=None)
    assert "Sin gastos ligados al vuelo." in html
    assert "Utilidad bruta" not in html
    assert "<h2>Facturación</h2>" not in html


def test_facturacion_cfdi() -> None:
    html = _html(
        cfdi_estatus="TIMBRADA",
        cfdi_folio="A-123",
        facturas=[
            {
                "serie": "A",
                "folio": "123",
                "estado": "TIMBRADA",
                "total": 40267.08,
                "moneda": "MXN",
                "fecha_timbrado": "2026-09-04",
            },
            {"serie": "A", "folio": "120", "estado": "CANCELADA", "cancelada": True},
        ],
    )
    assert "<h2>Facturación</h2>" in html
    assert "TIMBRADA · A-123" in html
    assert "A-123 · TIMBRADA · $40,267.08 MXN · 04/09/2026" in html
    assert "<s>A-120 · CANCELADA</s>" in html


def test_notas_escapadas_y_truncadas() -> None:
    larga = " ".join(f"palabra{i}" for i in range(120))
    html = _html(notas_internas=larga, notas_cliente="Cliente pide <hielo> extra.")
    assert "Cliente pide &lt;hielo&gt; extra." in html
    bloque = html[html.index("Notas internas") : html.index("Notas del cliente")]
    assert "…</div></td>" in bloque
    assert "palabra119" not in html
    assert len(_truncar(larga)) <= 281
    assert _truncar("a\nb\nc\nd\ne\nf\ng\nh").endswith("…")
    assert _truncar("corta") == "corta"
    assert "Notas internas" not in _html(notas_internas=None, notas_cliente=None)


# ===== Skew / payload mínimo =====


def test_payload_minimo_renderiza_y_campos_extra_se_ignoran() -> None:
    req = CotizacionInternaPdfRequest(campo_nuevo_del_futuro=1)
    html = _build_html(req)
    assert BANDA_INTERNA in html
    assert "Folio s/n" in html
    assert "Sin tramos." in html
    assert "Sin desglose (cotización sin precio)." in html
    assert "Sin partición" in html
    assert "Sin cobros registrados." in html
    assert "Sin gastos ligados al vuelo." in html
    assert "Documento interno · generado" in html


def test_sin_lineas_usa_escalares_espejo() -> None:
    html = _html(
        lineas=[],
        subtotal_usd=None,
        iva_base_usd=None,
        version_motor=None,
        calculado_at=None,
    )
    assert 'Servicio aéreo <span class="op">1.60 h × $950.00/hr</span>' in html
    assert "TUAS" in html and "$83.40" in html
    assert "Extras" in html and "$50.00" in html
    assert "Comisión del vendedor · Saab" in html and "$280.00" in html
    assert "&minus;$33.40" in html
    # Sin subtotal del API la plantilla NO lo deriva (cero recálculo): «—».
    assert 'Subtotal (sin IVA)</td><td class="val">—</td>' in html
    assert "TUA exento" in html and "HOL" in html


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
    tramos = [_tramo(i, "CUN" if i % 2 else "HOL", "HOL" if i % 2 else "CUN") for i in range(1, 7)]
    cobros = _payload()["cobros"] * 4
    cats = [
        {"categoria": f"C{i}", "etiqueta": f"Categoría {i}", "total_usd": 10.0, "n": 1}
        for i in range(5)
    ]
    html = _html(tramos=tramos, cobros=cobros, gastos_por_categoria=cats)
    assert len(HTML(string=html).render().pages) == 1
