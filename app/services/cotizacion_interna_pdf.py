"""PDF de la COTIZACIÓN INTERNA (8-sep-2026): UNA hoja carta para la oficina.

Pedido del cliente (foto de una cotización impresa anotada a mano):
administración necesita imprimir la cotización SIN las fotos/fichas del
avión que van al cliente, y con TODO lo interno a la vista: modelo y
matrícula (siempre, también por tramo), tacómetros y horas voladas vs
cotizadas, comisión del vendedor, T.C., partición del ingreso, cobros con su
comisión bancaria («Paywise 8.857 % = $3,236.36 → neto $33,303.64») y gastos.

NUNCA se manda al cliente: lleva banda «COTIZACIÓN INTERNA · uso exclusivo
de oficina» y pie «Documento interno · generado … · <usuario>».

Hermano de `cotizacion_pdf.py`: REUTILIZA (importa, no copia) el branding
(#dc2626 / #102a43), el logo, `_money`/`_monto` y los formatos de fecha. NO
importa `_estilos_base` a propósito: esa hoja es la del cliente (13 px,
márgenes de 2 cm, marca de agua y «Gracias por volar con VuelaTour» en el
pie) — aquí el CSS es propio, denso (8.5–9.5 pt) y con @page de 10–12 mm
para que todo quepa en una hoja. Tampoco aplica `_mostrar_matricula`: en el
documento interno la matrícula SIEMPRE se ve.

Aquí SOLO se pinta: el API arma el payload con el desglose canónico v1.3
del snapshot, cobrosEnUsd, particionIngresoVuelo/pagoVendedorUsd y las horas
derivadas de escalas/tacos. Las operaciones («1.60 h × $950.00/hr», «4 pax ×
$20.85») se muestran con los números que llegan; nunca se recalcula dinero.

Presupuesto de UNA hoja (≈ 733 pt útiles): cabecera + ficha ≈ 125, itinerario
≈ 14 pt/tramo, desglose | partición+gastos ≈ 200–270, cobros ≈ 14 + 24/cobro,
notas ≈ 55. Con más de ~8 tramos el itinerario puede pasar a una segunda hoja
(el `thead` se repite); los demás bloques evitan partirse (`page-break-inside`).
"""

from datetime import UTC, datetime
from html import escape

from app.schemas.reportes import (
    CotizacionInternaCobroPdf,
    CotizacionInternaLineaPdf,
    CotizacionInternaPdfRequest,
    CotizacionInternaTramoPdf,
)
from app.services.cotizacion_grupo_pdf import _RE_PAX_SUFIJO, _monto, _plural
from app.services.cotizacion_pdf import (
    _BRAND,
    _CANCUN,
    _NAVY,
    _fecha_dia,
    _logo_data_uri,
    _money,
)
from app.services.reporte_vuelo_pdf import _fecha, _pct, _tramos_txt

# Texto de la banda: es lo que distingue este documento del que va al cliente.
BANDA_INTERNA = "Cotización interna · uso exclusivo de oficina · no enviar al cliente"

# Truncado elegante de notas: prioridad a (1)–(4); las notas van al final.
_NOTAS_MAX_CHARS = 280
_NOTAS_MAX_LINEAS = 6
_FACTURAS_MAX = 3

_SEMAFORO_CLASES = {"verde": "sem-verde", "amarillo": "sem-amarillo", "rojo": "sem-rojo"}
_ORIGEN_TACO_ABREV = {"IA": "IA", "DEDUCIDO": "ded."}


# ===== Formato =====


def _fecha_corta(s: str | None, con_anio: bool = False) -> str:
    """Instante ISO → 'dd/mm HH:MM' en hora Cancún (con año si se pide).
    Un día de pared ('2026-09-03') NO se convierte (movería el día)."""
    if not s:
        return "—"
    try:
        if len(s) == 10 and "T" not in s:
            d = datetime.fromisoformat(s)
            return d.strftime("%d/%m/%Y" if con_anio else "%d/%m")
        dt = datetime.fromisoformat(s.replace("Z", "+00:00"))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=UTC)
        dt = dt.astimezone(_CANCUN)
        return dt.strftime("%d/%m/%Y %H:%M" if con_anio else "%d/%m %H:%M")
    except ValueError:
        return escape(s)


def _hora(s: str | None) -> str:
    """Solo la hora Cancún de un instante ('14:32'); '' si no hay."""
    if not s:
        return ""
    try:
        dt = datetime.fromisoformat(s.replace("Z", "+00:00"))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=UTC)
        return dt.astimezone(_CANCUN).strftime("%H:%M")
    except ValueError:
        return escape(s)


def _horas(v: float | None, dec: int = 2) -> str:
    return "—" if v is None else f"{v:.{dec}f} h"


def _taco(v: float | None) -> str:
    return "" if v is None else f"{v:.1f}"


def _tc(v: float | None) -> str:
    return "—" if not v else f"{v:g}"


def _pct_banco(v: float) -> str:
    """% de comisión bancaria como lo escribe la oficina: 2 decimales si no
    hay más precisión (2.90 %), 4 si la hay (8.8570 %, foto del cliente).
    El valor llega en PUNTOS porcentuales (cobro_vuelo.comision_banco_pct)."""
    dec = 2 if round(v, 2) == round(v, 4) else 4
    return f"{v:.{dec}f} %"


def _truncar(
    txt: str | None,
    max_chars: int = _NOTAS_MAX_CHARS,
    max_lineas: int = _NOTAS_MAX_LINEAS,
) -> str:
    """Recorta a N caracteres (en un espacio) y a N líneas, cerrando con «…».
    Devuelve texto plano (sin escapar)."""
    if not txt:
        return ""
    limpio = txt.strip()
    lineas = limpio.splitlines()
    cortado = False
    if len(lineas) > max_lineas:
        limpio = "\n".join(lineas[:max_lineas])
        cortado = True
    if len(limpio) > max_chars:
        corte = limpio[:max_chars]
        espacio = corte.rfind(" ")
        if espacio > max_chars * 0.6:
            corte = corte[:espacio]
        limpio = corte.rstrip(" ,;:·-")
        cortado = True
    return limpio + "…" if cortado else limpio


def _css_str(txt: str) -> str:
    """Texto seguro dentro de un string CSS entre comillas dobles."""
    return txt.replace("\\", "\\\\").replace('"', '\\"').replace("\n", " ")


def _generado_txt(r: CotizacionInternaPdfRequest) -> str:
    """'8 sep 2026 10:12' en hora Cancún: `generado_cancun` (pared, del API)
    manda; si no, `generado` (ISO) convertido; si no, ahora."""
    if r.generado_cancun and len(r.generado_cancun) >= 10:
        dia = _fecha_dia(r.generado_cancun[:10])
        hora = r.generado_cancun[11:16].strip()
        if dia:
            return f"{dia} {hora}".strip()
    dt: datetime | None = None
    if r.generado:
        try:
            dt = datetime.fromisoformat(r.generado.replace("Z", "+00:00"))
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=UTC)
        except ValueError:
            dt = None
    if dt is None:
        dt = datetime.now(UTC)
    dt = dt.astimezone(_CANCUN)
    return f"{_fecha_dia(dt.strftime('%Y-%m-%d'))} {dt.strftime('%H:%M')}"


def _pie_texto(r: CotizacionInternaPdfRequest) -> str:
    pie = f"Documento interno · generado {_generado_txt(r)} (hora Cancún)"
    if r.generado_por:
        pie += f" · {r.generado_por.strip()}"
    return pie


# ===== Piezas HTML =====


def _kv(label: str, valor_html: str, clase: str = "") -> str:
    """Fila clave/valor. `valor_html` YA viene escapado por quien la arma."""
    cls = f' class="{clase}"' if clase else ""
    return f'<tr{cls}><td class="k">{escape(label)}</td><td class="v">{valor_html}</td></tr>'


def _op(txt: str) -> str:
    """Operación/aclaración en gris junto a un concepto."""
    return f' <span class="op">{escape(txt)}</span>' if txt else ""


def _tag(txt: str, clase: str = "") -> str:
    cls = f"tag {clase}".strip()
    return f'<span class="{cls}">{escape(txt)}</span>'


def _header_html(r: CotizacionInternaPdfRequest) -> str:
    logo = _logo_data_uri("logo-vuelatour-blanco.png")
    logo_html = f'<img class="logo" src="{logo}" alt=""/>' if logo else ""
    folio = f"Folio #{escape(r.folio)}" if r.folio else "Folio s/n"
    if r.version is not None:
        folio += f" · v{r.version}"
    estado = r.estado_label or r.estado or ""
    if r.tipo:
        estado = f"{estado} · {r.tipo}".strip(" ·")
    return f"""
  <div class="header">
    <div class="izq">{logo_html}<div>
      <div class="titulo">Cotización interna</div>
      <div class="sub">VuelaTour — Aero Charter Cancún</div></div></div>
    <div class="folio">{folio}<span class="sub">{escape(estado)}</span></div>
  </div>
  <div class="banda">{escape(BANDA_INTERNA)}</div>"""


def _ficha_html(r: CotizacionInternaPdfRequest) -> str:
    """Ficha en dos columnas: cliente/condiciones | avión/tripulación/fechas."""
    izq: list[str] = []
    cliente = escape(r.cliente or "Cliente")
    if r.es_broker:
        cliente += _tag("Broker")
    if r.es_interno:
        cliente += _tag("Cliente interno", "ambar")
    if r.cliente_rfc:
        cliente += _op(f"RFC {r.cliente_rfc}")
    izq.append(_kv("Cliente", cliente))
    izq.append(_kv("Razón social", escape(r.razon_social) if r.razon_social else "—"))
    fecha = _fecha(r.fecha)
    if r.fecha_confirmacion:
        fecha += _op(f"confirmada {_fecha(r.fecha_confirmacion)}")
    izq.append(_kv("Fecha de cotización", fecha))
    tarifa = escape(r.tarifa_tipo_label or r.tarifa_tipo or "—")
    if r.tarifa_hora_usd:
        tarifa += f" · {_money(r.tarifa_hora_usd)}/hr"
    marcas = []
    if r.tarifa_preferencial:
        marcas.append("tarifa preferencial del cliente")
    if r.tarifa_override:
        marcas.append("tarifa manual")
    tarifa += _op(" · ".join(marcas))
    izq.append(_kv("Tarifa", tarifa))
    metodo = escape(r.metodo_cobro_label or r.metodo_cobro or "—")
    if r.comision_billpocket_pct:
        metodo += _op(f"comisión terminal {r.comision_billpocket_pct:g} %")
    izq.append(_kv("Método de cobro", metodo))
    izq.append(_kv("T.C. USD→MXN", _tc(r.tc_usd_mxn)))
    vendedor = escape(r.vendedor) if r.vendedor else "—"
    if r.cotizado_por:
        vendedor += _op(f"cotizó {r.cotizado_por}")
    izq.append(_kv("Vendedor", vendedor))

    der: list[str] = []
    if r.es_externo and r.avion_externo:
        avion = escape(r.avion_externo) + _tag("Externo", "ambar")
        if r.operador_externo:
            avion += _op(f"operador {r.operador_externo}")
    else:
        partes = [p for p in (r.aeronave_cotizada_modelo, r.aeronave_cotizada_matricula) if p]
        avion = escape(" · ".join(partes)) if partes else "—"
    der.append(_kv("Avión cotizado", avion))
    if r.aeronave_operativa:
        der.append(
            _kv(
                "Avión operativo",
                f'<span class="ambar">{escape(r.aeronave_operativa)}</span>'
                + _op("difiere del cotizado"),
            )
        )
    tripulacion = escape(r.piloto) if r.piloto else "—"
    if r.copiloto:
        tripulacion += f" · copiloto {escape(r.copiloto)}"
    if r.apoyos:
        tripulacion += _op("apoyo " + ", ".join(r.apoyos))
    der.append(_kv("Piloto", tripulacion))
    der.append(_kv("Traslado inicial", _fecha(r.fecha_traslado_inicial)))
    der.append(_kv("Traslado final", _fecha(r.fecha_traslado_final)))
    ruta = escape(r.ruta) if r.ruta else "—"
    ruta += f" · {_plural(r.pasajeros, 'pasajero', 'pasajeros')}"
    der.append(_kv("Ruta", ruta))
    tags: list[str] = []
    if r.cotizacion_abierta:
        tags.append(_tag("Cotización abierta", "ambar"))
    if r.itinerario_operativo:
        tags.append(_tag("Itinerario operativo"))
    if r.grupo_folio:
        g = f"Grupo {r.grupo_folio}"
        if r.grupo_posicion and r.grupo_total_aviones:
            g += f" · avión {r.grupo_posicion} de {r.grupo_total_aviones}"
        tags.append(_tag(g))
    if r.combinado_con_folio:
        tags.append(_tag(f"Combinado con #{r.combinado_con_folio}"))
    if tags:
        der.append(_kv("Marcas", "".join(tags)))

    return (
        '<table class="cols"><tr>'
        f'<td class="col"><table class="kv">{"".join(izq)}</table></td>'
        f'<td class="col"><table class="kv">{"".join(der)}</table></td>'
        "</tr></table>"
    )


def _marcas_tramo(t: CotizacionInternaTramoPdf) -> str:
    marcas: list[str] = []
    if t.cancelado:
        m = "Cancelado"
        if t.cancelada_motivo:
            m += f": {t.cancelada_motivo}"
        marcas.append(m)
    if t.es_ferry:
        marcas.append("Ferry")
    if t.solo_operativa:
        marcas.append("Operativo")
    if t.es_sobrevuelo:
        marcas.append("Sobrevuelo")
    if t.requiere_pernocta:
        marcas.append("Pernocta" + (f" {_money(t.pernocta_usd)}" if t.pernocta_usd else ""))
    if t.revision_requerida:
        marcas.append("Revisar")
    return escape(" · ".join(marcas))


def _taco_celda(valor: float | None, origen: str | None) -> str:
    txt = _taco(valor)
    abrev = _ORIGEN_TACO_ABREV.get((origen or "").upper())
    if txt and abrev:
        txt += f' <span class="muted">{abrev}</span>'
    return txt


def _itinerario_html(r: CotizacionInternaPdfRequest) -> str:
    """Tabla compacta por tramo + pie con horas cotizadas / voladas / cobrables."""
    tramos = sorted(r.tramos, key=lambda t: (t.orden, t.orden_real or 0))
    filas: list[str] = []
    for t in tramos:
        tramo = f"{escape(t.origen)} → {escape(t.destino)}"
        if t.piloto and t.piloto != r.piloto:
            tramo += f'<br><span class="muted">{escape(t.piloto)}</span>'
        plan = _fecha_corta(t.fecha_plan)
        real = " – ".join(h for h in (_hora(t.hora_salida), _hora(t.hora_llegada)) if h)
        if real:
            plan += f' <span class="muted">real {real}</span>'
        pax = "" if t.pasajeros is None else str(t.pasajeros)
        h_cot = "" if t.horas_cotizadas is None else f"{t.horas_cotizadas:.2f}"
        cls = ' class="cancelado"' if t.cancelado else ""
        filas.append(
            f"<tr{cls}><td>{t.orden}</td>"
            f'<td class="tramo">{tramo}</td>'
            f'<td class="num">{pax}</td>'
            f"<td>{plan}</td>"
            f"<td>{escape(t.matricula) if t.matricula else '—'}</td>"
            f'<td class="num">{_taco_celda(t.taco_salida, t.taco_salida_origen)}</td>'
            f'<td class="num">{_taco_celda(t.taco_llegada, t.taco_llegada_origen)}</td>'
            f'<td class="num">{_taco(t.horas_taco)}</td>'
            f'<td class="num">{h_cot}</td>'
            f'<td class="muted">{_marcas_tramo(t)}</td></tr>'
        )
    if not filas:
        filas.append('<tr><td colspan="10" class="muted">Sin tramos.</td></tr>')

    # Pie: horas cotizadas (con su composición) vs voladas vs cobrables.
    partes: list[str] = []
    cot = f"Cotizadas {_horas(r.horas_cotizadas_hr)}"
    comp = [
        f"{lbl} {v:g}"
        for lbl, v in (
            ("vuelo", r.vuelo_hr),
            ("calzos", r.calzos_hr),
            ("sobrevuelo", r.sobrevuelo_hr),
        )
        if v
    ]
    if len(comp) > 1:
        cot += f" ({' + '.join(comp)})"
    partes.append(cot)
    if r.horas_voladas_hr is None:
        partes.append("Voladas (tacos) —")
    else:
        voladas = f"Voladas (tacos) {_horas(r.horas_voladas_hr, 1)}"
        if r.delta_horas_hr is not None:
            # voladas − cotizadas: lo calcula el API; aquí solo se escribe.
            voladas += f" · Δ {r.delta_horas_hr:+.1f} h"
        partes.append(voladas)
    cobrable = f"Cobrables {_horas(r.tiempo_cobrable_hr)}"
    notas = []
    if r.hora_minima_aplicada:
        notas.append("hora mínima aplicada")
    if r.cobrable_override:
        notas.append("cobrable manual")
    if notas:
        cobrable += f" ({', '.join(notas)})"
    partes.append(cobrable)
    pie = f'<tr><td colspan="10">{escape(" · ".join(partes))}</td></tr>'

    return f"""
  <h2>Itinerario</h2>
  <table class="grid itin"><thead><tr>
    <th>#</th><th>Tramo</th><th class="num">Pax</th><th>Plan</th><th>Matrícula</th>
    <th class="num">Taco sal.</th><th class="num">Taco lleg.</th><th class="num">H taco</th>
    <th class="num">H cot.</th><th>Marcas</th>
  </tr></thead><tbody>{"".join(filas)}</tbody><tfoot>{pie}</tfoot></table>"""


def _concepto_operacion(
    ln: CotizacionInternaLineaPdf, r: CotizacionInternaPdfRequest
) -> tuple[str, str]:
    """(concepto, operación en gris) de una línea canónica. La operación se
    arma SOLO con cantidad/unitario que manda el API; el monto no se toca."""
    clave = (ln.clave or "").upper()
    concepto = (ln.concepto or ln.clave or "Concepto").strip()
    if ln.exento:
        pax = f"{ln.cantidad:g} pax · " if ln.cantidad else ""
        return concepto, f"{pax}exento"
    if clave == "COMISION_VENDEDOR":
        nombre = (r.comision_vendedor_nombre or "").strip()
        if nombre and nombre.lower() not in concepto.lower():
            concepto += f" · {nombre}"
        partes: list[str] = []
        modo = (r.comision_vendedor_modo or "").upper()
        if modo == "POR_HORA" and ln.cantidad is not None and ln.unitario is not None:
            partes.append(f"{ln.cantidad:.2f} h × {_money(ln.unitario)}/hr")
        elif modo == "FIJA":
            partes.append("fija")
        if r.pago_vendedor_usd is not None:
            con_iva = " c/IVA" if r.iva_comision_vendedor_usd else ""
            partes.append(f"pago al vendedor{con_iva} {_money(r.pago_vendedor_usd)}")
        return concepto, " · ".join(partes)
    if clave == "AJUSTE":
        if r.total_pactado_usd is not None:
            return concepto, f"a precio pactado {_money(r.total_pactado_usd)}"
        if ln.monto_usd < 0:
            return concepto, "descuento"
        if r.redondeo_auto_usd:
            return concepto, "redondeo automático"
        return concepto, ""
    if ln.cantidad is None or ln.unitario is None or "×" in concepto:
        return concepto, ""
    unit = _money(ln.unitario)
    if clave == "TIEMPO_VUELO":
        op = f"{ln.cantidad:.2f} h × {unit}/hr"
        if r.hora_minima_aplicada:
            op += " · hora mínima"
        if r.cobrable_override:
            op += " · cobrable manual"
        return concepto, op
    if clave == "TUAS":
        concepto = _RE_PAX_SUFIJO.sub("", concepto) or concepto
        op = f"{ln.cantidad:g} pax × {unit}"
    else:
        op = f"{ln.cantidad:g} × {unit}"
    if (ln.moneda or "USD").upper() == "MXN":
        op += " MXN"
        if ln.monto_nativo is not None:
            op += f" = ${ln.monto_nativo:,.2f} MXN"
        if ln.tc_aplicado:
            op += f" · T.C. {ln.tc_aplicado:g}"
    if clave == "EXTRA" and ln.aplica_iva is False:
        op += " · sin IVA"
    return concepto, op


def _lineas_fallback(r: CotizacionInternaPdfRequest) -> list[CotizacionInternaLineaPdf]:
    """Sin líneas canónicas (API viejo / sin snapshot): una por componente ≠ 0
    con los escalares espejo del vuelo."""
    lineas: list[CotizacionInternaLineaPdf] = []
    if r.subtotal_vuelo_usd:
        lineas.append(
            CotizacionInternaLineaPdf(
                clave="TIEMPO_VUELO",
                concepto="Servicio aéreo",
                monto_usd=r.subtotal_vuelo_usd,
                cantidad=r.tiempo_cobrable_hr,
                unitario=r.tarifa_hora_usd,
            )
        )
    if r.tuas_usd:
        lineas.append(
            CotizacionInternaLineaPdf(clave="TUAS", concepto="TUAS", monto_usd=r.tuas_usd)
        )
    if r.extras_total_usd:
        lineas.append(
            CotizacionInternaLineaPdf(
                clave="EXTRA", concepto="Extras", monto_usd=r.extras_total_usd
            )
        )
    if r.comision_vendedor_usd:
        lineas.append(
            CotizacionInternaLineaPdf(
                clave="COMISION_VENDEDOR",
                concepto="Comisión del vendedor",
                monto_usd=r.comision_vendedor_usd,
            )
        )
    if r.ajuste_final_usd:
        lineas.append(
            CotizacionInternaLineaPdf(
                clave="AJUSTE",
                concepto="Descuento" if r.ajuste_final_usd < 0 else "Ajuste",
                monto_usd=r.ajuste_final_usd,
            )
        )
    if r.viaticos_pernocta_usd:
        lineas.append(
            CotizacionInternaLineaPdf(
                clave="PERNOCTA",
                concepto="Viáticos por pernocta",
                monto_usd=r.viaticos_pernocta_usd,
            )
        )
    return lineas


def _desglose_html(r: CotizacionInternaPdfRequest) -> str:
    """Desglose COMPLETO interno: líneas canónicas en su orden (IVA aparte,
    en los totales), subtotal sin IVA, IVA con su base, total USD y MXN."""
    lineas = list(r.lineas) or _lineas_fallback(r)
    filas: list[str] = []
    hay_exento = any(ln.exento for ln in lineas)
    ultimo_tuas = max(
        (i for i, ln in enumerate(lineas) if (ln.clave or "").upper() == "TUAS"), default=-1
    )
    for i, ln in enumerate(lineas):
        clave = (ln.clave or "").upper()
        if clave == "IVA":
            continue
        concepto, op = _concepto_operacion(ln, r)
        monto = "—" if ln.exento else _monto(ln.monto_usd)
        cls = ' class="exento"' if ln.exento else ""
        filas.append(
            f'<tr{cls}><td class="lbl">{escape(concepto)}{_op(op)}</td>'
            f'<td class="val">{monto}</td></tr>'
        )
        if i == ultimo_tuas and r.tuas_exentos and not hay_exento:
            filas.append(
                '<tr class="exento"><td class="lbl">TUA exento'
                f'{_op(", ".join(r.tuas_exentos))}</td>'
                '<td class="val">—</td></tr>'
            )
    if not filas:
        filas.append(
            '<tr><td class="lbl muted" colspan="2">Sin desglose (cotización sin precio).</td></tr>'
        )

    # El subtotal (total − IVA) lo manda el API; sin dato se deja en blanco.
    subtotal = "—" if r.subtotal_usd is None else _monto(r.subtotal_usd)
    filas.append(
        '<tr class="sub-row"><td class="lbl">Subtotal (sin IVA)</td>'
        f'<td class="val">{subtotal}</td></tr>'
    )
    iva_op = f"sobre {_money(r.iva_base_usd)}" if r.iva_base_usd is not None else ""
    if r.iva_nota:
        iva_op = f"{iva_op} · {r.iva_nota}".strip(" ·")
    filas.append(
        f'<tr><td class="lbl">IVA {r.iva_pct:g} %{_op(iva_op)}</td>'
        f'<td class="val">{_money(r.iva_usd)}</td></tr>'
    )
    filas.append(
        f'<tr class="total-row"><td>Total USD</td><td class="val">{_monto(r.total_usd)}</td></tr>'
    )
    if r.total_mxn is not None:
        mxn_op = f"T.C. {r.tc_usd_mxn:g}" if r.tc_usd_mxn else ""
        if r.mxn_nativos:
            mxn_op = f"{mxn_op} · incluye ${r.mxn_nativos:,.2f} MXN nativos".strip(" ·")
        filas.append(
            f'<tr class="mxn-row"><td>Total MXN{_op(mxn_op)}</td>'
            f'<td class="val">{_monto(r.total_mxn)} MXN</td></tr>'
        )
    motor = []
    if r.version_motor:
        motor.append(f"motor {r.version_motor}")
    if r.calculado_at:
        motor.append(f"calculado {_fecha(r.calculado_at)}")
    if motor:
        filas.append(f'<tr><td colspan="2" class="muted">{escape(" · ".join(motor))}</td></tr>')
    return (
        '<h2>Desglose interno</h2>'
        f'<table class="totales"><tbody>{"".join(filas)}</tbody></table>'
    )


def _particion_html(r: CotizacionInternaPdfRequest) -> str:
    """Partición del ingreso (particionIngresoVuelo): venta del avión vs
    ingreso VuelaTour, por aeronave si es multi-avión, pago al vendedor y neto."""
    filas: list[str] = []
    if r.venta_avion_usd is None and r.otros_ingresos_vuelatour_usd is None:
        filas.append(
            '<tr><td colspan="2" class="muted">Sin partición (cotización sin precio).</td></tr>'
        )
    else:
        venta = _monto(r.venta_avion_usd or 0)
        if r.iva_avion_usd is not None:
            venta += _op(f"incluye IVA {_money(r.iva_avion_usd)}")
        filas.append(_kv("Venta del avión", venta))
        otros = _monto(r.otros_ingresos_vuelatour_usd or 0)
        otros += _op(
            "TUAS/extras/pernocta/comisión + su IVA"
            + (f" {_money(r.iva_vuelatour_usd)}" if r.iva_vuelatour_usd is not None else "")
        )
        filas.append(_kv("Ingreso VuelaTour", otros))
        if len(r.participacion_aviones) > 1:
            for p in r.participacion_aviones:
                tramos = _tramos_txt(p.tramos)
                txt = _pct(p.factor) + (f" ({tramos})" if tramos else "")
                if p.venta_usd is not None:
                    txt += f" = {_money(p.venta_usd)}"
                filas.append(_kv(f"· {p.matricula or '—'}", escape(txt)))
    if r.pago_vendedor_usd is not None:
        detalle = f"comisión {_money(r.comision_vendedor_usd)}"
        if r.iva_comision_vendedor_usd:
            detalle += f" + IVA {_money(r.iva_comision_vendedor_usd)}"
        if r.comision_vendedor_nombre:
            detalle += f" · {r.comision_vendedor_nombre}"
        filas.append(
            _kv("Pago al vendedor", f"&minus;{_money(r.pago_vendedor_usd)}" + _op(detalle))
        )
    if r.neto_vuelatour_usd is not None:
        filas.append(
            _kv(
                "Neto VuelaTour",
                f"<b>{_monto(r.neto_vuelatour_usd)}</b>" + _op("total − pago al vendedor"),
            )
        )
    avisos: list[str] = []
    if r.particion_inconsistente:
        avisos.append('<span class="rojo">Partición inconsistente con el desglose: revisar.</span>')
    if r.particion_fuente == "columnas":
        avisos.append("Partición de columnas espejo (sin snapshot).")
    if avisos:
        filas.append(f'<tr><td colspan="2" class="muted">{" ".join(avisos)}</td></tr>')
    return f'<h2>Partición del ingreso</h2><table class="kv">{"".join(filas)}</table>'


def _cobro_fila(c: CotizacionInternaCobroPdf, con_usd: bool) -> str:
    metodo = escape(c.metodo_label or c.metodo or "—")
    if c.es_reembolso or c.monto < 0:
        metodo += '<br><span class="muted">Reembolso</span>'
    moneda = (c.moneda or "USD").upper()
    sufijo = "" if moneda == "USD" else f' <span class="muted">{escape(moneda)}</span>'
    bruto = _monto(c.monto) + sufijo
    if c.comision_pct is not None or c.comision_monto:
        com = _pct_banco(c.comision_pct) if c.comision_pct is not None else ""
        if c.comision_monto:
            com = f"{com} = {_money(c.comision_monto)}".strip(" =")
    else:
        com = "—"
    neto = (_monto(c.neto) + sufijo) if c.neto is not None else "—"
    ref_partes = [p for p in (c.referencia, c.cuenta_destino) if p]
    ref = escape(" · ".join(ref_partes)) if ref_partes else "—"
    sub: list[str] = []
    if c.sobre_grupo_folio:
        s = f"Sobre {c.sobre_grupo_folio}"
        if c.sobre_grupo_monto_total is not None:
            s += f" {_money(c.sobre_grupo_monto_total)} {c.sobre_grupo_moneda or ''}".rstrip()
        if c.grupo_factor is not None:
            s += f" × {_pct(c.grupo_factor)}"
        sub.append(s)
    if c.notas:
        sub.append(_truncar(c.notas, 80, 1))
    if sub:
        ref += f'<br><span class="muted">{escape(" · ".join(sub))}</span>'
    if c.conciliado is None:
        conc = "—"
    else:
        conc = '<span class="verde">Sí</span>' if c.conciliado else "No"
    usd = ""
    if con_usd:
        # cobrosEnUsd: un cobro USD vale su bruto; uno MXN sin ningún TC queda
        # FUERA de la suma y se marca (nunca desaparece en silencio).
        monto_usd = c.monto_usd
        if monto_usd is None and moneda == "USD":
            monto_usd = c.monto
        usd_txt = (
            _monto(monto_usd) if monto_usd is not None else '<span class="rojo">sin T.C.</span>'
        )
        usd = f'<td class="num">{usd_txt}</td>'
    return (
        f"<tr><td>{_fecha_corta(c.fecha, con_anio=True)}</td><td>{metodo}</td>"
        f'<td class="num">{bruto}</td><td class="num">{_tc(c.tc)}</td>'
        f'<td class="num">{com}</td><td class="num">{neto}</td>{usd}'
        f"<td>{ref}</td><td>{conc}</td></tr>"
    )


def _cobros_html(r: CotizacionInternaPdfRequest) -> str:
    """Cobros del vuelo (cobro_vuelo, partes de sobre incluidas) y el resumen:
    cobrado (cobrosEnUsd), comisiones bancarias, neto, saldo y semáforo."""
    con_usd = any((c.moneda or "USD").upper() != "USD" for c in r.cobros)
    ncols = 9 if con_usd else 8
    if r.cobros:
        filas = "".join(_cobro_fila(c, con_usd) for c in r.cobros)
    else:
        filas = f'<tr><td colspan="{ncols}" class="muted">Sin cobros registrados.</td></tr>'
    th_usd = '<th class="num">Equiv. USD</th>' if con_usd else ""

    resumen = [f"Cobrado {_monto(r.total_cobrado_usd)} USD"]
    if r.comision_banco_usd:
        resumen.append(f"comisiones banco &minus;{_money(r.comision_banco_usd)}")
        if r.total_cobrado_neto_usd is not None:
            resumen.append(f"neto {_monto(r.total_cobrado_neto_usd)}")
    # Saldo = total − cobrado, calculado por el API (saldo_usd); sin dato «—».
    if r.saldo_usd is None:
        resumen.append("Saldo —")
    else:
        saldo_txt = f"Saldo {_monto(r.saldo_usd)}"
        if r.saldo_usd < -0.005:
            saldo_txt += " (sobrecobro)"
        resumen.append(saldo_txt)
    sem_cls = _SEMAFORO_CLASES.get((r.semaforo_cobro or "").lower(), "sem-gris")
    sem_lbl = r.semaforo_cobro_label or (r.semaforo_cobro_key or "—").replace("_", " ").capitalize()
    resumen.append(f'<span class="sem {sem_cls}"></span>{escape(sem_lbl)}')
    pie = f'<tr><td colspan="{ncols}">{" · ".join(resumen)}</td></tr>'
    if r.cobros_sin_tc_count:
        pie += (
            f'<tr><td colspan="{ncols}" class="rojo aviso">'
            f"OJO: {_plural(r.cobros_sin_tc_count, 'cobro', 'cobros')} en MXN por "
            f"${r.cobros_sin_tc_mxn:,.2f} SIN tipo de cambio: fuera de la suma.</td></tr>"
        )
    return f"""
  <div class="bloque"><h2>Cobros</h2>
  <table class="grid"><thead><tr>
    <th>Fecha</th><th>Método</th><th class="num">Bruto</th><th class="num">T.C.</th>
    <th class="num">Comisión banco</th><th class="num">Neto</th>{th_usd}
    <th>Referencia / cuenta</th><th>Conc.</th>
  </tr></thead><tbody>{filas}</tbody><tfoot>{pie}</tfoot></table></div>"""


def _gastos_html(r: CotizacionInternaPdfRequest) -> str:
    """Gastos del vuelo por categoría (USD) + utilidad bruta de REFERENCIA."""
    filas: list[str] = []
    cats = sorted(r.gastos_por_categoria, key=lambda g: -g.total_usd)
    if not cats and r.gastos_total_usd is None and r.costo_externo_usd is None:
        filas.append('<tr><td colspan="2" class="muted">Sin gastos ligados al vuelo.</td></tr>')
    else:
        # Una fila por categoría tal cual llega (el API ya agregó y convirtió
        # a USD; el enum tiene 19 categorías, un vuelo real usa ~5). Nada se
        # re-suma aquí.
        for g in cats:
            lbl = g.etiqueta or g.categoria or "Gasto"
            if g.n:
                lbl += f" ({g.n})"
            filas.append(_kv(lbl, _monto(g.total_usd)))
        if r.costo_externo_usd is not None:
            filas.append(_kv("Costo del operador externo", _monto(r.costo_externo_usd)))
        if r.gastos_total_usd is not None:
            filas.append(_kv("Total gastos", f"<b>{_monto(r.gastos_total_usd)}</b>", "sub-row"))
    if r.gastos_sin_tc_count:
        filas.append(
            '<tr><td colspan="2" class="rojo aviso">OJO: '
            f"{_plural(r.gastos_sin_tc_count, 'gasto', 'gastos')} en MXN por "
            f"${r.gastos_sin_tc_mxn:,.2f} sin tipo de cambio: no entran al total.</td></tr>"
        )
    if r.utilidad_bruta_usd is not None:
        if r.utilidad_base == "total":
            base_lbl, base_val = "total", r.total_usd
        else:
            base_lbl, base_val = "cobrado", r.total_cobrado_usd
        op = (
            f"{base_lbl} {_money(base_val)} − gastos {_money(r.gastos_total_usd or 0)}"
            " · no es el reparto"
        )
        color = "verde" if r.utilidad_bruta_usd >= 0 else "rojo"
        filas.append(
            _kv(
                "Utilidad bruta (referencia)",
                f'<b class="{color}">{_monto(r.utilidad_bruta_usd)}</b>' + _op(op),
            )
        )
    return f'<h2>Gastos del vuelo</h2><table class="kv">{"".join(filas)}</table>'


def _facturacion_html(r: CotizacionInternaPdfRequest) -> str:
    if not (r.cfdi_estatus or r.facturas or r.facturado):
        return ""
    filas: list[str] = []
    estatus = r.cfdi_estatus or ("Facturado" if r.facturado else "—")
    txt = escape(estatus)
    if r.cfdi_folio:
        txt += f" · {escape(r.cfdi_folio)}"
    filas.append(_kv("CFDI", txt))
    for f in r.facturas[:_FACTURAS_MAX]:
        ident = "-".join(p for p in (f.serie, f.folio) if p) or (f.uuid_fiscal or "")[:8] or "s/n"
        partes = [ident]
        if f.estado:
            partes.append(f.estado)
        if f.total is not None:
            partes.append(f"{_money(f.total)} {f.moneda or ''}".rstrip())
        if f.fecha_timbrado:
            partes.append(_fecha(f.fecha_timbrado))
        if f.facturado_a_nombre:
            partes.append(f.facturado_a_nombre)
        linea = escape(" · ".join(partes))
        if f.cancelada:
            linea = f'<s>{linea}</s> <span class="rojo">cancelada</span>'
        filas.append(f'<tr><td colspan="2" class="muted">{linea}</td></tr>')
    if len(r.facturas) > _FACTURAS_MAX:
        filas.append(
            f'<tr><td colspan="2" class="muted">… {len(r.facturas) - _FACTURAS_MAX} más</td></tr>'
        )
    return f'<h2>Facturación</h2><table class="kv">{"".join(filas)}</table>'


def _notas_html(r: CotizacionInternaPdfRequest) -> str:
    internas = _truncar(r.notas_internas)
    cliente = _truncar(r.notas_cliente)
    if not (internas or cliente):
        return ""
    return (
        '<table class="cols bloque"><tr>'
        '<td class="col"><h2>Notas internas</h2>'
        f'<div class="notas-txt">{escape(internas) or "—"}</div></td>'
        '<td class="col"><h2>Notas del cliente</h2>'
        f'<div class="notas-txt">{escape(cliente) or "—"}</div></td>'
        "</tr></table>"
    )


def _estilos_interno(pie: str) -> str:
    """CSS propio del documento interno (denso, una hoja carta). Branding
    compartido con el PDF del cliente vía `_BRAND`/`_NAVY`."""
    fuente = "'Helvetica Neue', Arial, sans-serif"
    return f"""
  @page {{
    size: Letter;
    margin: 10mm 12mm 11mm;
    @bottom-left {{ content: "{_css_str(pie)}"; font-size: 7pt; color: #9ca3af;
                    font-family: {fuente}; }}
    @bottom-right {{ content: "Página " counter(page) " de " counter(pages);
                     font-size: 7pt; color: #9ca3af; font-family: {fuente}; }}
  }}
  * {{ font-family: {fuente}; color: #1d1d1d; box-sizing: border-box; }}
  body {{ margin: 0; font-size: 8.5pt; line-height: 1.2; }}
  .header {{ background: {_NAVY}; color: #fff; padding: 5px 9px; border-radius: 5px;
             display: flex; align-items: center; justify-content: space-between; }}
  .header .izq {{ display: flex; align-items: center; }}
  .header .logo {{ height: 16px; margin-right: 9px; }}
  .header .titulo {{ font-size: 10.5pt; font-weight: 800; color: #fff; line-height: 1.1; }}
  .header .sub {{ font-size: 7pt; color: #9fb3c8; display: block; }}
  .header .folio {{ text-align: right; color: #fff; font-size: 9pt; font-weight: 700; }}
  .banda {{ margin: 3px 0 4px; padding: 1px 8px; border-left: 3px solid {_BRAND};
            background: #fef2f2; color: {_BRAND}; font-size: 7.5pt; font-weight: 800;
            letter-spacing: 1px; text-transform: uppercase; }}
  h2 {{ font-size: 8pt; text-transform: uppercase; letter-spacing: .8px; color: {_BRAND};
        border-bottom: 1px solid #e5e7eb; margin: 5px 0 1px; padding-bottom: 1px; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 8.5pt; }}
  table.kv td {{ padding: 1px 3px; vertical-align: top; }}
  table.kv td.k {{ color: #6b7280; width: 32%; white-space: nowrap; }}
  table.grid {{ font-size: 8pt; }}
  table.grid th, table.grid td {{ border: 1px solid #e5e7eb; padding: 1px 4px;
                                  text-align: left; vertical-align: top; }}
  table.grid th {{ background: #f3f4f6; font-size: 7pt; text-transform: uppercase;
                   color: #374151; letter-spacing: .3px; }}
  table.grid tfoot td {{ font-weight: 700; background: #f7f7f8; color: {_NAVY}; }}
  table.grid tfoot td.aviso {{ font-weight: 400; background: #fff; }}
  thead {{ display: table-header-group; }}
  .num {{ text-align: right !important; white-space: nowrap; }}
  .muted {{ color: #6b7280; font-size: 7.5pt; }}
  .op {{ color: #6b7280; font-size: 7.5pt; font-weight: 400; }}
  .ambar {{ color: #b45309; }}  .rojo {{ color: {_BRAND}; }}  .verde {{ color: #15803d; }}
  .tag {{ display: inline-block; border: 1px solid #d1d5db; border-radius: 3px; padding: 0 3px;
          font-size: 7pt; color: #374151; margin-left: 3px; }}
  .tag.ambar {{ border-color: #f59e0b; }}
  .totales td {{ padding: 1px 3px; vertical-align: top; }}
  .totales .val {{ text-align: right; font-weight: 600; white-space: nowrap; }}
  .totales .exento td {{ color: #6b7280; }}
  .sub-row td {{ border-top: 1px solid #d1d5db; font-weight: 700; color: {_NAVY}; }}
  .total-row td {{ border-top: 2px solid {_NAVY}; font-size: 10pt; font-weight: 800;
                   color: {_BRAND}; padding-top: 3px; }}
  .mxn-row td {{ font-weight: 700; color: {_NAVY}; }}
  .cols {{ width: 100%; border-collapse: separate; border-spacing: 0; }}
  .cols td.col {{ vertical-align: top; width: 50%; padding: 0; }}
  .cols td.col:first-child {{ padding-right: 9px; }}
  .cols td.col:last-child {{ padding-left: 9px; }}
  .sem {{ display: inline-block; width: 7px; height: 7px; border-radius: 50%;
          vertical-align: middle; margin-right: 3px; }}
  .sem-verde {{ background: #16a34a; }}  .sem-amarillo {{ background: #f59e0b; }}
  .sem-rojo {{ background: {_BRAND}; }}  .sem-gris {{ background: #9ca3af; }}
  .bloque {{ page-break-inside: avoid; }}
  tr.cancelado td {{ color: #9ca3af; }}
  tr.cancelado td.tramo {{ text-decoration: line-through; }}
  .notas-txt {{ font-size: 8pt; color: #374151; white-space: pre-wrap; }}"""


def _build_html(r: CotizacionInternaPdfRequest) -> str:
    pie = _pie_texto(r)
    izquierda = _desglose_html(r) + _facturacion_html(r)
    derecha = _particion_html(r) + _gastos_html(r)
    cuerpo = (
        '<table class="cols bloque"><tr>'
        f'<td class="col">{izquierda}</td>'
        f'<td class="col">{derecha}</td>'
        "</tr></table>"
    )
    return f"""<!doctype html>
<html><head><meta charset="utf-8"><title>Cotización interna #{escape(r.folio)}</title><style>
{_estilos_interno(pie)}
</style></head><body>
{_header_html(r)}
  {_ficha_html(r)}
  {_itinerario_html(r)}
  {cuerpo}
  {_cobros_html(r)}
  {_notas_html(r)}
</body></html>"""


def render_cotizacion_interna_pdf(req: CotizacionInternaPdfRequest) -> bytes:
    from weasyprint import HTML  # import perezoso

    return HTML(string=_build_html(req)).write_pdf()
