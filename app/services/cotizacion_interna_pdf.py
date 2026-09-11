"""PDF de la COTIZACIÓN INTERNA v2 (8-sep-2026): UNA hoja carta para la oficina.

Feedback de administración (8-sep, con la foto de su formato de siempre):
SOLO lo de la COTIZACIÓN. La fecha protagonista es el DÍA DEL VUELO (la de
cotización/confirmación va en pequeño); los tramos se desglosan como en su
hoja — RUTA («Cancun-Merida») · FECHA («26-jun») · DISTANCIA MILLAS ·
TIEMPO VUELO («01:18», incluye calzos) · COSTO POR HORA VUELO · TOTAL POR
TRAMO, con fila TOTAL en USD —; TUAS solo las que SE COBRARON; desglose
canónico con comisión del vendedor, ajuste/redondeo, IVA y total USD/MXN;
cobros compactos con su comisión bancaria; notas internas. NADA de operación
(tacómetros, horas voladas, piloto del tramo, traslados) ni partición /
gastos / utilidad / CFDI: eso vive en el reporte del vuelo. ÚNICA excepción
(11-sep-2026, pedido de la oficina): debajo del «Avión cotizado» se pinta
«Avión utilizado: …» cuando el API manda el campo NUEVO `aeronave_utilizada`
— saber si salió el avión que se cotizó es dato de la cotización, no de la
operación. El legado `aeronave_operativa` de la v1 sigue sin pintarse.

NUNCA se manda al cliente: lleva banda «COTIZACIÓN INTERNA · uso exclusivo
de oficina» y pie «Documento interno · generado … · <usuario>».

Hermano de `cotizacion_pdf.py`: REUTILIZA (importa, no copia) el branding
(#dc2626 / #102a43), el logo, `_money`/`_monto` y los formatos de fecha. NO
importa `_estilos_base` a propósito: esa hoja es la del cliente (13 px,
márgenes de 2 cm, marca de agua y «Gracias por volar con VuelaTour» en el
pie) — aquí el CSS es propio y con @page de 9–12 mm. Tampoco aplica
`_mostrar_matricula`: en el documento interno la matrícula SIEMPRE se ve.

AIRE (11-sep-2026, captura de la oficina: «todo muy junto»): la v2 salió a
8.5 pt con interlineado 1.2 y celdas de 1 px, y la hoja se leía apretada
arriba con el tercio de abajo vacío. Ahora el cuerpo va a 9.5 pt / 1.3, las
tablas NUNCA bajan de 9.5 pt (encabezados y aclaraciones en gris, 8–8.5 pt:
son apoyo), las celdas tienen 3–4 px y cada bloque se separa del anterior
con el margen superior de su `h2`. Misma información, mismos números.

Aquí SOLO se pinta. El API arma el payload desde el desglose canónico v1.3
del snapshot: cada tramo trae su `tiempo_hr` (con calzo), su tarifa y su
`total_usd`; la diferencia contra la línea TIEMPO_VUELO canónica llega como
`tramos_ajuste_usd` con su motivo (hora mínima / sobrevuelo / horas pactadas
/ redondeo) y se pinta como una fila más para que Σ tramos + ajuste ==
«Servicio aéreo». Nunca se recalcula dinero; a lo sumo se formatea
(«01:18», «26-jun», «8.8570 % = $3,236.36») o se re-suma una columna
informativa (millas).

Presupuesto de UNA hoja tras el aire (≈ 738 pt / 984 px útiles, medido con
una render de prueba): base fija (cabecera + banda + resumen + ficha +
desglose|horas + cobros con su thead + notas) ≈ 730 px, y cada FILA de más
≈ 40 px — cuenta cada tramo, cada cobro y las dos filas de la fila de
ajuste. O sea: UNA hoja mientras `filas ≲ 6` (el caso real de la oficina
—2 tramos, 1 cobro— usa el 86 %). Antes cabían ~8 tramos porque la tabla
iba a 8 pt: ese es el precio de que se lea. Con más filas la tabla pasa a
una segunda hoja —el `thead` se repite y ninguna fila se parte a la
mitad—; los bloques llevan `page-break-inside: avoid`.
"""

from datetime import UTC, datetime
from html import escape

from app.schemas.reportes import (
    CotizacionInternaCobroPdf,
    CotizacionInternaLineaPdf,
    CotizacionInternaPdfRequest,
    CotizacionInternaTramoCotizadoPdf,
    CotizacionInternaTuaCobradaPdf,
)
from app.services.cotizacion_grupo_pdf import _RE_PAX_SUFIJO, _monto, _plural
from app.services.cotizacion_pdf import (
    _BRAND,
    _CANCUN,
    _MESES_ES,
    _NAVY,
    _fecha_dia,
    _logo_data_uri,
    _money,
)
from app.services.reporte_vuelo_pdf import _fecha

# Texto de la banda: es lo que distingue este documento del que va al cliente.
BANDA_INTERNA = "Cotización interna · uso exclusivo de oficina · no enviar al cliente"
# Nota al pie de la tabla de tramos (regla del motor: 0.15 h por aterrizaje).
NOTA_TRAMOS = "Tiempo de vuelo en hh:mm e incluye calzos"

# Truncado elegante de notas: prioridad a la cotización; las notas van al final.
_NOTAS_MAX_CHARS = 280
_NOTAS_MAX_LINEAS = 5

_SEMAFORO_CLASES = {"verde": "sem-verde", "amarillo": "sem-amarillo", "rojo": "sem-rojo"}


# ===== Formato =====


def _instante_cancun(s: str) -> datetime | None:
    """Instante ISO → datetime en hora Cancún; None si no se puede parsear."""
    try:
        dt = datetime.fromisoformat(s.replace("Z", "+00:00"))
    except ValueError:
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=UTC)
    return dt.astimezone(_CANCUN)


def _dia_pared(s: str | None) -> str | None:
    """'YYYY-MM-DD' (día de pared) se respeta tal cual — convertirlo movería el
    día; un instante ISO se lleva a su día en Cancún. None si no hay dato."""
    if not s:
        return None
    if len(s) == 10 and "T" not in s:
        return s
    dt = _instante_cancun(s)
    return dt.strftime("%Y-%m-%d") if dt else None


def _fecha_corta(s: str | None, con_anio: bool = False) -> str:
    """Instante ISO → 'dd/mm HH:MM' en hora Cancún (con año si se pide).
    Un día de pared ('2026-09-03') NO se convierte (movería el día)."""
    if not s:
        return "—"
    if len(s) == 10 and "T" not in s:
        try:
            d = datetime.fromisoformat(s)
        except ValueError:
            return escape(s)
        return d.strftime("%d/%m/%Y" if con_anio else "%d/%m")
    dt = _instante_cancun(s)
    if dt is None:
        return escape(s)
    return dt.strftime("%d/%m/%Y %H:%M" if con_anio else "%d/%m %H:%M")


def _dia_mes(s: str | None) -> str:
    """Día del tramo como lo escribe administración: '2026-06-26' → '26-jun'."""
    dia = _dia_pared(s)
    if not dia:
        return "—"
    try:
        d = datetime.strptime(dia[:10], "%Y-%m-%d")
    except ValueError:
        return escape(dia)
    return f"{d.day}-{_MESES_ES[d.month - 1]}"


def _dia_largo(s: str | None) -> str:
    """'2026-06-26' → '26 jun 2026' (fecha protagonista); '' si no hay."""
    return _fecha_dia(_dia_pared(s))


def _hhmm(h: float | None) -> str:
    """Horas decimales → 'hh:mm' (1.3 → '01:18', 0.4 → '00:24'). Misma regla
    que `horasAHhmm` del API; solo se usa si el API no mandó `tiempo_hhmm`."""
    if h is None:
        return "—"
    m = max(0, round(h * 60))
    return f"{m // 60:02d}:{m % 60:02d}"


def _millas(v: float | None) -> str:
    """Millas náuticas como en la hoja de administración: '157', '157.3'."""
    if v is None:
        return "—"
    txt = f"{v:,.1f}"
    return txt[:-2] if txt.endswith(".0") else txt


def _horas(v: float | None, dec: int = 2) -> str:
    return "—" if v is None else f"{v:.{dec}f} h"


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
    dt = _instante_cancun(r.generado) if r.generado else None
    if dt is None:
        dt = datetime.now(UTC).astimezone(_CANCUN)
    return f"{_fecha_dia(dt.strftime('%Y-%m-%d'))} {dt.strftime('%H:%M')}"


def _pie_texto(r: CotizacionInternaPdfRequest) -> str:
    pie = f"Documento interno · generado {_generado_txt(r)} (hora Cancún)"
    if r.generado_por:
        pie += f" · {r.generado_por.strip()}"
    return pie


def _servicio_aereo_usd(r: CotizacionInternaPdfRequest) -> float:
    """Línea TIEMPO_VUELO canónica (ancla de la tabla de tramos); sin líneas,
    el escalar espejo `subtotal_vuelo_usd`. Nunca se suma aquí."""
    for ln in r.lineas:
        if (ln.clave or "").upper() == "TIEMPO_VUELO":
            return ln.monto_usd
    return r.subtotal_vuelo_usd


def _hay_ajuste(r: CotizacionInternaPdfRequest) -> bool:
    return abs(r.tramos_ajuste_usd) >= 0.005


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


def _avion_html(r: CotizacionInternaPdfRequest) -> str:
    """«Piper Seneca V · N4142R» (matrícula SIEMPRE visible) o el avión externo."""
    if r.es_externo and r.avion_externo:
        avion = escape(r.avion_externo) + _tag("Externo", "ambar")
        if r.operador_externo:
            avion += _op(f"operador {r.operador_externo}")
        return avion
    partes = [p for p in (r.aeronave_cotizada_modelo, r.aeronave_cotizada_matricula) if p]
    return escape(" · ".join(partes)) if partes else "—"


def _avion_utilizado_txt(r: CotizacionInternaPdfRequest) -> str:
    """«N4142R · Piper Seneca V» del avión REALMENTE utilizado (campo nuevo
    `aeronave_utilizada`, 11-sep-2026). TOLERANTE a propósito: acepta el
    texto ya armado por el API o un objeto {matricula, modelo}; sin dato —o
    con otra forma— devuelve "" y la segunda línea no se pinta."""
    v = r.aeronave_utilizada
    if isinstance(v, str):
        return v.strip()
    if isinstance(v, dict):
        partes = [
            str(v.get(k)).strip()
            for k in ("matricula", "modelo")
            if v.get(k) not in (None, "")
        ]
        return " · ".join(p for p in partes if p)
    return ""


def _resumen_html(r: CotizacionInternaPdfRequest) -> str:
    """Franja de resumen: FECHA DEL VUELO (protagonista, grande), avión
    COTIZADO (con la matrícula, siempre) y ruta cotizada con pasajeros. Si
    el API manda `aeronave_utilizada`, debajo del cotizado va la segunda
    línea «Avión utilizado: …» — la oficina necesita ver los dos cuando no
    salió el avión que se cotizó; si el API avisa que DIFIEREN, esa línea va
    en ámbar con la marca «Distinto al cotizado»."""
    fecha = _dia_largo(r.fecha_vuelo) or "Por definir"
    fin = _dia_largo(r.fecha_vuelo_fin)
    fin_html = f'<div class="sub">al {escape(fin)}</div>' if fin else ""
    utilizada = _avion_utilizado_txt(r)
    # ⚠ Se cotizó con un avión y vuela OTRO: lo decide el API comparando por
    # ID (`aeronave_cotizada_vs_utilizada_difiere`) — el texto no basta,
    # dos aviones pueden compartir modelo. La oficina tiene que verlo de un
    # vistazo, no deducirlo comparando dos matrículas en gris.
    difiere = r.aeronave_cotizada_vs_utilizada_difiere is True
    utilizada_html = (
        f'<div class="sub{" ambar" if difiere else ""}">Avión utilizado: '
        f"{escape(utilizada)}"
        + (_tag("Distinto al cotizado", "ambar") if difiere else "")
        + "</div>"
        if utilizada
        else ""
    )
    ruta = escape(r.ruta) if r.ruta else "—"
    ruta += _op(f"· {_plural(r.pasajeros, 'pasajero', 'pasajeros')}")
    return f"""
  <table class="resumen"><tr>
    <td class="fecha"><div class="lbl">Fecha del vuelo</div>
      <div class="big">{escape(fecha)}</div>{fin_html}</td>
    <td><div class="lbl">Avión cotizado</div><div class="med">{_avion_html(r)}</div>
      {utilizada_html}</td>
    <td><div class="lbl">Ruta cotizada</div><div class="med">{ruta}</div></td>
  </tr></table>"""


def _ficha_html(r: CotizacionInternaPdfRequest) -> str:
    """Ficha en dos columnas: cliente/condiciones | vendedor/tripulación/marcas.
    La fecha de cotización va al final y en pequeño (no es protagonista)."""
    izq: list[str] = []
    cliente = escape(r.cliente or "Cliente")
    if r.es_broker:
        cliente += _tag("Broker")
    if r.es_interno:
        cliente += _tag("Cliente interno", "ambar")
    izq.append(_kv("Cliente", cliente))
    razon = escape(r.razon_social) if r.razon_social else "—"
    if r.cliente_rfc:
        razon += _op(f"RFC {r.cliente_rfc}")
    izq.append(_kv("Razón social", razon))
    tarifa = escape(r.tarifa_tipo_label or r.tarifa_tipo or "—")
    if r.tarifa_hora_usd is not None:
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

    der: list[str] = []
    vendedor = escape(r.vendedor) if r.vendedor else "—"
    if r.cotizado_por:
        vendedor += _op(f"cotizó {r.cotizado_por}")
    der.append(_kv("Vendedor", vendedor))
    if r.piloto or r.copiloto:
        tripulacion = escape(r.piloto) if r.piloto else "—"
        if r.copiloto:
            tripulacion += f" · copiloto {escape(r.copiloto)}"
        der.append(_kv("Piloto", tripulacion))
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
    cotizada = _fecha(r.fecha)
    if r.fecha_confirmacion:
        cotizada += f" · confirmada {_fecha(r.fecha_confirmacion)}"
    der.append(_kv("Cotizada", escape(cotizada), "small"))

    return (
        '<table class="cols"><tr>'
        f'<td class="col"><table class="kv">{"".join(izq)}</table></td>'
        f'<td class="col"><table class="kv">{"".join(der)}</table></td>'
        "</tr></table>"
    )


def _tramo_fila(t: CotizacionInternaTramoCotizadoPdf) -> str:
    """Una fila de la tabla de administración: RUTA · FECHA · MILLAS · TIEMPO ·
    COSTO/HR · TOTAL. Debajo de la ruta, en gris, el par IATA y las marcas."""
    ruta = t.ruta or "-".join(
        p for p in (t.origen_nombre or t.origen_iata, t.destino_nombre or t.destino_iata) if p
    )
    detalles: list[str] = []
    if t.consolidado:
        detalles.append("tramos consolidados")
    elif t.origen_iata and t.destino_iata:
        detalles.append(f"{t.origen_iata}–{t.destino_iata}")
    if t.es_ferry:
        detalles.append("ferry")
    if t.pernocta:
        detalles.append("pernocta" + (f" {_money(t.pernocta_usd)}" if t.pernocta_usd else ""))
    ruta_html = escape(ruta or "—")
    if detalles:
        ruta_html += f'<br><span class="muted">{escape(" · ".join(detalles))}</span>'
    tarifa = "—" if t.tarifa_hora_usd is None else _money(t.tarifa_hora_usd)
    tiempo = t.tiempo_hhmm or _hhmm(t.tiempo_hr)
    return (
        f'<tr><td class="ruta">{ruta_html}</td>'
        f"<td>{_dia_mes(t.fecha)}</td>"
        f'<td class="num">{_millas(t.millas)}</td>'
        f'<td class="num">{escape(tiempo)}</td>'
        f'<td class="num">{tarifa}</td>'
        f'<td class="num">{_monto(t.total_usd)}</td></tr>'
    )


def _tramos_html(r: CotizacionInternaPdfRequest) -> str:
    """Tabla de tramos EXACTAMENTE como la hoja de administración + fila TOTAL
    (Σ tramos, del API) y, solo si hay diferencia contra la línea canónica,
    la fila de ajuste con su motivo y el «Servicio aéreo» resultante."""
    tramos = sorted(r.tramos_cotizados, key=lambda t: t.orden)
    if not tramos:
        return (
            "<h2>Tramos cotizados</h2>"
            '<table class="grid tramos"><tbody><tr><td class="muted">'
            "Sin tramos cotizados.</td></tr></tbody></table>"
        )
    filas = "".join(_tramo_fila(t) for t in tramos)
    # Millas: re-suma informativa de la columna (solo si todos los tramos las traen).
    millas = (
        _millas(sum(t.millas for t in tramos if t.millas is not None))
        if all(t.millas is not None for t in tramos)
        else ""
    )
    tiempo_total = r.tramos_tiempo_total_hhmm or _hhmm(r.tramos_tiempo_total_hr)
    pie = (
        '<tr class="total"><td>TOTAL</td><td></td>'
        f'<td class="num">{millas}</td><td class="num">{escape(tiempo_total)}</td><td></td>'
        f'<td class="num">{_monto(r.tramos_total_usd)} USD</td></tr>'
    )
    if _hay_ajuste(r):
        motivo = (r.tramos_ajuste_motivo or "Ajuste de horas").strip()
        pie += (
            f'<tr class="ajuste"><td colspan="5">{escape(motivo)}'
            f'{_op("servicio aéreo cotizado − Σ tramos")}</td>'
            f'<td class="num">{_monto(r.tramos_ajuste_usd)}</td></tr>'
            '<tr class="total"><td colspan="5">Servicio aéreo</td>'
            f'<td class="num">{_monto(_servicio_aereo_usd(r))} USD</td></tr>'
        )
    notas = [NOTA_TRAMOS + (f" ({r.calzos_hr:g} h en total)" if r.calzos_hr else "")]
    notas.append("distancia en millas náuticas")
    if any(t.consolidado for t in tramos):
        notas.append("cotización sin desglose por tramo: una sola fila con los totales")
    return f"""
  <h2>Tramos cotizados</h2>
  <table class="grid tramos"><thead><tr>
    <th>Ruta</th><th>Fecha</th><th class="num">Distancia millas</th>
    <th class="num">Tiempo vuelo</th><th class="num">Costo por hora vuelo</th>
    <th class="num">Total por tramo</th>
  </tr></thead><tbody>{filas}</tbody><tfoot>{pie}</tfoot></table>
  <div class="nota">{escape(" · ".join(notas))}.</div>"""


def _fila_desglose(concepto: str, op: str, monto_html: str) -> str:
    return (
        f'<tr><td class="lbl">{escape(concepto)}{_op(op)}</td>'
        f'<td class="val">{monto_html}</td></tr>'
    )


def _tua_fila(t: CotizacionInternaTuaCobradaPdf) -> str:
    """«TUA CUN · 4 pax × $20.85» — solo TUAS cobradas (las exentas no viajan)."""
    concepto = f"TUA {t.iata}".strip()
    op = f"{t.pax:g} pax × {_money(t.unitario)}"
    if (t.moneda or "USD").upper() == "MXN":
        op += f" MXN = ${t.total_nativo:,.2f} MXN"
        if t.tc_aplicado:
            op += f" · T.C. {t.tc_aplicado:g}"
    return _fila_desglose(concepto, op, _monto(t.total_usd))


def _concepto_operacion(
    ln: CotizacionInternaLineaPdf, r: CotizacionInternaPdfRequest
) -> tuple[str, str]:
    """(concepto, operación en gris) de una línea canónica. La operación se
    arma SOLO con cantidad/unitario que manda el API; el monto no se toca."""
    clave = (ln.clave or "").upper()
    concepto = (ln.concepto or ln.clave or "Concepto").strip()
    if clave == "TIEMPO_VUELO":
        # Etiqueta de administración; la operación sale del snapshot (horas
        # cobrables × tarifa) y, si hubo ajuste, se enlaza con la tabla.
        horas = ln.cantidad if ln.cantidad is not None else r.tiempo_cobrable_hr
        tarifa = ln.unitario if ln.unitario is not None else r.tarifa_hora_usd
        partes: list[str] = []
        if horas is not None and tarifa is not None:
            partes.append(f"{horas:.2f} h × {_money(tarifa)}/hr")
        if r.hora_minima_aplicada:
            partes.append("hora mínima")
        if r.cobrable_override:
            partes.append("cobrable manual")
        if _hay_ajuste(r):
            signo = "+" if r.tramos_ajuste_usd > 0 else "−"
            partes.append(
                f"Σ tramos {_money(r.tramos_total_usd)} {signo} "
                f"{_money(abs(r.tramos_ajuste_usd))}"
            )
        return "Servicio aéreo", " · ".join(partes)
    if clave == "COMISION_VENDEDOR":
        nombre = (r.comision_vendedor_nombre or "").strip()
        if nombre and nombre.lower() not in concepto.lower():
            concepto += f" · {nombre}"
        partes = []
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
    """Desglose interno: líneas canónicas en su orden (IVA aparte, en los
    totales), TUAS solo las cobradas (detalle `tuas_cobradas` en el lugar de
    las líneas TUAS), subtotal sin IVA, IVA con su base, total USD y MXN."""
    lineas = list(r.lineas) or _lineas_fallback(r)
    filas: list[str] = []
    tuas_pintadas = False
    for ln in lineas:
        clave = (ln.clave or "").upper()
        if clave == "IVA":
            continue
        if clave == "TUAS":
            # Exentas (línea sintética legado o monto 0) NO se muestran.
            if ln.exento or (not ln.monto_usd and not r.tuas_cobradas):
                continue
            if r.tuas_cobradas:
                # El detalle por aeropuerto sustituye a las líneas TUAS 1:1
                # (misma suma: `total_usd` de cada fila == línea canónica).
                if not tuas_pintadas:
                    filas.extend(_tua_fila(t) for t in r.tuas_cobradas)
                    tuas_pintadas = True
                continue
        concepto, op = _concepto_operacion(ln, r)
        filas.append(_fila_desglose(concepto, op, _monto(ln.monto_usd)))
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
    iva_op = f"{r.iva_pct:g} % de {_money(r.iva_base_usd)}" if r.iva_base_usd is not None else ""
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
        "<h2>Desglose de la cotización</h2>"
        f'<table class="totales"><tbody>{"".join(filas)}</tbody></table>'
    )


def _horas_html(r: CotizacionInternaPdfRequest) -> str:
    """Horas del snapshot (vuelo + calzos + sobrevuelo → cotizadas → cobrables)
    y tarifa: explican la fila de ajuste de la tabla. Solo lo que venga."""
    filas: list[str] = []
    for lbl, v in (("Vuelo", r.vuelo_hr), ("Calzos", r.calzos_hr), ("Sobrevuelo", r.sobrevuelo_hr)):
        if v:
            filas.append(_kv(lbl, _horas(v)))
    if r.horas_cotizadas_hr is not None:
        filas.append(_kv("Cotizadas", _horas(r.horas_cotizadas_hr)))
    if r.tiempo_cobrable_hr is not None:
        notas = []
        if r.hora_minima_aplicada:
            notas.append("hora mínima aplicada")
        if r.cobrable_override:
            notas.append("cobrable manual")
        cobrables = f"<b>{_horas(r.tiempo_cobrable_hr)}</b>" + _op(", ".join(notas))
        filas.append(_kv("Cobrables", cobrables))
    if r.tarifa_hora_usd is not None:
        filas.append(_kv("Tarifa", f"{_money(r.tarifa_hora_usd)}/hr"))
    if not filas:
        return ""
    return f'<h2>Horas cotizadas</h2><table class="kv">{"".join(filas)}</table>'


def _cobro_fila(c: CotizacionInternaCobroPdf, con_usd: bool) -> str:
    metodo = escape(c.metodo_label or c.metodo or "—")
    sub: list[str] = []
    if c.es_reembolso or c.monto < 0:
        sub.append("Reembolso")
    ref_partes = [p for p in (c.referencia, c.cuenta_destino) if p]
    if ref_partes:
        sub.append(" · ".join(ref_partes))
    if c.sobre_grupo_folio:
        s = f"Sobre {c.sobre_grupo_folio}"
        if c.sobre_grupo_monto_total is not None:
            s += f" {_money(c.sobre_grupo_monto_total)} {c.sobre_grupo_moneda or ''}".rstrip()
        if c.grupo_factor is not None:
            s += f" × {round(c.grupo_factor * 100, 2):g} %"
        sub.append(s)
    if c.notas:
        sub.append(_truncar(c.notas, 60, 1))
    if sub:
        metodo += f'<br><span class="muted">{escape(" · ".join(sub))}</span>'
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
        if c.tc and moneda != "USD":
            usd_txt += f' <span class="muted">T.C. {c.tc:g}</span>'
        usd = f'<td class="num">{usd_txt}</td>'
    return (
        f"<tr><td>{_fecha_corta(c.fecha, con_anio=True)}</td><td>{metodo}</td>"
        f'<td class="num">{bruto}</td><td class="num">{com}</td><td class="num">{neto}</td>'
        f"{usd}<td>{conc}</td></tr>"
    )


def _cobros_html(r: CotizacionInternaPdfRequest) -> str:
    """Cobros del vuelo (cobro_vuelo, partes de sobre incluidas), compactos:
    fecha, método (referencia en gris), bruto, comisión bancaria, neto,
    conciliado; resumen cobrado (cobrosEnUsd), comisiones, neto, saldo y semáforo."""
    con_usd = any((c.moneda or "USD").upper() != "USD" for c in r.cobros)
    ncols = 7 if con_usd else 6
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
    <th>Fecha</th><th>Método</th><th class="num">Bruto</th>
    <th class="num">Comisión banco</th><th class="num">Neto</th>{th_usd}<th>Conc.</th>
  </tr></thead><tbody>{filas}</tbody><tfoot>{pie}</tfoot></table></div>"""


def _notas_html(r: CotizacionInternaPdfRequest) -> str:
    """Solo las notas INTERNAS, recortadas. Las del cliente no son de este documento."""
    internas = _truncar(r.notas_internas)
    if not internas:
        return ""
    return (
        '<div class="bloque"><h2>Notas internas</h2>'
        f'<div class="notas-txt">{escape(internas)}</div></div>'
    )


def _estilos_interno(pie: str) -> str:
    """CSS propio del documento interno (UNA hoja carta). Branding compartido
    con el PDF del cliente vía `_BRAND`/`_NAVY`.

    Aire (11-sep-2026, captura de la oficina: «todo muy junto»): la v2 salió
    a 8.5 pt con interlineado 1.2 y celdas de 1 px — ilegible de un vistazo.
    Reglas de esta hoja, en orden de prioridad:
      1. NINGÚN texto de tabla baja de 9.5 pt (las aclaraciones en gris y los
         encabezados van a 8.5 pt: son apoyo, no el dato).
      2. Interlineado 1.3 y celdas de 3 px verticales / 6 px horizontales
         (antes 1 px × 4 px): las filas respiran.
      3. Cada bloque se separa del anterior con el margen SUPERIOR de su
         `h2` (9 px, antes 5) — por eso «Desglose de la cotización» y «Horas
         cotizadas», que van lado a lado, nunca se tocan: además del margen
         llevan 16 px de canal y una línea divisoria.
      4. `@page` de 9 mm arriba / 12 mm a los lados / 10 mm abajo.
    La MISMA información sigue cabiendo en UNA hoja (ver el presupuesto del
    encabezado del módulo): lo que se encogió es el número de tramos/cobros
    que caben, no el contenido.
    """
    fuente = "'Helvetica Neue', Arial, sans-serif"
    return f"""
  @page {{
    size: Letter;
    margin: 9mm 12mm 10mm;
    @bottom-left {{ content: "{_css_str(pie)}"; font-size: 7.5pt; color: #9ca3af;
                    font-family: {fuente}; }}
    @bottom-right {{ content: "Página " counter(page) " de " counter(pages);
                     font-size: 7.5pt; color: #9ca3af; font-family: {fuente}; }}
  }}
  * {{ font-family: {fuente}; color: #1d1d1d; box-sizing: border-box; }}
  body {{ margin: 0; font-size: 9.5pt; line-height: 1.3; }}
  .header {{ background: {_NAVY}; color: #fff; padding: 7px 10px; border-radius: 5px;
             display: flex; align-items: center; justify-content: space-between; }}
  .header .izq {{ display: flex; align-items: center; }}
  .header .logo {{ height: 18px; margin-right: 10px; }}
  .header .titulo {{ font-size: 11.5pt; font-weight: 800; color: #fff; line-height: 1.2; }}
  .header .sub {{ font-size: 8pt; color: #9fb3c8; display: block; }}
  .header .folio {{ text-align: right; color: #fff; font-size: 9.5pt; font-weight: 700; }}
  .banda {{ margin: 5px 0 7px; padding: 3px 9px; border-left: 3px solid {_BRAND};
            background: #fef2f2; color: {_BRAND}; font-size: 8pt; font-weight: 800;
            letter-spacing: 1px; text-transform: uppercase; }}
  h2 {{ font-size: 9pt; text-transform: uppercase; letter-spacing: .8px; color: {_BRAND};
        border-bottom: 1px solid #e5e7eb; margin: 9px 0 4px; padding-bottom: 2px; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 9.5pt; }}
  table.resumen {{ margin: 0 0 4px; border-collapse: separate; border-spacing: 6px 0; }}
  table.resumen td {{ width: 33.3%; vertical-align: top; padding: 5px 8px;
                      border: 1px solid #e5e7eb; border-radius: 4px; background: #f9fafb; }}
  table.resumen td.fecha {{ background: #fef2f2; border-color: #fecaca; }}
  table.resumen .lbl {{ font-size: 7.5pt; text-transform: uppercase; letter-spacing: .6px;
                        color: #6b7280; }}
  table.resumen .big {{ font-size: 14pt; font-weight: 800; color: {_BRAND}; line-height: 1.2; }}
  table.resumen .med {{ font-size: 10.5pt; font-weight: 700; color: {_NAVY}; line-height: 1.3; }}
  table.resumen .sub {{ font-size: 8.5pt; color: #6b7280; line-height: 1.35; }}
  /* Segunda línea del avión cuando DIFIERE del cotizado: gana al gris de
     .sub (misma especificidad + .ambar) para que la marca se vea. */
  table.resumen .sub.ambar {{ color: #b45309; font-weight: 700; }}
  table.kv td {{ padding: 2px 4px; vertical-align: top; }}
  table.kv td.k {{ color: #6b7280; width: 32%; white-space: nowrap; }}
  table.kv tr.small td {{ font-size: 8.5pt; color: #6b7280; }}
  table.grid {{ font-size: 9.5pt; }}
  table.grid th, table.grid td {{ border: 1px solid #e5e7eb; padding: 3px 6px;
                                  text-align: left; vertical-align: top; }}
  table.grid th {{ background: #f3f4f6; font-size: 8.5pt; text-transform: uppercase;
                   color: #374151; letter-spacing: .3px; }}
  table.grid tfoot td {{ font-weight: 700; background: #f7f7f8; color: {_NAVY}; }}
  table.grid tfoot td.aviso {{ font-weight: 400; background: #fff; }}
  table.grid.tramos td.ruta {{ font-weight: 600; }}
  table.grid.tramos tfoot tr.ajuste td {{ font-weight: 400; background: #fffbeb; color: #92400e; }}
  thead {{ display: table-header-group; }}
  /* Si la tabla pasa a otra hoja: el encabezado se repite (arriba) y ninguna
     fila se parte a la mitad. */
  table.grid tr {{ page-break-inside: avoid; }}
  .num {{ text-align: right !important; white-space: nowrap; }}
  .muted {{ color: #6b7280; font-size: 8pt; font-weight: 400; }}
  .op {{ color: #6b7280; font-size: 8pt; font-weight: 400; }}
  .nota {{ font-size: 8pt; color: #6b7280; margin: 3px 0 2px; }}
  .ambar {{ color: #b45309; }}  .rojo {{ color: {_BRAND}; }}  .verde {{ color: #15803d; }}
  .tag {{ display: inline-block; border: 1px solid #d1d5db; border-radius: 3px; padding: 0 4px;
          font-size: 8pt; color: #374151; margin-left: 4px; font-weight: 400; }}
  .tag.ambar {{ border-color: #f59e0b; }}
  .totales td {{ padding: 2px 4px; vertical-align: top; }}
  .totales .val {{ text-align: right; font-weight: 600; white-space: nowrap; }}
  .sub-row td {{ border-top: 1px solid #d1d5db; font-weight: 700; color: {_NAVY};
                 padding-top: 4px; }}
  .total-row td {{ border-top: 2px solid {_NAVY}; font-size: 10.5pt; font-weight: 800;
                   color: {_BRAND}; padding-top: 4px; }}
  .mxn-row td {{ font-weight: 700; color: {_NAVY}; }}
  .cols {{ width: 100%; border-collapse: separate; border-spacing: 0; }}
  .cols td.col {{ vertical-align: top; width: 50%; padding: 0; }}
  .cols td.col:first-child {{ padding-right: 16px; }}
  .cols td.col:last-child {{ padding-left: 16px; }}
  .cols.desglose td.col:first-child {{ width: 58%; }}
  .cols.desglose td.col:last-child {{ width: 42%; border-left: 1px solid #eef0f3; }}
  .sem {{ display: inline-block; width: 8px; height: 8px; border-radius: 50%;
          vertical-align: middle; margin-right: 4px; }}
  .sem-verde {{ background: #16a34a; }}  .sem-amarillo {{ background: #f59e0b; }}
  .sem-rojo {{ background: {_BRAND}; }}  .sem-gris {{ background: #9ca3af; }}
  .bloque {{ page-break-inside: avoid; }}
  .notas-txt {{ font-size: 9pt; color: #374151; white-space: pre-wrap; line-height: 1.4; }}"""


def _build_html(r: CotizacionInternaPdfRequest) -> str:
    pie = _pie_texto(r)
    cuerpo = (
        '<table class="cols desglose bloque"><tr>'
        f'<td class="col">{_desglose_html(r)}</td>'
        f'<td class="col">{_horas_html(r)}</td>'
        "</tr></table>"
    )
    return f"""<!doctype html>
<html><head><meta charset="utf-8"><title>Cotización interna #{escape(r.folio)}</title><style>
{_estilos_interno(pie)}
</style></head><body>
{_header_html(r)}
  {_resumen_html(r)}
  {_ficha_html(r)}
  {_tramos_html(r)}
  {cuerpo}
  {_cobros_html(r)}
  {_notas_html(r)}
</body></html>"""


def render_cotizacion_interna_pdf(req: CotizacionInternaPdfRequest) -> bytes:
    from weasyprint import HTML  # import perezoso

    return HTML(string=_build_html(req)).write_pdf()
