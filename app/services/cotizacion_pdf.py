"""Genera el PDF de una cotización con WeasyPrint (HTML → PDF).

Rediseño 26-ago-2026 (pedido del cliente): formato profesional con
membrete/marca de agua del logo, MAPA de la ruta, SIN horas de vuelo
(el cliente no debe ver tiempo cobrable ni tarifa por hora) y fotos
exterior/interior del avión cotizado al final.

Mantiene la identidad visual del admin: rojo de marca #dc2626, navy #102a43.
El import de WeasyPrint es perezoso para que el servicio arranque aunque la
librería (y sus libs de sistema: pango/cairo) no esté instalada.
"""

import base64
import json
from datetime import datetime
from functools import lru_cache
from html import escape
from pathlib import Path
from zoneinfo import ZoneInfo

from app.schemas.reportes import CotizacionPdfRequest, MapaPuntoPdf

_BRAND = "#dc2626"
_NAVY = "#102a43"
# Todo se muestra en hora de Cancún (Quintana Roo, UTC−5, sin horario de verano).
_CANCUN = ZoneInfo("America/Cancun")
TZ_NOTA = "Horarios en hora de Cancún (UTC−5)."

_ASSETS = Path(__file__).resolve().parent.parent / "assets"


def _money(v: float) -> str:
    return f"${v:,.2f}"


def _fecha_legible(s: str | None) -> str:
    if not s:
        return "Por confirmar"
    try:
        dt = datetime.fromisoformat(s.replace("Z", "+00:00"))
        # Si viene sin zona, se asume UTC; luego se convierte a Cancún.
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=ZoneInfo("UTC"))
        return dt.astimezone(_CANCUN).strftime("%d/%m/%Y %H:%M")
    except ValueError:
        return s


@lru_cache(maxsize=1)
def _logo_data_uri(nombre: str = "logo-vuelatour-blanco.png") -> str | None:
    """Logo como data-URI (el HTML del PDF no puede pedir archivos remotos)."""
    try:
        raw = (_ASSETS / nombre).read_bytes()
        return "data:image/png;base64," + base64.b64encode(raw).decode()
    except OSError:
        return None


# ===== Mapa de ruta (26-ago): mismo GeoJSON de la península que usa el =====
# panel, con la MISMA idea de proyección (identidad sobre lon/lat con la Y
# reflejada y ajuste al viewport). Un fit lineal replica geoIdentity().
_VIEW_W, _VIEW_H = 600, 420
_MARGEN = 28


@lru_cache(maxsize=1)
def _peninsula_paths() -> tuple[list[str], tuple[float, float, float, float, float]]:
    """Paths SVG de la península + parámetros (minx, maxy, escala, offx, offy)."""
    data = json.loads((_ASSETS / "yucatan-peninsula.json").read_text())
    anillos: list[list[list[float]]] = []
    for f in data.get("features", []):
        g = f.get("geometry", {})
        if g.get("type") == "Polygon":
            anillos.extend(g["coordinates"])
        elif g.get("type") == "MultiPolygon":
            for poly in g["coordinates"]:
                anillos.extend(poly)
    xs = [pt[0] for an in anillos for pt in an]
    ys = [pt[1] for an in anillos for pt in an]
    minx, maxx, miny, maxy = min(xs), max(xs), min(ys), max(ys)
    esc = min(
        (_VIEW_W - 2 * _MARGEN) / (maxx - minx),
        (_VIEW_H - 2 * _MARGEN) / (maxy - miny),
    )
    offx = (_VIEW_W - (maxx - minx) * esc) / 2
    offy = (_VIEW_H - (maxy - miny) * esc) / 2
    params = (minx, maxy, esc, offx, offy)

    def xy(lon: float, lat: float) -> tuple[float, float]:
        return (offx + (lon - minx) * esc, offy + (maxy - lat) * esc)

    paths: list[str] = []
    for an in anillos:
        puntos = [xy(p[0], p[1]) for p in an]
        d = "M" + " L".join(f"{x:.1f},{y:.1f}" for x, y in puntos) + " Z"
        paths.append(d)
    return paths, params


def _xy(lon: float, lat: float) -> tuple[float, float]:
    _, (minx, maxy, esc, offx, offy) = _peninsula_paths()
    return (offx + (lon - minx) * esc, offy + (maxy - lat) * esc)


def _mapa_svg(puntos: list[MapaPuntoPdf]) -> str:
    """SVG del itinerario: península + arcos numerados (ferry punteado).

    ZOOM a la ruta (27-ago, pedido del cliente): el viewBox se ajusta al
    bounding box de los puntos con aire — igual que el mapa del admin — para
    que rutas cortas (CUN–CZM) no se vean diminutas. Trazos, radios y textos
    se escalan por el factor de zoom para imprimir al tamaño de siempre.
    """
    if not puntos:
        return ""
    paths, _ = _peninsula_paths()

    orden = sorted(puntos, key=lambda x: x.orden)
    pts_xy: list[tuple[float, float]] = []
    for p in orden:
        pts_xy.append(_xy(p.o_lon, p.o_lat))
        pts_xy.append(_xy(p.d_lon, p.d_lat))
    xs = [x for x, _ in pts_xy]
    ys = [y for _, y in pts_xy]
    span = max(max(xs) - min(xs), max(ys) - min(ys), 1.0)
    # Aire proporcional + fijo (los arcos curvan hacia afuera y las
    # etiquetas cuelgan bajo el punto).
    pad = span * 0.30 + 18.0
    bx0, by0 = min(xs) - pad, min(ys) - pad
    bw, bh = (max(xs) - min(xs)) + 2 * pad, (max(ys) - min(ys)) + 2 * pad
    # Conservar el aspecto del lienzo.
    aspecto = _VIEW_W / _VIEW_H
    if bw / bh > aspecto:
        extra = bw / aspecto - bh
        by0 -= extra / 2
        bh = bw / aspecto
    else:
        extra = bh * aspecto - bw
        bx0 -= extra / 2
        bw = bh * aspecto
    # Límite de acercamiento (~3.2x): conservar contexto geográfico.
    min_w = _VIEW_W / 3.2
    if bw < min_w:
        cx, cy = bx0 + bw / 2, by0 + bh / 2
        bw, bh = min_w, min_w / aspecto
        bx0, by0 = cx - bw / 2, cy - bh / 2
    # Nunca más lejos que la vista completa original.
    if bw > _VIEW_W:
        bx0, by0, bw, bh = 0.0, 0.0, float(_VIEW_W), float(_VIEW_H)
    k = bw / _VIEW_W  # factor de escala de trazos/textos (1 = sin zoom)

    fondo = "".join(
        f'<path d="{d}" fill="#e8eef5" stroke="#b7c6d6" stroke-width="{1 * k:.2f}"/>'
        for d in paths
    )
    arcos: list[str] = []
    marcadores: list[str] = []
    etiquetas: dict[str, tuple[float, float]] = {}
    for p in orden:
        x1, y1 = _xy(p.o_lon, p.o_lat)
        x2, y2 = _xy(p.d_lon, p.d_lat)
        dx, dy = x2 - x1, y2 - y1
        largo = max((dx * dx + dy * dy) ** 0.5, 1.0)
        # Curvatura perpendicular (los tramos de ida y regreso se separan solos).
        cx = (x1 + x2) / 2 - dy / largo * largo * 0.18
        cy = (y1 + y2) / 2 + dx / largo * largo * 0.18
        dash = f' stroke-dasharray="{6 * k:.1f} {5 * k:.1f}"' if p.es_ferry else ""
        arcos.append(
            f'<path d="M{x1:.1f},{y1:.1f} Q{cx:.1f},{cy:.1f} {x2:.1f},{y2:.1f}" '
            f'fill="none" stroke="{_BRAND}" stroke-width="{2.4 * k:.2f}"{dash}/>'
        )
        # Badge numerado en el punto medio del arco (t = 0.5 del bezier).
        bx = 0.25 * x1 + 0.5 * cx + 0.25 * x2
        by = 0.25 * y1 + 0.5 * cy + 0.25 * y2
        arcos.append(
            f'<circle cx="{bx:.1f}" cy="{by:.1f}" r="{9 * k:.2f}" fill="{_BRAND}"/>'
            f'<text x="{bx:.1f}" y="{by + 3.4 * k:.1f}" text-anchor="middle" '
            f'font-size="{10 * k:.2f}" font-weight="700" fill="#ffffff">{p.orden}</text>'
        )
        etiquetas[p.origen_iata.upper()] = (x1, y1)
        etiquetas[p.destino_iata.upper()] = (x2, y2)
    for iata, (x, y) in etiquetas.items():
        marcadores.append(
            f'<circle cx="{x:.1f}" cy="{y:.1f}" r="{5 * k:.2f}" fill="#ffffff" '
            f'stroke="{_BRAND}" stroke-width="{2.5 * k:.2f}"/>'
            f'<text x="{x:.1f}" y="{y + 16 * k:.1f}" text-anchor="middle" '
            f'font-size="{10 * k:.2f}" '
            f'font-weight="700" fill="{_NAVY}">{escape(iata)}</text>'
        )
    return (
        '<div class="mapa">'
        f'<svg viewBox="{bx0:.1f} {by0:.1f} {bw:.1f} {bh:.1f}" '
        'xmlns="http://www.w3.org/2000/svg">'
        f"{fondo}{''.join(arcos)}{''.join(marcadores)}</svg></div>"
    )


def _build_html(r: CotizacionPdfRequest) -> str:
    # Matrículas OCULTAS en la cotización (regla 26-ago): el cliente no debe
    # ver qué avión es — EXCEPTO el VGV, que sí se comercializa por matrícula.
    mostrar_matricula = bool(r.matricula and "VGV" in r.matricula.upper())
    # Título con la RUTA COMPLETA (26-ago): "CUN → CTM → CUN", no solo
    # origen→destino. Fuente más chica si la ruta es larga (multiescala).
    if r.escalas:
        ordenadas = sorted(r.escalas, key=lambda x: x.orden)
        puntos = [ordenadas[0].origen] + [e.destino for e in ordenadas]
        ruta_titulo = " → ".join(escape(pt) for pt in puntos)
    else:
        ruta_titulo = f"{escape(r.origen)} → {escape(r.destino)}"
    n_puntos = ruta_titulo.count("→") + 1
    ruta_font = "26px" if n_puntos <= 4 else ("20px" if n_puntos <= 6 else "16px")

    # Mapa JUNTO al itinerario en la hoja 1 (26-ago v3, pedido del cliente):
    # tabla de tramos a la izquierda, mapa a la derecha (layout de tabla —
    # WeasyPrint lo respeta siempre). Sin puntos de mapa, la tabla va sola.
    mapa_html = _mapa_svg(r.mapa_puntos)
    escalas_html = ""
    if r.escalas:
        filas = "".join(
            f"<tr><td>{e.orden}</td><td>{escape(e.origen)} → {escape(e.destino)}</td></tr>"
            for e in sorted(r.escalas, key=lambda x: x.orden)
        )
        tabla_itin = (
            '<table class="grid"><thead><tr><th>#</th><th>Tramo</th></tr></thead>'
            f"<tbody>{filas}</tbody></table>"
        )
        if not r.mostrar_itinerario:
            # Itinerario OCULTO por configuración de la cotización (27-ago):
            # queda solo el mapa de la ruta, centrado.
            escalas_html = (
                f'<h2>La ruta</h2><div class="mapa-solo">{mapa_html}</div>'
                if mapa_html
                else ""
            )
        elif mapa_html:
            cuerpo_itin = (
                '<table class="itin-row"><tr>'
                f'<td class="itin-tabla">{tabla_itin}</td>'
                f'<td class="itin-mapa">{mapa_html}</td>'
                "</tr></table>"
            )
            escalas_html = f"""
        <h2>Itinerario</h2>
        {cuerpo_itin}"""
        else:
            escalas_html = f"""
        <h2>Itinerario</h2>
        {tabla_itin}"""

    notas_html = (
        f'<div class="notas"><strong>Notas:</strong> {escape(r.notas)}</div>' if r.notas else ""
    )

    # ----- Desglose SIN horas (26-ago, regla del cliente): ni tiempo
    # cobrable ni tarifa por hora — el servicio se presenta como monto. -----
    filas: list[str] = []

    def fila(lbl: str, val: str) -> None:
        filas.append(f'<tr><td class="lbl">{escape(lbl)}</td><td class="val">{val}</td></tr>')

    # Tarifa por hora VISIBLE solo si la cotización lo pide (27-ago).
    if (
        r.mostrar_tarifa_hora
        and r.tiempo_cobrable_hr
        and r.tarifa_hora_usd
        and r.tiempo_cobrable_hr > 0
        and r.tarifa_hora_usd > 0
    ):
        fila(
            f"Servicio aéreo ({r.tiempo_cobrable_hr:g} h × "
            f"{_money(r.tarifa_hora_usd)}/hr)",
            _money(r.subtotal_usd),
        )
    else:
        fila("Servicio aéreo", _money(r.subtotal_usd))
    if not r.tuas_detalle:
        fila("TUAS", _money(r.tuas_usd))
    elif len(r.tuas_detalle) == 1:
        fila(r.tuas_detalle[0], _money(r.tuas_usd))
    else:
        for det in r.tuas_detalle:
            fila(det, "")
        fila("TUAS (total)", _money(r.tuas_usd))
    for e in r.extras:
        lbl = e.concepto or "Extra"
        if e.moneda == "MXN" and e.monto_nativo is not None:
            lbl += f" · ${e.monto_nativo:,.2f} MXN"
        fila(lbl, _money(e.monto_usd))
    if r.viaticos_pernocta_usd > 0:
        fila("Viáticos por pernocta", _money(r.viaticos_pernocta_usd))
    if r.descuento_usd > 0:
        fila("Descuento", f"&minus;{_money(r.descuento_usd)}")
    # Subtotal SIN IVA antes del IVA (27-ago, pedido del cliente). Se deriva
    # del total canónico (total − IVA) para cuadrar exacto con el desglose.
    filas.append(
        '<tr class="sub-row"><td class="lbl">Subtotal (sin IVA)</td>'
        f'<td class="val">{_money(r.total_usd - r.iva_usd)}</td></tr>'
    )
    fila(f"IVA ({r.iva_pct:.0f}%)", _money(r.iva_usd))
    desglose_html = "".join(filas)
    total_row_html = (
        f'<tr class="total-row"><td>Total ({escape(r.moneda)})</td>'
        f'<td class="val">{_money(r.total_usd)}</td></tr>'
    )

    total_mxn_html = ""
    if r.total_mxn is not None:
        tc_txt = f" (T.C. {r.tc_usd_mxn:g})" if r.tc_usd_mxn else ""
        total_mxn_html = (
            f'<tr class="total-mxn"><td>Total MXN{escape(tc_txt)}</td>'
            f'<td class="val">{_money(r.total_mxn)} MXN</td></tr>'
        )

    # ----- Marca de agua + logo del membrete -----
    logo_blanco = _logo_data_uri("logo-vuelatour-blanco.png")
    logo_marca = _logo_data_uri("logo-vuelatour.png")
    logo_header_html = (
        f'<img class="logo" src="{logo_blanco}" alt=""/>' if logo_blanco else ""
    )
    marca_html = (
        f'<div class="marca"><img src="{logo_marca}" alt=""/></div>' if logo_marca else ""
    )


    # ----- Página "La aeronave" (26-ago v2, mockup del cliente) -----
    # Exterior ANCHO arriba; abajo interior + tarjeta "De un vistazo";
    # tira de características comerciales al pie. SIN matrícula en esta hoja
    # (regla del cliente — aplica también al VGV; la página 1 conserva su
    # propia regla de matrícula).
    vistazo: list[tuple[str, str]] = []
    if r.avion_pasajeros:
        vistazo.append(("Pasajeros", f"{r.avion_pasajeros} máx."))
    if r.avion_velocidad_kts:
        kmh = round(r.avion_velocidad_kts * 1.852)
        vistazo.append(
            ("Velocidad crucero", f"{r.avion_velocidad_kts:g} kt / {kmh} km/h")
        )
    if r.avion_tiempo_tramo_hr and r.avion_tiempo_tramo_hr > 0:
        th = int(r.avion_tiempo_tramo_hr)
        tm = int(round((r.avion_tiempo_tramo_hr - th) * 60))
        if tm == 60:
            th, tm = th + 1, 0
        vistazo.append(("Tiempo de vuelo", f"{th}:{tm:02d} h por tramo"))
    if r.avion_num_motores:
        if r.avion_motor_hp:
            vistazo.append(
                (
                    "Motor" if r.avion_num_motores == 1 else "Motores",
                    f"{r.avion_num_motores} × {r.avion_motor_hp} HP",
                )
            )
        else:
            vistazo.append(
                ("Motor", "Monomotor" if r.avion_num_motores == 1 else "Bimotor")
            )
    vistazo_html = ""
    if vistazo:
        filas_vz = "".join(
            f'<div class="vz-lbl">{escape(lb)}</div>'
            f'<div class="vz-val">{escape(vl)}</div>'
            for lb, vl in vistazo
        )
        vistazo_html = (
            '<div class="vistazo"><div class="vz-titulo">De un vistazo</div>'
            f"{filas_vz}</div>"
        )

    chips = " &nbsp;·&nbsp; ".join(
        escape(c.strip()) for c in r.avion_caracteristicas if c and c.strip()
    )
    caracts_html = f'<div class="caracts">{chips}</div>' if chips else ""

    # La foto ANCHA de arriba: exterior; si solo hay interior, sube esa.
    foto_ancha = r.foto_exterior or r.foto_interior
    foto_abajo = r.foto_interior if r.foto_exterior else None
    ancha_html = (
        f'<img class="foto-ancha" src="{foto_ancha}" alt=""/>' if foto_ancha else ""
    )
    # Fila inferior: interior + tarjeta lado a lado (tabla: WeasyPrint la
    # respeta siempre); sin interior, la tarjeta ocupa la fila completa.
    if foto_abajo and vistazo_html:
        fila_html = (
            '<table class="av-row"><tr>'
            f'<td class="av-foto"><img src="{foto_abajo}" alt=""/></td>'
            f'<td class="av-card">{vistazo_html}</td>'
            "</tr></table>"
        )
    elif foto_abajo:
        fila_html = f'<img class="foto-ancha" src="{foto_abajo}" alt=""/>'
    else:
        fila_html = vistazo_html

    fotos_html = ""
    if ancha_html or fila_html or caracts_html:
        titulo_avion = escape(r.avion_modelo) if r.avion_modelo else "La aeronave"
        fotos_html = f"""
        <div class="av-titulo">{titulo_avion}</div>
        <div class="av-linea"></div>
        {ancha_html}
        {fila_html}
        {caracts_html}"""

    detalles_html = (
        f'<div class="detalles">{fotos_html}</div>' if fotos_html else ""
    )

    return f"""<!doctype html>
<html><head><meta charset="utf-8"><style>
  /* Pie en TODAS las hojas (26-ago v3): la web al centro y el paginado a
     la derecha; la hoja 1 además conserva su leyenda de horarios/gracias. */
  @page {{
    size: Letter;
    margin: 1.8cm 2cm 2.2cm;
    @bottom-center {{
      content: "www.vuelatour.com";
      font-size: 10px;
      color: #9ca3af;
      font-family: 'Helvetica Neue', Arial, sans-serif;
    }}
    @bottom-right {{
      content: "Página " counter(page) " de " counter(pages);
      font-size: 9px;
      color: #9ca3af;
      font-family: 'Helvetica Neue', Arial, sans-serif;
    }}
  }}
  @page :first {{
    margin-bottom: 2.6cm;
    @bottom-center {{
      content: "{TZ_NOTA} \\A Gracias por volar con VuelaTour, "
               "Aero Charter Cancún. \\A www.vuelatour.com";
      white-space: pre;
      font-size: 10px;
      color: #9ca3af;
      text-align: center;
      font-family: 'Helvetica Neue', Arial, sans-serif;
    }}
  }}
  * {{ font-family: 'Helvetica Neue', Arial, sans-serif; color: #1d1d1d; }}
  /* Marca de agua: fija = se repite en TODAS las páginas (WeasyPrint). */
  .marca {{ position: fixed; top: 34%; left: 0; right: 0; text-align: center;
            opacity: 0.05; z-index: -1; }}
  .marca img {{ width: 78%; transform: rotate(-18deg); }}
  .header {{ background: {_NAVY}; color: #fff; padding: 16px 24px; border-radius: 10px;
             display: flex; align-items: center; justify-content: space-between; }}
  .header .logo {{ height: 30px; }}
  .header .titulos p {{ margin: 2px 0 0; color: #9fb3c8; font-size: 11px;
                        text-align: right; }}
  .header .titulos h1 {{ margin: 0; font-size: 17px; color: #fff; text-align: right; }}
  .meta {{ display: flex; justify-content: space-between; margin: 18px 0 14px; font-size: 13px; }}
  .route {{ font-size: 26px; font-weight: 800; color: {_NAVY}; margin: 6px 0 2px; }}
  h2 {{ font-size: 12px; text-transform: uppercase; letter-spacing: 1.2px; color: #6b7280;
        margin: 20px 0 8px; border-bottom: 1px solid #e5e7eb; padding-bottom: 4px; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
  .grid th, .grid td {{ border: 1px solid #e5e5e5; padding: 6px 10px; text-align: left; }}
  .grid th {{ background: #f7f7f8; }}
  /* Mapa junto al itinerario (26-ago v3): dos columnas en la hoja 1. */
  .itin-row {{ width: 100%; border-collapse: separate; border-spacing: 0; }}
  .itin-row td {{ vertical-align: top; }}
  .itin-tabla {{ width: 55%; padding-right: 12px; }}
  .itin-mapa {{ width: 45%; }}
  .mapa-solo {{ width: 60%; margin: 0 auto; }}
  .mapa {{ width: 100%; border: 1px solid #e5e7eb;
           border-radius: 10px; overflow: hidden; background: #f8fafc; }}
  .mapa svg {{ width: 100%; display: block; }}
  .totales td {{ padding: 7px 0; }}
  .totales .lbl {{ color: #6b7280; }}
  .totales .val {{ text-align: right; font-weight: 600; }}
  .sub-row td {{ border-top: 1px solid #d1d5db; padding-top: 8px; font-weight: 700;
                 color: {_NAVY}; }}
  .total-row td {{ border-top: 2px solid {_NAVY}; padding-top: 12px; font-size: 18px;
                   font-weight: 800; color: {_BRAND}; }}
  .total-mxn td {{ font-size: 13px; font-weight: 700; color: {_NAVY}; }}
  .notas {{ margin-top: 20px; font-size: 12px; color: #374151; }}
  /* Página 2 (26-ago): SOLO imágenes (mapa de la ruta + fotos del avión);
     la página 1 lleva cotización + traslados + itinerario. */
  .detalles {{ page-break-before: always; }}
  /* Página "La aeronave" (26-ago v2, mockup del cliente). */
  .av-titulo {{ font-size: 24px; font-weight: 800; color: #111827; margin: 0; }}
  .av-linea {{ width: 3.2cm; height: 4px; background: {_BRAND}; margin: 6px 0 14px; }}
  .foto-ancha {{ width: 100%; height: 8.6cm; object-fit: cover; border-radius: 12px;
                 border: 1px solid #e5e7eb; margin-bottom: 12px; }}
  .av-row {{ width: 100%; border-collapse: separate; border-spacing: 0; }}
  .av-row td {{ vertical-align: top; }}
  .av-foto {{ width: 62%; padding-right: 12px; }}
  .av-foto img {{ width: 100%; height: 8.2cm; object-fit: cover; border-radius: 12px;
                  border: 1px solid #e5e7eb; }}
  .vistazo {{ background: {_NAVY}; border-top: 6px solid {_BRAND};
              border-radius: 12px; padding: 16px 18px; height: 8.2cm;
              box-sizing: border-box; }}
  .vz-titulo {{ color: #d6bf8e; font-size: 11px; font-weight: 700;
               letter-spacing: 3px; text-transform: uppercase;
               margin-bottom: 12px; }}
  .vz-lbl {{ color: #94a3b8; font-size: 10px; text-transform: uppercase;
             letter-spacing: 1px; margin-top: 9px; }}
  .vz-val {{ color: #ffffff; font-size: 16px; font-weight: 700; margin-top: 1px; }}
  .caracts {{ margin-top: 12px; background: #f3f4f6; border-radius: 10px;
              padding: 11px 14px; font-size: 12px; color: #374151;
              text-align: center; }}
</style></head><body>
  {marca_html}
  <div class="header">
    {logo_header_html}
    <div class="titulos">
      <h1>{escape(r.empresa)}</h1>
      <p>Cotización de servicio aéreo</p>
    </div>
  </div>

  <div class="meta">
    <div><strong>Folio:</strong> #{escape(r.folio)}<br><strong>Cliente:</strong> {escape(r.cliente)}</div>
    <div style="text-align:right"><strong>Fecha de cotización:</strong> {_fecha_legible(r.fecha)}<br>
      <strong>Tipo:</strong> {escape(r.tipo)}</div>
  </div>

  <div class="route" style="font-size:{ruta_font}">{ruta_titulo}</div>
  <div style="font-size:13px;color:#374151">
    {r.pasajeros} {'pasajero' if r.pasajeros == 1 else 'pasajeros'}{f" · {escape(r.avion_externo)}" if r.avion_externo else (f" · {escape(r.matricula)}" if mostrar_matricula else "")}
  </div>

  <h2>Traslados</h2>
  <table class="grid"><tbody>
    <tr><td>Traslado inicial</td><td>{_fecha_legible(r.fecha_traslado_inicial)}</td></tr>
    <tr><td>Traslado final</td><td>{_fecha_legible(r.fecha_traslado_final)}</td></tr>
  </tbody></table>
  {escalas_html}

  <h2>Desglose</h2>
  <table class="totales"><tbody>
    {desglose_html}
    {total_row_html}
    {total_mxn_html}
  </tbody></table>
  {notas_html}

  {detalles_html}

</body></html>"""


def render_cotizacion_pdf(req: CotizacionPdfRequest) -> bytes:
    from weasyprint import HTML  # import perezoso

    return HTML(string=_build_html(req)).write_pdf()
