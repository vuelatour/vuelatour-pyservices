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
    """SVG del itinerario: península + arcos numerados (ferry punteado)."""
    if not puntos:
        return ""
    paths, _ = _peninsula_paths()
    fondo = "".join(
        f'<path d="{d}" fill="#e8eef5" stroke="#b7c6d6" stroke-width="1"/>' for d in paths
    )
    arcos: list[str] = []
    marcadores: list[str] = []
    etiquetas: dict[str, tuple[float, float]] = {}
    for p in sorted(puntos, key=lambda x: x.orden):
        x1, y1 = _xy(p.o_lon, p.o_lat)
        x2, y2 = _xy(p.d_lon, p.d_lat)
        dx, dy = x2 - x1, y2 - y1
        largo = max((dx * dx + dy * dy) ** 0.5, 1.0)
        # Curvatura perpendicular (los tramos de ida y regreso se separan solos).
        cx = (x1 + x2) / 2 - dy / largo * largo * 0.18
        cy = (y1 + y2) / 2 + dx / largo * largo * 0.18
        dash = ' stroke-dasharray="6 5"' if p.es_ferry else ""
        arcos.append(
            f'<path d="M{x1:.1f},{y1:.1f} Q{cx:.1f},{cy:.1f} {x2:.1f},{y2:.1f}" '
            f'fill="none" stroke="{_BRAND}" stroke-width="2.4"{dash}/>'
        )
        # Badge numerado en el punto medio del arco (t = 0.5 del bezier).
        bx = 0.25 * x1 + 0.5 * cx + 0.25 * x2
        by = 0.25 * y1 + 0.5 * cy + 0.25 * y2
        arcos.append(
            f'<circle cx="{bx:.1f}" cy="{by:.1f}" r="9" fill="{_BRAND}"/>'
            f'<text x="{bx:.1f}" y="{by + 3.4:.1f}" text-anchor="middle" '
            f'font-size="10" font-weight="700" fill="#ffffff">{p.orden}</text>'
        )
        etiquetas[p.origen_iata.upper()] = (x1, y1)
        etiquetas[p.destino_iata.upper()] = (x2, y2)
    for iata, (x, y) in etiquetas.items():
        marcadores.append(
            f'<circle cx="{x:.1f}" cy="{y:.1f}" r="5" fill="#ffffff" '
            f'stroke="{_BRAND}" stroke-width="2.5"/>'
            f'<text x="{x:.1f}" y="{y + 16:.1f}" text-anchor="middle" font-size="10" '
            f'font-weight="700" fill="{_NAVY}">{escape(iata)}</text>'
        )
    return (
        '<div class="mapa">'
        f'<svg viewBox="0 0 {_VIEW_W} {_VIEW_H}" xmlns="http://www.w3.org/2000/svg">'
        f"{fondo}{''.join(arcos)}{''.join(marcadores)}</svg></div>"
    )


def _build_html(r: CotizacionPdfRequest) -> str:
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

    escalas_html = ""
    if r.escalas:
        filas = "".join(
            f"<tr><td>{e.orden}</td><td>{escape(e.origen)} → {escape(e.destino)}</td></tr>"
            for e in sorted(r.escalas, key=lambda x: x.orden)
        )
        escalas_html = f"""
        <h2>Itinerario</h2>
        <table class="grid"><thead><tr><th>#</th><th>Tramo</th></tr></thead>
        <tbody>{filas}</tbody></table>"""

    notas_html = (
        f'<div class="notas"><strong>Notas:</strong> {escape(r.notas)}</div>' if r.notas else ""
    )

    # ----- Desglose SIN horas (26-ago, regla del cliente): ni tiempo
    # cobrable ni tarifa por hora — el servicio se presenta como monto. -----
    filas: list[str] = []

    def fila(lbl: str, val: str) -> None:
        filas.append(f'<tr><td class="lbl">{escape(lbl)}</td><td class="val">{val}</td></tr>')

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

    mapa_html = _mapa_svg(r.mapa_puntos)

    # ----- Fotos del avión (exterior / interior) al final -----
    fotos_html = ""
    fotos = [
        ("Exterior", r.foto_exterior),
        ("Interior", r.foto_interior),
    ]
    fotos = [(t, f) for t, f in fotos if f]
    if fotos:
        celdas = "".join(
            f'<figure><img src="{f}" alt=""/><figcaption>{escape(t)}</figcaption></figure>'
            for t, f in fotos
        )
        titulo_av = f" · {escape(r.matricula)}" if r.matricula else ""
        fotos_html = f"""
        <h2>La aeronave{titulo_av}</h2>
        <div class="fotos-grid cols-{len(fotos)}">{celdas}</div>"""

    return f"""<!doctype html>
<html><head><meta charset="utf-8"><style>
  @page {{ size: Letter; margin: 1.8cm 2cm; }}
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
  /* Un CUARTO de la hoja (pedido 26-ago): mitad de ancho, centrado. */
  .mapa {{ width: 50%; margin: 0 auto; border: 1px solid #e5e7eb;
           border-radius: 10px; overflow: hidden; background: #f8fafc; }}
  .mapa svg {{ width: 100%; display: block; }}
  .totales td {{ padding: 7px 0; }}
  .totales .lbl {{ color: #6b7280; }}
  .totales .val {{ text-align: right; font-weight: 600; }}
  .total-row td {{ border-top: 2px solid {_NAVY}; padding-top: 12px; font-size: 18px;
                   font-weight: 800; color: {_BRAND}; }}
  .total-mxn td {{ font-size: 13px; font-weight: 700; color: {_NAVY}; }}
  .notas {{ margin-top: 20px; font-size: 12px; color: #374151; }}
  /* Página 2 (26-ago): los DETALLES del vuelo (mapa, traslados, itinerario
     y fotos) van en su propia página; la página 1 es solo la cotización. */
  .detalles {{ page-break-before: always; }}
  .fotos-grid figure {{ page-break-inside: avoid; }}
  .fotos-grid {{ display: flex; gap: 12px; }}
  .fotos-grid figure {{ margin: 0; flex: 1; }}
  .fotos-grid.cols-1 figure {{ flex: none; width: 100%; }}
  .fotos-grid img {{ width: 100%; border-radius: 10px; border: 1px solid #e5e7eb; }}
  .fotos-grid figcaption {{ font-size: 11px; color: #6b7280; margin-top: 4px;
                            text-align: center; }}
  .footer {{ margin-top: 28px; font-size: 11px; color: #9ca3af; text-align: center; }}
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
    {r.pasajeros} {'pasajero' if r.pasajeros == 1 else 'pasajeros'}{f" · {escape(r.matricula)}" if r.matricula else ""}
  </div>

  <h2>Desglose</h2>
  <table class="totales"><tbody>
    {desglose_html}
    {total_row_html}
    {total_mxn_html}
  </tbody></table>
  {notas_html}

  <div class="detalles">
    <h2>Detalles del vuelo</h2>
    {mapa_html}
    <h2>Traslados</h2>
    <table class="grid"><tbody>
      <tr><td>Traslado inicial</td><td>{_fecha_legible(r.fecha_traslado_inicial)}</td></tr>
      <tr><td>Traslado final</td><td>{_fecha_legible(r.fecha_traslado_final)}</td></tr>
    </tbody></table>
    {escalas_html}
    {fotos_html}
  </div>

  <div class="footer">{TZ_NOTA}<br>Gracias por volar con VuelaTour — Aero Charter Cancún.</div>
</body></html>"""


def render_cotizacion_pdf(req: CotizacionPdfRequest) -> bytes:
    from weasyprint import HTML  # import perezoso

    return HTML(string=_build_html(req)).write_pdf()
