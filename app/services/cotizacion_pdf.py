"""Genera el PDF de una cotización con WeasyPrint (HTML → PDF).

Mantiene la identidad visual del admin: rojo de marca #dc2626, navy #102a43.
El import de WeasyPrint es perezoso para que el servicio arranque aunque la
librería (y sus libs de sistema: pango/cairo) no esté instalada.
"""

from datetime import datetime
from html import escape
from zoneinfo import ZoneInfo

from app.schemas.reportes import CotizacionPdfRequest

_BRAND = "#dc2626"
_NAVY = "#102a43"
# Todo se muestra en hora de Cancún (Quintana Roo, UTC−5, sin horario de verano).
_CANCUN = ZoneInfo("America/Cancun")
TZ_NOTA = "Horarios en hora de Cancún (UTC−5)."


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


def _build_html(r: CotizacionPdfRequest) -> str:
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

    tiempo = f"{r.tiempo_cobrable_hr:.1f} hr" if r.tiempo_cobrable_hr is not None else "—"
    tarifa = _money(r.tarifa_hora_usd) if r.tarifa_hora_usd is not None else "—"
    notas_html = (
        f'<div class="notas"><strong>Notas:</strong> {escape(r.notas)}</div>' if r.notas else ""
    )

    # ----- Desglose: cada renglón que compone el total (mismo orden que el
    # motor). Los TUAS se pintan POR AEROPUERTO con su moneda (requisito del
    # cliente); sin detalle (cotización vieja) se conserva la línea única. -----
    filas: list[str] = []

    def fila(lbl: str, val: str) -> None:
        filas.append(f'<tr><td class="lbl">{escape(lbl)}</td><td class="val">{val}</td></tr>')

    fila("Tiempo cobrable", tiempo)
    fila("Tarifa por hora", tarifa)
    fila("Subtotal", _money(r.subtotal_usd))
    if not r.tuas_detalle:
        fila("TUAS", _money(r.tuas_usd))
    elif len(r.tuas_detalle) == 1:
        # Caso común (un aeropuerto): el concepto detallado sustituye a "TUAS".
        fila(r.tuas_detalle[0], _money(r.tuas_usd))
    else:
        # Varios aeropuertos: detalle por línea + total USD para que la columna
        # de montos siga sumando el total exacto.
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

    # Línea final en pesos: total MXN EXACTO por composición (viene calculado
    # del motor; aquí solo se muestra) con el TC congelado de la cotización.
    total_mxn_html = ""
    if r.total_mxn is not None:
        tc_txt = f" (T.C. {r.tc_usd_mxn:g})" if r.tc_usd_mxn else ""
        total_mxn_html = (
            f'<tr class="total-mxn"><td>Total MXN{escape(tc_txt)}</td>'
            f'<td class="val">{_money(r.total_mxn)} MXN</td></tr>'
        )

    return f"""<!doctype html>
<html><head><meta charset="utf-8"><style>
  @page {{ size: Letter; margin: 2cm; }}
  * {{ font-family: 'Helvetica Neue', Arial, sans-serif; color: #1d1d1d; }}
  .header {{ background: {_NAVY}; color: #fff; padding: 18px 24px; border-radius: 10px; }}
  .header h1 {{ margin: 0; font-size: 20px; }}
  .header p {{ margin: 2px 0 0; color: #9fb3c8; font-size: 12px; }}
  .meta {{ display: flex; justify-content: space-between; margin: 20px 0; font-size: 13px; }}
  .route {{ font-size: 26px; font-weight: 800; color: {_NAVY}; margin: 8px 0 4px; }}
  h2 {{ font-size: 13px; text-transform: uppercase; letter-spacing: 1px; color: #6b7280;
        margin: 22px 0 8px; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
  .grid th, .grid td {{ border: 1px solid #e5e5e5; padding: 6px 10px; text-align: left; }}
  .grid th {{ background: #f7f7f8; }}
  .totales td {{ padding: 7px 0; }}
  .totales .lbl {{ color: #6b7280; }}
  .totales .val {{ text-align: right; font-weight: 600; }}
  .total-row td {{ border-top: 2px solid {_NAVY}; padding-top: 12px; font-size: 18px;
                   font-weight: 800; color: {_BRAND}; }}
  .total-mxn td {{ font-size: 13px; font-weight: 700; color: {_NAVY}; }}
  .notas {{ margin-top: 22px; font-size: 12px; color: #374151; }}
  .footer {{ margin-top: 30px; font-size: 11px; color: #9ca3af; text-align: center; }}
</style></head><body>
  <div class="header">
    <h1>{escape(r.empresa)}</h1>
    <p>Cotización de servicio aéreo</p>
  </div>

  <div class="meta">
    <div><strong>Folio:</strong> #{escape(r.folio)}<br><strong>Cliente:</strong> {escape(r.cliente)}</div>
    <div style="text-align:right"><strong>Fecha:</strong> {_fecha_legible(r.fecha)}<br>
      <strong>Tipo:</strong> {escape(r.tipo)}</div>
  </div>

  <div class="route">{escape(r.origen)} → {escape(r.destino)}</div>
  <div style="font-size:13px;color:#374151">
    {r.pasajeros} {'pasajero' if r.pasajeros == 1 else 'pasajeros'}
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

  <div class="footer">{TZ_NOTA}<br>Gracias por volar con VuelaTour — Aero Charter Cancún.</div>
</body></html>"""


def render_cotizacion_pdf(req: CotizacionPdfRequest) -> bytes:
    from weasyprint import HTML  # import perezoso

    return HTML(string=_build_html(req)).write_pdf()
