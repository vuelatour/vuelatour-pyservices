"""Genera el PDF del RECIBO DE PAGO de un cobro (documento NO fiscal).

Clona el patrón visual de cotizacion_pdf.py: membrete navy con el logo
blanco, rojo de marca, marca de agua con el logo, hoja Letter con pie
www.vuelatour.com. El API (vuelatour-api) manda TODO ya calculado — folio
del recibo, cobrado a la fecha, saldo, liquidado — y aquí SOLO se pinta:
jamás se recalcula dinero (regla de este microservicio).

El import de WeasyPrint es perezoso para que el servicio arranque aunque la
librería (y sus libs de sistema: pango/cairo) no esté instalada.
"""

import base64
from datetime import datetime
from functools import lru_cache
from html import escape
from pathlib import Path
from zoneinfo import ZoneInfo

from app.schemas.reportes import ReciboPdfRequest

_BRAND = "#dc2626"
_NAVY = "#102a43"
# Todo se muestra en hora de Cancún (Quintana Roo, UTC−5, sin horario de verano).
_CANCUN = ZoneInfo("America/Cancun")

_ASSETS = Path(__file__).resolve().parent.parent / "assets"


def _money(v: float) -> str:
    # Los reembolsos (negativos) se pintan con el signo "menos" tipográfico.
    return f"−${abs(v):,.2f}" if v < 0 else f"${v:,.2f}"


def _fecha_legible(s: str | None, con_hora: bool = True) -> str:
    """ISO → dd/mm/aaaa [HH:MM] en hora Cancún.

    Una fecha PURA (sin componente de hora) se formatea tal cual, sin
    conversión de zona: tratarla como medianoche UTC la movería al día
    anterior en Cancún (regla de fechas del workspace).
    """
    if not s:
        return "—"
    try:
        if "T" not in s and " " not in s:
            return datetime.strptime(s[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
        dt = datetime.fromisoformat(s.replace("Z", "+00:00"))
        # Si viene sin zona, se asume UTC; luego se convierte a Cancún.
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=ZoneInfo("UTC"))
        dt = dt.astimezone(_CANCUN)
        return dt.strftime("%d/%m/%Y %H:%M") if con_hora else dt.strftime("%d/%m/%Y")
    except ValueError:
        return s


@lru_cache(maxsize=2)
def _logo_data_uri(nombre: str) -> str | None:
    """Logo como data-URI (el HTML del PDF no puede pedir archivos remotos)."""
    try:
        raw = (_ASSETS / nombre).read_bytes()
        return "data:image/png;base64," + base64.b64encode(raw).decode()
    except OSError:
        return None


def _build_html(r: ReciboPdfRequest) -> str:
    # ----- Marca de agua + logo del membrete (mismo patrón que cotización) -----
    logo_blanco = _logo_data_uri("logo-vuelatour-blanco.png")
    logo_marca = _logo_data_uri("logo-vuelatour.png")
    logo_header_html = f'<img class="logo" src="{logo_blanco}" alt=""/>' if logo_blanco else ""
    marca_html = f'<div class="marca"><img src="{logo_marca}" alt=""/></div>' if logo_marca else ""

    # ----- Bloque del pago: monto grande en SU moneda -----
    tc_html = ""
    if r.moneda == "MXN" and (r.tc_usd_mxn or r.equivalente_usd is not None):
        partes: list[str] = []
        if r.tc_usd_mxn:
            partes.append(f"T.C. {r.tc_usd_mxn:g}")
        if r.equivalente_usd is not None:
            partes.append(f"equivale a {_money(r.equivalente_usd)} USD")
        tc_html = f'<div class="pago-tc">{escape(" · ".join(partes))}</div>'

    detalle_pago: list[tuple[str, str]] = []
    if r.metodo:
        detalle_pago.append(("Método de pago", r.metodo))
    if r.cuenta_destino:
        detalle_pago.append(("Cuenta destino", r.cuenta_destino))
    if r.referencia:
        detalle_pago.append(("Referencia", r.referencia))
    detalle_pago_html = "".join(
        f'<tr><td class="lbl">{escape(lb)}</td><td class="val">{escape(vl)}</td></tr>'
        for lb, vl in detalle_pago
    )

    # ----- Resumen del vuelo: total / cobrado neto / saldo — o LIQUIDADO -----
    if r.liquidado:
        saldo_html = (
            '<tr class="total-row"><td>Estado de la cuenta</td>'
            '<td class="val"><span class="sello">LIQUIDADO</span></td></tr>'
        )
    else:
        saldo_html = (
            '<tr class="total-row"><td>Saldo pendiente</td>'
            f'<td class="val">{_money(r.saldo_pendiente_usd)} USD</td></tr>'
        )
    resumen_html = f"""
    <h2>Estado de cuenta del vuelo</h2>
    <table class="totales"><tbody>
      <tr><td class="lbl">Total de la cotización</td>
          <td class="val">{_money(r.total_cotizacion_usd)} USD</td></tr>
      <tr><td class="lbl">Cobrado a la fecha (neto)</td>
          <td class="val">{_money(r.cobrado_a_la_fecha_usd)} USD</td></tr>
      {saldo_html}
    </tbody></table>"""

    # ----- Historial de abonos previos (si el API lo manda) -----
    historial_html = ""
    if r.cobros_previos:
        filas = "".join(
            "<tr>"
            f"<td>{_fecha_legible(a.fecha, con_hora=False)}</td>"
            f"<td>{escape(a.etiqueta or 'Abono')}</td>"
            f'<td class="num{" neg" if a.monto < 0 else ""}">'
            f"{_money(a.monto)} {escape(a.moneda)}</td>"
            "</tr>"
            for a in r.cobros_previos
        )
        historial_html = f"""
    <h2>Historial de abonos</h2>
    <table class="grid"><thead>
      <tr><th>Fecha</th><th>Concepto</th><th class="num">Monto</th></tr>
    </thead><tbody>{filas}</tbody></table>"""

    notas_html = (
        f'<div class="notas"><strong>Notas:</strong> {escape(r.notas)}</div>' if r.notas else ""
    )
    sin_tc_html = (
        f'<div class="aviso">AVISO: {escape(r.sin_tc_nota)}</div>' if r.sin_tc_nota else ""
    )

    # Sobre de grupo: el folio ya viene como "G-12" (sin '#') y `ruta` trae

    # el concepto del grupo; un vuelo normal sigue como "Vuelo: #n · ruta".

    if r.grupo_folio:

        referencia_html = (

            f"<strong>Grupo:</strong> {escape(r.grupo_folio)} · {escape(r.ruta)}"

        )

    else:

        referencia_html = (

            f"<strong>Vuelo:</strong> #{escape(r.vuelo_folio)} · {escape(r.ruta)}"

        )


    fecha_vuelo_html = (
        f"<br><strong>Fecha del vuelo:</strong> {_fecha_legible(r.fecha_vuelo, con_hora=False)}"
        if r.fecha_vuelo
        else ""
    )

    return f"""<!doctype html>
<html><head><meta charset="utf-8"><style>
  @page {{
    size: Letter;
    margin: 1.8cm 2cm 2.2cm;
    @bottom-center {{
      content: "www.vuelatour.com";
      font-size: 10px;
      color: #9ca3af;
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
  .folio {{ font-size: 26px; font-weight: 800; color: {_NAVY}; margin: 18px 0 2px; }}
  .meta {{ display: flex; justify-content: space-between; margin: 10px 0 14px;
           font-size: 13px; }}
  h2 {{ font-size: 12px; text-transform: uppercase; letter-spacing: 1.2px; color: #6b7280;
        margin: 20px 0 8px; border-bottom: 1px solid #e5e7eb; padding-bottom: 4px; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
  .grid th, .grid td {{ border: 1px solid #e5e5e5; padding: 6px 10px; text-align: left; }}
  .grid th {{ background: #f7f7f8; }}
  .grid .num {{ text-align: right; }}
  .grid .neg {{ color: {_BRAND}; }}
  /* Tarjeta del pago: monto protagonista en su moneda. */
  .pago {{ background: #f8fafc; border: 1px solid #e5e7eb; border-left: 6px solid {_BRAND};
           border-radius: 10px; padding: 16px 18px; margin-top: 6px; }}
  .pago-lbl {{ font-size: 11px; text-transform: uppercase; letter-spacing: 1.2px;
               color: #6b7280; }}
  .pago-monto {{ font-size: 30px; font-weight: 800; color: {_BRAND}; margin: 2px 0; }}
  .pago-tc {{ font-size: 12px; color: #374151; margin-bottom: 6px; }}
  .pago table {{ margin-top: 8px; }}
  .pago td {{ padding: 4px 0; }}
  .pago .lbl {{ color: #6b7280; width: 40%; }}
  .pago .val {{ font-weight: 600; }}
  .totales td {{ padding: 7px 0; }}
  .totales .lbl {{ color: #6b7280; }}
  .totales .val {{ text-align: right; font-weight: 600; }}
  .total-row td {{ border-top: 2px solid {_NAVY}; padding-top: 12px; font-size: 16px;
                   font-weight: 800; color: {_BRAND}; }}
  .sello {{ display: inline-block; border: 3px solid #15803d; color: #15803d;
            border-radius: 8px; padding: 3px 14px; font-size: 16px; font-weight: 800;
            letter-spacing: 2px; transform: rotate(-3deg); }}
  .aviso {{ margin-top: 14px; background: #fffbeb; border: 1px solid #fcd34d;
            border-radius: 8px; padding: 10px 12px; font-size: 12px; color: #92400e; }}
  .notas {{ margin-top: 14px; font-size: 12px; color: #374151; }}
  .leyenda {{ margin-top: 26px; padding-top: 10px; border-top: 1px solid #e5e7eb;
              font-size: 11px; color: #6b7280; text-align: center; }}
</style></head><body>
  {marca_html}
  <div class="header">
    {logo_header_html}
    <div class="titulos">
      <h1>{escape(r.empresa)}</h1>
      <p>Recibo de pago</p>
    </div>
  </div>

  <div class="folio">Recibo {escape(r.folio_recibo) or "de pago"}</div>
  <div class="meta">
    <div><strong>Cliente:</strong> {escape(r.cliente) or "Cliente"}<br>
      {referencia_html}{fecha_vuelo_html}</div>
    <div style="text-align:right"><strong>Fecha del cobro:</strong>
      {_fecha_legible(r.fecha_cobro)}</div>
  </div>

  <h2>Pago recibido</h2>
  <div class="pago">
    <div class="pago-lbl">Monto recibido</div>
    <div class="pago-monto">{_money(r.monto)} {escape(r.moneda)}</div>
    {tc_html}
    <table><tbody>{detalle_pago_html}</tbody></table>
  </div>

  {resumen_html}
  {historial_html}
  {sin_tc_html}
  {notas_html}

  <div class="leyenda">Este recibo no es un comprobante fiscal (CFDI).</div>
</body></html>"""


def render_recibo_pdf(req: ReciboPdfRequest) -> bytes:
    from weasyprint import HTML  # import perezoso

    return HTML(string=_build_html(req)).write_pdf()
