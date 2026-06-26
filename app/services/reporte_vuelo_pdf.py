"""Reporte consolidado de UN vuelo (WeasyPrint, HTML → PDF).

Reúne en una hoja: datos del vuelo, cotización, ingreso (cobros), tacómetro por
tramo, combustible y gastos. Misma identidad visual que la cotización.
"""

from datetime import datetime
from html import escape
from zoneinfo import ZoneInfo

from app.schemas.reportes import ReporteVueloRequest

_BRAND = "#dc2626"
_NAVY = "#102a43"
_CANCUN = ZoneInfo("America/Cancun")


def _money(v: float | None, moneda: str = "USD") -> str:
    if v is None:
        return "—"
    return f"${v:,.2f} {moneda}".rstrip()


def _fecha(s: str | None) -> str:
    if not s:
        return "—"
    try:
        dt = datetime.fromisoformat(s.replace("Z", "+00:00"))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=ZoneInfo("UTC"))
        return dt.astimezone(_CANCUN).strftime("%d/%m/%Y %H:%M")
    except ValueError:
        return s


def _row(label: str, value: str) -> str:
    return (
        f'<tr><td class="k">{escape(label)}</td>'
        f'<td class="v">{value}</td></tr>'
    )


def _seccion(titulo: str, cuerpo: str) -> str:
    return f'<h2>{escape(titulo)}</h2>{cuerpo}'


def _build_html(r: ReporteVueloRequest) -> str:
    # --- Resumen del vuelo ---
    resumen = "<table class='kv'>"
    resumen += _row("Folio", f"#{escape(r.folio)}")
    resumen += _row("Cliente", escape(r.cliente or "—"))
    resumen += _row("Ruta", escape(r.ruta or "—"))
    resumen += _row("Tipo / Estado", escape(f"{r.tipo} · {r.estado}".strip(" ·")))
    resumen += _row("Aeronave", escape(r.aeronave or "—"))
    piloto = r.piloto or "—"
    if r.copiloto:
        piloto += f" / {r.copiloto} (copiloto)"
    resumen += _row("Piloto", escape(piloto))
    resumen += _row("Pasajeros", str(r.pasajeros))
    resumen += _row("Fecha de vuelo", _fecha(r.fecha_vuelo))
    resumen += _row("Traslado final", _fecha(r.fecha_traslado_final))
    resumen += "</table>"

    # --- Cotización ---
    cot = "<table class='kv'>"
    if r.tarifa_tipo:
        cot += _row("Tarifa", escape(r.tarifa_tipo))
    if r.tarifa_hora_usd:
        cot += _row("USD / hora", _money(r.tarifa_hora_usd))
    if r.tiempo_cobrable_hr:
        cot += _row("Tiempo cobrable", f"{r.tiempo_cobrable_hr:.2f} hr")
    cot += _row("Subtotal", _money(r.subtotal_usd))
    cot += _row("TUAS", _money(r.tuas_usd))
    if r.extras_total_usd:
        cot += _row("Extras", _money(r.extras_total_usd))
    if r.ajuste_final_usd:
        cot += _row("Ajuste", _money(r.ajuste_final_usd))
    cot += _row("IVA", _money(r.iva_usd))
    cot += _row("<b>Total</b>", f"<b>{_money(r.total_usd)}</b>")
    if r.total_mxn:
        tc = f" (TC {r.tc_usd_mxn:.2f})" if r.tc_usd_mxn else ""
        cot += _row("Total MXN", f"{_money(r.total_mxn, 'MXN')}{escape(tc)}")
    if r.metodo_cobro:
        cot += _row("Método de cobro", escape(r.metodo_cobro))
    cot += "</table>"

    # --- Ingreso (cobros) ---
    if r.cobros:
        filas = "".join(
            f"<tr><td>{_fecha(c.fecha)}</td><td>{escape(c.concepto or '—')}</td>"
            f"<td class='num'>{_money(c.monto, c.moneda or 'USD')}</td></tr>"
            for c in r.cobros
        )
        ingreso = (
            "<table class='grid'><thead><tr><th>Fecha</th><th>Método</th>"
            f"<th class='num'>Monto</th></tr></thead><tbody>{filas}</tbody></table>"
        )
    else:
        ingreso = "<p class='muted'>Sin cobros registrados.</p>"
    ingreso += (
        f"<p class='tot'>Cobrado: <b>{_money(r.total_cobrado_usd)}</b> · "
        f"Saldo: <b>{_money(r.saldo_usd)}</b></p>"
    )

    # --- Tacómetro por tramo ---
    if r.tramos:
        filas = "".join(
            f"<tr><td>{t.orden}</td><td>{escape(t.ruta)}</td>"
            f"<td class='num'>{'' if t.taco_salida is None else f'{t.taco_salida:.1f}'}</td>"
            f"<td class='num'>{'' if t.taco_llegada is None else f'{t.taco_llegada:.1f}'}</td>"
            f"<td class='num'>{'' if t.horas is None else f'{t.horas:.1f}'}</td></tr>"
            for t in r.tramos
        )
        taco = (
            "<table class='grid'><thead><tr><th>#</th><th>Tramo</th>"
            "<th class='num'>Salida</th><th class='num'>Llegada</th>"
            f"<th class='num'>Horas</th></tr></thead><tbody>{filas}</tbody></table>"
        )
    else:
        taco = "<p class='muted'>Sin lecturas de tacómetro.</p>"

    # --- Combustible ---
    if r.combustible:
        filas = "".join(
            f"<tr><td>{_fecha(c.fecha)}</td><td>{escape(c.detalle or c.concepto or '—')}</td>"
            f"<td class='num'>{_money(c.monto, c.moneda or 'MXN')}</td></tr>"
            for c in r.combustible
        )
        comb = (
            "<table class='grid'><thead><tr><th>Fecha</th><th>Detalle</th>"
            f"<th class='num'>Monto</th></tr></thead><tbody>{filas}</tbody></table>"
        )
    else:
        comb = "<p class='muted'>Sin cargas de combustible.</p>"

    # --- Gastos ---
    if r.gastos:
        filas = "".join(
            f"<tr><td>{_fecha(g.fecha)}</td><td>{escape(g.concepto or '—')}</td>"
            f"<td>{escape(g.detalle or '')}</td>"
            f"<td class='num'>{_money(g.monto, g.moneda or 'MXN')}</td></tr>"
            for g in r.gastos
        )
        gastos = (
            "<table class='grid'><thead><tr><th>Fecha</th><th>Categoría</th>"
            "<th>Proveedor</th><th class='num'>Monto</th></tr></thead>"
            f"<tbody>{filas}</tbody></table>"
        )
    else:
        gastos = "<p class='muted'>Sin gastos registrados.</p>"

    generado = _fecha(r.generado)
    notas = f"<p class='muted'>{escape(r.notas)}</p>" if r.notas else ""

    return f"""<!doctype html>
<html><head><meta charset="utf-8"><style>
  @page {{ size: letter; margin: 28px 34px; }}
  * {{ font-family: 'Helvetica Neue', Arial, sans-serif; color: {_NAVY}; }}
  .head {{ display:flex; justify-content:space-between; align-items:flex-start;
           border-bottom:3px solid {_BRAND}; padding-bottom:8px; margin-bottom:14px; }}
  .brand {{ font-size:22px; font-weight:800; color:{_BRAND}; }}
  .sub {{ font-size:11px; color:#627d98; }}
  h1 {{ font-size:16px; margin:0; }}
  h2 {{ font-size:12px; text-transform:uppercase; letter-spacing:.5px;
        color:{_BRAND}; border-bottom:1px solid #e1e8f0; padding-bottom:3px;
        margin:16px 0 6px; }}
  table {{ width:100%; border-collapse:collapse; font-size:11px; }}
  table.kv td {{ padding:3px 4px; vertical-align:top; }}
  table.kv td.k {{ color:#627d98; width:40%; }}
  table.grid th, table.grid td {{ border:1px solid #e1e8f0; padding:4px 6px; text-align:left; }}
  table.grid th {{ background:#f0f4f8; font-size:10px; text-transform:uppercase; }}
  .num {{ text-align:right; }}
  .muted {{ color:#829ab1; font-size:11px; }}
  .tot {{ font-size:12px; margin-top:6px; }}
  .cols {{ display:flex; gap:24px; }}
  .cols > div {{ flex:1; }}
</style></head><body>
  <div class="head">
    <div><div class="brand">VuelaTour</div><div class="sub">Aero Charter Cancún</div></div>
    <div style="text-align:right"><h1>Reporte de vuelo #{escape(r.folio)}</h1>
      <div class="sub">Generado {generado} · hora de Cancún (UTC−5)</div></div>
  </div>
  <div class="cols">
    <div>{_seccion("Resumen del vuelo", resumen)}</div>
    <div>{_seccion("Cotización", cot)}</div>
  </div>
  {_seccion("Ingreso (cobros)", ingreso)}
  {_seccion("Tacómetro por tramo", taco)}
  {_seccion("Combustible", comb)}
  {_seccion("Gastos", gastos)}
  {notas}
</body></html>"""


def render_reporte_vuelo_pdf(req: ReporteVueloRequest) -> bytes:
    from weasyprint import HTML  # import perezoso

    return HTML(string=_build_html(req)).write_pdf()
