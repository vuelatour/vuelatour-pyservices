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
        # Fecha SOLO-DÍA (fecha_gasto: "2026-08-21"): es un día de pared, NO
        # un instante — convertirla a Cancún la corría al día anterior.
        if len(s) == 10 and "T" not in s:
            return datetime.fromisoformat(s).strftime("%d/%m/%Y")
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
    pax = str(r.pasajeros)
    if r.pasajeros_nombres:
        pax += f" — {r.pasajeros_nombres}"
    resumen += _row("Pasajeros", escape(pax))
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
    # Detalle de TUAS por aeropuerto CON su moneda (requisito del cliente):
    # sub-líneas informativas — la fila numérica de arriba cuadra la suma.
    for det in r.tuas_detalle:
        cot += (
            "<tr><td colspan='2' class='muted' "
            f"style='font-size:10px;padding:0 4px 3px 14px'>{escape(det)}</td></tr>"
        )
    if r.viaticos_pernocta_usd:
        cot += _row("Pernocta (vi&aacute;ticos)", _money(r.viaticos_pernocta_usd))
    if r.extras_total_usd:
        cot += _row("Extras", _money(r.extras_total_usd))
    if r.ajuste_final_usd:
        cot += _row("Ajuste", _money(r.ajuste_final_usd))
    cot += _row("IVA", _money(r.iva_usd))
    cot += _row("<b>Total</b>", f"<b>{_money(r.total_usd)}</b>")
    if r.total_mxn:
        tc = f" (TC {r.tc_usd_mxn:.2f})" if r.tc_usd_mxn else ""
        cot += _row("Total MXN", f"{_money(r.total_mxn, 'MXN')}{escape(tc)}")
    # Comisión del vendedor (interna): el cliente paga el total completo; el
    # neto es lo que queda a VuelaTour (lo que fluye al reparto).
    if r.comision_vendedor_usd:
        quien = f" ({r.comision_vendedor_nombre})" if r.comision_vendedor_nombre else ""
        cot += _row(
            f"Comisi&oacute;n vendedor{escape(quien)}",
            f"&minus;{_money(r.comision_vendedor_usd)}",
        )
        neto = (
            r.neto_vuelatour_usd
            if r.neto_vuelatour_usd is not None
            else r.total_usd - r.comision_vendedor_usd
        )
        cot += _row("<b>Neto VuelaTour</b>", f"<b>{_money(neto)}</b>")
    if r.metodo_cobro:
        cot += _row("Método de cobro", escape(r.metodo_cobro))
    cot += "</table>"

    # --- Ingreso (cobros) ---
    if r.cobros:
        filas = "".join(
            f"<tr><td>{_fecha(c.fecha)}</td><td>{escape(c.concepto or '—')}"
            # Comisión bancaria del cobro (el banco depositó monto − comisión).
            + (
                f"<br/><span class='muted' style='font-size:10px'>{escape(c.detalle)}</span>"
                if c.detalle
                else ""
            )
            + "</td>"
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
    # Total ANTES de comisión bancaria vs neto DESPUÉS (lo que entró al banco).
    if r.comision_banco_usd:
        neto_banco = (
            r.total_cobrado_neto_usd
            if r.total_cobrado_neto_usd is not None
            else r.total_cobrado_usd - r.comision_banco_usd
        )
        ingreso += (
            f"<p class='tot'>Comisiones bancarias: <b>&minus;{_money(r.comision_banco_usd)}</b> · "
            f"Neto recibido (después de comisión): <b>{_money(neto_banco)}</b></p>"
        )

    # --- Tacómetro por tramo ---
    if r.tramos:
        filas = "".join(
            f"<tr><td>{t.orden}</td>"
            f"<td>{escape(t.ruta)}"
            + (f"<br/><span class='muted' style='font-size:10px'>{escape(t.pasajeros_nombres)}</span>" if t.pasajeros_nombres else "")
            + "</td>"
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

    # --- Horas cotizadas vs voladas (utilidad operativa) ---
    if r.horas_cotizadas_hr is not None or r.horas_voladas_hr is not None:
        delta = r.horas_delta_hr
        color = "#16a34a" if (delta or 0) > 0 else ("#dc2626" if (delta or 0) < 0 else "#64748b")
        signo = "+" if (delta or 0) > 0 else ""
        fmt = lambda v: "—" if v is None else f"{v:.1f} hrs"
        notas_html = (
            "<ul style='margin:6px 0 0 16px;padding:0'>"
            + "".join(f"<li class='muted' style='font-size:11px'>{escape(nt)}</li>" for nt in r.notas_horas)
            + "</ul>"
            if r.notas_horas
            else ""
        )
        horas_cmp = (
            "<table class='grid'><thead><tr>"
            "<th class='num'>Horas cotizadas</th><th class='num'>Horas voladas</th>"
            "<th class='num'>Diferencia</th></tr></thead><tbody><tr>"
            f"<td class='num'>{fmt(r.horas_cotizadas_hr)}</td>"
            f"<td class='num'>{fmt(r.horas_voladas_hr)}</td>"
            f"<td class='num' style='color:{color};font-weight:700'>"
            f"{'—' if delta is None else f'{signo}{delta:.1f} hrs'}</td>"
            "</tr></tbody></table>" + notas_html
        )
    else:
        horas_cmp = ""

    # --- Combustible (con litros y $/litro, como el control del equipo) ---
    if r.combustible:
        def _litros(c):
            if not c.litros:
                return "<td class='num'>—</td><td class='num'>—</td>"
            xl = (
                f"{_money(round((c.monto or 0) / c.litros, 2), c.moneda or 'MXN')}"
                if c.monto
                else "—"
            )
            return f"<td class='num'>{c.litros:g} L</td><td class='num'>{xl}</td>"

        filas = "".join(
            f"<tr><td>{_fecha(c.fecha)}</td><td>{escape(c.detalle or c.concepto or '—')}</td>"
            + _litros(c)
            + f"<td class='num'>{_money(c.monto, c.moneda or 'MXN')}</td></tr>"
            for c in r.combustible
        )
        comb = (
            "<table class='grid'><thead><tr><th>Fecha</th><th>Detalle</th>"
            "<th class='num'>Litros</th><th class='num'>$ x litro</th>"
            f"<th class='num'>Monto</th></tr></thead><tbody>{filas}</tbody></table>"
        )
        if r.combustible_total_usd:
            comb += f"<p class='tot'>Total gasolina: <b>{_money(r.combustible_total_usd)}</b></p>"
        comb += (
            "<p class='muted'>El combustible se controla por AVIÓN y por MES "
            "(Balance por avión → hoja 'combustible'); aquí solo se muestran "
            "las cargas ligadas a este vuelo.</p>"
        )
    else:
        comb = (
            "<p class='muted'>Sin cargas ligadas a este vuelo. El combustible "
            "se controla por AVIÓN y por MES (Balance por avión → hoja "
            "'combustible').</p>"
        )

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
    if r.gastos and r.gastos_total_usd:
        gastos += f"<p class='tot'>Total gastos: <b>{_money(r.gastos_total_usd)}</b></p>"
    if r.gastos_sin_tc_count:
        gastos += (
            f"<p class='muted' style='color:{_BRAND}'>OJO: {r.gastos_sin_tc_count} gasto(s) en MXN "
            f"por ${r.gastos_sin_tc_mxn:,.2f} SIN tipo de cambio: no entran al balance USD.</p>"
        )

    # --- Balance del vuelo (remanente → ganancia → %, como el Excel del
    # equipo). Los montos vienen calculados del API; el MXN es equivalente
    # al TC del vuelo (solo display). ---
    tc = r.tc_usd_mxn if (r.tc_usd_mxn or 0) > 0 else None

    def _signed(v: float, moneda: str = "USD") -> str:
        # Signo real en el número (no solo color): una ganancia negativa debe
        # verse negativa, igual que en el XLSX del mismo vuelo.
        return ("&minus;" if v < 0 else "") + _money(abs(v), moneda)

    def _bal_row(label: str, usd, bold=False, color=None, signo=1):
        if usd is None:
            return ""
        val = signo * usd  # las deducciones llegan positivas con signo=-1
        style = f"font-weight:{'700' if bold else '400'};" + (f"color:{color};" if color else "")
        mxn = (
            f"<td class='num' style='{style}'>{_signed(round(val * tc, 2), 'MXN')}</td>"
            if tc
            else ""
        )
        return (
            f"<tr><td style='{style}'>{escape(label)}</td>"
            f"<td class='num' style='{style}'>{_signed(val)}</td>{mxn}</tr>"
        )

    if r.remanente_usd is not None or r.ganancia_final_usd is not None:
        venta_sin_iva = r.venta_sin_iva_usd or max(r.total_usd - r.iva_usd, 0)
        g_color = "#16a34a" if (r.ganancia_final_usd or 0) >= 0 else "#dc2626"
        balance_rows = (
            _bal_row("Venta total (c/IVA)", r.total_usd, bold=True)
            + _bal_row("Venta sin IVA", venta_sin_iva)
            + _bal_row("(−) Gasolina", r.combustible_total_usd or 0, signo=-1)
            + _bal_row("(−) Gastos del vuelo", r.gastos_total_usd or 0, signo=-1)
            + _bal_row("REMANENTE (venta − costo)", r.remanente_usd, bold=True)
            + (
                _bal_row(
                    "(−) Comisión vendedor"
                    + (f" ({r.comision_vendedor_nombre})" if r.comision_vendedor_nombre else ""),
                    r.comision_vendedor_usd,
                    signo=-1,
                )
                if r.comision_vendedor_usd
                else ""
            )
            + (
                _bal_row("(−) Comisiones bancarias", r.comision_banco_usd, signo=-1)
                if r.comision_banco_usd
                else ""
            )
            + _bal_row("GANANCIA DEL VUELO", r.ganancia_final_usd, bold=True, color=g_color)
            + _bal_row("Ganancia x hora", r.ganancia_x_hr_usd)
        )
        pct_html = ""
        if r.ganancia_pct is not None:
            pct_html = (
                f"<tr><td style='font-weight:700'>% de ganancia (sobre venta s/IVA)</td>"
                f"<td class='num' style='font-weight:700;color:{g_color}'>"
                f"{r.ganancia_pct * 100:.1f}%</td>"
                + ("<td></td>" if tc else "")
                + "</tr>"
            )
        head_mxn = f"<th class='num'>MXN (T.C. {tc:g})</th>" if tc else ""
        balance = (
            "<table class='grid'><thead><tr><th>Concepto</th>"
            f"<th class='num'>USD</th>{head_mxn}</tr></thead>"
            f"<tbody>{balance_rows}{pct_html}</tbody></table>"
        )
    else:
        balance = ""

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
  {_seccion("Horas cotizadas vs voladas", horas_cmp) if horas_cmp else ""}
  {_seccion("Combustible", comb)}
  {_seccion("Gastos", gastos)}
  {_seccion("Balance del vuelo", balance) if balance else ""}
  {notas}
</body></html>"""


def render_reporte_vuelo_pdf(req: ReporteVueloRequest) -> bytes:
    from weasyprint import HTML  # import perezoso

    return HTML(string=_build_html(req)).write_pdf()
