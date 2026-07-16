"""Vista previa del PDF de una factura SIN timbrar (WeasyPrint, HTML → PDF).

Renderiza el TimbrarRequest tal cual llegaría al PAC para que el operador
revise emisor/receptor/conceptos ANTES de timbrar. NO es un CFDI: no hay
sello ni folio fiscal, y la hoja lo grita (banda + marca de agua) para que
nadie la confunda con una factura real. Misma identidad visual que el
reporte de vuelo (brand #dc2626, navy #102a43).
"""

from html import escape

from app.schemas.facturacion import TimbrarRequest

_BRAND = "#dc2626"
_NAVY = "#102a43"

# Catálogos SAT mínimos para que el operador lea códigos sin memorizarlos.
# Solo display: si llega un código fuera del mapa se muestra el código pelón.
_FORMAS_PAGO = {
    "01": "Efectivo",
    "02": "Cheque",
    "03": "Transferencia",
    "04": "Tarjeta de crédito",
    "28": "Tarjeta de débito",
    "99": "Por definir",
}
_METODOS_PAGO = {
    "PUE": "Pago en una sola exhibición",
    "PPD": "Pago en parcialidades o diferido",
}
_TIPOS_COMPROBANTE = {"I": "Ingreso", "E": "Egreso (nota de crédito)"}


def _money(v: float, moneda: str) -> str:
    return f"${v:,.2f} {escape(moneda)}"


def _codigo(codigo: str, catalogo: dict[str, str]) -> str:
    """"03 — Transferencia" o el código solo si no está en el catálogo."""
    desc = catalogo.get(codigo)
    return f"{escape(codigo)} — {escape(desc)}" if desc else escape(codigo)


def _kv(label: str, value: str) -> str:
    """Fila clave/valor; value ya viene escapado por el llamador."""
    return f'<tr><td class="k">{escape(label)}</td><td class="v">{value}</td></tr>'


def _build_html(req: TimbrarRequest) -> str:
    # --- Conceptos: mismo criterio de IVA que facturama.py (el iva explícito
    # del API manda sobre el recálculo para no descuadrar por centavos). ---
    filas = []
    subtotal_total = 0.0
    iva_total = 0.0
    for c in req.conceptos:
        subtotal = round(c.cantidad * c.valor_unitario, 2)
        iva = c.iva if c.iva is not None else round(subtotal * c.tasa_iva, 2)
        subtotal_total = round(subtotal_total + subtotal, 2)
        iva_total = round(iva_total + iva, 2)
        filas.append(
            f"<tr><td>{escape(c.clave_prod_serv)}</td>"
            f"<td>{escape(c.descripcion)}</td>"
            f"<td class='num'>{c.cantidad:g}</td>"
            f"<td class='num'>{_money(c.valor_unitario, req.moneda)}</td>"
            f"<td class='num'>{_money(iva, req.moneda)}</td>"
            f"<td class='num'>{_money(round(subtotal + iva, 2), req.moneda)}</td></tr>"
        )
    total = round(subtotal_total + iva_total, 2)

    # --- Emisor / Receptor lado a lado ---
    emisor = (
        "<table class='kv'>"
        + _kv("Razón social", escape(req.emisor.nombre))
        + _kv("RFC", escape(req.emisor.rfc))
        + _kv("Régimen fiscal", escape(req.emisor.regimen_fiscal))
        + _kv("Lugar de expedición", escape(req.lugar_expedicion))
        + "</table>"
    )
    receptor = (
        "<table class='kv'>"
        + _kv("Nombre", escape(req.receptor.nombre))
        + _kv("RFC", escape(req.receptor.rfc))
        + _kv("Régimen fiscal", escape(req.receptor.regimen_fiscal))
        + _kv("Uso CFDI", escape(req.receptor.uso_cfdi))
        + _kv("CP (domicilio fiscal)", escape(req.receptor.domicilio_fiscal))
        + "</table>"
    )

    # --- Datos del comprobante ---
    comprobante = "<table class='kv'>"
    comprobante += _kv("Tipo", _codigo(req.tipo_comprobante or "I", _TIPOS_COMPROBANTE))
    if req.serie or req.folio:
        serie_folio = " ".join(p for p in (req.serie, req.folio) if p)
        comprobante += _kv("Serie / Folio", escape(serie_folio))
    comprobante += _kv("Moneda", escape(req.moneda))
    comprobante += _kv("Forma de pago", _codigo(req.forma_pago, _FORMAS_PAGO))
    comprobante += _kv("Método de pago", _codigo(req.metodo_pago, _METODOS_PAGO))
    comprobante += _kv("Referencia", escape(req.referencia))
    comprobante += "</table>"

    avisos = ""
    if req.informacion_global:
        g = req.informacion_global
        avisos += (
            f"<p class='aviso'>Factura global: periodicidad {escape(g.periodicidad)} · "
            f"mes {escape(g.meses)} · año {g.anio}</p>"
        )
    if req.cfdi_relacionado_uuid:
        avisos += (
            f"<p class='aviso'>CFDI relacionado (tipo {escape(req.tipo_relacion or '—')}): "
            f"{escape(req.cfdi_relacionado_uuid)}</p>"
        )

    conceptos_html = (
        "<table class='grid'><thead><tr><th>ClaveProdServ</th><th>Descripción</th>"
        "<th class='num'>Cant.</th><th class='num'>Valor unitario</th>"
        "<th class='num'>IVA</th><th class='num'>Total</th></tr></thead>"
        f"<tbody>{''.join(filas)}</tbody></table>"
    )

    totales = (
        "<table class='totales'>"
        f"<tr><td class='k'>Subtotal</td>"
        f"<td class='num'>{_money(subtotal_total, req.moneda)}</td></tr>"
        f"<tr><td class='k'>IVA</td><td class='num'>{_money(iva_total, req.moneda)}</td></tr>"
        f"<tr class='gran'><td class='k'>TOTAL</td>"
        f"<td class='num'>{_money(total, req.moneda)}</td></tr>"
        "</table>"
    )

    return f"""<!doctype html>
<html><head><meta charset="utf-8"><style>
  @page {{ size: letter; margin: 28px 34px; }}
  * {{ font-family: 'Helvetica Neue', Arial, sans-serif; color: {_NAVY}; }}
  /* Marca de agua diagonal: fixed para que cruce TODAS las páginas. */
  .marca {{ position: fixed; top: 42%; left: -6%; width: 115%;
            transform: rotate(-32deg); font-size: 92px; font-weight: 800;
            color: {_BRAND}; opacity: 0.08; text-align: center;
            letter-spacing: 6px; z-index: 0; }}
  .banda {{ background: {_BRAND}; color: #fff; text-align: center;
            font-size: 13px; font-weight: 800; letter-spacing: 1px;
            padding: 7px 10px; margin-bottom: 12px; }}
  .banda * {{ color: #fff; }}
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
  table.kv td.k {{ color:#627d98; width:42%; }}
  table.grid th, table.grid td {{ border:1px solid #e1e8f0; padding:4px 6px; text-align:left; }}
  table.grid th {{ background:#f0f4f8; font-size:10px; text-transform:uppercase; }}
  .num {{ text-align:right; }}
  .cols {{ display:flex; gap:24px; }}
  .cols > div {{ flex:1; }}
  .aviso {{ font-size:11px; color:#627d98; margin:4px 0; }}
  table.totales {{ width:44%; margin-left:auto; margin-top:10px; font-size:12px; }}
  table.totales td {{ padding:4px 6px; }}
  table.totales td.k {{ color:#627d98; text-align:right; }}
  table.totales tr.gran td {{ font-size:16px; font-weight:800;
        border-top:2px solid {_BRAND}; color:{_NAVY}; }}
  .pie {{ margin-top:22px; border-top:1px solid #e1e8f0; padding-top:8px;
          font-size:10px; color:#829ab1; }}
</style></head><body>
  <div class="marca">VISTA PREVIA</div>
  <div class="banda">VISTA PREVIA — SIN VALIDEZ FISCAL</div>
  <div class="head">
    <div><div class="brand">VuelaTour</div><div class="sub">Aero Charter Cancún</div></div>
    <div style="text-align:right"><h1>Vista previa de factura</h1>
      <div class="sub">Documento SIN timbrar · sin sello ni folio fiscal</div></div>
  </div>
  <div class="cols">
    <div><h2>Emisor</h2>{emisor}</div>
    <div><h2>Receptor</h2>{receptor}</div>
  </div>
  <h2>Comprobante</h2>
  {comprobante}
  {avisos}
  <h2>Conceptos</h2>
  {conceptos_html}
  {totales}
  <p class="pie">Este documento es una vista previa generada por VuelaTour.
  NO es un CFDI; no tiene sello ni folio fiscal.</p>
</body></html>"""


def render_factura_preview_pdf(req: TimbrarRequest) -> bytes:
    from weasyprint import HTML  # import perezoso: la lib nativa no está en todas las máquinas

    return HTML(string=_build_html(req)).write_pdf()
