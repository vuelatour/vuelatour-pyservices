"""PDF de la COTIZACIÓN DE GRUPO (4-sep-2026): varios aviones para un mismo
cliente, UN documento y UN total.

Hermano de `cotizacion_pdf.py` y REUTILIZA (importa, no copia) su branding
(#dc2626 / #102a43), estilos, mapa de ruta (modo local/amplio), formato de
fechas, la tabla de itinerario y la ficha "De un vistazo" de la aeronave.

Aquí SOLO se pinta: el consolidado, los totales y el precio por persona
llegan YA calculados del API (Σ de los desgloses canónicos de los hijos
vivos) — jamás se recalcula dinero en pyservices.

Páginas:
  1. "Cotización de grupo G-n": cliente, fecha, ruta, "Grupo de N pasajeros
     · N aeronaves", itinerario (columna Fecha solo si algún tramo la trae)
     junto al mapa, desglose consolidado, subtotal / IVA / TOTAL USD (+ MXN
     y T.C.), precio por persona (toggle) y notas.
  2. "Flota asignada" (toggle `mostrar_anexo_aviones`): una fila por avión;
     subtotal y tarifa solo con sus toggles.
  3. "Las aeronaves": una hoja por MODELO distinto con foto y ficha.

Reglas del cliente heredadas del PDF de un avión:
  - Matrícula OCULTA salvo VGV (`_mostrar_matricula`, fuente única); en la
    hoja de la aeronave nunca se imprime.
  - NUNCA se pinta la comisión del vendedor ni el redondeo hacia arriba
    (el API ya los absorbe en "Servicio aéreo"; aquí se filtra por clave
    por si acaso).
  - Sin horas de vuelo ni tarifa por hora salvo `mostrar_tarifa`.
  - TUAS (5-sep-2026, «que se vea sutilmente la operación»): cuando la
    línea trae cantidad (pax gravados) y unitario nativo se pinta
    «TUA CZA · 44 pax × $20.85[ MXN]»; sin ellos, el concepto tal cual
    («TUA CZA · 44 pax»). El monto nunca se toca.
"""

import re
from html import escape

from app.schemas.reportes import (
    CotizacionGrupoAvionPdf,
    CotizacionGrupoLineaPdf,
    CotizacionGrupoPdfRequest,
)
from app.services.cotizacion_pdf import (
    _NAVY,
    _estilos_base,
    _fecha_legible,
    _ficha_aeronave_html,
    _itinerario_html,
    _logo_data_uri,
    _mapa_svg,
    _money,
    _mostrar_matricula,
    _vistazo_motores,
    _vistazo_velocidad,
)

# Claves que JAMÁS ve el cliente aunque algún API las mande como línea.
_CLAVES_OCULTAS = frozenset({"COMISION_VENDEDOR", "REDONDEO"})
# El IVA se pinta en el bloque de totales (subtotal / IVA / total), no como
# línea del cuerpo: si se dejara pasar saldría dos veces.
_CLAVES_TOTALES = frozenset({"IVA"})


def _folio_texto(r: CotizacionGrupoPdfRequest) -> str:
    if r.folio_grupo:
        return r.folio_grupo
    return f"G-{r.folio}" if r.folio else "G-s/n"


def _plural(n: int, singular: str, plural: str) -> str:
    return f"{n} {singular if n == 1 else plural}"


def _monto(v: float) -> str:
    """Monto con signo tipográfico: los negativos (descuento) con '−'."""
    return f"&minus;{_money(-v)}" if v < 0 else _money(v)


# Sufijo "· 44 pax" del concepto TUAS legado ("TUA CZA · 44 pax"): se quita
# antes de agregar "44 pax × $20.85" para no repetir los pasajeros.
_RE_PAX_SUFIJO = re.compile(r"\s*·\s*\d+\s*pax\s*$", re.IGNORECASE)


def _etiqueta_linea(ln: CotizacionGrupoLineaPdf) -> str:
    """Concepto TAL CUAL viene. Agrega "cantidad × unitario" SOLO si la línea
    los trae y el concepto no los incluye ya (el API los escribe en el
    concepto — "Tour Chichén Itzá · 44 × $85.00" — pero se tolera un payload
    que solo mande los números).

    TUAS (5-sep): la cantidad son PASAJEROS gravados, así que la operación
    se lee «TUA CZA · 44 pax × $20.85» (+ " MXN" si el unitario es nativo en
    pesos). El API manda el concepto "TUA CZA" cuando el unitario es
    uniforme; si llegara el legado "TUA CZA · 44 pax" junto con los números
    se quita ese sufijo para no duplicar el pax. Sin cantidad/unitario la
    línea queda como hoy. El monto no se toca."""
    lbl = ln.concepto or ln.clave or "Concepto"
    if ln.cantidad is None or ln.unitario is None or "×" in lbl:
        return lbl
    if (ln.clave or "").upper() == "TUAS":
        lbl = _RE_PAX_SUFIJO.sub("", lbl) or lbl
        lbl += f" · {ln.cantidad:g} pax × {_money(ln.unitario)}"
    else:
        lbl += f" · {ln.cantidad:g} × {_money(ln.unitario)}"
    if (ln.moneda or "USD").upper() == "MXN":
        lbl += " MXN"
    return lbl


def _lineas_cliente(r: CotizacionGrupoPdfRequest) -> list[tuple[str, str]]:
    """Cuerpo del desglose: (etiqueta, monto) en el ORDEN del API.

    Filtra defensivamente lo que nunca ve el cliente: comisión del vendedor,
    redondeo (también un AJUSTE positivo = redondeo hacia arriba) y el IVA
    (va en los totales). Sin líneas (API viejo) el cuerpo se arma con los
    escalares del payload, mismo criterio que el PDF de un avión.
    """
    filas: list[tuple[str, str]] = []
    for ln in r.desglose_consolidado:
        clave = (ln.clave or "").upper()
        if clave in _CLAVES_OCULTAS or clave in _CLAVES_TOTALES:
            continue
        if clave == "AJUSTE" and ln.monto_usd > 0:
            continue
        if (ln.concepto or "").strip().lower() == "redondeo":
            continue
        filas.append((_etiqueta_linea(ln), _monto(ln.monto_usd)))
    if filas:
        return filas

    # ----- Fallback (skew): escalares del payload -----
    n_av = r.aviones_total or len(r.aviones)
    lbl = "Servicio aéreo"
    if n_av:
        lbl += f" · {_plural(n_av, 'aeronave', 'aeronaves')}"
    filas.append((lbl, _money(r.servicio_aereo_usd)))
    if not r.tuas_detalle:
        filas.append(("TUAS", _money(r.tuas_usd)))
    elif len(r.tuas_detalle) == 1:
        filas.append((r.tuas_detalle[0], _money(r.tuas_usd)))
    else:
        filas.extend((det, "") for det in r.tuas_detalle)
        filas.append(("TUAS (total)", _money(r.tuas_usd)))
    for e in r.extras:
        lbl = e.concepto or "Extra"
        if e.cantidad is not None and e.unitario is not None and "×" not in lbl:
            lbl += f" · {e.cantidad:g} × {_money(e.unitario)}"
        if e.moneda == "MXN" and e.monto_nativo is not None:
            lbl += f" · ${e.monto_nativo:,.2f} MXN"
        filas.append((lbl, _money(e.monto_usd)))
    if r.viaticos_pernocta_usd > 0:
        filas.append(("Viáticos por pernocta", _money(r.viaticos_pernocta_usd)))
    if r.descuento_usd > 0:
        filas.append(("Descuento", f"&minus;{_money(r.descuento_usd)}"))
    return filas


def _desglose_html(r: CotizacionGrupoPdfRequest) -> str:
    """Tabla de desglose + totales (subtotal / IVA / total USD / MXN) y el
    precio por persona (toggle). Todo viene del API; el único derivado es
    el subtotal cuando un API viejo no lo manda (total − IVA, como en el
    PDF de un avión)."""
    filas = [
        f'<tr><td class="lbl">{escape(lbl)}</td><td class="val">{val}</td></tr>'
        for lbl, val in _lineas_cliente(r)
    ]
    subtotal = r.subtotal_usd if r.subtotal_usd is not None else r.total_usd - r.iva_usd
    filas.append(
        '<tr class="sub-row"><td class="lbl">Subtotal (sin IVA)</td>'
        f'<td class="val">{_money(subtotal)}</td></tr>'
    )
    filas.append(
        f'<tr><td class="lbl">IVA ({r.iva_pct:.0f}%)</td>'
        f'<td class="val">{_money(r.iva_usd)}</td></tr>'
    )
    filas.append(
        f'<tr class="total-row"><td>Total ({escape(r.moneda)})</td>'
        f'<td class="val">{_money(r.total_usd)}</td></tr>'
    )
    if r.total_mxn is not None:
        tc_txt = f" (T.C. {r.tc_usd_mxn:g})" if r.tc_usd_mxn else ""
        filas.append(
            f'<tr class="total-mxn"><td>Total MXN{escape(tc_txt)}</td>'
            f'<td class="val">{_money(r.total_mxn)} MXN</td></tr>'
        )
    # Precio por persona: SOLO si la cabecera lo pide y el API lo mandó
    # (nunca se divide aquí).
    if r.mostrar_precio_por_persona and r.precio_por_persona_usd:
        lbl = "Precio por persona"
        if r.pasajeros_total:
            lbl += f" ({_plural(r.pasajeros_total, 'pasajero', 'pasajeros')})"
        filas.append(
            f'<tr class="pp-row"><td>{escape(lbl)}</td>'
            f'<td class="val">{_money(r.precio_por_persona_usd)} {escape(r.moneda)}</td></tr>'
        )
    return f'<table class="totales"><tbody>{"".join(filas)}</tbody></table>'


def _anexo_flota_html(r: CotizacionGrupoPdfRequest) -> str:
    """Hoja "Flota asignada" (toggle `mostrar_anexo_aviones`): # / aeronave
    (modelo + matrícula solo VGV) / asientos / pasajeros / salidas
    (rotaciones) [/ horas + tarifa con `mostrar_tarifa`] [/ subtotal con
    `mostrar_subtotal_por_avion`]. El pie pinta pasajeros_total y total_usd
    del API — no se re-suman columnas."""
    if not r.mostrar_anexo_aviones or not r.aviones:
        return ""
    aviones = sorted(r.aviones, key=lambda a: a.posicion)
    ths = [
        "<th>#</th>",
        "<th>Aeronave</th>",
        '<th class="num">Asientos</th>',
        '<th class="num">Pasajeros</th>',
        "<th>Salidas</th>",
    ]
    if r.mostrar_tarifa:
        ths += ['<th class="num">Horas</th>', '<th class="num">Tarifa</th>']
    if r.mostrar_subtotal_por_avion:
        ths.append('<th class="num">Subtotal (USD)</th>')

    filas: list[str] = []
    for i, a in enumerate(aviones, start=1):
        nombre = escape(a.modelo) if a.modelo else "Aeronave"
        if _mostrar_matricula(a.matricula):
            nombre += f" · {escape(a.matricula or '')}"
        asientos = str(a.asientos) if a.asientos is not None else "—"
        vueltas = a.rotaciones if a.rotaciones and a.rotaciones > 0 else 1
        tds = [
            f"<td>{a.posicion or i}</td>",
            f"<td>{nombre}</td>",
            f'<td class="num">{asientos}</td>',
            f'<td class="num">{a.pasajeros}</td>',
            f"<td>{_plural(vueltas, 'vuelta', 'vueltas')}</td>",
        ]
        if r.mostrar_tarifa:
            horas = f"{a.tiempo_hr:g} h" if a.tiempo_hr else "—"
            tarifa = f"{_money(a.tarifa_hora_usd)}/hr" if a.tarifa_hora_usd else "—"
            tds += [f'<td class="num">{horas}</td>', f'<td class="num">{tarifa}</td>']
        if r.mostrar_subtotal_por_avion:
            sub = _money(a.subtotal_usd) if a.subtotal_usd is not None else "—"
            tds.append(f'<td class="num">{sub}</td>')
        filas.append(f"<tr>{''.join(tds)}</tr>")

    pie = [
        '<td colspan="3">Total del grupo</td>',
        f'<td class="num">{r.pasajeros_total or "—"}</td>',
        "<td></td>",
    ]
    if r.mostrar_tarifa:
        horas_t = f"{r.horas_total_hr:g} h" if r.horas_total_hr else ""
        pie += [f'<td class="num">{horas_t}</td>', "<td></td>"]
    if r.mostrar_subtotal_por_avion:
        pie.append(f'<td class="num">{_money(r.total_usd)}</td>')

    n_av = len(aviones)
    intro = _plural(n_av, "aeronave", "aeronaves")
    if r.pasajeros_total:
        intro += f" para {_plural(r.pasajeros_total, 'pasajero', 'pasajeros')}"
    nota = (
        '<p class="anexo-nota">Subtotales por aeronave con IVA incluido.</p>'
        if r.mostrar_subtotal_por_avion
        else ""
    )
    return (
        '<div class="anexo">'
        '<div class="av-titulo">Flota asignada</div><div class="av-linea"></div>'
        f'<p class="anexo-intro">{escape(intro)}.</p>'
        f'<table class="grid"><thead><tr>{"".join(ths)}</tr></thead>'
        f'<tbody>{"".join(filas)}</tbody>'
        f'<tfoot><tr>{"".join(pie)}</tr></tfoot></table>'
        f"{nota}</div>"
    )


def _primero(lista: list[CotizacionGrupoAvionPdf], attr: str):
    """Primer valor no vacío de `attr` entre los aviones del modelo."""
    for a in lista:
        v = getattr(a, attr)
        if v:
            return v
    return None


def _src_foto(data_uri: str | None, url: str | None) -> str | None:
    """Data URI (lo normal) o la URL pública como respaldo; escapado para
    ir dentro de src="…"."""
    foto = data_uri or url
    return escape(foto, quote=True) if foto else None


def _fichas_aeronaves_html(r: CotizacionGrupoPdfRequest) -> str:
    """Hojas "Las aeronaves": UNA por modelo distinto (en orden de posición),
    con la foto del primer avión del modelo que la traiga y la ficha "De un
    vistazo". Sin foto y sin datos de ficha, el modelo se omite. Nunca lleva
    matrícula (regla del cliente, también para el VGV)."""
    grupos: dict[str, list[CotizacionGrupoAvionPdf]] = {}
    for a in sorted(r.aviones, key=lambda a: a.posicion):
        grupos.setdefault((a.modelo or "").strip().upper(), []).append(a)

    bloques: list[str] = []
    for lista in grupos.values():
        con_foto = next(
            (
                a
                for a in lista
                if a.foto_exterior or a.foto_interior or a.foto_exterior_url or a.foto_interior_url
            ),
            None,
        )
        ref = con_foto or lista[0]
        vistazo: list[tuple[str, str]] = []
        asientos = _primero(lista, "asientos")
        if asientos:
            vistazo.append(("Pasajeros", f"{asientos} máx."))
        kts = _primero(lista, "velocidad_kts")
        if kts:
            vistazo.append(_vistazo_velocidad(kts))
        motores = _primero(lista, "num_motores")
        if motores:
            vistazo.append(_vistazo_motores(motores, _primero(lista, "motor_hp")))
        foto_ext = _src_foto(ref.foto_exterior, ref.foto_exterior_url)
        foto_int = _src_foto(ref.foto_interior, ref.foto_interior_url)
        if not (vistazo or foto_ext or foto_int):
            continue
        # Cuántos aviones de este modelo van en el viaje (conteo, no dinero).
        pax = sum(a.pasajeros for a in lista)
        en_viaje = _plural(len(lista), "aeronave", "aeronaves")
        if pax:
            en_viaje += f" · {_plural(pax, 'pasajero', 'pasajeros')}"
        vistazo.append(("En este viaje", en_viaje))
        ficha = _ficha_aeronave_html(
            _primero(lista, "modelo"),
            foto_ext,
            foto_int,
            vistazo,
            _primero(lista, "caracteristicas") or [],
        )
        if ficha:
            bloques.append(f'<div class="detalles">{ficha}</div>')
    return "".join(bloques)


def _estilos_grupo() -> str:
    """CSS propio del PDF de grupo (se suma a `_estilos_base`)."""
    return f"""
  .grid th.num, .grid td.num {{ text-align: right; white-space: nowrap; }}
  .grid tfoot td {{ font-weight: 700; background: #f7f7f8; color: {_NAVY}; }}
  .pp-row td {{ border-top: 1px solid #d1d5db; padding-top: 10px; font-size: 14px;
                font-weight: 800; color: {_NAVY}; }}
  .anexo {{ page-break-before: always; }}
  .anexo-intro {{ font-size: 13px; color: #374151; margin: 0 0 12px; }}
  .anexo-nota {{ font-size: 11px; color: #6b7280; margin-top: 8px; }}"""


def _build_html(r: CotizacionGrupoPdfRequest) -> str:
    folio = _folio_texto(r)

    # Título = ruta VISIBLE resuelta por el API; sin ella (skew) se arma de
    # la plantilla; sin tramos, el nombre del grupo.
    if r.ruta:
        ruta_titulo = escape(r.ruta)
    elif r.itinerario:
        ordenadas = sorted(r.itinerario, key=lambda x: x.orden)
        puntos = [ordenadas[0].origen] + [t.destino for t in ordenadas]
        ruta_titulo = " → ".join(escape(pt) for pt in puntos)
    else:
        ruta_titulo = escape(r.nombre) if r.nombre else "Cotización de grupo"
    n_puntos = ruta_titulo.count("→") + 1
    ruta_font = "26px" if n_puntos <= 4 else ("20px" if n_puntos <= 6 else "16px")

    n_av = r.aviones_total or len(r.aviones)
    partes: list[str] = []
    if r.pasajeros_total:
        partes.append(f"Grupo de {_plural(r.pasajeros_total, 'pasajero', 'pasajeros')}")
    if n_av:
        partes.append(_plural(n_av, "aeronave", "aeronaves"))
    subtitulo = escape(" · ".join(partes))

    # Itinerario + mapa: mismo bloque (y mismo mapa local/amplio) que el PDF
    # de un avión.
    mapa_html = _mapa_svg(r.mapa_puntos)
    itinerario_html = _itinerario_html(r.itinerario, mapa_html, r.mostrar_itinerario)

    notas_html = (
        f'<div class="notas"><strong>Notas:</strong> {escape(r.notas)}</div>' if r.notas else ""
    )
    condiciones_html = (
        f'<div class="notas"><strong>Condiciones:</strong> {escape(r.condiciones)}</div>'
        if r.condiciones
        else ""
    )

    logo_blanco = _logo_data_uri("logo-vuelatour-blanco.png")
    logo_marca = _logo_data_uri("logo-vuelatour.png")
    logo_header_html = f'<img class="logo" src="{logo_blanco}" alt=""/>' if logo_blanco else ""
    marca_html = (
        f'<div class="marca"><img src="{logo_marca}" alt=""/></div>' if logo_marca else ""
    )
    viaje_html = f"<br><strong>Viaje:</strong> {escape(r.nombre)}" if r.nombre else ""

    return f"""<!doctype html>
<html><head><meta charset="utf-8"><style>
{_estilos_base()}{_estilos_grupo()}
</style></head><body>
  {marca_html}
  <div class="header">
    {logo_header_html}
    <div class="titulos">
      <h1>{escape(r.empresa)}</h1>
      <p>Cotización de grupo</p>
    </div>
  </div>

  <div class="meta">
    <div><strong>Folio:</strong> {escape(folio)}<br>
      <strong>Cliente:</strong> {escape(r.cliente)}</div>
    <div style="text-align:right"><strong>Fecha de salida:</strong> {_fecha_legible(r.fecha)}
      {viaje_html}</div>
  </div>

  <div class="route" style="font-size:{ruta_font}">{ruta_titulo}</div>
  <div style="font-size:13px;color:#374151">{subtitulo}</div>
  {itinerario_html}

  <h2>Desglose</h2>
  {_desglose_html(r)}
  {notas_html}
  {condiciones_html}

  {_anexo_flota_html(r)}
  {_fichas_aeronaves_html(r)}

</body></html>"""


def render_cotizacion_grupo_pdf(req: CotizacionGrupoPdfRequest) -> bytes:
    from weasyprint import HTML  # import perezoso

    return HTML(string=_build_html(req)).write_pdf()
