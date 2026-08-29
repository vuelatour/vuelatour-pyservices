"""Render del PDF de reparto de utilidades (doc 5.9 — reporte 'vital')."""

from io import BytesIO

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    KeepTogether,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from app.schemas.reparto import (
    RepartoAvion,
    RepartoOtrosIngresosDesglose,
    RepartoPdfRequest,
    RepartoVueloLinea,
)

BRAND = colors.HexColor("#0F4C81")
LIGHT = colors.HexColor("#EEF2F7")
MUTED = colors.HexColor("#5B6470")


def _nota_tc_oficial(avion) -> str | None:
    """Nota del TC oficial de respaldo del avión (29-ago-2026), o None."""
    partes = []
    g = getattr(avion, "gastos_tc_oficial", None)
    if g is not None and g.count > 0:
        partes.append(f"{g.count} gasto(s) MXN sin TC capturado (${g.monto_mxn:,.2f} MXN)")
    c = getattr(avion, "cobros_tc_oficial_count", 0) or 0
    if c > 0:
        partes.append(f"{c} vuelo(s) con cobros MXN sin TC")
    if not partes:
        return None
    return (
        " y ".join(partes)
        + " se convirtieron con el TC oficial de referencia del día "
        "(open.er-api / BCE); ya están dentro de las cifras."
    )


def _nota_tc_oficial_global(tc) -> str | None:
    """Nota del periodo (resumen): vuelos/gastos convertidos con TC oficial."""
    if tc is None:
        return None
    gastos = tc.gastos.count if tc.gastos is not None else 0
    if tc.vuelos <= 0 and gastos <= 0:
        return None
    partes = []
    if gastos > 0:
        partes.append(f"{gastos} gasto(s) MXN (${tc.gastos.monto_mxn:,.2f} MXN)")
    if tc.vuelos > 0:
        partes.append(f"{tc.vuelos} vuelo(s) con cobros MXN")
    fuentes = f" Fuente(s): {' / '.join(tc.fuentes)}." if tc.fuentes else ""
    return (
        "Tipo de cambio de referencia: "
        + " y ".join(partes)
        + " sin TC capturado se convirtieron con el TC oficial del día de la "
        "cotización o del gasto (open.er-api / BCE); ya están dentro de las cifras."
        + fuentes
    )


def _usd(value: float) -> str:
    return f"${value:,.2f}"


def _pct(factor: float | None) -> str:
    """0.5 → '50 %' (hasta 2 decimales, sin ceros de más)."""
    return "—" if factor is None else f"{round(factor * 100, 2):g} %"


def _desglose_txt(d: RepartoOtrosIngresosDesglose | None) -> str | None:
    """'TUAs $100.00 · comisión del vendedor $1,000.00 · IVA $176.00' con la
    composición COTIZADA de otros ingresos VuelaTour; None si no hay nada."""
    if d is None:
        return None
    partes = [
        f"{label} {_usd(val)}"
        for label, val in (
            ("TUAs", d.tuas_usd),
            ("extras", d.extras_usd),
            ("pernocta", d.pernocta_usd),
            ("comisión del vendedor", d.comision_usd),
            ("IVA", d.iva_usd),
        )
        if val
    ]
    return " · ".join(partes) or None


def _folio_vuelo(v: RepartoVueloLinea) -> str:
    """'#105 · 50 % (CUN-MID)' cuando el vuelo fue MULTI-AVIÓN (regla B,
    28-ago-2026: la venta del avión se repartió por tramo); '#105' si no.
    Helvetica no dibuja '→' (sale una caja): se sustituye por '-'."""
    texto = f"#{v.folio}" if v.folio is not None else "—"
    if v.participacion is not None and v.participacion < 0.9999:
        texto += f" · {_pct(v.participacion)}"
        if v.tramos_avion:
            texto += f" ({v.tramos_avion.replace('→', '-')})"
    return texto


def render_reparto_pdf(req: RepartoPdfRequest) -> bytes:
    """Genera el PDF del reparto de utilidades y devuelve los bytes."""
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        topMargin=18 * mm,
        bottomMargin=18 * mm,
        leftMargin=20 * mm,
        rightMargin=20 * mm,
        title=f"Reparto {req.periodo_desde} a {req.periodo_hasta}",
    )

    base = getSampleStyleSheet()
    s_marca = ParagraphStyle("marca", parent=base["Title"], fontSize=18, textColor=BRAND)
    s_doc = ParagraphStyle(
        "doc", parent=base["Normal"], fontSize=14, textColor=BRAND, alignment=2
    )
    s_meta = ParagraphStyle(
        "meta", parent=base["Normal"], fontSize=8, textColor=MUTED, alignment=2
    )
    s_sub = ParagraphStyle("sub", parent=base["Normal"], fontSize=9, textColor=MUTED)
    s_avion = ParagraphStyle(
        "avion", parent=base["Normal"], fontSize=12, textColor=BRAND, spaceBefore=4
    )
    s_seccion = ParagraphStyle(
        "seccion", parent=base["Normal"], fontSize=11, textColor=BRAND
    )
    s_foot = ParagraphStyle(
        "foot", parent=base["Normal"], fontSize=8, textColor=MUTED, leading=11
    )

    story: list = []

    # ---- Encabezado ----
    encabezado = Table(
        [
            [
                Paragraph("Vuela Tour", s_marca),
                Paragraph("REPARTO DE UTILIDADES", s_doc),
            ],
            [
                Paragraph("Aero Charter Cancún S.A. de C.V.", s_sub),
                Paragraph(
                    f"Periodo: {req.periodo_desde} a {req.periodo_hasta}<br/>"
                    f"Generado: {req.generado} (hora de Cancún, UTC−5)",
                    s_meta,
                ),
            ],
        ],
        colWidths=[95 * mm, 75 * mm],
    )
    encabezado.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ]
        )
    )
    story.append(encabezado)
    story.append(Spacer(1, 3 * mm))
    story.append(_regla())
    story.append(Spacer(1, 5 * mm))

    # ---- Resumen global ----
    total_ingresos = sum(a.ingresos_cobrado_usd for a in req.aviones)
    total_saldo = sum(a.saldo_usd for a in req.aviones)
    # comisiones_venta_usd llega 0 desde el 28-ago-2026 (regla A: la comisión
    # del vendedor es ingreso/pago de VuelaTour, no costo del avión); se suma
    # solo por compat con payloads viejos.
    total_gastos = sum(
        _gastos_avion(a) + a.comisiones_venta_usd for a in req.aviones
    )
    # Otros ingresos de VuelaTour (TUAs/extras/pernocta cobrados): solo
    # informativo, FUERA del saldo y del reparto. Viene del API; si un API
    # viejo no manda el global, se re-suman los por avión solo para mostrar.
    total_otros = req.otros_ingresos_vuelatour_total_usd or sum(
        a.otros_ingresos_vuelatour_usd for a in req.aviones
    )
    resumen = Table(
        [
            [
                "Venta del avión cobrada",
                "Otros ingresos VuelaTour",
                "Gastos del periodo",
                "Saldo a repartir",
            ],
            [_usd(total_ingresos), _usd(total_otros), _usd(total_gastos), _usd(total_saldo)],
        ],
        colWidths=[42.5 * mm, 42.5 * mm, 42.5 * mm, 42.5 * mm],
    )
    resumen.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), LIGHT),
                ("TEXTCOLOR", (0, 0), (-1, 0), MUTED),
                ("FONTSIZE", (0, 0), (-1, 0), 8),
                ("FONTSIZE", (0, 1), (-1, 1), 13),
                ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
                ("TEXTCOLOR", (0, 1), (-1, 1), BRAND),
                # Otros ingresos VuelaTour: informativo (gris), no se reparte.
                ("TEXTCOLOR", (1, 1), (1, 1), MUTED),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ("BOX", (0, 0), (-1, -1), 0.5, LIGHT),
                ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.white),
            ]
        )
    )
    story.append(resumen)
    # Composición cotizada de "Otros ingresos VuelaTour" (regla A: incluye la
    # comisión del vendedor) — solo si el API manda el desglose.
    desglose_txt = _desglose_txt(req.otros_ingresos_vuelatour_desglose)
    if desglose_txt:
        story.append(Spacer(1, 2 * mm))
        story.append(
            Paragraph(
                "Otros ingresos VuelaTour del periodo (cotizado, no se reparte): "
                f"{desglose_txt}. El pago de la comisión al vendedor sale de "
                "VuelaTour (otros movimientos), no del avión.",
                s_sub,
            )
        )
    # Nota global del TC oficial de respaldo (29-ago-2026): cuántos vuelos y
    # gastos del periodo entraron convertidos con él (open.er-api / BCE).
    nota_global = _nota_tc_oficial_global(req.tc_oficial)
    if nota_global:
        story.append(Spacer(1, 2 * mm))
        story.append(Paragraph(nota_global, s_sub))
    story.append(Spacer(1, 8 * mm))

    # ---- Por aeronave ----
    if not req.aviones:
        story.append(Paragraph("Sin aeronaves en el periodo.", s_sub))

    for avion in req.aviones:
        story.append(_bloque_avion(avion, s_avion))
        story.append(Spacer(1, 6 * mm))

    # ---- Resumen por socio ----
    por_socio: dict[str, float] = {}
    for avion in req.aviones:
        for linea in avion.reparto:
            por_socio[linea.socio_nombre] = (
                por_socio.get(linea.socio_nombre, 0.0) + linea.monto_usd
            )
    if por_socio:
        story.append(_regla())
        story.append(Spacer(1, 3 * mm))
        story.append(Paragraph("Resumen por socio", s_seccion))
        story.append(Spacer(1, 2 * mm))
        filas = [["Socio", "Total a recibir"]]
        for nombre, monto in sorted(
            por_socio.items(), key=lambda kv: kv[1], reverse=True
        ):
            filas.append([nombre, _usd(monto)])
        tabla = Table(filas, colWidths=[120 * mm, 50 * mm])
        tabla.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), BRAND),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, -1), 9),
                    ("ALIGN", (1, 0), (1, -1), "RIGHT"),
                    ("TOPPADDING", (0, 0), (-1, -1), 5),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                    ("LEFTPADDING", (0, 0), (-1, -1), 8),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, LIGHT]),
                ]
            )
        )
        story.append(tabla)

    story.append(Spacer(1, 10 * mm))
    story.append(_regla())
    story.append(Spacer(1, 3 * mm))
    story.append(
        Paragraph(
            "Solo se reparte lo cobrado. El pendiente de cobro (parte avión) se "
            "distribuye cuando entra el pago; la deuda total del cliente incluye "
            "además TUAs/extras/pernocta/comisión del vendedor. La venta del avión = "
            "tiempo de vuelo + ajuste + IVA proporcional (en vuelos multi-avión, la "
            "parte de cada matrícula en partes iguales por tramo vendido; los "
            "ferries/tramos operativos no reparten); los TUAs, extras, viáticos de "
            "pernocta y la comisión del vendedor cobrados son ingreso de VuelaTour: "
            "no entran al saldo ni se reparten, y el pago de la comisión al vendedor "
            "sale de VuelaTour, no del avión (detalle en 'otros movimientos' del "
            "Balance general). Montos en USD. "
            "Vuela Tour · Aero Charter Cancún.",
            s_foot,
        )
    )

    doc.build(story)
    return buffer.getvalue()


def _gastos_avion(a: RepartoAvion) -> float:
    return (
        a.gastos_directos_usd
        + a.gastos_indirectos_usd
        + a.permisos_usd
        + a.otros_usd
        + a.reserva_overhaul_usd
    )


def _bloque_avion(avion: RepartoAvion, estilo_titulo: ParagraphStyle) -> KeepTogether:
    titulo = Paragraph(f"{avion.matricula} — {avion.modelo}", estilo_titulo)

    filas: list[list[str]] = [
        ["Horas voladas del periodo", f"{avion.horas_voladas_hr:.1f} hr"],
        ["Venta del avión cobrada", _usd(avion.ingresos_cobrado_usd)],
    ]
    # Filas informativas en gris bajo la venta, FUERA de la cascada (no
    # suman al saldo ni se reparten):
    #  - otros ingresos de VuelaTour (TUAs/extras/pernocta cobrados);
    #  - pendiente de cobro = parte del AVIÓN; si el cliente debe más
    #    (TUAs/extras/pernocta sin cobrar) se anota la deuda total.
    info_rows: list[int] = []
    if avion.otros_ingresos_vuelatour_usd:
        info_rows.append(len(filas))
        filas.append([
            "Otros ingresos VuelaTour (TUAs/extras/pernocta/comisión vendedor) — no se reparten",
            _usd(avion.otros_ingresos_vuelatour_usd),
        ])
    # Regla A (28-ago-2026): la comisión del vendedor es ingreso de VuelaTour
    # (ya dentro de la línea anterior) y su pago sale de VuelaTour — se anota
    # solo si el API manda el desglose.
    desg = avion.otros_ingresos_vuelatour_desglose
    if desg is not None and desg.comision_usd:
        info_rows.append(len(filas))
        filas.append([
            "   incluye comisión del vendedor cotizada (pre-IVA) — su pago no es costo del avión",
            _usd(desg.comision_usd),
        ])
    pendiente = avion.pendiente_cobro_usd or 0.0
    bruto = avion.pendiente_bruto_usd or 0.0
    if pendiente or bruto:
        texto = "Pendiente de cobro (parte avión) — se reparte al cobrar"
        if bruto > pendiente + 0.005:
            texto = f"Pendiente de cobro (parte avión) — deuda total del cliente {_usd(bruto)}"
        info_rows.append(len(filas))
        filas.append([texto, _usd(pendiente)])
    filas += [
        # Comisiones de venta: solo payloads viejos (antes del 28-ago-2026);
        # hoy el API manda 0 y la fila NO se imprime (regla A).
        *(
            [["(-) Comisiones de venta", _usd(-avion.comisiones_venta_usd)]]
            if avion.comisiones_venta_usd
            else []
        ),
        ["(-) Gastos directos", _usd(-avion.gastos_directos_usd)],
        ["(-) Gastos indirectos", _usd(-avion.gastos_indirectos_usd)],
        ["(-) Permisos", _usd(-avion.permisos_usd)],
        ["(-) Fijos (prorrateo + reparto manual)", _usd(-avion.otros_usd)],
        ["(-) Reserva overhaul", _usd(-avion.reserva_overhaul_usd)],
        ["Saldo disponible", _usd(avion.saldo_usd)],
    ]
    cascada = Table(filas, colWidths=[120 * mm, 50 * mm])
    estilo: list = [
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LINEABOVE", (0, -1), (-1, -1), 0.6, MUTED),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("TEXTCOLOR", (0, -1), (-1, -1), BRAND),
        ("BACKGROUND", (0, -1), (-1, -1), LIGHT),
    ]
    for info_idx in info_rows:
        estilo += [
            ("TEXTCOLOR", (0, info_idx), (-1, info_idx), MUTED),
            ("FONTSIZE", (0, info_idx), (-1, info_idx), 7.5),
            ("FONTNAME", (0, info_idx), (-1, info_idx), "Helvetica-Oblique"),
        ]
    cascada.setStyle(TableStyle(estilo))

    bloque: list = [titulo, Spacer(1, 2 * mm), cascada]

    # Detalle de vuelos (solo si el API lo manda): el folio lleva ' · 50 %'
    # cuando el vuelo fue MULTI-AVIÓN (regla B: la venta del avión se
    # repartió por tramo entre las matrículas).
    if avion.vuelos:
        vfilas = [["Vuelo", "Cliente", "Cobrado avión", "Pendiente avión"]]
        for v in avion.vuelos:
            vfilas.append([
                _folio_vuelo(v),
                v.cliente or "",
                _usd(v.cobrado_usd or 0.0),
                _usd(v.pendiente_usd or 0.0),
            ])
        vtabla = Table(vfilas, colWidths=[60 * mm, 60 * mm, 25 * mm, 25 * mm])
        vtabla.setStyle(
            TableStyle(
                [
                    ("FONTSIZE", (0, 0), (-1, -1), 7.5),
                    ("TEXTCOLOR", (0, 0), (-1, -1), MUTED),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("ALIGN", (2, 0), (-1, -1), "RIGHT"),
                    ("TOPPADDING", (0, 0), (-1, -1), 2),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                    ("LINEBELOW", (0, 0), (-1, 0), 0.4, LIGHT),
                ]
            )
        )
        bloque.append(Spacer(1, 2 * mm))
        bloque.append(vtabla)

    # Advertencias de integridad: dinero que NO pudo entrar al balance. El
    # supervisor debe verlo aquí mismo, no descubrirlo cuadrando a mano.
    avisos = []
    if avion.gastos_sin_tc_mxn > 0:
        avisos.append(
            f"Gastos MXN sin tipo de cambio por ${avion.gastos_sin_tc_mxn:,.2f} MXN excluidos del balance."
        )
    if avion.cobros_sin_tc_mxn > 0:
        avisos.append(
            f"Cobros MXN sin tipo de cambio por ${avion.cobros_sin_tc_mxn:,.2f} MXN excluidos del ingreso."
        )
    if avion.reserva_incompleta:
        avisos.append("Sin tarifa de reserva de overhaul configurada pese a horas voladas.")
    # Vigencias de socios traslapadas o incompletas: el reparto impreso saldría
    # doble (o corto). La web lo delata con un badge; el papel debe decirlo igual.
    if avion.reparto and round(avion.reparto_porcentaje_total, 2) != 100:
        avisos.append(
            f"Los porcentajes de socios suman {avion.reparto_porcentaje_total:g}% (no 100%): "
            "revisar vigencias en el catálogo de socios antes de dispersar."
        )
    if avisos:
        aviso_style = ParagraphStyle(
            "aviso", fontSize=7.5, textColor=colors.HexColor("#b45309"), leading=10
        )
        for a in avisos:
            bloque.append(Spacer(1, 1 * mm))
            bloque.append(Paragraph(f"AVISO: {a}", aviso_style))

    # Nota INFORMATIVA (no advertencia, 29-ago-2026): montos MXN sin tipo de
    # cambio capturado que SÍ entraron, convertidos con el TC oficial de
    # referencia del día (open.er-api / BCE). Ya están dentro de las cifras.
    nota_tc = _nota_tc_oficial(avion)
    if nota_tc:
        nota_style = ParagraphStyle("nota_tc", fontSize=7.5, textColor=MUTED, leading=10)
        bloque.append(Spacer(1, 1 * mm))
        bloque.append(Paragraph(f"Nota: {nota_tc}", nota_style))

    if avion.reparto:
        rfilas = [["Socio", "%", "Monto"]]
        for linea in avion.reparto:
            rfilas.append(
                [linea.socio_nombre, f"{linea.porcentaje:g}%", _usd(linea.monto_usd)]
            )
        rtabla = Table(rfilas, colWidths=[100 * mm, 25 * mm, 45 * mm])
        rtabla.setStyle(
            TableStyle(
                [
                    ("FONTSIZE", (0, 0), (-1, -1), 8),
                    ("TEXTCOLOR", (0, 0), (-1, 0), MUTED),
                    ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
                    ("TOPPADDING", (0, 0), (-1, -1), 3),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                    ("LINEBELOW", (0, 0), (-1, 0), 0.4, LIGHT),
                ]
            )
        )
        bloque.append(Spacer(1, 2 * mm))
        bloque.append(rtabla)

    return KeepTogether(bloque)


def _regla() -> Table:
    regla = Table([[""]], colWidths=[170 * mm], rowHeights=[0.1])
    regla.setStyle(TableStyle([("LINEBELOW", (0, 0), (-1, -1), 1, BRAND)]))
    return regla
