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

from app.schemas.reparto import RepartoAvion, RepartoPdfRequest

BRAND = colors.HexColor("#0F4C81")
LIGHT = colors.HexColor("#EEF2F7")
MUTED = colors.HexColor("#5B6470")


def _usd(value: float) -> str:
    return f"${value:,.2f}"


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
    total_gastos = sum(_gastos_avion(a) for a in req.aviones)
    resumen = Table(
        [
            ["Ingresos cobrados", "Gastos del periodo", "Saldo a repartir"],
            [_usd(total_ingresos), _usd(total_gastos), _usd(total_saldo)],
        ],
        colWidths=[56 * mm, 56 * mm, 58 * mm],
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
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ("BOX", (0, 0), (-1, -1), 0.5, LIGHT),
                ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.white),
            ]
        )
    )
    story.append(resumen)
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
            "Solo se reparte lo cobrado. El saldo pendiente de cobro se distribuye "
            "cuando entra el pago. Montos en USD. Vuela Tour · Aero Charter Cancún.",
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

    filas = [
        ["Ingresos cobrados", _usd(avion.ingresos_cobrado_usd)],
        ["(-) Gastos directos", _usd(-avion.gastos_directos_usd)],
        ["(-) Gastos indirectos", _usd(-avion.gastos_indirectos_usd)],
        ["(-) Permisos", _usd(-avion.permisos_usd)],
        ["(-) Otros gastos (prorrateados)", _usd(-avion.otros_usd)],
        ["(-) Reserva overhaul", _usd(-avion.reserva_overhaul_usd)],
        ["Saldo disponible", _usd(avion.saldo_usd)],
    ]
    cascada = Table(filas, colWidths=[120 * mm, 50 * mm])
    cascada.setStyle(
        TableStyle(
            [
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
        )
    )

    bloque: list = [titulo, Spacer(1, 2 * mm), cascada]

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
