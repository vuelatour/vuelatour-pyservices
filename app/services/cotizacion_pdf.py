"""Render del PDF de cotizacion con ReportLab (doc 5.1 / 4.2)."""

from io import BytesIO

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from app.schemas.cotizacion import CotizacionPdfRequest

BRAND = colors.HexColor("#0F4C81")
LIGHT = colors.HexColor("#EEF2F7")
MUTED = colors.HexColor("#5B6470")


def _usd(value: float) -> str:
    return f"${value:,.2f}"


def render_cotizacion_pdf(req: CotizacionPdfRequest) -> bytes:
    """Genera el PDF de la cotizacion y devuelve los bytes."""
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        topMargin=18 * mm,
        bottomMargin=18 * mm,
        leftMargin=20 * mm,
        rightMargin=20 * mm,
        title=f"Cotizacion {req.folio}",
    )

    base = getSampleStyleSheet()
    s_marca = ParagraphStyle(
        "marca", parent=base["Title"], fontSize=20, textColor=BRAND, spaceAfter=0
    )
    s_sub = ParagraphStyle(
        "sub", parent=base["Normal"], fontSize=9, textColor=MUTED
    )
    s_doc = ParagraphStyle(
        "doc",
        parent=base["Normal"],
        fontSize=16,
        textColor=BRAND,
        alignment=2,
        leading=18,
    )
    s_doc_meta = ParagraphStyle(
        "docmeta", parent=base["Normal"], fontSize=9, textColor=MUTED, alignment=2
    )
    s_label = ParagraphStyle(
        "label", parent=base["Normal"], fontSize=8, textColor=MUTED
    )
    s_value = ParagraphStyle("value", parent=base["Normal"], fontSize=10)
    s_foot = ParagraphStyle(
        "foot", parent=base["Normal"], fontSize=8, textColor=MUTED, leading=11
    )

    story: list = []

    # ---- Encabezado ----
    encabezado = Table(
        [
            [
                Paragraph("Vuela Tour", s_marca),
                Paragraph(f"COTIZACIÓN<br/>Folio #{req.folio} · v.{req.version}", s_doc),
            ],
            [
                Paragraph("Aero Charter Cancún S.A. de C.V.", s_sub),
                Paragraph(f"Fecha: {req.fecha}", s_doc_meta),
            ],
        ],
        colWidths=[100 * mm, 70 * mm],
    )
    encabezado.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )
    story.append(encabezado)
    story.append(Spacer(1, 4 * mm))
    story.append(_regla())
    story.append(Spacer(1, 5 * mm))

    # ---- Datos de la cotizacion ----
    datos = [
        ("Cliente", req.cliente),
        ("Ruta", req.ruta),
        ("Aeronave", req.aeronave),
        ("Tipo de vuelo", req.tipo_vuelo),
        ("Pasajeros", str(req.pasajeros)),
    ]
    info = Table(
        [[Paragraph(k, s_label), Paragraph(v, s_value)] for k, v in datos],
        colWidths=[35 * mm, 135 * mm],
    )
    info.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("LINEBELOW", (0, 0), (-1, -2), 0.4, LIGHT),
            ]
        )
    )
    story.append(info)
    story.append(Spacer(1, 7 * mm))

    # ---- Desglose ----
    filas = [["Concepto", "Monto USD"]]
    for linea in req.lineas:
        filas.append([linea.concepto, _usd(linea.monto_usd)])

    n_lineas = len(req.lineas)
    filas.append(["Subtotal", _usd(req.subtotal_usd)])
    if req.tuas_usd:
        filas.append(["TUAS", _usd(req.tuas_usd)])
    if req.iva_usd:
        filas.append(["IVA", _usd(req.iva_usd)])
    filas.append(["TOTAL", _usd(req.total_usd)])

    tabla = Table(filas, colWidths=[125 * mm, 45 * mm])
    estilo = [
        ("BACKGROUND", (0, 0), (-1, 0), BRAND),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("LINEBELOW", (0, 1), (-1, n_lineas), 0.4, LIGHT),
        ("LINEABOVE", (0, n_lineas + 1), (-1, n_lineas + 1), 0.6, MUTED),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, -1), (-1, -1), 11),
        ("TEXTCOLOR", (0, -1), (-1, -1), BRAND),
        ("BACKGROUND", (0, -1), (-1, -1), LIGHT),
    ]
    tabla.setStyle(TableStyle(estilo))
    story.append(tabla)

    if req.tc_usd_mxn:
        story.append(Spacer(1, 3 * mm))
        total_mxn = req.total_usd * req.tc_usd_mxn
        story.append(
            Paragraph(
                f"Tipo de cambio referencia: ${req.tc_usd_mxn:,.4f} MXN/USD · "
                f"Total aproximado: ${total_mxn:,.2f} MXN",
                s_foot,
            )
        )

    if req.notas:
        story.append(Spacer(1, 7 * mm))
        story.append(Paragraph("Notas", s_label))
        story.append(Spacer(1, 1 * mm))
        story.append(Paragraph(req.notas, s_value))

    story.append(Spacer(1, 12 * mm))
    story.append(_regla())
    story.append(Spacer(1, 3 * mm))
    story.append(
        Paragraph(
            "Cotización en dólares estadounidenses. Sujeta a disponibilidad de "
            "aeronave y a confirmación. El avión siempre regresa a su base. "
            "Vuela Tour · Aero Charter Cancún.",
            s_foot,
        )
    )

    doc.build(story)
    return buffer.getvalue()


def _regla() -> Table:
    regla = Table([[""]], colWidths=[170 * mm], rowHeights=[0.1])
    regla.setStyle(TableStyle([("LINEBELOW", (0, 0), (-1, -1), 1, BRAND)]))
    return regla
