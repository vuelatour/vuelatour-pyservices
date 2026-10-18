"""Render del PDF de reparto de utilidades (doc 5.9 — reporte 'vital')."""

from io import BytesIO
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.lib.utils import simpleSplit
from reportlab.pdfbase.pdfmetrics import stringWidth
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
    RepartoCierre,
    RepartoCierreVuelo,
    RepartoGastoCategoriaLinea,
    RepartoOtroIngresoLinea,
    RepartoOtrosIngresosDesglose,
    RepartoPdfRequest,
    RepartoVueloLinea,
)
from app.services import cierre_mes
from app.services.reparto_comisiones import (
    ETIQUETA_FILA,
    LINEA_COMISION_VENDEDOR,
    absorbe_comisiones,
    comisiones_al_avion,
    frase_regla,
    linea_comision_vendedor,
    solo_gasto_real,
)

BRAND = colors.HexColor("#0F4C81")
LIGHT = colors.HexColor("#EEF2F7")
MUTED = colors.HexColor("#5B6470")

# ===== Otros ingresos ASIGNADOS a un avión (13-oct-2026, API 0.0.75) =====
# Invariante «INGRESO ASIGNADO A UN AVIÓN» del API: los ingresos registrados
# en Ingresos con ese avión (reembolso de su servicio, venta de una pieza
# suya…), netos de comisión bancaria, SUMAN a su saldo (`saldo_usd` ya los
# trae: aquí jamás se vuelven a sumar). Textos EXACTOS del contrato (es-MX).
FILA_OTROS_INGRESOS_AVION = (
    "(+) Otros ingresos del avión (registrados en Ingresos, netos de comisión)"
)
LINEA_OTROS_INGRESOS_SIN_TC = (
    "Otros ingresos del avión en MXN sin tipo de cambio — no entraron al saldo"
)
# Detalle gris bajo la fila (+): hasta 5 ingresos y «y N más».
MAX_LINEAS_OTROS_INGRESOS = 5
ENCABEZADOS_RESUMEN_CON_OTROS_INGRESOS = [
    "Venta del avión cobrada",
    "Otros ingresos de aviones",
    "Otros ingresos VuelaTour",
    "Gastos del periodo",
    "Saldo a repartir",
]
# Ancho útil de la columna de etiquetas de la cascada (120 mm − relleno) y
# letra de las filas grises: el detalle de un ingreso se RECORTA a ese ancho
# (un concepto largo se encimaría con el monto).
_ANCHO_ETIQUETA_PT = 120 * mm - 12
_FUENTE_GRIS = ("Helvetica-Oblique", 7.5)


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


# Pie del documento. La cláusula de en medio depende de la regla de
# comisiones (6-oct-2026): sin comisiones absorbidas, la de siempre.
_PIE_INICIO = (
    "Solo se reparte lo cobrado. El pendiente de cobro (parte avión) se "
    "distribuye cuando entra el pago; la deuda total del cliente incluye "
    "además TUAs/extras/pernocta/comisión del vendedor. La venta del avión = "
    "tiempo de vuelo + ajuste + IVA proporcional (en vuelos multi-avión, la "
    "parte de cada matrícula en partes iguales por tramo vendido; los "
    "ferries/tramos operativos no reparten); los TUAs, extras, viáticos de "
    "pernocta y la comisión del vendedor cobrados son ingreso de VuelaTour: "
    "no entran al saldo ni se reparten"
)
_PIE_COMISION_VENDEDOR = (
    ", y el pago de la comisión al vendedor sale de VuelaTour, no del avión "
    "(detalle en 'otros movimientos' del Balance general)."
)
_PIE_FIN = " Montos en USD. Vuela Tour · Aero Charter Cancún."
# Línea gris bajo «Otros ingresos VuelaTour» cuando el avión NO absorbe
# comisiones (la de siempre; con comisiones, `linea_comision_vendedor(req)`:
# «su provisión…» o, con el API 0.0.74, «pago «(avión)» resta…»).
_LINEA_COMISION_VENDEDOR_VUELATOUR = (
    "   incluye comisión del vendedor cotizada (pre-IVA) — su pago no es costo del avión"
)


def _pie(req: RepartoPdfRequest) -> str:
    """Pie del PDF. Con comisiones a cargo del avión (regla de septiembre
    2026) la frase «el pago de la comisión al vendedor sale de VuelaTour, no
    del avión» ya no es verdad: dice qué absorbe el avión y dónde resta."""
    if not comisiones_al_avion(req):
        return _PIE_INICIO + _PIE_COMISION_VENDEDOR + _PIE_FIN
    return (
        f"{_PIE_INICIO}. {frase_regla(req)}: restan antes del saldo en "
        f"{ETIQUETA_FILA} (detalle en 'otros movimientos' del Balance "
        f"general).{_PIE_FIN}"
    )


def _nota_otros_ingresos(req: RepartoPdfRequest, desglose_txt: str) -> str:
    """Párrafo bajo el resumen con la composición cotizada de «Otros ingresos
    VuelaTour». Con comisiones a cargo del avión ya no dice que el pago de la
    comisión «no es del avión»."""
    base = f"Otros ingresos VuelaTour del periodo (cotizado, no se reparte): {desglose_txt}. "
    if not comisiones_al_avion(req):
        return base + (
            "El pago de la comisión al vendedor sale de VuelaTour (otros "
            "movimientos), no del avión."
        )
    if solo_gasto_real(req):
        # API 0.0.74 (13-oct-2026, invariante 61, dos categorías): el avión
        # carga el pago «Comisión del vendedor (avión)» («otros movimientos»
        # lo trae con egreso $0.00) y el de «(VuelaTour)» lo paga VuelaTour.
        return base + (
            f"{frase_regla(req)}: restan en {ETIQUETA_FILA}. Lo cobrado al "
            "cliente por la comisión sigue siendo de VuelaTour; el pago "
            "capturado como «Comisión del vendedor (avión)» lo carga el avión "
            "('otros movimientos' del Balance general lo muestra con egreso "
            "$0.00) y el de «Comisión del vendedor (VuelaTour)» lo paga "
            "VuelaTour ahí mismo."
        )
    return base + (
        f"{frase_regla(req)}: restan en {ETIQUETA_FILA}. Lo cobrado al "
        "cliente por la comisión sigue siendo de VuelaTour, que le paga al "
        "vendedor (otros movimientos)."
    )


# Ancho útil de la columna «Vuelo» del detalle (60 mm menos el relleno de la
# celda): la marca del cierre se parte a este ancho.
_ANCHO_FOLIO_PT = 60 * mm - 12


def _folio_vuelo(v: RepartoVueloLinea) -> str:
    """'#105 · 50 % (CUN-MID)' cuando el vuelo fue MULTI-AVIÓN (regla B,
    28-ago-2026: la venta del avión se repartió por tramo); '#105' si no.
    Helvetica no dibuja '→' (sale una caja): se sustituye por '-'. Con la
    marca del cierre de mes (16-oct-2026) va DEBAJO, partida al ancho de la
    columna (una celda de texto de la tabla no envuelve sola: se encimaría
    con el cliente): '#105\nViene de septiembre 2026'."""
    texto = f"#{v.folio}" if v.folio is not None else "—"
    if v.participacion is not None and v.participacion < 0.9999:
        texto += f" · {_pct(v.participacion)}"
        if v.tramos_avion:
            texto += f" ({v.tramos_avion.replace('→', '-')})"
    if v.cierre_marca:
        renglones = simpleSplit(
            v.cierre_marca.replace("→", "-"), "Helvetica", 7.5, _ANCHO_FOLIO_PT
        )
        texto += "\n" + "\n".join(renglones)
    return texto


# ===== CIERRE DE MES (16-oct-2026, API 0.0.88, invariante 89 del API) =====
# Un vuelo COMPLETADO sin cobrar por completo al cerrar el mes pasa COMPLETO
# al mes siguiente y se reparte cuando está cobrado. El API manda en `cierre`
# el estado del mes y las listas (`diferidos_salen`, `diferidos_entran`,
# `pendientes_que_pasarian`), ya con montos; aquí solo se imprimen: al final
# de cada avión sus vuelos y, en la portada, el sello del cierre. Sin `cierre`
# el PDF es el de siempre (flujos de contenido idénticos).
def _linea_vuelo_cierre(v: RepartoCierreVuelo) -> str:
    """«#316 Mike Nelson $2,076.40 (cobrado $0.00, por cobrar $2,076.40)» —
    las partes que no vienen se omiten; '→' no se dibuja en Helvetica. Va
    dentro de un Paragraph: el texto del API se ESCAPA (un «&» o «<» en el
    nombre del cliente rompería el marcado). Con el `motivo` del API, los
    montos al corte y lo cobrado después («(cobrado al 30-sep $0.00,
    cobrado después $2,076.40 el 14-oct)», `cierre_mes.detalle_cobro`)."""
    partes = [f"#{v.folio}" if v.folio else "—"]
    if v.cliente:
        partes.append(v.cliente)
    if v.venta_usd is not None:
        partes.append(_usd(v.venta_usd))
    detalle = cierre_mes.detalle_cobro(
        v,
        cobrado=v.cobrado_usd,
        por_cobrar=v.por_cobrar_usd,
        despues=v.cobrado_posterior_usd,
        fecha_despues=v.fecha_cobro_posterior,
        dinero=_usd,
    )
    texto = " ".join(partes)
    if detalle:
        texto += f" ({', '.join(detalle)})"
    return escape(texto.replace("→", "-"))


def _parrafos_cierre_avion(
    cierre: RepartoCierre | None, matricula: str | None, estilo: ParagraphStyle
) -> list[Paragraph]:
    """Las líneas del cierre que tocan a ESTE avión (por `matricula`), una por
    lista con vuelos; `matricula` None ⇒ las que no son de ningún avión del
    reparto (externos)."""
    if cierre is None:
        return []
    parrafos: list[Paragraph] = []
    for attr, plantilla, mes_attr in cierre_mes.LISTAS_REPARTO:
        items = [v for v in getattr(cierre, attr) if v.matricula == matricula]
        if not items:
            continue
        titulo = cierre_mes.titulo_lista_reparto(plantilla, items, mes_attr)
        lineas = " · ".join(_linea_vuelo_cierre(v) for v in items)
        parrafos.append(Paragraph(f"<b>{titulo}:</b> {lineas}", estilo))
    return parrafos


# Grupo del desglose de gastos (mismas etiquetas que el panel).
_GRUPO_LABEL = {
    "DIRECTO": "Directo",
    "INDIRECTO": "Indirecto",
    "PERMISO": "Permiso",
    "FIJO": "Fijo (manual)",
    "EXCLUIDO": "Excluido",
}


def _filas_gastos_categoria(
    gastos: list[RepartoGastoCategoriaLinea],
) -> list[list[str]]:
    """Filas (sin encabezado) de la tabla 'Gastos por categoría' del avión.
    La categoría se imprime con su etiqueta amable (`etiqueta`, homologada
    con panel/app/API, 2-sep-2026) y, si el API no la manda, con el código
    tal cual."""
    return [
        [
            g.etiqueta or g.categoria,
            _GRUPO_LABEL.get(g.grupo or "", g.grupo or ""),
            str(g.count),
            _usd(g.usd),
        ]
        for g in gastos
    ]


def _recortar(texto: str, ancho_pt: float, fuente: str, tamano: float) -> str:
    """`texto` recortado (con «...») para que quepa en `ancho_pt` con esa
    letra: Helvetica no envuelve dentro de una celda de tabla."""
    if stringWidth(texto, fuente, tamano) <= ancho_pt:
        return texto
    while texto and stringWidth(texto + "...", fuente, tamano) > ancho_pt:
        texto = texto[:-1]
    return texto.rstrip(" ·") + "..."


def _linea_otro_ingreso(linea: RepartoOtroIngresoLinea) -> list[str]:
    """Renglón gris del detalle: «   ING-4 · Reembolsos… · concepto» y lo que
    sumó en dólares (o «sin T.C.» si no entró al saldo). Las partes vacías se
    omiten y «→» (Helvetica no lo dibuja) se vuelve «-»."""
    partes = [x.strip() for x in (linea.clave, linea.categoria, linea.concepto) if x and x.strip()]
    texto = "   " + " · ".join(partes).replace("→", "-")
    return [
        _recortar(texto, _ANCHO_ETIQUETA_PT, *_FUENTE_GRIS),
        _usd(linea.usd) if linea.usd is not None else "sin T.C.",
    ]


def _filas_otros_ingresos_avion(avion: RepartoAvion) -> tuple[list[list[str]], list[int]]:
    """Filas de la cascada del avión con sus otros ingresos asignados
    (13-oct-2026) y, aparte, los índices (relativos) de las filas GRISES. La
    fila (+) SUMA (no es gris) y solo sale con monto ≠ 0; después, hasta
    `MAX_LINEAS_OTROS_INGRESOS` ingresos y «y N más»; al final, los MXN sin
    tipo de cambio (no entraron al saldo). Sin nada de eso (API ≤ 0.0.74),
    ninguna fila: el PDF de siempre."""
    filas: list[list[str]] = []
    grises: list[int] = []
    if avion.otros_ingresos_avion_usd:
        filas.append([FILA_OTROS_INGRESOS_AVION, _usd(avion.otros_ingresos_avion_usd)])
    lineas = avion.otros_ingresos_avion
    for linea in lineas[:MAX_LINEAS_OTROS_INGRESOS]:
        grises.append(len(filas))
        filas.append(_linea_otro_ingreso(linea))
    if len(lineas) > MAX_LINEAS_OTROS_INGRESOS:
        grises.append(len(filas))
        filas.append([f"   y {len(lineas) - MAX_LINEAS_OTROS_INGRESOS} más", ""])
    if avion.otros_ingresos_avion_sin_tc_mxn:
        grises.append(len(filas))
        filas.append(
            [LINEA_OTROS_INGRESOS_SIN_TC, f"${avion.otros_ingresos_avion_sin_tc_mxn:,.2f} MXN"]
        )
    return filas, grises


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

    s_cierre = ParagraphStyle(
        "cierre", parent=base["Normal"], fontSize=7.5, textColor=MUTED, leading=10
    )
    s_aviso_cierre = ParagraphStyle(
        "aviso_cierre", fontSize=8, textColor=colors.HexColor("#b91c1c"), leading=11
    )

    story: list = []

    # ---- Encabezado ----
    # Sello del cierre de mes (16-oct-2026): «Mes cerrado el 16-oct-2026 por
    # Ale» bajo la fecha de generación; sin `cierre` (o mes abierto), nada.
    sello_cierre = cierre_mes.texto_estado_cierre(req.cierre)
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
                    f"Generado: {req.generado} (hora de Cancún, UTC−5)"
                    + (f"<br/><b>{escape(sello_cierre)}</b>" if sello_cierre else ""),
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
    # comisiones_venta_usd: 0 desde el 28-ago-2026 (regla A) y, desde la
    # regla de septiembre 2026 (API 0.0.65), las comisiones que absorbe cada
    # avión (banco + lo del vendedor: su provisión hasta el API 0.0.73, el
    # pago «Comisión del vendedor (avión)» desde el 0.0.74): restan antes del
    # saldo, así que «Gastos del periodo» las incluye y venta − gastos =
    # saldo.
    total_gastos = sum(
        _gastos_avion(a) + a.comisiones_venta_usd for a in req.aviones
    )
    # Otros ingresos de VuelaTour (TUAs/extras/pernocta cobrados): solo
    # informativo, FUERA del saldo y del reparto. Viene del API; si un API
    # viejo no manda el global, se re-suman los por avión solo para mostrar.
    total_otros = req.otros_ingresos_vuelatour_total_usd or sum(
        a.otros_ingresos_vuelatour_usd for a in req.aviones
    )
    # Otros ingresos ASIGNADOS a los aviones (13-oct-2026, API 0.0.75): ya
    # van dentro de cada `saldo_usd`; con ellos el resumen gana su columna
    # para que venta + otros ingresos de aviones − gastos = saldo. Sin ellos
    # (0), las 4 columnas de siempre, byte a byte.
    total_oi = sum(a.otros_ingresos_avion_usd for a in req.aviones)
    if total_oi:
        resumen = Table(
            [
                list(ENCABEZADOS_RESUMEN_CON_OTROS_INGRESOS),
                [
                    _usd(total_ingresos),
                    _usd(total_oi),
                    _usd(total_otros),
                    _usd(total_gastos),
                    _usd(total_saldo),
                ],
            ],
            colWidths=[34 * mm] * 5,
        )
        gris_otros = 2
    else:
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
        gris_otros = 1
    estilo_resumen: list = [
        ("BACKGROUND", (0, 0), (-1, 0), LIGHT),
        ("TEXTCOLOR", (0, 0), (-1, 0), MUTED),
        ("FONTSIZE", (0, 0), (-1, 0), 8),
        ("FONTSIZE", (0, 1), (-1, 1), 13),
        ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
        ("TEXTCOLOR", (0, 1), (-1, 1), BRAND),
        # Otros ingresos VuelaTour: informativo (gris), no se reparte.
        ("TEXTCOLOR", (gris_otros, 1), (gris_otros, 1), MUTED),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("BOX", (0, 0), (-1, -1), 0.5, LIGHT),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.white),
    ]
    if total_oi:
        # 5 columnas de 34 mm: los encabezados de 8 pt («Otros ingresos de
        # aviones» mide 93 pt) caben con menos relleno lateral.
        estilo_resumen += [
            ("LEFTPADDING", (0, 0), (-1, -1), 2),
            ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ]
    resumen.setStyle(TableStyle(estilo_resumen))
    story.append(resumen)
    # Composición cotizada de "Otros ingresos VuelaTour" (regla A: incluye la
    # comisión del vendedor) — solo si el API manda el desglose.
    desglose_txt = _desglose_txt(req.otros_ingresos_vuelatour_desglose)
    if desglose_txt:
        story.append(Spacer(1, 2 * mm))
        story.append(Paragraph(_nota_otros_ingresos(req, desglose_txt), s_sub))
    # Nota global del TC oficial de respaldo (29-ago-2026): cuántos vuelos y
    # gastos del periodo entraron convertidos con él (open.er-api / BCE).
    nota_global = _nota_tc_oficial_global(req.tc_oficial)
    if nota_global:
        story.append(Spacer(1, 2 * mm))
        story.append(Paragraph(nota_global, s_sub))
    # Cierre de mes (16-oct-2026): si el API detectó que los números de hoy
    # ya no son los del cierre, el aviso va arriba, en rojo.
    if req.cierre is not None and req.cierre.coincide_con_snapshot is False:
        story.append(Spacer(1, 2 * mm))
        story.append(Paragraph(f"AVISO: {cierre_mes.AVISO_SNAPSHOT}", s_aviso_cierre))
    story.append(Spacer(1, 8 * mm))

    # ---- Por aeronave ----
    if not req.aviones:
        story.append(Paragraph("Sin aeronaves en el periodo.", s_sub))

    for avion in req.aviones:
        story.append(
            _bloque_avion(
                avion,
                s_avion,
                linea_comision=linea_comision_vendedor(req),
                cierre=_parrafos_cierre_avion(req.cierre, avion.matricula, s_cierre),
            )
        )
        story.append(Spacer(1, 6 * mm))
    # Vuelos del cierre cuya matrícula no tiene bloque arriba (externos o un
    # avión sin dinero en el periodo): no se pierden en silencio.
    fuera = cierre_mes.matriculas_sin_bloque(req.cierre, (a.matricula for a in req.aviones))
    if fuera:
        bloque_fuera: list = [
            Paragraph(cierre_mes.TITULO_REPARTO_SIN_AVION, s_avion),
            Spacer(1, 2 * mm),
        ]
        for matricula in fuera:
            if matricula:
                bloque_fuera.append(Paragraph(f"<b>{escape(matricula)}</b>", s_sub))
            bloque_fuera += _parrafos_cierre_avion(req.cierre, matricula, s_cierre)
        story.append(KeepTogether(bloque_fuera))
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
    story.append(Paragraph(_pie(req), s_foot))

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


def _bloque_avion(
    avion: RepartoAvion,
    estilo_titulo: ParagraphStyle,
    *,
    linea_comision: str = LINEA_COMISION_VENDEDOR,
    cierre: list[Paragraph] | None = None,
) -> KeepTogether:
    """`cierre` (16-oct-2026): los párrafos del cierre de mes de ESTE avión
    (vuelos que pasan / vienen / pasarían), al final del bloque; vacío o
    None ⇒ el bloque de siempre."""
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
    # solo si el API manda el desglose. Con la regla de septiembre 2026 el
    # avión absorbe lo que se le paga al vendedor (resta abajo, en
    # Comisiones): la línea lo dice en lugar de «su pago no es costo del
    # avión» — `linea_comision` = «su provisión…» o, con el API 0.0.74,
    # «pago «(avión)» resta…».
    desg = avion.otros_ingresos_vuelatour_desglose
    if desg is not None and desg.comision_usd:
        info_rows.append(len(filas))
        filas.append(
            [
                linea_comision
                if absorbe_comisiones(avion)
                else _LINEA_COMISION_VENDEDOR_VUELATOUR,
                _usd(desg.comision_usd),
            ]
        )
    pendiente = avion.pendiente_cobro_usd or 0.0
    bruto = avion.pendiente_bruto_usd or 0.0
    if pendiente or bruto:
        texto = "Pendiente de cobro (parte avión) — se reparte al cobrar"
        if bruto > pendiente + 0.005:
            texto = f"Pendiente de cobro (parte avión) — deuda total del cliente {_usd(bruto)}"
        info_rows.append(len(filas))
        filas.append([texto, _usd(pendiente)])
    # Otros ingresos ASIGNADOS al avión (13-oct-2026, API 0.0.75): la fila (+)
    # SUMA (ya dentro de `saldo_usd`) y su detalle va en gris; sin ellos,
    # ninguna fila (el PDF de siempre).
    filas_oi, grises_oi = _filas_otros_ingresos_avion(avion)
    info_rows += [len(filas) + i for i in grises_oi]
    filas += filas_oi
    filas += [
        # Comisiones que absorbe el avión: desde la regla de septiembre 2026
        # (API 0.0.65) la parte del avión de la comisión bancaria + lo del
        # vendedor (su provisión hasta el API 0.0.73; el pago «Comisión del
        # vendedor (avión)» desde el 0.0.74). 0 (regla A del 28-ago, o sin
        # comisiones) ⇒ la fila NO se imprime.
        *(
            [[ETIQUETA_FILA, _usd(-avion.comisiones_venta_usd)]]
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

    # Gastos por categoría (solo si el API lo manda, 2-sep-2026): misma tabla
    # del desglose del panel; la categoría sale con su etiqueta amable.
    if avion.gastos_por_categoria:
        gfilas = [["Categoría de gasto", "Grupo", "#", "USD"]]
        gfilas += _filas_gastos_categoria(avion.gastos_por_categoria)
        gtabla = Table(gfilas, colWidths=[85 * mm, 35 * mm, 15 * mm, 35 * mm])
        gtabla.setStyle(
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
        bloque.append(gtabla)

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

    # Cierre de mes (16-oct-2026): al final del avión, sus vuelos diferidos.
    for parrafo in cierre or []:
        bloque.append(Spacer(1, 1.5 * mm))
        bloque.append(parrafo)

    return KeepTogether(bloque)


def _regla() -> Table:
    regla = Table([[""]], colWidths=[170 * mm], rowHeights=[0.1])
    regla.setStyle(TableStyle([("LINEBELOW", (0, 0), (-1, -1), 1, BRAND)]))
    return regla
