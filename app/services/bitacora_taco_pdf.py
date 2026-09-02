"""Bitácoras de vuelo imprimibles por componente (WeasyPrint, HTML → PDF).

Réplica del formato manual del equipo (hojas "Imprimir planeador" y
"MOTOR - HÉLICE" de la plantilla Excel): una fila por vuelo con fecha,
tacómetro inicial, horas voladas, tacómetro final y la ruta en minúsculas.
El API manda UNA tira por bitácora física (planeador, motor, hélice) con el
tiempo acumulado del componente YA derivado de su base capturada — aquí solo
se pinta: cada tira en su propia página, angosta a propósito para recortarla
y pegarla en el libro que le toca. Los tiempos de planeador, motor y hélice
suelen ser distintos aunque compartan tacómetro, fechas y rutas.

Compatibilidad: sigue aceptando el payload LEGADO (``formato`` + ``filas``
planas con ``helice_*``) porque pyservices se despliega ANTES que el API; se
convierte a una sola tira equivalente en ``_tiras_normalizadas``.
"""

from datetime import datetime
from html import escape
from zoneinfo import ZoneInfo

from app.schemas.reportes import BitacoraTacoRequest, BitacoraTira, BitacoraTiraFila

_NAVY = "#102a43"
_CANCUN = ZoneInfo("America/Cancun")
# Abreviaturas fijas: el locale del contenedor no garantiza es-MX.
_MESES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
_SIN_FILAS = "Sin vuelos con tacómetro en el periodo"


def _fecha_corta(s: str | None) -> str:
    """ISO → "08-may" en hora Cancún (formato de la plantilla del equipo)."""
    if not s:
        return "—"
    try:
        dt = datetime.fromisoformat(s.replace("Z", "+00:00"))
        if dt.tzinfo is not None:
            dt = dt.astimezone(_CANCUN)
        return f"{dt.day:02d}-{_MESES[dt.month - 1]}"
    except ValueError:
        return s


def _fecha_larga(s: str | None) -> str:
    if not s:
        return "—"
    try:
        dt = datetime.fromisoformat(s.replace("Z", "+00:00"))
        if dt.tzinfo is not None:
            dt = dt.astimezone(_CANCUN)
        return dt.strftime("%d/%m/%Y")
    except ValueError:
        return s


def _num(v: float) -> str:
    return f"{v:,.1f}"


def _tiempo(v: float | None) -> str:
    """Tiempo del componente; None ⇒ "—" (se llena a mano en el libro)."""
    return _num(v) if v is not None else "—"


def _anio(desde: str | None, filas: list[BitacoraTiraFila]) -> str:
    """Año del encabezado "Fecha 2026" (del rango o de la primera fila)."""
    for s in (desde, filas[0].fecha if filas else None):
        if s and len(s) >= 4 and s[:4].isdigit():
            return s[:4]
    return ""


def _tiras_normalizadas(r: BitacoraTacoRequest) -> list[BitacoraTira]:
    """Tiras a renderizar, en el orden en que llegan.

    Payload nuevo ⇒ ``r.tiras`` tal cual (mandan aunque también vengan filas
    planas). Payload LEGADO (sin tiras) ⇒ UNA tira equivalente al PDF
    histórico:
      - MOTOR_HELICE: 7 columnas, tiempos = ``helice_inicial``/``helice_final``.
      - PLANEADOR (o cualquier otro valor): 5 columnas, solo tacómetro.
    Sin tiras ni filas también se devuelve la tira legada vacía, para que el
    PDF diga "Sin vuelos…" en lugar de salir en blanco.
    """
    if r.tiras:
        return list(r.tiras)
    bimotor = r.formato == "MOTOR_HELICE"
    filas = [
        BitacoraTiraFila(
            fecha=f.fecha,
            taco_inicial=f.taco_inicial,
            horas=f.horas,
            taco_final=f.taco_final,
            tiempo_inicial=f.helice_inicial if bimotor else None,
            tiempo_final=f.helice_final if bimotor else None,
            ruta=f.ruta,
        )
        for f in r.filas
    ]
    if bimotor:
        return [
            BitacoraTira(
                tipo="HELICE",
                titulo="Bitácora motor–hélice",
                etiqueta="Tiempo hélice",
                con_tiempo=True,
                filas=filas,
            )
        ]
    return [
        BitacoraTira(
            tipo="MOTOR",
            titulo="Bitácora de tacómetro",
            etiqueta="Tiempo motor",
            con_tiempo=False,
            filas=filas,
        )
    ]


def _tira_html(t: BitacoraTira, r: BitacoraTacoRequest, primera: bool) -> str:
    """Una tira = encabezado + nota + tabla + pie, en su propia página."""
    anio = _anio(r.desde, t.filas)
    fecha_head = f"Fecha<br/>{anio}" if anio else "Fecha"
    if t.con_tiempo:
        # Réplica de la hoja "MOTOR - HÉLICE" del equipo: tiempo del
        # componente junto a cada tacómetro.
        etq = escape(t.etiqueta)
        thead = (
            "<tr>"
            f"<th style='width:52px'>{fecha_head}</th>"
            "<th style='width:62px'>Tacómetro<br/>inicial</th>"
            f"<th style='width:70px'>{etq}<br/>inicial</th>"
            "<th style='width:44px'>Tiempo<br/>de vuelo</th>"
            "<th style='width:62px'>Tacómetro<br/>final</th>"
            f"<th style='width:70px'>{etq}<br/>final</th>"
            "<th>Ruta</th>"
            "</tr>"
        )
        filas = "".join(
            "<tr>"
            f"<td class='c'>{escape(_fecha_corta(f.fecha))}</td>"
            f"<td class='n'>{_num(f.taco_inicial)}</td>"
            f"<td class='n'>{_tiempo(f.tiempo_inicial)}</td>"
            f"<td class='n'>{_num(f.horas)}</td>"
            f"<td class='n'>{_num(f.taco_final)}</td>"
            f"<td class='n'>{_tiempo(f.tiempo_final)}</td>"
            f"<td class='r'>{escape(f.ruta)}</td>"
            "</tr>"
            for f in t.filas
        )
        cols = 7
    else:
        # Tira histórica de tacómetro (hoja "Imprimir planeador").
        thead = (
            "<tr>"
            f"<th style='width:58px'>{fecha_head}</th>"
            "<th style='width:66px'>Tacómetro<br/>inicial</th>"
            "<th style='width:44px'>Horas</th>"
            "<th style='width:66px'>Tacómetro<br/>final</th>"
            "<th>Ruta</th>"
            "</tr>"
        )
        filas = "".join(
            "<tr>"
            f"<td class='c'>{escape(_fecha_corta(f.fecha))}</td>"
            f"<td class='n'>{_num(f.taco_inicial)}</td>"
            f"<td class='n'>{_num(f.horas)}</td>"
            f"<td class='n'>{_num(f.taco_final)}</td>"
            f"<td class='r'>{escape(f.ruta)}</td>"
            "</tr>"
            for f in t.filas
        )
        cols = 5
    if not filas:
        filas = f"<tr><td colspan='{cols}' class='c muted'>{_SIN_FILAS}</td></tr>"

    titulo = f"{escape(t.titulo)} · {escape(r.matricula)}"
    if r.modelo:
        titulo += f" · {escape(r.modelo)}"
    rango = ""
    if r.desde or r.hasta:
        rango = f"{_fecha_larga(r.desde)} — {_fecha_larga(r.hasta)}"
    generado = _fecha_larga(r.generado) if r.generado else ""
    nota = f"<div class='nota'>{escape(t.nota)}</div>" if t.nota else ""
    clases = f"tira ancho-{cols}" + ("" if primera else " salto")

    return f"""<section class="{clases}">
  <div class="head">
    <div class="titulo">{titulo}</div>
    <div class="sub">{escape(rango)}</div>
  </div>
  {nota}
  <table>
    <thead>{thead}</thead>
    <tbody>{filas}</tbody>
  </table>
  <div class="pie">Generado {escape(generado)} · hora de Cancún (UTC−5) · VuelaTour</div>
</section>"""


def _build_html(r: BitacoraTacoRequest) -> str:
    tiras = _tiras_normalizadas(r)
    cuerpo = "\n".join(_tira_html(t, r, primera=i == 0) for i, t in enumerate(tiras))

    # Cada tira es angosta a propósito (como el área de impresión de la
    # plantilla): cabe en la página del libro tras recortarla. Cada tira va en
    # su propia página; el thead se repite solo en cada página (WeasyPrint).
    return f"""<!doctype html>
<html><head><meta charset="utf-8"><style>
  @page {{ size: letter; margin: 36px 40px; }}
  * {{ font-family: 'Helvetica Neue', Arial, sans-serif; color: {_NAVY}; }}
  .tira.salto {{ page-break-before: always; }}
  .ancho-7 .head, .ancho-7 .nota, .ancho-7 table, .ancho-7 .pie {{ width: 520px; }}
  .ancho-5 .head, .ancho-5 .nota, .ancho-5 table, .ancho-5 .pie {{ width: 440px; }}
  .head {{ display: flex; justify-content: space-between;
           align-items: baseline; margin-bottom: 6px; }}
  .titulo {{ font-size: 12px; font-weight: 800; }}
  .sub {{ font-size: 9px; color: #627d98; }}
  .nota {{ font-size: 9px; color: #627d98; margin: -2px 0 5px; }}
  table {{ border-collapse: collapse; font-size: 10px; }}
  th, td {{ border: 1px solid #9aa8b5; padding: 2.5px 6px; }}
  th {{ background: #f0f4f8; font-size: 8.5px; text-transform: uppercase;
        letter-spacing: .4px; text-align: center; }}
  td.c {{ text-align: center; white-space: nowrap; }}
  td.n {{ text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }}
  td.r {{ text-transform: lowercase; }}
  .muted {{ color: #829ab1; }}
  .pie {{ font-size: 8px; color: #829ab1; margin-top: 4px; }}
</style></head><body>
{cuerpo}
</body></html>"""


def render_bitacora_taco_pdf(req: BitacoraTacoRequest) -> bytes:
    from weasyprint import HTML  # import perezoso

    return HTML(string=_build_html(req)).write_pdf()
