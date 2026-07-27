"""Tira imprimible de bitácora de tacómetros (WeasyPrint, HTML → PDF).

Réplica del formato manual del equipo (hoja "Imprimir planeador" de la
plantilla Excel, formato MONOMOTOR): una fila por vuelo con fecha, tacómetro
inicial, horas voladas, tacómetro final y la ruta en minúsculas. La tira es
angosta a propósito: se imprime, se recorta y se pega en la bitácora física
del avión (planeador y motor comparten lectura en monomotor).
"""

from datetime import datetime
from html import escape
from zoneinfo import ZoneInfo

from app.schemas.reportes import BitacoraTacoRequest

_NAVY = "#102a43"
_CANCUN = ZoneInfo("America/Cancun")
# Abreviaturas fijas: el locale del contenedor no garantiza es-MX.
_MESES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]


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


def _build_html(r: BitacoraTacoRequest) -> str:
    filas = "".join(
        "<tr>"
        f"<td class='c'>{escape(_fecha_corta(f.fecha))}</td>"
        f"<td class='n'>{_num(f.taco_inicial)}</td>"
        f"<td class='n'>{_num(f.horas)}</td>"
        f"<td class='n'>{_num(f.taco_final)}</td>"
        f"<td class='r'>{escape(f.ruta)}</td>"
        "</tr>"
        for f in r.filas
    )
    if not filas:
        filas = "<tr><td colspan='5' class='c muted'>Sin vuelos con tacómetro en el periodo</td></tr>"

    rango = ""
    if r.desde or r.hasta:
        rango = f"{_fecha_larga(r.desde)} — {_fecha_larga(r.hasta)}"
    generado = _fecha_larga(r.generado) if r.generado else ""
    modelo = f" · {escape(r.modelo)}" if r.modelo else ""

    # La tira mide ~11.6 cm de ancho (como el área de impresión de la
    # plantilla): cabe en la página de la bitácora tras recortarla. El thead
    # se repite solo en cada página (WeasyPrint).
    return f"""<!doctype html>
<html><head><meta charset="utf-8"><style>
  @page {{ size: letter; margin: 36px 40px; }}
  * {{ font-family: 'Helvetica Neue', Arial, sans-serif; color: {_NAVY}; }}
  .head {{ width: 440px; display: flex; justify-content: space-between;
           align-items: baseline; margin-bottom: 6px; }}
  .titulo {{ font-size: 12px; font-weight: 800; }}
  .sub {{ font-size: 9px; color: #627d98; }}
  table {{ width: 440px; border-collapse: collapse; font-size: 10px; }}
  th, td {{ border: 1px solid #9aa8b5; padding: 2.5px 6px; }}
  th {{ background: #f0f4f8; font-size: 8.5px; text-transform: uppercase;
        letter-spacing: .4px; text-align: center; }}
  td.c {{ text-align: center; white-space: nowrap; }}
  td.n {{ text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }}
  td.r {{ text-transform: lowercase; }}
  .muted {{ color: #829ab1; }}
  .pie {{ width: 440px; font-size: 8px; color: #829ab1; margin-top: 4px; }}
</style></head><body>
  <div class="head">
    <div class="titulo">Bitácora de tacómetro · {escape(r.matricula)}{modelo}</div>
    <div class="sub">{escape(rango)}</div>
  </div>
  <table>
    <thead><tr>
      <th style="width:58px">Fecha</th>
      <th style="width:66px">Tacómetro<br/>inicial</th>
      <th style="width:44px">Horas</th>
      <th style="width:66px">Tacómetro<br/>final</th>
      <th>Ruta</th>
    </tr></thead>
    <tbody>{filas}</tbody>
  </table>
  <div class="pie">Generado {escape(generado)} · hora de Cancún (UTC−5) · VuelaTour</div>
</body></html>"""


def render_bitacora_taco_pdf(req: BitacoraTacoRequest) -> bytes:
    from weasyprint import HTML  # import perezoso

    return HTML(string=_build_html(req)).write_pdf()
