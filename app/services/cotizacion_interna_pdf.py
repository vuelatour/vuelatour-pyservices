"""PDF de la COTIZACIÓN INTERNA v2 (8-sep-2026): UNA hoja carta para la oficina.

Feedback de administración (8-sep, con la foto de su formato de siempre):
SOLO lo de la COTIZACIÓN. La fecha protagonista es el DÍA DEL VUELO (la de
cotización/confirmación va en pequeño); los tramos se desglosan como en su
hoja — RUTA («Cancun-Merida») · FECHA («26-jun») · DISTANCIA MILLAS ·
TIEMPO VUELO (HRS) («1.19» en horas decimales desde el 24-sep-2026, incluye
calzos) · COSTO POR HORA VUELO · TOTAL POR
TRAMO, con fila TOTAL en USD —; TUAS solo las que SE COBRARON; desglose
canónico con comisión del vendedor, ajuste/redondeo, IVA y total USD/MXN;
cobros compactos con su comisión bancaria; notas internas. NADA de operación
(tacómetros, horas voladas, piloto del tramo, traslados) ni partición /
gastos / utilidad / CFDI: eso vive en el reporte del vuelo. ÚNICA excepción
(11-sep-2026, pedido de la oficina): debajo del «Avión cotizado» se pinta
«Avión utilizado: …» cuando el API manda el campo NUEVO `aeronave_utilizada`
— saber si salió el avión que se cotizó es dato de la cotización, no de la
operación. El legado `aeronave_operativa` de la v1 sigue sin pintarse.

NUNCA se manda al cliente: lleva banda «COTIZACIÓN INTERNA · uso exclusivo
de oficina» y pie «Documento interno · generado … · <usuario>».

Hermano de `cotizacion_pdf.py`: REUTILIZA (importa, no copia) el branding
(#dc2626 / #102a43), el logo, `_money`/`_monto` y los formatos de fecha. NO
importa `_estilos_base` a propósito: esa hoja es la del cliente (13 px,
márgenes de 2 cm, marca de agua y «Gracias por volar con VuelaTour» en el
pie) — aquí el CSS es propio y con @page de 9–12 mm. Tampoco aplica
`_mostrar_matricula`: en el documento interno la matrícula SIEMPRE se ve.

AIRE (11-sep-2026, captura de la oficina: «todo muy junto»): la v2 salió a
8.5 pt con interlineado 1.2 y celdas de 1 px, y la hoja se leía apretada
arriba con el tercio de abajo vacío. Ahora el cuerpo va a 9.5 pt / 1.3, las
tablas NUNCA bajan de 9.5 pt (encabezados y aclaraciones en gris, 8–8.5 pt:
son apoyo), las celdas tienen 3–4 px y cada bloque se separa del anterior
con el margen superior de su `h2`. Misma información, mismos números.

Aquí SOLO se pinta. El API arma el payload desde el desglose canónico v1.3
del snapshot: cada tramo trae su `tiempo_hr` (con calzo), su tarifa y su
`total_usd`; la diferencia contra la línea TIEMPO_VUELO canónica llega como
`tramos_ajuste_usd` con su motivo (hora mínima / sobrevuelo / horas pactadas
/ redondeo) y se pinta como una fila más para que Σ tramos + ajuste ==
«Servicio aéreo». Nunca se recalcula dinero; a lo sumo se formatea
(«1.19», «26-jun», «8.8570 % = $3,236.36») o se re-suma una columna
informativa (millas).

TIEMPO EN HORAS DECIMALES (24-sep-2026, API 0.0.33). Pedido del cliente:
«la parte de tiempo de vuelo, lo podemos manejar solo en decimales por
favor? … se nos hacen raros los tiempos». En hh:mm la CUN→PTU→CUN decía
«01:12» + «01:12» y TOTAL «02:23» (cada tramo valía 1.19166… h = 71.5 min).
Ahora la columna va con 2 decimales fijos y los tramos SUMAN el total: se
pintan `tiempo_horas` / `tramos_tiempo_total_horas` del API, repartidos por
residuo mayor (`repartirHorasDecimales`); con un payload viejo se reparte
aquí con el espejo `_repartir_horas_decimales`. `tiempo_hhmm` se ignora.

PODA de repeticiones (22-sep-2026, capturas del cliente sobre la #329, que
salía en DOS hojas). Se quitó lo que ya estaba dicho en otro renglón:
  · las filas «Método de cobro» y «T.C. USD→MXN» de la ficha — el método
    PREVISTO subió a la cabecera del bloque «Cobros» (ahí también vive el %
    de terminal pactado) y el T.C. sigue en la fila «Total MXN» y en la
    columna «Equiv. USD» del cobro;
  · el nombre largo del tramo («Cancún-Playa del Carmen»): la fila ahora
    abre con la ABREVIATURA «CUN–PCE» y las marcas en gris en la MISMA
    línea (el nombre largo vuelve solo como respaldo: fila consolidada o
    sin los dos IATA);
  · el detalle gris del IVA («16 % de $X · Pago facturable…»): con el orden
    nuevo el renglón de arriba YA es la base gravable, así que el gris solo
    aparece si de verdad difieren;
  · el gris del servicio aéreo («1.75 h × $1,650.00/hr · cobrable manual ·
    Σ tramos…»): eso mismo está en el bloque «Horas cotizadas» (Cobrables +
    Tarifa), donde la fila «Cobrables» ahora lleva el motivo Y el importe
    del ajuste — sin esa mudanza NO se puede borrar el gris: en 6 de cada
    10 cotizaciones la tabla cierra en «TOTAL Σ tramos» y el desglose en
    «Servicio aéreo», y el ajuste es lo único que los concilia.

Conceptos SIN IVA (22-sep-2026, mismo pedido): cuando hay algún concepto
exento (pernocta, extra con `aplica_iva=false`, comisión de terminal) Y el
vuelo lleva IVA, el renglón sobre el IVA pasa a ser la BASE GRAVABLE
(«Subtotal gravable» = `iva_base_usd`) y los exentos bajan DEBAJO del IVA,
bajo el rótulo «No causan IVA». La partición es de PRESENTACIÓN: ningún
monto cambia y la hace la función compartida `particionar_por_iva`
(`cotizacion_pdf`), que verifica las dos identidades (Σ gravables == base;
base + IVA + Σ exentos == total) y, si no cuadran, DEGRADA al layout de
siempre. Sin conceptos exentos la hoja sale exactamente igual que antes.

Presupuesto de UNA hoja tras la poda (≈ 738 pt / 984 px útiles, medido con
la fuente INCRUSTADA): base fija (cabecera + banda + resumen + ficha de 3
filas + desglose|horas + cobros con su thead) ≈ 530 px y cada FILA de más
≈ 40 px — cuenta cada tramo, cada TUA cobrada, cada extra, cada cobro y la
pareja ajuste + «Servicio aéreo» de la tfoot. O sea: UNA hoja mientras
`filas ≲ 11` (antes ≲ 6; el máximo real en producción es 10). Con más
filas la tabla pasa a una segunda hoja —el `thead` se repite y ninguna
fila se parte a la mitad—; «Notas internas» y el par desglose+horas llevan
`page-break-inside: avoid`, pero el bloque de cobros NO: si algo se
desborda, que bajen dos filas de cobro y no el bloque entero.

La fuente va INCRUSTADA (Arimo, la misma del PDF del cliente, vía
`_estilos_fuente`): sin ella el contenedor de Railway cae a su sans por
defecto (~8-12 % más ancho que el de esta Mac) y «cabe en una hoja» se
vuelve lotería — es la razón de que la #329 cupiera aquí y no allá.

HOJA INTERNA COMPARTIDA CON EL PANEL (2.1 del rediseño del cotizador,
22-sep-2026). La pantalla de la cotización pasa a ser esta hoja (el PDF del
cliente queda como salida), así que el documento se comparte igual que la
hoja del cliente desde el 8-sep: el CUERPO del CSS vive en
`app/static/cotizacion-interna.css` con TODO selector colgado de la raíz
`.cot-interna` (`CLASE_RAIZ`), el marcado se envuelve en
`<div class="cot-interna">` (`_cuerpo_interno_html`, fuente ÚNICA: la usan
el PDF y la vista previa) y el panel pide el MISMO cuerpo por
`POST /reportes/cotizacion-interna/preview-html` y el MISMO CSS por
`GET /reportes/cotizacion-interna/hoja.css` (fuente incrustada + cuerpo).
Solo se quedan del lado del papel las reglas `@page` y el margen del
`<body>` (`_estilos_page_interno`): en el panel no hay página que maquetar.
Regla: un estilo de esta hoja se cambia en el .css UNA vez y sale en el PDF,
en la vista previa y en el panel; jamás copiar CSS al panel ni renombrar las
clases del marcado.
"""

import math
from datetime import UTC, datetime
from functools import lru_cache
from html import escape

from app.schemas.reportes import (
    CotizacionInternaCobroPdf,
    CotizacionInternaLineaPdf,
    CotizacionInternaPdfRequest,
    CotizacionInternaTramoCotizadoPdf,
    CotizacionInternaTuaCobradaPdf,
)
from app.services._formato import _tc_txt
from app.services.cotizacion_grupo_pdf import _RE_PAX_SUFIJO, _monto, _plural
from app.services.cotizacion_pdf import (
    _CANCUN,
    _MESES_ES,
    _STATIC,
    ETIQUETA_BASE_GRAVABLE,
    ETIQUETA_SIN_IVA,
    ETIQUETA_SUBTOTAL,
    TOLERANCIA_USD,
    LineaIva,
    _estilos_fuente,
    _fecha_dia,
    _logo_data_uri,
    _money,
    particionar_por_iva,
)
from app.services.reporte_vuelo_pdf import _fecha

# Texto de la banda: es lo que distingue este documento del que va al cliente.
BANDA_INTERNA = "Cotización interna · uso exclusivo de oficina · no enviar al cliente"
# Clase RAÍZ del documento interno (2.1, 22-sep-2026): la lleva el <div> que
# envuelve el cuerpo en el PDF y en la vista previa, y el contenedor de la
# hoja interna en el panel; TODO selector de `cotizacion-interna.css` cuelga
# de ella. Hermana de `cotizacion_pdf.CLASE_RAIZ` ("cot-hoja", la del cliente).
CLASE_RAIZ = "cot-interna"
# Nota al pie de la tabla de tramos (regla del motor: 0.15 h por aterrizaje).
# Horas DECIMALES desde el 24-sep-2026 (API 0.0.33, pedido del cliente: «la
# parte de tiempo de vuelo, lo podemos manejar solo en decimales»): con hh:mm
# «01:12» + «01:12» no daba «02:23». Mismo texto en el panel (`NOTA_TRAMOS`).
NOTA_TRAMOS = "Tiempo de vuelo en horas decimales (1.50 = 1 h 30 min) e incluye calzos"
# Encabezado de la columna (el CSS lo pone en mayúsculas: «TIEMPO VUELO (HRS)»).
ENCABEZADO_TIEMPO = "Tiempo vuelo (hrs)"

# Truncado elegante de notas: prioridad a la cotización; las notas van al final.
_NOTAS_MAX_CHARS = 280
_NOTAS_MAX_LINEAS = 5

_SEMAFORO_CLASES = {"verde": "sem-verde", "amarillo": "sem-amarillo", "rojo": "sem-rojo"}


# ===== Formato =====


def _instante_cancun(s: str) -> datetime | None:
    """Instante ISO → datetime en hora Cancún; None si no se puede parsear."""
    try:
        dt = datetime.fromisoformat(s.replace("Z", "+00:00"))
    except ValueError:
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=UTC)
    return dt.astimezone(_CANCUN)


def _dia_pared(s: str | None) -> str | None:
    """'YYYY-MM-DD' (día de pared) se respeta tal cual — convertirlo movería el
    día; un instante ISO se lleva a su día en Cancún. None si no hay dato."""
    if not s:
        return None
    if len(s) == 10 and "T" not in s:
        return s
    dt = _instante_cancun(s)
    return dt.strftime("%Y-%m-%d") if dt else None


def _fecha_corta(s: str | None, con_anio: bool = False) -> str:
    """Instante ISO → 'dd/mm HH:MM' en hora Cancún (con año si se pide).
    Un día de pared ('2026-09-03') NO se convierte (movería el día)."""
    if not s:
        return "—"
    if len(s) == 10 and "T" not in s:
        try:
            d = datetime.fromisoformat(s)
        except ValueError:
            return escape(s)
        return d.strftime("%d/%m/%Y" if con_anio else "%d/%m")
    dt = _instante_cancun(s)
    if dt is None:
        return escape(s)
    return dt.strftime("%d/%m/%Y %H:%M" if con_anio else "%d/%m %H:%M")


def _dia_mes(s: str | None) -> str:
    """Día del tramo como lo escribe administración: '2026-06-26' → '26-jun'."""
    dia = _dia_pared(s)
    if not dia:
        return "—"
    try:
        d = datetime.strptime(dia[:10], "%Y-%m-%d")
    except ValueError:
        return escape(dia)
    return f"{d.day}-{_MESES_ES[d.month - 1]}"


def _dia_largo(s: str | None) -> str:
    """'2026-06-26' → '26 jun 2026' (fecha protagonista); '' si no hay."""
    return _fecha_dia(_dia_pared(s))


_MICRO_POR_HORA = 1_000_000
_MICRO_POR_CENTESIMA = _MICRO_POR_HORA // 100


def _micro_horas(h: float) -> int:
    """Horas → micro-horas ENTERAS (nunca negativo): se redondea sin flotantes.
    `floor(x + 0.5)` y no `round()`: es el `Math.round` del API (Python
    redondea los .5 al par) — el espejo tiene que ser exacto."""
    return max(0, math.floor(h * _MICRO_POR_HORA + 0.5))


def _centesimas_txt(c: int) -> str:
    """Centésimas de hora → '1.19' (2 decimales FIJOS)."""
    return f"{c // 100}.{c % 100:02d}"


def _horas_decimal(h: float) -> str:
    """Horas decimales → '1.19' con 2 decimales FIJOS, medio hacia arriba en
    aritmética ENTERA (1.005 → '1.01'; `f"{1.005:.2f}"` daría '1.00').
    Espejo EXACTO de `horasADecimal` del API (`tramos-costeados.util.ts`)."""
    return _centesimas_txt(
        (_micro_horas(h) + _MICRO_POR_CENTESIMA // 2) // _MICRO_POR_CENTESIMA
    )


def _repartir_horas_decimales(
    tiempos: list[float | None],
) -> tuple[list[str | None], str]:
    """Columna «TIEMPO VUELO (HRS)» con la SUMA CUADRADA — RESPALDO para un
    payload sin `tramos_tiempo_total_horas` (API previo al 0.0.33); con él,
    manda lo que calculó el API.

    Espejo EXACTO de `repartirHorasDecimales` del API: total =
    `_horas_decimal(Σ)`; cada tramo baja a su centésima de piso y las que
    faltan para el total se reparten de una en una al residuo MÁS GRANDE
    (empate ⇒ orden del tramo). Así Σ tramos mostrados == total mostrado y
    cada tramo queda a ≤ 0.01 de su propio redondeo. Un tiempo None es un
    tramo sin tiempo: celda None («—») que no suma. Mismos casos congelados en
    `tests/test_cotizacion_interna_pdf.py` y en los specs del API y el panel."""
    micros = [None if h is None or not math.isfinite(h) else _micro_horas(h) for h in tiempos]
    suma = sum(m for m in micros if m is not None)
    total = (suma + _MICRO_POR_CENTESIMA // 2) // _MICRO_POR_CENTESIMA
    cent: list[int | None] = [None if m is None else m // _MICRO_POR_CENTESIMA for m in micros]
    faltan = total - sum(c for c in cent if c is not None)
    candidatos = sorted(
        (
            (-(m % _MICRO_POR_CENTESIMA), i)
            for i, m in enumerate(micros)
            if m is not None and m % _MICRO_POR_CENTESIMA > 0
        ),
    )
    for _, i in candidatos:
        if faltan <= 0:
            break
        cent[i] = (cent[i] or 0) + 1
        faltan -= 1
    return [None if c is None else _centesimas_txt(c) for c in cent], _centesimas_txt(total)


def _millas(v: float | None) -> str:
    """Millas náuticas como en la hoja de administración: '157', '157.3'."""
    if v is None:
        return "—"
    txt = f"{v:,.1f}"
    return txt[:-2] if txt.endswith(".0") else txt


def _horas(v: float | None, dec: int = 2) -> str:
    return "—" if v is None else f"{v:.{dec}f} h"


def _pct_banco(v: float) -> str:
    """% de comisión bancaria como lo escribe la oficina: 2 decimales si no
    hay más precisión (2.90 %), 4 si la hay (8.8570 %, foto del cliente).
    El valor llega en PUNTOS porcentuales (cobro_vuelo.comision_banco_pct)."""
    dec = 2 if round(v, 2) == round(v, 4) else 4
    return f"{v:.{dec}f} %"


def _truncar(
    txt: str | None,
    max_chars: int = _NOTAS_MAX_CHARS,
    max_lineas: int = _NOTAS_MAX_LINEAS,
) -> str:
    """Recorta a N caracteres (en un espacio) y a N líneas, cerrando con «…».
    Devuelve texto plano (sin escapar)."""
    if not txt:
        return ""
    limpio = txt.strip()
    lineas = limpio.splitlines()
    cortado = False
    if len(lineas) > max_lineas:
        limpio = "\n".join(lineas[:max_lineas])
        cortado = True
    if len(limpio) > max_chars:
        corte = limpio[:max_chars]
        espacio = corte.rfind(" ")
        if espacio > max_chars * 0.6:
            corte = corte[:espacio]
        limpio = corte.rstrip(" ,;:·-")
        cortado = True
    return limpio + "…" if cortado else limpio


def _css_str(txt: str) -> str:
    """Texto seguro dentro de un string CSS entre comillas dobles."""
    return txt.replace("\\", "\\\\").replace('"', '\\"').replace("\n", " ")


def _generado_txt(r: CotizacionInternaPdfRequest) -> str:
    """'8 sep 2026 10:12' en hora Cancún: `generado_cancun` (pared, del API)
    manda; si no, `generado` (ISO) convertido; si no, ahora."""
    if r.generado_cancun and len(r.generado_cancun) >= 10:
        dia = _fecha_dia(r.generado_cancun[:10])
        hora = r.generado_cancun[11:16].strip()
        if dia:
            return f"{dia} {hora}".strip()
    dt = _instante_cancun(r.generado) if r.generado else None
    if dt is None:
        dt = datetime.now(UTC).astimezone(_CANCUN)
    return f"{_fecha_dia(dt.strftime('%Y-%m-%d'))} {dt.strftime('%H:%M')}"


def _pie_texto(r: CotizacionInternaPdfRequest) -> str:
    pie = f"Documento interno · generado {_generado_txt(r)} (hora Cancún)"
    if r.generado_por:
        pie += f" · {r.generado_por.strip()}"
    return pie


def _servicio_aereo_usd(r: CotizacionInternaPdfRequest) -> float:
    """Línea TIEMPO_VUELO canónica (ancla de la tabla de tramos); sin líneas,
    el escalar espejo `subtotal_vuelo_usd`. Nunca se suma aquí."""
    for ln in r.lineas:
        if (ln.clave or "").upper() == "TIEMPO_VUELO":
            return ln.monto_usd
    return r.subtotal_vuelo_usd


def _hay_ajuste(r: CotizacionInternaPdfRequest) -> bool:
    return abs(r.tramos_ajuste_usd) >= 0.005


# ===== Piezas HTML =====


def _kv(label: str, valor_html: str, clase: str = "") -> str:
    """Fila clave/valor. `valor_html` YA viene escapado por quien la arma."""
    cls = f' class="{clase}"' if clase else ""
    return f'<tr{cls}><td class="k">{escape(label)}</td><td class="v">{valor_html}</td></tr>'


def _op(txt: str) -> str:
    """Operación/aclaración en gris junto a un concepto."""
    return f' <span class="op">{escape(txt)}</span>' if txt else ""


def _tag(txt: str, clase: str = "") -> str:
    cls = f"tag {clase}".strip()
    return f'<span class="{cls}">{escape(txt)}</span>'


def _header_html(r: CotizacionInternaPdfRequest) -> str:
    logo = _logo_data_uri("logo-vuelatour-blanco.png")
    logo_html = f'<img class="logo" src="{logo}" alt=""/>' if logo else ""
    folio = f"Folio #{escape(r.folio)}" if r.folio else "Folio s/n"
    if r.version is not None:
        folio += f" · v{r.version}"
    estado = r.estado_label or r.estado or ""
    if r.tipo:
        estado = f"{estado} · {r.tipo}".strip(" ·")
    return f"""
  <div class="header">
    <div class="izq">{logo_html}<div>
      <div class="titulo">Cotización interna</div>
      <div class="sub">VuelaTour — Aero Charter Cancún</div></div></div>
    <div class="folio">{folio}<span class="sub">{escape(estado)}</span></div>
  </div>
  <div class="banda">{escape(BANDA_INTERNA)}</div>"""


def _avion_html(r: CotizacionInternaPdfRequest) -> str:
    """«Piper Seneca V · N4142R» (matrícula SIEMPRE visible) o el avión externo."""
    if r.es_externo and r.avion_externo:
        avion = escape(r.avion_externo) + _tag("Externo", "ambar")
        if r.operador_externo:
            avion += _op(f"operador {r.operador_externo}")
        return avion
    partes = [p for p in (r.aeronave_cotizada_modelo, r.aeronave_cotizada_matricula) if p]
    return escape(" · ".join(partes)) if partes else "—"


def _avion_utilizado_txt(r: CotizacionInternaPdfRequest) -> str:
    """«N4142R · Piper Seneca V» del avión REALMENTE utilizado (campo nuevo
    `aeronave_utilizada`, 11-sep-2026). TOLERANTE a propósito: acepta el
    texto ya armado por el API o un objeto {matricula, modelo}; sin dato —o
    con otra forma— devuelve "" y la segunda línea no se pinta."""
    v = r.aeronave_utilizada
    if isinstance(v, str):
        return v.strip()
    if isinstance(v, dict):
        partes = [
            str(v.get(k)).strip()
            for k in ("matricula", "modelo")
            if v.get(k) not in (None, "")
        ]
        return " · ".join(p for p in partes if p)
    return ""


def _resumen_html(r: CotizacionInternaPdfRequest) -> str:
    """Franja de resumen: FECHA DEL VUELO (protagonista, grande), avión
    COTIZADO (con la matrícula, siempre) y ruta cotizada con pasajeros. Si
    el API manda `aeronave_utilizada`, debajo del cotizado va la segunda
    línea «Avión utilizado: …» — la oficina necesita ver los dos cuando no
    salió el avión que se cotizó; si el API avisa que DIFIEREN, esa línea va
    en ámbar con la marca «Distinto al cotizado»."""
    fecha = _dia_largo(r.fecha_vuelo) or "Por definir"
    fin = _dia_largo(r.fecha_vuelo_fin)
    fin_html = f'<div class="sub">al {escape(fin)}</div>' if fin else ""
    utilizada = _avion_utilizado_txt(r)
    # ⚠ Se cotizó con un avión y vuela OTRO: lo decide el API comparando por
    # ID (`aeronave_cotizada_vs_utilizada_difiere`) — el texto no basta,
    # dos aviones pueden compartir modelo. La oficina tiene que verlo de un
    # vistazo, no deducirlo comparando dos matrículas en gris.
    difiere = r.aeronave_cotizada_vs_utilizada_difiere is True
    utilizada_html = (
        f'<div class="sub{" ambar" if difiere else ""}">Avión utilizado: '
        f"{escape(utilizada)}"
        + (_tag("Distinto al cotizado", "ambar") if difiere else "")
        + "</div>"
        if utilizada
        else ""
    )
    ruta = escape(r.ruta) if r.ruta else "—"
    ruta += _op(f"· {_plural(r.pasajeros, 'pasajero', 'pasajeros')}")
    return f"""
  <table class="resumen"><tr>
    <td class="fecha"><div class="lbl">Fecha del vuelo</div>
      <div class="big">{escape(fecha)}</div>{fin_html}</td>
    <td><div class="lbl">Avión cotizado</div><div class="med">{_avion_html(r)}</div>
      {utilizada_html}</td>
    <td><div class="lbl">Ruta cotizada</div><div class="med">{ruta}</div></td>
  </tr></table>"""


def _metodo_previsto_txt(r: CotizacionInternaPdfRequest) -> str:
    """«previsto: Transferencia · comisión terminal 8.86 %»: el método de cobro
    PACTADO, que el 22-sep-2026 dejó de tener fila propia en la ficha y pasó a
    la cabecera del bloque «Cobros» (cuesta cero altura y queda junto a los
    cobros reales). NO es decorativo: es el campo que decide si la cotización
    lleva IVA 16 % o 0 %, y es el ÚNICO lugar del documento donde se ve el
    porcentaje de terminal pactado (distinto de la comisión BANCARIA de cada
    cobro). Sin método → '' (la cabecera queda como antes)."""
    metodo = (r.metodo_cobro_label or r.metodo_cobro or "").strip()
    if not metodo:
        return ""
    txt = f"previsto: {metodo}"
    if r.comision_billpocket_pct:
        txt += f" · comisión terminal {r.comision_billpocket_pct:g} %"
    return txt


def _ficha_html(r: CotizacionInternaPdfRequest) -> str:
    """Ficha en dos columnas: cliente/condiciones | vendedor/tripulación/marcas.
    La fecha de cotización va al final y en pequeño (no es protagonista).

    22-sep-2026: SIN las filas «Método de cobro» y «T.C. USD→MXN» (el cliente
    las marcó como repetidas). El método vive ahora en la cabecera de
    «Cobros» (`_metodo_previsto_txt`) y el T.C. en la fila «Total MXN» del
    desglose y en la columna «Equiv. USD» del cobro."""
    izq: list[str] = []
    cliente = escape(r.cliente or "Cliente")
    if r.es_broker:
        cliente += _tag("Broker")
    if r.es_interno:
        cliente += _tag("Cliente interno", "ambar")
    izq.append(_kv("Cliente", cliente))
    razon = escape(r.razon_social) if r.razon_social else "—"
    if r.cliente_rfc:
        razon += _op(f"RFC {r.cliente_rfc}")
    izq.append(_kv("Razón social", razon))
    tarifa = escape(r.tarifa_tipo_label or r.tarifa_tipo or "—")
    if r.tarifa_hora_usd is not None:
        tarifa += f" · {_money(r.tarifa_hora_usd)}/hr"
    marcas = []
    if r.tarifa_preferencial:
        marcas.append("tarifa preferencial del cliente")
    if r.tarifa_override:
        marcas.append("tarifa manual")
    tarifa += _op(" · ".join(marcas))
    izq.append(_kv("Tarifa", tarifa))

    der: list[str] = []
    vendedor = escape(r.vendedor) if r.vendedor else "—"
    if r.cotizado_por:
        vendedor += _op(f"cotizó {r.cotizado_por}")
    der.append(_kv("Vendedor", vendedor))
    if r.piloto or r.copiloto:
        tripulacion = escape(r.piloto) if r.piloto else "—"
        if r.copiloto:
            tripulacion += f" · copiloto {escape(r.copiloto)}"
        der.append(_kv("Piloto", tripulacion))
    tags: list[str] = []
    if r.cotizacion_abierta:
        tags.append(_tag("Cotización abierta", "ambar"))
    if r.itinerario_operativo:
        tags.append(_tag("Itinerario operativo"))
    if r.grupo_folio:
        g = f"Grupo {r.grupo_folio}"
        if r.grupo_posicion and r.grupo_total_aviones:
            g += f" · avión {r.grupo_posicion} de {r.grupo_total_aviones}"
        tags.append(_tag(g))
    if r.combinado_con_folio:
        tags.append(_tag(f"Combinado con #{r.combinado_con_folio}"))
    if tags:
        der.append(_kv("Marcas", "".join(tags)))
    cotizada = _fecha(r.fecha)
    if r.fecha_confirmacion:
        cotizada += f" · confirmada {_fecha(r.fecha_confirmacion)}"
    der.append(_kv("Cotizada", escape(cotizada), "small"))

    return (
        '<table class="cols"><tr>'
        f'<td class="col"><table class="kv">{"".join(izq)}</table></td>'
        f'<td class="col"><table class="kv">{"".join(der)}</table></td>'
        "</tr></table>"
    )


def _tramo_fila(t: CotizacionInternaTramoCotizadoPdf, tiempo: str) -> str:
    """Una fila de la tabla de administración: RUTA · FECHA · MILLAS · TIEMPO ·
    COSTO/HR · TOTAL.

    `tiempo` llega YA en horas decimales («1.19», de `_tiempos_de_tabla`): la
    celda de un tramo depende de TODA la tabla (suma cuadrada), así que no se
    formatea fila por fila.

    La celda RUTA abre con la ABREVIATURA «CUN–PCE» (22-sep-2026: el cliente
    marcó el nombre largo como repetido) y las marcas —ferry, pernocta— van
    en gris en la MISMA línea, sin `<br>`: la fila baja de dos renglones a
    uno (89 → 40 px), que es de dónde sale la mitad de la hoja recuperada.

    RESPALDO (obligatorio): la abreviatura solo existe si vienen los DOS
    IATA y la fila no es la consolidada; si no, se conserva el nombre largo
    de siempre (`ruta` o el join de nombres) — promoverla sin respaldo
    dejaría la fila consolidada SIN ruta."""
    largo = t.ruta or "-".join(
        p for p in (t.origen_nombre or t.origen_iata, t.destino_nombre or t.destino_iata) if p
    )
    abrev = (
        f"{t.origen_iata}–{t.destino_iata}"
        if (t.origen_iata and t.destino_iata and not t.consolidado)
        else ""
    )
    detalles: list[str] = []
    if t.consolidado:
        detalles.append("tramos consolidados")
    if t.es_ferry:
        detalles.append("ferry")
    if t.pernocta:
        detalles.append("pernocta" + (f" {_money(t.pernocta_usd)}" if t.pernocta_usd else ""))
    ruta_html = escape(abrev or largo or "—")
    if detalles:
        ruta_html += f'<span class="muted"> · {escape(" · ".join(detalles))}</span>'
    tarifa = "—" if t.tarifa_hora_usd is None else _money(t.tarifa_hora_usd)
    return (
        f'<tr><td class="ruta">{ruta_html}</td>'
        f"<td>{_dia_mes(t.fecha)}</td>"
        f'<td class="num">{_millas(t.millas)}</td>'
        f'<td class="num">{escape(tiempo)}</td>'
        f'<td class="num">{tarifa}</td>'
        f'<td class="num">{_monto(t.total_usd)}</td></tr>'
    )


def _tiempos_de_tabla(
    r: CotizacionInternaPdfRequest, tramos: list[CotizacionInternaTramoCotizadoPdf]
) -> tuple[list[str], str]:
    """Celdas de «TIEMPO VUELO (HRS)» (en el orden de `tramos`) y su fila TOTAL.

    API 0.0.33+ (viaja `tramos_tiempo_total_horas`): se PINTA lo que mandó el
    API (`tiempo_horas` por tramo, ya cuadrado con el total; None ⇒ «—»).
    Payload anterior: se reparte aquí con el MISMO criterio
    (`_repartir_horas_decimales`) sobre `tiempo_hr`, para que la tabla también
    cuadre durante el despliegue."""
    if r.tramos_tiempo_total_horas:
        return [t.tiempo_horas or "—" for t in tramos], r.tramos_tiempo_total_horas
    celdas, total = _repartir_horas_decimales([t.tiempo_hr for t in tramos])
    return [c or "—" for c in celdas], total


def _tramos_html(r: CotizacionInternaPdfRequest) -> str:
    """Tabla de tramos EXACTAMENTE como la hoja de administración + fila TOTAL
    (Σ tramos, del API) y, solo si hay diferencia contra la línea canónica,
    la fila de ajuste con su motivo y el «Servicio aéreo» resultante."""
    tramos = sorted(r.tramos_cotizados, key=lambda t: t.orden)
    if not tramos:
        return (
            "<h2>Tramos cotizados</h2>"
            '<table class="grid tramos"><tbody><tr><td class="muted">'
            "Sin tramos cotizados.</td></tr></tbody></table>"
        )
    # TIEMPO en horas decimales con la suma cuadrada (Σ celdas == TOTAL).
    tiempos, tiempo_total = _tiempos_de_tabla(r, tramos)
    filas = "".join(_tramo_fila(t, tiempo) for t, tiempo in zip(tramos, tiempos, strict=True))
    # Millas: re-suma informativa de la columna (solo si todos los tramos las traen).
    millas = (
        _millas(sum(t.millas for t in tramos if t.millas is not None))
        if all(t.millas is not None for t in tramos)
        else ""
    )
    pie = (
        '<tr class="total"><td>TOTAL</td><td></td>'
        f'<td class="num">{millas}</td><td class="num">{escape(tiempo_total)}</td><td></td>'
        f'<td class="num">{_monto(r.tramos_total_usd)} USD</td></tr>'
    )
    if _hay_ajuste(r):
        motivo = (r.tramos_ajuste_motivo or "Ajuste de horas").strip()
        pie += (
            f'<tr class="ajuste"><td colspan="5">{escape(motivo)}'
            f'{_op("servicio aéreo cotizado − Σ tramos")}</td>'
            f'<td class="num">{_monto(r.tramos_ajuste_usd)}</td></tr>'
            '<tr class="total"><td colspan="5">Servicio aéreo</td>'
            f'<td class="num">{_monto(_servicio_aereo_usd(r))} USD</td></tr>'
        )
    notas = [NOTA_TRAMOS + (f" ({r.calzos_hr:g} h en total)" if r.calzos_hr else "")]
    notas.append("distancia en millas náuticas")
    if any(t.consolidado for t in tramos):
        notas.append("cotización sin desglose por tramo: una sola fila con los totales")
    return f"""
  <h2>Tramos cotizados</h2>
  <table class="grid tramos"><thead><tr>
    <th>Ruta</th><th>Fecha</th><th class="num">Distancia millas</th>
    <th class="num">{ENCABEZADO_TIEMPO}</th><th class="num">Costo por hora vuelo</th>
    <th class="num">Total por tramo</th>
  </tr></thead><tbody>{filas}</tbody><tfoot>{pie}</tfoot></table>
  <div class="nota">{escape(" · ".join(notas))}.</div>"""


def _fila_desglose(concepto: str, op: str, monto_html: str) -> str:
    return (
        f'<tr><td class="lbl">{escape(concepto)}{_op(op)}</td>'
        f'<td class="val">{monto_html}</td></tr>'
    )


def _tua_fila(t: CotizacionInternaTuaCobradaPdf) -> str:
    """«TUA CUN · 4 pax × $20.85» — solo TUAS cobradas (las exentas no viajan)."""
    concepto = f"TUA {t.iata}".strip()
    op = f"{t.pax:g} pax × {_money(t.unitario)}"
    if (t.moneda or "USD").upper() == "MXN":
        op += f" MXN = ${t.total_nativo:,.2f} MXN"
        if t.tc_aplicado:
            op += f" · T.C. {_tc_txt(t.tc_aplicado)}"
    return _fila_desglose(concepto, op, _monto(t.total_usd))


def _concepto_operacion(
    ln: CotizacionInternaLineaPdf, r: CotizacionInternaPdfRequest
) -> tuple[str, str]:
    """(concepto, operación en gris) de una línea canónica. La operación se
    arma SOLO con cantidad/unitario que manda el API; el monto no se toca."""
    clave = (ln.clave or "").upper()
    concepto = (ln.concepto or ln.clave or "Concepto").strip()
    if clave == "TIEMPO_VUELO":
        # Etiqueta de administración, SIN gris (22-sep-2026): «1.75 h ×
        # $1,650.00/hr», «hora mínima», «cobrable manual» y «Σ tramos ±
        # ajuste» son exactamente las filas «Cobrables» y «Tarifa» del
        # bloque «Horas cotizadas» — el cliente marcó esta línea como
        # repetida. El enlace con la tabla de tramos (Σ tramos ± ajuste) NO
        # se pierde: vive en el `.op` de «Cobrables» (ver `_horas_html`).
        return "Servicio aéreo", ""
    if clave == "COMISION_VENDEDOR":
        nombre = (r.comision_vendedor_nombre or "").strip()
        if nombre and nombre.lower() not in concepto.lower():
            concepto += f" · {nombre}"
        partes: list[str] = []
        modo = (r.comision_vendedor_modo or "").upper()
        if modo == "POR_HORA" and ln.cantidad is not None and ln.unitario is not None:
            partes.append(f"{ln.cantidad:.2f} h × {_money(ln.unitario)}/hr")
        elif modo == "FIJA":
            partes.append("fija")
        if r.pago_vendedor_usd is not None:
            con_iva = " c/IVA" if r.iva_comision_vendedor_usd else ""
            partes.append(f"pago al vendedor{con_iva} {_money(r.pago_vendedor_usd)}")
        return concepto, " · ".join(partes)
    if clave == "AJUSTE":
        if r.total_pactado_usd is not None:
            return concepto, f"a precio pactado {_money(r.total_pactado_usd)}"
        if ln.monto_usd < 0:
            return concepto, "descuento"
        if r.redondeo_auto_usd:
            return concepto, "redondeo automático"
        return concepto, ""
    if ln.cantidad is None or ln.unitario is None or "×" in concepto:
        return concepto, ""
    unit = _money(ln.unitario)
    if clave == "TUAS":
        concepto = _RE_PAX_SUFIJO.sub("", concepto) or concepto
        op = f"{ln.cantidad:g} pax × {unit}"
    else:
        op = f"{ln.cantidad:g} × {unit}"
    if (ln.moneda or "USD").upper() == "MXN":
        op += " MXN"
        if ln.monto_nativo is not None:
            op += f" = ${ln.monto_nativo:,.2f} MXN"
        if ln.tc_aplicado:
            op += f" · T.C. {_tc_txt(ln.tc_aplicado)}"
    if clave == "EXTRA" and ln.aplica_iva is False:
        op += " · sin IVA"
    return concepto, op


def _lineas_fallback(r: CotizacionInternaPdfRequest) -> list[CotizacionInternaLineaPdf]:
    """Sin líneas canónicas (API viejo / sin snapshot): una por componente ≠ 0
    con los escalares espejo del vuelo."""
    lineas: list[CotizacionInternaLineaPdf] = []
    if r.subtotal_vuelo_usd:
        lineas.append(
            CotizacionInternaLineaPdf(
                clave="TIEMPO_VUELO",
                concepto="Servicio aéreo",
                monto_usd=r.subtotal_vuelo_usd,
                cantidad=r.tiempo_cobrable_hr,
                unitario=r.tarifa_hora_usd,
            )
        )
    if r.tuas_usd:
        lineas.append(
            CotizacionInternaLineaPdf(clave="TUAS", concepto="TUAS", monto_usd=r.tuas_usd)
        )
    if r.extras_total_usd:
        lineas.append(
            CotizacionInternaLineaPdf(
                clave="EXTRA", concepto="Extras", monto_usd=r.extras_total_usd
            )
        )
    if r.comision_vendedor_usd:
        lineas.append(
            CotizacionInternaLineaPdf(
                clave="COMISION_VENDEDOR",
                concepto="Comisión del vendedor",
                monto_usd=r.comision_vendedor_usd,
            )
        )
    if r.ajuste_final_usd:
        lineas.append(
            CotizacionInternaLineaPdf(
                clave="AJUSTE",
                concepto="Descuento" if r.ajuste_final_usd < 0 else "Ajuste",
                monto_usd=r.ajuste_final_usd,
            )
        )
    if r.viaticos_pernocta_usd:
        lineas.append(
            CotizacionInternaLineaPdf(
                clave="PERNOCTA",
                concepto="Viáticos por pernocta",
                monto_usd=r.viaticos_pernocta_usd,
            )
        )
    return lineas


def _sin_iva(ln: CotizacionInternaLineaPdf) -> bool:
    """¿La línea canónica NO causa IVA? La PERNOCTA por su clave (el motor la
    publica «Viáticos por pernocta (sin IVA)» y la suma después del IVA) y
    cualquier línea que el API marque `aplica_iva=False`: extras exentos y la
    comisión de terminal sintetizada. `None` = el API no lo dice → gravable,
    como siempre."""
    return (ln.clave or "").upper() == "PERNOCTA" or ln.aplica_iva is False


def _desglose_html(r: CotizacionInternaPdfRequest) -> str:
    """Desglose interno: líneas canónicas en su orden (IVA aparte, en los
    totales), TUAS solo las cobradas (detalle `tuas_cobradas` en el lugar de
    las líneas TUAS), subtotal, IVA, total USD y MXN.

    22-sep-2026, dos cambios de PRESENTACIÓN (ningún monto se mueve):
      · los conceptos que NO causan IVA bajan DEBAJO del IVA bajo el rótulo
        «No causan IVA» y el renglón de arriba pasa a ser «Subtotal
        gravable» (= `iva_base_usd`), que es sobre lo que se calcula el
        16 %. Lo decide `particionar_por_iva` (compartida con el PDF del
        cliente y el de grupo), que degrada al layout de siempre si las dos
        identidades no cuadran;
      · el gris del IVA se reduce: la nota larga del API (`iva_nota`) ya no
        se pinta —el cliente la marcó como repetida— y «16 % de $X» solo
        aparece si el renglón de arriba NO es la base (con la partición
        activa lo es por construcción, así que no se pinta)."""
    lineas = list(r.lineas) or _lineas_fallback(r)
    cuerpo: list[LineaIva] = []
    tuas_pintadas = False
    for ln in lineas:
        clave = (ln.clave or "").upper()
        if clave == "IVA":
            continue
        if clave == "TUAS":
            # Exentas (línea sintética legado o monto 0) NO se muestran.
            if ln.exento or (not ln.monto_usd and not r.tuas_cobradas):
                continue
            if r.tuas_cobradas:
                # El detalle por aeropuerto sustituye a las líneas TUAS 1:1
                # (misma suma: `total_usd` de cada fila == línea canónica).
                if not tuas_pintadas:
                    cuerpo.extend(
                        LineaIva(t.total_usd, False, _tua_fila(t)) for t in r.tuas_cobradas
                    )
                    tuas_pintadas = True
                continue
        concepto, op = _concepto_operacion(ln, r)
        cuerpo.append(
            LineaIva(ln.monto_usd, _sin_iva(ln), _fila_desglose(concepto, op, _monto(ln.monto_usd)))
        )

    particion = particionar_por_iva(cuerpo, r.iva_base_usd, r.iva_usd, r.total_usd, r.iva_pct)
    filas = [ln.fila for ln in particion.gravables]
    if not filas and not particion.exentos:
        filas.append(
            '<tr><td class="lbl muted" colspan="2">Sin desglose (cotización sin precio).</td></tr>'
        )

    if particion.activa:
        # El renglón sobre el IVA es la BASE GRAVABLE, no «total − IVA».
        filas.append(
            f'<tr class="sub-row"><td class="lbl">{ETIQUETA_BASE_GRAVABLE}</td>'
            f'<td class="val">{_monto(particion.base_usd)}</td></tr>'
        )
        subtotal_impreso: float | None = particion.base_usd
    else:
        # El subtotal (total − IVA) lo manda el API; sin dato se deja en blanco.
        subtotal = "—" if r.subtotal_usd is None else _monto(r.subtotal_usd)
        filas.append(
            f'<tr class="sub-row"><td class="lbl">{ETIQUETA_SUBTOTAL}</td>'
            f'<td class="val">{subtotal}</td></tr>'
        )
        subtotal_impreso = r.subtotal_usd
    # Gris del IVA SOLO si el renglón de arriba no es ya la base (si lo es,
    # «16 % de $2,987.50» repetiría el número que está justo encima).
    difiere_de_la_base = r.iva_base_usd is not None and (
        subtotal_impreso is None or abs(subtotal_impreso - r.iva_base_usd) >= TOLERANCIA_USD
    )
    iva_op = f"{r.iva_pct:g} % de {_money(r.iva_base_usd)}" if difiere_de_la_base else ""
    filas.append(
        f'<tr><td class="lbl">IVA {r.iva_pct:g} %{_op(iva_op)}</td>'
        f'<td class="val">{_money(r.iva_usd)}</td></tr>'
    )
    if particion.exentos:
        filas.append(
            f'<tr class="exentos-row"><td class="lbl" colspan="2">{ETIQUETA_SIN_IVA}</td></tr>'
        )
        filas.extend(ln.fila for ln in particion.exentos)
    filas.append(
        f'<tr class="total-row"><td>Total USD</td><td class="val">{_monto(r.total_usd)}</td></tr>'
    )
    if r.total_mxn is not None:
        mxn_op = f"T.C. {_tc_txt(r.tc_usd_mxn)}" if r.tc_usd_mxn else ""
        if r.mxn_nativos:
            mxn_op = f"{mxn_op} · incluye ${r.mxn_nativos:,.2f} MXN nativos".strip(" ·")
        filas.append(
            f'<tr class="mxn-row"><td>Total MXN{_op(mxn_op)}</td>'
            f'<td class="val">{_monto(r.total_mxn)} MXN</td></tr>'
        )
    motor = []
    if r.version_motor:
        motor.append(f"motor {r.version_motor}")
    if r.calculado_at:
        motor.append(f"calculado {_fecha(r.calculado_at)}")
    if motor:
        filas.append(f'<tr><td colspan="2" class="muted">{escape(" · ".join(motor))}</td></tr>')
    return (
        "<h2>Desglose de la cotización</h2>"
        f'<table class="totales"><tbody>{"".join(filas)}</tbody></table>'
    )


def _horas_html(r: CotizacionInternaPdfRequest) -> str:
    """Horas del snapshot (vuelo + calzos + sobrevuelo → cotizadas → cobrables)
    y tarifa: explican la fila de ajuste de la tabla. Solo lo que venga.

    La fila «Cobrables» carga desde el 22-sep-2026 el MOTIVO y el IMPORTE del
    ajuste («Horas pactadas 1.75 h: +$165.00 sobre Σ tramos»). Es el
    refuerzo que permite borrar el gris del «Servicio aéreo»: sin él, en las
    6 de cada 10 cotizaciones que tienen ajuste la hoja diría «TOTAL
    $2,722.50» en la tabla y «Servicio aéreo $2,887.50» en el desglose sin
    nada que los concilie. Cuesta 0 px: es un `.op` en una fila que ya
    existía."""
    filas: list[str] = []
    for lbl, v in (("Vuelo", r.vuelo_hr), ("Calzos", r.calzos_hr), ("Sobrevuelo", r.sobrevuelo_hr)):
        if v:
            filas.append(_kv(lbl, _horas(v)))
    if r.horas_cotizadas_hr is not None:
        filas.append(_kv("Cotizadas", _horas(r.horas_cotizadas_hr)))
    if r.tiempo_cobrable_hr is not None:
        notas = []
        if r.hora_minima_aplicada:
            notas.append("hora mínima aplicada")
        if r.cobrable_override:
            notas.append("cobrable manual")
        if _hay_ajuste(r):
            signo = "+" if r.tramos_ajuste_usd > 0 else "−"
            motivo = (r.tramos_ajuste_motivo or "Ajuste de horas").strip()
            notas.append(
                f"{motivo}: {signo}{_money(abs(r.tramos_ajuste_usd))} sobre Σ tramos"
            )
        cobrables = f"<b>{_horas(r.tiempo_cobrable_hr)}</b>" + _op(" · ".join(notas))
        filas.append(_kv("Cobrables", cobrables))
    if r.tarifa_hora_usd is not None:
        filas.append(_kv("Tarifa", f"{_money(r.tarifa_hora_usd)}/hr"))
    if not filas:
        return ""
    return f'<h2>Horas cotizadas</h2><table class="kv">{"".join(filas)}</table>'


def _cobro_fila(c: CotizacionInternaCobroPdf, con_usd: bool) -> str:
    metodo = escape(c.metodo_label or c.metodo or "—")
    sub: list[str] = []
    if c.es_reembolso or c.monto < 0:
        sub.append("Reembolso")
    ref_partes = [p for p in (c.referencia, c.cuenta_destino) if p]
    if ref_partes:
        sub.append(" · ".join(ref_partes))
    if c.sobre_grupo_folio:
        s = f"Sobre {c.sobre_grupo_folio}"
        if c.sobre_grupo_monto_total is not None:
            s += f" {_money(c.sobre_grupo_monto_total)} {c.sobre_grupo_moneda or ''}".rstrip()
        if c.grupo_factor is not None:
            s += f" × {round(c.grupo_factor * 100, 2):g} %"
        sub.append(s)
    if c.notas:
        sub.append(_truncar(c.notas, 60, 1))
    if sub:
        metodo += f'<br><span class="muted">{escape(" · ".join(sub))}</span>'
    moneda = (c.moneda or "USD").upper()
    sufijo = "" if moneda == "USD" else f' <span class="muted">{escape(moneda)}</span>'
    bruto = _monto(c.monto) + sufijo
    if c.comision_pct is not None or c.comision_monto:
        com = _pct_banco(c.comision_pct) if c.comision_pct is not None else ""
        if c.comision_monto:
            com = f"{com} = {_money(c.comision_monto)}".strip(" =")
    else:
        com = "—"
    neto = (_monto(c.neto) + sufijo) if c.neto is not None else "—"
    if c.conciliado is None:
        conc = "—"
    else:
        conc = '<span class="verde">Sí</span>' if c.conciliado else "No"
    usd = ""
    if con_usd:
        # cobrosEnUsd: un cobro USD vale su bruto; uno MXN sin ningún TC queda
        # FUERA de la suma y se marca (nunca desaparece en silencio).
        monto_usd = c.monto_usd
        if monto_usd is None and moneda == "USD":
            monto_usd = c.monto
        usd_txt = (
            _monto(monto_usd) if monto_usd is not None else '<span class="rojo">sin T.C.</span>'
        )
        if c.tc and moneda != "USD":
            usd_txt += f' <span class="muted">T.C. {_tc_txt(c.tc)}</span>'
        usd = f'<td class="num">{usd_txt}</td>'
    return (
        f"<tr><td>{_fecha_corta(c.fecha, con_anio=True)}</td><td>{metodo}</td>"
        f'<td class="num">{bruto}</td><td class="num">{com}</td><td class="num">{neto}</td>'
        f"{usd}<td>{conc}</td></tr>"
    )


def _cobros_html(r: CotizacionInternaPdfRequest) -> str:
    """Cobros del vuelo (cobro_vuelo, partes de sobre incluidas), compactos:
    fecha, método (referencia en gris), bruto, comisión bancaria, neto,
    conciliado; resumen cobrado (cobrosEnUsd), comisiones, neto, saldo y semáforo.

    El `<h2>` lleva el método PREVISTO (22-sep-2026, ver `_metodo_previsto_txt`)
    y la fila vacía también: en el 19 % de las cotizaciones no hay ningún
    cobro registrado y el método sigue siendo el que explica el IVA.

    El bloque NO lleva `.bloque` (`page-break-inside: avoid`) a propósito:
    cada FILA ya lo lleva y el `thead` se repite, así que en el peor caso se
    va a la hoja 2 un par de filas — no el bloque entero con media hoja en
    blanco detrás."""
    con_usd = any((c.moneda or "USD").upper() != "USD" for c in r.cobros)
    ncols = 7 if con_usd else 6
    previsto = _metodo_previsto_txt(r)
    if r.cobros:
        filas = "".join(_cobro_fila(c, con_usd) for c in r.cobros)
    else:
        vacia = "Sin cobros registrados." + (f" · {previsto}" if previsto else "")
        filas = f'<tr><td colspan="{ncols}" class="muted">{escape(vacia)}</td></tr>'
    th_usd = '<th class="num">Equiv. USD</th>' if con_usd else ""

    resumen = [f"Cobrado {_monto(r.total_cobrado_usd)} USD"]
    if r.comision_banco_usd:
        resumen.append(f"comisiones banco &minus;{_money(r.comision_banco_usd)}")
        if r.total_cobrado_neto_usd is not None:
            resumen.append(f"neto {_monto(r.total_cobrado_neto_usd)}")
    # Saldo = total − cobrado, calculado por el API (saldo_usd); sin dato «—».
    if r.saldo_usd is None:
        resumen.append("Saldo —")
    else:
        saldo_txt = f"Saldo {_monto(r.saldo_usd)}"
        if r.saldo_usd < -0.005:
            saldo_txt += " (sobrecobro)"
        resumen.append(saldo_txt)
    sem_cls = _SEMAFORO_CLASES.get((r.semaforo_cobro or "").lower(), "sem-gris")
    sem_lbl = r.semaforo_cobro_label or (r.semaforo_cobro_key or "—").replace("_", " ").capitalize()
    resumen.append(f'<span class="sem {sem_cls}"></span>{escape(sem_lbl)}')
    pie = f'<tr><td colspan="{ncols}">{" · ".join(resumen)}</td></tr>'
    if r.cobros_sin_tc_count:
        pie += (
            f'<tr><td colspan="{ncols}" class="rojo aviso">'
            f"OJO: {_plural(r.cobros_sin_tc_count, 'cobro', 'cobros')} en MXN por "
            f"${r.cobros_sin_tc_mxn:,.2f} SIN tipo de cambio: fuera de la suma.</td></tr>"
        )
    return f"""
  <div><h2>Cobros{_op(previsto)}</h2>
  <table class="grid"><thead><tr>
    <th>Fecha</th><th>Método</th><th class="num">Bruto</th>
    <th class="num">Comisión banco</th><th class="num">Neto</th>{th_usd}<th>Conc.</th>
  </tr></thead><tbody>{filas}</tbody><tfoot>{pie}</tfoot></table></div>"""


def _notas_html(r: CotizacionInternaPdfRequest) -> str:
    """Solo las notas INTERNAS, recortadas. Las del cliente no son de este documento."""
    internas = _truncar(r.notas_internas)
    if not internas:
        return ""
    return (
        '<div class="bloque"><h2>Notas internas</h2>'
        f'<div class="notas-txt">{escape(internas)}</div></div>'
    )


def _estilos_page_interno(pie: str) -> str:
    """Reglas de PAPEL del documento interno: `@page` (Carta, márgenes de
    9/12/10 mm, pie «Documento interno · generado … · usuario» a la izquierda
    y el paginado a la derecha) y el `margin: 0` del <body>. SOLO tienen
    sentido en WeasyPrint — el panel no las recibe (no hay página que
    maquetar) y por eso NO viven en `cotizacion-interna.css`, igual que
    `_estilos_page()` en la hoja del cliente. La familia se repite aquí
    porque WeasyPrint no hereda la tipografía dentro de los márgenes de
    página."""
    fuente = "Arimo, 'Helvetica Neue', Arial, sans-serif"
    return f"""  @page {{
    size: Letter;
    margin: 9mm 12mm 10mm;
    @bottom-left {{ content: "{_css_str(pie)}"; font-size: 7.5pt; color: #9ca3af;
                    font-family: {fuente}; }}
    @bottom-right {{ content: "Página " counter(page) " de " counter(pages);
                     font-size: 7.5pt; color: #9ca3af; font-family: {fuente}; }}
  }}
  body {{ margin: 0; }}
"""


@lru_cache(maxsize=1)
def _estilos_cuerpo_interno() -> str:
    """CSS del CUERPO del documento interno (cabecera navy, banda roja,
    resumen, ficha, tablas de tramos y cobros, desglose/horas, cobros y
    notas): lo comparten el PDF, la vista previa y la HOJA INTERNA del panel
    (Fase 2.2). Desde el 22-sep-2026 vive en `app/static/cotizacion-interna.css`
    (fuente única; aquí solo se lee) con TODO selector acotado a
    `.cot-interna` (`CLASE_RAIZ`) — mismo patrón que `cotizacion-hoja.css`
    con `.cot-hoja`.

    Aire (11-sep-2026, captura de la oficina: «todo muy junto»): la v2 salió
    a 8.5 pt con interlineado 1.2 y celdas de 1 px — ilegible de un vistazo.
    Reglas de esta hoja, en orden de prioridad:
      1. NINGÚN texto de tabla baja de 9.5 pt (las aclaraciones en gris y los
         encabezados van a 8.5 pt: son apoyo, no el dato).
      2. Interlineado 1.3 y celdas de 3 px verticales / 6 px horizontales
         (antes 1 px × 4 px): las filas respiran.
      3. Cada bloque se separa del anterior con el margen SUPERIOR de su
         `h2` (9 px, antes 5) — por eso «Desglose de la cotización» y «Horas
         cotizadas», que van lado a lado, nunca se tocan: además del margen
         llevan 16 px de canal y una línea divisoria.
    La MISMA información sigue cabiendo en UNA hoja (ver el presupuesto del
    encabezado del módulo): lo que se encogió es el número de tramos/cobros
    que caben, no el contenido."""
    return (_STATIC / "cotizacion-interna.css").read_text(encoding="utf-8")


def _estilos_hoja_interna() -> str:
    """Fuente incrustada + cuerpo: EXACTAMENTE lo que el panel recibe en
    `GET /reportes/cotizacion-interna/hoja.css` y lo que lleva el PDF
    (contiguo, para que un test lo verifique como substring).

    FUENTE INCRUSTADA (22-sep-2026): Arimo (OFL, métrica de Arial) vía
    `_estilos_fuente()` del PDF del cliente — el MISMO archivo, importado y
    no copiado. Antes esta hoja declaraba 'Helvetica Neue', Arial sin
    `@font-face` y el contenedor de Railway (sin paquetes de fuentes) caía a
    su sans por defecto, ~8-12 % más ancho: lo que cabía en una hoja aquí se
    desbordaba allá — y el test de «una hoja» medía con métricas que
    producción no usa."""
    return _estilos_fuente() + _estilos_cuerpo_interno()


def _estilos_interno(pie: str) -> str:
    """CSS COMPLETO del PDF interno: reglas de papel (`@page` + el margen del
    <body>) + lo que recibe el panel (fuente incrustada + cuerpo, CONTIGUO,
    para que un test lo verifique como substring). Mismo reparto que
    `_estilos_base()` en la hoja del cliente. El branding (#dc2626 / #102a43)
    ya viaja resuelto dentro del archivo estático, así que este módulo dejó
    de necesitar `_BRAND`/`_NAVY`."""
    return _estilos_page_interno(pie) + _estilos_hoja_interna()


def _cuerpo_interno_html(r: CotizacionInternaPdfRequest) -> str:
    """EL DOCUMENTO interno, sin `<head>` ni CSS: la raíz
    `<div class="cot-interna">` con cabecera + banda, resumen, ficha, tramos,
    desglose|horas, cobros y notas internas.

    Fuente ÚNICA del marcado (22-sep-2026): lo usan el PDF (`_build_html`) y
    la vista previa del panel (`render_cotizacion_interna_preview_html`) —
    jamás una réplica. El <div> existe para que TODO selector del CSS pueda
    colgar de `.cot-interna` y la hoja interna del panel se pinte con el
    MISMO marcado y las MISMAS clases que el papel."""
    cuerpo = (
        '<table class="cols desglose bloque"><tr>'
        f'<td class="col">{_desglose_html(r)}</td>'
        f'<td class="col">{_horas_html(r)}</td>'
        "</tr></table>"
    )
    return f"""<div class="{CLASE_RAIZ}">
{_header_html(r)}
  {_resumen_html(r)}
  {_ficha_html(r)}
  {_tramos_html(r)}
  {cuerpo}
  {_cobros_html(r)}
  {_notas_html(r)}
</div>"""


def _build_html(r: CotizacionInternaPdfRequest) -> str:
    return f"""<!doctype html>
<html><head><meta charset="utf-8"><title>Cotización interna #{escape(r.folio)}</title><style>
{_estilos_interno(_pie_texto(r))}
</style></head><body>
{_cuerpo_interno_html(r)}
</body></html>"""


def render_cotizacion_interna_pdf(req: CotizacionInternaPdfRequest) -> bytes:
    from weasyprint import HTML  # import perezoso

    return HTML(string=_build_html(req)).write_pdf()


def render_cotizacion_interna_preview_html(req: CotizacionInternaPdfRequest) -> str:
    """Vista previa del documento interno para el panel (2.1, 22-sep-2026):
    el MISMO `_cuerpo_interno_html` del PDF con el MISMO payload — solo el
    cuerpo (`<div class="cot-interna">…</div>`), sin `<head>`, sin `<style>`
    y sin `@page`. El CSS viaja aparte por
    `GET /reportes/cotizacion-interna/hoja.css`, así que la hoja del panel y
    el papel se pintan con el mismo marcado y la misma hoja de estilos. NO
    importa WeasyPrint (es HTML, no PDF)."""
    return _cuerpo_interno_html(req)
