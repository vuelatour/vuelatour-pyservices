"""Lectura DETERMINISTA del PDF de una factura EMITIDA (24-sep-2026).

Pedido de Ale: «que haya una [sección] de facturas emitidas para las que hace
Mari manualmente … Mari las estaría adjuntando en PDF». Mari suelta el PDF (la
representación impresa del CFDI que le dio su PAC) en el diálogo «Registrar
factura» del panel; el API (`facturas-emitidas/leer-archivo`) lo manda aquí y
prellena serie, folio, UUID, fecha, RFC, montos, moneda y método. Ella revisa
y guarda. AQUÍ NO SE GUARDA NADA ni se decide nada de negocio: duplicados,
razón social emisora y cliente sugerido viven en el API.

SIN IA: pypdf saca el texto y expresiones regulares TOLERANTES buscan los
campos. Cada PAC arma su PDF distinto, así que:
  · Ningún campo es obligatorio: lo que no se encuentra sale `None` y el API
    arma el aviso «No encontré: …». Mejor vacío que inventado.
  · Nunca 500. Solo lo que NO es un PDF utilizable es 422 (`PdfNoValidoError`:
    base64 roto, sin `%PDF-`, más de 11 MB). Un PDF protegido, escaneado,
    roto o que tarda demasiado ⇒ 200 con `texto_extraido=False` + aviso.
  · pypdf corre en un hilo con TOPE DE TIEMPO (`TOPE_SEGUNDOS`): un PDF
    malformado puede dejarlo en un ciclo largo y el API corta a los 30 s.

Trampas del ejemplo real (representación impresa de Seguros Inbursa, 4.0):
  · Etiquetas en una línea y valores en la siguiente:
        Emisor: RFC emisor: Régimen fiscal emisor:
        SEGUROS INBURSA, S.A., GRUPO FINANCIERO INBURSA SIN9408027L7 601
  · Más abajo hay OTRA línea «Emisor: 26300 Póliza: …» sin RFC: una línea
    sin RFC válido jamás pisa al emisor ya encontrado.
  · «RFC proveedor de certificación: AUR100128NN3» y la cadena original
    («||1.1|uuid|fecha|AUR100128NN3|…») traen el RFC del PAC: se excluye.
  · «Subtotal:» e «Importe total con letra» NO son el total; «Prima total:» sí.
  · «No. de serie del certificado» no es la serie; «Folio fiscal» no es el
    folio.
  · Entre líneas hay basura binaria (el QR): se limpia antes de buscar.
"""

from __future__ import annotations

import base64
import binascii
import logging
import re
import threading
import time
import unicodedata
from collections.abc import Callable
from dataclasses import dataclass, field
from datetime import date
from io import BytesIO
from typing import Any

from app.schemas.facturacion import LeerPdfEmitidaResponse

try:  # sin pypdf el router igual arranca (y la lectura degrada a captura manual)
    from pypdf import PdfReader
except ImportError:  # pragma: no cover - depende del entorno
    PdfReader = None  # type: ignore[assignment,misc]

logger = logging.getLogger("facturacion.emitida_pdf")

# El API topa la subida en 10 MB; 11 MB deja holgura al redondeo del base64.
MAX_BYTES = 11 * 1024 * 1024
MAX_PAGINAS = 5
TOPE_SEGUNDOS = 20.0
# Menos que esto (letras + dígitos) = no hay capa de texto (PDF escaneado).
MIN_ALFANUMERICOS = 30
# Tolerancia del cuadre subtotal + IVA = total (centavos de redondeo del PAC).
TOL_CUADRE = 0.02

AVISO_PROTEGIDO = "El PDF está protegido con contraseña: captura los datos a mano."
AVISO_TIEMPO = "No pude leer el PDF a tiempo: captura los datos a mano."
AVISO_ILEGIBLE = "No pude leer el texto del PDF (parece dañado): captura los datos a mano."
AVISO_SIN_TEXTO = "El PDF no tiene texto (parece escaneado)."
AVISO_PAGINAS = f"Solo se leyeron las primeras {MAX_PAGINAS} páginas."
AVISO_VARIOS_UUID = "Encontré varios folios fiscales; captura el UUID a mano."
AVISO_CAMPOS = "No pude interpretar el texto del PDF: captura los datos a mano."


class PdfNoValidoError(ValueError):
    """El archivo NO es un PDF utilizable ⇒ 422 (el API lo traduce a aviso)."""


# ---------------------------------------------------------------------------
# Entrada
# ---------------------------------------------------------------------------

_PREFIJO_DATA_URL = re.compile(r"^data:[^;,]*;base64,", re.IGNORECASE)


def decodificar_pdf(pdf_b64: str) -> bytes:
    """Bytes del PDF o `PdfNoValidoError` (base64 roto, demasiado grande, no-PDF)."""
    s = _PREFIJO_DATA_URL.sub("", (pdf_b64 or "").strip())
    s = re.sub(r"\s+", "", s)
    if not s:
        raise PdfNoValidoError("No llegó ningún PDF")
    # 4·⌈n/3⌉ es lo más largo que mide el base64 de n bytes: más largo que eso
    # ya pesa más del tope y ni se decodifica.
    if len(s) > 4 * ((MAX_BYTES + 2) // 3):
        raise PdfNoValidoError("El PDF pesa más de 10 MB")
    try:
        raw = base64.b64decode(s, validate=True)
    except (binascii.Error, ValueError) as e:
        raise PdfNoValidoError("El PDF no llegó en base64 válido") from e
    if len(raw) > MAX_BYTES:
        raise PdfNoValidoError("El PDF pesa más de 10 MB")
    if b"%PDF-" not in raw[:1024]:
        raise PdfNoValidoError("El archivo no es un PDF")
    return raw


# ---------------------------------------------------------------------------
# Texto (pypdf dentro del tope de tiempo)
# ---------------------------------------------------------------------------


@dataclass
class _Extraccion:
    texto: str | None
    paginas: int = 0
    avisos: list[str] = field(default_factory=list)


def _extraer_texto(raw: bytes) -> _Extraccion:
    """Texto de las primeras `MAX_PAGINAS` páginas. Corre DENTRO del tope."""
    if PdfReader is None:
        logger.error("pypdf no está instalado: no se puede leer el PDF de la factura")
        return _Extraccion(None, 0, [AVISO_ILEGIBLE])
    lector = PdfReader(BytesIO(raw), strict=False)
    if lector.is_encrypted:
        # Muchos PAC cifran con contraseña de USUARIO vacía (solo restringen
        # imprimir/copiar): esos se abren con "". Con contraseña real, no.
        try:
            abierto = bool(lector.decrypt(""))
        except Exception:  # noqa: BLE001 — cifrado raro o falta la librería
            abierto = False
        if not abierto:
            return _Extraccion(None, 0, [AVISO_PROTEGIDO])
    paginas = len(lector.pages)
    avisos = [AVISO_PAGINAS] if paginas > MAX_PAGINAS else []
    partes: list[str] = []
    fallidas = 0
    for i in range(min(paginas, MAX_PAGINAS)):
        try:
            partes.append(lector.pages[i].extract_text() or "")
        except Exception:  # noqa: BLE001 — una página rota no tumba las demás
            fallidas += 1
    texto = "\n".join(partes)
    if fallidas and not texto.strip():
        return _Extraccion(None, paginas, [*avisos, AVISO_ILEGIBLE])
    return _Extraccion(texto, paginas, avisos)


def _con_tope(fn: Callable[[Any], Any], arg: Any, segundos: float) -> tuple[str, Any]:
    """Corre `fn(arg)` con tope de `segundos`: ('ok', resultado) ·
    ('error', excepción) · ('tiempo', None). Jamás propaga excepciones.

    Hilo DAEMON (no un ThreadPoolExecutor): un pypdf atorado no se puede
    matar, pero así tampoco detiene el apagado del proceso en un redeploy.
    """
    resultado: list[tuple[str, Any]] = []

    def trabajo() -> None:
        try:
            resultado.append(("ok", fn(arg)))
        except BaseException as e:  # noqa: BLE001 — se reporta como aviso
            resultado.append(("error", e))

    hilo = threading.Thread(target=trabajo, name="leer-pdf-emitida", daemon=True)
    hilo.start()
    hilo.join(segundos)
    if hilo.is_alive() or not resultado:
        return "tiempo", None
    return resultado[0]


def _extraer_con_tope(raw: bytes) -> _Extraccion:
    """`_extraer_texto` con tope de `TOPE_SEGUNDOS`; jamás propaga excepciones."""
    estado, r = _con_tope(_extraer_texto, raw, TOPE_SEGUNDOS)
    if estado == "tiempo":
        logger.warning("Lectura del PDF de factura emitida superó %.0f s", TOPE_SEGUNDOS)
        return _Extraccion(None, 0, [AVISO_TIEMPO])
    if estado == "error":
        logger.warning("pypdf no pudo leer el PDF de factura emitida: %r", r)
        return _Extraccion(None, 0, [AVISO_ILEGIBLE])
    return r


def normalizar_texto_pdf(texto: str) -> str:
    """NFKC, sin caracteres de control (la basura binaria del QR), espacios
    colapsados y sin líneas vacías. Idempotente."""
    t = unicodedata.normalize("NFKC", texto or "")
    t = t.replace("\r\n", "\n").replace("\r", "\n")
    t = "".join(
        ch if ch == "\n" or not unicodedata.category(ch).startswith("C") else " " for ch in t
    )
    lineas = (" ".join(linea.split()) for linea in t.split("\n"))
    return "\n".join(linea for linea in lineas if linea)


# ---------------------------------------------------------------------------
# Punto de entrada
# ---------------------------------------------------------------------------


def leer_pdf_emitida(pdf_b64: str) -> LeerPdfEmitidaResponse:
    """Campos de la representación impresa de un CFDI. `PdfNoValidoError` ⇒ 422."""
    raw = decodificar_pdf(pdf_b64)
    inicio = time.monotonic()
    ext = _extraer_con_tope(raw)
    avisos = list(ext.avisos)
    if ext.texto is None:
        return LeerPdfEmitidaResponse(texto_extraido=False, paginas=ext.paginas, avisos=avisos)
    texto = normalizar_texto_pdf(ext.texto)
    if sum(ch.isalnum() for ch in texto) < MIN_ALFANUMERICOS:
        avisos.append(AVISO_SIN_TEXTO)
        return LeerPdfEmitidaResponse(texto_extraido=False, paginas=ext.paginas, avisos=avisos)
    # La interpretación comparte el MISMO tope (lo que sobró de la lectura,
    # mínimo 1 s): con un texto patológico (miles de etiquetas apiladas, que
    # es cuadrático) el endpoint igual responde antes de que el API corte a
    # los 30 s; el hilo atorado termina solo, como el de pypdf.
    restante = max(TOPE_SEGUNDOS - (time.monotonic() - inicio), 1.0)
    estado, r = _con_tope(extraer_campos, texto, restante)
    if estado != "ok":
        if estado == "tiempo":
            logger.warning("Interpretar el PDF de factura emitida superó el tope")
            avisos.append(AVISO_TIEMPO)
        else:  # un texto raro no se vuelve 500
            logger.error("Error interpretando el texto de un PDF de factura emitida", exc_info=r)
            avisos.append(AVISO_CAMPOS)
        return LeerPdfEmitidaResponse(texto_extraido=True, paginas=ext.paginas, avisos=avisos)
    campos, avisos_campos = r
    return LeerPdfEmitidaResponse(
        **campos, texto_extraido=True, paginas=ext.paginas, avisos=[*avisos, *avisos_campos]
    )


# ---------------------------------------------------------------------------
# Campos (funciones PURAS sobre el texto)
# ---------------------------------------------------------------------------

CAMPOS = (
    "serie",
    "folio",
    "uuid",
    "fecha_emision",
    "emisor_rfc",
    "emisor_nombre",
    "receptor_rfc",
    "receptor_nombre",
    "subtotal",
    "iva",
    "total",
    "moneda",
    "metodo_pago",
    "forma_pago",
)


def extraer_campos(texto: str) -> tuple[dict, list[str]]:
    """(campos, avisos) de un texto de PDF. Lo no hallado sale `None`."""
    texto = normalizar_texto_pdf(texto)
    lineas = texto.split("\n")
    avisos: list[str] = []
    c: dict = dict.fromkeys(CAMPOS)

    c["uuid"], aviso = _extraer_uuid(texto)
    _agregar(avisos, aviso)
    c["serie"], c["folio"] = _extraer_serie_folio(lineas)
    c["fecha_emision"] = _extraer_fecha(lineas)
    partes = _extraer_partes(lineas, texto)
    c.update(partes)

    c["subtotal"] = _monto_de(lineas, _ETQS_SUBTOTAL, _PREFIJO_MONTO)
    c["iva"] = _monto_de(lineas, _ETQS_IVA, _PREFIJO_IVA, excluir_antes=_RETENCION_ANTES)
    ajustes = _ajustes_cuadre(lineas)
    c["total"] = _extraer_total(lineas, c["subtotal"], c["iva"], ajustes)
    _agregar(avisos, _aviso_cuadre(c["subtotal"], c["iva"], c["total"], ajustes))

    c["moneda"], aviso = _extraer_moneda(lineas)
    _agregar(avisos, aviso)
    c["metodo_pago"] = _extraer_metodo(lineas)
    c["forma_pago"] = _extraer_forma(lineas)
    _agregar(avisos, _aviso_tipo_comprobante(lineas))
    return c, avisos


def _agregar(avisos: list[str], aviso: str | None) -> None:
    if aviso and aviso not in avisos:
        avisos.append(aviso)


def _valor_tras(lineas: list[str], i: int, fin: int) -> str:
    """Lo que sigue a una etiqueta: el resto de la línea; si quedó vacío y la
    etiqueta cerró con «:», la línea siguiente."""
    resto = lineas[i][fin:].strip()
    if resto:
        return resto
    if lineas[i][:fin].rstrip().endswith(":") and i + 1 < len(lineas):
        return lineas[i + 1]
    return ""


# Corta un valor libre (un nombre) donde empieza la siguiente etiqueta.
_SIGUIENTE_ETIQUETA = re.compile(
    r"\s(?:R\.?\s?F\.?\s?C\.?(?![A-Za-z])|R[eé]gimen|Uso\s+(?:de\s+)?CFDI|C[oó]digo\s+postal"
    r"|C\.P\.|Domicilio|Folio|Serie|Fecha|Lugar|Tipo\s+de|Residencia|N[uú]m\.?\s+de)",
    re.IGNORECASE,
)


# --- UUID -------------------------------------------------------------------

_HEX = "[0-9A-Fa-f]"
# Patrón EXACTO 8-4-4-4-12 con guiones y sin hex pegado a los lados: los
# sellos (base64) y cadenas nunca lo cumplen. `-\s?` tolera el corte de
# renglón justo después de un guion.
_UUID_RE = re.compile(
    rf"(?<![0-9A-Fa-f-])({_HEX}{{8}}-\s?{_HEX}{{4}}-\s?{_HEX}{{4}}-\s?{_HEX}{{4}}-\s?{_HEX}{{12}})"
    r"(?![0-9A-Fa-f-])"
)
_ETQS_UUID = (
    re.compile(r"folio\s*fiscal", re.IGNORECASE),  # también «Folio Fiscal (UUID)»
    re.compile(r"(?<![A-Za-z])UUID", re.IGNORECASE),
)
# «UUID relacionado», «CFDI relacionado», «folio fiscal sustituido»: es el de
# OTRA factura (re-emisión / nota de crédito), nunca el de esta.
_CONTEXTO_RELACIONADO = re.compile(r"relacionad|sustitu|tipo\s+de\s+relaci", re.IGNORECASE)


def _norm_uuid(s: str) -> str:
    return re.sub(r"\s", "", s).upper()


def _ventana(texto: str, ini: int) -> str:
    """Desde `ini` hasta el final de la línea SIGUIENTE."""
    fin = texto.find("\n", ini)
    if fin == -1:
        return texto[ini:]
    fin2 = texto.find("\n", fin + 1)
    return texto[ini:] if fin2 == -1 else texto[ini:fin2]


def _contexto_etiqueta(texto: str, ini: int, fin: int, *, con_titulo: bool) -> str:
    """La etiqueta con lo que la rodea en su línea. Con `con_titulo`, si la
    etiqueta ABRE la línea y la línea anterior es un TÍTULO (sin UUID), esa
    también: el bloque «CFDI relacionados» pone su título arriba y «UUID: …»
    abajo."""
    inicio_linea = texto.rfind("\n", 0, ini) + 1
    contexto = texto[max(inicio_linea, ini - 60) : fin + 25]
    if con_titulo and inicio_linea and not texto[inicio_linea:ini].strip():
        previa = texto[texto.rfind("\n", 0, inicio_linea - 1) + 1 : inicio_linea]
        if not _UUID_RE.search(previa):
            contexto = previa + contexto
    return contexto


def _extraer_uuid(texto: str) -> tuple[str | None, str | None]:
    relacionados: set[str] = set()
    for n, etiqueta in enumerate(_ETQS_UUID):
        for m in etiqueta.finditer(texto):
            # «Folio fiscal» es la leyenda del SAT para el UUID PROPIO: solo su
            # misma línea decide; el «UUID:» suelto también mira el título.
            contexto = _contexto_etiqueta(texto, m.start(), m.end(), con_titulo=n > 0)
            if _CONTEXTO_RELACIONADO.search(contexto):
                fin_linea = texto.find("\n", m.end())
                r = _UUID_RE.search(texto[m.end() : fin_linea if fin_linea != -1 else None])
                if r:
                    relacionados.add(_norm_uuid(r.group(1)))
                continue
            v = _UUID_RE.search(_ventana(texto, m.end()))
            if v:
                return _norm_uuid(v.group(1)), None
    # Sin etiqueta propia: el único UUID distinto del texto (típicamente el de
    # la cadena original del timbre), sin contar los de CFDI relacionados.
    distintos: list[str] = []
    for v in _UUID_RE.finditer(texto):
        u = _norm_uuid(v.group(1))
        if u not in distintos and u not in relacionados:
            distintos.append(u)
    if len(distintos) == 1:
        return distintos[0], None
    if len(distintos) > 1:
        return None, AVISO_VARIOS_UUID
    return None, None


# --- Serie y folio ----------------------------------------------------------

_ETQ_SERIE_Y_FOLIO = re.compile(
    r"(?<![A-Za-z])Serie\s*(?:y|e|/|-|&)\s*Folio(?![A-Za-z])\s*:?\s*", re.IGNORECASE
)
# «A-123» · «A 123» · «A123» · «123» (la serie, si viene, son letras).
_VAL_SERIE_FOLIO = re.compile(
    r"(?:(?P<serie>[A-Za-z]{1,10})\s?[-/]?\s?)?(?P<folio>\d[\d-]{0,39})(?!\d)"
)
# «Serie:» con dos puntos: «No. de serie del certificado …» NO lo cumple.
_ETQ_SERIE = re.compile(r"(?<![A-Za-z])Serie\s*:\s*", re.IGNORECASE)
_VAL_SERIE = re.compile(r"(?P<serie>[A-Za-z0-9][A-Za-z0-9-]{0,24})(?![A-Za-z0-9-])(?!\s?:)")
# «Folio» NUNCA seguido de «fiscal» (ese es el UUID).
_ETQ_FOLIO = re.compile(
    r"(?<![A-Za-z])Folio(?![A-Za-z])"
    r"(?!\s*(?:fiscal|del\s+SAT|SAT\b|de\s+sustituci|relacionad))"
    r"\s*(?:No\.?|N[uú]m(?:ero)?\.?|#)?\s*:?\s*",
    re.IGNORECASE,
)
# «No. de Serie: 0000100000…» es el número del certificado, no la serie.
_NO_DE_ANTES = re.compile(r"(?:No\.?|N[uú]m(?:ero)?\.?|N°)\s*(?:de\s*)?$", re.IGNORECASE)


def _serie_ok(s: str | None) -> str | None:
    s = (s or "").strip(" -").upper()
    return s[:25] or None


def _folio_ok(f: str | None) -> str | None:
    f = (f or "").strip(" -")
    return f[:40] or None


def _serie_folio_de(valor: str) -> re.Match[str] | None:
    """`_VAL_SERIE_FOLIO` con un candado: una «serie» separada del número SOLO
    por un espacio tiene que venir en MAYÚSCULAS («A 123» sí; «interno 555»
    de «Folio interno 555», o «Fecha 2026…» de la línea siguiente, no)."""
    v = _VAL_SERIE_FOLIO.match(valor)
    if v is None or v.group("serie") is None:
        return v
    solo_espacio = not valor[v.end("serie") : v.start("folio")].strip()
    return v if (not solo_espacio or v.group("serie").isupper()) else None


# Tras la serie, en la línea siguiente, solo puede seguir OTRA etiqueta
# («A Folio: 123»); un número suelto («Fecha 2026-09-24») delata que ese
# renglón no es el valor de «Serie:».
_TRAS_SERIE_OK = re.compile(r"\s*$|\s+[^\d:]{2,40}:")


def _extraer_serie_folio(lineas: list[str]) -> tuple[str | None, str | None]:
    for i, linea in enumerate(lineas):
        for m in _ETQ_SERIE_Y_FOLIO.finditer(linea):
            v = _serie_folio_de(_valor_tras(lineas, i, m.end()))
            if v:
                return _serie_ok(v.group("serie")), _folio_ok(v.group("folio"))

    serie: str | None = None
    for i, linea in enumerate(lineas):
        for m in _ETQ_SERIE.finditer(linea):
            if _NO_DE_ANTES.search(linea[: m.start()]):
                continue
            valor = _valor_tras(lineas, i, m.end())
            v = _VAL_SERIE.match(valor)
            if v and not linea[m.end() :].strip() and not _TRAS_SERIE_OK.match(valor, v.end()):
                continue  # vino de la línea siguiente y no es SOLO la serie
            # 15+ dígitos = número de certificado, no una serie.
            if v and not (v.group("serie").isdigit() and len(v.group("serie")) >= 15):
                serie = _serie_ok(v.group("serie"))
                break
        if serie:
            break

    folio: str | None = None
    prefijo: str | None = None
    # Primero las etiquetas con «:»/«#»/«No.» («Folio: 777»); una «Folio» sin
    # ellas («Folio interno 555») solo si no hubo otra.
    etiquetas_folio = sorted(
        ((i, m) for i, linea in enumerate(lineas) for m in _ETQ_FOLIO.finditer(linea)),
        key=lambda im: not re.search(r"[:#]|N[oOuUúÚ]", im[1].group(0)),
    )
    for i, m in etiquetas_folio:
        v = _serie_folio_de(_valor_tras(lineas, i, m.end()))
        if v:
            prefijo, folio = _serie_ok(v.group("serie")), _folio_ok(v.group("folio"))
            break

    if prefijo and folio:
        if serie is None:
            serie = prefijo  # «Folio: A-123» ⇒ serie A, folio 123
        elif prefijo != serie:
            folio = _folio_ok(f"{prefijo}-{folio}")  # no se descarta lo impreso
    return serie, folio


# --- Fecha de emisión -------------------------------------------------------

_ETQS_FECHA = (
    re.compile(r"fecha\s+(?:y\s+hora\s+)?(?:de\s+)?(?:expedici[oó]n|emisi[oó]n)", re.IGNORECASE),
    re.compile(r"(?<![A-Za-z])(?:fecha|expedida|emitida|expedido)\s*:", re.IGNORECASE),
)
# Año primero: «2026-09-24» y también «2026/09/24» (mismo separador).
_FECHA_ISO = re.compile(r"(?<!\d)(\d{4})([-/])(\d{1,2})\2(\d{1,2})(?!\d)")
_FECHA_DMA = re.compile(r"(?<!\d)(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})(?!\d)")
_FECHA_TEXTO = re.compile(
    r"(?<!\d)(\d{1,2})(?:\s+de\s+|[\s/.-]+)([A-Za-zé]{3,10})\.?(?:\s+del?\s+|[\s/.-]+)(\d{4})(?!\d)",
    re.IGNORECASE,
)
_MESES = {
    "ene": 1, "jan": 1, "feb": 2, "mar": 3, "abr": 4, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "ago": 8, "aug": 8, "sep": 9, "set": 9, "oct": 10, "nov": 11, "dic": 12,
    "dec": 12,
}  # fmt: skip


def _fecha_iso(a: int, m: int, d: int) -> str | None:
    if not 2000 <= a <= 2099:
        return None
    try:
        return date(a, m, d).isoformat()
    except ValueError:
        return None


def _primera_fecha(s: str) -> str | None:
    """La primera fecha válida de `s` como 'YYYY-MM-DD' (día de pared impreso,
    sin conversión de zona)."""
    candidatos: list[tuple[int, str]] = []
    for m in _FECHA_ISO.finditer(s):
        f = _fecha_iso(int(m.group(1)), int(m.group(3)), int(m.group(4)))
        if f:
            candidatos.append((m.start(), f))
    for m in _FECHA_DMA.finditer(s):
        d, mes, a = int(m.group(1)), int(m.group(2)), int(m.group(3))
        # México escribe día/mes; si no existe, se prueba mes/día.
        f = _fecha_iso(a, mes, d) or _fecha_iso(a, d, mes)
        if f:
            candidatos.append((m.start(), f))
    for m in _FECHA_TEXTO.finditer(s):
        mes = _MESES.get(m.group(2)[:3].lower())
        f = _fecha_iso(int(m.group(3)), mes, int(m.group(1))) if mes else None
        if f:
            candidatos.append((m.start(), f))
    return min(candidatos)[1] if candidatos else None


def _extraer_fecha(lineas: list[str]) -> str | None:
    for etiqueta in _ETQS_FECHA:
        for i, linea in enumerate(lineas):
            for m in etiqueta.finditer(linea):
                f = _primera_fecha(linea[m.end() :][:120])
                if f is None and i + 1 < len(lineas):
                    f = _primera_fecha(lineas[i + 1][:120])
                if f:
                    return f
    return None


# --- Emisor y receptor ------------------------------------------------------

_RFC_PATRON = r"[A-ZÑ&]{3,4}\d{6}[A-Z0-9]{3}"
_RFC_RE = re.compile(rf"(?<![A-ZÑ&0-9])({_RFC_PATRON})(?![A-Z0-9])")
_ETQ_RFC = r"R\.?\s?F\.?\s?C\.?"
_ROLES = {"emisor": r"emisor", "receptor": r"(?:receptor|cliente)"}
_ROL_CUALQUIERA = r"(?:emisor|receptor|cliente)"


def _etiqueta_rfc_de(rol: str) -> re.Pattern[str]:
    """«RFC emisor» · «RFC del receptor» · «Emisor RFC:» · «R.F.C. Cliente»."""
    return re.compile(
        rf"(?<![A-Za-z]){_ETQ_RFC}\s*(?:del\s+)?{rol}(?![A-Za-z])\s*:?"
        rf"|(?<![A-Za-z]){rol}\s*\(?{_ETQ_RFC}\)?\s*:",
        re.IGNORECASE,
    )


_ETQ_RFC_ROL = {k: _etiqueta_rfc_de(v) for k, v in _ROLES.items()}
_ETQ_RFC_CUALQUIER_ROL = _etiqueta_rfc_de(_ROL_CUALQUIERA)
_ETQ_NOMBRE_ROL = {
    k: re.compile(
        rf"(?<![A-Za-z])(?:nombre|raz[oó]n\s+social|denominaci[oó]n(?:\s+social)?)\s*"
        rf"(?:del\s+)?{v}(?![A-Za-z])\s*:\s*",
        re.IGNORECASE,
    )
    for k, v in _ROLES.items()
}
# «Emisor: …» / «Receptor: …» / «Cliente: …» al inicio de la línea.
_LINEA_ROL = re.compile(rf"^(?:datos\s+(?:del\s+)?)?({_ROL_CUALQUIERA})\s*:\s*(.*)$", re.IGNORECASE)
# Encabezado de bloque: la línea es SOLO «Emisor» / «Datos del receptor».
_ENCABEZADO_ROL = re.compile(
    rf"^(?:datos\s+(?:del\s+)?)?({_ROL_CUALQUIERA})(?:\s+del\s+(?:comprobante|cfdi))?\s*:?$",
    re.IGNORECASE,
)
_LINEA_RFC = re.compile(rf"^{_ETQ_RFC}(?![A-Za-z])\s*:?\s*", re.IGNORECASE)
_LINEA_NOMBRE = re.compile(
    r"^(?:nombre|raz[oó]n\s+social|denominaci[oó]n(?:\s+social)?)\s*:?\s*(.+)$", re.IGNORECASE
)
# En el encabezado, antes de «RFC emisor:», hay una columna de NOMBRE.
_NOMBRE_ANTES = re.compile(
    rf"(?<![A-Za-z])(?:{_ROL_CUALQUIERA}|nombre|raz[oó]n\s+social)(?![A-Za-z])", re.IGNORECASE
)
# RFC del PAC: «RFC proveedor de certificación: X» y la cadena original del TFD.
_RFC_PAC = re.compile(
    rf"{_ETQ_RFC}\s*(?:del\s+)?(?:proveedor|PAC)[^:\n]*:\s*({_RFC_PATRON})", re.IGNORECASE
)
_CADENA_TFD = re.compile(rf"\|\|\s*1\.[01]\s*\|[^|]*\|[^|]*\|\s*({_RFC_PATRON})\s*\|")
_LINEA_DEL_PAC = re.compile(r"proveedor|certificaci|(?<![A-Za-z])PAC(?![A-Za-z])", re.IGNORECASE)


def _rfc_valido(rfc: str) -> bool:
    """Las 6 cifras son AAMMDD: descarta tiras de letras+números al azar."""
    fecha = rfc[-9:-3]
    return 1 <= int(fecha[2:4]) <= 12 and 1 <= int(fecha[4:6]) <= 31


def _rfcs(s: str) -> list[re.Match[str]]:
    return [m for m in _RFC_RE.finditer(s) if _rfc_valido(m.group(1))]


def _rfcs_excluidos(texto: str) -> set[str]:
    return {m.group(1).upper() for m in _RFC_PAC.finditer(texto)} | {
        m.group(1) for m in _CADENA_TFD.finditer(texto)
    }


def _limpiar_nombre(s: str | None) -> str | None:
    s = " ".join((s or "").split()).strip(" ,;:-|/")
    s = re.sub(r"^(?:nombre|raz[oó]n\s+social)\s*:?\s*", "", s, flags=re.IGNORECASE)
    s = re.sub(rf"\s*{_ETQ_RFC}\s*:?$", "", s, flags=re.IGNORECASE).strip(" ,;:-|/")
    if len(re.findall(r"[^\W\d_]", s)) < 2:
        return None
    return s[:300]


def _cortar_en_etiqueta(s: str) -> str:
    m = _SIGUIENTE_ETIQUETA.search(" " + s)
    s = s[: m.start()] if m else s
    r = _RFC_RE.search(s)
    return s[: r.start()] if r else s


def _por_etiqueta_rfc(lineas: list[str], rol: str) -> tuple[str | None, str | None]:
    """«RFC emisor: X» en la misma línea, o el formato etiquetas-en-una-línea /
    valores-en-la-siguiente del ejemplo real (nombre = texto antes del RFC)."""
    etiqueta = _ETQ_RFC_ROL[rol]
    for i, linea in enumerate(lineas):
        for m in etiqueta.finditer(linea):
            en_linea = _rfcs(linea[m.end() :])
            if en_linea:
                return en_linea[0].group(1), None
            # Columna k del encabezado ⇒ k-ésimo RFC de la línea de valores.
            k = len(_ETQ_RFC_CUALQUIER_ROL.findall(linea[: m.start()]))
            for j in (i + 1, i + 2):
                if j >= len(lineas):
                    break
                valores = _rfcs(lineas[j])
                if not valores:
                    continue
                if k >= len(valores):
                    break
                v = valores[k]
                nombre = None
                if j == i + 1 and _NOMBRE_ANTES.search(linea[: m.start()]):
                    desde = valores[k - 1].end() if k else 0
                    nombre = _limpiar_nombre(lineas[j][desde : v.start()])
                return v.group(1), nombre
    return None, None


def _por_linea_rol(
    lineas: list[str], rol: str, excluidos: set[str]
) -> tuple[str | None, str | None]:
    """«Receptor: MAQAR MACHINERY RFC: MMA150622P83» o «Emisor: ACME SA» con
    «RFC: X» en la línea siguiente. Una línea «Emisor:» SIN RFC válido (la
    trampa «Emisor: 26300 Póliza: …») no aporta nada."""
    nombre_suelto: str | None = None
    for i, linea in enumerate(lineas):
        m = _LINEA_ROL.match(linea)
        if not m or not re.fullmatch(_ROLES[rol], m.group(1), re.IGNORECASE):
            continue
        resto = m.group(2)
        validos = [r for r in _rfcs(resto) if r.group(1) not in excluidos]
        if validos:
            r = validos[0]
            nombre = _limpiar_nombre(resto[: r.start()])
            despues = resto[r.end() :]
            if nombre is None and ":" not in despues:
                nombre = _limpiar_nombre(re.sub(r"\s\d{3}(?:\s.*)?$", "", despues))
            return r.group(1), nombre
        if not resto or ":" in resto or re.search(r"\d", resto):
            continue
        sig = lineas[i + 1] if i + 1 < len(lineas) else ""
        et = _LINEA_RFC.match(sig)
        rs = _rfcs(sig[et.end() :]) if et else []
        if rs and rs[0].start() == 0 and rs[0].group(1) not in excluidos:
            return rs[0].group(1), _limpiar_nombre(resto)
        nombre_suelto = nombre_suelto or _limpiar_nombre(resto)
    return None, nombre_suelto


def _por_bloque(lineas: list[str], rol: str, excluidos: set[str]) -> tuple[str | None, str | None]:
    """Bloque con encabezado propio: «Receptor» / «Nombre: X» / «RFC: Y»."""
    for i, linea in enumerate(lineas):
        m = _ENCABEZADO_ROL.match(linea)
        if not m or not re.fullmatch(_ROLES[rol], m.group(1), re.IGNORECASE):
            continue
        rfc: str | None = None
        nombre: str | None = None
        for j in range(i + 1, min(i + 7, len(lineas))):
            actual = lineas[j]
            if _ENCABEZADO_ROL.match(actual) or _LINEA_ROL.match(actual):
                break
            et = _LINEA_RFC.match(actual)
            if et:
                rs = [r for r in _rfcs(actual[et.end() :]) if r.group(1) not in excluidos]
                if rfc is None and rs:
                    rfc = rs[0].group(1)
                continue
            mn = _LINEA_NOMBRE.match(actual)
            if mn and nombre is None:
                nombre = _limpiar_nombre(_cortar_en_etiqueta(mn.group(1)))
            elif nombre is None and j == i + 1 and ":" not in actual and not _rfcs(actual):
                nombre = _limpiar_nombre(actual)
        if rfc:
            return rfc, nombre
    return None, None


def _por_etiqueta_nombre(lineas: list[str], rol: str) -> str | None:
    """«Nombre del receptor: X» / «Razón social emisor: X» en la misma línea."""
    for linea in lineas:
        for m in _ETQ_NOMBRE_ROL[rol].finditer(linea):
            nombre = _limpiar_nombre(_cortar_en_etiqueta(linea[m.end() :]))
            if nombre:
                return nombre
    return None


def _rfc_generico(lineas: list[str], excluidos: set[str]) -> str | None:
    """Último recurso para el EMISOR: un «RFC: X» suelto (el encabezado de la
    empresa en muchos PAC). Solo se usa si el receptor ya se identificó y el
    candidato es distinto a él y al del PAC."""
    for linea in lineas:
        if _LINEA_DEL_PAC.search(linea):
            continue
        for m in re.finditer(rf"(?<![A-Za-z]){_ETQ_RFC}(?![A-Za-z])\s*:?\s*", linea, re.IGNORECASE):
            rs = _rfcs(linea[m.end() :])
            if rs and rs[0].start() == 0 and rs[0].group(1) not in excluidos:
                return rs[0].group(1)
    return None


def _extraer_partes(lineas: list[str], texto: str) -> dict[str, str | None]:
    excluidos = _rfcs_excluidos(texto)
    res: dict[str, str | None] = {}
    for rol in _ROLES:
        rfc1, nombre1 = _por_etiqueta_rfc(lineas, rol)
        rfc2, nombre2 = _por_linea_rol(lineas, rol, excluidos)
        rfc3, nombre3 = (None, None) if (rfc1 or rfc2) else _por_bloque(lineas, rol, excluidos)
        res[f"{rol}_rfc"] = rfc1 or rfc2 or rfc3
        res[f"{rol}_nombre"] = _por_etiqueta_nombre(lineas, rol) or nombre1 or nombre2 or nombre3
    if res["emisor_rfc"] is None and res["receptor_rfc"]:
        res["emisor_rfc"] = _rfc_generico(lineas, excluidos | {res["receptor_rfc"]})
    return res


# --- Montos -----------------------------------------------------------------

# «6,336.80» · «$1,160.00» · «1160.00» · «$1,160» — con 2 decimales, o con
# «$» si es entero. «0.160000» (tasa) y «16.00%» NO son montos.
_MONTO_RE = re.compile(
    r"\$?\s?(?P<n>\d{1,3}(?:,\d{3})+(?:\.\d{1,2})?|\d+\.\d{2})(?![\d%]|\s?%|[.,]\d)"
    r"|\$\s?(?P<e>\d+)(?![\d.,%])"
)
# Moneda pegada a la etiqueta o al monto: «Total MXN:», «(USD)», «M.N.» y el
# «US$ 1,160.00» de algunos PAC (el «$» lo consume `_MONTO_RE`).
_MONEDA_PREFIJO = r"\(?\s?(?:MXN|USD|US(?=\s?\$)|M\.\s?N\.?)\s?\)?(?![A-Za-z])"
_PREFIJO_MONTO = re.compile(rf"(?:[\s:=]|{_MONEDA_PREFIJO})*", re.IGNORECASE)
# El IVA además salta «16%», «(16 %)», «Tasa 16%» y «Trasladado».
# «Trasladado» se salta solo si lo sigue la tasa, «:» o «$»: en el renglón de
# concepto «IVA Traslado 1,000.00 Tasa 16.00% 160.00» el primer número es la
# BASE, no el impuesto.
_PREFIJO_IVA = re.compile(
    rf"(?:[\s:=]|{_MONEDA_PREFIJO}"
    r"|traslad[a-z]*(?=\s?(?:\(|tasa|\d{1,2}(?:\.\d+)?\s?%|:|\$))"
    r"|\(?\s?(?:tasa\s?)?\d{1,2}(?:\.\d+)?\s?%\s?\)?)*",
    re.IGNORECASE,
)
# Una retención además salta el impuesto retenido: «Retención ISR: 100.00»,
# «Retención de IVA (10.6667 %): 106.67».
_PREFIJO_RETENCION = re.compile(
    rf"(?:[\s:=]|{_MONEDA_PREFIJO}"
    r"|(?:de(?:l)?\s)?(?:I\.?\s?V\.?\s?A\.?|I\.?\s?S\.?\s?R\.?)(?![A-Za-z])"
    r"|\(?\s?(?:tasa\s?)?\d{1,2}(?:\.\d+)?\s?%\s?\)?)*",
    re.IGNORECASE,
)
# «Retención IVA: 106.67» / «Ret. IVA» / «Impuestos retenidos IVA»: ese «IVA»
# es lo que el cliente RETIENE, no el IVA trasladado de la factura.
_RETENCION_ANTES = re.compile(
    r"(?<![A-Za-z])ret(?:enci[oó]n(?:es)?|enid[oa]s?|\.)?\s*(?:de(?:l)?\s+)?$", re.IGNORECASE
)
# Línea que es SOLO una etiqueta con «:» («Subtotal:», «IVA 16%:»).
_SOLO_ETIQUETA = re.compile(r"[^\d:]*(?:\d{1,2}(?:\.\d+)?\s?%[^\d:]*)?:")
# Línea que es SOLO un monto.
_SOLO_MONTO = re.compile(
    r"\$?\s?(?:\d{1,3}(?:,\d{3})+(?:\.\d{1,2})?|\d+\.\d{2})(?:\s?(?:MXN|USD))?", re.IGNORECASE
)

_ETQS_SUBTOTAL = (re.compile(r"(?<![A-Za-z])Sub\s?-?\s?total(?![A-Za-z])", re.IGNORECASE),)
_ETQS_IVA = (
    re.compile(r"(?<![A-Za-z])(?:I\.\s?V\.\s?A\.?|IVA)(?![A-Za-z])", re.IGNORECASE),
    re.compile(r"(?<![A-Za-z])(?:Total\s+(?:de\s+)?)?Impuestos?\s+trasladados?", re.IGNORECASE),
)
# En orden de preferencia; «Subtotal» e «Importe total con letra» nunca
# llegan a dar monto (el valor debe ir PEGADO a la etiqueta).
_ETQS_TOTAL = (
    re.compile(r"(?<![A-Za-z])Total\s+a\s+pagar(?![A-Za-z])", re.IGNORECASE),
    re.compile(
        r"(?<![A-Za-z])Total\s+(?:de\s+la\s+|del\s+)?(?:factura|comprobante|neto|general)"
        r"(?![A-Za-z])",
        re.IGNORECASE,
    ),
    re.compile(r"(?<![A-Za-z])Importe\s+total(?![A-Za-z])", re.IGNORECASE),
    re.compile(r"(?<![A-Za-z])Prima\s+total(?![A-Za-z])", re.IGNORECASE),
    re.compile(r"(?<![A-Za-z])(?<!Sub\s)(?<!Sub-)Total(?![A-Za-z])", re.IGNORECASE),
)
_ETQS_DESCUENTO = (re.compile(r"(?<![A-Za-z])Descuentos?(?![A-Za-z])", re.IGNORECASE),)
_ETQS_RETENCION = (
    re.compile(r"(?<![A-Za-z])(?:retenid[oa]s?|retenci[oó]n(?:es)?)(?![A-Za-z])", re.IGNORECASE),
)


def _a_float(m: re.Match[str]) -> float:
    return round(float((m.group("n") or m.group("e")).replace(",", "")), 2)


def _pila_de_etiquetas(lineas: list[str], i: int) -> tuple[int, int]:
    """[a, b]: renglones consecutivos que son SOLO etiqueta, alrededor de i."""
    a = i
    while a > 0 and _SOLO_ETIQUETA.fullmatch(lineas[a - 1]):
        a -= 1
    b = i
    while b + 1 < len(lineas) and _SOLO_ETIQUETA.fullmatch(lineas[b + 1]):
        b += 1
    return a, b


def _monto_en_columna(lineas: list[str], i: int, a: int, b: int) -> float | None:
    """Etiquetas apiladas y montos apilados debajo (pypdf separa en renglones
    los objetos de texto de una tabla): «Subtotal:\\nIVA:\\nTotal:\\n1,000.00
    \\n160.00\\n1,160.00». Se empareja por posición."""
    valores: list[str] = []
    j = b + 1
    while j < len(lineas) and _SOLO_MONTO.fullmatch(lineas[j]):
        valores.append(lineas[j])
        j += 1
    if len(valores) < b - a + 1:
        return None
    m = _MONTO_RE.search(valores[i - a])
    return _a_float(m) if m else None


def _monto_en_fila(lineas: list[str], i: int, inicio_etiqueta: int) -> float | None:
    """«Subtotal: IVA: Total:» y en la línea siguiente «1,000.00 160.00
    1,160.00»: la etiqueta k se lleva el monto k (mismo número de ambos)."""
    linea = lineas[i]
    if i + 1 >= len(lineas):
        return None
    montos = list(_MONTO_RE.finditer(lineas[i + 1]))
    if len(montos) != linea.count(":"):
        return None
    return _a_float(montos[linea[:inicio_etiqueta].count(":")])


def _monto_tras_etiqueta(
    lineas: list[str], i: int, m: re.Match[str], prefijo: re.Pattern[str]
) -> float | None:
    linea = lineas[i]
    resto = linea[m.end() :]
    p = prefijo.match(resto)
    corte = p.end() if p else 0
    monto = _MONTO_RE.match(resto[corte:])
    if monto:
        return _a_float(monto)
    if linea.count(":") >= 2 and not _MONTO_RE.search(linea):
        en_fila = _monto_en_fila(lineas, i, m.start())
        if en_fila is not None:
            return en_fila
    if resto[corte:].strip():
        return None  # otra cosa pegada a la etiqueta: no es su monto
    if _SOLO_ETIQUETA.fullmatch(linea):
        a, b = _pila_de_etiquetas(lineas, i)
        if b > a:  # varias etiquetas apiladas: solo vale el emparejamiento
            return _monto_en_columna(lineas, i, a, b)
    # «Total:» y el monto SOLO en la línea siguiente (un único monto: con
    # varios no se sabe cuál es el suyo).
    if ":" in resto[:corte] and i + 1 < len(lineas):
        sig = lineas[i + 1]
        p2 = prefijo.match(sig)
        monto = _MONTO_RE.match(sig[p2.end() if p2 else 0 :])
        if monto and len(list(_MONTO_RE.finditer(sig))) == 1:
            return _a_float(monto)
    return None


def _montos_por_etiqueta(
    lineas: list[str],
    etiquetas: tuple[re.Pattern[str], ...],
    prefijo: re.Pattern[str],
    excluir_antes: re.Pattern[str] | None = None,
) -> list[list[float]]:
    """Por cada etiqueta (en su orden de preferencia), los montos hallados en
    orden de aparición. `excluir_antes` descarta la etiqueta si lo que la
    precede en su renglón lo cumple («Retención IVA» no es el IVA)."""
    return [
        [
            v
            for i, linea in enumerate(lineas)
            for m in etiqueta.finditer(linea)
            if not (excluir_antes and excluir_antes.search(linea[: m.start()]))
            if (v := _monto_tras_etiqueta(lineas, i, m, prefijo)) is not None
        ]
        for etiqueta in etiquetas
    ]


def _monto_de(
    lineas: list[str],
    etiquetas: tuple[re.Pattern[str], ...],
    prefijo: re.Pattern[str],
    excluir_antes: re.Pattern[str] | None = None,
) -> float | None:
    """Monto de la etiqueta más específica; dentro de ella el ÚLTIMO hallado
    (el cuadro de totales va al final, después de los conceptos)."""
    for hallados in _montos_por_etiqueta(lineas, etiquetas, prefijo, excluir_antes):
        if hallados:
            return hallados[-1]
    return None


def _ajustes_cuadre(lineas: list[str]) -> list[float]:
    """Lo que puede separar subtotal + IVA del total impreso: descuento y
    retenciones (IVA y/o ISR). Devuelve los NETOS posibles (descuento +
    retenciones) contra los que se acepta el cuadre: sin ajuste, cada
    retención sola, todas sumadas y todas menos la mayor (cuando además de
    cada renglón viene impreso «Total de retenciones»)."""
    descuento = _monto_de(lineas, _ETQS_DESCUENTO, _PREFIJO_MONTO) or 0.0
    grupos = _montos_por_etiqueta(lineas, _ETQS_RETENCION, _PREFIJO_RETENCION)
    rets = [v for grupo in grupos for v in grupo]
    opciones_ret = {0.0, *rets}
    if rets:
        opciones_ret |= {sum(rets), sum(rets) - max(rets)}
    return sorted({round(d + r, 2) for d in {0.0, descuento} for r in opciones_ret})


def _cuadra(subtotal: float, iva: float, total: float, ajustes: list[float]) -> bool:
    return any(abs(subtotal + iva - a - total) <= TOL_CUADRE for a in ajustes)


def _extraer_total(
    lineas: list[str], subtotal: float | None, iva: float | None, ajustes: list[float]
) -> float | None:
    """Total impreso. Con subtotal e IVA leídos gana, en orden de preferencia
    de etiqueta, el monto que CUADRA (un «Importe total: 1,000.00» de un
    renglón de concepto no le gana al «Total: 1,160.00» del cuadro). Si
    ninguno cuadra, el de siempre (etiqueta más específica, último hallado) y
    el aviso de cuadre lo señala."""
    candidatos = _montos_por_etiqueta(lineas, _ETQS_TOTAL, _PREFIJO_MONTO)
    if subtotal is not None and iva is not None:
        for hallados in candidatos:
            for v in reversed(hallados):
                if _cuadra(subtotal, iva, v, ajustes):
                    return v
    for hallados in candidatos:
        if hallados:
            return hallados[-1]
    return None


def _fmt_dinero(n: float) -> str:
    """Regla de dinero del sistema: entero sin decimales, si no EXACTAMENTE 2."""
    n = round(n, 2) + 0.0
    return f"${n:,.0f}" if n == int(n) else f"${n:,.2f}"


def _aviso_cuadre(
    subtotal: float | None, iva: float | None, total: float | None, ajustes: list[float]
) -> str | None:
    """Subtotal + IVA debe dar el total (salvo descuento/retenciones impresos):
    si no, algo se leyó mal y Mari debe revisarlo ANTES de guardar."""
    if subtotal is None or iva is None or total is None:
        return None
    if _cuadra(subtotal, iva, total, ajustes):
        return None
    return (
        f"Subtotal {_fmt_dinero(subtotal)} + IVA {_fmt_dinero(iva)} no da el total impreso "
        f"{_fmt_dinero(total)}: revisa los montos antes de guardar."
    )


# --- Moneda, método, forma y tipo -------------------------------------------

_ETQ_MONEDA = re.compile(r"(?<![A-Za-z])Moneda(?:\s+del\s+comprobante)?\s*:?\s*", re.IGNORECASE)
_CODIGO_MONEDA = re.compile(r"\(?([A-Z]{3})(?![A-Za-z])")
# Catálogo c_Moneda del SAT (= ISO 4217), acotado a lo que un PDF plausible
# trae. Un código de 3 letras FUERA de aquí no es moneda: en el encabezado de
# la tabla de conceptos «CANTIDAD MONEDA IVA IMPORTE» el «IVA» no debe cortar
# la búsqueda antes de llegar a «MONEDA: MXN».
_MONEDAS_ISO = {
    "MXN", "USD", "XXX", "EUR", "CAD", "GBP", "JPY", "CHF", "CNY", "AUD", "NZD",
    "BRL", "ARS", "COP", "CLP", "PEN", "UYU", "BOB", "PYG", "VES", "GTQ", "CRC",
    "DOP", "HNL", "NIO", "PAB", "CUP", "BZD", "SEK", "NOK", "DKK", "KRW", "INR",
    "HKD", "SGD", "TWD",
}  # fmt: skip
# «DLS» / «DLLS» / «US$»: como se escribe «dólares» en México.
_ALIAS_USD = re.compile(r"\(?(?:DLLS?|DLS|US\s?\$)(?![A-Za-z])", re.IGNORECASE)


def _codigo_moneda(valor: str) -> str | None:
    m = _CODIGO_MONEDA.match(valor)
    if m and m.group(1) in _MONEDAS_ISO:
        return m.group(1)
    if _ALIAS_USD.match(valor):
        return "USD"
    v = unicodedata.normalize("NFKD", valor[:40]).encode("ascii", "ignore").decode().lower()
    if re.match(r"(?:pesos?(?![a-z])|m\.?\s?n\.?(?![a-z])|moneda nacional)", v):
        return "MXN"
    if re.match(r"dolar", v):
        return "USD"
    if re.match(r"euro", v):
        return "EUR"
    return None


def _extraer_moneda(lineas: list[str]) -> tuple[str | None, str | None]:
    for i, linea in enumerate(lineas):
        for m in _ETQ_MONEDA.finditer(linea):
            codigo = _codigo_moneda(_valor_tras(lineas, i, m.end()))
            if codigo is None:
                continue  # «MONEDA NACIONAL» del importe con letra, etc.
            if codigo in ("MXN", "XXX"):
                return "MXN", None
            if codigo == "USD":
                return "USD", None
            return None, f"Moneda {codigo}: el registro solo maneja MXN y USD."
    return None, None


_ETQ_METODO = re.compile(r"M[eé]todo\s+de\s+pago\s*:?\s*", re.IGNORECASE)
_ETQ_FORMA = re.compile(r"Forma\s+de\s+pago\s*:?\s*", re.IGNORECASE)
# Catálogo c_FormaPago del SAT (evita tomar un «16» cualquiera).
_FORMAS_PAGO = {
    "01", "02", "03", "04", "05", "06", "08", "12", "13", "14", "15", "17",
    "23", "24", "25", "26", "27", "28", "29", "30", "31", "99",
}  # fmt: skip
_FORMAS_POR_NOMBRE = (
    (r"efectivo", "01"),
    (r"cheque", "02"),
    (r"transferencia", "03"),
    (r"tarjeta de credito", "04"),
    (r"monedero", "05"),
    (r"dinero electronico", "06"),
    (r"compensacion", "17"),
    (r"tarjeta de debito", "28"),
    (r"tarjeta de servicio", "29"),
    (r"anticipo", "30"),
    (r"por definir", "99"),
)


def _sin_acentos(s: str) -> str:
    return unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode().lower()


# Dónde empieza la SIGUIENTE etiqueta en la misma línea del valor de pago.
_ETIQUETA_SIGUIENTE_PAGO = re.compile(
    r"\s(?:(?:Forma|M[eé]todo|Condiciones)\s+de\s+pago|Moneda|Tipo\s+de|Uso\s+(?:de\s+)?CFDI"
    r"|Lugar|Fecha|Cuenta|Banco|R[eé]gimen|Exportaci[oó]n|[^\s:]+:)",
    re.IGNORECASE,
)


def _hasta_siguiente_etiqueta(valor: str) -> str:
    """«PUE - Pago en una sola exhibición Forma de pago: 03» ⇒ hasta antes de
    «Forma de pago:»."""
    m = _ETIQUETA_SIGUIENTE_PAGO.search(valor)
    return valor[: m.start()] if m else valor


def _extraer_metodo(lineas: list[str]) -> str | None:
    for i, linea in enumerate(lineas):
        for m in _ETQ_METODO.finditer(linea):
            valor = _hasta_siguiente_etiqueta(_valor_tras(lineas, i, m.end()))
            codigo = re.search(r"(?<![A-Za-z])(PUE|PPD)(?![A-Za-z])", valor, re.IGNORECASE)
            if codigo:
                return codigo.group(1).upper()
            v = _sin_acentos(valor)
            if "una sola exhibicion" in v:
                return "PUE"
            if "parcialidades" in v or "diferido" in v:
                return "PPD"
    return None


def _extraer_forma(lineas: list[str]) -> str | None:
    for i, linea in enumerate(lineas):
        for m in _ETQ_FORMA.finditer(linea):
            valor = _hasta_siguiente_etiqueta(_valor_tras(lineas, i, m.end()))
            codigo = re.match(r"\(?(\d{1,2})(?!\d)", valor) or re.search(r"\((\d{1,2})\)", valor)
            if codigo and codigo.group(1).zfill(2) in _FORMAS_PAGO:
                return codigo.group(1).zfill(2)
            v = _sin_acentos(valor)
            for patron, clave in _FORMAS_POR_NOMBRE:
                if re.search(patron, v):
                    return clave
    return None


# «Tipo de comprobante» (PAC) · «Efecto de comprobante» (factura del SAT).
_ETQ_TIPO = re.compile(
    r"(?<![A-Za-z])(?:Tipo|Efecto)\s+(?:del?\s+)?comprobante\s*:?\s*", re.IGNORECASE
)
_TIPOS_NO_INGRESO = {"E": "Egreso", "P": "Pago", "T": "Traslado", "N": "Nómina"}


def _aviso_tipo_comprobante(lineas: list[str]) -> str | None:
    """El registro es de facturas de INGRESO: una nota de crédito (Egreso) o
    un complemento de Pago no deben entrar sin que Mari lo note."""
    for i, linea in enumerate(lineas):
        for m in _ETQ_TIPO.finditer(linea):
            valor = _sin_acentos(_valor_tras(lineas, i, m.end()))
            t = re.match(r"\(?([iepnt])\)?(?![a-z])|(ingreso|egreso|pago|traslado|nomina)", valor)
            if not t:
                continue
            letra = (t.group(1) or t.group(2))[0].upper()
            if letra in _TIPOS_NO_INGRESO:
                return (
                    f"Este PDF es de un CFDI de tipo {_TIPOS_NO_INGRESO[letra]}, "
                    "no una factura de ingreso: revisa que sea el archivo correcto."
                )
            return None
    return None
