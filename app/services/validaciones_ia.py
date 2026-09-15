"""Validaciones DETERMINISTAS de las respuestas de IA (15-sep-2026).

Todo lo que se puede comprobar con aritmética o con un catálogo NO se le
pregunta al modelo: se comprueba aquí, en Python, después de la respuesta.
Funciones PURAS (sin red, sin reloj propio salvo el `hoy` que se inyecta):
se prueban sin llamar a Anthropic.

POLÍTICA DE CONFIANZA (la misma en todos los endpoints):
  · Falla una validación del dato PRINCIPAL (el monto del ticket, los litros
    de la carga, el RFC de la constancia) → `confianza` se fuerza a ≤ 0.3 y
    se explica en `advertencias`. El dato NO se borra: el operador lo ve con
    el aviso y decide.
  · Falla una validación de un dato SECUNDARIO (el desglose de renglones,
    la matrícula, la terminación de tarjeta) → ese campo se limpia o se
    marca y se agrega la advertencia, pero la confianza del dato principal
    NO se castiga (un total perfectamente legible no deja de serlo porque el
    modelo se equivocara al desglosar).
"""

from __future__ import annotations

import re
import unicodedata
from datetime import date, datetime, timedelta, timezone

# Tope de confianza cuando una validación del dato principal falla.
CONFIANZA_SOSPECHOSA = 0.3
# Tolerancia absoluta en pesos para cuadrar sumas de renglones.
TOL_SUMA = 0.02
# Tolerancia relativa para identidades de producto (litros × precio).
TOL_REL = 0.01


# ---------------------------------------------------------------------------
# Texto
# ---------------------------------------------------------------------------


def normalizar_texto(s: str | None) -> str:
    """MAYÚSCULAS sin acentos y con espacios colapsados ("Cozumel"→"COZUMEL")."""
    if not s:
        return ""
    t = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode("ascii")
    return " ".join(t.upper().split())


# Prefijos de agregador que ensucian la descripción del banco.
_PREFIJOS_AGREGADOR = ("MERPAGO*", "PINPE*", "CLIP*", "BILLPOCKET*", "SEL ", "EC ", "ES ")


def normalizar_texto_banco(s: str | None) -> str:
    """Descripción del banco lista para comparar: sin acentos, sin prefijos de
    agregador y sin los números sueltos que solo son folios."""
    t = normalizar_texto(s)
    for pref in _PREFIJOS_AGREGADOR:
        if t.startswith(pref):
            t = t[len(pref) :]
    t = t.replace("*", " ")
    t = re.sub(r"\b\d{3,}\b", " ", t)
    return " ".join(t.split())


def normalizar_referencia(s: str | None) -> str:
    """Referencia bancaria comparable: solo [A-Z0-9] en mayúsculas."""
    return re.sub(r"[^A-Z0-9]", "", normalizar_texto(s))


def terminacion_de_referencia(
    referencia: str | None, terminaciones_validas: list[str] | tuple[str, ...] | None = None
) -> str | None:
    """Últimos 4 dígitos de una referencia larga, SOLO si empatan con una
    terminación de tarjeta conocida.

    '0025830577' → '0577' y '9155656256' → '6256' cuando esas terminaciones
    existen en el catálogo. Una referencia corta de 6 dígitos ('174465') que
    no empata con nada devuelve None: jamás se inventa una tarjeta.

    Sin catálogo (`terminaciones_validas` vacío) devuelve los últimos 4
    dígitos de una referencia de al menos 8 dígitos, que es la forma en que
    el banco incrusta la tarjeta; nunca de una más corta.
    """
    solo_digitos = re.sub(r"\D", "", referencia or "")
    if len(solo_digitos) < 4:
        return None
    ultimos = solo_digitos[-4:]
    validas = {re.sub(r"\D", "", t)[-4:] for t in (terminaciones_validas or []) if t}
    if validas:
        return ultimos if ultimos in validas else None
    return ultimos if len(solo_digitos) >= 8 else None


# ---------------------------------------------------------------------------
# Números
# ---------------------------------------------------------------------------


def num(v: object) -> float | None:
    """float de un valor del modelo; None si no es número (bool NO es número)."""
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace("$", "").replace(",", "").strip())
    except (TypeError, ValueError):
        return None


def cuadra(a: float | None, b: float | None, tol: float = TOL_SUMA) -> bool:
    """|a − b| ≤ tol (False si falta alguno)."""
    if a is None or b is None:
        return False
    return abs(a - b) <= tol


def validar_suma_conceptos(
    conceptos: list[dict], monto: float | None, tol: float = TOL_SUMA
) -> str | None:
    """Advertencia si el desglose no suma el total; None si cuadra o no aplica."""
    if not conceptos or monto is None:
        return None
    suma = round(sum(float(c.get("monto") or 0) for c in conceptos), 2)
    if cuadra(suma, monto, tol):
        return None
    return (
        f"El desglose suma {suma:,.2f} y el total del documento es {monto:,.2f}: "
        "se descartaron los renglones (captúralos a mano si los necesitas)."
    )


def validar_identidad_combustible(
    litros: float | None, precio_litro: float | None, total: float | None, iva_pct: float = 16.0
) -> tuple[str | None, bool]:
    """Comprueba litros × precio_litro ≈ total.

    Devuelve (advertencia, principal_sospechoso). Si la diferencia es
    EXACTAMENTE el IVA (precio sin IVA contra total con IVA) no es un error:
    se informa sin castigar la confianza.
    """
    if litros is None or precio_litro is None or total is None:
        return None, False
    if litros <= 0 or precio_litro <= 0 or total <= 0:
        return "Litros, precio por litro o total no positivos en el ticket.", True
    esperado = litros * precio_litro
    tol = max(total * TOL_REL, 1.0)
    if abs(esperado - total) <= tol:
        return None, False
    con_iva = esperado * (1 + iva_pct / 100.0)
    if abs(con_iva - total) <= tol:
        return (
            f"El precio por litro parece SIN IVA ({esperado:,.2f}) y el total CON IVA "
            f"({total:,.2f}); revisa cuál necesitas.",
            False,
        )
    return (
        f"litros × precio por litro = {esperado:,.2f} pero el ticket dice {total:,.2f}: "
        "revisa la cantidad y el precio antes de guardar.",
        True,
    )


def rango_plausible(valor: float | None, minimo: float, maximo: float) -> bool:
    return valor is not None and minimo <= valor <= maximo


# ---------------------------------------------------------------------------
# Fechas
# ---------------------------------------------------------------------------


def fecha_iso(valor: object) -> date | None:
    """`date` de un 'YYYY-MM-DD' (o 'YYYY-MM-DDTHH:MM…'); None si no lo es."""
    if not valor:
        return None
    txt = str(valor).strip()[:10]
    try:
        return datetime.strptime(txt, "%Y-%m-%d").date()
    except ValueError:
        return None


def validar_fecha_documento(
    valor: object, hoy: date | None = None, dias_atras: int = 365, dias_adelante: int = 1
) -> tuple[str | None, str | None]:
    """(fecha válida o None, advertencia o None) para la fecha de un ticket.

    Una fecha del futuro o de hace años casi siempre es un misread del año
    (2026 leído como 2016): se conserva el texto pero se avisa.
    """
    if valor is None or str(valor).strip() == "":
        return None, None
    f = fecha_iso(valor)
    if f is None:
        return None, f"La fecha «{valor}» no se entendió (se esperaba YYYY-MM-DD)."
    ref = hoy or date.today()
    if f > ref + timedelta(days=dias_adelante):
        return f.isoformat(), f"La fecha {f.isoformat()} es del futuro: verifica el año."
    if f < ref - timedelta(days=dias_atras):
        return (
            f.isoformat(),
            f"La fecha {f.isoformat()} tiene más de {dias_atras} días: verifica el año.",
        )
    return f.isoformat(), None


def dias_entre(a: object, b: object) -> int | None:
    """Días de `a` a `b` ('YYYY-MM-DD'); None si falta alguna."""
    fa, fb = fecha_iso(a), fecha_iso(b)
    if fa is None or fb is None:
        return None
    return (fb - fa).days


# Todo el sistema opera en hora de Cancún (UTC−5): comparar días en UTC
# corre la fecha 5 horas y manda los vuelos de la tarde-noche al día
# siguiente (invariante 4 del CLAUDE.md del workspace).
TZ_CANCUN = timezone(timedelta(hours=-5))


def dia_cancun(valor: object) -> date | None:
    """Día de PARED en Cancún de un ISO con o sin zona.

    Sin zona horaria se asume que ya viene en hora de Cancún (así lo captura
    la app); con zona (Z o ±HH:MM) se convierte antes de tomar el día."""
    if not valor:
        return None
    txt = str(valor).strip()
    if len(txt) == 10:
        return fecha_iso(txt)
    try:
        dt = datetime.fromisoformat(txt.replace("Z", "+00:00"))
    except ValueError:
        return fecha_iso(txt)
    if dt.tzinfo is None:
        return dt.date()
    return dt.astimezone(TZ_CANCUN).date()


def dias_entre_cancun(a: object, b: object) -> int | None:
    """Días de calendario CANCÚN entre dos instantes (ISO o 'YYYY-MM-DD')."""
    fa, fb = dia_cancun(a), dia_cancun(b)
    if fa is None or fb is None:
        return None
    return (fb - fa).days


# ---------------------------------------------------------------------------
# Catálogos
# ---------------------------------------------------------------------------


def normalizar_matricula(raw: object) -> str | None:
    """Matrícula con formato plausible (XA-/XB-/N…), en MAYÚSCULAS.

    OJO: las mexicanas NO llevan dígitos (XB-PEV, XA-VGV): exigirlos —como
    hacía el parser del ticket— tiraba en silencio la matrícula de media
    flota. Las estadounidenses sí: N4142R, N58BT.
    """
    if not isinstance(raw, str):
        return None
    m = raw.strip().upper().replace(" ", "")
    if not 4 <= len(m) <= 8:
        return None
    if re.fullmatch(r"X[AB]-?[A-Z]{3}", m):
        return m if "-" in m else f"{m[:2]}-{m[2:]}"
    if re.fullmatch(r"N[0-9]{1,5}[A-Z]{0,2}", m):
        return m
    return None


def matricula_conocida(matricula: str | None, flota: list[str] | tuple[str, ...]) -> bool:
    if not matricula:
        return False
    objetivo = matricula.upper().replace("-", "")
    return any(objetivo == m.upper().replace("-", "") for m in flota)


def terminacion_4(raw: object) -> str | None:
    """Últimos 4 dígitos de una terminación de tarjeta; None si no hay 4."""
    if raw is None or isinstance(raw, bool):
        return None
    d = re.sub(r"\D", "", str(raw))
    return d[-4:] if len(d) >= 4 else None


_RFC_RE = re.compile(r"^[A-ZÑ&]{3,4}\d{6}[A-Z0-9]{3}$")
_RFC_TABLA = "0123456789ABCDEFGHIJKLMN&OPQRSTUVWXYZ"


def _valor_rfc(ch: str) -> int:
    if ch == " ":
        return 37
    if ch == "Ñ":
        return 38
    idx = _RFC_TABLA.find(ch)
    return idx if idx >= 0 else 0


def rfc_digito_verificador(rfc: str) -> str | None:
    """Dígito verificador que le TOCA a un RFC (última posición)."""
    base = rfc[:-1]
    if len(base) not in (11, 12):
        return None
    if len(base) == 11:  # persona moral: se rellena con un espacio al frente
        base = " " + base
    suma = sum(_valor_rfc(ch) * (13 - i) for i, ch in enumerate(base))
    residuo = suma % 11
    if residuo == 0:
        return "0"
    if residuo == 1:
        return "A"
    return str(11 - residuo)


def validar_rfc(raw: object) -> tuple[str | None, str | None]:
    """(RFC normalizado o None, advertencia o None).

    Estructura (3-4 letras + AAMMDD + homoclave) y DÍGITO VERIFICADOR. Un RFC
    inventado revienta el timbrado semanas después: mejor null y que el
    operador lo escriba.
    """
    if not isinstance(raw, str) or not raw.strip():
        return None, None
    rfc = raw.strip().upper().replace(" ", "").replace("-", "")
    if not _RFC_RE.match(rfc):
        return None, f"El RFC «{rfc}» no tiene forma de RFC: captúralo a mano."
    esperado = rfc_digito_verificador(rfc)
    if esperado is not None and rfc[-1] != esperado:
        return None, (
            f"El RFC «{rfc}» no pasa el dígito verificador (debería terminar en "
            f"«{esperado}»): revísalo en la constancia."
        )
    return rfc, None


# ---------------------------------------------------------------------------
# Confianza
# ---------------------------------------------------------------------------


def confianza_de(valor: object, default: float = 0.0) -> float:
    """Confianza 0..1 tolerante a basura del modelo (null, string, 1.7…)."""
    v = num(valor)
    if v is None:
        return default
    return min(max(v, 0.0), 1.0)


def confianza_calibrada(confianza: float, sospechoso: bool) -> float:
    """Tope duro cuando el dato principal no pasó su validación."""
    return min(confianza, CONFIANZA_SOSPECHOSA) if sospechoso else confianza


def limpiar_advertencias(valores: list[str | None]) -> list[str]:
    """Advertencias no vacías, sin repetir y en orden."""
    out: list[str] = []
    for v in valores:
        if v and v not in out:
            out.append(v)
    return out
