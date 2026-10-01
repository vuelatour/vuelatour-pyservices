"""Errores de Claude → texto para el operador (FUENTE ÚNICA, 1-oct-2026).

Caso real: un ADMIN vio «IA no disponible: Claude no disponible (400).
Captura a mano.» y la oficina entendió que «Claude no está disponible». La
causa era el SALDO de créditos agotado: Anthropic responde HTTP 400
``invalid_request_error`` con «Your credit balance is too low to access the
Anthropic API…». El código de estado solo no le dice nada a quien opera:
aquí se traduce cada error del SDK a una frase en español que dice QUÉ pasó y
QUÉ hacer.

Todos los routers que llaman a Claude usan estas funciones (``detail`` del
502 y ``motivo`` de la constancia fiscal); el API NestJS reenvía el texto tal
cual y la app lo pinta como «IA no disponible: <texto>. Captura a mano.».
Por eso, CONTRATO: el texto NUNCA termina en punto (quien lo muestra agrega
«. Captura a mano…»: con punto final salía «..») y NO dice «captura a mano»
(lo dice quien lo muestra; la constancia fiscal lo agrega aquí porque su
respuesta degrada en vez de ser un 502).

Funciones PURAS salvo ``registrar_error_claude`` (solo escribe en el log).
Nunca se imprime la llave ni los encabezados de la petición.
"""

from __future__ import annotations

import logging
import re

# --- Textos de UI (es-MX), SIN punto final --------------------------------

TEXTO_SIN_SALDO = (
    "Sin saldo de créditos de IA en Anthropic: hay que recargar (Plans & Billing) "
    "y registrar el nuevo saldo en Configuración → Consumo de IA"
)
TEXTO_LIMITE_GASTO = (
    "Se alcanzó el límite de gasto de IA configurado en Anthropic: "
    "avisa a sistemas para que lo suban"
)
TEXTO_ARCHIVO_PESADO = (
    "La foto o el archivo pesa demasiado para la IA (máx. 5 MB por foto, "
    "32 MB por PDF): toma otra foto o recórtala"
)
TEXTO_FOTO_DIMENSIONES = (
    "La foto es demasiado grande en pixeles para la IA: tómala con menor "
    "resolución o recórtala"
)
TEXTO_FORMATO_DISTINTO = (
    "La foto llegó marcada con un formato distinto al real: reintenta y, "
    "si se repite, avisa a sistemas"
)
TEXTO_FORMATO_IMAGEN = (
    "Formato de archivo no soportado por la IA: usa una foto JPG o PNG, o un PDF"
)
TEXTO_FOTO_ILEGIBLE = "La IA no pudo leer la foto: toma otra con mejor luz y enfoque"
TEXTO_DEMASIADAS_FOTOS = "Demasiadas fotos en una sola lectura: léelas en dos tandas"
TEXTO_DOCUMENTO_LARGO = "El documento es demasiado largo para la IA: pártelo en varios archivos"
TEXTO_LLAVE_INVALIDA = "La llave de la IA no es válida o venció: avisa a sistemas"
TEXTO_MODELO_INEXISTENTE = "El modelo de IA configurado no existe: avisa a sistemas"
TEXTO_LIMITE_PETICIONES = (
    "La IA está saturada (límite de peticiones): espera un minuto y reintenta"
)
TEXTO_IA_CAIDA = "La IA está saturada o caída por el momento: reintenta en unos minutos"

# Constancia fiscal: su respuesta degrada (no es un 502) y siempre debe decir
# que los datos se capturan a mano.
SUFIJO_CAPTURA_MANUAL = "captura los datos manualmente"

# Largo máximo del mensaje original que se muestra en el caso genérico.
LARGO_MENSAJE_GENERICO = 120
# Largo máximo del mensaje en el log (un 5xx puede traer una página HTML).
LARGO_MENSAJE_LOG = 500

# --- Reglas por TEXTO del error (minúsculas; la primera que coincide gana) --
#
# El orden importa: las frases más específicas van antes. Mensajes reales de
# Anthropic que cubre cada una (todas llegan como 400 invalid_request_error):
#   «Your credit balance is too low to access the Anthropic API…»
#   «You have reached your specified API usage limits. You will regain access…»
#   «image dimensions exceed max allowed size: 8000 pixels»
#   «The image was specified using the image/jpeg media type, but the image
#    appears to be a image/png image» (la evita `media_type_real`; si vuelve,
#    el formato que tiene la foto SÍ es soportado: no se pide «usa JPG o PNG»)
#   «…image exceeds 5 MB maximum…»
#   «…media_type: Input should be 'image/jpeg', 'image/png'…»
#   «prompt is too long: …» / «input length and `max_tokens` exceed context limit…»

_REGLAS_POR_TEXTO: tuple[tuple[tuple[str, ...], str], ...] = (
    (("credit balance",), TEXTO_SIN_SALDO),
    (("usage limit",), TEXTO_LIMITE_GASTO),
    (("image dimensions",), TEXTO_FOTO_DIMENSIONES),
    (("image appears to be",), TEXTO_FORMATO_DISTINTO),
    (("image exceeds", "too large", "exceeds 5 mb", "file size"), TEXTO_ARCHIVO_PESADO),
    (("media_type", "media type"), TEXTO_FORMATO_IMAGEN),
    (("could not process image", "image could not be processed"), TEXTO_FOTO_ILEGIBLE),
    (("too many images",), TEXTO_DEMASIADAS_FOTOS),
    (("prompt is too long", "too many tokens", "context limit"), TEXTO_DOCUMENTO_LARGO),
)

# Prefijo que el SDK antepone a `e.message` («Error code: 400 - {…}» o, con
# la respuesta cerrada sin leer, solo «Error code: 400»).
_PREFIJO_SDK = re.compile(r"^error code:\s*\d+\s*(?:-\s*)?", re.IGNORECASE)


def _status(e: BaseException) -> int | None:
    status = getattr(e, "status_code", None)
    return status if isinstance(status, int) else None


def mensaje_error_claude(e: BaseException) -> str:
    """Mensaje ORIGINAL del error, en una sola línea.

    Usa ``error.message`` del cuerpo JSON de Anthropic
    (``{"type": "error", "error": {"type": …, "message": …}}``) y, si no
    viene, ``e.message`` del SDK (p. ej. «Error code: 400 - {…}» o el texto
    crudo de una respuesta que no era JSON)."""
    texto: object = None
    cuerpo = getattr(e, "body", None)
    if isinstance(cuerpo, dict):
        error = cuerpo.get("error")
        if isinstance(error, dict):
            texto = error.get("message")
    if not isinstance(texto, str) or not texto.strip():
        texto = getattr(e, "message", None)
    if not isinstance(texto, str) or not texto.strip():
        texto = str(e)
    return " ".join(texto.split())


def tipo_error_claude(e: BaseException) -> str:
    """Tipo del error de Anthropic (``invalid_request_error``,
    ``rate_limit_error``…) o, sin cuerpo, el nombre de la clase del SDK."""
    tipo = getattr(e, "type", None)
    if isinstance(tipo, str) and tipo:
        return tipo
    return type(e).__name__


def detalle_error_claude(e: BaseException) -> str:
    """Texto en español y accionable para el operador a partir de un error
    del SDK de Anthropic (``anthropic.APIStatusError`` y subclases).

    Decide primero por el TEXTO del error (saldo, tamaño, formato…) y después
    por el código HTTP; lo que no se reconoce cae en «Claude no disponible
    (<status>): <mensaje original recortado>» (o «Claude no disponible
    (<status>)» si el mensaje no tiene nada legible). Nunca termina en punto."""
    mensaje = mensaje_error_claude(e)
    texto = mensaje.lower()
    for claves, salida in _REGLAS_POR_TEXTO:
        if any(clave in texto for clave in claves):
            return salida

    status = _status(e)
    if status in (401, 403):
        return TEXTO_LLAVE_INVALIDA
    if status == 404 and "model" in texto:
        return TEXTO_MODELO_INEXISTENTE
    if status == 413:
        # «request_too_large»: el cuerpo de la petición (fotos/PDF) se pasó.
        return TEXTO_ARCHIVO_PESADO
    if status == 429:
        return TEXTO_LIMITE_PETICIONES
    if status is not None and status >= 500:
        # 500 / 503 / 529 (sobrecarga) y cualquier otro 5xx transitorio.
        return TEXTO_IA_CAIDA

    codigo = f" ({status})" if status is not None else ""
    recorte = _sin_punto_final(_mensaje_visible(mensaje)[:LARGO_MENSAJE_GENERICO])
    if not recorte:
        return f"Claude no disponible{codigo}"
    return f"Claude no disponible{codigo}: {recorte}"


def _sin_punto_final(texto: str) -> str:
    return texto.strip().rstrip(" .")


def _mensaje_visible(mensaje: str) -> str:
    """Lo que SÍ se le puede enseñar al operador del mensaje original: sin
    el prefijo «Error code: N - » del SDK y nunca un dict/lista de Python ni
    una página HTML (un cuerpo sin ``error.message`` llega como
    «Error code: 400 - {'type': 'error', …}»). El log guarda el completo."""
    visible = _PREFIJO_SDK.sub("", mensaje, count=1).strip()
    if visible.startswith(("{", "[", "<")):
        return ""
    return visible


def detalle_error_claude_captura_manual(e: BaseException) -> str:
    """Igual que ``detalle_error_claude`` pero garantiza que el texto diga
    que los datos se capturan a mano (constancia fiscal, que degrada en vez
    de responder 502). Si el texto ya lo dice, no se repite."""
    detalle = detalle_error_claude(e)
    if "captura" in detalle.lower():
        return detalle
    return f"{detalle}; {SUFIJO_CAPTURA_MANUAL}"


def registrar_error_claude(logger: logging.Logger, e: BaseException) -> None:
    """Deja en el log el status, el tipo y el mensaje del error de Claude
    (y el request-id para soporte). Nunca la llave ni los encabezados."""
    logger.warning(
        "Claude API error %s (%s): %s [request_id=%s]",
        _status(e),
        tipo_error_claude(e),
        mensaje_error_claude(e)[:LARGO_MENSAJE_LOG],
        getattr(e, "request_id", None),
    )
