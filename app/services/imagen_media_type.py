"""Tipo REAL de una imagen en base64 por sus bytes mágicos (1-oct-2026).

El ``media_type`` llega del cliente sin revisar (el API lo pasa tal cual):
una captura PNG etiquetada ``image/jpeg`` hace que Anthropic responda 400
«The image was specified using the image/jpeg media type, but the image
appears to be a image/png image» y la lectura muere sin que el operador
pueda hacer nada. Antes de armar el bloque ``image`` se deduce el tipo por
los primeros bytes; si no se reconoce, se respeta el declarado.

Función PURA: solo decodifica los primeros caracteres del base64.
"""

from __future__ import annotations

import base64
import binascii

# 32 caracteres de base64 = 24 bytes: alcanzan para la firma de WEBP (12).
_CARACTERES_A_LEER = 32


def _primeros_bytes(data_base64: str) -> bytes:
    texto = data_base64.lstrip()
    # Tolera un data URL («data:image/png;base64,…») aunque el API no lo mande.
    if texto.startswith("data:"):
        coma = texto.find(",", 0, 100)
        if coma != -1:
            texto = texto[coma + 1 :]
    trozo = "".join(texto[:256].split())[:_CARACTERES_A_LEER]
    trozo = trozo[: len(trozo) - len(trozo) % 4]
    if not trozo:
        return b""
    try:
        return base64.b64decode(trozo, validate=False)
    except (binascii.Error, ValueError):
        return b""


def media_type_real(data_base64: str | None, declarado: str | None) -> str | None:
    """``image/jpeg`` / ``image/png`` / ``image/gif`` / ``image/webp`` según
    los bytes; si no se reconoce (o no hay datos), el ``declarado``."""
    if not data_base64:
        return declarado
    cabeza = _primeros_bytes(data_base64)
    if cabeza.startswith(b"\xff\xd8\xff"):
        return "image/jpeg"
    if cabeza.startswith(b"\x89PNG\r\n\x1a\n"):
        return "image/png"
    if cabeza.startswith((b"GIF87a", b"GIF89a")):
        return "image/gif"
    if len(cabeza) >= 12 and cabeza[:4] == b"RIFF" and cabeza[8:12] == b"WEBP":
        return "image/webp"
    return declarado
