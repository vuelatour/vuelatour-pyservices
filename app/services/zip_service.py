"""Ensamblado genérico de archivos (base64) en un .zip en memoria."""

from __future__ import annotations

import base64
import zipfile
from io import BytesIO

from app.schemas.zip import ZipRequest


def _decodificar(b64: str) -> bytes:
    # validate=True: sin él, b64decode descarta caracteres inválidos en
    # silencio y el archivo entraría corrupto al zip (se toleran saltos
    # de línea/espacios, comunes en base64 envuelto).
    return base64.b64decode("".join(b64.split()), validate=True)


def render_zip(req: ZipRequest) -> bytes:
    # Un base64 ilegible NO se omite en silencio: el zip del cierre mensual
    # saldría incompleto (sin facturas/reportes) sin que nadie lo note.
    ilegibles: list[str] = []
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for archivo in req.archivos:
            try:
                contenido = _decodificar(archivo.contenido_b64)
            except Exception:
                ilegibles.append(archivo.nombre)
                continue
            zf.writestr(archivo.nombre, contenido)
    if ilegibles:
        raise ValueError(
            "No se pudo leer el contenido (base64 inválido) de: "
            + ", ".join(ilegibles)
            + ". El zip saldría incompleto; corrige esos archivos y reintenta."
        )
    return buf.getvalue()
