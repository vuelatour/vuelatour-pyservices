"""Ensamblado genérico de archivos (base64) en un .zip en memoria."""

from __future__ import annotations

import base64
import zipfile
from io import BytesIO

from app.schemas.zip import ZipRequest


def render_zip(req: ZipRequest) -> bytes:
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for archivo in req.archivos:
            try:
                contenido = base64.b64decode(archivo.contenido_b64)
            except Exception:
                continue
            zf.writestr(archivo.nombre, contenido)
    return buf.getvalue()
