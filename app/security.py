"""Autenticacion servicio-a-servicio entre NestJS y el microservicio.

Por compatibilidad se aceptan dos esquemas de token:
  - X-Internal-Token (INTERNAL_SHARED_TOKEN) — routers de bloques 1-4.
  - X-Service-Token (SERVICE_TOKEN) — routers de dramirez (pdf/reparto).
Pueden configurarse con el mismo valor en ambas variables de entorno.
"""

import logging
import secrets

from fastapi import Header, HTTPException, status

from app.config import get_settings
from app.services.modelo_ia import fijar_modelo_de_peticion, modelo_pedido

logger = logging.getLogger("security")


async def require_internal_token(
    x_internal_token: str | None = Header(default=None),
    x_ia_modelo: str | None = Header(default=None, alias="X-IA-Modelo"),
) -> None:
    """Valida el header X-Internal-Token contra INTERNAL_SHARED_TOKEN y deja el
    modelo de Claude de ESTA petición (header X-IA-Modelo, 2-oct-2026).

    Es `async def` A PROPÓSITO: FastAPI corre una dependencia `def` en un hilo
    con una COPIA del contexto, y el ContextVar fijado ahí se pierde antes de
    llegar al endpoint (se leería siempre el modelo del servidor). Async, se
    fija en el contexto de la petición y el endpoint (`def`, en su hilo) lo
    hereda. Se fija SIEMPRE —None sin header—: jamás queda el de otra petición.
    """
    fijar_modelo_de_peticion(x_ia_modelo)
    expected = get_settings().internal_shared_token
    if not expected:
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail="INTERNAL_SHARED_TOKEN no configurado",
        )
    if not x_internal_token or not secrets.compare_digest(x_internal_token, expected):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Token interno inválido",
        )
    # Solo con token válido (un anónimo no llena el log). El id no es secreto.
    if x_ia_modelo is not None and modelo_pedido(x_ia_modelo) is None:
        logger.warning(
            "X-IA-Modelo inválido (%r): se usa el modelo del servidor", x_ia_modelo[:100]
        )


def require_service_token(x_service_token: str | None = Header(default=None)) -> None:
    """Dependency de FastAPI que valida el token de servicio (X-Service-Token)."""
    settings = get_settings()
    if not settings.service_token:
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail="SERVICE_TOKEN no configurado en el microservicio",
        )
    if not x_service_token or not secrets.compare_digest(x_service_token, settings.service_token):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Token de servicio invalido o ausente",
        )
