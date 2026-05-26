"""Autenticacion servicio-a-servicio entre NestJS y el microservicio.

Por compatibilidad se aceptan dos esquemas de token:
  - X-Internal-Token (INTERNAL_SHARED_TOKEN) — routers de bloques 1-4.
  - X-Service-Token (SERVICE_TOKEN) — routers de dramirez (pdf/reparto).
Pueden configurarse con el mismo valor en ambas variables de entorno.
"""

import secrets

from fastapi import Header, HTTPException, status

from app.config import get_settings


def require_internal_token(x_internal_token: str | None = Header(default=None)) -> None:
    """Valida el header X-Internal-Token contra INTERNAL_SHARED_TOKEN."""
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
