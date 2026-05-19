"""Autenticacion servicio-a-servicio: NestJS manda el header X-Service-Token."""

from fastapi import Header, HTTPException, status

from app.config import get_settings


def require_service_token(x_service_token: str | None = Header(default=None)) -> None:
    """Dependency de FastAPI que valida el token de servicio."""
    settings = get_settings()
    if not settings.service_token:
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail="SERVICE_TOKEN no configurado en el microservicio",
        )
    if x_service_token != settings.service_token:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Token de servicio invalido o ausente",
        )
