import secrets

from fastapi import Header, HTTPException, status

from app.config import get_settings


def require_internal_token(x_internal_token: str | None = Header(default=None)) -> None:
    """Valida el header X-Internal-Token contra INTERNAL_SHARED_TOKEN.

    Si el token no está configurado en el entorno, se rechaza todo: evita
    exponer el endpoint sin protección por un .env incompleto.
    """
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
