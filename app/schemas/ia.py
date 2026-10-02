"""Esquemas de la configuración de IA (2-oct-2026)."""

from pydantic import BaseModel, Field


class IaModeloResponse(BaseModel):
    """Respuesta de `GET /ia/modelo`.

    El API la usa para mostrar en Configuración → Créditos de IA cuál es el
    modelo del servidor (lo que se usa cuando nadie eligió otro)."""

    default_servidor: str = Field(
        description="Modelo del servidor: variable de entorno ANTHROPIC_MODEL."
    )
    efectivo: str = Field(
        description=(
            "Modelo que usaría ESTA petición: el del header X-IA-Modelo si es "
            "válido; si no, el del servidor."
        )
    )
