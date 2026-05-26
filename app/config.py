"""Configuracion del microservicio, leida de variables de entorno / .env."""

from functools import lru_cache

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=(".env.local", ".env"),
        env_file_encoding="utf-8",
        extra="ignore",
    )

    app_env: str = "development"
    port: int = 8000

    # Anthropic (Claude) — vision para lectura de tacómetros y extracción.
    anthropic_api_key: str = ""
    anthropic_model: str = "claude-sonnet-4-6"
    anthropic_timeout_s: float = 30.0
    anthropic_max_retries: int = 3

    # Token compartido pyservices <-> NestJS (header X-Internal-Token).
    internal_shared_token: str = ""
    # Token de servicio de dramirez (header X-Service-Token). Puede ser igual al anterior.
    service_token: str = ""

    # FEL (PAC) — timbrado CFDI 4.0. Vacío = facturación deshabilitada.
    fel_usuario: str = ""
    fel_password: str = ""
    fel_wsdl_url: str = "https://app.fel.mx/WSTimbrado33Test/WSCFDI33.svc?WSDL"  # pruebas por defecto
    fel_modo: str = "test"  # test | prod
    fel_timeout_s: float = 40.0

    @property
    def fel_configurado(self) -> bool:
        return bool(self.fel_usuario and self.fel_password)


@lru_cache
def get_settings() -> Settings:
    return Settings()
