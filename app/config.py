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
    # Opus tarda 20-60s en una factura densa: 30s cortaba la lectura a medias.
    # Si la var ANTHROPIC_TIMEOUT_S existe en el entorno, debe ser >= 90.
    anthropic_timeout_s: float = 90.0
    anthropic_max_retries: int = 2

    # Token compartido pyservices <-> NestJS (header X-Internal-Token).
    internal_shared_token: str = ""
    # Token de servicio de dramirez (header X-Service-Token). Puede ser igual al anterior.
    service_token: str = ""

    # FEL (PAC) — timbrado CFDI 4.0. Vacío = facturación deshabilitada.
    # OJO: el usuario de timbrado es DISTINTO al usuario del portal FEL En Línea.
    fel_usuario: str = ""
    fel_password: str = ""
    fel_wsdl_url: str = "https://app.fel.mx/WSTimbrado33Test/WSCFDI33.svc?WSDL"  # pruebas por defecto
    fel_modo: str = "test"  # test | prod

    # PAC de timbrado: 'fel' (SOAP, requiere Usuario de Timbrado) o
    # 'facturama' (REST api-lite multiemisor; Basic Auth con el usuario y
    # contraseña de la cuenta — Facturama NO usa API key).
    facturacion_pac: str = "fel"  # fel | facturama
    facturama_user: str = ""
    facturama_password: str = ""
    facturama_modo: str = "sandbox"  # sandbox | prod
    fel_timeout_s: float = 40.0
    # PFX (PKCS12) de cancelación, generado a partir del CSD del emisor
    # (guía "creación PFX" de FEL). Sin él no se puede cancelar ante el SAT.
    fel_pfx_b64: str = ""
    fel_pfx_password: str = ""

    @property
    def fel_configurado(self) -> bool:
        return bool(self.fel_usuario and self.fel_password)


@lru_cache
def get_settings() -> Settings:
    return Settings()
