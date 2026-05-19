"""Configuracion del microservicio, leida de variables de entorno / .env."""

from functools import lru_cache

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=".env", extra="ignore")

    app_env: str = "development"
    port: int = 8000
    # Token compartido con NestJS para autenticar las llamadas entre servicios.
    service_token: str = ""


@lru_cache
def get_settings() -> Settings:
    return Settings()
