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

    # Anthropic (Claude) — vision para lectura de tacómetros.
    anthropic_api_key: str = ""
    anthropic_model: str = "claude-sonnet-4-6"
    anthropic_timeout_s: float = 30.0
    anthropic_max_retries: int = 3

    # Token compartido pyservices <-> NestJS (header X-Internal-Token).
    internal_shared_token: str = ""


@lru_cache
def get_settings() -> Settings:
    return Settings()
