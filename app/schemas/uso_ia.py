"""Consumo de tokens de una llamada a Claude (campo ADITIVO en respuestas IA).

El API (NestJS) usa este bloque para registrar el costo en `ia_uso`. Los
campos cache_* existen porque TODOS los system prompts llevan cache_control
ephemeral: `input_tokens` NO incluye los tokens de caché (creación 1.25x,
lectura 0.10x de la tarifa input) — sin ellos el costo se subestima.
"""

from pydantic import BaseModel, Field


class UsoIA(BaseModel):
    modelo: str = Field(default="", description="Modelo REAL servido (resp.model)")
    input_tokens: int = Field(default=0, description="Tokens de entrada (sin caché)")
    output_tokens: int = Field(default=0, description="Tokens de salida")
    cache_creation_input_tokens: int = Field(
        default=0, description="Tokens escritos a caché (se cobran a 1.25x input)"
    )
    cache_read_input_tokens: int = Field(
        default=0, description="Tokens leídos de caché (se cobran a 0.10x input)"
    )
