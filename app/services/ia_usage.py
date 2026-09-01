"""Helper único: extrae el consumo de tokens de una respuesta del SDK.

Contrato: NUNCA levanta — el registro de consumo es best-effort y un fallo
aquí jamás debe romper la llamada de visión que sí funcionó.
"""

from app.schemas.uso_ia import UsoIA


def _int_de(obj: object, campo: str) -> int:
    valor = getattr(obj, campo, 0)
    try:
        return int(valor) if valor is not None else 0
    except (TypeError, ValueError):
        return 0


def uso_ia_de(resp: object) -> UsoIA:
    """UsoIA a partir de `resp` (Message del SDK de Anthropic).

    Usa `resp.model` (el modelo REAL servido, no settings.anthropic_model) y
    `resp.usage.*` con getattr defensivo (None → 0).
    """
    try:
        modelo = getattr(resp, "model", "") or ""
        usage = getattr(resp, "usage", None)
        if usage is None:
            return UsoIA(modelo=str(modelo))
        return UsoIA(
            modelo=str(modelo),
            input_tokens=_int_de(usage, "input_tokens"),
            output_tokens=_int_de(usage, "output_tokens"),
            cache_creation_input_tokens=_int_de(usage, "cache_creation_input_tokens"),
            cache_read_input_tokens=_int_de(usage, "cache_read_input_tokens"),
        )
    except Exception:  # noqa: BLE001 — degradar, nunca romper la lectura
        return UsoIA()
