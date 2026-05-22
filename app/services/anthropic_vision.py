import json
from functools import lru_cache

import anthropic

from app.config import get_settings
from app.schemas.vision import TacometroRequest, TacometroResponse

_SYSTEM = (
    "Eres un asistente de operaciones de aviación. Lees el HORÓMETRO/TACÓMETRO "
    "(HOBBS) de una aeronave a partir de una foto del instrumento. El valor es "
    "un contador de horas de operación con normalmente una décima (ej. 1234.5). "
    "Devuelves SOLO un objeto JSON, sin texto adicional ni ```fences```, con las "
    "claves exactas:\n"
    '  "lectura": número con la lectura en horas, o null si no se puede leer.\n'
    '  "confianza": número entre 0 y 1.\n'
    '  "legible": true/false según si el display se distingue.\n'
    '  "notas": string breve en español (reflejo, borrosa, dígito parcial, etc.).\n'
    "Si dudas entre dos dígitos, elige el más probable y baja la confianza. "
    "No inventes dígitos que no ves: si faltan, usa null y explica en notas."
)

_USER_PROMPT = (
    "Lee el contador de horas (HOBBS/tacómetro) en esta foto y responde con el "
    "JSON indicado. Considera el último dígito como décima si el display lo separa."
)


@lru_cache
def _client() -> anthropic.Anthropic:
    s = get_settings()
    return anthropic.Anthropic(
        api_key=s.anthropic_api_key,
        timeout=s.anthropic_timeout_s,
        max_retries=s.anthropic_max_retries,
    )


def _image_block(req: TacometroRequest) -> dict:
    if req.image_base64:
        return {
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": req.media_type,
                "data": req.image_base64,
            },
        }
    return {"type": "image", "source": {"type": "url", "url": req.image_url}}


def _extract_json(text: str) -> dict:
    t = text.strip()
    if t.startswith("```"):
        t = t.split("```", 2)[1] if "```" in t[3:] else t[3:]
        t = t.removeprefix("json").strip()
    start, end = t.find("{"), t.rfind("}")
    if start == -1 or end == -1:
        raise ValueError(f"Respuesta sin JSON: {text[:200]}")
    return json.loads(t[start : end + 1])


def leer_tacometro(req: TacometroRequest) -> TacometroResponse:
    s = get_settings()
    resp = _client().messages.create(
        model=s.anthropic_model,
        max_tokens=512,
        system=[{"type": "text", "text": _SYSTEM, "cache_control": {"type": "ephemeral"}}],
        messages=[
            {
                "role": "user",
                "content": [_image_block(req), {"type": "text", "text": _USER_PROMPT}],
            }
        ],
    )
    text = next((b.text for b in resp.content if b.type == "text"), "")
    data = _extract_json(text)

    lectura = data.get("lectura")
    return TacometroResponse(
        lectura=float(lectura) if isinstance(lectura, (int, float)) else None,
        confianza=float(data.get("confianza", 0.0)),
        legible=bool(data.get("legible", lectura is not None)),
        notas=str(data.get("notas", "")),
        modelo=s.anthropic_model,
    )
