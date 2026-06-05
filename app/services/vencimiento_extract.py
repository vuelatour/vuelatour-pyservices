import json
from functools import lru_cache

import anthropic

from app.config import get_settings
from app.schemas.vencimiento import VencimientoExtraerRequest, VencimientoExtraerResponse

_SYSTEM = (
    "Eres un asistente documental para una empresa de aviación. A partir de un "
    "documento de vencimiento renovado (póliza de seguro, certificado, permiso o "
    "tarjeta de circulación de una aeronave o vehículo) extraes sus datos clave. "
    "Devuelves SOLO un objeto JSON, sin texto adicional ni ```fences```, con las "
    "claves exactas:\n"
    '  "matricula": matrícula/placa de la aeronave o vehículo (ej. "XA-ABC"), o null.\n'
    '  "tipo_documento": tipo del documento (ej. "Póliza de seguro", "Seguro", '
    '"Tarjeta de circulación", "Certificado"), o null.\n'
    '  "fecha_vigencia": fecha de INICIO de vigencia en formato YYYY-MM-DD, o null.\n'
    '  "fecha_vencimiento": fecha de VENCIMIENTO/expiración en formato YYYY-MM-DD, o null.\n'
    '  "emisor": aseguradora o entidad que emite el documento, o null.\n'
    '  "confianza": número entre 0 y 1.\n'
    '  "notas": string breve en español con cualquier observación.\n'
    "No inventes datos que no aparezcan: usa null. Si una fecha viene en otro "
    "formato (ej. DD/MM/AAAA), conviértela a YYYY-MM-DD. La fecha_vencimiento es la "
    "más importante: es hasta cuándo es válido el documento."
)

_PROMPT = (
    "Extrae los datos de vencimiento de este documento y responde con el JSON "
    "indicado. Prioriza identificar la fecha de vencimiento (hasta cuándo es válido)."
)


@lru_cache
def _client() -> anthropic.Anthropic:
    s = get_settings()
    return anthropic.Anthropic(
        api_key=s.anthropic_api_key,
        timeout=s.anthropic_timeout_s,
        max_retries=s.anthropic_max_retries,
    )


def _extract_json(text: str) -> dict:
    t = text.strip()
    if t.startswith("```"):
        t = t.split("```", 2)[1] if "```" in t[3:] else t[3:]
        t = t.removeprefix("json").strip()
    start, end = t.find("{"), t.rfind("}")
    if start == -1 or end == -1:
        raise ValueError(f"Respuesta sin JSON: {text[:200]}")
    return json.loads(t[start : end + 1])


def _source_block(req: VencimientoExtraerRequest) -> dict:
    if req.pdf_base64:
        return {
            "type": "document",
            "source": {
                "type": "base64",
                "media_type": "application/pdf",
                "data": req.pdf_base64,
            },
        }
    return {
        "type": "image",
        "source": {
            "type": "base64",
            "media_type": req.media_type,
            "data": req.image_base64,
        },
    }


def extraer_vencimiento(req: VencimientoExtraerRequest) -> VencimientoExtraerResponse:
    s = get_settings()
    resp = _client().messages.create(
        model=s.anthropic_model,
        max_tokens=1024,
        system=[{"type": "text", "text": _SYSTEM, "cache_control": {"type": "ephemeral"}}],
        messages=[
            {
                "role": "user",
                "content": [_source_block(req), {"type": "text", "text": _PROMPT}],
            }
        ],
    )
    text = next((b.text for b in resp.content if b.type == "text"), "")
    data = _extract_json(text)

    return VencimientoExtraerResponse(
        matricula=str(data["matricula"]) if data.get("matricula") else None,
        tipo_documento=str(data["tipo_documento"]) if data.get("tipo_documento") else None,
        fecha_vigencia=str(data["fecha_vigencia"]) if data.get("fecha_vigencia") else None,
        fecha_vencimiento=str(data["fecha_vencimiento"]) if data.get("fecha_vencimiento") else None,
        emisor=str(data["emisor"]) if data.get("emisor") else None,
        confianza=float(data.get("confianza", 0.0)),
        notas=str(data.get("notas", "")),
        modelo=s.anthropic_model,
    )
