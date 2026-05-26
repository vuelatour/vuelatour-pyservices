import json
from functools import lru_cache

import anthropic

from app.config import get_settings
from app.schemas.compras import CompraExtraerRequest, CompraExtraerResponse, CompraLinea

_SYSTEM = (
    "Eres un asistente de compras para una empresa de aviación. A partir de una "
    "factura u orden de compra en PDF (ej. Aircraft Spruce, en inglés), extraes "
    "las líneas de producto. Devuelves SOLO un objeto JSON, sin texto adicional "
    "ni ```fences```, con las claves exactas:\n"
    '  "proveedor": nombre del proveedor, o null.\n'
    '  "fecha": fecha de la orden YYYY-MM-DD, o null.\n'
    '  "moneda": código de moneda (normalmente "USD").\n'
    '  "lineas": arreglo de objetos con "nombre" (descripción del producto), '
    '"numero_parte" (part number o null), "cantidad" (número), '
    '"precio_unitario_usd" (número o null), "total_usd" (número o null).\n'
    '  "subtotal_usd", "shipping_usd", "impuestos_usd", "total_usd": números o null.\n'
    '  "confianza": número entre 0 y 1.\n'
    '  "notas": string breve en español.\n'
    "Incluye solo líneas de producto reales (no encabezados ni totales). No "
    "inventes part numbers: usa null si no aparece."
)

_PROMPT = (
    "Extrae las líneas de producto de esta factura/orden de compra y responde con "
    "el JSON indicado."
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


def _num(v) -> float | None:
    return float(v) if isinstance(v, (int, float)) else None


def extraer_compra(req: CompraExtraerRequest) -> CompraExtraerResponse:
    s = get_settings()
    resp = _client().messages.create(
        model=s.anthropic_model,
        max_tokens=2048,
        system=[{"type": "text", "text": _SYSTEM, "cache_control": {"type": "ephemeral"}}],
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "document",
                        "source": {
                            "type": "base64",
                            "media_type": "application/pdf",
                            "data": req.pdf_base64,
                        },
                    },
                    {"type": "text", "text": _PROMPT},
                ],
            }
        ],
    )
    text = next((b.text for b in resp.content if b.type == "text"), "")
    data = _extract_json(text)

    lineas: list[CompraLinea] = []
    for raw in data.get("lineas", []):
        if not isinstance(raw, dict):
            continue
        nombre = raw.get("nombre")
        if not nombre:
            continue
        lineas.append(
            CompraLinea(
                nombre=str(nombre),
                numero_parte=str(raw["numero_parte"]) if raw.get("numero_parte") else None,
                cantidad=_num(raw.get("cantidad")) or 1,
                precio_unitario_usd=_num(raw.get("precio_unitario_usd")),
                total_usd=_num(raw.get("total_usd")),
            )
        )

    return CompraExtraerResponse(
        proveedor=str(data["proveedor"]) if data.get("proveedor") else None,
        fecha=str(data["fecha"]) if data.get("fecha") else None,
        moneda=str(data.get("moneda", "USD")),
        lineas=lineas,
        subtotal_usd=_num(data.get("subtotal_usd")),
        shipping_usd=_num(data.get("shipping_usd")),
        impuestos_usd=_num(data.get("impuestos_usd")),
        total_usd=_num(data.get("total_usd")),
        confianza=float(data.get("confianza", 0.0)),
        notas=str(data.get("notas", "")),
        modelo=s.anthropic_model,
    )
