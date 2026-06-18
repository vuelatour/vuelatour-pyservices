import json
from functools import lru_cache

import anthropic

from app.config import get_settings
from app.schemas.vision import (
    CombustibleTicketResponse,
    GastoTicketRequest,
    GastoTicketResponse,
    TacometroRequest,
    TacometroResponse,
)

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


def _image_block(req: TacometroRequest | GastoTicketRequest) -> dict:
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


_TICKET_SYSTEM = (
    "Eres un asistente de captura de gastos para una empresa de aviación. A partir "
    "de la foto de un ticket o recibo, extraes los datos del gasto. Devuelves SOLO "
    "un objeto JSON, sin texto adicional ni ```fences```, con las claves exactas:\n"
    '  "monto": número con el TOTAL pagado (no subtotales), o null si ilegible.\n'
    '  "moneda": "MXN" o "USD" según el ticket (default "MXN" en México).\n'
    '  "fecha": fecha del ticket en formato YYYY-MM-DD, o null.\n'
    '  "proveedor": nombre del comercio/proveedor, o null.\n'
    '  "concepto": descripción breve de lo comprado, o null.\n'
    '  "categoria_sugerida": una de GAS, ATERRIZAJE, TUAS, FBO, COMIDA, HOTEL, '
    "TAXI, REFACCION, PERMISO, FIJO, OTRO (la más probable), o null.\n"
    '  "medio_pago": "EFECTIVO", "TARJETA_CORP" o "TRANSFERENCIA" segun el ticket '
    "(DEBITO/CREDITO/VISA/MASTERCARD/TARJETA = TARJETA_CORP; EFECTIVO/CASH = "
    "EFECTIVO; SPEI/TRANSFERENCIA = TRANSFERENCIA), o null.\n"
    '  "tarjeta_terminacion": ultimos 4 digitos de la tarjeta si aparecen, como string de 4 digitos, o null.\n'
    '  "confianza": número entre 0 y 1.\n'
    '  "legible": true/false según si el ticket se distingue.\n'
    '  "notas": string breve en español con cualquier observación.\n'
    "No inventes datos que no aparezcan: usa null. GAS es combustible/turbosina; "
    "FBO es servicio de aeropuerto; TUAS es tarifa de uso de aeropuerto."
)

_TICKET_PROMPT = (
    "Extrae los datos de gasto de este ticket y responde con el JSON indicado. "
    "Prioriza el TOTAL final, no subtotales ni impuestos por separado."
)


def leer_ticket_gasto(req: GastoTicketRequest) -> GastoTicketResponse:
    s = get_settings()
    resp = _client().messages.create(
        model=s.anthropic_model,
        max_tokens=512,
        system=[{"type": "text", "text": _TICKET_SYSTEM, "cache_control": {"type": "ephemeral"}}],
        messages=[
            {
                "role": "user",
                "content": [_image_block(req), {"type": "text", "text": _TICKET_PROMPT}],
            }
        ],
    )
    text = next((b.text for b in resp.content if b.type == "text"), "")
    data = _extract_json(text)

    monto = data.get("monto")
    moneda = data.get("moneda")
    categoria = data.get("categoria_sugerida")
    valid_cats = {
        "GAS", "ATERRIZAJE", "TUAS", "FBO", "COMIDA", "HOTEL",
        "TAXI", "REFACCION", "PERMISO", "FIJO", "OTRO",
    }
    return GastoTicketResponse(
        monto=float(monto) if isinstance(monto, (int, float)) else None,
        moneda=moneda if moneda in ("MXN", "USD") else None,
        fecha=str(data["fecha"]) if data.get("fecha") else None,
        proveedor=str(data["proveedor"]) if data.get("proveedor") else None,
        concepto=str(data["concepto"]) if data.get("concepto") else None,
        categoria_sugerida=categoria if categoria in valid_cats else None,
        medio_pago=(
            data.get("medio_pago")
            if data.get("medio_pago") in ("EFECTIVO", "TARJETA_CORP", "TRANSFERENCIA")
            else None
        ),
        tarjeta_terminacion=(
            str(data["tarjeta_terminacion"]).strip()[-4:]
            if data.get("tarjeta_terminacion") and str(data["tarjeta_terminacion"]).strip()
            else None
        ),
        confianza=float(data.get("confianza", 0.0)),
        legible=bool(data.get("legible", monto is not None)),
        notas=str(data.get("notas", "")),
        modelo=s.anthropic_model,
    )


_COMBUSTIBLE_SYSTEM = (
    "Eres un asistente de captura de cargas de combustible de aviación (turbosina "
    "Jet A o avgas 100LL). A partir de la foto del ticket de combustible extraes los "
    "datos. Devuelves SOLO un objeto JSON, sin texto adicional ni ```fences```, con "
    "las claves exactas:\n"
    '  "litros": litros cargados (si viene en galones, conviértelo: 1 gal = 3.78541 L), o null.\n'
    '  "precio_litro": precio por litro, o null.\n'
    '  "total": total pagado, o null.\n'
    '  "moneda": "MXN" o "USD", o null.\n'
    '  "aeropuerto": código IATA/ICAO o nombre del aeropuerto/FBO, o null.\n'
    '  "tipo_combustible": "TURBOSINA" o "AVGAS" según el ticket, o null.\n'
    '  "fecha": fecha YYYY-MM-DD, o null.\n'
    '  "hora": hora de la carga en formato HH:MM de 24 horas, o null.\n'
    '  "proveedor": nombre del proveedor/FBO, o null.\n'
    '  "tarjeta_terminacion": ultimos 4 digitos de la tarjeta de pago si aparecen '
    'en el ticket (p. ej. "**** 1234" o "TARJETA ...1234"), como string de 4 digitos, o null.\n'
    '  "medio_pago": "EFECTIVO", "TARJETA_CORP" o "TRANSFERENCIA" segun el ticket, o null.\n'
    '  "confianza": número entre 0 y 1.\n'
    '  "legible": true/false.\n'
    '  "notas": string breve en español.\n'
    "No inventes datos: usa null. Si el ticket indica galones, convierte a litros."
)

_COMBUSTIBLE_PROMPT = (
    "Extrae los datos de esta carga de combustible y responde con el JSON indicado. "
    "Si la cantidad viene en galones (GAL), conviértela a litros."
)


def leer_ticket_combustible(req: GastoTicketRequest) -> CombustibleTicketResponse:
    s = get_settings()
    resp = _client().messages.create(
        model=s.anthropic_model,
        max_tokens=512,
        system=[{"type": "text", "text": _COMBUSTIBLE_SYSTEM, "cache_control": {"type": "ephemeral"}}],
        messages=[
            {
                "role": "user",
                "content": [_image_block(req), {"type": "text", "text": _COMBUSTIBLE_PROMPT}],
            }
        ],
    )
    text = next((b.text for b in resp.content if b.type == "text"), "")
    data = _extract_json(text)

    def _num(v: object) -> float | None:
        return float(v) if isinstance(v, (int, float)) else None

    moneda = data.get("moneda")
    tipo = data.get("tipo_combustible")
    return CombustibleTicketResponse(
        litros=_num(data.get("litros")),
        precio_litro=_num(data.get("precio_litro")),
        total=_num(data.get("total")),
        moneda=moneda if moneda in ("MXN", "USD") else None,
        aeropuerto=str(data["aeropuerto"]) if data.get("aeropuerto") else None,
        tipo_combustible=tipo if tipo in ("TURBOSINA", "AVGAS") else None,
        fecha=str(data["fecha"]) if data.get("fecha") else None,
        hora=str(data["hora"]) if data.get("hora") else None,
        proveedor=str(data["proveedor"]) if data.get("proveedor") else None,
        tarjeta_terminacion=(
            str(data["tarjeta_terminacion"]).strip()[-4:]
            if data.get("tarjeta_terminacion") and str(data["tarjeta_terminacion"]).strip()
            else None
        ),
        medio_pago=(
            data.get("medio_pago")
            if data.get("medio_pago") in ("EFECTIVO", "TARJETA_CORP", "TRANSFERENCIA")
            else None
        ),
        confianza=float(data.get("confianza", 0.0)),
        legible=bool(data.get("legible", data.get("total") is not None)),
        notas=str(data.get("notas", "")),
        modelo=s.anthropic_model,
    )
