"""Match IA gasto→vuelo: Claude elige el vuelo más probable entre candidatos.

Mismo patrón que la sugerencia de conciliación: los candidatos vienen del API
(deterministas) y el modelo SOLO puede escoger uno de ellos o ninguno.
"""

import json

from app.config import get_settings
from app.schemas.gastos import GastoVueloSugerirRequest, GastoVueloSugerirResponse
from app.services.estado_cuenta import _client, _extract_json

_SYSTEM = (
    "Eres un asistente de una operadora de vuelos chárter en Cancún. Un piloto "
    "capturó un gasto sin indicar a qué vuelo pertenece. Te doy el gasto y la "
    "lista de vuelos CANDIDATOS (vuelos de ese piloto en fechas cercanas). "
    "Responde SOLO un objeto JSON, sin texto adicional, con las claves:\n"
    '  "vuelo_id_sugerido": id del vuelo más probable, o null si ninguno encaja '
    "razonablemente.\n"
    '  "confianza": número entre 0 y 1.\n'
    '  "razon": explicación breve en español (ej. "voló ese día a CZM y el gasto '
    'es de un FBO de Cozumel").\n'
    "Considera: la fecha del gasto vs la fecha del vuelo (mismo día pesa mucho; "
    "hoteles/comidas pueden ser del día anterior o siguiente si hubo pernocta); "
    "el lugar/proveedor del gasto vs los aeropuertos de la ruta; la categoría "
    "(aterrizaje/FBO ocurren en aeropuertos de la ruta; un supermercado no ata a "
    "un vuelo). Usa SOLO ids de la lista. Si nada encaja, vuelo_id_sugerido=null "
    "con confianza 0."
)


def sugerir_vuelo_para_gasto(
    req: GastoVueloSugerirRequest,
) -> GastoVueloSugerirResponse:
    s = get_settings()

    if not req.candidatos:
        return GastoVueloSugerirResponse(
            vuelo_id_sugerido=None,
            confianza=0.0,
            razon="Sin vuelos del piloto en fechas cercanas.",
            modelo=s.anthropic_model,
        )

    ids_validos = {c.vuelo_id for c in req.candidatos}
    payload = json.dumps(
        {
            "gasto": req.gasto.model_dump(),
            "candidatos": [c.model_dump() for c in req.candidatos],
        },
        ensure_ascii=False,
    )
    resp = _client().messages.create(
        model=s.anthropic_model,
        max_tokens=512,
        system=[{"type": "text", "text": _SYSTEM, "cache_control": {"type": "ephemeral"}}],
        messages=[
            {
                "role": "user",
                "content": (
                    "Gasto sin asignar y vuelos candidatos (JSON):\n"
                    f"{payload}\n\n"
                    "Elige el vuelo más probable y responde con el JSON indicado."
                ),
            }
        ],
    )
    text = next((b.text for b in resp.content if b.type == "text"), "")
    data = _extract_json(text)

    sugerido = data.get("vuelo_id_sugerido")
    sugerido = str(sugerido) if sugerido is not None else None
    # No confíes en ids inventados por el modelo: solo candidatos reales.
    if sugerido not in ids_validos:
        sugerido = None

    return GastoVueloSugerirResponse(
        vuelo_id_sugerido=sugerido,
        confianza=float(data.get("confianza", 0.0)) if sugerido else 0.0,
        razon=str(data.get("razon", "")),
        modelo=s.anthropic_model,
    )
