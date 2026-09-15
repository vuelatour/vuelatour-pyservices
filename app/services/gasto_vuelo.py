"""Match IA gasto→vuelo: Claude elige el vuelo más probable entre candidatos.

Mismo patrón que la sugerencia de conciliación: los candidatos vienen del API
(deterministas) y el modelo SOLO puede escoger uno de ellos o ninguno. Lo que
se puede comprobar con fechas y aeropuertos se comprueba aquí, en Python, y
acota la confianza que el modelo puede reclamar.
"""

import json
import re

from app.config import get_settings
from app.schemas.gastos import (
    GastoParaMatch,
    GastoVueloSugerirRequest,
    GastoVueloSugerirResponse,
    VueloCandidato,
)
from app.services._dominio import AEROPUERTOS, sistema_con_dominio
from app.services.estado_cuenta import _client, _extract_json
from app.services.ia_usage import uso_ia_de
from app.services.validaciones_ia import (
    confianza_de,
    dias_entre_cancun,
    limpiar_advertencias,
    normalizar_texto,
)

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
    '  "evidencias": arreglo de strings con los HECHOS que citas, uno por hecho '
    '(ej. "mismo día 2026-09-07", "la ruta toca CZM y el gasto es de Cozumel", '
    '"es el único vuelo del piloto ese día"). Si no puedes citar ningún hecho, '
    "deja el arreglo vacío y devuelve null.\n"
    '  "motivo_sin_match": cuando el sugerido es null, UNA frase con el motivo '
    "(ej. «el gasto es de un aeropuerto que ningún candidato tocó»). null si "
    "sí hay sugerencia.\n"
    "CRITERIOS, del más fuerte al más débil:\n"
    "  1. AEROPUERTO: un gasto de aeropuerto (aterrizaje, plataforma, TUA, FBO, "
    "combustible) SOLO puede ser de un vuelo que TOCÓ ese aeropuerto. Si el "
    "lugar/proveedor del gasto nombra una ciudad o código IATA que no está en "
    "la ruta del candidato, ese candidato queda descartado.\n"
    "  2. FECHA (todas las fechas, del gasto y del vuelo, están en hora de "
    "Cancún UTC−5; compáralas como días de calendario de Cancún, no en UTC): "
    "el mismo día pesa mucho. Excepción de PERNOCTA: el HOTEL de la noche del "
    "día de llegada y el DESAYUNO/taxi de la mañana siguiente pertenecen al "
    "vuelo que se quedó a dormir, aunque la fecha sea del día siguiente; una "
    "comida de la noche anterior a un vuelo de madrugada, igual.\n"
    "  3. CATEGORÍA: aterrizaje/FBO/TUAS/GAS ocurren en un aeropuerto de la "
    "ruta; comida/taxi/hotel son viáticos de la tripulación y pueden estar "
    "lejos del aeropuerto; un súper o una refacción rara vez atan a un vuelo.\n"
    "  4. Si un candidato tiene estado CANCELADO, elígelo solo cuando el gasto "
    "claramente sea de ese vuelo (p. ej. voló a recoger y lo cancelaron); ante "
    "duda prefiere el vuelo NO cancelado del mismo día.\n"
    "Usa SOLO ids de la lista, nunca inventes uno. Si dos candidatos encajan "
    "igual de bien, devuelve null con motivo «hay dos vuelos posibles»: es "
    "mejor que el operador lo decida a colgarle el gasto al vuelo equivocado, "
    "porque ese gasto cambia la utilidad que se reparte."
)


def _codigos_ruta(ruta: str | None) -> set[str]:
    """Códigos IATA de una ruta ('CUN → CZM → CUN' → {CUN, CZM})."""
    if not ruta:
        return set()
    return {t for t in re.findall(r"\b[A-Z]{3}\b", normalizar_texto(ruta))}


def _lugares_del_gasto(gasto: GastoParaMatch) -> set[str]:
    """Códigos IATA que el gasto menciona (por código o por nombre de ciudad)."""
    texto = normalizar_texto(" ".join(x for x in (gasto.lugar, gasto.notas) if x))
    if not texto:
        return set()
    encontrados = {t for t in re.findall(r"\b[A-Z]{3}\b", texto) if t in AEROPUERTOS}
    for iata, nombre in AEROPUERTOS.items():
        if normalizar_texto(nombre) and normalizar_texto(nombre) in texto:
            encontrados.add(iata)
    return encontrados


def _evidencias_deterministas(
    gasto: GastoParaMatch, vuelo: VueloCandidato, candidatos: list[VueloCandidato]
) -> tuple[list[str], float]:
    """Hechos comprobables + tope de confianza que permiten."""
    ev: list[str] = []
    topes: list[float] = []

    dias = dias_entre_cancun(vuelo.fecha_vuelo, gasto.fecha)
    if dias is not None:
        if dias == 0:
            ev.append("El gasto es del mismo día del vuelo.")
        elif abs(dias) == 1:
            ev.append(
                f"El gasto es del día {'siguiente' if dias > 0 else 'anterior'} al vuelo "
                "(pernocta)."
            )
            topes.append(0.85)
        else:
            ev.append(f"El gasto está a {abs(dias)} días del vuelo.")
            topes.append(0.6)

    ruta = _codigos_ruta(vuelo.ruta)
    lugares = _lugares_del_gasto(gasto)
    if lugares and ruta:
        comunes = lugares & ruta
        if comunes:
            ev.append(f"La ruta del vuelo toca {', '.join(sorted(comunes))}, como el gasto.")
        else:
            ev.append(
                f"El gasto es de {', '.join(sorted(lugares))} y la ruta del vuelo "
                f"({', '.join(sorted(ruta))}) no lo toca."
            )
            topes.append(0.4)

    if (vuelo.estado or "").upper() == "CANCELADO":
        ev.append("El vuelo está CANCELADO: confirma que el gasto sí ocurrió.")
        topes.append(0.6)

    mismo_dia = [
        c
        for c in candidatos
        if c.fecha_vuelo and gasto.fecha and dias_entre_cancun(c.fecha_vuelo, gasto.fecha) == 0
    ]
    if len(mismo_dia) == 1 and mismo_dia[0].vuelo_id == vuelo.vuelo_id:
        ev.append("Es el único vuelo del piloto ese día.")

    return ev, min(topes) if topes else 0.95


def sugerir_vuelo_para_gasto(
    req: GastoVueloSugerirRequest,
) -> GastoVueloSugerirResponse:
    s = get_settings()

    if not req.candidatos:
        return GastoVueloSugerirResponse(
            vuelo_id_sugerido=None,
            confianza=0.0,
            razon="Sin vuelos del piloto en fechas cercanas.",
            motivo_sin_match="El piloto que capturó el gasto no tiene vuelos en esas fechas.",
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
        max_tokens=800,
        system=sistema_con_dominio(_SYSTEM),
        messages=[
            {
                "role": "user",
                "content": (
                    "Gasto sin asignar y vuelos candidatos (JSON):\n"
                    f"{payload}\n\n"
                    "Elige el vuelo más probable y responde con el JSON indicado. "
                    "Cita hechos en «evidencias»; si no tienes ninguno, devuelve null."
                ),
            }
        ],
    )
    uso = uso_ia_de(resp)
    text = next((b.text for b in resp.content if b.type == "text"), "")
    data = _extract_json(text)

    sugerido = data.get("vuelo_id_sugerido")
    sugerido = str(sugerido) if sugerido is not None else None
    # No confíes en ids inventados por el modelo: solo candidatos reales.
    if sugerido not in ids_validos:
        sugerido = None
    elegido = next((c for c in req.candidatos if c.vuelo_id == sugerido), None)

    evidencias: list[str] = []
    confianza = 0.0
    if elegido is not None:
        duras, tope = _evidencias_deterministas(req.gasto, elegido, req.candidatos)
        evidencias = limpiar_advertencias(
            [*duras, *[str(x).strip()[:200] for x in (data.get("evidencias") or [])]]
        )
        confianza = min(confianza_de(data.get("confianza")), tope)

    razon = str(data.get("razon", ""))[:300]
    if elegido is not None and not razon:
        razon = evidencias[0] if evidencias else "Coincidencia propuesta por la IA."

    motivo = data.get("motivo_sin_match")
    motivo = str(motivo)[:300] if motivo else None
    if elegido is None and not motivo:
        motivo = "Ningún vuelo del piloto encaja con la fecha y el lugar del gasto."

    return GastoVueloSugerirResponse(
        vuelo_id_sugerido=sugerido,
        confianza=confianza,
        razon=razon,
        modelo=s.anthropic_model,
        uso_ia=uso,
        evidencias=evidencias[:8],
        motivo_sin_match=motivo if elegido is None else None,
    )
