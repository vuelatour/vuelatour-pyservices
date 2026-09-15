import json
from datetime import date
from functools import lru_cache

import anthropic

from app.config import get_settings
from app.schemas.vencimiento import VencimientoExtraerRequest, VencimientoExtraerResponse
from app.services._dominio import sistema_con_dominio
from app.services.ia_usage import uso_ia_de
from app.services.validaciones_ia import (
    confianza_calibrada,
    confianza_de,
    fecha_iso,
    limpiar_advertencias,
    normalizar_texto,
)

_SYSTEM = (
    "Eres un asistente documental para una empresa de aviación. A partir de un "
    "documento de vencimiento renovado (póliza de seguro, certificado, permiso o "
    "tarjeta de circulación de una aeronave o vehículo) extraes sus datos clave. "
    "Devuelves SOLO un objeto JSON, sin texto adicional ni ```fences```, con las "
    "claves exactas:\n"
    '  "matriculas": arreglo con TODAS las matrículas/placas que cubre el '
    "documento. Una póliza de FLOTA ampara varias aeronaves: devuélvelas "
    "todas, en el orden en que aparecen. Si no cubre ninguna identificable, "
    "arreglo vacío.\n"
    '  "matricula": la PRINCIPAL de esas matrículas (la primera), o null.\n'
    '  "tipo_documento": tipo del documento. Si junto al documento te doy un '
    "CATÁLOGO de tipos, usa el texto EXACTO del que corresponda; si ninguno "
    "encaja, escribe el que dice el documento y anótalo en advertencias.\n"
    '  "numero_poliza": número de póliza / folio / certificado del documento, '
    "o null. Es lo que evita registrar dos veces la misma renovación.\n"
    '  "fecha_vigencia": fecha de INICIO de vigencia en formato YYYY-MM-DD, o null.\n'
    '  "fecha_vencimiento": fecha de VENCIMIENTO/expiración en formato YYYY-MM-DD, o null.\n'
    '  "emisor": aseguradora o entidad que emite el documento (AFAC, la '
    "aseguradora, el gobierno del estado), o null.\n"
    '  "confianza": número entre 0 y 1.\n'
    '  "notas": string breve en español con cualquier observación.\n'
    '  "advertencias": arreglo de strings con lo que no puedas confirmar '
    "(fechas ilegibles, varias vigencias, matrícula dudosa). Vacío si todo "
    "está claro.\n"
    "No inventes datos que no aparezcan: usa null. Las fechas mexicanas vienen "
    "DD/MM/AAAA (07/09/2026 = 7 de septiembre): conviértelas a YYYY-MM-DD. La "
    "fecha_vencimiento es la más importante: es hasta cuándo es válido el "
    "documento, y de ella salen las alertas de la operación."
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
    prompt = _PROMPT
    contexto: list[str] = []
    flota = [str(m).strip() for m in (req.matriculas_flota or []) if str(m).strip()]
    if flota:
        contexto.append("MATRÍCULAS/PLACAS registradas: " + ", ".join(flota) + ".")
    tipos = [str(t).strip() for t in (req.tipos_documento or []) if str(t).strip()]
    if tipos:
        contexto.append("CATÁLOGO de tipos de documento: " + "; ".join(tipos) + ".")
    if contexto:
        prompt = prompt + "\n" + " ".join(contexto)
    resp = _client().messages.create(
        model=s.anthropic_model,
        max_tokens=1024,
        system=sistema_con_dominio(_SYSTEM),
        messages=[
            {
                "role": "user",
                "content": [_source_block(req), {"type": "text", "text": prompt}],
            }
        ],
    )
    uso = uso_ia_de(resp)
    text = next((b.text for b in resp.content if b.type == "text"), "")
    data = _extract_json(text)

    avisos: list[str | None] = [
        str(x).strip()[:200] for x in (data.get("advertencias") or []) if str(x).strip()
    ]

    matriculas = [
        str(m).strip().upper()
        for m in (data.get("matriculas") or [])
        if isinstance(m, str) and m.strip()
    ]
    principal = str(data["matricula"]).strip().upper() if data.get("matricula") else None
    if principal and principal not in matriculas:
        matriculas.insert(0, principal)
    if not principal and matriculas:
        principal = matriculas[0]
    if len(matriculas) > 1:
        avisos.append(
            "El documento cubre varias matrículas ("
            + ", ".join(matriculas[:6])
            + "): regístralo en cada una."
        )
    conocidas = {normalizar_texto(m).replace("-", "") for m in flota}
    if principal and conocidas and normalizar_texto(principal).replace("-", "") not in conocidas:
        avisos.append(
            f"La matrícula/placa {principal} no está registrada en el sistema: verifícala."
        )

    tipo = str(data["tipo_documento"]).strip() if data.get("tipo_documento") else None
    if tipo and tipos and normalizar_texto(tipo) not in {normalizar_texto(t) for t in tipos}:
        avisos.append(f"El tipo «{tipo}» no está en el catálogo: elígelo a mano.")

    # Las dos fechas son el dato PRINCIPAL: de ellas salen las alertas.
    vigencia = fecha_iso(data.get("fecha_vigencia"))
    vencimiento = fecha_iso(data.get("fecha_vencimiento"))
    sospechoso = False
    if data.get("fecha_vencimiento") and vencimiento is None:
        avisos.append(
            f"No se entendió la fecha de vencimiento «{data.get('fecha_vencimiento')}»."
        )
        sospechoso = True
    if vigencia and vencimiento and vencimiento <= vigencia:
        avisos.append(
            f"El vencimiento ({vencimiento}) no es posterior al inicio de vigencia "
            f"({vigencia}): revisa si se invirtieron."
        )
        sospechoso = True
    if vencimiento:
        anio = vencimiento.year
        hoy = date.today().year
        if not hoy - 5 <= anio <= hoy + 10:
            avisos.append(f"El año de vencimiento ({anio}) no es plausible: revísalo.")
            sospechoso = True

    return VencimientoExtraerResponse(
        matricula=principal,
        matriculas=matriculas,
        tipo_documento=tipo,
        numero_poliza=str(data["numero_poliza"]) if data.get("numero_poliza") else None,
        fecha_vigencia=vigencia.isoformat() if vigencia else None,
        fecha_vencimiento=vencimiento.isoformat() if vencimiento else None,
        emisor=str(data["emisor"]) if data.get("emisor") else None,
        confianza=confianza_calibrada(confianza_de(data.get("confianza")), sospechoso),
        notas=str(data.get("notas", "")),
        advertencias=limpiar_advertencias(avisos),
        modelo=s.anthropic_model,
        uso_ia=uso,
    )
