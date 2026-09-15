import json
from functools import lru_cache

import anthropic

from app.config import get_settings
from app.schemas.compras import CompraExtraerRequest, CompraExtraerResponse, CompraLinea
from app.services._dominio import sistema_con_dominio
from app.services.ia_usage import uso_ia_de
from app.services.validaciones_ia import (
    confianza_calibrada,
    confianza_de,
    cuadra,
    limpiar_advertencias,
)

_SYSTEM = (
    "Eres un asistente de compras para una empresa de aviación. A partir de una "
    "factura u orden de compra en PDF (ej. Aircraft Spruce, en inglés), extraes "
    "las líneas de producto. Devuelves SOLO un objeto JSON, sin texto adicional "
    "ni ```fences```, con las claves exactas:\n"
    '  "proveedor": nombre del proveedor, o null.\n'
    '  "fecha": fecha de la orden YYYY-MM-DD, o null.\n'
    '  "numero_orden": el número de ORDEN/INVOICE/PO del proveedor tal como lo '
    'imprime (ej. "Invoice 1234567", "Order #A-9981" -> "1234567", "A-9981"), '
    "o null. Es la llave con la que el sistema evita capturar dos veces la "
    "misma compra: no lo inventes.\n"
    '  "tracking": número de guía/tracking del embarque si aparece, o null.\n'
    '  "moneda": código de moneda (normalmente "USD").\n'
    '  "lineas": arreglo de objetos con "nombre" (descripción del producto), '
    '"numero_parte" (part number o null), "cantidad" (número), '
    '"precio_unitario_usd" (número o null), "total_usd" (número o null).\n'
    '  "subtotal_usd": suma de las líneas de producto.\n'
    '  "shipping_usd": envío + handling JUNTOS (freight, shipping, handling, '
    "S&H): el sistema los prorratea igual sobre el costo de las piezas.\n"
    '  "impuestos_usd": taxes + duties/customs/aranceles JUNTOS.\n'
    '  "total_usd": el gran total de la factura.\n'
    '  "confianza": número entre 0 y 1.\n'
    '  "notas": string breve en español.\n'
    '  "advertencias": arreglo de strings con lo que no cuadre (una línea sin '
    "precio, backordered, un total que no suma). Vacío si todo está claro.\n"
    "CONVENCIONES DE AVIACIÓN: los part numbers llevan guiones y letras "
    "(MS20470AD4-4, AN960-10L) — cópialos EXACTOS. Unidades: ea/each = pieza, "
    "qt = cuarto de galón, gal, ft, in. Una línea marcada BACKORDERED o B/O no "
    "se embarcó: inclúyela con su cantidad y dilo en advertencias. La moneda "
    "es USD salvo que el documento diga otra cosa.\n"
    "ARITMÉTICA: por línea, cantidad × precio_unitario = total de la línea; la "
    "suma de las líneas = subtotal; subtotal + shipping + impuestos = total. Si "
    "algo de eso no cuadra, NO ajustes los números a mano: transcribe lo que "
    "ves, baja la confianza y dilo en advertencias.\n"
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
        system=sistema_con_dominio(_SYSTEM),
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
    uso = uso_ia_de(resp)
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

    subtotal = _num(data.get("subtotal_usd"))
    shipping = _num(data.get("shipping_usd"))
    impuestos = _num(data.get("impuestos_usd"))
    total = _num(data.get("total_usd"))

    # --- Validaciones deterministas (15-sep-2026): el costo prorrateado de
    # cada refacción sale de estos números; si no cuadran, no pueden salir
    # como seguros.
    avisos: list[str | None] = [
        str(x).strip()[:200] for x in (data.get("advertencias") or []) if str(x).strip()
    ]
    sospechoso = False

    suma_lineas = round(sum(ln.total_usd or 0 for ln in lineas), 2)
    if lineas and subtotal is not None and not cuadra(suma_lineas, subtotal):
        avisos.append(
            f"Las líneas suman {suma_lineas:,.2f} USD y el subtotal dice {subtotal:,.2f}: "
            "revisa si falta alguna partida."
        )
        sospechoso = True

    if total is not None and subtotal is not None:
        armado = round(subtotal + (shipping or 0) + (impuestos or 0), 2)
        if not cuadra(armado, total):
            avisos.append(
                f"Subtotal + envío + impuestos = {armado:,.2f} USD y el total dice "
                f"{total:,.2f}: revisa los cargos del pie."
            )
            sospechoso = True

    for ln in lineas:
        if ln.cantidad <= 0:
            avisos.append(f"La línea «{ln.nombre[:40]}» viene con cantidad {ln.cantidad}.")
            sospechoso = True
        elif ln.precio_unitario_usd is not None and ln.total_usd is not None:
            esperado = round(ln.cantidad * ln.precio_unitario_usd, 2)
            if not cuadra(esperado, ln.total_usd, 0.05):
                avisos.append(
                    f"En «{ln.nombre[:40]}», {ln.cantidad} × {ln.precio_unitario_usd:,.2f} = "
                    f"{esperado:,.2f} y la línea dice {ln.total_usd:,.2f}."
                )
                sospechoso = True

    return CompraExtraerResponse(
        proveedor=str(data["proveedor"]) if data.get("proveedor") else None,
        fecha=str(data["fecha"]) if data.get("fecha") else None,
        numero_orden=str(data["numero_orden"]) if data.get("numero_orden") else None,
        tracking=str(data["tracking"]) if data.get("tracking") else None,
        moneda=str(data.get("moneda", "USD")),
        lineas=lineas,
        subtotal_usd=subtotal,
        shipping_usd=shipping,
        impuestos_usd=impuestos,
        total_usd=total,
        confianza=confianza_calibrada(confianza_de(data.get("confianza")), sospechoso),
        notas=str(data.get("notas", "")),
        advertencias=limpiar_advertencias(avisos),
        modelo=s.anthropic_model,
        uso_ia=uso,
    )
