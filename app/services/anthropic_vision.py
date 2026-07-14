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
    "(HOBBS) de una aeronave a partir de una foto del instrumento.\n"
    "FORMATO FIJO de la lectura: 4 dígitos enteros + 1 décima, es decir NNNN.N "
    "(ej. 1555.8). El tambor de la derecha —más pequeño, normalmente recuadrado, "
    "de otro color y marcado con '1/10'— es SIEMPRE la décima, nunca un entero. "
    "Lee exactamente 4 ruedas enteras + esa décima.\n"
    "Si en la foto aparece más de un medidor (p. ej. ENGINE/ELT/otro de "
    "referencia), IGNORA los secundarios: reporta solo el HOBBS principal. No "
    "mezcles dígitos de dos medidores distintos.\n"
    "Devuelves SOLO un objeto JSON, sin texto adicional ni ```fences```, con las "
    "claves exactas:\n"
    '  "lectura": número NNNN.N con la lectura en horas, o null si no se puede leer.\n'
    '  "confianza": número entre 0 y 1.\n'
    '  "legible": true/false según si el display se distingue.\n'
    '  "notas": string breve en español (reflejo, borrosa, dígito parcial, etc.).\n'
    "Si dudas entre dos dígitos, elige el más probable y baja la confianza. "
    "No inventes dígitos que no ves: si faltan, usa null y explica en notas."
)

_USER_PROMPT = (
    "Lee el contador de horas (HOBBS/tacómetro) en esta foto y responde con el "
    "JSON indicado. Recuerda: 4 enteros + 1 décima (NNNN.N); el tambor pequeño "
    "marcado '1/10' es la décima, no un entero; ignora cualquier medidor "
    "secundario que aparezca."
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


def _excel_to_text(excel_base64: str, filename: str | None) -> str:
    """Convierte la factura en Excel/CSV a texto tabular (todas las hojas,
    truncado) para que Claude extraiga los datos como si leyera el documento."""
    import base64 as _b64
    import io

    import pandas as pd

    raw = _b64.b64decode(excel_base64)
    nombre = (filename or "").lower()
    partes: list[str] = []
    if nombre.endswith(".csv"):
        df = pd.read_csv(io.BytesIO(raw), dtype=str, keep_default_na=False)
        partes.append(df.to_csv(index=False))
    elif nombre.endswith(".xls") and not nombre.endswith(".xlsx"):
        raise ValueError("Formato .xls (Excel viejo) no soportado: guárdalo como .xlsx")
    else:
        hojas = pd.read_excel(io.BytesIO(raw), dtype=str, sheet_name=None)
        for nombre_hoja, df in list(hojas.items())[:3]:
            df = df.fillna("")
            partes.append(f"--- HOJA: {nombre_hoja} ---\n{df.to_csv(index=False)}")
    texto = "\n\n".join(partes).strip()
    if not texto:
        raise ValueError("El archivo Excel/CSV está vacío o no se pudo leer")
    # Techo defensivo: una factura no necesita más; evita reventar tokens.
    return texto[:15000]


def _gasto_source_blocks(req: GastoTicketRequest) -> list[dict]:
    """Bloques de contenido para el ticket: 1 foto, N fotos (hojas del mismo
    documento) o un PDF como bloque document (visión nativa)."""
    if req.pdf_base64:
        return [
            {
                "type": "document",
                "source": {
                    "type": "base64",
                    "media_type": "application/pdf",
                    "data": req.pdf_base64,
                },
            }
        ]
    if req.images:
        blocks: list[dict] = []
        for img in req.images:
            if img.image_base64:
                blocks.append(
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": img.media_type,
                            "data": img.image_base64,
                        },
                    }
                )
            else:
                blocks.append({"type": "image", "source": {"type": "url", "url": img.image_url}})
        return blocks
    return [_image_block(req)]


def _extract_json(text: str) -> dict:
    t = text.strip()
    if t.startswith("```"):
        t = t.split("```", 2)[1] if "```" in t[3:] else t[3:]
        t = t.removeprefix("json").strip()
    start, end = t.find("{"), t.rfind("}")
    if start == -1 or end == -1:
        raise ValueError(f"Respuesta sin JSON: {text[:300]}")
    try:
        return json.loads(t[start : end + 1])
    except json.JSONDecodeError as e:
        # Incluir el contenido en el error: sin esto el log solo decía la
        # posición del fallo y no QUÉ devolvió el modelo (indepurable).
        raise ValueError(
            f"JSON inválido ({e.msg} pos {e.pos}): …{t[max(0, e.pos - 80) : e.pos + 80]}…"
        ) from e


def leer_tacometro(req: TacometroRequest) -> TacometroResponse:
    s = get_settings()
    user_text = _USER_PROMPT
    if isinstance(req.ultimo, (int, float)) and req.ultimo > 0:
        # Ancla de magnitud: la lectura anterior acota el rango esperado y evita
        # que la IA confunda la décima con un entero (ej. 1555.8 vs 15558.1).
        user_text = (
            f"{_USER_PROMPT}\nLa lectura anterior de esta aeronave fue "
            f"{req.ultimo:.1f} hrs. La nueva debe ser parecida (sube ~1 a 3 hrs por "
            f"vuelo) y NUNCA menor; si tu lectura se aleja mucho de ese valor, "
            f"revisa que no hayas tomado la décima como entero."
        )
    resp = _client().messages.create(
        model=s.anthropic_model,
        max_tokens=800,
        system=[{"type": "text", "text": _SYSTEM, "cache_control": {"type": "ephemeral"}}],
        messages=[
            {
                "role": "user",
                "content": [_image_block(req), {"type": "text", "text": user_text}],
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
    '  "categoria_sugerida": una de GAS, OPERACIONES, COMIDA, HOTEL, '
    "TAXI, REFACCION, PERMISO, FIJO, OTRO (la más probable), o null.\n"
    '  "medio_pago": "EFECTIVO", "TARJETA_CORP" o "TRANSFERENCIA" segun el ticket '
    "(DEBITO/CREDITO/VISA/MASTERCARD/TARJETA = TARJETA_CORP; EFECTIVO/CASH = "
    "EFECTIVO; SPEI/TRANSFERENCIA = TRANSFERENCIA), o null.\n"
    '  "tarjeta_terminacion": ultimos 4 digitos de la tarjeta si aparecen, como string de 4 digitos, o null.\n'
    '  "conceptos": lista [{"concepto": string, "monto": número}] con los RENGLONES '
    "del ticket/factura SOLO si desglosa claramente varios conceptos (ej. "
    "aterrizaje, plataforma de pernocta, embarque, servicio FBO, T.U.A.); máximo 8; "
    "[] si es un solo concepto o no se distinguen. OJO CON LOS DESCUENTOS: si un "
    "renglón tiene columna Descuento (muy común en T.U.A.), el monto del renglón es "
    "el NETO = importe − descuento (ej. importe $605.18 con descuento $5.18 → monto "
    "600.00). IMPORTANTE: si la factura muestra subtotal + IVA por separado, agrega "
    'el IVA como un renglón más ({"concepto": "IVA 16%", "monto": X}) para que la '
    "SUMA de renglones (ya con descuentos aplicados) sea exactamente el TOTAL "
    "pagado; si aun así no cuadra, [].\n"
    '  "matricula": matrícula de la aeronave si aparece en el documento '
    '(observaciones/referencias; formato XA-ABC, XB-ABC o N123XX), o null.\n'
    '  "confianza": número entre 0 y 1.\n'
    '  "legible": true/false según si el ticket se distingue.\n'
    '  "notas": string breve en español con cualquier observación.\n'
    "No inventes datos que no aparezcan: usa null. GAS es combustible/turbosina; "
    "OPERACIONES agrupa aterrizaje, FBO y servicios de aeropuerto/plataforma."
)

_TICKET_PROMPT = (
    "Extrae los datos de gasto de este ticket y responde con el JSON indicado. "
    "Prioriza el TOTAL final, no subtotales ni impuestos por separado."
)


def leer_ticket_gasto(req: GastoTicketRequest) -> GastoTicketResponse:
    s = get_settings()
    if req.excel_base64:
        # Hoja de cálculo: viaja como TEXTO tabular (Claude no lee xlsx nativo).
        texto = _excel_to_text(req.excel_base64, req.excel_filename)
        blocks: list[dict] = [
            {
                "type": "text",
                "text": (
                    "CONTENIDO DE LA FACTURA (hoja de cálculo convertida a texto, "
                    f"separado por comas):\n{texto}"
                ),
            }
        ]
        prompt = (
            "El contenido anterior es la factura en formato tabular. "
        ) + _TICKET_PROMPT
    else:
        blocks = _gasto_source_blocks(req)
        prompt = _TICKET_PROMPT
        if len(blocks) > 1:
            prompt = (
                f"Las {len(blocks)} imágenes son HOJAS del MISMO documento (una sola "
                "factura de varias páginas): léelas en orden como un solo comprobante. "
            ) + _TICKET_PROMPT
    resp = _client().messages.create(
        model=s.anthropic_model,
        max_tokens=2000,
        system=[{"type": "text", "text": _TICKET_SYSTEM, "cache_control": {"type": "ephemeral"}}],
        messages=[
            {
                "role": "user",
                "content": [*blocks, {"type": "text", "text": prompt}],
            }
        ],
    )
    if resp.stop_reason == "max_tokens":
        raise ValueError("Respuesta truncada por max_tokens (subir el límite)")
    text = next((b.text for b in resp.content if b.type == "text"), "")
    data = _extract_json(text)

    monto = data.get("monto")
    moneda = data.get("moneda")
    categoria = data.get("categoria_sugerida")
    valid_cats = {
        "GAS", "OPERACIONES", "ATERRIZAJE", "TUAS", "FBO", "COMIDA", "HOTEL",
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
        conceptos=_parse_conceptos(data.get("conceptos")),
        matricula=_parse_matricula(data.get("matricula")),
        confianza=float(data.get("confianza", 0.0)),
        legible=bool(data.get("legible", monto is not None)),
        notas=str(data.get("notas", "")),
        modelo=s.anthropic_model,
    )


def _parse_matricula(raw: object) -> str | None:
    """Matrícula plausible (XA-/XB-/N…): mayúsculas, 4-7 caracteres."""
    if not isinstance(raw, str):
        return None
    m = raw.strip().upper().replace(" ", "")
    if 4 <= len(m) <= 7 and m[0] in ("X", "N") and any(c.isdigit() for c in m):
        return m
    return None


def _parse_conceptos(raw: object) -> list[dict]:
    """Renglones {concepto, monto} válidos (máx 8); lo demás se descarta."""
    if not isinstance(raw, list):
        return []
    out: list[dict] = []
    for item in raw[:8]:
        if not isinstance(item, dict):
            continue
        concepto = str(item.get("concepto", "")).strip()
        monto = item.get("monto")
        if concepto and isinstance(monto, (int, float)) and monto > 0:
            out.append({"concepto": concepto[:80], "monto": round(float(monto), 2)})
    return out


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
    '  "tipo_combustible": "TURBOSINA" o "AVGAS" según el ticket (GASAVION/'
    'GAS AVION/100LL = AVGAS; JET A/TURBOSINA = TURBOSINA), o null.\n'
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
        max_tokens=1000,
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
