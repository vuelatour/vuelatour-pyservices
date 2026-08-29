import json
import unicodedata
from functools import lru_cache

import anthropic

from app.config import get_settings
from app.schemas.vision import (
    CombustibleTicketResponse,
    ConstanciaFiscalRequest,
    ConstanciaFiscalResponse,
    GastoTicketRequest,
    GastoTicketResponse,
    ImagenFuente,
    InventarioEmpaque,
    InventarioItemRequest,
    InventarioItemResponse,
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
    '  "calidad_foto": "ALTA" si los dígitos se ven nítidos y completos; "MEDIA" '
    "si se leen pero hay reflejo, ángulo, sombra o algo de desenfoque; \"BAJA\" si "
    "la foto está borrosa/oscura/movida y algún dígito PODRÍA estar equivocado "
    "(en especial la décima del tambor pequeño).\n"
    '  "notas": string BREVE en español (máximo ~200 caracteres, UNA o dos '
    'frases). OBLIGATORIO cuando calidad_foto no es "ALTA": di QUÉ estorba y '
    'CUÁL dígito es el dudoso (ej. "foto borrosa: la décima podría ser 8 o 9"). '
    'No repitas la lectura ni narres todo el dial.\n'
    "Si dudas entre dos dígitos, elige el más probable, baja la confianza y "
    'reporta calidad_foto "BAJA". No inventes dígitos que no ves: si faltan, usa '
    "null y explica en notas. Es MUCHO peor entregar una lectura equivocada como "
    "segura que admitir la duda: de esa lectura salen las horas de toda la flota."
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


def _image_block(req: TacometroRequest | GastoTicketRequest | ImagenFuente) -> dict:
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


def _bloques_imagenes(imagenes: list[ImagenFuente]) -> list[dict]:
    """Bloques image (base64 o URL) para N fotos, en el orden recibido."""
    return [_image_block(img) for img in imagenes]


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
        return _bloques_imagenes(req.images)
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
    confianza = float(data.get("confianza", 0.0))
    calidad = str(data.get("calidad_foto", "") or "").strip().upper()
    if calidad not in {"ALTA", "MEDIA", "BAJA"}:
        # Modelo viejo o respuesta sin el campo: se deduce de la confianza para
        # no perder el aviso (nunca se asume ALTA "por defecto").
        calidad = "ALTA" if confianza >= 0.9 else "BAJA" if confianza < 0.7 else "MEDIA"
    return TacometroResponse(
        lectura=float(lectura) if isinstance(lectura, (int, float)) else None,
        confianza=confianza,
        legible=bool(data.get("legible", lectura is not None)),
        # Tope duro además del prompt: una nota kilométrica rompía la captura
        # aguas abajo (el API la valida por largo).
        notas=str(data.get("notas", ""))[:400],
        calidad_foto=calidad,
        modelo=s.anthropic_model,
    )


_TICKET_SYSTEM = (
    "Eres un asistente de captura de gastos para una empresa de aviación. A partir "
    "de la foto de un ticket o recibo, extraes los datos del gasto. Devuelves SOLO "
    "un objeto JSON, sin texto adicional ni ```fences```, con las claves exactas:\n"
    '  "monto": número con el TOTAL pagado (no subtotales), o null si ilegible. '
    "Si el total cobrado YA INCLUYE una línea de PROPINA/TIP, el monto es ese "
    "total CON la propina (lo que se cargó a la tarjeta); NUNCA sumes tú una "
    "propina al total impreso.\n"
    '  "propina": número con la PROPINA/TIP solo si está PAGADA e INCLUIDA en '
    "el TOTAL del ticket, o null si no aparece. IGNORA propinas SUGERIDAS u "
    'opcionales (ej. "propina sugerida 10%", tablas 10/15/20%) que NO están '
    "sumadas al total: en ese caso propina = null. NUNCA la inventes ni la "
    "calcules.\n"
    '  "moneda": "MXN" o "USD" según el ticket (default "MXN" en México).\n'
    '  "fecha": fecha del ticket en formato YYYY-MM-DD, o null.\n'
    '  "proveedor": nombre del comercio/proveedor, o null.\n'
    '  "folio": numero de folio/remision/ticket/factura impreso en el documento '
    '(ej. "Remision: 2622242310" → "2622242310"; "Folio A-1234" → "A-1234"; '
    '"FACTURA No. FEDCUN 193422" → "FEDCUN 193422" — las facturas de aeropuerto '
    "lo traen en el recuadro superior derecho). Solo el identificador, sin la "
    "palabra Folio/Remision. Si hay varios, el de la remision o folio principal "
    "(NO el folio fiscal UUID). Si no aparece: null. Este dato evita capturas "
    "duplicadas: no lo inventes.\n"
    '  "concepto": descripción breve de lo comprado, o null.\n'
    '  "categoria_sugerida": una de GAS, GASOLINA, OPERACIONES, ATERRIZAJE, '
    "TUAS, FBO, COMIDA, HOTEL, TAXI, REFACCION, SERVICIOS, PERMISO, FIJO, "
    "OTRO (la más probable), o "
    "null. Guía: SERVICIOS = servicio/mano de obra contratada al AVIÓN "
    "(taller, mantenimiento, lavado, inspección — factura de servicio, no de "
    "piezas; si el ticket es de PIEZAS/refacciones usa REFACCION). "
    "ATERRIZAJE = cuotas de aterrizaje/estacionamiento de pista "
    "o plataforma de la AERONAVE; TUAS = tarifa de uso de aeropuerto (TUA); "
    "FBO = servicios FBO/handling. OJO: el ESTACIONAMIENTO DE AUTOS (pensión "
    "o parking del aeropuerto/ciudad donde la tripulación deja el coche) es "
    "un viático de la tripulación → TAXI, NUNCA OPERACIONES ni ATERRIZAJE. "
    "Los VIÁTICOS de la tripulación tienen categoría propia — NO los mandes a "
    "OTRO: COMIDA = restaurante, desayuno/comida/cena, torta, café, alimentos "
    "de la tripulación (aunque sea tienda de conveniencia u OXXO); TAXI = "
    "taxi, Uber/DiDi, mototaxi, autobús, ferry, estacionamiento de autos y "
    "cualquier transporte terrestre/marítimo de la tripulación; HOTEL = "
    "hospedaje de la "
    "tripulación (incluye impuestos del hotel como saneamiento). OTRO es solo "
    "para lo que no es de la tripulación: comisariato (snacks/bebidas para "
    "los PASAJEROS), mensajería/paquetería, papelería, compras de oficina y "
    "varios sin categoría clara. Desempate: si el ticket no deja ver para "
    "quién es (súper/OXXO, ferry, restaurante sin más señas), asume "
    "TRIPULACIÓN (COMIDA/TAXI/HOTEL); señales de PASAJEROS mandan a OTRO "
    "(varios boletos en el mismo ticket, consumo a nombre del cliente, "
    "leyendas que mencionen pasajeros). OJO: gasolina/diésel de un VEHÍCULO "
    "TERRESTRE (gasolinera Pemex, Gulf, OXXO GAS — que NO es la tienda OXXO; "
    "Magna/Premium/Regular/diésel) NUNCA es GAS ni COMIDA: es GASOLINA "
    "(combustible de coche/camioneta). GAS es SOLO combustible de AVIACIÓN "
    "(ASA, GAFSACOMM, AVGAS, turbosina, gasavión). TAXI queda para "
    "transporte pagado (taxi/Uber/ferry), no para cargar gasolina.\n"
    '  "medio_pago": "EFECTIVO", "TARJETA_CORP" o "TRANSFERENCIA" segun el ticket '
    "(DEBITO/CREDITO/VISA/MASTERCARD/TARJETA = TARJETA_CORP; EFECTIVO/CASH = "
    "EFECTIVO; SPEI/TRANSFERENCIA = TRANSFERENCIA), o null. OJO: muchas fotos "
    "traen ENGRAPADO o encimado sobre la factura un voucher chico de terminal "
    'bancaria (Banamex/Global Payments/EVO: "TOTAL $", "TARJETA", '
    '"AUTORIZACION", "Debit/Credit", "COPIA CLIENTE/NEGOCIO") — ese voucher ES '
    "la evidencia del pago: medio_pago = TARJETA_CORP aunque la factura no "
    "mencione el medio.\n"
    '  "tarjeta_terminacion": ultimos 4 digitos de la tarjeta si aparecen, como '
    "string de 4 digitos, o null. Tambien se leen del voucher bancario "
    'engrapado (ej. "TARJETA 6250 Debito" → "6250").\n'
    '  "litros": SOLO si es ticket de COMBUSTIBLE (gasavión/AVGAS/turbosina/'
    "Jet A): litros cargados; si viene en galones conviértelo (1 gal = "
    "3.78541 L). Cualquier otro ticket: null. Las horas del balance dependen "
    "de este dato: no lo inventes.\n"
    '  "conceptos": lista [{"concepto": string, "monto": número}] con el desglose '
    "de la factura. PRIORIDAD 1 — TABLA RESUMEN: muchas facturas de aeropuerto "
    "(ASUR/ADO) traen al pie una tabla resumen por SECCIONES (Operaciones, Taxi, "
    "Autobuses, Tarifa TUA, Combustible, FOB) con columnas SubTotal/IVA/Total; si "
    "existe, usa ESA tabla: un concepto por sección con Total distinto de cero, "
    'con "monto" = columna TOTAL (IVA YA INCLUIDO), eligiendo solo las secciones '
    "cuya suma dé exactamente el TOTAL pagado de la factura, y NO agregues renglón "
    "de IVA aparte. PRIORIDAD 2 — sin tabla resumen: usa los RENGLONES del "
    "ticket/factura SOLO si desglosa claramente varios conceptos (ej. aterrizaje, "
    "plataforma de pernocta, embarque, servicio FBO, T.U.A.); máximo 8; [] si es "
    "un solo concepto o no se distinguen. EXCEPCIÓN al caso de un solo concepto: "
    "si el ÚNICO renglón es TUA o FBO (ej. factura CFDI que solo cobra 'TUA "
    "NACIONAL'), SÍ regresa ese concepto con su renglón de IVA si la factura lo "
    "desglosa (ej. total $939.60 → [{'concepto': 'TUA NACIONAL', 'monto': 810.00}, "
    "{'concepto': 'IVA 16%', 'monto': 129.60}]) — la separación contable del "
    "balance depende de ese desglose. CLAVES DE SERVICIO: algunos aeropuertos "
    "(ej. Cozumel) imprimen renglones genéricos 'Servicio' con una CLAVE numérica "
    "de concepto — conserva SIEMPRE la clave en el texto del concepto (ej. "
    "'Servicio (clave 230700)') y, si la factura trae en otra parte la "
    "descripción textual de esa clave (Descripción del CFDI, leyendas, tabla de "
    "tarifas), usa el nombre real del servicio conservando la clave (ej. 'Tarifa "
    "de Uso de Aeropuerto (clave 230700)'). OJO: esas facturas suelen traer DOS "
    "columnas de clave — 'Clave' (la del CONCEPTO del aeropuerto, 6 dígitos, ej. "
    "210100/230700, distinta por renglón) y 'Clave SAT' (catálogo genérico, 8 "
    "dígitos, ej. 78141805, IGUAL en todos los renglones): conserva la 'Clave' "
    "del concepto, NUNCA la Clave SAT. OJO CON LOS DESCUENTOS: si un renglón "
    "tiene columna Descuento (muy común en T.U.A.), el monto del renglón es el "
    "NETO = importe − descuento (ej. importe $605.18 con descuento $5.18 → monto "
    "600.00). En este caso, si la factura muestra subtotal + IVA por separado, "
    'agrega el IVA como un renglón más ({"concepto": "IVA 16%", "monto": X}) para '
    "que la SUMA de renglones (ya con descuentos aplicados) sea exactamente el "
    "TOTAL pagado; si aun así no cuadra, [].\n"
    '  "matricula": matrícula de la aeronave si aparece en el documento '
    '(observaciones/referencias; formato XA-ABC, XB-ABC o N123XX), o null.\n'
    '  "confianza": número entre 0 y 1.\n'
    '  "legible": true/false según si el ticket se distingue.\n'
    '  "notas": string breve en español con cualquier observación.\n'
    "No inventes datos que no aparezcan: usa null. GAS es SOLO combustible de "
    "AVIACIÓN (gasavión/AVGAS/turbosina/Jet A). Para facturas de aeropuerto: "
    "OPERACIONES es la categoría de las MIXTAS o generales (aterrizaje + "
    "servicios + FBO juntos, prefiérela sobre la legada ATERRIZAJE); usa TUAS "
    "solo si la factura es PURAMENTE TUA, y FBO solo si es puramente "
    "handling/FBO."
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
        "GAS", "GASOLINA", "OPERACIONES", "ATERRIZAJE", "TUAS", "FBO",
        "COMIDA", "HOTEL", "TAXI", "REFACCION", "PERMISO", "FIJO", "OTRO",
        # SERVICIOS (29-ago): servicio contratado al avión — SÍ está en el
        # prompt (mano de obra/taller, a diferencia de REFACCION = piezas).
        "SERVICIOS",
        # No están en el prompt (no se leen de un ticket) pero se toleran
        # si el modelo las eco-devuelve en un reanálisis: VISITA (gasto de
        # visitante, 27-ago), PERSONAL_DUENO y NOMINA (sueldos, 29-ago).
        "VISITA", "PERSONAL_DUENO", "NOMINA",
    }
    monto_f = float(monto) if isinstance(monto, (int, float)) else None
    # Propina solo si es coherente (0 < propina < monto): una lectura donde
    # "propina" >= total es un misread y dejaría el ticket en <= 0 en el
    # autofill de panel/app — se descarta aquí (punto único) conservando el
    # monto, que es el dato sagrado.
    propina = data.get("propina")
    propina_f = (
        float(propina)
        if isinstance(propina, (int, float))
        and monto_f is not None
        and 0 < propina < monto_f
        else None
    )
    return GastoTicketResponse(
        monto=monto_f,
        propina=propina_f,
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
        litros=(
            float(data["litros"])
            if isinstance(data.get("litros"), (int, float)) and data["litros"] > 0
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


# Catálogo c_RegimenFiscal del SAT: la constancia muestra el NOMBRE del
# régimen; el CFDI 4.0 exige el CÓDIGO. La IA mapea con esta tabla y aquí
# validamos que el código devuelto exista (parsing defensivo).
_REGIMENES_SAT: frozenset[str] = frozenset(
    {
        "601", "603", "605", "606", "607", "608", "610", "611", "612", "614",
        "615", "616", "620", "621", "622", "623", "624", "625", "626",
    }
)

_CONSTANCIA_SYSTEM = (
    "Eres un asistente fiscal mexicano. El documento es una CONSTANCIA DE "
    "SITUACIÓN FISCAL emitida por el SAT (cédula de identificación fiscal). "
    "Extraes los datos que exige el CFDI 4.0. Devuelves SOLO un objeto JSON, "
    "sin texto adicional ni ```fences```, con las claves exactas:\n"
    '  "rfc": el RFC del contribuyente (12-13 caracteres, MAYÚSCULAS), o null.\n'
    '  "razon_social": para persona MORAL la "Denominación/Razón Social" SIN el '
    "régimen societario ('S.A. DE C.V.', 'S. DE R.L.', etc.), que en la "
    "constancia aparece SEPARADO como 'Régimen de capital' — el CFDI 4.0 exige "
    "el nombre EXACTO sin ese sufijo. Para persona FÍSICA el nombre completo: "
    "Nombre(s) + Primer Apellido + Segundo Apellido. Mantener tal cual aparece "
    "en la constancia. O null.\n"
    '  "regimen_fiscal": el CÓDIGO c_RegimenFiscal de 3 dígitos del régimen '
    "VIGENTE más reciente del apartado 'Regímenes'. La constancia muestra el "
    "NOMBRE; mapéalo con este catálogo: 601 General de Ley Personas Morales; "
    "603 Personas Morales con Fines no Lucrativos; 605 Sueldos y Salarios; "
    "606 Arrendamiento; 607 Enajenación o Adquisición de Bienes; 608 Demás "
    "ingresos; 610 Residentes en el Extranjero; 611 Dividendos; 612 Personas "
    "Físicas con Actividades Empresariales y Profesionales; 614 Intereses; "
    "615 Obtención de premios; 616 Sin obligaciones fiscales; 620 Sociedades "
    "Cooperativas; 621 Incorporación Fiscal; 622 Actividades AGAPES; 623 "
    "Opcional para Grupos de Sociedades; 624 Coordinados; 625 Plataformas "
    "Tecnológicas; 626 Régimen Simplificado de Confianza (RESICO). O null si "
    "no puedes mapearlo.\n"
    '  "regimen_descripcion": el nombre del régimen tal cual aparece en la '
    "constancia, o null.\n"
    '  "cp": código postal del domicilio fiscal (5 dígitos, como string), o null.\n'
    '  "domicilio": el domicilio fiscal completo en UNA línea (calle, número, '
    "colonia, municipio/alcaldía, estado), o null.\n"
    '  "confianza": número entre 0 y 1.\n'
    '  "legible": true/false según si el documento se distingue.\n'
    '  "motivo": string breve en español (por qué no es legible, u '
    "observaciones), o null.\n"
    "No inventes datos que no aparezcan en el documento: usa null. Si el "
    "documento NO es una constancia de situación fiscal del SAT, devuelve "
    "legible=false y explica en motivo."
)

_CONSTANCIA_PROMPT = (
    "Extrae los datos fiscales de esta Constancia de Situación Fiscal del SAT "
    "y responde con el JSON indicado. Recuerda: razón social SIN régimen "
    "societario para persona moral, y el código de 3 dígitos del régimen "
    "VIGENTE más reciente."
)


def _constancia_blocks(req: ConstanciaFiscalRequest) -> list[dict]:
    """PDF como bloque document (visión nativa, igual que gasto-ticket) o
    una imagen base64."""
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
    return [
        {
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": req.media_type,
                "data": req.image_base64,
            },
        }
    ]


def _parse_rfc(raw: object) -> str | None:
    """RFC plausible: 12 (moral) o 13 (física) caracteres alfanuméricos."""
    if not isinstance(raw, str):
        return None
    rfc = raw.strip().upper().replace(" ", "").replace("-", "")
    if 12 <= len(rfc) <= 13 and rfc.isalnum():
        return rfc
    return None


def _parse_cp(raw: object) -> str | None:
    """CP mexicano: exactamente 5 dígitos (acepta int del modelo)."""
    if raw is None:
        return None
    cp = str(raw).strip()
    if len(cp) == 4 and cp.isdigit():
        # El modelo a veces devuelve int y pierde el cero inicial (CDMX).
        cp = f"0{cp}"
    return cp if len(cp) == 5 and cp.isdigit() else None


def _str_or_none(raw: object) -> str | None:
    if isinstance(raw, str) and raw.strip():
        return raw.strip()
    return None


def leer_constancia_fiscal(req: ConstanciaFiscalRequest) -> ConstanciaFiscalResponse:
    s = get_settings()
    resp = _client().messages.create(
        model=s.anthropic_model,
        max_tokens=1000,
        system=[
            {"type": "text", "text": _CONSTANCIA_SYSTEM, "cache_control": {"type": "ephemeral"}}
        ],
        messages=[
            {
                "role": "user",
                "content": [*_constancia_blocks(req), {"type": "text", "text": _CONSTANCIA_PROMPT}],
            }
        ],
    )
    text = next((b.text for b in resp.content if b.type == "text"), "")
    data = _extract_json(text)

    regimen = _str_or_none(data.get("regimen_fiscal"))
    rfc = _parse_rfc(data.get("rfc"))
    # Confianza defensiva: un valor raro del modelo (null/string) no debe
    # tirar la extracción completa — el contrato es degradar, nunca 500.
    conf = data.get("confianza")
    confianza = min(max(float(conf), 0.0), 1.0) if isinstance(conf, (int, float)) else 0.0
    return ConstanciaFiscalResponse(
        disponible=True,
        legible=bool(data.get("legible", rfc is not None)),
        rfc=rfc,
        razon_social=_str_or_none(data.get("razon_social")),
        # Solo códigos del catálogo c_RegimenFiscal: un código inventado
        # rompería el timbrado; mejor null y que el operador lo elija.
        regimen_fiscal=regimen if regimen in _REGIMENES_SAT else None,
        regimen_descripcion=_str_or_none(data.get("regimen_descripcion")),
        cp=_parse_cp(data.get("cp")),
        domicilio=_str_or_none(data.get("domicilio")),
        confianza=confianza,
        motivo=_str_or_none(data.get("motivo")),
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
    '  "tipo_combustible": "TURBOSINA" o "AVGAS" según el ticket (GASAVION/'
    'GAS AVION/100LL = AVGAS; JET A/TURBOSINA = TURBOSINA), o null.\n'
    '  "fecha": fecha YYYY-MM-DD, o null.\n'
    '  "hora": hora de la carga en formato HH:MM de 24 horas, o null.\n'
    '  "proveedor": nombre del proveedor/FBO, o null.\n'
    '  "folio": numero de folio/remision impreso en el ticket (ej. '
    '"Remision: 2622242310" → "2622242310"). Solo el identificador, sin la '
    "palabra Folio/Remision. Si no aparece: null. Este dato evita capturas "
    "duplicadas: no lo inventes.\n"
    '  "tarjeta_terminacion": ultimos 4 digitos de la tarjeta de pago si aparecen '
    'en el ticket o en un voucher bancario engrapado (p. ej. "**** 1234" o '
    '"TARJETA ...1234"), como string de 4 digitos, o null.\n'
    '  "medio_pago": "EFECTIVO", "TARJETA_CORP" o "TRANSFERENCIA" segun el ticket, '
    "o null. Un voucher de terminal bancaria engrapado/encimado (TOTAL, TARJETA, "
    "AUTORIZACION, Debit/Credit) = TARJETA_CORP.\n"
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
        system=[
            {"type": "text", "text": _COMBUSTIBLE_SYSTEM, "cache_control": {"type": "ephemeral"}}
        ],
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


# --------------------------------------------------------------------------
# Inventario: ficha del producto a partir de fotos (28-ago-2026)
# --------------------------------------------------------------------------

_INVENTARIO_SYSTEM = (
    "Eres el asistente de bodega de una empresa de aviación chárter (VuelaTour, "
    "Cancún). A partir de una o varias FOTOS DEL MISMO PRODUCTO (distintos ángulos, "
    "etiqueta, código de barras o su caja/empaque) llenas la ficha del ítem de "
    "inventario. Devuelves SOLO un objeto JSON, sin texto adicional ni ```fences```, "
    "con las claves exactas:\n"
    '  "nombre": nombre corto y útil para bodega (marca + producto + especificación + '
    'presentación), ej. "AeroShell W15W-50 aceite motor pistón 1 qt", o null.\n'
    '  "marca": marca/fabricante (ej. "AeroShell"), o null.\n'
    '  "numero_parte": número de parte / product code del FABRICANTE tal como está '
    'impreso (en una caja AeroShell aparece como "Product code 550050835"), o null. '
    "No lo confundas con el código de barras.\n"
    '  "codigo_barras": los DÍGITOS impresos bajo las barras del código de la UNIDAD '
    "(UPC-A de 12 dígitos, EAN-13 de 13…), como string SIN espacios ni guiones "
    '(ej. impreso "0 21400 06215 3" → "021400062153"). Léelos dígito por dígito de '
    "la impresión; si no son legibles con seguridad → null. NUNCA inventes ni "
    '"completes" dígitos.\n'
    '  "contenido": contenido/presentación de la UNIDAD (ej. "946 mL", "1 qt (946 mL)", '
    '"5 L", "1 pieza"), o null.\n'
    '  "unidad": unidad de medida sugerida para controlar existencias ("pieza", '
    '"botella", "litro", "galón", "juego", "kit", "metro"…), o null.\n'
    '  "categoria": la categoría que mejor encaje de la lista CATEGORÍAS que recibirás '
    "(compara de forma laxa: mayúsculas, acentos y plural no importan; devuelve el "
    "texto EXACTO de la lista). Si ninguna encaja, propón una nueva y corta (1-2 "
    "palabras), o null.\n"
    '  "descripcion": descripción de ficha en 1-2 líneas: tipo de producto, uso y '
    'especificación/norma (ej. "Aceite semisintético multigrado SAE J1899 para '
    'motores de pistón de aviación"), o null.\n'
    '  "empaque": si ALGUNA foto es la CAJA/EMPAQUE que agrupa varias unidades: '
    '{"nombre": "Caja de N", "factor": N (unidades por empaque; "6 - 1qt bottles" → 6), '
    '"codigo_barras": los dígitos del código de la CAJA (normalmente ITF-14 de 14 '
    'dígitos, ej. impreso "0 00 21400 06216 0" → "00021400062160"), o null si no es '
    "legible}. Si ninguna foto es un empaque → null.\n"
    '  "confianza": número entre 0 y 1 (seguridad de la identificación y de los códigos).\n'
    '  "notas_ia": string breve en español con dudas u observaciones (p. ej. "código '
    'parcialmente legible", "la foto 2 es la caja"), o null.\n'
    "Reglas: todas las fotos son del MISMO producto; combina lo que se ve en cada una. "
    "Un código de 12-13 dígitos en la botella/pieza es de la unidad; uno de 14 dígitos "
    "en una caja es del empaque. Si dudas de un dígito, prefiere null y explícalo en "
    "notas_ia: un código inventado contamina el inventario. No inventes marca ni número "
    "de parte: usa null cuando no aparezcan."
)


def _norm_codigo_barras(raw: object) -> str | None:
    """Dígitos del código sin espacios ni guiones; None si vacío."""
    if raw is None or isinstance(raw, bool):
        return None
    if isinstance(raw, (int, float)):
        # El prompt pide string; si el modelo mandó número ya perdió ceros a
        # la izquierda, pero el valor sigue siendo útil para el operador.
        s = str(int(raw)) if float(raw).is_integer() else str(raw)
    else:
        s = "".join(str(raw).split()).replace("-", "")
    return s or None


def _str_o_numero(raw: object) -> str | None:
    """Texto; un número entero del modelo (product code) se regresa sin '.0'."""
    if isinstance(raw, bool) or raw is None:
        return None
    if isinstance(raw, (int, float)):
        return str(int(raw)) if float(raw).is_integer() else str(raw)
    return _str_or_none(raw)


def _norm_texto_laxo(s: str) -> str:
    t = unicodedata.normalize("NFD", s)
    t = "".join(ch for ch in t if not unicodedata.combining(ch))
    return " ".join(t.lower().split())


def _empata_categoria(propuesta: object, catalogo: list[str]) -> str | None:
    """Categoría del catálogo (texto EXACTO) que empata de forma laxa con la
    propuesta del modelo; si ninguna empata, la propuesta como categoría nueva."""
    p = _str_or_none(propuesta)
    if p is None:
        return None
    np_ = _norm_texto_laxo(p)
    for c in catalogo:
        if _norm_texto_laxo(c) == np_:
            return c
    for c in catalogo:
        nc = _norm_texto_laxo(c)
        if nc.rstrip("s") == np_.rstrip("s") or (len(np_) >= 4 and (np_ in nc or nc in np_)):
            return c
    return p


def _factor_empaque(raw: object) -> int | None:
    if raw is None or isinstance(raw, bool):
        return None
    if isinstance(raw, (int, float)):
        return int(round(raw)) if raw >= 1 else None
    s = str(raw).strip()
    return int(s) if s.isdigit() and int(s) >= 1 else None


def _concilia_codigos(
    unidad: str | None, empaque: str | None, escaneados: list[str]
) -> tuple[str | None, str | None, str | None]:
    """Los códigos escaneados son la VERDAD: el resultado solo puede contener
    códigos de esa lista. El modelo decide cuál es de la caja (por default el
    primero es el de la unidad). Regresa (unidad, empaque, nota)."""
    nota: str | None = None
    if not escaneados:
        if unidad and unidad == empaque:
            return unidad, None, "Mismo código para unidad y empaque; se dejó solo en la unidad."
        return unidad, empaque, None
    validos = set(escaneados)
    if unidad not in validos and empaque not in validos:
        unidad = escaneados[0]
        empaque = escaneados[1] if len(escaneados) > 1 else None
        nota = "Códigos tomados del lector (la IA no los leyó igual en las fotos)."
    elif unidad not in validos:
        # El modelo ubicó el escaneado en la caja: la unidad es el otro (si hay).
        resto = [c for c in escaneados if c != empaque]
        unidad = resto[0] if resto else None
    elif empaque is not None and empaque not in validos:
        resto = [c for c in escaneados if c != unidad]
        empaque = resto[0] if resto else None
    if empaque is None:
        sobrantes = [c for c in escaneados if c != unidad]
        if sobrantes:
            empaque = sobrantes[0]  # el segundo código escaneado es la caja
    if empaque == unidad:
        empaque = None
    return unidad, empaque, nota


def leer_producto_inventario(req: InventarioItemRequest) -> InventarioItemResponse:
    s = get_settings()
    blocks = _bloques_imagenes(req.images)
    escaneados = list(
        dict.fromkeys(
            c for c in (_norm_codigo_barras(x) for x in (req.codigos_escaneados or [])) if c
        )
    )
    categorias = [c.strip() for c in req.categorias if c and c.strip()]

    partes: list[str] = []
    if len(blocks) > 1:
        partes.append(
            f"Las {len(blocks)} fotos son del MISMO producto (distintos ángulos, etiqueta "
            "o su caja/empaque): combínalas para llenar una sola ficha."
        )
    else:
        partes.append("La foto es del producto que se dará de alta en inventario.")
    if categorias:
        partes.append("CATEGORÍAS disponibles: " + "; ".join(categorias) + ".")
    else:
        partes.append("No hay catálogo de categorías: propón una corta.")
    if escaneados:
        lista = "; ".join(f"{i}) {c}" for i, c in enumerate(escaneados, start=1))
        partes.append(
            f"CÓDIGOS ESCANEADOS con lector (son la VERDAD; cópialos tal cual, no los "
            f"reinventes ni alteres): {lista}. El primero es el código de la UNIDAD, salvo "
            "que las fotos muestren claramente que corresponde a la caja/empaque (entonces "
            "va en empaque.codigo_barras y el otro, si lo hay, en codigo_barras)."
        )
    partes.append("Responde SOLO con el JSON indicado.")
    prompt = "\n".join(partes)

    resp = _client().messages.create(
        model=s.anthropic_model,
        max_tokens=1500,
        system=[
            {"type": "text", "text": _INVENTARIO_SYSTEM, "cache_control": {"type": "ephemeral"}}
        ],
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

    emp_raw = data.get("empaque")
    emp_nombre: str | None = None
    emp_factor: int | None = None
    emp_codigo: str | None = None
    if isinstance(emp_raw, dict):
        emp_nombre = _str_or_none(emp_raw.get("nombre"))
        emp_factor = _factor_empaque(emp_raw.get("factor"))
        emp_codigo = _norm_codigo_barras(emp_raw.get("codigo_barras"))
    codigo, emp_codigo, nota_codigos = _concilia_codigos(
        _norm_codigo_barras(data.get("codigo_barras")), emp_codigo, escaneados
    )
    if emp_nombre is None and emp_factor:
        emp_nombre = f"Caja de {emp_factor}"
    empaque = (
        InventarioEmpaque(nombre=emp_nombre, factor=emp_factor, codigo_barras=emp_codigo)
        if (emp_nombre or emp_factor or emp_codigo)
        else None
    )

    notas = [n for n in (_str_or_none(data.get("notas_ia")), nota_codigos) if n]
    conf = data.get("confianza")
    confianza = (
        min(max(float(conf), 0.0), 1.0)
        if isinstance(conf, (int, float)) and not isinstance(conf, bool)
        else 0.0
    )
    return InventarioItemResponse(
        nombre=_str_or_none(data.get("nombre")),
        marca=_str_or_none(data.get("marca")),
        numero_parte=_str_o_numero(data.get("numero_parte")),
        codigo_barras=codigo,
        categoria=_empata_categoria(data.get("categoria"), categorias),
        unidad=_str_or_none(data.get("unidad")),
        contenido=_str_o_numero(data.get("contenido")),
        descripcion=_str_or_none(data.get("descripcion")),
        empaque=empaque,
        confianza=confianza,
        notas_ia=" ".join(notas) or None,
        modelo=s.anthropic_model,
    )
