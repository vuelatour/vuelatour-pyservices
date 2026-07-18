import base64
import io
import json
import re
from functools import lru_cache

from app.config import get_settings
from app.schemas.conciliacion import (
    ConciliacionParseRequest,
    ConciliacionParseResponse,
    ConciliacionSugerirRequest,
    ConciliacionSugerirResponse,
    MovimientoParseado,
)


def _norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", s.lower())


def _find_col(cols: list[str], *needles: str) -> str | None:
    for c in cols:
        n = _norm(c)
        if any(needle in n for needle in needles):
            return c
    return None


def _to_float(v) -> float | None:
    if v is None:
        return None
    s = str(v).strip().replace("$", "").replace(",", "")
    if s in ("", "-", "nan"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _parse_tabular(req: ConciliacionParseRequest, formato: str) -> ConciliacionParseResponse:
    import pandas as pd  # lazy

    raw = base64.b64decode(req.file_base64)
    buf = io.BytesIO(raw)
    if formato == "csv":
        df = pd.read_csv(buf, dtype=str, keep_default_na=False)
    else:
        df = pd.read_excel(buf, dtype=str)

    cols = [str(c) for c in df.columns]
    col_fecha = _find_col(cols, "fecha", "date")
    col_desc = _find_col(cols, "concepto", "descrip", "detalle", "referencia", "movimiento")
    col_cargo = _find_col(cols, "cargo", "retiro", "debito", "debit")
    col_abono = _find_col(cols, "abono", "deposito", "credito", "credit")
    col_monto = _find_col(cols, "monto", "importe", "amount")

    movimientos: list[MovimientoParseado] = []
    for _, row in df.iterrows():
        fecha = None
        if col_fecha:
            try:
                ts = pd.to_datetime(str(row[col_fecha]), dayfirst=True, errors="coerce")
                if ts is not None and not pd.isna(ts):
                    fecha = ts.strftime("%Y-%m-%d")
            except (ValueError, TypeError):
                fecha = None

        desc = str(row[col_desc]).strip() if col_desc else None
        ref = None

        cargo = _to_float(row[col_cargo]) if col_cargo else None
        abono = _to_float(row[col_abono]) if col_abono else None
        if cargo and cargo != 0:
            movimientos.append(
                MovimientoParseado(fecha=fecha, descripcion=desc, monto=abs(cargo), tipo="CARGO", referencia=ref)
            )
        elif abono and abono != 0:
            movimientos.append(
                MovimientoParseado(fecha=fecha, descripcion=desc, monto=abs(abono), tipo="ABONO", referencia=ref)
            )
        elif col_monto:
            m = _to_float(row[col_monto])
            if m is not None and m != 0:
                movimientos.append(
                    MovimientoParseado(
                        fecha=fecha,
                        descripcion=desc,
                        monto=abs(m),
                        tipo="ABONO" if m > 0 else "CARGO",
                        referencia=ref,
                    )
                )

    notas = "" if movimientos else "No se reconocieron columnas de monto. Revisa el formato del archivo."
    return ConciliacionParseResponse(
        movimientos=movimientos, total=len(movimientos), formato=formato, notas=notas
    )


_SYSTEM = (
    "Extraes movimientos de un estado de cuenta bancario en PDF. Devuelves SOLO un "
    "objeto JSON, sin texto adicional ni ```fences```, con la clave \"movimientos\": "
    "un arreglo de objetos con \"fecha\" (YYYY-MM-DD o null), \"descripcion\" (string), "
    "\"monto\" (número positivo), \"tipo\" (\"CARGO\" si es salida/retiro, \"ABONO\" si es "
    "entrada/depósito) y \"referencia\" (string o null). Incluye solo movimientos "
    "reales, no encabezados ni saldos. El JSON debe ser COMPACTO (sin espacios ni "
    "saltos de línea) y las descripciones breves (máx ~10 palabras): los estados de "
    "cuenta traen cientos de movimientos y TODOS deben caber en la respuesta."
)


@lru_cache
def _client():
    import anthropic

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


_USA_CSV = (
    "Exporta el estado de cuenta en CSV o Excel desde el portal del banco e "
    "impórtalo: es el formato exacto y preferido."
)


def _parse_pdf(req: ConciliacionParseRequest) -> ConciliacionParseResponse:
    s = get_settings()
    # Timeout propio: generar cientos de movimientos tarda más que el timeout
    # global (90s) — la importación es manual y el operador espera.
    resp = _client().with_options(timeout=240.0).messages.create(
        model=s.anthropic_model,
        # Un estado de cuenta mensual trae cientos de movimientos: con 4096 la
        # respuesta se TRUNCABA a media estructura y el json.loads reventaba
        # con un error críptico ("Expecting ',' delimiter...").
        max_tokens=16384,
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
                            "data": req.file_base64,
                        },
                    },
                    {"type": "text", "text": "Extrae los movimientos y responde con el JSON indicado."},
                ],
            }
        ],
    )
    # Truncado por límite de salida: la conciliación exige el universo COMPLETO
    # de movimientos — importar una lista parcial en silencio es peor que
    # fallar con instrucciones claras.
    if resp.stop_reason == "max_tokens":
        raise ValueError(
            f"El PDF tiene demasiados movimientos para leerse completo con IA. {_USA_CSV}"
        )
    text = next((b.text for b in resp.content if b.type == "text"), "")
    try:
        data = _extract_json(text)
    except (json.JSONDecodeError, ValueError) as e:
        raise ValueError(
            "La IA no devolvió una respuesta interpretable al leer el PDF. "
            f"Reintenta, o mejor: {_USA_CSV}"
        ) from e

    movimientos: list[MovimientoParseado] = []
    for raw in data.get("movimientos", []):
        if not isinstance(raw, dict):
            continue
        monto = _to_float(raw.get("monto"))
        if monto is None or monto == 0:
            continue
        tipo = raw.get("tipo")
        movimientos.append(
            MovimientoParseado(
                fecha=str(raw["fecha"]) if raw.get("fecha") else None,
                descripcion=str(raw["descripcion"]) if raw.get("descripcion") else None,
                monto=abs(monto),
                tipo="CARGO" if tipo == "CARGO" else "ABONO",
                referencia=str(raw["referencia"]) if raw.get("referencia") else None,
            )
        )
    return ConciliacionParseResponse(
        movimientos=movimientos, total=len(movimientos), formato="pdf", modelo=s.anthropic_model
    )


def parsear_estado_cuenta(req: ConciliacionParseRequest) -> ConciliacionParseResponse:
    name = req.filename.lower()
    if name.endswith(".csv"):
        return _parse_tabular(req, "csv")
    if name.endswith(".xlsx"):
        return _parse_tabular(req, "excel")
    if name.endswith(".xls"):
        # pandas necesitaría xlrd (no instalado) → mismo mensaje que visión.
        raise ValueError("Formato .xls (Excel viejo) no soportado: guárdalo como .xlsx")
    if name.endswith(".pdf"):
        return _parse_pdf(req)
    raise ValueError(f"Formato no soportado: {req.filename}")


_SUGERIR_SYSTEM = (
    "Eres un asistente de conciliación bancaria para una empresa de aviación. "
    "Recibes UN movimiento bancario sin conciliar y una lista de gastos candidatos "
    "(ya filtrados por fecha y monto cercanos). Tu trabajo es elegir el gasto que "
    "con más probabilidad corresponde a ese movimiento. Devuelves SOLO un objeto "
    "JSON, sin texto adicional ni ```fences```, con las claves exactas:\n"
    '  "gasto_id_sugerido": el "id" del gasto candidato más probable, o null si '
    "ninguno encaja razonablemente.\n"
    '  "confianza": número entre 0 y 1 (qué tan seguro estás del match).\n'
    '  "razon": explicación breve en español (ej. "mismo proveedor y monto exacto", '
    '"diferencia de $3 = comisión bancaria", "fechas a 1 día").\n'
    "Considera: pequeñas diferencias de monto suelen ser comisiones o redondeos; el "
    "proveedor del gasto puede aparecer dentro de la descripción del movimiento; "
    "fechas cercanas (no idénticas) son normales por el tiempo de procesamiento "
    "bancario. Usa solo ids que estén en la lista de candidatos. Si la lista está "
    "vacía o nada encaja, devuelve gasto_id_sugerido=null con confianza 0."
)


def sugerir_conciliacion(req: ConciliacionSugerirRequest) -> ConciliacionSugerirResponse:
    s = get_settings()

    movimiento = {
        "fecha": req.movimiento.fecha,
        "monto": req.movimiento.monto,
        "descripcion": req.movimiento.descripcion,
    }
    candidatos = [
        {"id": c.id, "fecha": c.fecha, "monto": c.monto, "proveedor": c.proveedor}
        for c in req.candidatos
    ]
    ids_validos = {c.id for c in req.candidatos}

    if not candidatos:
        return ConciliacionSugerirResponse(
            gasto_id_sugerido=None,
            confianza=0.0,
            razon="No hay gastos candidatos para este movimiento.",
            modelo=s.anthropic_model,
        )

    payload = json.dumps(
        {"movimiento": movimiento, "candidatos": candidatos}, ensure_ascii=False
    )
    resp = _client().messages.create(
        model=s.anthropic_model,
        max_tokens=512,
        system=[{"type": "text", "text": _SUGERIR_SYSTEM, "cache_control": {"type": "ephemeral"}}],
        messages=[
            {
                "role": "user",
                "content": (
                    "Movimiento sin conciliar y gastos candidatos (JSON):\n"
                    f"{payload}\n\n"
                    "Elige el match más probable y responde con el JSON indicado."
                ),
            }
        ],
    )
    text = next((b.text for b in resp.content if b.type == "text"), "")
    data = _extract_json(text)

    sugerido = data.get("gasto_id_sugerido")
    sugerido = str(sugerido) if sugerido is not None else None
    # No confíes en ids inventados por el modelo: solo acepta candidatos reales.
    if sugerido not in ids_validos:
        sugerido = None

    return ConciliacionSugerirResponse(
        gasto_id_sugerido=sugerido,
        confianza=float(data.get("confianza", 0.0)) if sugerido else 0.0,
        razon=str(data.get("razon", "")),
        modelo=s.anthropic_model,
    )
