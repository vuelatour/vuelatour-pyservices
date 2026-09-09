import base64
import io
import json
import re
import unicodedata
from functools import lru_cache

from app.config import get_settings
from app.schemas.conciliacion import (
    ConciliacionParseRequest,
    ConciliacionParseResponse,
    ConciliacionSugerirRequest,
    ConciliacionSugerirResponse,
    MapeoColumnasPaywise,
    MovimientoParseado,
)
from app.services.ia_usage import uso_ia_de


def _norm(s: str) -> str:
    """Encabezado normalizado: sin acentos (ó→o), minúsculas, solo [a-z0-9].

    Transliterar (y no solo borrar) los acentos importa: "Comisión" debe
    dar "comision" para que las agujas de Paywise/banco lo reconozcan."""
    sin_acentos = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]", "", sin_acentos.lower())


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


def _leer_tabla(req: ConciliacionParseRequest, formato: str):
    """DataFrame de texto del CSV/XLSX (celdas vacías = "")."""
    import pandas as pd  # lazy

    raw = base64.b64decode(req.file_base64)
    buf = io.BytesIO(raw)
    if formato == "csv":
        return pd.read_csv(buf, dtype=str, keep_default_na=False)
    df = pd.read_excel(buf, dtype=str)
    return df.fillna("")


def _fecha_de(v) -> str | None:
    import pandas as pd  # lazy

    txt = str(v).strip()
    if not txt or txt.lower() in ("nan", "nat", "none"):
        return None
    try:
        # Paywise exporta "dd/mm/yyyy HH:MM" (dayfirst, formato mexicano) o
        # ISO "yyyy-mm-dd" (pandas 3 con dayfirst=True invierte mes/día en
        # ISO: se parsea aparte).
        es_iso = re.match(r"^\d{4}-\d{2}-\d{2}", txt) is not None
        ts = pd.to_datetime(txt, dayfirst=not es_iso, errors="coerce")
    except (ValueError, TypeError):
        return None
    if ts is None or pd.isna(ts):
        return None
    return ts.strftime("%Y-%m-%d")


# ---------------------------------------------------------------------------
# PAYWISE (9-sep-2026): estado de cuenta de la PASARELA de cobro. Cada fila es
# una operación con BRUTO (lo que pagó el cliente), COMISIÓN retenida y NETO
# depositado. Se normaliza a movimientos bancarios: `monto` = NETO, tipo ABONO
# (pagos) o CARGO (reembolsos/contracargos), con bruto/comisión ADITIVOS y la
# referencia (ID de operación) poblada para el cotejo del API.
# ---------------------------------------------------------------------------

# Encabezados que identifican cada columna (sub-cadenas normalizadas, sin
# acentos ni signos). Orden = prioridad. Tolerantes a variantes de nombre.
_PW_FECHA = (
    "fechadeoperacion", "fechaoperacion", "fechadepago", "fechapago",
    "fechadetransaccion", "fechatransaccion", "fecha", "date",
)
_PW_BRUTO = (
    "montobruto", "importebruto", "bruto", "montototal", "importetotal",
    "totalcobrado", "montocobrado", "montodelpago", "montopagado",
    "gross", "amount", "total", "monto", "importe",
)
_PW_COMISION = (
    "comisiontotal", "comision", "commission", "fee", "cargoporservicio", "costodelservicio",
)
_PW_NETO = (
    "montoneto", "importeneto", "neto", "netodepositado", "depositado",
    "deposito", "net", "aliquidar", "liquidado", "montoliquidado",
)
_PW_REF = (
    "idoperacion", "iddeoperacion", "idtransaccion", "iddetransaccion",
    "idpago", "numerodeautorizacion", "noautorizacion", "autorizacion",
    "referencia", "folio", "orderid", "transactionid",
)
# "id" solo por coincidencia EXACTA: por contención ganaría "Cantidad" o
# "Validación" cuando el archivo no trae columna de referencia.
_PW_REF_EXACTO = ("id",)
# Forma de BANCO (cargo + abono en columnas separadas): jamás es Paywise
# aunque traiga una columna "Comisión" — la detección automática se abstiene
# (el mapeo manual sigue disponible).
_PW_BANCO_CARGO = ("cargo", "retiro", "debito", "debit")
_PW_BANCO_ABONO = ("abono", "deposito", "credito", "credit")
_PW_ESTATUS = ("estatus", "estado", "status")
_PW_DESC = ("concepto", "descripcion", "detalle", "cliente", "nombre", "producto")
_PW_TARJETA = ("ultimos4", "ultimos4digitos", "terminacion", "tarjeta", "card", "last4")
_PW_TIPO = ("tipodeoperacion", "tipooperacion", "tipodemovimiento", "tipo", "operacion")

# Estatus que NO son dinero depositado (se omiten y se cuentan en notas).
_PW_ESTATUS_OMITIR = ("rechaz", "declin", "cancel", "pendiente", "fallid", "error", "expir")
# Estatus / tipo que indican salida de dinero (reembolso o contracargo).
_PW_SALIDA = ("reembols", "devol", "contracargo", "chargeback", "refund", "reverso")


def _find_col_exact_first(cols: list[str], needles: tuple[str, ...]) -> str | None:
    """Columna cuyo nombre normalizado ES una aguja (exacto) o la CONTIENE.

    Prioridad: coincidencia exacta con la primera aguja que la tenga; luego
    contención en el orden de las agujas. Evita que "montoneto" gane la
    búsqueda de "monto" cuando existe la columna "monto" a secas.
    """
    normed = [(c, _norm(c)) for c in cols]
    for needle in needles:
        for c, n in normed:
            if n == needle:
                return c
    for needle in needles:
        for c, n in normed:
            if needle in n:
                return c
    return None


def _detectar_columnas_paywise(cols: list[str]) -> dict[str, str | None] | None:
    """Mapa {fecha, bruto, comision, neto, referencia, estatus, descripcion,
    tarjeta, tipo} si los encabezados parecen de Paywise; None si no.

    Criterio: hay fecha Y comisión Y (bruto O neto). Sin comisión no es un
    estado de cuenta de pasarela (un banco no la desglosa por fila)."""
    if _find_col(cols, *_PW_BANCO_CARGO) and _find_col(cols, *_PW_BANCO_ABONO):
        return None
    fecha = _find_col_exact_first(cols, _PW_FECHA)
    comision = _find_col_exact_first(cols, _PW_COMISION)
    neto = _find_col_exact_first(cols, _PW_NETO)
    # El bruto no debe ser la misma columna que neto/comisión: se excluyen.
    restantes = [c for c in cols if c not in (neto, comision)]
    bruto = _find_col_exact_first(restantes, _PW_BRUTO)
    if not fecha or not comision or not (bruto or neto):
        return None
    otros = [c for c in cols if c not in (fecha, comision, neto, bruto)]
    return {
        "fecha": fecha,
        "bruto": bruto,
        "comision": comision,
        "neto": neto,
        "referencia": _find_col_exact_first(otros, _PW_REF)
        or next((c for c in otros if _norm(c) in _PW_REF_EXACTO), None),
        "estatus": _find_col_exact_first(otros, _PW_ESTATUS),
        "descripcion": _find_col_exact_first(otros, _PW_DESC),
        "tarjeta": _find_col_exact_first(otros, _PW_TARJETA),
        "tipo": _find_col_exact_first(otros, _PW_TIPO),
    }


def _mapeo_a_columnas(mapeo: MapeoColumnasPaywise, cols: list[str]) -> dict[str, str | None]:
    """Mapeo manual → columnas reales (tolerante a mayúsculas/espacios)."""
    por_norm = {_norm(c): c for c in cols}

    def col(nombre: str | None) -> str | None:
        if not nombre:
            return None
        if nombre in cols:
            return nombre
        return por_norm.get(_norm(nombre))

    pedidas = {
        "fecha": mapeo.fecha,
        "bruto": mapeo.bruto,
        "comision": mapeo.comision,
        "neto": mapeo.neto,
        "referencia": mapeo.referencia,
        "estatus": mapeo.estatus,
        "descripcion": mapeo.descripcion,
    }
    out: dict[str, str | None] = {k: col(v) for k, v in pedidas.items()}
    faltan = [v for k, v in pedidas.items() if v and not out[k]]
    if faltan:
        raise ValueError(
            "Columnas del mapeo que no existen en el archivo: " + ", ".join(faltan)
        )
    if not out["bruto"] and not out["neto"]:
        raise ValueError("El mapeo necesita al menos la columna de bruto o la de neto.")
    out["tarjeta"] = None
    out["tipo"] = None
    return out


def _celda(row, cols_map: dict[str, str | None], k: str) -> str:
    c = cols_map.get(k)
    if not c:
        return ""
    v = row[c]
    txt = "" if v is None else str(v).strip()
    return "" if txt.lower() in ("nan", "none") else txt


def _parse_paywise(df, cols_map: dict[str, str | None]) -> ConciliacionParseResponse:
    movimientos: list[MovimientoParseado] = []
    omitidos = 0
    sin_fecha = 0
    derivados = 0
    for _, row in df.iterrows():

        def get(k: str, _row=row) -> str:
            return _celda(_row, cols_map, k)

        estatus = get("estatus") or None
        tipo_txt = get("tipo")
        est_n = _norm(estatus or "")
        tipo_n = _norm(tipo_txt)
        if est_n and any(x in est_n for x in _PW_ESTATUS_OMITIR):
            omitidos += 1
            continue
        bruto = _to_float(get("bruto")) if cols_map.get("bruto") else None
        comision = _to_float(get("comision")) if cols_map.get("comision") else None
        neto = _to_float(get("neto")) if cols_map.get("neto") else None
        # Completar por diferencia: neto = bruto − comisión (y viceversa).
        if neto is None and bruto is not None:
            neto = bruto - (comision or 0.0)
            derivados += 1
        if bruto is None and neto is not None:
            bruto = neto + (comision or 0.0)
        if comision is None and bruto is not None and neto is not None:
            comision = bruto - neto
        if neto is None or bruto is None:
            continue
        fecha = _fecha_de(row[cols_map["fecha"]])
        if fecha is None:
            sin_fecha += 1
        salida = (
            neto < 0
            or any(x in est_n for x in _PW_SALIDA)
            or any(x in tipo_n for x in _PW_SALIDA)
        )
        neto_abs = round(abs(neto), 2)
        if neto_abs == 0:
            continue
        ref = get("referencia") or None
        tarjeta = get("tarjeta")
        partes = [p for p in (get("descripcion"), tipo_txt or None) if p]
        # Solo si de verdad son dígitos ("**** 4242"): una columna "Tipo de
        # tarjeta" (Crédito) no debe colarse como "tarjeta dito".
        if tarjeta and tarjeta[-4:].isdigit():
            partes.append(f"tarjeta {tarjeta[-4:]}")
        if ref:
            partes.append(f"ref {ref}")
        desc = " · ".join(partes) if partes else ("Paywise reembolso" if salida else "Paywise")
        movimientos.append(
            MovimientoParseado(
                fecha=fecha,
                descripcion=desc,
                monto=neto_abs,
                tipo="CARGO" if salida else "ABONO",
                referencia=ref[:120] if ref else None,
                monto_bruto=round(abs(bruto), 2),
                comision=round(abs(comision or 0.0), 2),
                estatus=estatus,
            )
        )
    notas: list[str] = []
    if not movimientos:
        notas.append("No se encontraron operaciones con monto en el archivo de Paywise.")
    if omitidos:
        notas.append(
            f"{omitidos} operación(es) omitida(s) por estatus (rechazada/cancelada/pendiente)."
        )
    if sin_fecha:
        notas.append(f"{sin_fecha} operación(es) sin fecha legible (no se importan).")
    if derivados:
        notas.append(
            f"{derivados} neto(s) calculado(s) como bruto − comisión (el archivo no trae neto)."
        )
    return ConciliacionParseResponse(
        movimientos=movimientos,
        total=len(movimientos),
        formato="paywise",
        notas=" ".join(notas),
        columnas=[str(c) for c in df.columns],
    )


def _parse_tabular(req: ConciliacionParseRequest, formato: str) -> ConciliacionParseResponse:
    import pandas as pd  # lazy

    df = _leer_tabla(req, formato)
    cols = [str(c) for c in df.columns]

    # PAYWISE: mapeo manual del panel manda; si no, detección por encabezados.
    if req.mapeo is not None:
        return _parse_paywise(df, _mapeo_a_columnas(req.mapeo, cols))
    detectadas = _detectar_columnas_paywise(cols)
    if detectadas is not None:
        return _parse_paywise(df, detectadas)

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
    if movimientos and "paywise" in req.filename.lower():
        notas = (
            "El archivo parece de Paywise pero no se reconocieron sus columnas "
            "(bruto/comisión/neto): se importó como banco genérico. Usa el mapeo "
            "manual de columnas para leerlo como Paywise."
        )
    return ConciliacionParseResponse(
        movimientos=movimientos, total=len(movimientos), formato=formato, notas=notas,
        columnas=cols,
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
    uso = uso_ia_de(resp)
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
        movimientos=movimientos,
        total=len(movimientos),
        formato="pdf",
        modelo=s.anthropic_model,
        uso_ia=uso,
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
    uso = uso_ia_de(resp)
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
        uso_ia=uso,
    )
