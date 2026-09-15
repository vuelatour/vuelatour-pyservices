import base64
import csv
import io
import json
import re
import unicodedata
from collections import Counter
from functools import lru_cache

from app.config import get_settings
from app.schemas.conciliacion import (
    ConciliacionParseRequest,
    ConciliacionParseResponse,
    ConciliacionSugerirRequest,
    ConciliacionSugerirResponse,
    GastoCandidato,
    MapeoColumnasPaywise,
    MovimientoParseado,
    MovimientoSinConciliar,
    SugerenciaAlternativa,
)
from app.services._dominio import TC_USD_MXN_MAX, TC_USD_MXN_MIN, sistema_con_dominio
from app.services.ia_usage import uso_ia_de
from app.services.validaciones_ia import (
    confianza_de,
    dias_entre,
    normalizar_texto_banco,
    terminacion_4,
    terminacion_de_referencia,
)


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
    s = str(v).strip().replace("$", "").replace(" ", "").replace(",", "")
    if s in ("", "-", "nan", "none"):
        return None
    # Contabilidad mexicana: los cargos a veces vienen entre paréntesis.
    negativo = s.startswith("(") and s.endswith(")")
    if negativo:
        s = s[1:-1]
    try:
        valor = float(s)
    except ValueError:
        return None
    return -valor if negativo else valor


# --- Lectura tolerante de CSV/XLSX (15-sep-2026) ---------------------------
# Los exports reales de Scotiabank/BBVA traen PREÁMBULO (número de cuenta,
# periodo, saldo inicial) antes de los encabezados, llegan en latin-1 y a
# veces separados por ';'. Antes cualquiera de las tres cosas rompía el parse
# entero (UnicodeDecodeError críptico o cero movimientos sin explicación).

_ENCODINGS = ("utf-8-sig", "utf-8", "latin-1")
_DELIMITADORES = (",", ";", "\t", "|")
# Agujas para reconocer la FILA de encabezados dentro del archivo.
_AGUJAS_FECHA = ("fecha", "date")
_AGUJAS_MONTO = (
    "monto", "importe", "amount", "cargo", "abono", "retiro", "deposito",
    "debito", "credito", "debit", "credit", "saldo", "balance",
)


def _decodificar(raw: bytes) -> str:
    for enc in _ENCODINGS:
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    # Último recurso: nunca reventar por un byte suelto.
    return raw.decode("latin-1", errors="replace")


def _delimitador(texto: str) -> str:
    """Separador más CONSISTENTE de las primeras líneas.

    No gana el que más columnas produce (una coma dentro de "1,234.56"
    ganaría en un archivo separado por ';'), sino aquel cuyo número de
    columnas se repite más veces."""
    muestra = "\n".join(texto.splitlines()[:25])
    mejor, mejor_score = ",", (False, 0, 0)
    for d in _DELIMITADORES:
        try:
            filas = list(csv.reader(io.StringIO(muestra), delimiter=d))
        except csv.Error:
            continue
        anchos = Counter(len(f) for f in filas if any(c.strip() for c in f))
        if not anchos:
            continue
        ancho, veces = anchos.most_common(1)[0]
        score = (ancho > 1, veces, ancho)
        if score > mejor_score:
            mejor, mejor_score = d, score
    return mejor


def _filas_crudas(req: ConciliacionParseRequest, formato: str) -> list[list[str]]:
    """Todas las filas del archivo como texto, SIN suponer encabezado."""
    raw = base64.b64decode(req.file_base64)
    if formato == "csv":
        texto = _decodificar(raw)
        lector = csv.reader(io.StringIO(texto), delimiter=_delimitador(texto))
        return [[("" if c is None else str(c).strip()) for c in fila] for fila in lector]

    import pandas as pd  # lazy

    df = pd.read_excel(io.BytesIO(raw), dtype=str, header=None).fillna("")
    return [[("" if c is None else str(c).strip()) for c in fila] for fila in df.values.tolist()]


def _fila_encabezado(filas: list[list[str]]) -> int:
    """Índice de la fila que trae los ENCABEZADOS (0 si no hay preámbulo)."""
    for i, fila in enumerate(filas[:25]):
        normal = [_norm(c) for c in fila]
        con_texto = sum(1 for n in normal if n)
        if con_texto < 2:
            continue
        tiene_fecha = any(any(a in n for a in _AGUJAS_FECHA) for n in normal)
        tiene_monto = any(any(a in n for a in _AGUJAS_MONTO) for n in normal)
        if tiene_fecha and tiene_monto:
            return i
    return 0


def _nombres_columnas(fila: list[str]) -> list[str]:
    """Nombres únicos y no vacíos (pandas los necesita para indexar por nombre)."""
    out: list[str] = []
    vistos: dict[str, int] = {}
    for i, valor in enumerate(fila):
        nombre = str(valor).strip()
        if not nombre or nombre.lower() in ("nan", "none"):
            nombre = f"columna_{i + 1}"
        if nombre in vistos:
            vistos[nombre] += 1
            nombre = f"{nombre}_{vistos[nombre]}"
        else:
            vistos[nombre] = 1
        out.append(nombre)
    return out


def _leer_tabla(req: ConciliacionParseRequest, formato: str):
    """DataFrame de texto del CSV/XLSX (celdas vacías = ""), saltando el
    preámbulo del banco si lo hay."""
    import pandas as pd  # lazy

    filas = _filas_crudas(req, formato)
    if not filas:
        return pd.DataFrame()
    h = _fila_encabezado(filas)
    columnas = _nombres_columnas(filas[h])
    ancho = len(columnas)
    datos = [
        (fila + [""] * ancho)[:ancho]
        for fila in filas[h + 1 :]
        if any(str(c).strip() for c in fila)
    ]
    return pd.DataFrame(datos, columns=columnas, dtype=str).fillna("")


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
    # La REFERENCIA tiene columna propia (15-sep-2026): ahí viaja la
    # terminación de la tarjeta ('0025830577' ⇒ 0577), que es lo que desempata
    # dos cargos del mismo monto el mismo día. Antes se tiraba (ref = None).
    col_ref = _find_col(
        cols,
        "referencia", "autorizacion", "clavederastreo", "folio",
        "numerodeoperacion", "idoperacion", "numerodemovimiento",
    )
    otros = [c for c in cols if c != col_ref]
    # La descripción NO compite con la referencia; si el archivo solo trae
    # "Referencia" como texto, esa misma columna hace de descripción (como
    # antes de separarlas).
    col_desc = (
        _find_col(otros, "concepto", "descrip", "detalle", "movimiento", "beneficiario")
        or col_ref
    )
    col_cargo = _find_col(cols, "cargo", "retiro", "debito", "debit")
    col_abono = _find_col(cols, "abono", "deposito", "credito", "credit")
    col_monto = _find_col(cols, "monto", "importe", "amount")
    col_saldo = _find_col(cols, "saldo", "balance")

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
        ref_txt = str(row[col_ref]).strip() if col_ref else ""
        ref = ref_txt[:120] if ref_txt and ref_txt.lower() not in ("nan", "none") else None
        saldo = _to_float(row[col_saldo]) if col_saldo else None

        cargo = _to_float(row[col_cargo]) if col_cargo else None
        abono = _to_float(row[col_abono]) if col_abono else None
        if cargo and cargo != 0:
            movimientos.append(
                MovimientoParseado(
                    fecha=fecha, descripcion=desc, monto=abs(cargo), tipo="CARGO",
                    referencia=ref, saldo_posterior=saldo,
                )
            )
        elif abono and abono != 0:
            movimientos.append(
                MovimientoParseado(
                    fecha=fecha, descripcion=desc, monto=abs(abono), tipo="ABONO",
                    referencia=ref, saldo_posterior=saldo,
                )
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
                        saldo_posterior=saldo,
                    )
                )

    notas = (
        ""
        if movimientos
        else "No se reconocieron columnas de monto. Revisa el formato del archivo."
    )
    if movimientos and "paywise" in req.filename.lower():
        notas = (
            "El archivo parece de Paywise pero no se reconocieron sus columnas "
            "(bruto/comisión/neto): se importó como banco genérico. Usa el mapeo "
            "manual de columnas para leerlo como Paywise."
        )
    return ConciliacionParseResponse(
        movimientos=movimientos, total=len(movimientos), formato=formato, notas=notas,
        columnas=cols,
        advertencias=_advertencias_saldos(movimientos),
    )


def _advertencias_saldos(movimientos: list[MovimientoParseado]) -> list[str]:
    """Cadena de saldos: saldo[i] = saldo[i−1] − cargo + abono (±0.01).

    Si no cuadra, entre esas dos líneas FALTA un movimiento (extracción
    incompleta) o sobra uno. Es el detector más barato que existe de un
    estado de cuenta importado a medias; solo corre si el archivo trae saldo.
    """
    con_saldo = [m for m in movimientos if m.saldo_posterior is not None]
    if len(con_saldo) < 2:
        return []
    rotos: list[str] = []
    previo = con_saldo[0]
    for m in con_saldo[1:]:
        delta = -m.monto if m.tipo == "CARGO" else m.monto
        esperado = round((previo.saldo_posterior or 0.0) + delta, 2)
        if abs(esperado - (m.saldo_posterior or 0.0)) > 0.01:
            rotos.append(
                f"{previo.fecha or '?'} → {m.fecha or '?'} "
                f"({(m.descripcion or '')[:40]})"
            )
        previo = m
    if not rotos:
        return []
    detalle = "; ".join(rotos[:3]) + (f"; y {len(rotos) - 3} más" if len(rotos) > 3 else "")
    return [
        "El saldo corrido del archivo no cuadra en "
        f"{len(rotos)} punto(s): pueden faltar movimientos entre esas líneas "
        f"({detalle}). Revisa el archivo antes de dar la cuenta por conciliada."
    ]


_SYSTEM = (
    "Eres un TRANSCRIPTOR de estados de cuenta bancarios mexicanos en PDF. No "
    "resumes, no interpretas, no traduces: transcribes renglón por renglón.\n\n"
    "POR QUÉ IMPORTA LA LITERALIDAD: la descripción que devuelves es la llave "
    "con la que el sistema detecta si un archivo YA se importó. Si parafraseas, "
    "el mismo PDF subido dos veces entra DUPLICADO y las cuentas del mes se "
    "rompen. Copia la descripción EXACTA que imprime el banco (mayúsculas, "
    "asteriscos, abreviaturas y números incluidos); no la acortes ni la "
    "\"limpies\". Si es larguísima, córtala al final con los primeros 80 "
    "caracteres, nunca la reescribas.\n\n"
    "SALIDA: SOLO un objeto JSON, sin texto adicional ni ```fences```, COMPACTO "
    "(sin espacios ni saltos de línea innecesarios), con estas claves:\n"
    '  "movimientos": arreglo, en el MISMO ORDEN del estado de cuenta, de objetos:\n'
    '      "fecha": "YYYY-MM-DD". El PDF la imprime DD/MM/AAAA («07/09/2026» = 7 '
    "de septiembre) o solo día y mes («07 SEP», «07-SEP»): en ese caso toma el "
    "AÑO del periodo del encabezado. null solo si de plano no hay fecha.\n"
    '      "descripcion": el concepto LITERAL del banco (ver arriba).\n'
    '      "monto": número POSITIVO, siempre (el signo lo dice "tipo"). Los '
    "montos vienen con coma de miles: «1,234.56» son 1234.56. Un monto entre "
    "paréntesis o con signo menos es un CARGO con monto positivo.\n"
    '      "tipo": "CARGO" si el dinero SALE de la cuenta (retiro, compra, pago, '
    'columna «Cargos»/«Retiros»/«Débito»), "ABONO" si ENTRA (depósito, '
    "transferencia recibida, columna «Abonos»/«Depósitos»/«Crédito»).\n"
    '      "referencia": la referencia/autorización del renglón, ÍNTEGRA y tal '
    "cual (no la recortes: sus últimos 4 dígitos suelen ser la terminación de "
    "la tarjeta con la que se pagó, y con eso se identifica quién gastó). Si el "
    "renglón no trae, null.\n"
    '      "saldo_posterior": el SALDO que imprime el banco a la derecha de ese '
    "renglón (número, puede ser negativo), o null si esa columna no existe. NO "
    "lo calcules tú: cópialo.\n"
    '  "periodo_inicio" / "periodo_fin": "YYYY-MM-DD" del periodo del '
    "encabezado, o null.\n"
    '  "saldo_inicial" / "saldo_final": números del encabezado/resumen, o null.\n'
    '  "total_cargos" / "total_abonos": totales del resumen del estado de cuenta '
    "(números positivos), o null si el PDF no los imprime.\n"
    '  "advertencias": arreglo de strings en español con lo que te haya '
    "generado duda (páginas ilegibles, renglones cortados, movimientos en otra "
    "moneda, columnas ambiguas). Vacío si todo se leyó limpio.\n\n"
    "FORMATOS FRECUENTES (usa el que reconozcas; si no, aplica el criterio "
    "general de arriba):\n"
    "  · SCOTIABANK: FECHA · CONCEPTO/DESCRIPCIÓN · REFERENCIA (o SUCURSAL) · "
    "RETIROS/CARGOS · DEPÓSITOS/ABONOS · SALDO. Las compras con tarjeta traen "
    "la terminación dentro de la referencia numérica larga.\n"
    "  · BBVA: OPER · LIQ · COD. DESCRIPCIÓN · CARGOS · ABONOS · SALDO. Usa la "
    "fecha de OPERACIÓN (la primera), no la de liquidación, y copia el código y "
    "la descripción juntos tal como se imprimen.\n"
    "  · BANORTE: FECHA · DESCRIPCIÓN/ESTABLECIMIENTO · MONTO DEL DEPÓSITO · "
    "MONTO DEL RETIRO · SALDO.\n"
    "  · SANTANDER: FECHA · FOLIO · DESCRIPCIÓN · DEPÓSITOS · RETIROS · SALDO. "
    "El FOLIO es la referencia.\n"
    "  · Una sola columna de IMPORTE con signo: negativo = CARGO, positivo = "
    "ABONO, y el monto va en positivo con su tipo.\n\n"
    "REGLAS DURAS:\n"
    "  1. Incluye TODAS las filas de movimiento de TODAS las páginas, en orden. "
    "No omitas ninguna por parecer repetida: dos cargos idénticos el mismo día "
    "son normales y los dos deben aparecer.\n"
    "  2. NO incluyas encabezados, subtotales, «SALDO ANTERIOR», «SALDO FINAL», "
    "leyendas legales ni publicidad: esos números van en las claves del "
    "encabezado, no en \"movimientos\".\n"
    "  3. No inventes NADA. Si un renglón está cortado o ilegible, omítelo y "
    'dilo en "advertencias" con la fecha y lo que alcances a leer.\n'
    "  4. Si un movimiento está en otra moneda que la cuenta, transcríbelo con "
    'el monto que aparece y menciónalo en "advertencias".'
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


def _prompt_pdf(req: ConciliacionParseRequest) -> str:
    partes = ["Transcribe TODOS los movimientos y responde con el JSON indicado."]
    if req.banco:
        partes.append(
            f"El estado de cuenta es de {req.banco}: usa su formato de columnas."
        )
    if req.cuenta_moneda:
        partes.append(
            f"La cuenta es en {req.cuenta_moneda}; si ves un movimiento en otra "
            "moneda, transcríbelo y anótalo en advertencias."
        )
    partes.append(
        "Recuerda: descripción LITERAL, referencia íntegra, saldo copiado (no "
        "calculado) y ninguna fila de más ni de menos."
    )
    return " ".join(partes)


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
                            "data": req.file_base64,
                        },
                    },
                    {"type": "text", "text": _prompt_pdf(req)},
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
                referencia=str(raw["referencia"])[:120] if raw.get("referencia") else None,
                saldo_posterior=_to_float(raw.get("saldo_posterior")),
            )
        )
    advertencias = _advertencias_pdf(data, movimientos)
    return ConciliacionParseResponse(
        movimientos=movimientos,
        total=len(movimientos),
        formato="pdf",
        modelo=s.anthropic_model,
        uso_ia=uso,
        notas=(
            "La IA leyó el PDF: revisa las advertencias antes de conciliar."
            if advertencias
            else ""
        ),
        advertencias=advertencias,
    )


def _advertencias_pdf(data: dict, movimientos: list[MovimientoParseado]) -> list[str]:
    """Comprobaciones DETERMINISTAS sobre lo que devolvió el modelo.

    Una extracción con IA puede saltarse renglones sin truncarse (el
    `stop_reason == max_tokens` no lo detecta). Tres redes: la cadena de
    saldos, los totales impresos y el periodo del documento."""
    avisos: list[str] = list(_advertencias_saldos(movimientos))

    # Totales del resumen contra lo transcrito.
    for clave, tipo, etiqueta in (
        ("total_cargos", "CARGO", "cargos"),
        ("total_abonos", "ABONO", "abonos"),
    ):
        impreso = _to_float(data.get(clave))
        if impreso is None or impreso == 0:
            continue
        suma = round(sum(m.monto for m in movimientos if m.tipo == tipo), 2)
        if abs(abs(impreso) - suma) > 0.5:
            avisos.append(
                f"Los {etiqueta} transcritos suman {suma:,.2f} y el estado de cuenta "
                f"dice {abs(impreso):,.2f}: falta(n) movimiento(s) por leer. "
                f"{_USA_CSV}"
            )

    # Fechas fuera del periodo del documento.
    ini, fin = data.get("periodo_inicio"), data.get("periodo_fin")
    if isinstance(ini, str) and isinstance(fin, str) and len(ini) == 10 and len(fin) == 10:
        fuera = [m.fecha for m in movimientos if m.fecha and not (ini <= m.fecha <= fin)]
        if fuera:
            avisos.append(
                f"{len(fuera)} movimiento(s) con fecha fuera del periodo {ini}…{fin} "
                f"(ej. {fuera[0]}): revisa el año que leyó la IA."
            )

    sin_fecha = sum(1 for m in movimientos if not m.fecha)
    if sin_fecha:
        avisos.append(f"{sin_fecha} movimiento(s) sin fecha legible.")

    for extra in data.get("advertencias", []) or []:
        texto = str(extra).strip()
        if texto and texto not in avisos:
            avisos.append(texto[:300])
    return avisos[:10]


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
    "Eres el asistente de CONCILIACIÓN BANCARIA de la operadora. Recibes UN "
    "movimiento del banco sin conciliar y una lista de GASTOS CANDIDATOS que el "
    "sistema ya filtró por fecha, monto y moneda. Tu trabajo es decir cuál de "
    "esos gastos pagó ese movimiento — o decir claramente que ninguno.\n\n"
    "LA IA PROPONE, LA PERSONA CONFIRMA: nada de lo que digas se liga solo. "
    "Por eso vale mucho más una respuesta honesta («no encaja ninguno») que un "
    "match forzado: un gasto ligado al cargo equivocado descuadra el cierre del "
    "mes y nadie lo nota hasta el reparto de utilidades.\n\n"
    "JERARQUÍA DE EVIDENCIA (de más a menos):\n"
    "  1. MONTO EXACTO (hasta el centavo) — evidencia FUERTE.\n"
    "  2. TERMINACIÓN DE TARJETA: si el movimiento trae "
    "`terminacion_tarjeta_detectada` y es IGUAL a `tarjeta_terminacion` del "
    "gasto, es evidencia FUERTE (identifica a la persona que pagó). Si son "
    "DISTINTAS, es evidencia FUERTE EN CONTRA: no lo elijas aunque el monto "
    "cuadre, porque lo pagó otra tarjeta.\n"
    "  3. DESCRIPCIÓN DEL BANCO contra `proveedor`, `lugar` y `nota` del gasto "
    "(la nota es la primera línea que escribió quien capturó). Usa los alias "
    "del contexto de dominio: «ASUR CANCUN» ≈ «Aeropuerto de Cancún», «A I DE "
    "CHETUMAL» ≈ «Aeropuerto de Chetumal», «ASA MERIDA» ≈ combustible en "
    "Mérida, «MERPAGO*TAQUERIA» ≈ una comida. Evidencia MEDIA.\n"
    "  4. FECHA: sola es evidencia DÉBIL. El cargo del banco aparece el MISMO "
    "día del gasto o hasta 3 días DESPUÉS (las tarjetas tardan en reflejar). "
    "Un cargo ANTERIOR al gasto solo se explica con un pago anticipado: "
    "sospéchalo y dilo.\n\n"
    "CASOS ESPECIALES:\n"
    "  · PAGOS PARCIALES: si un candidato trae `faltante` mayor que 0, ese "
    "gasto ya está pagado a medias; compara el movimiento contra el FALTANTE, "
    "no contra el monto total.\n"
    "  · MONEDA CRUZADA: un cargo en MXN puede pagar un gasto en USD. Solo es "
    "creíble si el `tc_implicito` que trae el candidato es un tipo de cambio "
    f"plausible ({TC_USD_MXN_MIN:.0f}–{TC_USD_MXN_MAX:.0f} MXN por USD).\n"
    "  · DIFERENCIAS DE MONTO: en un CARGO de tarjeta NO hay comisión que "
    "explique una diferencia — si no cuadra al centavo, di que no cuadra. Una "
    "diferencia sí es normal en un ABONO de pasarela (Paywise retiene su "
    "comisión) o en propinas de terminal.\n"
    "  · Si el movimiento es un ABONO (entra dinero), casi nunca corresponde a "
    "un gasto: dilo en la razón y baja mucho la confianza.\n\n"
    "SALIDA: SOLO un objeto JSON, sin texto adicional ni ```fences```, con las "
    "claves exactas:\n"
    '  "gasto_id_sugerido": el "id" EXACTO de un gasto de la lista, o null. '
    "NUNCA inventes un id ni lo modifiques.\n"
    '  "confianza": número 0..1, calibrado así: 0.9–1.0 monto exacto MÁS '
    "(tarjeta igual o proveedor/lugar que empata); 0.7–0.89 monto exacto sin "
    "más señales, o tarjeta igual con monto muy cercano; 0.4–0.69 solo "
    "coincidencia de descripción o de fecha; menos de 0.4 corazonada. Si no "
    "llegas a 0.4, mejor devuelve null.\n"
    '  "razon": UNA frase en español, concreta, para que el operador decida.\n'
    '  "evidencias": arreglo de strings con los HECHOS que citas, uno por '
    'hecho (ej. "monto exacto 125.82", "terminación 0577 = tarjeta del gasto", '
    '"AEROPUERTO DE COZUMEL ≈ lugar «Aeropuerto de Cozumel»", "el cargo es 1 '
    'día posterior al gasto"). Si no puedes citar ningún hecho, deja el '
    "arreglo vacío y devuelve null en gasto_id_sugerido.\n"
    '  "alternativas": hasta 3 objetos {"gasto_id", "confianza", "razon"} con '
    "las siguientes opciones más probables (solo ids de la lista), de más a "
    "menos probable. Vacío si no hay.\n"
    '  "motivo_sin_match": cuando gasto_id_sugerido es null, UNA frase que '
    "diga por qué (ej. «ningún candidato coincide en monto ni en tarjeta», «el "
    "único monto que cuadra es de otra tarjeta», «parece una comisión del "
    "banco, no un gasto capturado»). null si sí hay sugerencia."
)

# Palabras que no distinguen nada al comparar descripción del banco vs gasto.
_VACIAS = {
    "DE", "DEL", "LA", "EL", "LOS", "LAS", "SA", "CV", "SAPI", "SC", "SRL",
    "PAGO", "COMPRA", "SERVICIO", "SERVICIOS", "TARJETA", "MEXICO", "CANCUN",
}


def _tokens(texto: str | None) -> set[str]:
    return {
        t for t in normalizar_texto_banco(texto).split() if len(t) >= 4 and t not in _VACIAS
    }


def _payload_movimiento(mov: MovimientoSinConciliar, terminacion: str | None) -> dict:
    """Lo que ve el modelo del movimiento. Antes solo iban fecha, monto y
    descripción: sin referencia ni tarjeta era imposible desempatar."""
    return {
        "fecha": mov.fecha,
        "monto": mov.monto,
        "tipo": mov.tipo,
        "descripcion": mov.descripcion,
        "referencia": mov.referencia,
        "terminacion_tarjeta_detectada": terminacion,
        "cuenta": mov.cuenta_alias,
        "cuenta_moneda": mov.cuenta_moneda,
    }


def _payload_candidato(c: GastoCandidato) -> dict:
    """Ficha del gasto SIN campos vacíos (el JSON viaja en cada llamada)."""
    campos = {
        "id": c.id,
        "fecha": c.fecha,
        "monto": c.monto,
        "moneda": c.moneda,
        "proveedor": c.proveedor,
        "lugar": c.lugar,
        "nota": (c.nota or "")[:160] or None,
        "categoria": c.categoria,
        "medio_pago": c.medio_pago,
        "tarjeta_terminacion": c.tarjeta_terminacion,
        "matricula": c.matricula,
        "vuelo_folio": c.vuelo_folio,
        "capturado_por": c.capturado_por,
        "monto_vinculado": c.monto_vinculado,
        "faltante": c.faltante,
        "tc_implicito": c.tc_implicito,
    }
    return {k: v for k, v in campos.items() if v is not None}


def _evidencias_deterministas(
    mov: MovimientoSinConciliar, c: GastoCandidato, terminacion: str | None
) -> tuple[list[str], float]:
    """Hechos que NO dependen del modelo + tope de confianza que permiten.

    El servicio conoce los montos, las fechas y las tarjetas: no hay razón
    para creerle al modelo un 0.95 cuando la aritmética dice otra cosa. El
    tope final es el MÍNIMO de todos (cualquier señal en contra manda).
    """
    ev: list[str] = []
    topes: list[float] = []

    if mov.tipo == "ABONO":
        ev.append("El movimiento es un ABONO (entra dinero): un gasto se paga con un CARGO.")
        topes.append(0.3)

    monedas_distintas = bool(
        c.moneda and mov.cuenta_moneda and c.moneda.upper() != mov.cuenta_moneda.upper()
    )
    objetivo = c.faltante if (c.faltante is not None and 0 < c.faltante < c.monto) else c.monto
    parcial = objetivo != c.monto

    if monedas_distintas:
        tc = c.tc_implicito if c.tc_implicito else (mov.monto / c.monto if c.monto else None)
        if tc and TC_USD_MXN_MIN <= tc <= TC_USD_MXN_MAX:
            ev.append(
                f"Moneda cruzada ({c.moneda} → {mov.cuenta_moneda}) con tipo de cambio "
                f"implícito {tc:.2f}, dentro de la banda razonable."
            )
            topes.append(0.8)
        else:
            ev.append(
                f"Moneda cruzada ({c.moneda} → {mov.cuenta_moneda}) con tipo de cambio "
                f"implícito {tc:.2f} fuera de la banda "
                f"{TC_USD_MXN_MIN:.0f}–{TC_USD_MXN_MAX:.0f}."
                if tc
                else f"El gasto está en {c.moneda} y la cuenta en {mov.cuenta_moneda}."
            )
            topes.append(0.4)
    elif abs(mov.monto - objetivo) <= 0.01:
        ev.append(
            f"Cubre exactamente el faltante del gasto ({objetivo:,.2f})."
            if parcial
            else f"Monto exacto: {mov.monto:,.2f}."
        )
        topes.append(0.9 if parcial else 0.95)
    elif objetivo:
        dif = abs(mov.monto - objetivo)
        pct = dif / objetivo * 100
        etiqueta = "faltante" if parcial else "monto"
        ev.append(f"Diferencia de {dif:,.2f} ({pct:.2f} %) contra el {etiqueta} del gasto.")
        topes.append(0.7 if pct <= 0.5 else 0.5)

    term_gasto = terminacion_4(c.tarjeta_terminacion)
    if terminacion and term_gasto:
        if terminacion == term_gasto:
            ev.append(f"Terminación {terminacion}: es la tarjeta con la que se pagó el gasto.")
            topes.append(0.95)
        else:
            ev.append(
                f"La tarjeta del movimiento termina en {terminacion} y la del gasto en "
                f"{term_gasto}: lo pagó otra tarjeta."
            )
            topes.append(0.3)

    dias = dias_entre(c.fecha, mov.fecha)
    if dias is not None:
        if dias == 0:
            ev.append("Mismo día que el gasto.")
        elif 0 < dias <= 3:
            ev.append(f"El cargo aparece {dias} día(s) después del gasto (normal en tarjeta).")
        elif dias > 3:
            ev.append(f"El cargo aparece {dias} días después del gasto.")
            if dias > 7:
                topes.append(0.6)
        else:
            ev.append(
                f"El cargo es {abs(dias)} día(s) ANTERIOR al gasto: solo cuadra si se "
                "pagó por adelantado."
            )
            if dias < -1:
                topes.append(0.6)

    comunes = _tokens(mov.descripcion) & (
        _tokens(c.proveedor) | _tokens(c.lugar) | _tokens(c.nota)
    )
    if comunes:
        ev.append(
            f"La descripción del banco y el gasto comparten «{' '.join(sorted(comunes)[:3])}»."
        )

    return ev, min(topes) if topes else 0.5


def _alternativas(
    data: dict, ids_validos: set[str], elegido: str | None
) -> list[SugerenciaAlternativa]:
    """Top 3 alternativas del modelo, con ids validados y sin repetir."""
    out: list[SugerenciaAlternativa] = []
    vistos = {elegido} if elegido else set()
    for raw in data.get("alternativas", []) or []:
        if not isinstance(raw, dict) or len(out) >= 3:
            continue
        gid = raw.get("gasto_id")
        gid = str(gid) if gid is not None else None
        if gid is None or gid not in ids_validos or gid in vistos:
            continue
        vistos.add(gid)
        out.append(
            SugerenciaAlternativa(
                gasto_id=gid,
                confianza=confianza_de(raw.get("confianza")),
                razon=str(raw.get("razon", ""))[:200],
            )
        )
    return out


def sugerir_conciliacion(req: ConciliacionSugerirRequest) -> ConciliacionSugerirResponse:
    s = get_settings()
    ids_validos = {c.id for c in req.candidatos}

    if not req.candidatos:
        return ConciliacionSugerirResponse(
            gasto_id_sugerido=None,
            confianza=0.0,
            razon="No hay gastos candidatos para este movimiento.",
            motivo_sin_match=(
                "Sin candidatos: ningún gasto sin conciliar cerca en fecha, monto y moneda."
            ),
            modelo=s.anthropic_model,
        )

    # La terminación de tarjeta del movimiento: la que ya dedujo el API o la
    # que se extrae de la referencia, SOLO si empata con la tarjeta de algún
    # candidato (jamás se inventa una).
    terminaciones = [c.tarjeta_terminacion for c in req.candidatos if c.tarjeta_terminacion]
    terminacion = terminacion_4(req.movimiento.terminacion_tarjeta_detectada) or (
        terminacion_de_referencia(req.movimiento.referencia, terminaciones)
    )

    payload = json.dumps(
        {
            "movimiento": _payload_movimiento(req.movimiento, terminacion),
            "candidatos": [_payload_candidato(c) for c in req.candidatos],
        },
        ensure_ascii=False,
    )
    resp = _client().messages.create(
        model=s.anthropic_model,
        max_tokens=900,
        system=sistema_con_dominio(_SUGERIR_SYSTEM),
        messages=[
            {
                "role": "user",
                "content": (
                    "Movimiento sin conciliar y gastos candidatos (JSON):\n"
                    f"{payload}\n\n"
                    "Elige el match más probable y responde con el JSON indicado. "
                    "Cita hechos en «evidencias»; si no tienes ninguno, devuelve null."
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
    elegido = next((c for c in req.candidatos if c.id == sugerido), None)

    evidencias: list[str] = []
    confianza = 0.0
    if elegido is not None:
        duras, tope = _evidencias_deterministas(req.movimiento, elegido, terminacion)
        evidencias = list(duras)
        for extra in data.get("evidencias", []) or []:
            texto = str(extra).strip()[:200]
            if texto and texto not in evidencias:
                evidencias.append(texto)
        # La confianza del modelo NUNCA sube por encima de lo que permiten los
        # hechos comprobables (fiabilidad numérica sagrada).
        confianza = min(confianza_de(data.get("confianza")), tope)

    razon = str(data.get("razon", ""))[:300]
    if elegido is not None and not razon:
        razon = evidencias[0] if evidencias else "Coincidencia propuesta por la IA."

    motivo = data.get("motivo_sin_match")
    motivo = str(motivo)[:300] if motivo else None
    if elegido is None and not motivo:
        motivo = (
            "Ningún candidato coincide lo suficiente (revisa monto, tarjeta y fecha)."
        )

    return ConciliacionSugerirResponse(
        gasto_id_sugerido=sugerido,
        confianza=confianza,
        razon=razon,
        modelo=s.anthropic_model,
        uso_ia=uso,
        evidencias=evidencias[:8],
        alternativas=_alternativas(data, ids_validos, sugerido),
        motivo_sin_match=motivo if elegido is None else None,
    )
