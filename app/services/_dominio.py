"""Contexto de dominio compartido por TODOS los prompts de IA (15-sep-2026).

Antes cada prompt volvía a explicar (o se olvidaba de explicar) lo mismo:
la flota, los aeropuertos, el IVA, el formato de fecha mexicano. Aquí vive
UNA sola vez y se inyecta como PRIMER bloque `system` con `cache_control`
en cada llamada a Claude:

    system=sistema_con_dominio(_MI_PROMPT)

Ventajas: se actualiza en un solo lugar, el bloque es idéntico en todos los
endpoints (el caché de Anthropic lo cobra ~10 % después del primer uso) y
todos los prompts hablan el mismo idioma.

REGLA: las listas de aeropuertos/proveedores son ORIENTATIVAS (el modelo
puede devolver algo que no esté en ellas). La ÚNICA lista autoritativa es
`MATRICULAS_FLOTA`, y aun así una matrícula fuera de la lista NO se borra:
se conserva con una advertencia (la flota cambia sin que este archivo se
entere; el API tiene la última palabra). Quien quiera validar contra la
flota VIVA puede mandarla en el request (campo `matriculas_flota`).
"""

from __future__ import annotations

# Fecha de la última revisión del contenido (viaja en el prompt: si un día
# una respuesta rara huele a contexto viejo, aquí se ve de qué fecha es).
DOMINIO_VERSION = "2026-09-24"

# Flota conocida (aviones chárter). Referencia para advertir matrículas
# alucinadas; NO es un filtro que borre datos.
MATRICULAS_FLOTA: tuple[str, ...] = (
    "N4142R",
    "N990GG",
    "XB-PEV",
    "XA-VGV",
    "XB-ANU",
    "N58BT",
    "XB-IJP",
    "N621TX",
)

# Aeropuertos que opera la flota (IATA → nombre largo). Lista NO exhaustiva.
AEROPUERTOS: dict[str, str] = {
    "CUN": "Cancún",
    "CZM": "Cozumel",
    "MID": "Mérida",
    "CTM": "Chetumal",
    "CME": "Ciudad del Carmen",
    "MTT": "Minatitlán",
    "HOL": "Holbox",
    "TQO": "Tulum",
    "VSA": "Villahermosa",
    "VER": "Veracruz",
    "PBC": "Puebla",
    "MEX": "Ciudad de México",
    "TLC": "Toluca",
}

# Leyendas tal como aparecen en tickets y en el estado de cuenta del banco,
# con lo que significan. Alimenta el prompt de conciliación (desempate
# descripción del banco ↔ proveedor/lugar/nota del gasto).
ALIAS_BANCO: tuple[tuple[str, str], ...] = (
    ("ASUR", "administrador aeroportuario: aterrizaje, plataforma, TUA"),
    ("A I DE", "«Aeropuerto Internacional de …» (ej. A I DE CHETUMAL)"),
    ("AEROPUERTO DE", "cuotas del aeropuerto de esa ciudad"),
    ("ASA", "Aeropuertos y Servicios Auxiliares: combustible de aviación"),
    ("GAFSACOMM", "combustible de aviación"),
    ("GASOL CARIBE", "combustible de aviación (AVGAS/turbosina)"),
    ("REAL AERO", "combustible / servicios de aviación"),
    ("AFAC", "autoridad aeronáutica: permisos y derechos"),
    ("VIP SAESA", "servicios de pista / FBO"),
    ("GOB EDO DE QROO", "derechos estatales de Quintana Roo"),
    ("MERPAGO", "Mercado Pago: el comercio real va DESPUÉS del asterisco"),
    ("PINPE", "agregador de pago: el comercio real va después del asterisco"),
    ("CLIP", "agregador de pago (terminal): comercio real tras el asterisco"),
    ("BILLPOCKET", "agregador de pago (terminal)"),
    (
        "POCKET DE LATINOAMERICA",
        "BillPocket: depósito AGRUPADO de la terminal de cobro (ABONO neto de comisión)",
    ),
    ("PAYWISE", "pasarela de cobro de la empresa (deposita NETO, retiene comisión)"),
    ("SEL TRASPASO ENTRE CUENTAS", "traspaso interno: NO es gasto ni cobro"),
    ("COMISION", "comisión del banco (sin gasto capturado detrás)"),
    ("IVA COMISION", "IVA de la comisión del banco"),
    ("INTERES", "intereses del banco"),
    # 24-sep-2026 (conciliación de INGRESOS): lo que se ve en los ABONOS.
    ("REV", "reverso: el banco devolvió un cargo anterior (NO es ingreso ni gasto nuevo)"),
    ("SPEI", "transferencia interbancaria: suele traer el nombre del ordenante"),
)

IVA_PCT = 16.0
MONEDAS: tuple[str, ...] = ("MXN", "USD")
# Banda plausible del tipo de cambio USD→MXN (la usa el cruce implícito del
# API; aquí sirve para que el modelo no proponga conversiones absurdas).
TC_USD_MXN_MIN = 15.0
TC_USD_MXN_MAX = 25.0
ZONA_HORARIA = "hora de Cancún (UTC−5)"


def _lista_aeropuertos() -> str:
    return " · ".join(f"{iata} {nombre}" for iata, nombre in AEROPUERTOS.items())


def _lista_alias() -> str:
    return "\n".join(f"  · {patron} = {que_es}" for patron, que_es in ALIAS_BANCO)


BLOQUE_DOMINIO = (
    f"CONTEXTO DE DOMINIO — VuelaTour / Aero Charter Cancún (v{DOMINIO_VERSION})\n"
    "Operadora de vuelos chárter con base en Cancún, México. Este bloque es "
    "contexto de referencia: NO cambia el formato de salida que te pide la "
    "instrucción principal.\n\n"
    "FLOTA (matrículas conocidas): "
    + ", ".join(MATRICULAS_FLOTA)
    + ".\n"
    "Formato de matrícula: XA-/XB- (México) o N seguido de números y letras "
    "(EE. UU.). Si en el documento ves una matrícula distinta, repórtala tal "
    "cual y dilo en las notas; NUNCA inventes una matrícula ni «corrijas» la "
    "que ves para que se parezca a una de la lista.\n\n"
    "AEROPUERTOS FRECUENTES (código IATA; la lista NO es exhaustiva):\n  "
    + _lista_aeropuertos()
    + ".\n\n"
    "PROVEEDORES Y LEYENDAS FRECUENTES (así aparecen en tickets y en el "
    "estado de cuenta del banco):\n" + _lista_alias() + "\n\n"
    "DINERO: solo hay dos monedas, MXN y USD. IVA mexicano = "
    f"{IVA_PCT:.0f} %. Los cobros de aeropuerto suelen venir con IVA "
    "INCLUIDO. Los montos se imprimen con coma de miles y punto decimal: "
    "«$1,234.56» son mil doscientos treinta y cuatro pesos con 56 centavos, "
    "no 1.23. Un tipo de cambio USD→MXN plausible está entre "
    f"{TC_USD_MXN_MIN:.0f} y {TC_USD_MXN_MAX:.0f}.\n\n"
    "FECHAS: en México se imprimen DD/MM/AAAA — «07/09/2026» es el 7 de "
    "septiembre de 2026, no el 9 de julio. Devuelve SIEMPRE YYYY-MM-DD. Si "
    "el documento solo imprime día y mes («07 SEP»), toma el año del "
    "encabezado o del periodo del documento. Toda la operación se mide en "
    f"{ZONA_HORARIA}: no conviertas a UTC ni a otra zona.\n\n"
    "HONESTIDAD (regla que manda sobre todas las demás): no inventes ningún "
    "dato que no esté en el documento — usa null. Si una comprobación "
    "aritmética o de catálogo no te cuadra (los renglones no suman el total, "
    "la matrícula no parece de la flota, una fecha imposible, un dígito que "
    "no alcanzas a leer), NO entregues el dato como bueno: baja «confianza» "
    "a 0.3 o menos y explica el problema en «notas». De estos números salen "
    "el cierre mensual y el reparto de utilidades: es MUCHO peor un dato "
    "equivocado que se ve seguro, que admitir la duda."
)


def bloque_sistema_dominio() -> dict:
    """Bloque `system` con el contexto de dominio (cacheado)."""
    return {
        "type": "text",
        "text": BLOQUE_DOMINIO,
        "cache_control": {"type": "ephemeral"},
    }


def sistema_con_dominio(prompt: str) -> list[dict]:
    """`system` de dos bloques: dominio compartido + prompt del endpoint.

    El dominio va PRIMERO a propósito: es el prefijo idéntico en todas las
    llamadas del servicio, así que el caché de Anthropic lo reutiliza entre
    endpoints distintos. Ambos bloques llevan `cache_control` (2 de los 4
    puntos de corte permitidos).
    """
    return [
        bloque_sistema_dominio(),
        {"type": "text", "text": prompt, "cache_control": {"type": "ephemeral"}},
    ]
