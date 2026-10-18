"""Cierre de mes formal y vuelos que pasan al mes siguiente (16-oct-2026,
API 0.0.88, invariante 89 del API).

Pedido del cliente: «una opción para cerrar el mes … y los vuelos que aún no
se cobren se marquen en el balance y se muevan para el siguiente mes hasta
que se cobre y complete». El API decide TODO (qué vuelo pasa, a qué mes, con
qué leyenda, qué suma y qué no); aquí solo se PINTAN los campos aditivos que
manda: `cierre_marca` de cada fila («Pasa a octubre 2026 (pendiente de cobro:
cobrado $0.00 de $36,632.89 MXN)» / «Viene de septiembre 2026» / «Pendiente
de cobro: pasará a octubre 2026 al cerrar»), `fila_sombra` (fila del mes de
ORIGEN de un vuelo diferido: horas reales y dinero en null, así ninguna SUM
la cuenta) y el bloque raíz `cierre` con las listas `pasan` / `vienen` /
`pasarian` ya calculadas.

Este módulo tiene SOLO texto de presentación compartido por el balance
(Excel) y el reparto (PDF/XLSX): la etiqueta es-MX de un mes, el estado del
cierre («Mes cerrado el 16-oct-2026 por Ale») y los títulos de las listas.
Ningún número: los montos y las leyendas llegan YA armados del API.
"""

from __future__ import annotations

from datetime import date, datetime, timedelta, timezone

# Cancún no tiene horario de verano: UTC−5 fijo (CLAUDE.md del workspace).
_CANCUN = timezone(timedelta(hours=-5))

_MESES_ES = [
    "enero",
    "febrero",
    "marzo",
    "abril",
    "mayo",
    "junio",
    "julio",
    "agosto",
    "septiembre",
    "octubre",
    "noviembre",
    "diciembre",
]
_MESES_CORTOS = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]

# Colores de la marca (Excel): ámbar = sale o pasará; azul claro = viene.
FILL_SALE = "FCD34D"  # el ámbar de los libros (`AMBER` del balance)
FILL_VIENE = "DBEAFE"
COLOR_TEXTO_SALE = "92400E"
COLOR_TEXTO_VIENE = "1E3A8A"

ESTADO_CERRADO = "CERRADO"
ESTADO_REABIERTO = "REABIERTO"
ESTADO_ABIERTO = "ABIERTO"

# Motivo con que el API pasa un vuelo al mes siguiente (`MotivoDiferir`):
# COBRO_POSTERIOR = sin cobro completo al corte, pero cobrado DESPUÉS.
MOTIVO_COBRO_POSTERIOR = "COBRO_POSTERIOR"


def etiqueta_mes(mes: str | None) -> str | None:
    """'2026-10' ⇒ «octubre 2026» (la misma etiqueta que arma el API en sus
    leyendas); texto que no sea YYYY-MM se devuelve tal cual; None ⇒ None."""
    if not mes:
        return None
    partes = mes.strip().split("-")
    if len(partes) >= 2 and partes[0].isdigit() and partes[1].isdigit():
        m = int(partes[1])
        if 1 <= m <= 12:
            return f"{_MESES_ES[m - 1]} {partes[0]}"
    return mes


def fecha_corta(ts: str | None) -> str | None:
    """Sello del cierre como lo lee la oficina: '2026-10-16T19:02:11.123Z' ⇒
    «16-oct-2026» en hora de Cancún; un texto que no sea fecha se devuelve
    tal cual (el API podría mandarlo ya legible); None ⇒ None."""
    if not ts:
        return None
    try:
        dt = datetime.fromisoformat(ts.replace("Z", "+00:00"))
    except ValueError:
        return ts
    if dt.tzinfo is not None:
        dt = dt.astimezone(_CANCUN)
    return f"{dt.day:02d}-{_MESES_CORTOS[dt.month - 1]}-{dt.year}"


def texto_estado_cierre(cierre) -> str | None:
    """«Mes cerrado el 16-oct-2026 por Ale» / «Mes reabierto el 17-oct-2026
    por Ale» o None cuando el mes está abierto (o no hay dato). `cierre` es
    cualquier objeto con `estado`, `cerrado_at`, `cerrado_por_nombre`
    (el `BalanceCierre` del balance o el `RepartoCierre` del reparto)."""
    if cierre is None:
        return None
    estado = (getattr(cierre, "estado", None) or "").strip().upper()
    if estado not in (ESTADO_CERRADO, ESTADO_REABIERTO):
        return None
    verbo = "cerrado" if estado == ESTADO_CERRADO else "reabierto"
    sello = getattr(cierre, "reabierto_at", None) if estado == ESTADO_REABIERTO else None
    quien = getattr(cierre, "reabierto_por_nombre", None) if estado == ESTADO_REABIERTO else None
    sello = sello or getattr(cierre, "cerrado_at", None)
    quien = quien or getattr(cierre, "cerrado_por_nombre", None)
    texto = f"Mes {verbo}"
    fecha = fecha_corta(sello)
    if fecha:
        texto += f" el {fecha}"
    if quien and str(quien).strip():
        texto += f" por {str(quien).strip()}"
    return texto


def marca_es_entrada(v) -> bool:
    """¿La marca de la fila se pinta como ENTRADA (azul: «Viene de …»)? Todo
    lo demás (sale o pasará) va en ámbar. Una fila que vino de otro mes y
    que además pasará al cerrar este («… pasará …») va en ámbar: lo urgente
    es el cobro. `v` = fila del balance o del Libro Dinero."""
    marca = (getattr(v, "cierre_marca", None) or "").lower()
    if getattr(v, "fila_sombra", False) or getattr(v, "diferido_a", None):
        return False
    if "pasará" in marca or "pasara" in marca:
        return False
    return bool(getattr(v, "diferido_de", None)) or marca.startswith("viene")


def dia_corto(dia: str | None) -> str | None:
    """'2026-09-30' ⇒ «30-sep», '2026-10-04' ⇒ «4-oct»: el día corto de las
    leyendas del API (`diaCortoDeDia`). Lo que no sea AAAA-MM-DD ⇒ None."""
    if not isinstance(dia, str):
        return None
    t = dia.strip()
    if len(t) < 10 or t[4] != "-" or t[7] != "-":
        return None
    if not (t[:4].isdigit() and t[5:7].isdigit() and t[8:10].isdigit()):
        return None
    mes = int(t[5:7])
    if not 1 <= mes <= 12:
        return None
    return f"{int(t[8:10])}-{_MESES_CORTOS[mes - 1]}"


def _ultimo_dia_del_mes(mes: str | None, *, anterior: bool = False) -> str | None:
    """'2026-09' ⇒ '2026-09-30'; con `anterior`, el del mes de antes
    ('2026-10' ⇒ '2026-09-30'). Lo que no sea AAAA-MM ⇒ None."""
    if not isinstance(mes, str):
        return None
    partes = mes.strip().split("-")
    if len(partes) < 2 or not (partes[0].isdigit() and partes[1].isdigit()):
        return None
    anio, m = int(partes[0]), int(partes[1])
    if not 1 <= m <= 12:
        return None
    primero = date(anio, m, 1)
    if anterior:
        return (primero - timedelta(days=1)).isoformat()
    siguiente = date(anio + 1, 1, 1) if m == 12 else date(anio, m + 1, 1)
    return (siguiente - timedelta(days=1)).isoformat()


def _motivo(it) -> str | None:
    m = getattr(it, "motivo", None)
    return m.strip().upper() if isinstance(m, str) and m.strip() else None


def es_cobro_posterior(it) -> bool:
    """¿El API pasó este vuelo porque se cobró DESPUÉS del corte?"""
    return _motivo(it) == MOTIVO_COBRO_POSTERIOR


def corte_del_vuelo(it) -> str | None:
    """Día (AAAA-MM-DD) del corte con que el API midió «cobrado» para este
    vuelo, SOLO cuando el API manda su `motivo` (sin él ⇒ None y todo se
    pinta como siempre): `corte` si viene; si no, el último día del mes
    anterior a su `destino` (`corteDeDestino` del API) o, en un vuelo que
    VIENE, el último día de su mes de `origen` (los montos de esa lista son
    los del cierre de su mes)."""
    if _motivo(it) is None:
        return None
    corte = getattr(it, "corte", None)
    if dia_corto(corte):
        return corte
    return _ultimo_dia_del_mes(getattr(it, "destino", None), anterior=True) or (
        _ultimo_dia_del_mes(getattr(it, "origen", None))
    )


def primer_corte(items) -> str | None:
    """El corte de la primera línea que lo tenga (las de una lista se
    midieron al mismo corte; el encabezado lo nombra una vez)."""
    for it in items:
        c = corte_del_vuelo(it)
        if c:
            return c
    return None


def titulo_pasan(destino: str | None, corte: str | None = None) -> str:
    """Encabezado de la lista de vuelos que SALEN del mes (no suman). Con el
    corte (API que manda el motivo): «sin cobro completo al 30-sep» — un
    vuelo cobrado DESPUÉS del corte también pasa, y «pendientes de cobro» se
    leía como «no se han cobrado» (caso #351, 16-oct-2026)."""
    mes = etiqueta_mes(destino)
    dia = dia_corto(corte)
    if dia:
        return (
            f"Pasan a {mes} (sin cobro completo al {dia}: no suman en este mes)"
            if mes
            else f"Pasan al mes siguiente (sin cobro completo al {dia}: no suman en este mes)"
        )
    return (
        f"Pasan a {mes} (pendientes de cobro al cierre: no suman en este mes)"
        if mes
        else "Pasan al mes siguiente (pendientes de cobro al cierre: no suman en este mes)"
    )


def titulo_vienen(origen: str | None, corte: str | None = None) -> str:
    """Encabezado de la lista de vuelos que ENTRAN al mes (suman completos).
    `corte` se acepta por simetría y no cambia el texto."""
    mes = etiqueta_mes(origen)
    return (
        f"Vienen de {mes} (suman completos en este mes; sus horas quedaron en su mes real)"
        if mes
        else "Vienen de un mes anterior (suman completos en este mes; sus horas quedaron "
        "en su mes real)"
    )


def titulo_pasarian(destino: str | None, corte: str | None = None) -> str:
    """Encabezado de la lista de un mes ABIERTO: hoy suman; al cerrar pasan.
    Con el corte, sin «Pendientes de cobro:» al frente (la lista también
    trae los cobrados después del corte, que igual pasan)."""
    mes = etiqueta_mes(destino)
    dia = dia_corto(corte)
    if dia:
        return (
            f"Pasarán a {mes} al cerrar el mes: sin cobro completo al {dia} (hoy sí suman)"
            if mes
            else f"Pasarán al mes siguiente al cerrar el mes: sin cobro completo al {dia} "
            "(hoy sí suman)"
        )
    return (
        f"Pendientes de cobro: pasarán a {mes} al cerrar el mes (hoy sí suman)"
        if mes
        else "Pendientes de cobro: pasarán al mes siguiente al cerrar el mes (hoy sí suman)"
    )


def _pesos(x: float) -> str:
    return f"${x:,.2f}"


def nombre_vuelo(it) -> str:
    """La CLAVE que arma el API («#316 · Mike Nelson»); sin ella, «#folio ·
    cliente» con lo que venga."""
    if it.clave:
        return it.clave
    nombre = f"#{it.folio}" if it.folio else "—"
    return f"{nombre} · {it.cliente}" if it.cliente else nombre


def detalle_cobro(it, *, cobrado, por_cobrar, despues, fecha_despues, dinero) -> list[str]:
    """Partes «cobrado …» de un vuelo de las listas del cierre (Libro Dinero
    y reparto PDF); `dinero` formatea un monto en la moneda de quien llama.

    Sin `motivo` (API previo): «cobrado $X», «por cobrar $Y», byte a byte
    las de siempre. Con él, los montos van AL CORTE y se dice lo cobrado
    después (caso #351, 16-oct-2026: «aparece que aún no se cobran pero sí
    están cobrados»):
      - COBRO_POSTERIOR: «cobrado al 30-sep $0.00» + «cobrado después
        $2,076.40 el 14-oct» (sin el monto: «cobrado el 14-oct, después del
        30-sep» o «cobrado después del 30-sep» — un mes ya cerrado solo
        guardó el motivo), sin «por cobrar».
      - SIN_COBRO / COBRO_PARCIAL: «cobrado al 30-sep $X», «por cobrar al
        30-sep $Y» y, si ya entró algo después, «cobrado después $Z»."""
    dia = dia_corto(corte_del_vuelo(it))
    if dia is None:
        partes = [
            f"cobrado {dinero(cobrado)}" if cobrado is not None else None,
            f"por cobrar {dinero(por_cobrar)}" if por_cobrar is not None else None,
        ]
        return [p for p in partes if p]
    partes = [f"cobrado al {dia} {dinero(cobrado)}" if cobrado is not None else None]
    cuando = dia_corto(fecha_despues)
    hay_despues = isinstance(despues, int | float) and despues > 0.005
    despues_txt = (
        f"cobrado después {dinero(despues)}" + (f" el {cuando}" if cuando else "")
        if hay_despues
        else None
    )
    if es_cobro_posterior(it):
        sin_monto = (
            f"cobrado el {cuando}, después del {dia}" if cuando else f"cobrado después del {dia}"
        )
        partes.append(despues_txt or sin_monto)
    else:
        if por_cobrar is not None:
            partes.append(f"por cobrar al {dia} {dinero(por_cobrar)}")
        partes.append(despues_txt)
    return [p for p in partes if p]


def linea_vuelo_libro(it) -> str:
    """Un vuelo de las listas del cierre en UNA línea de texto (Libro Dinero):
    «#316 · Mike Nelson — venta $36,632.89 MXN, cobrado $0.00, por cobrar
    $36,632.89». Montos del API tal cual; las partes que no vienen se
    omiten. Con el motivo del API, montos al corte y lo cobrado después
    (`detalle_cobro`)."""
    nombre = nombre_vuelo(it)
    montos = [f"venta {_pesos(it.venta_mxn)} MXN" if it.venta_mxn is not None else None]
    montos += detalle_cobro(
        it,
        cobrado=it.cobrado_mxn,
        por_cobrar=it.por_cobrar_mxn,
        despues=getattr(it, "cobrado_despues_mxn", None),
        fecha_despues=getattr(it, "fecha_cobro_despues", None),
        dinero=_pesos,
    )
    montos = [m for m in montos if m]
    return nombre + (f" — {', '.join(montos)}" if montos else "")


def listas_balance(cierre) -> list[tuple[str, str, list]]:
    """(clave de la lista, título, vuelos) de las listas del balance / Libro
    Dinero que traen vuelos, en orden: pasan, vienen, pasarían."""
    if cierre is None:
        return []
    salida = []
    for attr, titulo, mes_attr in (
        ("pasan", titulo_pasan, "destino"),
        ("vienen", titulo_vienen, "origen"),
        ("pasarian", titulo_pasarian, "destino"),
    ):
        items = getattr(cierre, attr)
        if items:
            salida.append((attr, titulo(primer_mes(items, mes_attr), primer_corte(items)), items))
    return salida


def primer_mes(items, attr: str) -> str | None:
    """El `destino`/`origen` de la primera línea que lo traiga (todas las de
    una lista apuntan al mismo mes; el encabezado lo nombra una vez)."""
    for it in items:
        v = getattr(it, attr, None)
        if v:
            return v
    return None


# ---- Reparto (PDF y Excel): mismos títulos en los dos documentos ----------
TITULO_REPARTO_SALEN = "Pasan a {mes} (pendientes de cobro, no se reparten)"
TITULO_REPARTO_ENTRAN = "Vienen de {mes} (cobrados: se reparten en este mes)"
TITULO_REPARTO_PASARIAN = "Pendientes de cobro: pasarán a {mes} al cerrar el mes"
TITULO_REPARTO_SIN_AVION = "Vuelos del cierre sin avión en este reparto"
# Con el motivo del API (caso #351): el corte en vez de «pendientes de cobro»,
# que se leía como «no se han cobrado» aunque se cobraran después del corte.
_TITULO_REPARTO_CON_CORTE = {
    TITULO_REPARTO_SALEN: "Pasan a {mes} (sin cobro completo al {dia}, no se reparten)",
    TITULO_REPARTO_PASARIAN: "Pasarán a {mes} al cerrar el mes (sin cobro completo al {dia})",
}
AVISO_SNAPSHOT = (
    "Los números de hoy no coinciden con los del cierre (cambió un porcentaje "
    "de socio o un tipo de cambio después de cerrar): reabre el mes y ciérralo "
    "de nuevo antes de dispersar."
)
# (lista de `RepartoCierre`, plantilla del título, llave del mes que nombra).
LISTAS_REPARTO = (
    ("diferidos_salen", TITULO_REPARTO_SALEN, "destino"),
    ("diferidos_entran", TITULO_REPARTO_ENTRAN, "origen"),
    ("pendientes_que_pasarian", TITULO_REPARTO_PASARIAN, "destino"),
)


def titulo_lista_reparto(plantilla: str, items, attr: str) -> str:
    """Título de una lista del reparto con el mes de su primera línea (y, si
    el API manda el motivo, el día del corte)."""
    mes = etiqueta_mes(primer_mes(items, attr)) or "otro mes"
    dia = dia_corto(primer_corte(items))
    con_corte = _TITULO_REPARTO_CON_CORTE.get(plantilla)
    if dia and con_corte:
        return con_corte.format(mes=mes, dia=dia)
    return plantilla.format(mes=mes)


def hay_motivo(items) -> bool:
    """¿Alguna línea trae el `motivo` del API? (encabezados «al cierre»)."""
    return any(_motivo(it) is not None for it in items)


def matriculas_sin_bloque(cierre, matriculas_del_reparto) -> list[str | None]:
    """Matrículas de las listas del cierre que no tienen bloque en el reparto
    (externos o un avión sin dinero en el periodo), en orden de aparición:
    ningún vuelo diferido se pierde en silencio."""
    if cierre is None:
        return []
    propias = set(matriculas_del_reparto)
    fuera: list[str | None] = []
    for attr, _plantilla, _mes in LISTAS_REPARTO:
        for v in getattr(cierre, attr):
            if v.matricula not in propias and v.matricula not in fuera:
                fuera.append(v.matricula)
    return fuera
