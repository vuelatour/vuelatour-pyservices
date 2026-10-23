"""Formatos compartidos por TODOS los documentos (17-sep-2026; ampliado el
13-oct-2026).

Viven aquí dos cosas, las dos de PRESENTACIÓN (ninguna toca dinero):
1. Cómo se ESCRIBE un tipo de cambio USD→MXN (`_tc_txt`, abajo).
2. Las columnas CLIENTE / RUTA / Nº VUELO CLIENTE que acompañan a la CLAVE
   en todas las hojas con vuelos (encabezados, anchos y el ancho POR
   CONTENIDO con tope de la columna CLIENTE: `ancho_columna_por_contenido`,
   al final del módulo).

T.C. (17-sep-2026):

Por qué existe (pedido del cliente sobre la cotización #314): el operador
captura el T.C. con los decimales que hacen cuadrar los pesos
(5,885.25 USD × 16.991632 = $100,000.00 MXN), pero la hoja lo imprimía con
`:g` — seis CIFRAS SIGNIFICATIVAS, o sea «16.9916». Quien volvía a
multiplicar con ese texto obtenía $99,999.81 y el documento dejaba de
cuadrar consigo mismo («cuando son muchos decimales como que siempre
cambia»).

REGLA: un T.C. se imprime con HASTA 6 DECIMALES y sin ceros de cola
(16.991632 → «16.991632»; 16.9916 → «16.9916»; 18 → «18»). Es el MISMO
texto que da `fmtTc` en el panel (vuelatour-next) y la misma precisión que
persiste el API (`numeric(12,6)`): los tres tienen que decir lo mismo o el
operador vuelve a ver dos números distintos para el mismo dato.

Aquí NO se redondea dinero ni se recalcula nada: solo se formatea. El total
en MXN de un vuelo se LEE del API (`monto_total_mxn`), jamás se recalcula
con este texto.

CLAVE sin apodos (23-oct-2026, API 0.0.105; pedido del cliente del 23-oct
sobre «reporte horas FLOTA» y 'otros movimientos': «siempre debemos ocultar
los apodos con la clave del vuelo, para dejar solo las columnas de (Num de
cotización, nombre completo, fecha, ruta y consecutivo del vuelo)» y «mejor
eliminar esos apodos de todos lados, y de hecho faltan las columnas de
nombre, ruta en algunos otros lados»). Del 13-oct al 23-oct la CLAVE llegaba
armada con el ID por cliente («#376 · january02-211026 · January Davila») y
`clave_cliente` solo ensanchaba la columna A. Desde el API 0.0.105 la CLAVE
es SOLO el folio («#376») y cada fila de vuelo trae `cliente` (nombre
completo), `ruta` y `consecutivo_cliente` (el «02» de «january02»), que se
pintan en columnas propias JUNTO a la CLAVE: CLIENTE y Nº VUELO CLIENTE en
todas las hojas con vuelos (y RUTA donde la hoja no la tenía). La CLAVE se
pinta TAL CUAL llega (con un API viejo sigue trayendo el apodo: aquí no se
parte ni se inventa nada; CLIENTE queda vacía) y `clave_cliente` se ignora.
"""

from __future__ import annotations

from collections.abc import Iterable

# Decimales que persiste el API (`numeric(12,6)`) y máximo que se imprime.
TC_DECIMALES = 6


def _tc_txt(v: float | int | None) -> str:
    """Tipo de cambio como texto: hasta `TC_DECIMALES` decimales, sin ceros
    de cola ni punto huérfano.

    `None` → cadena vacía: quien llama decide si pinta «—» o se salta la
    línea (así ningún documento inventa un T.C. que no llegó).
    """
    if v is None:
        return ""
    txt = f"{float(v):.{TC_DECIMALES}f}"
    return txt.rstrip("0").rstrip(".") if "." in txt else txt


# Columnas que acompañan a la CLAVE en TODAS las hojas con vuelos
# (23-oct-2026, API 0.0.105): el nombre completo del cliente y el consecutivo
# del vuelo de ESE cliente. Los encabezados en MAYÚSCULAS son los de las hojas
# que así escriben los suyos (maestra, cobranza, 'dinero-vlos'); 'otros
# movimientos' y 'Otros ingresos' van en minúsculas como el resto de su fila.
ENCABEZADO_CLIENTE = "CLIENTE"
ENCABEZADO_RUTA = "RUTA"
ENCABEZADO_NUM_VUELO_CLIENTE = "Nº VUELO\nCLIENTE"
# Ancho de CLIENTE: POR CONTENIDO entre una base y un tope (un nombre
# larguísimo no ensancha la hoja sin límite: Excel lo corta como siempre).
ANCHO_CLIENTE_BASE = 22
ANCHO_CLIENTE_TOPE = 40
ANCHO_NUM_VUELO_CLIENTE = 9
# Tope de la columna CLAVE del Libro Dinero cuando crece por la marca del
# cierre de mes («vtbeh · Viene de septiembre 2026»).
ANCHO_CLAVE_TOPE_DINERO = 40


def ancho_columna_por_contenido(textos: Iterable[str | None], *, base: int, tope: int) -> int:
    """Ancho de una columna de texto: `min(max(base, largo + 2), tope)`, con
    `largo` = el texto PINTADO más largo (0 sin textos: la base). Nunca se
    ANGOSTA por debajo de la base ni crece sobre el tope."""
    largo = max((len(t) for t in textos if t), default=0)
    return min(max(int(base), largo + 2), int(tope))


def ancho_columna_cliente(nombres: Iterable[str | None]) -> int:
    """Ancho de la columna CLIENTE (igual en todas las hojas)."""
    return ancho_columna_por_contenido(nombres, base=ANCHO_CLIENTE_BASE, tope=ANCHO_CLIENTE_TOPE)
