"""Formatos compartidos por TODOS los documentos (17-sep-2026; ampliado el
13-oct-2026).

Viven aquí dos cosas, las dos de PRESENTACIÓN (ninguna toca dinero):
1. Cómo se ESCRIBE un tipo de cambio USD→MXN (`_tc_txt`, abajo).
2. El ANCHO de la columna CLAVE de los libros cuando la clave trae el ID de
   vuelo por cliente (`ancho_columna_clave`, al final del módulo).

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

CLAVE con ID de vuelo por cliente (13-oct-2026, API 0.0.75, decisión 6 del
lote del cliente): el API arma la CLAVE con el ID junto al folio («#376 ·
january02-211026 · January Davila» en la hoja maestra y la cobranza, «#376 ·
january02-211026» en 'otros movimientos', «january02-211026» en el Libro
Dinero) y manda `clave_cliente` SOLO cuando el vuelo tiene ID. Ese campo NO
se pinta: solo decide que la columna A se ensanche POR CONTENIDO, con tope.
Sin él (API previo o migración sin aplicar), el ancho de siempre: libro byte
a byte el de antes.
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


# Tope del ancho de la columna CLAVE con ID de vuelo por cliente: balances
# (hoja maestra, cobranza, 'otros movimientos') y Libro Dinero ('dinero-vlos',
# 'Otros ingresos'). Un nombre de cliente larguísimo no ensancha la hoja sin
# límite: Excel lo corta como hoy.
ANCHO_CLAVE_TOPE_BALANCE = 48
ANCHO_CLAVE_TOPE_DINERO = 40


def hay_id_cliente(filas: Iterable[object]) -> bool:
    """¿Alguna fila trae el ID de vuelo por cliente (`clave_cliente` con
    valor)? Filas sin el atributo (otro modelo, API previo) cuentan como no."""
    return any(getattr(f, "clave_cliente", None) for f in filas)


def ancho_columna_clave(
    textos: Iterable[str | None], *, base: int, tope: int, con_id_cliente: bool
) -> int:
    """Ancho de la columna CLAVE.

    - Sin ID de vuelo por cliente (`con_id_cliente` False): `base` TAL CUAL,
      como int — el libro sale byte a byte el de antes.
    - Con ID: `min(max(base, largo + 2), tope)`, con `largo` = el texto
      PINTADO más largo (el «·» cuenta 1; 0 sin textos). Un nombre
      larguísimo no ensancha la hoja sin límite: Excel lo corta como hoy.
    """
    if not con_id_cliente:
        return int(base)
    largo = max((len(t) for t in textos if t), default=0)
    return min(max(int(base), largo + 2), int(tope))
