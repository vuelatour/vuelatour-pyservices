"""Formatos numéricos compartidos por TODOS los documentos (17-sep-2026).

Hoy vive aquí UNA sola cosa: cómo se ESCRIBE un tipo de cambio USD→MXN.

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
"""

from __future__ import annotations

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
