"""Evaluador INDEPENDIENTE de las fórmulas de un .xlsx, solo para tests
(5-oct-2026, «fórmulas visibles» del balance por avión).

No reutiliza el evaluador de `app.services.xlsx_formulas` (ese decide en
producción si una fórmula se queda): aquí la fórmula se TRADUCE a una
expresión de Python con expresiones regulares y se evalúa sobre el libro YA
GUARDADO, leído dos veces — con fórmulas (`data_only=False`) y con la caché
(`data_only=True`). Cubre lo que generan los balances: referencias A1 /
$A$1 / 'hoja'!A1, rangos, SUM, AVERAGE, AVERAGEIF, SUMIF, ROUND, + − × ÷ y
paréntesis. Una celda con fórmula se evalúa con el valor COMPLETO de las
fórmulas que cita (como Excel), no con su caché redondeada.

`verificar_libro` exige además que toda fórmula lleve ROUND exterior salvo
las referencias puras y la lista blanca `ENCABEZADOS_SIN_ROUND` (revisión
5-oct-2026: sin esa regla, quitar un ROUND seguía cuadrando a la vista).
"""

from __future__ import annotations

import re
from decimal import ROUND_HALF_UP, Decimal
from io import BytesIO

from openpyxl import load_workbook
from openpyxl.utils.cell import column_index_from_string

_TEXTO = re.compile(r'"(?:[^"]|"")*"')
_REF = re.compile(
    r"(?:'(?P<hoja>(?:[^']|'')+)'!)?"
    r"\$?(?P<c1>[A-Z]{1,3})\$?(?P<f1>\d+)"
    r"(?::\$?(?P<c2>[A-Z]{1,3})\$?(?P<f2>\d+))?"
    r"(?![A-Za-z0-9_(])"
)
_FUNCION = re.compile(r"\b(AVERAGEIF|SUMIF|AVERAGE|SUM|ROUND)\(")


def round_excel(x: float, n: float) -> float:
    """ROUND de Excel: medio hacia afuera sobre 15 cifras significativas."""
    d = Decimal(format(x, ".15g"))
    return float(d.quantize(Decimal(1).scaleb(-int(n)), rounding=ROUND_HALF_UP))


def _es_numero(x: object) -> bool:
    return isinstance(x, (int, float)) and not isinstance(x, bool)


def _cumple(valor: object, criterio: str) -> bool:
    m = re.match(r"^(<>|>=|<=|>|<|=)?(.*)$", criterio)
    op, obj = (m.group(1) or "="), m.group(2)
    try:
        num = float(obj)
    except ValueError:
        texto = "" if valor is None else str(valor)
        return (texto.upper() == obj.upper()) == (op == "=")
    if not _es_numero(valor):
        return op == "<>"
    v = float(valor)
    return {
        "=": v == num,
        "<>": v != num,
        ">": v > num,
        "<": v < num,
        ">=": v >= num,
        "<=": v <= num,
    }[op]


class Rango(list):
    """Lista de valores de un rango (None = vacía)."""


def _aplanar(args) -> list:
    out: list = []
    for a in args:
        out.extend(a if isinstance(a, Rango) else [a])
    return out


def _SUM(*args) -> float:  # noqa: N802 - nombre de la función de Excel
    total = 0.0
    for x in _aplanar(args):
        if _es_numero(x):
            total += float(x)
    return total


def _AVERAGE(*args) -> float:  # noqa: N802
    nums = [float(x) for x in _aplanar(args) if _es_numero(x)]
    if not nums:
        raise ZeroDivisionError("AVERAGE sin números")
    total = 0.0
    for x in nums:
        total += x
    return total / len(nums)


def _AVERAGEIF(rango: Rango, criterio: str, promedio: Rango | None = None) -> float:  # noqa: N802
    datos = promedio if promedio is not None else rango
    return _AVERAGE(Rango(d for p, d in zip(rango, datos, strict=True) if _cumple(p, criterio)))


def _SUMIF(rango: Rango, criterio: str, suma: Rango | None = None) -> float:  # noqa: N802
    datos = suma if suma is not None else rango
    return _SUM(Rango(d for p, d in zip(rango, datos, strict=True) if _cumple(p, criterio)))


def _ROUND(x, n) -> float:  # noqa: N802
    return round_excel(float(x or 0), n)


class Libro:
    """Un .xlsx con fórmulas: `formulas()` las lista y `evaluar()` calcula el
    valor que Excel les daría."""

    def __init__(self, data: bytes) -> None:
        self.con_formulas = load_workbook(BytesIO(data))
        self.con_cache = load_workbook(BytesIO(data), data_only=True)
        self._memo: dict[tuple[str, str], float] = {}

    def formulas(self) -> list[tuple[str, str, str]]:
        """(hoja, celda, fórmula sin «=») de TODAS las celdas con fórmula."""
        return [
            (ws.title, c.coordinate, c.value[1:])
            for ws in self.con_formulas.worksheets
            for fila in ws.iter_rows()
            for c in fila
            if c.data_type == "f"
        ]

    def cache(self, hoja: str, celda: str):
        return self.con_cache[hoja][celda].value

    def _valor(self, hoja: str, fila: int, col: int):
        ws = self.con_formulas[hoja]
        c = ws.cell(row=fila, column=col)
        if c.data_type == "f":
            return self.evaluar(hoja, c.coordinate)
        return c.value

    def _celda(self, hoja: str, letra: str, fila: str):
        v = self._valor(hoja, int(fila), column_index_from_string(letra))
        return 0.0 if v is None else v

    def _rango(self, hoja: str, c1: str, f1: str, c2: str, f2: str) -> Rango:
        a, b = sorted((column_index_from_string(c1), column_index_from_string(c2)))
        x, y = sorted((int(f1), int(f2)))
        return Rango(self._valor(hoja, f, c) for f in range(x, y + 1) for c in range(a, b + 1))

    def _traducir(self, trozo: str, hoja: str) -> str:
        def ref(m: re.Match) -> str:
            h = (m.group("hoja") or hoja).replace("''", "'")
            if m.group("c2"):
                return (
                    f"_r({h!r},{m.group('c1')!r},{m.group('f1')!r},"
                    f"{m.group('c2')!r},{m.group('f2')!r})"
                )
            return f"_c({h!r},{m.group('c1')!r},{m.group('f1')!r})"

        return _FUNCION.sub(lambda m: f"_{m.group(1)}(", _REF.sub(ref, trozo))

    def evaluar(self, hoja: str, celda: str) -> float:
        clave = (hoja, celda)
        if clave in self._memo:
            return self._memo[clave]
        texto = self.con_formulas[hoja][celda].value[1:]
        partes, pos = [], 0
        for m in _TEXTO.finditer(texto):
            partes.append(self._traducir(texto[pos : m.start()], hoja))
            partes.append(repr(m.group(0)[1:-1].replace('""', '"')))
            pos = m.end()
        partes.append(self._traducir(texto[pos:], hoja))
        espacio = {
            "_c": self._celda,
            "_r": self._rango,
            "_SUM": _SUM,
            "_AVERAGE": _AVERAGE,
            "_AVERAGEIF": _AVERAGEIF,
            "_SUMIF": _SUMIF,
            "_ROUND": _ROUND,
        }
        valor = float(eval("".join(partes), {"__builtins__": {}}, espacio))  # noqa: S307
        self._memo[clave] = valor
        return valor


def es_round_externo(formula: str) -> bool:
    """¿La fórmula ENTERA es ROUND(…)? (el paréntesis que abre cierra al final)."""
    if not formula.startswith("ROUND("):
        return False
    nivel = 0
    for i, ch in enumerate(formula):
        if ch == "(":
            nivel += 1
        elif ch == ")":
            nivel -= 1
            if nivel == 0:
                return i == len(formula) - 1
    return False


# Las ÚNICAS fórmulas que no llevan ROUND exterior (revisión 5-oct-2026): el
# API no hace `round2` en su cadena (AE = Y/z, AF = AE/1.16, …; solo redondea
# al SERIALIZAR) y la fórmula tampoco, o el siguiente eslabón dejaría de
# reproducirlo. Se identifican por el ENCABEZADO de su columna y solo como
# aritmética de fila (sin funciones): un TOTALES de estas columnas con
# SUM/AVERAGE sí lleva ROUND. Cualquier otra fórmula sin ROUND exterior —
# salvo una referencia pura (=F14, ='hoja'!$C$5)— es un ROUND perdido: con la
# tolerancia «a la vista» el número seguiría cuadrando y la regresión
# llegaría a producción sin aviso.
ENCABEZADOS_SIN_ROUND = frozenset(
    {
        "COSTO TOTAL USD",
        "COSTO TOTAL USD S/IVA",
        "IVA PAGADO USD",
        "IVA PAGADO MXN",
        "COSTO X HORA USD",
        "COSTO X HORA USD S/IVA",
        "% COBRADO",
    }
)
_REFERENCIA_PURA = re.compile(r"^(?:'(?:[^']|'')+'!)?\$?[A-Z]{1,3}\$?\d+$")


def es_referencia_pura(formula: str) -> bool:
    """¿La fórmula solo copia otra celda (=F14, ='hoja'!$C$5)?"""
    return _REFERENCIA_PURA.match(formula) is not None


def _normalizar_encabezado(texto: str) -> str:
    return " ".join(texto.split()).upper()


def encabezado(libro: Libro, hoja: str, celda: str) -> str | None:
    """Encabezado de la columna de la celda: el primer TEXTO (no fórmula) que
    hay arriba de ella en la misma columna."""
    ws = libro.con_formulas[hoja]
    c = ws[celda]
    for fila in range(c.row - 1, 0, -1):
        arriba = ws.cell(row=fila, column=c.column)
        if arriba.data_type == "f":
            continue
        if isinstance(arriba.value, str) and arriba.value.strip():
            return _normalizar_encabezado(arriba.value)
    return None


def _tiene_funcion(formula: str) -> bool:
    return re.search(r"[A-Z][A-Z0-9.]*\(", _TEXTO.sub("", formula)) is not None


def _exacto(calculado: float, cache: float) -> bool:
    """Igualdad «exacta»: solo se tolera el último dígito con el que openpyxl
    escribe el número (%.16g)."""
    return abs(calculado - cache) <= 1e-9 * max(1.0, abs(cache))


def tolerancia(libro: Libro, hoja: str, celda: str) -> str:
    """'exacta' (ROUND exterior, referencia a un valor o a una fórmula
    exacta) o 'a la vista' (las de `ENCABEZADOS_SIN_ROUND`, que el API solo
    redondea al serializar). Revienta con un AssertionError legible si la
    fórmula no lleva ROUND y no está en la lista blanca."""
    formula = libro.con_formulas[hoja][celda].value[1:]
    if es_round_externo(formula):
        return "exacta"
    if es_referencia_pura(formula):
        m = _REF.match(formula)
        destino_hoja = (m.group("hoja") or hoja).replace("''", "'")
        destino = libro.con_formulas[destino_hoja][f"{m.group('c1')}{m.group('f1')}"]
        if destino.data_type != "f":
            return "exacta"
        return tolerancia(libro, destino_hoja, destino.coordinate)
    nombre = encabezado(libro, hoja, celda)
    assert nombre in ENCABEZADOS_SIN_ROUND and not _tiene_funcion(formula), (
        f"{hoja}!{celda} ={formula} (columna {nombre!r}) no lleva ROUND exterior y no "
        "está en ENCABEZADOS_SIN_ROUND: ¿se perdió el ROUND(…,2) del round2 del API?"
    )
    return "a la vista"


def verificar_libro(data: bytes) -> list[tuple[str, str, str, float, float]]:
    """Evalúa CADA fórmula y la compara con su caché (el número del API).
    Igualdad exacta salvo en la lista blanca de las que no redondean (ahí,
    medio centavo: el API redondea al serializar). Devuelve (hoja, celda,
    fórmula, calculado, caché); revienta con un AssertionError legible a la
    primera que no cuadre o que perdió su ROUND."""
    libro = Libro(data)
    salida = []
    for hoja, celda, formula in libro.formulas():
        cache = libro.cache(hoja, celda)
        assert _es_numero(cache), f"{hoja}!{celda} =({formula}) sin valor en caché: {cache!r}"
        calculado = libro.evaluar(hoja, celda)
        if tolerancia(libro, hoja, celda) == "exacta":
            assert _exacto(calculado, float(cache)), (
                f"{hoja}!{celda} ={formula} da {calculado!r}, el API {cache!r}"
            )
        else:
            assert abs(calculado - float(cache)) < 0.005, (
                f"{hoja}!{celda} ={formula} da {calculado!r}, el API {cache!r}"
            )
        salida.append((hoja, celda, formula, calculado, float(cache)))
    return salida
