"""Schema genérico para exportar cualquier tabla a Excel."""

from pydantic import BaseModel, Field


class TablaColumna(BaseModel):
    label: str
    # texto | money | numero | entero | pct
    tipo: str = "texto"


class CeldaResalte(BaseModel):
    """Celda de `filas` a resaltar (índices 0-based relativos a `filas`)."""

    fila: int
    col: int
    # Color hex del texto SIN "#" (default: naranja).
    color: str = "ED7D31"


class TablaHoja(BaseModel):
    """Una hoja del libro (ADITIVO, 9-sep-2026: auditoría Paywise en 3 hojas).

    Mismos campos que la tabla raíz; `titulo` es el NOMBRE de la pestaña.
    """

    titulo: str
    subtitulo: str | None = None
    columnas: list[TablaColumna]
    filas: list[list] = Field(default_factory=list)
    totales: list | None = None
    resumen_titulo: str | None = None
    resumen: list[list] | None = None
    resaltes: list[CeldaResalte] | None = None


class TablaXlsxRequest(BaseModel):
    titulo: str
    subtitulo: str | None = None
    columnas: list[TablaColumna]
    # Cada fila es una lista de valores alineada a `columnas`.
    filas: list[list] = Field(default_factory=list)
    # Fila de totales opcional, alineada a `columnas` (None en celdas vacías).
    totales: list | None = None
    # Bloque RESUMEN opcional arriba de la tabla: pares [etiqueta, valor]
    # (el valor numérico se pinta como moneda). P. ej. total por categoría.
    resumen_titulo: str | None = None
    resumen: list[list] | None = None
    # ADITIVO: celdas de `filas` a resaltar (p. ej. montos SIN conciliar en
    # naranja). None/[] = render idéntico al de siempre para el resto de
    # callers de tabla-xlsx.
    resaltes: list[CeldaResalte] | None = None
    # ADITIVO: varias hojas en el mismo libro. Con `hojas`, la tabla raíz NO
    # se pinta (solo nombra el libro); sin ellas, render idéntico al de siempre.
    hojas: list[TablaHoja] | None = None
