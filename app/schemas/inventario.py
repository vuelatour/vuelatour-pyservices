"""Alta masiva de inventario (28-ago-2026): plantilla XLSX + parseo estructural.

La oficina descarga la plantilla (una fila = un ítem de bodega, con su
empaque/caja opcional y su existencia inicial), la llena y la sube. Aquí
SOLO se renderiza el libro y se convierte de vuelta a filas crudas con tipos
básicos; toda la validación de negocio (nombre/categoría obligatorios,
códigos únicos, costo/TC de la entrada inicial) vive en vuelatour-api.

Esquemas ADITIVOS: todos los campos con default para tolerar skew de deploy.
"""

from pydantic import BaseModel, Field


class PlantillaInventarioRequest(BaseModel):
    """Catálogos vigentes para los dropdowns de la plantilla (los manda el API)."""

    categorias: list[str] = Field(
        default_factory=list, description="Categorías existentes en inventario_item"
    )
    unidades: list[str] = Field(
        default_factory=list, description="Unidades de medida sugeridas (pieza, botella, litro…)"
    )
    monedas: list[str] = Field(default_factory=list, description="Monedas del costo (MXN/USD)")


class ParseInventarioRequest(BaseModel):
    archivo_base64: str
    filename: str = ""


class FilaInventario(BaseModel):
    """Valores CRUDOS de una fila de la plantilla, en tipos básicos.

    Aquí NO se valida negocio: eso lo hace el API. Los numéricos llegan como
    float si la celda fue numérica; si el texto no se pudo convertir se
    regresa la cadena cruda para que el API la reporte con claridad. Los
    códigos de barras se normalizan (sin espacios) y siempre viajan como str
    (aunque Excel los haya guardado como número).
    """

    fila: int = Field(description="Número de fila real en el Excel/CSV (1-based)")
    nombre: str | None = None
    marca: str | None = None
    categoria: str | None = None
    numero_parte: str | None = None
    codigo: str | None = Field(
        default=None, description="Código de barras / SKU de la UNIDAD, normalizado (sin espacios)"
    )
    unidad: str | None = Field(default=None, description="Unidad de medida (pieza, botella…)")
    descripcion: str | None = None
    ubicacion: str | None = None
    stock_minimo: float | str | None = None
    existencia_inicial: float | str | None = None
    costo_unitario: float | str | None = None
    moneda: str | None = None
    tipo_cambio: float | str | None = None
    empaque_nombre: str | None = Field(default=None, description="Ej. 'Caja de 6'")
    empaque_factor: float | str | None = Field(
        default=None, description="Unidades del ítem por empaque (caja de 6 → 6)"
    )
    empaque_codigo: str | None = Field(
        default=None, description="Código de barras del EMPAQUE, normalizado (sin espacios)"
    )
    notas: str | None = None
    avisos: list[str] = Field(
        default_factory=list,
        description=(
            "Advertencias estructurales de la lectura (p. ej. código de barras que Excel "
            "guardó como número y pudo perder ceros a la izquierda). El API las puede "
            "anexar a sus mensajes."
        ),
    )


class ParseInventarioResponse(BaseModel):
    filas: list[FilaInventario] = Field(default_factory=list)
