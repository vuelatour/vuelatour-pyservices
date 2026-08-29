import logging

from fastapi import APIRouter, Depends, HTTPException, Response, status

from app.schemas.inventario import (
    ParseInventarioRequest,
    ParseInventarioResponse,
    PlantillaInventarioRequest,
)
from app.security import require_internal_token
from app.services.inventario_masivo import parse_inventario, render_plantilla_inventario

logger = logging.getLogger("inventario")

router = APIRouter(
    prefix="/inventario", tags=["inventario"], dependencies=[Depends(require_internal_token)]
)

XLSX_MEDIA = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


@router.post("/plantilla")
def plantilla_inventario(req: PlantillaInventarioRequest) -> Response:
    """Plantilla XLSX de alta masiva de inventario (dropdowns con catálogos)."""
    xlsx_bytes = render_plantilla_inventario(req)
    return Response(
        content=xlsx_bytes,
        media_type=XLSX_MEDIA,
        headers={"Content-Disposition": 'attachment; filename="plantilla-inventario.xlsx"'},
    )


@router.post("/parse", response_model=ParseInventarioResponse)
def parse_inventario_endpoint(req: ParseInventarioRequest) -> ParseInventarioResponse:
    """Convierte la plantilla llenada (XLSX/CSV) a filas crudas; el API valida negocio."""
    try:
        return parse_inventario(req)
    except ValueError as e:
        logger.warning("Plantilla de inventario no parseable: %s", e)
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY, detail=str(e)
        ) from e
