import logging

import anthropic
from fastapi import APIRouter, Depends, HTTPException, Response, status

from app.schemas.gastos import (
    GastoVueloSugerirRequest,
    GastoVueloSugerirResponse,
    ParseCombustibleRequest,
    ParseCombustibleResponse,
    PlantillaCombustibleRequest,
)
from app.security import require_internal_token
from app.services.combustible_masivo import (
    parse_combustible,
    render_plantilla_combustible,
)
from app.services.gasto_vuelo import sugerir_vuelo_para_gasto

logger = logging.getLogger("gastos")

router = APIRouter(
    prefix="/gastos", tags=["gastos"], dependencies=[Depends(require_internal_token)]
)

XLSX_MEDIA = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


@router.post("/sugerir-vuelo", response_model=GastoVueloSugerirResponse)
def sugerir_vuelo(req: GastoVueloSugerirRequest) -> GastoVueloSugerirResponse:
    try:
        return sugerir_vuelo_para_gasto(req)
    except anthropic.APIStatusError as e:
        logger.warning("Claude API error %s: %s", e.status_code, e.message)
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail=f"Claude no disponible ({e.status_code})",
        ) from e
    except (ValueError, KeyError) as e:
        logger.warning("Sugerencia gasto→vuelo no parseable: %s", e)
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="No se pudo interpretar la sugerencia",
        ) from e


@router.post("/plantilla-combustible")
def plantilla_combustible(req: PlantillaCombustibleRequest) -> Response:
    """Plantilla XLSX de carga masiva de combustible (dropdowns con catálogos)."""
    xlsx_bytes = render_plantilla_combustible(req)
    return Response(
        content=xlsx_bytes,
        media_type=XLSX_MEDIA,
        headers={
            "Content-Disposition": 'attachment; filename="plantilla-combustible.xlsx"'
        },
    )


@router.post("/parse-combustible", response_model=ParseCombustibleResponse)
def parse_combustible_endpoint(req: ParseCombustibleRequest) -> ParseCombustibleResponse:
    """Convierte la plantilla llenada (XLSX/CSV) a filas crudas; el API valida negocio."""
    try:
        return parse_combustible(req)
    except ValueError as e:
        logger.warning("Plantilla de combustible no parseable: %s", e)
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY, detail=str(e)
        ) from e
