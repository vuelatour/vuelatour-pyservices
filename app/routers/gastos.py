import logging

import anthropic
from fastapi import APIRouter, Depends, HTTPException, status

from app.schemas.gastos import GastoVueloSugerirRequest, GastoVueloSugerirResponse
from app.security import require_internal_token
from app.services.gasto_vuelo import sugerir_vuelo_para_gasto

logger = logging.getLogger("gastos")

router = APIRouter(
    prefix="/gastos", tags=["gastos"], dependencies=[Depends(require_internal_token)]
)


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
