import logging

import anthropic
from fastapi import APIRouter, Depends, HTTPException, status

from app.schemas.compras import CompraExtraerRequest, CompraExtraerResponse
from app.security import require_internal_token
from app.services.compras_extract import extraer_compra

logger = logging.getLogger("compras")

router = APIRouter(prefix="/compras", tags=["compras"], dependencies=[Depends(require_internal_token)])


@router.post("/extraer", response_model=CompraExtraerResponse)
def extraer(req: CompraExtraerRequest) -> CompraExtraerResponse:
    try:
        return extraer_compra(req)
    except anthropic.APIStatusError as e:
        logger.warning("Claude API error %s: %s", e.status_code, e.message)
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail=f"Claude no disponible ({e.status_code})",
        ) from e
    except (ValueError, KeyError) as e:
        logger.warning("Respuesta de Claude no parseable: %s", e)
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="No se pudo interpretar el PDF de compra",
        ) from e
