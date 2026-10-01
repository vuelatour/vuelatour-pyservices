import logging

import anthropic
from fastapi import APIRouter, Depends, HTTPException, status

from app.schemas.compras import CompraExtraerRequest, CompraExtraerResponse
from app.security import require_internal_token
from app.services.compras_extract import extraer_compra
from app.services.ia_errores import detalle_error_claude, registrar_error_claude

logger = logging.getLogger("compras")

router = APIRouter(prefix="/compras", tags=["compras"], dependencies=[Depends(require_internal_token)])


@router.post("/extraer", response_model=CompraExtraerResponse)
def extraer(req: CompraExtraerRequest) -> CompraExtraerResponse:
    try:
        return extraer_compra(req)
    except anthropic.APIStatusError as e:
        registrar_error_claude(logger, e)
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail=detalle_error_claude(e),
        ) from e
    except (ValueError, KeyError) as e:
        logger.warning("Respuesta de Claude no parseable: %s", e)
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="No se pudo interpretar el PDF de compra",
        ) from e
