import logging

import anthropic
from fastapi import APIRouter, Depends, HTTPException, status

from app.schemas.vencimiento import VencimientoExtraerRequest, VencimientoExtraerResponse
from app.security import require_internal_token
from app.services.ia_errores import detalle_error_claude, registrar_error_claude
from app.services.vencimiento_extract import extraer_vencimiento

logger = logging.getLogger("vencimientos")

router = APIRouter(
    prefix="/vencimientos", tags=["vencimientos"], dependencies=[Depends(require_internal_token)]
)


@router.post("/extraer", response_model=VencimientoExtraerResponse)
def extraer(req: VencimientoExtraerRequest) -> VencimientoExtraerResponse:
    try:
        return extraer_vencimiento(req)
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
            detail="No se pudo interpretar el documento de vencimiento",
        ) from e
