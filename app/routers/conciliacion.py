import logging

import anthropic
from fastapi import APIRouter, Depends, HTTPException, status

from app.schemas.conciliacion import ConciliacionParseRequest, ConciliacionParseResponse
from app.security import require_internal_token
from app.services.estado_cuenta import parsear_estado_cuenta

logger = logging.getLogger("conciliacion")

router = APIRouter(
    prefix="/conciliacion", tags=["conciliacion"], dependencies=[Depends(require_internal_token)]
)


@router.post("/parse", response_model=ConciliacionParseResponse)
def parse(req: ConciliacionParseRequest) -> ConciliacionParseResponse:
    try:
        return parsear_estado_cuenta(req)
    except ModuleNotFoundError as e:
        logger.warning("Dependencia faltante: %s", e)
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail="Falta pandas/openpyxl en el servidor para leer CSV/Excel",
        ) from e
    except anthropic.APIStatusError as e:
        logger.warning("Claude API error %s: %s", e.status_code, e.message)
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail=f"Claude no disponible ({e.status_code})",
        ) from e
    except (ValueError, KeyError) as e:
        logger.warning("Estado de cuenta no parseable: %s", e)
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail=str(e),
        ) from e
