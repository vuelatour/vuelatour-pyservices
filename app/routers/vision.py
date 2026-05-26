import logging

import anthropic
from fastapi import APIRouter, Depends, HTTPException, status

from app.schemas.vision import (
    GastoTicketRequest,
    GastoTicketResponse,
    TacometroRequest,
    TacometroResponse,
)
from app.security import require_internal_token
from app.services.anthropic_vision import leer_tacometro, leer_ticket_gasto

logger = logging.getLogger("vision")

router = APIRouter(prefix="/vision", tags=["vision"], dependencies=[Depends(require_internal_token)])


@router.post("/tacometro", response_model=TacometroResponse)
def tacometro(req: TacometroRequest) -> TacometroResponse:
    try:
        return leer_tacometro(req)
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
            detail="No se pudo interpretar la lectura del tacómetro",
        ) from e


@router.post("/gasto", response_model=GastoTicketResponse)
def gasto(req: GastoTicketRequest) -> GastoTicketResponse:
    try:
        return leer_ticket_gasto(req)
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
            detail="No se pudo interpretar el ticket",
        ) from e
