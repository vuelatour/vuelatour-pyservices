import logging

import anthropic
from fastapi import APIRouter, Depends, HTTPException, status

from app.schemas.vision import (
    CombustibleTicketResponse,
    ConstanciaFiscalRequest,
    ConstanciaFiscalResponse,
    GastoTicketRequest,
    GastoTicketResponse,
    TacometroRequest,
    TacometroResponse,
)
from app.security import require_internal_token
from app.services.anthropic_vision import (
    leer_constancia_fiscal,
    leer_tacometro,
    leer_ticket_combustible,
    leer_ticket_gasto,
)

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
            # El motivo real ayuda al operador (ej. ".xls no soportado").
            detail=str(e)[:200] or "No se pudo interpretar el ticket",
        ) from e


@router.post("/constancia-fiscal", response_model=ConstanciaFiscalResponse)
def constancia_fiscal(req: ConstanciaFiscalRequest) -> ConstanciaFiscalResponse:
    """Lee una Constancia de Situación Fiscal del SAT (PDF o foto) para
    autocompletar los datos fiscales del cliente. Los fallos de IA degradan
    a disponible=false/legible=false (captura manual), nunca 500."""
    try:
        return leer_constancia_fiscal(req)
    except anthropic.APIStatusError as e:
        logger.warning("Claude API error %s: %s", e.status_code, e.message)
        return ConstanciaFiscalResponse(
            disponible=False,
            legible=False,
            motivo=f"IA no disponible ({e.status_code}); captura los datos manualmente.",
        )
    except anthropic.APIError as e:
        # Timeout / red caída: mismo destino que un 5xx del modelo.
        logger.warning("Claude API sin respuesta: %s", e)
        return ConstanciaFiscalResponse(
            disponible=False,
            legible=False,
            motivo="IA no disponible (sin conexión); captura los datos manualmente.",
        )
    except (ValueError, KeyError, TypeError) as e:
        # ValueError cubre también la ValidationError de pydantic (respuesta
        # del modelo fuera de rango): degradar, no 500.
        logger.warning("Respuesta de Claude no parseable: %s", e)
        return ConstanciaFiscalResponse(
            disponible=True,
            legible=False,
            motivo="No se pudo interpretar la constancia; captura los datos manualmente.",
        )


@router.post("/combustible", response_model=CombustibleTicketResponse)
def combustible(req: GastoTicketRequest) -> CombustibleTicketResponse:
    try:
        return leer_ticket_combustible(req)
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
            detail="No se pudo interpretar el ticket de combustible",
        ) from e
