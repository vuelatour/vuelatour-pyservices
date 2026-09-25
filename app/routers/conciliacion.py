import logging

import anthropic
from fastapi import APIRouter, Depends, HTTPException, status

from app.schemas.conciliacion import (
    ConciliacionParseRequest,
    ConciliacionParseResponse,
    ConciliacionSugerirAbonosRequest,
    ConciliacionSugerirAbonosResponse,
    ConciliacionSugerirRequest,
    ConciliacionSugerirResponse,
)
from app.security import require_internal_token
from app.services.conciliacion_abonos import (
    CODIGO_ILEGIBLE,
    CODIGO_TRUNCADA,
    RespuestaIlegibleError,
    RespuestaTruncadaError,
    sugerir_abonos,
)
from app.services.estado_cuenta import parsear_estado_cuenta, sugerir_conciliacion

logger = logging.getLogger("conciliacion")

router = APIRouter(
    prefix="/conciliacion", tags=["conciliacion"], dependencies=[Depends(require_internal_token)]
)


@router.post("/parse", response_model=ConciliacionParseResponse)
def parse(req: ConciliacionParseRequest) -> ConciliacionParseResponse:
    try:
        return parsear_estado_cuenta(req)
    # ImportError (no solo ModuleNotFoundError): pandas relanza ImportError
    # plano cuando falta una dependencia opcional de lectura.
    except ImportError as e:
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


@router.post("/sugerir", response_model=ConciliacionSugerirResponse)
def sugerir(req: ConciliacionSugerirRequest) -> ConciliacionSugerirResponse:
    try:
        return sugerir_conciliacion(req)
    except anthropic.APIStatusError as e:
        logger.warning("Claude API error %s: %s", e.status_code, e.message)
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail=f"Claude no disponible ({e.status_code})",
        ) from e
    except (ValueError, KeyError) as e:
        logger.warning("Sugerencia de conciliación no parseable: %s", e)
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="No se pudo interpretar la sugerencia de conciliación",
        ) from e


def _detalle(codigo: str, mensaje: str, uso_ia=None) -> dict:
    """Detalle JSON de los errores de `/sugerir-abonos`.

    El código va en `error` (convención de los errores del API) y repetido en
    `code`; `uso_ia` viaja cuando Claude SÍ respondió: los créditos se
    gastaron y el API los registra en `ia_uso` aunque la respuesta no sirva.
    """
    detalle: dict = {"error": codigo, "code": codigo, "message": mensaje}
    if uso_ia is not None:
        detalle["uso_ia"] = uso_ia.model_dump()
    return detalle


@router.post("/sugerir-abonos", response_model=ConciliacionSugerirAbonosResponse)
def sugerir_abonos_endpoint(
    req: ConciliacionSugerirAbonosRequest,
) -> ConciliacionSugerirAbonosResponse:
    """Conciliación de INGRESOS (24-sep-2026): propone qué cobro de vuelo,
    sobre de grupo o ingreso registrado es cada ABONO del banco — o qué es si
    no es ninguno. Nunca liga: el API valida y la persona confirma. Lotes de
    ≤ 10 abonos y ≤ 120 candidatos (más ⇒ 422 de pydantic)."""
    try:
        return sugerir_abonos(req)
    except RespuestaTruncadaError as e:
        logger.warning("Sugerir abonos: respuesta truncada por max_tokens")
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail=_detalle(
                CODIGO_TRUNCADA,
                "La respuesta del asistente llegó incompleta (límite de salida).",
                e.uso_ia,
            ),
        ) from e
    except anthropic.APIStatusError as e:
        logger.warning("Claude API error %s: %s", e.status_code, e.message)
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail=f"Claude no disponible ({e.status_code})",
        ) from e
    except anthropic.APIConnectionError as e:
        # Incluye APITimeoutError: sin esto un timeout salía como 500.
        logger.warning("Claude sin respuesta (conexión/timeout): %s", e)
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail="Claude no respondió a tiempo",
        ) from e
    except anthropic.APIError as e:
        # Resto de la familia del SDK (p. ej. APIResponseValidationError, que
        # no es APIStatusError ni de conexión): 502, nunca 500.
        logger.warning("Claude respondió algo inesperado: %s", e)
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail="Claude no disponible",
        ) from e
    except RespuestaIlegibleError as e:
        logger.warning("Sugerencia de abonos no parseable: %s", e.__cause__ or e)
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail=_detalle(CODIGO_ILEGIBLE, str(e), e.uso_ia),
        ) from e
    except (ValueError, KeyError) as e:
        logger.warning("Sugerencia de abonos no parseable: %s", e)
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="No se pudo interpretar la sugerencia de conciliación de abonos",
        ) from e
