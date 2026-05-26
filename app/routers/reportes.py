import logging

from fastapi import APIRouter, Depends, HTTPException, Response, status

from app.schemas.reportes import CotizacionPdfRequest
from app.security import require_internal_token
from app.services.cotizacion_pdf import render_cotizacion_pdf

logger = logging.getLogger("reportes")

router = APIRouter(
    prefix="/reportes",
    tags=["reportes"],
    dependencies=[Depends(require_internal_token)],
)


@router.post("/cotizacion")
def cotizacion_pdf(req: CotizacionPdfRequest) -> Response:
    try:
        pdf = render_cotizacion_pdf(req)
    except ImportError as e:
        logger.error("WeasyPrint no disponible: %s", e)
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail="Generador de PDF no instalado (WeasyPrint).",
        ) from e
    except Exception as e:  # noqa: BLE001
        logger.exception("Error generando PDF de cotización")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"No se pudo generar el PDF: {e}",
        ) from e
    return Response(content=pdf, media_type="application/pdf")
