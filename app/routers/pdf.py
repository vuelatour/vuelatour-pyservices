"""Endpoints de generacion de PDF. Protegidos con el token de servicio."""

from fastapi import APIRouter, Depends, Response

from app.schemas.cotizacion import CotizacionPdfRequest
from app.schemas.reparto import RepartoPdfRequest
from app.security import require_service_token
from app.services.cotizacion_pdf import render_cotizacion_pdf
from app.services.reparto_pdf import render_reparto_pdf

router = APIRouter(
    prefix="/pdf",
    tags=["pdf"],
    dependencies=[Depends(require_service_token)],
)


@router.post("/cotizacion")
def cotizacion_pdf(payload: CotizacionPdfRequest) -> Response:
    """Genera el PDF de una cotizacion y lo devuelve como application/pdf."""
    pdf_bytes = render_cotizacion_pdf(payload)
    filename = f"cotizacion-{payload.folio}-v{payload.version}.pdf"
    return Response(
        content=pdf_bytes,
        media_type="application/pdf",
        headers={"Content-Disposition": f'inline; filename="{filename}"'},
    )


@router.post("/reparto")
def reparto_pdf(payload: RepartoPdfRequest) -> Response:
    """Genera el PDF del reparto de utilidades del periodo."""
    pdf_bytes = render_reparto_pdf(payload)
    filename = f"reparto-{payload.periodo_desde}-a-{payload.periodo_hasta}.pdf"
    return Response(
        content=pdf_bytes,
        media_type="application/pdf",
        headers={"Content-Disposition": f'inline; filename="{filename}"'},
    )
