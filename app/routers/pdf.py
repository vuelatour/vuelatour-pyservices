"""Endpoints de generacion de PDF. Protegidos con el token de servicio."""

from fastapi import APIRouter, Depends, Response

from app.schemas.reparto import RepartoPdfRequest
from app.security import require_service_token
from app.services.reparto_pdf import render_reparto_pdf

router = APIRouter(
    prefix="/pdf",
    tags=["pdf"],
    dependencies=[Depends(require_service_token)],
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
