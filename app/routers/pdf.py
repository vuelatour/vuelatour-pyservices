"""Endpoints de generacion de PDF. Protegidos con el token interno (X-Internal-Token)."""

from fastapi import APIRouter, Depends, Response

from app.schemas.reparto import RepartoPdfRequest
from app.security import require_internal_token
from app.services.reparto_pdf import render_reparto_pdf
from app.services.reparto_xlsx import render_reparto_xlsx

router = APIRouter(
    prefix="/pdf",
    tags=["pdf"],
    dependencies=[Depends(require_internal_token)],
)

XLSX_MEDIA = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


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


@router.post("/reparto-xlsx")
def reparto_xlsx(payload: RepartoPdfRequest) -> Response:
    """Reporte mensual por avión en Excel (mismos datos del reparto)."""
    xlsx_bytes = render_reparto_xlsx(payload)
    filename = f"reporte-mensual-{payload.periodo_desde}-a-{payload.periodo_hasta}.xlsx"
    return Response(
        content=xlsx_bytes,
        media_type=XLSX_MEDIA,
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )
