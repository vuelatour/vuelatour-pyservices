"""Endpoints de generacion de PDF. Protegidos con el token interno (X-Internal-Token)."""

from fastapi import APIRouter, Depends, HTTPException, Response, status

from app.schemas.reparto import RepartoPdfRequest
from app.schemas.reportes import BalanceAvionRequest, ReporteVueloRequest
from app.schemas.tabla import TablaXlsxRequest
from app.schemas.zip import ZipRequest
from app.security import require_internal_token
from app.services.balance_avion_xlsx import render_balance_avion_xlsx
from app.services.reparto_pdf import render_reparto_pdf
from app.services.reparto_xlsx import render_reparto_xlsx
from app.services.reporte_vuelo_pdf import render_reporte_vuelo_pdf
from app.services.reporte_vuelo_xlsx import render_reporte_vuelo_xlsx
from app.services.tabla_xlsx import render_tabla_xlsx
from app.services.zip_service import render_zip

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


@router.post("/tabla-xlsx")
def tabla_xlsx(payload: TablaXlsxRequest) -> Response:
    """Export genérico de cualquier tabla a Excel (inventario, cardex, horas...)."""
    xlsx_bytes = render_tabla_xlsx(payload)
    return Response(
        content=xlsx_bytes,
        media_type=XLSX_MEDIA,
        headers={"Content-Disposition": 'attachment; filename="reporte.xlsx"'},
    )


@router.post("/reporte-vuelo")
def reporte_vuelo_pdf(payload: ReporteVueloRequest) -> Response:
    """Reporte consolidado de un vuelo (cotización, ingreso, tacómetro, gastos)."""
    pdf_bytes = render_reporte_vuelo_pdf(payload)
    return Response(
        content=pdf_bytes,
        media_type="application/pdf",
        headers={"Content-Disposition": f'inline; filename="vuelo-{payload.folio}.pdf"'},
    )


@router.post("/reporte-vuelo-xlsx")
def reporte_vuelo_xlsx(payload: ReporteVueloRequest) -> Response:
    """Reporte consolidado de un vuelo en Excel."""
    xlsx_bytes = render_reporte_vuelo_xlsx(payload)
    return Response(
        content=xlsx_bytes,
        media_type=XLSX_MEDIA,
        headers={"Content-Disposition": f'attachment; filename="vuelo-{payload.folio}.xlsx"'},
    )


@router.post("/balance-avion-xlsx")
def balance_avion_xlsx(payload: BalanceAvionRequest) -> Response:
    """Balance mensual por avión en Excel (réplica del control del equipo).

    El API manda todo precalculado; aquí SOLO se renderiza el libro.
    """
    xlsx_bytes = render_balance_avion_xlsx(payload)
    filename = (
        f"balance-{payload.matricula or 'avion'}-"
        f"{payload.periodo_desde or 's-f'}-{payload.periodo_hasta or 's-f'}.xlsx"
    )
    return Response(
        content=xlsx_bytes,
        media_type=XLSX_MEDIA,
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@router.post("/zip")
def zip_archivos(payload: ZipRequest) -> Response:
    """Ensambla archivos (base64) en un .zip. Usado por el cierre mensual."""
    try:
        zip_bytes = render_zip(payload)
    except ValueError as e:
        # Archivos ilegibles: mejor un error claro que un cierre incompleto.
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY, detail=str(e)
        ) from e
    return Response(
        content=zip_bytes,
        media_type="application/zip",
        headers={"Content-Disposition": 'attachment; filename="cierre.zip"'},
    )
