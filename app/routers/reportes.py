import logging
import re
from collections.abc import Callable
from typing import Any

from fastapi import APIRouter, Depends, HTTPException, Response, status

from app.schemas.reportes import CotizacionGrupoPdfRequest, CotizacionPdfRequest
from app.security import require_internal_token
from app.services.cotizacion_grupo_pdf import render_cotizacion_grupo_pdf
from app.services.cotizacion_pdf import render_cotizacion_pdf

logger = logging.getLogger("reportes")

router = APIRouter(
    prefix="/reportes",
    tags=["reportes"],
    dependencies=[Depends(require_internal_token)],
)


def _render_o_http(render: Callable[[Any], bytes], req: Any, que: str) -> bytes:
    """Corre el render y traduce sus fallos a HTTP: WeasyPrint ausente → 503,
    cualquier otro error → 500 con el mensaje (el API lo propaga al panel)."""
    try:
        return render(req)
    except ImportError as e:
        logger.error("WeasyPrint no disponible: %s", e)
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail="Generador de PDF no instalado (WeasyPrint).",
        ) from e
    except Exception as e:  # noqa: BLE001
        logger.exception("Error generando PDF de %s", que)
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"No se pudo generar el PDF: {e}",
        ) from e


@router.post("/cotizacion")
def cotizacion_pdf(req: CotizacionPdfRequest) -> Response:
    pdf = _render_o_http(render_cotizacion_pdf, req, "cotización")
    return Response(content=pdf, media_type="application/pdf")


@router.post("/cotizacion-grupo")
def cotizacion_grupo_pdf(req: CotizacionGrupoPdfRequest) -> Response:
    """PDF ÚNICO de la cotización de GRUPO (4-sep-2026): varios aviones para
    un mismo cliente con total consolidado, hoja "Flota asignada" y fichas
    por modelo. El API manda todo YA calculado; aquí SOLO se pinta."""
    pdf = _render_o_http(render_cotizacion_grupo_pdf, req, "cotización de grupo")
    folio = req.folio_grupo or (f"G-{req.folio}" if req.folio else "G")
    folio = re.sub(r"[^A-Za-z0-9_-]+", "", folio) or "G"
    return Response(
        content=pdf,
        media_type="application/pdf",
        headers={"Content-Disposition": f'inline; filename="cotizacion-grupo-{folio}.pdf"'},
    )
