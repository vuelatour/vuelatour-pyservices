import logging
import re
from collections.abc import Callable
from typing import Any

from fastapi import APIRouter, Depends, HTTPException, Response, status
from fastapi.responses import HTMLResponse

from app.schemas.reportes import (
    CotizacionGrupoPdfRequest,
    CotizacionInternaPdfRequest,
    CotizacionPdfRequest,
    MapaSvgRequest,
)
from app.security import require_internal_token
from app.services.cotizacion_grupo_pdf import render_cotizacion_grupo_pdf
from app.services.cotizacion_interna_pdf import render_cotizacion_interna_pdf
from app.services.cotizacion_pdf import (
    _estilos_hoja,
    _mapa_svg_elemento,
    render_cotizacion_pdf,
    render_cotizacion_preview_html,
)

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


@router.post("/cotizacion/preview-html", response_class=HTMLResponse)
def cotizacion_preview_html(req: CotizacionPdfRequest) -> HTMLResponse:
    """VISTA PREVIA de la hoja 1 de la cotización (rediseño del cotizador,
    8-sep-2026): el MISMO payload que `/cotizacion` y el MISMO `_build_html`,
    pero SOLO la hoja 1 con CSS de pantalla (sin `@page` ni hoja "La
    aeronave"; logo y mapa van inline como siempre). Devuelve HTML, no PDF:
    NO importa WeasyPrint. `Cache-Control: no-store` — es el borrador vivo
    del operador, jamás se cachea. Un fallo del armado → 500 con detalle
    (el API lo propaga al panel, que muestra "sin vista previa")."""
    try:
        html = render_cotizacion_preview_html(req)
    except Exception as e:  # noqa: BLE001
        logger.exception("Error generando la vista previa de cotización")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"No se pudo generar la vista previa: {e}",
        ) from e
    return HTMLResponse(content=html, headers={"Cache-Control": "no-store"})


@router.get("/cotizacion/hoja.css")
def cotizacion_hoja_css() -> Response:
    """CSS de la HOJA de cotización para el panel (form-as-document,
    8-sep-2026): `@font-face` incrustado (Arimo) + el cuerpo de la hoja, el
    MISMO texto que llevan el PDF y la vista previa (`_estilos_hoja`). Todo
    selector cuelga de `.cot-hoja`. Es estático por deploy: cacheable 1 h."""
    return Response(
        content=_estilos_hoja(),
        media_type="text/css; charset=utf-8",
        headers={"Cache-Control": "public, max-age=3600"},
    )


@router.post("/cotizacion/mapa-svg")
def cotizacion_mapa_svg(req: MapaSvgRequest) -> Response:
    """Mapa del itinerario como `image/svg+xml` para la hoja del panel
    (8-sep-2026): el MISMO `_mapa_svg` del PDF (modo local/amplio, badges
    numerados por `orden`, ferry punteado) sin el envoltorio `.mapa` — el
    panel lo inyecta inline dentro de su `<div class="mapa">`. Sin puntos →
    204 (la hoja no lleva mapa)."""
    svg = _mapa_svg_elemento(req.mapa_puntos)
    if not svg:
        return Response(status_code=status.HTTP_204_NO_CONTENT)
    return Response(
        content=svg,
        media_type="image/svg+xml",
        headers={"Cache-Control": "no-store"},
    )


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


@router.post("/cotizacion-interna")
def cotizacion_interna_pdf(req: CotizacionInternaPdfRequest) -> Response:
    """PDF de la COTIZACIÓN INTERNA v2 (8-sep-2026): UNA hoja carta para la
    oficina con SOLO lo de la cotización — fecha del vuelo, matrícula, tabla
    de tramos (ruta · fecha · millas · tiempo con calzos · costo/hr · total),
    desglose con comisión del vendedor, TUAS cobradas, cobros con comisión
    bancaria y notas internas. NUNCA va al cliente. El API manda todo YA
    calculado; aquí SOLO se pinta."""
    pdf = _render_o_http(render_cotizacion_interna_pdf, req, "cotización interna")
    folio = re.sub(r"[^A-Za-z0-9_-]+", "", req.folio or "") or "sn"
    return Response(
        content=pdf,
        media_type="application/pdf",
        headers={"Content-Disposition": f'inline; filename="cotizacion-interna-{folio}.pdf"'},
    )
