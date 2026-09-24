import base64
import logging
from xml.etree.ElementTree import ParseError

from fastapi import APIRouter, Depends, HTTPException, status

from app.config import get_settings
from app.schemas.facturacion import (
    CancelarRequest,
    CancelarResponse,
    FacturacionHealthResponse,
    FacturaPreviewResponse,
    LeerPdfEmitidaRequest,
    LeerPdfEmitidaResponse,
    TimbrarRequest,
    TimbrarResponse,
)
from app.schemas.recibida import FacturaRecibidaParsed, ParseRecibidaRequest
from app.security import require_internal_token
from app.services.cfdi_fel import cancelar, timbrar
from app.services.emitida_pdf import AVISO_ILEGIBLE, PdfNoValidoError, leer_pdf_emitida
from app.services.factura_preview_pdf import render_factura_preview_pdf
from app.services.facturama import cancelar_facturama, probar_conexion, timbrar_facturama
from app.services.recibida_parse import parse_cfdi

logger = logging.getLogger("facturacion")

router = APIRouter(
    prefix="/facturacion",
    tags=["facturacion"],
    dependencies=[Depends(require_internal_token)],
)


@router.post("/parse-recibida", response_model=FacturaRecibidaParsed)
def parse_recibida(req: ParseRecibidaRequest) -> FacturaRecibidaParsed:
    """Parsea un CFDI recibido (XML de proveedor) y extrae sus datos."""
    try:
        return parse_cfdi(req)
    # Un XML roto, un base64 inválido o un XML con DTD/entidades (rechazado
    # por seguridad) daban 500 sin explicación: 422 con el motivo real.
    except (ValueError, ParseError) as e:
        logger.warning("CFDI recibido no parseable: %s", e)
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail=str(e)[:200] or "El XML no es un CFDI legible",
        ) from e


@router.post("/leer-pdf-emitida", response_model=LeerPdfEmitidaResponse)
def leer_pdf_emitida_endpoint(req: LeerPdfEmitidaRequest) -> LeerPdfEmitidaResponse:
    """Lee el PDF (representación impresa del CFDI) de una factura que la
    oficina EMITIÓ a mano y devuelve lo que encontró para prellenar el
    registro. Determinista (pypdf + regex), sin IA y sin guardar nada.

    422 solo si NO es un PDF (base64 roto, sin `%PDF-`, > 11 MB). Protegido,
    escaneado, roto o lento ⇒ 200 con `texto_extraido=false` + aviso: el API
    degrada a captura manual. Jamás 500.
    """
    try:
        return leer_pdf_emitida(req.pdf_b64)
    except PdfNoValidoError as e:
        # 422 numérico: la constante cambió de nombre entre versiones de Starlette.
        raise HTTPException(status_code=422, detail=str(e)) from e
    except Exception:  # noqa: BLE001 — la lectura degrada, nunca tumba el diálogo
        logger.exception("Error inesperado leyendo el PDF de una factura emitida")
        return LeerPdfEmitidaResponse(texto_extraido=False, avisos=[AVISO_ILEGIBLE])


@router.get("/health", response_model=FacturacionHealthResponse)
def health() -> FacturacionHealthResponse:
    """Prueba las credenciales del PAC sin timbrar (no consume folios)."""
    s = get_settings()
    if not _usa_facturama():
        return FacturacionHealthResponse(
            ok=False,
            pac="fel",
            detalle=(
                "El health-check solo está implementado para Facturama "
                "(pon FACTURACION_PAC=facturama en Railway pyservices)."
            ),
        )
    ok, detalle = probar_conexion()
    return FacturacionHealthResponse(
        ok=ok, pac="facturama", modo=s.facturama_modo, detalle=detalle
    )


def _usa_facturama() -> bool:
    return get_settings().facturacion_pac == "facturama"


@router.post("/preview", response_model=FacturaPreviewResponse)
def preview_factura(req: TimbrarRequest) -> FacturaPreviewResponse:
    """PDF de vista previa de la factura SIN timbrar (sin validez fiscal).
    Mismo body que /timbrar pero NO requiere CSD (los campos csd_* tienen
    default vacío): no se sella nada, solo se renderiza para revisión."""
    try:
        pdf = render_factura_preview_pdf(req)
    except ImportError as e:
        logger.error("WeasyPrint no disponible: %s", e)
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail="Generador de PDF no instalado (WeasyPrint).",
        ) from e
    except Exception as e:  # noqa: BLE001
        logger.exception("Error generando la vista previa de la factura")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"No se pudo generar la vista previa: {e}",
        ) from e
    return FacturaPreviewResponse(pdf_b64=base64.b64encode(pdf).decode("ascii"))


@router.post("/timbrar", response_model=TimbrarResponse)
def timbrar_cfdi(req: TimbrarRequest) -> TimbrarResponse:
    return timbrar_facturama(req) if _usa_facturama() else timbrar(req)


@router.post("/nota-credito", response_model=TimbrarResponse)
def nota_credito(req: TimbrarRequest) -> TimbrarResponse:
    """Nota de crédito = CFDI de Egreso relacionado (tipo 01) a la factura original."""
    req.tipo_comprobante = "E"
    if not req.cfdi_relacionado_uuid:
        return TimbrarResponse(
            ok=False,
            error="La nota de crédito requiere cfdi_relacionado_uuid (UUID de la factura).",
        )
    req.tipo_relacion = req.tipo_relacion or "01"
    return timbrar_facturama(req) if _usa_facturama() else timbrar(req)


@router.post("/cancelar", response_model=CancelarResponse)
def cancelar_cfdi(req: CancelarRequest) -> CancelarResponse:
    # El PAC de cancelación lo decide la FACTURA, no el env: pac_id solo existe
    # cuando se timbró con Facturama; sin él, se timbró con FEL. Si ruteáramos
    # por FACTURACION_PAC y el env cambiara de PAC, una factura timbrada con el
    # PAC anterior se intentaría cancelar donde nunca existió. El env solo
    # decide el PAC de TIMBRADO (nuevas facturas).
    return cancelar_facturama(req) if req.pac_id else cancelar(req)
