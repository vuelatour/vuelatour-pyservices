from fastapi import APIRouter, Depends

from app.schemas.facturacion import (
    CancelarRequest,
    CancelarResponse,
    TimbrarRequest,
    TimbrarResponse,
)
from app.schemas.recibida import FacturaRecibidaParsed, ParseRecibidaRequest
from app.security import require_internal_token
from app.services.cfdi_fel import cancelar, timbrar
from app.services.recibida_parse import parse_cfdi

router = APIRouter(
    prefix="/facturacion",
    tags=["facturacion"],
    dependencies=[Depends(require_internal_token)],
)


@router.post("/parse-recibida", response_model=FacturaRecibidaParsed)
def parse_recibida(req: ParseRecibidaRequest) -> FacturaRecibidaParsed:
    """Parsea un CFDI recibido (XML de proveedor) y extrae sus datos."""
    return parse_cfdi(req)


@router.post("/timbrar", response_model=TimbrarResponse)
def timbrar_cfdi(req: TimbrarRequest) -> TimbrarResponse:
    return timbrar(req)


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
    return timbrar(req)


@router.post("/cancelar", response_model=CancelarResponse)
def cancelar_cfdi(req: CancelarRequest) -> CancelarResponse:
    return cancelar(req)
