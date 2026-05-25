from fastapi import APIRouter, Depends

from app.schemas.facturacion import TimbrarRequest, TimbrarResponse
from app.security import require_internal_token
from app.services.cfdi_fel import timbrar

router = APIRouter(
    prefix="/facturacion",
    tags=["facturacion"],
    dependencies=[Depends(require_internal_token)],
)


@router.post("/timbrar", response_model=TimbrarResponse)
def timbrar_cfdi(req: TimbrarRequest) -> TimbrarResponse:
    return timbrar(req)
