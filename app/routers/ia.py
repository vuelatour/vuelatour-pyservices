"""Configuración de IA visible para el API (2-oct-2026).

`GET /ia/modelo` → `{default_servidor, efectivo}`. El API la consulta
(best-effort) para pintar en Configuración «default del servidor» y para
confirmar qué modelo resulta con el header `X-IA-Modelo` que manda. Aquí no
se guarda nada: la elección vive en el API (`configuracion_sistema.ia_modelo`).
"""

from fastapi import APIRouter, Depends

from app.schemas.ia import IaModeloResponse
from app.security import require_internal_token
from app.services.modelo_ia import modelo_actual, modelo_del_servidor

router = APIRouter(prefix="/ia", tags=["ia"], dependencies=[Depends(require_internal_token)])


@router.get("/modelo", response_model=IaModeloResponse)
def modelo() -> IaModeloResponse:
    return IaModeloResponse(default_servidor=modelo_del_servidor(), efectivo=modelo_actual())
