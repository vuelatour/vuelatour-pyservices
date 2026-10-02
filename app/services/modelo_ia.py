"""Modelo de Claude POR PETICIÓN (2-oct-2026, API 0.0.51).

Pedido del cliente (Configuración → Créditos de IA): «dejar una opción en la
configuración para adaptar el modelo que quieran utilizar, aunque ahorita
dejaremos por default el que estamos usando actualmente».

La elección vive en el API (`configuracion_sistema`, clave `ia_modelo`): este
servicio NO guarda nada. El API manda el id en el header `X-IA-Modelo` de
cada petición SOLO cuando hay modelo configurado; `require_internal_token`
(app/security.py) lo deja en el ContextVar `modelo_ia_peticion` —un valor
NUEVO en cada petición: sin header ⇒ None, nunca el de la anterior— y
`modelo_actual()` es la ÚNICA fuente del `model=` de cada `messages.create`
y del campo `modelo` de las respuestas.

Reglas:
- Header ausente, vacío o que no cumple `PATRON_ID_MODELO` (la MISMA regex
  que valida el API al guardar) ⇒ el del servidor (`ANTHROPIC_MODEL`). Un API
  viejo, que no manda el header, funciona exactamente como antes.
- Aquí no hay catálogo ni tarifas: un id válido fuera del catálogo del API se
  usa tal cual (si Anthropic no lo conoce responde 404 y `ia_errores` lo
  traduce a texto para el operador).
- El ContextVar NO cruza a hilos de un `ThreadPoolExecutor`: quien reparta
  trabajo en hilos propios resuelve `modelo_actual()` ANTES y pasa el id
  (ver `estado_cuenta._leer_por_bloques`).
"""

import re
from contextvars import ContextVar

from app.config import get_settings

# Misma regex que `esIdModeloValido` del API (src/common/ia-modelo.util.ts) y
# del panel (lib/admin/ia-modelo.ts). Se evalúa con `fullmatch`: un salto de
# línea al final NO pasa.
PATRON_ID_MODELO = re.compile(r"^claude-[a-z0-9.-]{3,80}$")

# Modelo pedido por la petición en curso (el header crudo; se valida al leer).
modelo_ia_peticion: ContextVar[str | None] = ContextVar("modelo_ia_peticion", default=None)


def es_id_modelo_valido(valor: object) -> bool:
    """True si `valor` es un id de modelo aceptable (`claude-…`)."""
    return isinstance(valor, str) and PATRON_ID_MODELO.fullmatch(valor) is not None


def modelo_pedido(valor: object) -> str | None:
    """El id pedido ya limpio (sin espacios a los lados) si es válido; si no, None."""
    if not isinstance(valor, str):
        return None
    limpio = valor.strip()
    return limpio if es_id_modelo_valido(limpio) else None


def fijar_modelo_de_peticion(valor: str | None) -> None:
    """Deja el modelo de ESTA petición (también None: borra el de cualquier otra).

    La llama `require_internal_token` en TODAS las peticiones. Debe correr en
    el contexto de la petición (dependencia `async def`): una dependencia
    `def` corre en un hilo con una COPIA del contexto y el valor se pierde."""
    modelo_ia_peticion.set(valor)


def modelo_del_servidor() -> str:
    """El default del servidor: variable de entorno `ANTHROPIC_MODEL`."""
    return get_settings().anthropic_model


def modelo_actual() -> str:
    """Modelo a usar en esta petición: el del header si es válido, si no el del servidor."""
    return modelo_pedido(modelo_ia_peticion.get()) or modelo_del_servidor()
