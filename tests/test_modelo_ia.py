"""Modelo de Claude por petición: header `X-IA-Modelo` (2-oct-2026, API 0.0.51).

Pedido: «dejar una opción en la configuración para adaptar el modelo que
quieran utilizar, aunque ahorita dejaremos por default el que estamos usando».
El API manda el id elegido en Configuración en el header `X-IA-Modelo` (solo
cuando hay uno configurado); aquí se congela:

- `modelo_actual()`: header válido ⇒ ese; inválido/ausente ⇒ `ANTHROPIC_MODEL`;
- la dependencia es `async def` (una `def` corre en un hilo con COPIA del
  contexto y el modelo nunca llegaría al endpoint) y fija SIEMPRE el valor:
  sin header ⇒ None, jamás el de la petición anterior;
- aislamiento entre peticiones seguidas (en la MISMA tarea) y SIMULTÁNEAS (en
  un solo loop) con un router de visión real vía `httpx.ASGITransport` y el
  cliente de Anthropic simulado que captura `model` — nunca con `TestClient`
  sin `with`, que abre un loop y un contexto nuevos por petición;
- `GET /ia/modelo` → `{default_servidor, efectivo}`.
"""

import ast
import asyncio
import contextvars
import inspect
import json
import logging
import threading
from pathlib import Path
from types import SimpleNamespace

import httpx
import pytest
from fastapi.testclient import TestClient

from app.config import get_settings
from app.main import app
from app.security import require_internal_token
from app.services import anthropic_vision
from app.services.modelo_ia import (
    PATRON_ID_MODELO,
    es_id_modelo_valido,
    fijar_modelo_de_peticion,
    modelo_actual,
    modelo_del_servidor,
    modelo_ia_peticion,
    modelo_pedido,
)

TOKEN = "secreto-de-prueba"
# Distinto de cualquier modelo real: si una prueba lo ve, vino del servidor.
MODELO_SERVIDOR = "claude-servidor-prueba"
CATALOGO_API = (
    "claude-opus-4-8",
    "claude-opus-5-5",
    "claude-sonnet-5",
    "claude-sonnet-4-6",
    "claude-haiku-4-5-20251001",
)

# Solo para pruebas de UNA petición: sin `with`, cada llamada corre en su propio
# hilo, loop y contexto (ver «aislamiento entre peticiones» más abajo).
client = TestClient(app)

TICKET = {
    "monto": 939.60,
    "moneda": "MXN",
    "fecha": "2026-09-07",
    "proveedor": "Aeropuertos del Sureste",
    "concepto": "TUA",
    "categoria_sugerida": "TUAS",
    "confianza": 0.93,
    "legible": True,
    "notas": "",
}
IMG = {"image_base64": "AAAA", "media_type": "image/jpeg"}


@pytest.fixture
def entorno(monkeypatch, request):
    """Token de prueba y modelo del servidor conocido (no el de `.env.local`)."""
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    monkeypatch.setenv("ANTHROPIC_MODEL", MODELO_SERVIDOR)
    get_settings.cache_clear()
    # Aunque una aserción falle, el token de prueba no se queda en caché.
    request.addfinalizer(get_settings.cache_clear)


def _en_contexto_limpio(fn):
    """Corre `fn` en una COPIA del contexto: lo que fije no contamina otras pruebas."""
    return contextvars.copy_context().run(fn)


def _con_modelo(valor):
    def correr():
        fijar_modelo_de_peticion(valor)
        return modelo_actual()

    return _en_contexto_limpio(correr)


# --- Fuente única: regex y modelo_actual -----------------------------------


def test_regex_es_la_del_contrato() -> None:
    # Copia literal de `esIdModeloValido` (API) y de lib/admin/ia-modelo.ts (panel).
    assert PATRON_ID_MODELO.pattern == r"^claude-[a-z0-9.-]{3,80}$"


@pytest.mark.parametrize("modelo", CATALOGO_API)
def test_todo_el_catalogo_del_api_es_valido(modelo) -> None:
    assert es_id_modelo_valido(modelo)


@pytest.mark.parametrize(
    "valor",
    [
        "",
        "gpt-4o",
        "claude-",
        "claude-ab",  # menos de 3 caracteres tras «claude-»
        "Claude-Opus-4-8",  # mayúsculas
        "claude-opus 4-8",  # espacio interno
        "claude-opus-4-8\n",  # salto de línea al final
        "claude-opus_4_8",  # guion bajo
        "claude-" + "a" * 81,  # más de 80
        None,
        123,
    ],
)
def test_ids_invalidos(valor) -> None:
    assert not es_id_modelo_valido(valor)


def test_limites_de_largo() -> None:
    assert es_id_modelo_valido("claude-abc")
    assert es_id_modelo_valido("claude-" + "a" * 80)


def test_sin_peticion_es_el_del_servidor(entorno) -> None:
    assert modelo_del_servidor() == MODELO_SERVIDOR
    assert _en_contexto_limpio(modelo_actual) == MODELO_SERVIDOR
    assert _en_contexto_limpio(modelo_ia_peticion.get) is None


def test_header_valido_gana(entorno) -> None:
    assert _con_modelo("claude-sonnet-5") == "claude-sonnet-5"
    # Espacios a los lados no lo invalidan.
    assert _con_modelo("  claude-haiku-4-5-20251001 ") == "claude-haiku-4-5-20251001"
    assert modelo_pedido(" claude-opus-5-5 ") == "claude-opus-5-5"


@pytest.mark.parametrize("valor", [None, "", "   ", "gpt-4o", "Claude-Opus-4-8", "claude-x"])
def test_header_invalido_o_ausente_es_el_del_servidor(entorno, valor) -> None:
    assert _con_modelo(valor) == MODELO_SERVIDOR


# --- La dependencia --------------------------------------------------------


def test_dependencia_es_async_para_que_el_modelo_llegue_al_endpoint() -> None:
    # Una dependencia `def` corre en un hilo con COPIA del contexto: el
    # ContextVar fijado ahí se pierde y TODO usaría el modelo del servidor.
    assert inspect.iscoroutinefunction(require_internal_token)


def test_dependencia_sin_header_borra_el_de_la_peticion_anterior(entorno) -> None:
    async def escenario():
        fijar_modelo_de_peticion("claude-prueba-anterior")
        await require_internal_token(x_internal_token=TOKEN, x_ia_modelo=None)
        return modelo_ia_peticion.get(), modelo_actual()

    assert asyncio.run(escenario()) == (None, MODELO_SERVIDOR)


def test_header_invalido_se_avisa_en_el_log_solo_con_token(entorno, caplog) -> None:
    with caplog.at_level(logging.WARNING, logger="security"):
        sin_token = client.get("/ia/modelo", headers={"X-IA-Modelo": "gpt-4o"})
        assert sin_token.status_code == 401
        assert "X-IA-Modelo" not in caplog.text
        con_token = client.get(
            "/ia/modelo", headers={"X-Internal-Token": TOKEN, "X-IA-Modelo": "gpt-4o"}
        )
    assert con_token.status_code == 200
    assert "X-IA-Modelo inválido" in caplog.text
    assert "gpt-4o" in caplog.text
    assert TOKEN not in caplog.text


# --- GET /ia/modelo ---------------------------------------------------------


def test_ia_modelo_exige_token(entorno) -> None:
    assert client.get("/ia/modelo").status_code == 401
    assert client.get("/ia/modelo", headers={"X-Internal-Token": "otro"}).status_code == 401


def test_ia_modelo_sin_header_es_el_del_servidor(entorno) -> None:
    res = client.get("/ia/modelo", headers={"X-Internal-Token": TOKEN})
    assert res.status_code == 200
    assert res.json() == {"default_servidor": MODELO_SERVIDOR, "efectivo": MODELO_SERVIDOR}


def test_ia_modelo_con_header(entorno) -> None:
    res = client.get(
        "/ia/modelo", headers={"X-Internal-Token": TOKEN, "X-IA-Modelo": "claude-sonnet-5"}
    )
    assert res.json() == {"default_servidor": MODELO_SERVIDOR, "efectivo": "claude-sonnet-5"}
    invalido = client.get(
        "/ia/modelo", headers={"X-Internal-Token": TOKEN, "X-IA-Modelo": "gpt-4o"}
    )
    assert invalido.json() == {"default_servidor": MODELO_SERVIDOR, "efectivo": MODELO_SERVIDOR}


# --- Router de visión real: aislamiento entre peticiones --------------------
#
# Estas pruebas NO usan `client` (TestClient sin `with`): Starlette abre un
# portal —hilo, event loop y contexto NUEVOS— por cada petición, así que una
# fuga del modelo de una petición a la siguiente jamás se vería (revisión
# 2-oct-2026: con la dependencia fijando el ContextVar solo cuando llega el
# header, las pruebas con TestClient seguían en verde). `httpx.ASGITransport`
# corre la app en el MISMO loop y la MISMA tarea que la prueba, como uvicorn
# atiende varias peticiones (keep-alive) en un solo loop.


def _cliente_asgi() -> httpx.AsyncClient:
    return httpx.AsyncClient(transport=httpx.ASGITransport(app=app), base_url="http://pyservices")


def _headers(modelo: str | None) -> dict[str, str]:
    headers = {"X-Internal-Token": TOKEN}
    if modelo is not None:
        headers["X-IA-Modelo"] = modelo
    return headers


async def _leer_gasto(http: httpx.AsyncClient, modelo: str | None) -> dict:
    res = await http.post("/vision/gasto", json=IMG, headers=_headers(modelo))
    assert res.status_code == 200, res.text
    return res.json()


def test_vision_gasto_peticiones_seguidas_en_la_misma_tarea(entorno, claude_fake) -> None:
    """A → sin header → inválido → B → sin header, todas en UNA tarea: el
    contexto es el mismo para las cinco y solo la dependencia (que fija
    SIEMPRE, None sin header) evita que una arrastre el modelo de la anterior."""
    cliente = claude_fake(anthropic_vision, TICKET)
    pedidos = ["claude-prueba-uno", None, "gpt-4o", "claude-prueba-dos", None]

    async def escenario():
        async with _cliente_asgi() as http:
            respuestas = [await _leer_gasto(http, m) for m in pedidos]
            ia_modelo = []
            for m in ("claude-prueba-tres", None):
                res = await http.get("/ia/modelo", headers=_headers(m))
                ia_modelo.append(res.json()["efectivo"])
        # Lo que quedó en el contexto de ESTA tarea tras la última petición.
        return respuestas, ia_modelo, modelo_ia_peticion.get()

    respuestas, ia_modelo, queda = asyncio.run(escenario())

    esperados = [
        "claude-prueba-uno",
        MODELO_SERVIDOR,
        MODELO_SERVIDOR,
        "claude-prueba-dos",
        MODELO_SERVIDOR,
    ]
    # `model=` de cada messages.create: uno por petición, sin arrastre.
    assert [k["model"] for k in cliente.llamadas] == esperados
    # El campo `modelo` de la respuesta dice el mismo.
    assert [r["modelo"] for r in respuestas] == esperados
    # `uso_ia` sigue saliendo de `resp.model` (el modelo REAL servido).
    assert respuestas[0]["uso_ia"]["modelo"] == "claude-sonnet-4-6"
    # GET /ia/modelo en la misma tarea: con header y luego sin él.
    assert ia_modelo == ["claude-prueba-tres", MODELO_SERVIDOR]
    assert queda is None
    # Y nada se filtró al contexto de la prueba.
    assert modelo_ia_peticion.get() is None


def _respuesta():
    return SimpleNamespace(
        content=[SimpleNamespace(type="text", text=json.dumps(TICKET))],
        stop_reason="end_turn",
        model="claude-sonnet-4-6",
        usage=SimpleNamespace(
            input_tokens=1,
            output_tokens=1,
            cache_creation_input_tokens=0,
            cache_read_input_tokens=0,
        ),
    )


class _ClienteConBarrera:
    """`messages.create` que espera a que TODAS las peticiones estén en vuelo.

    El endpoint es `def` (corre en el threadpool), así que la barrera de hilos
    no bloquea el loop."""

    def __init__(self, partes: int) -> None:
        self.barrera = threading.Barrier(partes, timeout=10)
        self.modelos: list[str] = []
        self._lock = threading.Lock()

    @property
    def messages(self) -> "_ClienteConBarrera":
        return self

    def with_options(self, **_kwargs) -> "_ClienteConBarrera":
        return self

    def create(self, **kwargs):
        with self._lock:
            self.modelos.append(kwargs["model"])
        self.barrera.wait()
        return _respuesta()


def test_peticiones_simultaneas_en_un_loop_no_se_cruzan_el_modelo(entorno, monkeypatch) -> None:
    """Tres peticiones a la vez en UN loop (como uvicorn: una tarea por
    petición). El `modelo` de la respuesta se lee DESPUÉS de la barrera, cuando
    las tres dependencias ya corrieron: un estado compartido (variable global en
    vez de ContextVar) haría que al menos dos respuestas digan el último."""
    cliente = _ClienteConBarrera(partes=3)
    monkeypatch.setattr(anthropic_vision, "_client", lambda: cliente)
    pedidos = ["claude-prueba-uno", "claude-prueba-dos", None]

    async def escenario():
        async with _cliente_asgi() as http:
            return await asyncio.gather(*(_leer_gasto(http, m) for m in pedidos))

    resultados = asyncio.run(escenario())

    assert sorted(cliente.modelos) == sorted(
        ["claude-prueba-uno", "claude-prueba-dos", MODELO_SERVIDOR]
    )
    assert [r["modelo"] for r in resultados] == [
        "claude-prueba-uno",
        "claude-prueba-dos",
        MODELO_SERVIDOR,
    ]


# --- Fuente única en el código ----------------------------------------------

APP = Path(__file__).resolve().parent.parent / "app"


def _modulos():
    for ruta in sorted(APP.rglob("*.py")):
        yield ruta, ast.parse(ruta.read_text(encoding="utf-8"))


def test_nadie_lee_anthropic_model_salvo_modelo_ia() -> None:
    """`settings.anthropic_model` directo ignora lo elegido en Configuración."""
    culpables = [
        f"{ruta.relative_to(APP.parent)}:{nodo.lineno}"
        for ruta, arbol in _modulos()
        if ruta.name != "modelo_ia.py"
        for nodo in ast.walk(arbol)
        if isinstance(nodo, ast.Attribute) and nodo.attr == "anthropic_model"
    ]
    assert culpables == []


def test_toda_llamada_a_claude_usa_modelo_actual() -> None:
    """Cada `messages.create(model=…)` usa `modelo_actual()` o el id que se
    resolvió con él antes de pasar a un hilo (`modelo`)."""
    llamadas = []
    for ruta, arbol in _modulos():
        for nodo in ast.walk(arbol):
            if not (
                isinstance(nodo, ast.Call)
                and isinstance(nodo.func, ast.Attribute)
                and nodo.func.attr == "create"
                and isinstance(nodo.func.value, ast.Attribute)
                and nodo.func.value.attr == "messages"
            ):
                continue
            modelo = next((k.value for k in nodo.keywords if k.arg == "model"), None)
            ok = (
                isinstance(modelo, ast.Call)
                and isinstance(modelo.func, ast.Name)
                and modelo.func.id == "modelo_actual"
            ) or (isinstance(modelo, ast.Name) and modelo.id == "modelo")
            llamadas.append((f"{ruta.name}:{nodo.lineno}", ok))
    # 5 de visión + compras + vencimientos + abonos + gasto→vuelo + 2 de estado de cuenta.
    assert len(llamadas) == 11, llamadas
    assert [donde for donde, ok in llamadas if not ok] == []
