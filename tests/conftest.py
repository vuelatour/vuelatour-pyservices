"""Doble de prueba del SDK de Anthropic (15-sep-2026).

Ninguna prueba de este repo llama a Claude de verdad: se instala un cliente
falso que devuelve el JSON que queremos y guarda los kwargs con los que se
llamó (para verificar el payload y los bloques `system`).
"""

import json

import pytest


class _Bloque:
    def __init__(self, texto: str) -> None:
        self.type = "text"
        self.text = texto


class _Uso:
    input_tokens = 120
    output_tokens = 45
    cache_creation_input_tokens = 0
    cache_read_input_tokens = 0


class RespuestaFake:
    def __init__(self, payload, stop_reason: str = "end_turn") -> None:
        texto = payload if isinstance(payload, str) else json.dumps(payload, ensure_ascii=False)
        self.content = [_Bloque(texto)]
        self.stop_reason = stop_reason
        self.model = "claude-sonnet-4-6"
        self.usage = _Uso()


class _Messages:
    def __init__(self, cliente: "ClienteFake") -> None:
        self._cliente = cliente

    def create(self, **kwargs):
        self._cliente.llamadas.append(kwargs)
        return self._cliente.respuesta


class ClienteFake:
    """Sustituto de `anthropic.Anthropic` con `messages.create` y `with_options`."""

    def __init__(self, respuesta: RespuestaFake) -> None:
        self.respuesta = respuesta
        self.llamadas: list[dict] = []

    @property
    def messages(self) -> _Messages:
        return _Messages(self)

    def with_options(self, **_kwargs) -> "ClienteFake":
        return self

    # --- helpers de aserción -------------------------------------------
    @property
    def ultima(self) -> dict:
        assert self.llamadas, "no se llamó a Claude"
        return self.llamadas[-1]

    def texto_usuario(self) -> str:
        contenido = self.ultima["messages"][0]["content"]
        if isinstance(contenido, str):
            return contenido
        return "\n".join(b.get("text", "") for b in contenido if isinstance(b, dict))

    def texto_system(self) -> str:
        return "\n".join(b["text"] for b in self.ultima["system"])


@pytest.fixture
def claude_fake(monkeypatch):
    """`instalar(modulo, payload)` → ClienteFake ya parcheado en ese módulo."""

    def instalar(modulo, payload, stop_reason: str = "end_turn") -> ClienteFake:
        cliente = ClienteFake(RespuestaFake(payload, stop_reason))
        monkeypatch.setattr(modulo, "_client", lambda: cliente)
        return cliente

    return instalar
