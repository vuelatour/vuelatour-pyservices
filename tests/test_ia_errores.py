"""Errores de Claude → texto para el operador (1-oct-2026).

Caso real: «IA no disponible: Claude no disponible (400). Captura a mano.»
era el SALDO de créditos agotado. Aquí se congelan los textos de
`app/services/ia_errores.py` con errores REALES del SDK de Anthropic
(construidos sobre `httpx.Response`, igual que los arma el cliente) y se
comprueba que los routers los usan.
"""

import ast
import base64
import logging
import pathlib

import anthropic
import httpx
import pytest
from fastapi.testclient import TestClient

from app.config import get_settings
from app.main import app
from app.routers import vision as vision_router
from app.schemas.vencimiento import VencimientoExtraerRequest
from app.schemas.vision import ConstanciaFiscalRequest, GastoTicketRequest
from app.services import anthropic_vision, ia_errores, vencimiento_extract
from app.services.ia_errores import (
    TEXTO_ARCHIVO_PESADO,
    TEXTO_DEMASIADAS_FOTOS,
    TEXTO_DOCUMENTO_LARGO,
    TEXTO_FORMATO_DISTINTO,
    TEXTO_FORMATO_IMAGEN,
    TEXTO_FOTO_DIMENSIONES,
    TEXTO_FOTO_ILEGIBLE,
    TEXTO_IA_CAIDA,
    TEXTO_LIMITE_GASTO,
    TEXTO_LIMITE_PETICIONES,
    TEXTO_LLAVE_INVALIDA,
    TEXTO_MODELO_INEXISTENTE,
    TEXTO_SIN_SALDO,
    detalle_error_claude,
    detalle_error_claude_captura_manual,
    mensaje_error_claude,
    registrar_error_claude,
    tipo_error_claude,
)
from app.services.imagen_media_type import media_type_real

URL = "https://api.anthropic.com/v1/messages"
TOKEN = "secreto-de-prueba"
LLAVE_FALSA = "sk-ant-llave-que-jamas-debe-salir"

MENSAJE_SIN_SALDO = (
    "Your credit balance is too low to access the Anthropic API. "
    "Please go to Plans & Billing to upgrade or purchase credits."
)


def _error(
    clase: type[anthropic.APIStatusError],
    status: int,
    mensaje: str,
    tipo: str = "invalid_request_error",
    *,
    con_cuerpo: bool = True,
) -> anthropic.APIStatusError:
    """Error del SDK tal como lo arma `_make_status_error_from_response`."""
    cuerpo = {"type": "error", "error": {"type": tipo, "message": mensaje}}
    respuesta = httpx.Response(
        status,
        json=cuerpo,
        headers={"request-id": "req_prueba_123"},
        request=httpx.Request("POST", URL, headers={"x-api-key": LLAVE_FALSA}),
    )
    return clase(
        message=f"Error code: {status} - {cuerpo}",
        response=respuesta,
        body=cuerpo if con_cuerpo else None,
    )


def _error_crudo(status: int, texto: str) -> anthropic.APIStatusError:
    """Respuesta que NO era JSON: el SDK deja el texto en `body` y `message`."""
    respuesta = httpx.Response(status, text=texto, request=httpx.Request("POST", URL))
    return anthropic.APIStatusError(message=texto, response=respuesta, body=texto)


def _error_del_sdk(status: int, cuerpo: dict | str) -> anthropic.APIStatusError:
    """El error que arma el CLIENTE REAL del SDK (sin red: `MockTransport`),
    para no depender de cómo creemos que se construye."""

    def responder(_peticion: httpx.Request) -> httpx.Response:
        if isinstance(cuerpo, dict):
            return httpx.Response(status, json=cuerpo)
        return httpx.Response(status, text=cuerpo)

    cliente = anthropic.Anthropic(
        api_key=LLAVE_FALSA,
        max_retries=0,
        http_client=httpx.Client(transport=httpx.MockTransport(responder)),
    )
    with pytest.raises(anthropic.APIStatusError) as info:
        cliente.messages.create(
            model="claude-de-prueba",
            max_tokens=10,
            messages=[{"role": "user", "content": "hola"}],
        )
    return info.value


def _cuerpo(mensaje: str, tipo: str = "invalid_request_error") -> dict:
    return {"type": "error", "error": {"type": tipo, "message": mensaje}}


# --- Saldo agotado: el caso real -------------------------------------------


def test_sin_saldo_caso_real_400() -> None:
    e = _error(anthropic.BadRequestError, 400, MENSAJE_SIN_SALDO)
    assert detalle_error_claude(e) == (
        "Sin saldo de créditos de IA en Anthropic: hay que recargar (Plans & Billing) "
        "y registrar el nuevo saldo en Configuración → Consumo de IA"
    )
    assert detalle_error_claude(e) == TEXTO_SIN_SALDO


def test_sin_saldo_con_el_cliente_real_del_sdk() -> None:
    e = _error_del_sdk(400, _cuerpo(MENSAJE_SIN_SALDO))
    assert isinstance(e, anthropic.BadRequestError)
    assert detalle_error_claude(e) == TEXTO_SIN_SALDO


def test_como_lo_ve_la_app() -> None:
    """La app arma «IA no disponible: <motivo>. Captura a mano o reintenta.»
    (gasto_screen.dart): sin «..» ni «captura a mano» repetido."""
    e = _error(anthropic.BadRequestError, 400, MENSAJE_SIN_SALDO)
    aviso = f"IA no disponible: {detalle_error_claude(e)}. Captura a mano o reintenta."
    assert ".." not in aviso
    assert aviso.lower().count("captura a mano") == 1


def test_sin_saldo_sin_distinguir_mayusculas() -> None:
    e = _error(anthropic.BadRequestError, 400, "YOUR CREDIT BALANCE IS TOO LOW")
    assert detalle_error_claude(e) == TEXTO_SIN_SALDO


def test_sin_cuerpo_forma_real_del_sdk() -> None:
    # body None (respuesta cerrada antes de leerla): el SDK solo deja
    # «Error code: 400», sin texto. No se repite el código.
    respuesta = httpx.Response(400, request=httpx.Request("POST", URL))
    e = anthropic.BadRequestError("Error code: 400", response=respuesta, body=None)
    assert detalle_error_claude(e) == "Claude no disponible (400)"


def test_cuerpo_sin_error_message_no_muestra_el_dict() -> None:
    e = _error_del_sdk(400, {"type": "error", "error": {"type": "invalid_request_error"}})
    assert e.message.startswith("Error code: 400 - {")
    assert detalle_error_claude(e) == "Claude no disponible (400)"


# --- Reglas por texto -------------------------------------------------------


@pytest.mark.parametrize(
    ("mensaje", "esperado"),
    [
        (
            "messages.0.content.0.image.source.base64: image exceeds 5 MB maximum: "
            "6291456 bytes > 5242880 bytes",
            TEXTO_ARCHIVO_PESADO,
        ),
        ("The PDF file is too large", TEXTO_ARCHIVO_PESADO),
        ("Payload exceeds 5 MB per image", TEXTO_ARCHIVO_PESADO),
        ("Maximum file size exceeded", TEXTO_ARCHIVO_PESADO),
        (
            "messages.0.content.0.image.source.base64.media_type: Input should be "
            "'image/jpeg', 'image/png', 'image/gif' or 'image/webp'",
            TEXTO_FORMATO_IMAGEN,
        ),
        ("Image does not match the provided media type image/jpeg", TEXTO_FORMATO_IMAGEN),
        (
            "messages.0.content.0.image.source.base64.data: The image was specified using "
            "the image/jpeg media type, but the image appears to be a image/png image",
            TEXTO_FORMATO_DISTINTO,
        ),
        ("image dimensions exceed max allowed size: 8000 pixels", TEXTO_FOTO_DIMENSIONES),
        (
            "messages.0.content.1.image.source.base64.data: image dimensions exceed max "
            "allowed size for many-image requests: 2000 pixels",
            TEXTO_FOTO_DIMENSIONES,
        ),
        (
            "You have reached your specified API usage limits. You will regain access "
            "on 2026-11-01 at 00:00 UTC.",
            TEXTO_LIMITE_GASTO,
        ),
        (
            "input length and `max_tokens` exceed context limit: 197000 + 8192 > 200000, "
            "decrease input length or `max_tokens` and try again",
            TEXTO_DOCUMENTO_LARGO,
        ),
        ("Could not process image", TEXTO_FOTO_ILEGIBLE),
        ("The image could not be processed", TEXTO_FOTO_ILEGIBLE),
        ("Too many images in the request (max 100)", TEXTO_DEMASIADAS_FOTOS),
        ("prompt is too long: 215000 tokens > 200000 maximum", TEXTO_DOCUMENTO_LARGO),
        ("Too many tokens in the request", TEXTO_DOCUMENTO_LARGO),
    ],
)
def test_reglas_por_texto_400(mensaje: str, esperado: str) -> None:
    e = _error(anthropic.BadRequestError, 400, mensaje)
    assert detalle_error_claude(e) == esperado


def test_limite_de_gasto_con_el_cliente_real_del_sdk() -> None:
    e = _error_del_sdk(
        400,
        _cuerpo(
            "You have reached your specified API usage limits. "
            "You will regain access on 2026-11-01 at 00:00 UTC."
        ),
    )
    assert detalle_error_claude(e) == TEXTO_LIMITE_GASTO


def test_textos_congelados() -> None:
    assert TEXTO_LIMITE_GASTO == (
        "Se alcanzó el límite de gasto de IA configurado en Anthropic: "
        "avisa a sistemas para que lo suban"
    )
    assert TEXTO_ARCHIVO_PESADO == (
        "La foto o el archivo pesa demasiado para la IA (máx. 5 MB por foto, "
        "32 MB por PDF): toma otra foto o recórtala"
    )
    assert TEXTO_FOTO_DIMENSIONES == (
        "La foto es demasiado grande en pixeles para la IA: tómala con menor "
        "resolución o recórtala"
    )
    assert TEXTO_FORMATO_DISTINTO == (
        "La foto llegó marcada con un formato distinto al real: reintenta y, "
        "si se repite, avisa a sistemas"
    )
    assert TEXTO_FORMATO_IMAGEN == (
        "Formato de archivo no soportado por la IA: usa una foto JPG o PNG, o un PDF"
    )
    assert TEXTO_FOTO_ILEGIBLE == "La IA no pudo leer la foto: toma otra con mejor luz y enfoque"
    assert TEXTO_DEMASIADAS_FOTOS == "Demasiadas fotos en una sola lectura: léelas en dos tandas"
    assert TEXTO_DOCUMENTO_LARGO == (
        "El documento es demasiado largo para la IA: pártelo en varios archivos"
    )
    assert TEXTO_LLAVE_INVALIDA == "La llave de la IA no es válida o venció: avisa a sistemas"
    assert TEXTO_MODELO_INEXISTENTE == "El modelo de IA configurado no existe: avisa a sistemas"
    assert TEXTO_LIMITE_PETICIONES == (
        "La IA está saturada (límite de peticiones): espera un minuto y reintenta"
    )
    assert TEXTO_IA_CAIDA == (
        "La IA está saturada o caída por el momento: reintenta en unos minutos"
    )


def test_contrato_sin_punto_final_ni_captura_a_mano() -> None:
    """Quien muestra el texto agrega «. Captura a mano…»: un punto final
    daba «..» y un «captura a mano» propio salía repetido."""
    textos = [v for k, v in vars(ia_errores).items() if k.startswith("TEXTO_")]
    assert len(textos) == 13
    for texto in textos:
        assert not texto.endswith("."), texto
        assert "captura" not in texto.lower(), texto


# --- Reglas por status ------------------------------------------------------


@pytest.mark.parametrize(
    ("clase", "status", "mensaje", "tipo", "esperado"),
    [
        (
            anthropic.AuthenticationError,
            401,
            "invalid x-api-key",
            "authentication_error",
            TEXTO_LLAVE_INVALIDA,
        ),
        (
            anthropic.PermissionDeniedError,
            403,
            "Your API key does not have permission to use the specified resource.",
            "permission_error",
            TEXTO_LLAVE_INVALIDA,
        ),
        (
            anthropic.NotFoundError,
            404,
            "model: claude-que-no-existe",
            "not_found_error",
            TEXTO_MODELO_INEXISTENTE,
        ),
        (
            anthropic.RequestTooLargeError,
            413,
            "Request exceeds the maximum allowed number of bytes.",
            "request_too_large",
            TEXTO_ARCHIVO_PESADO,
        ),
        (
            anthropic.RateLimitError,
            429,
            "Number of request tokens has exceeded your per-minute rate limit",
            "rate_limit_error",
            TEXTO_LIMITE_PETICIONES,
        ),
        (
            anthropic.InternalServerError,
            500,
            "Internal server error",
            "api_error",
            TEXTO_IA_CAIDA,
        ),
        (anthropic.InternalServerError, 503, "Service unavailable", "api_error", TEXTO_IA_CAIDA),
        (anthropic.OverloadedError, 529, "Overloaded", "overloaded_error", TEXTO_IA_CAIDA),
    ],
)
def test_reglas_por_status(clase, status: int, mensaje: str, tipo: str, esperado: str) -> None:
    e = _error(clase, status, mensaje, tipo)
    assert e.status_code == status
    assert detalle_error_claude(e) == esperado


def test_404_sin_modelo_cae_al_generico() -> None:
    e = _error(anthropic.NotFoundError, 404, "Not found.", "not_found_error")
    assert detalle_error_claude(e) == "Claude no disponible (404): Not found"


# --- Caso genérico ----------------------------------------------------------


def test_generico_recorta_a_120_y_sin_saltos_de_linea() -> None:
    mensaje = "temperature: valor fuera de rango\n" + ("x" * 200)
    e = _error(anthropic.BadRequestError, 400, mensaje)
    detalle = detalle_error_claude(e)
    assert detalle.startswith("Claude no disponible (400): temperature: valor fuera de rango x")
    assert "\n" not in detalle
    cola = detalle.removeprefix("Claude no disponible (400): ")
    assert len(cola) == 120


def test_generico_con_respuesta_no_json() -> None:
    e = _error_crudo(409, "conflicto\r\n  raro")
    assert detalle_error_claude(e) == "Claude no disponible (409): conflicto raro"


def test_generico_sin_mensaje() -> None:
    e = _error_crudo(409, "")
    assert detalle_error_claude(e) == "Claude no disponible (409)"


def test_generico_nunca_muestra_html() -> None:
    e = _error_crudo(409, "<html><body>Conflict</body></html>")
    assert detalle_error_claude(e) == "Claude no disponible (409)"


def test_generico_quita_el_punto_final_del_mensaje() -> None:
    e = _error(anthropic.BadRequestError, 400, "temperature: valor fuera de rango.")
    assert detalle_error_claude(e) == (
        "Claude no disponible (400): temperature: valor fuera de rango"
    )


def test_mensaje_y_tipo_salen_del_cuerpo() -> None:
    e = _error(anthropic.BadRequestError, 400, MENSAJE_SIN_SALDO)
    assert mensaje_error_claude(e) == MENSAJE_SIN_SALDO
    assert tipo_error_claude(e) == "invalid_request_error"
    sin_cuerpo = _error(anthropic.BadRequestError, 400, "x", con_cuerpo=False)
    assert tipo_error_claude(sin_cuerpo) == "BadRequestError"


# --- Constancia fiscal: siempre dice que se captura a mano -------------------


def test_captura_manual_sin_saldo() -> None:
    e = _error(anthropic.BadRequestError, 400, MENSAJE_SIN_SALDO)
    assert detalle_error_claude_captura_manual(e) == (
        f"{TEXTO_SIN_SALDO}; captura los datos manualmente"
    )


def test_captura_manual_se_agrega_si_falta() -> None:
    e = _error(anthropic.RateLimitError, 429, "rate limited", "rate_limit_error")
    assert detalle_error_claude_captura_manual(e) == (
        "La IA está saturada (límite de peticiones): espera un minuto y reintenta; "
        "captura los datos manualmente"
    )


def test_captura_manual_no_se_repite_si_ya_la_trae() -> None:
    e = _error_crudo(409, "Captura bloqueada.")
    assert detalle_error_claude_captura_manual(e) == (
        "Claude no disponible (409): Captura bloqueada"
    )


# --- Log: status, tipo y mensaje; nunca la llave -----------------------------


def test_registrar_error_claude_sin_llaves(caplog) -> None:
    e = _error(anthropic.BadRequestError, 400, MENSAJE_SIN_SALDO)
    logger = logging.getLogger("prueba_ia_errores")
    with caplog.at_level(logging.WARNING, logger="prueba_ia_errores"):
        registrar_error_claude(logger, e)
    assert len(caplog.records) == 1
    texto = caplog.records[0].getMessage()
    assert "400" in texto
    assert "invalid_request_error" in texto
    assert "credit balance is too low" in texto
    assert "req_prueba_123" in texto
    assert LLAVE_FALSA not in caplog.text


# --- Routers ----------------------------------------------------------------

client = TestClient(app)


def _con_token(monkeypatch) -> dict:
    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", TOKEN)
    get_settings.cache_clear()
    return {"X-Internal-Token": TOKEN}


def test_router_vision_gasto_502_con_detalle_de_saldo(monkeypatch) -> None:
    headers = _con_token(monkeypatch)

    def _sin_saldo(_req):
        raise _error(anthropic.BadRequestError, 400, MENSAJE_SIN_SALDO)

    monkeypatch.setattr(vision_router, "leer_ticket_gasto", _sin_saldo)
    res = client.post(
        "/vision/gasto",
        json={"image_base64": "aGVsbG8=", "media_type": "image/jpeg"},
        headers=headers,
    )
    assert res.status_code == 502
    assert res.json() == {"detail": TEXTO_SIN_SALDO}


def test_router_constancia_degrada_con_motivo_de_saldo(monkeypatch) -> None:
    headers = _con_token(monkeypatch)

    def _sin_saldo(_req):
        raise _error(anthropic.BadRequestError, 400, MENSAJE_SIN_SALDO)

    monkeypatch.setattr(vision_router, "leer_constancia_fiscal", _sin_saldo)
    res = client.post(
        "/vision/constancia-fiscal",
        json={"pdf_base64": "JVBERi0xLjQ="},
        headers=headers,
    )
    assert res.status_code == 200
    cuerpo = res.json()
    assert cuerpo["disponible"] is False
    assert cuerpo["legible"] is False
    assert cuerpo["motivo"] == f"{TEXTO_SIN_SALDO}; captura los datos manualmente"


def test_todos_los_routers_usan_la_fuente_unica() -> None:
    """Cada `except anthropic.APIStatusError` de app/routers traduce con
    `ia_errores`: nadie vuelve a pintar «Claude no disponible (400)» a mano."""
    carpeta = pathlib.Path(ia_errores.__file__).resolve().parents[1] / "routers"
    manejadores = 0
    for archivo in sorted(carpeta.glob("*.py")):
        fuente = archivo.read_text(encoding="utf-8")
        assert "Claude no disponible ({e.status_code})" not in fuente, archivo.name
        for nodo in ast.walk(ast.parse(fuente)):
            if not isinstance(nodo, ast.ExceptHandler) or nodo.type is None:
                continue
            if ast.unparse(nodo.type) != "anthropic.APIStatusError":
                continue
            manejadores += 1
            cuerpo = "\n".join(ast.unparse(s) for s in nodo.body)
            assert "registrar_error_claude(logger, e)" in cuerpo, archivo.name
            assert "detalle_error_claude" in cuerpo, archivo.name
    # vision ×5 (con constancia), gastos, compras, vencimientos, conciliacion ×3.
    assert manejadores == 11


# --- media_type real por bytes mágicos ---------------------------------------
#
# Una captura PNG etiquetada image/jpeg daba 400 «…but the image appears to be
# a image/png image». El tipo se deduce de los bytes antes de llamar a Claude.

_JPEG = b"\xff\xd8\xff\xe0\x00\x10JFIF\x00" + b"\x00" * 20
_PNG = b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR" + b"\x00" * 20
_GIF = b"GIF89a" + b"\x00" * 30
_WEBP = b"RIFF\x24\x00\x00\x00WEBPVP8 " + b"\x00" * 20


def _b64(datos: bytes) -> str:
    return base64.b64encode(datos).decode()


@pytest.mark.parametrize(
    ("datos", "declarado", "esperado"),
    [
        (_PNG, "image/jpeg", "image/png"),
        (_JPEG, "image/png", "image/jpeg"),
        (_GIF, "image/jpeg", "image/gif"),
        (_WEBP, "image/jpeg", "image/webp"),
        (_JPEG, "image/jpeg", "image/jpeg"),
        (b"hello world, no soy imagen", "image/png", "image/png"),
    ],
)
def test_media_type_real(datos: bytes, declarado: str, esperado: str) -> None:
    assert media_type_real(_b64(datos), declarado) == esperado


def test_media_type_real_tolerante() -> None:
    assert media_type_real(None, "image/jpeg") == "image/jpeg"
    assert media_type_real("", "image/png") == "image/png"
    assert media_type_real("@@@@ no es base64 @@@@", "image/png") == "image/png"
    assert media_type_real(f"data:image/jpeg;base64,{_b64(_PNG)}", "image/jpeg") == "image/png"
    # base64 partido en renglones (76 columnas, como lo arma más de un cliente)
    partido = "\n".join(_b64(_PNG * 3)[i : i + 4] for i in range(0, 40, 4))
    assert media_type_real(partido, "image/jpeg") == "image/png"


_RESPUESTA_TICKET = {"monto": 100.0, "moneda": "MXN", "confianza": 0.9, "legible": True}


def test_bloque_de_gasto_usa_el_tipo_real(claude_fake) -> None:
    cliente = claude_fake(anthropic_vision, _RESPUESTA_TICKET)
    anthropic_vision.leer_ticket_gasto(
        GastoTicketRequest(image_base64=_b64(_PNG), media_type="image/jpeg")
    )
    bloque = cliente.ultima["messages"][0]["content"][0]
    assert bloque["type"] == "image"
    assert bloque["source"]["media_type"] == "image/png"


def test_bloque_de_constancia_usa_el_tipo_real(claude_fake) -> None:
    cliente = claude_fake(anthropic_vision, {"legible": False, "confianza": 0.1})
    anthropic_vision.leer_constancia_fiscal(
        ConstanciaFiscalRequest(image_base64=_b64(_JPEG), media_type="image/png")
    )
    bloque = cliente.ultima["messages"][0]["content"][0]
    assert bloque["source"]["media_type"] == "image/jpeg"


def test_bloque_de_vencimiento_usa_el_tipo_real() -> None:
    req = VencimientoExtraerRequest(image_base64=_b64(_PNG), media_type="image/jpeg")
    assert vencimiento_extract._source_block(req)["source"]["media_type"] == "image/png"
