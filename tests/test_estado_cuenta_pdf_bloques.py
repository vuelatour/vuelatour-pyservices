"""Estado de cuenta PDF por BLOQUES de páginas (29-sep-2026).

El Scotiabank mensual (1.18 MB) truncaba la respuesta de UNA llamada a 16k
tokens (422 «demasiados movimientos») y tardaba más que los topes de la
cadena. Ahora el PDF se parte en bloques de N páginas (default 3) que se leen
en paralelo y se fusionan en orden. Estas pruebas congelan:

- 7 páginas ⇒ 3 llamadas (1–3, 4–6, 7), fusión en ORDEN de páginas aunque
  los bloques terminen desordenados, `uso_ia` sumado y el contexto del
  encabezado SOLO en los bloques que no traen la página 1;
- ≤ N páginas ⇒ UNA llamada byte-idéntica a la de siempre;
- bloque truncado ⇒ se re-parte; una página sola truncada ⇒ el error de hoy;
- pypdf no puede abrir/partir el PDF ⇒ el camino de siempre;
- la cadena de saldos cruza bloques.

Ninguna prueba llama a Claude: el cliente falso lee con pypdf qué páginas
trae cada documento (cada página imprime «PAGINA n») y contesta por página.
"""

import base64
import json
import logging
import re
import threading
import time
from io import BytesIO
from types import SimpleNamespace

import pytest
from pypdf import PdfReader, PdfWriter
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

from app.config import Settings, get_settings
from app.schemas.conciliacion import ConciliacionParseRequest
from app.services import estado_cuenta
from app.services._dominio import sistema_con_dominio
from app.services.estado_cuenta import parsear_estado_cuenta

MSG_TRUNCADO = (
    "El PDF tiene demasiados movimientos para leerse completo con IA. "
    "Exporta el estado de cuenta en CSV o Excel desde el portal del banco e "
    "impórtalo: es el formato exacto y preferido."
)


# --- PDF sintético ----------------------------------------------------------


def _pdf(paginas: int) -> bytes:
    """PDF de `paginas` páginas; la 1 trae el encabezado del banco."""
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=letter)
    for n in range(1, paginas + 1):
        if n == 1:
            c.drawString(72, 720, "ESTADO DE CUENTA SCOTIABANK")
            c.drawString(72, 705, "PERIODO DEL 01/09/2026 AL 30/09/2026")
            c.drawString(72, 690, "CUENTA 0025830577 MONEDA MXN")
        c.drawString(72, 650, f"PAGINA {n}")
        c.showPage()
    c.save()
    return buf.getvalue()


def _b64(raw: bytes) -> str:
    return base64.b64encode(raw).decode("ascii")


def _req(raw: bytes) -> ConciliacionParseRequest:
    return ConciliacionParseRequest(
        filename="scotiabank.pdf", file_base64=_b64(raw), cuenta_moneda="MXN"
    )


# --- Cliente falso que contesta por página ---------------------------------


class _Uso:
    """Consumo proporcional a las páginas: la suma se puede comprobar."""

    def __init__(self, paginas: int) -> None:
        self.input_tokens = 100 * paginas
        self.output_tokens = 10 * paginas
        self.cache_creation_input_tokens = 1
        self.cache_read_input_tokens = 2


class _Respuesta:
    def __init__(self, payload, stop_reason: str, paginas: list[int]) -> None:
        texto = payload if isinstance(payload, str) else json.dumps(payload, ensure_ascii=False)
        self.content = [SimpleNamespace(type="text", text=texto)]
        self.stop_reason = stop_reason
        self.model = "claude-opus-4-8"
        self.usage = _Uso(len(paginas))


def _paginas_del_documento(kwargs: dict) -> list[int]:
    doc = kwargs["messages"][0]["content"][0]
    raw = base64.b64decode(doc["source"]["data"])
    lector = PdfReader(BytesIO(raw))
    return [
        int(re.search(r"PAGINA (\d+)", p.extract_text() or "").group(1)) for p in lector.pages
    ]


class ClientePorPaginas:
    """`messages.create` thread-safe que registra (páginas, kwargs) por llamada."""

    def __init__(self, responder) -> None:
        self.responder = responder
        self.llamadas: list[tuple[list[int], dict]] = []
        self.opciones: list[dict] = []
        self._lock = threading.Lock()

    @property
    def messages(self) -> "ClientePorPaginas":
        return self

    def with_options(self, **kwargs) -> "ClientePorPaginas":
        with self._lock:
            self.opciones.append(kwargs)
        return self

    def create(self, **kwargs):
        paginas = _paginas_del_documento(kwargs)
        with self._lock:
            self.llamadas.append((paginas, kwargs))
        payload, stop = self.responder(paginas)
        return _Respuesta(payload, stop, paginas)

    # --- helpers de aserción -------------------------------------------
    def paginas(self) -> list[list[int]]:
        return sorted(p for p, _ in self.llamadas)

    def kwargs_de(self, paginas: list[int]) -> dict:
        return next(k for p, k in self.llamadas if p == paginas)

    def texto_de(self, paginas: list[int]) -> str:
        contenido = self.kwargs_de(paginas)["messages"][0]["content"]
        return "\n".join(b["text"] for b in contenido if b["type"] == "text")


@pytest.fixture
def cliente_por_paginas(monkeypatch):
    def instalar(responder) -> ClientePorPaginas:
        cliente = ClientePorPaginas(responder)
        monkeypatch.setattr(estado_cuenta, "_client", lambda: cliente)
        return cliente

    return instalar


def _mov(n: int) -> dict:
    """Un cargo de 100 por página con saldo corrido CONSISTENTE."""
    return {
        "fecha": f"2026-09-{n:02d}",
        "descripcion": f"CARGO PAGINA {n}",
        "monto": 100,
        "tipo": "CARGO",
        "referencia": f"00258305{n:02d}",
        "saldo_posterior": 10000 - 100 * n,
    }


def _responder(total: int, *, truncar=lambda paginas: False, omitir=(), sin_totales=False):
    """Contesta como lo haría el modelo con la nota de páginas del bloque."""

    def responder(paginas: list[int]):
        if truncar(paginas):
            return {"movimientos": [_mov(n) for n in paginas[:1]]}, "max_tokens"
        con_p1 = 1 in paginas
        payload = {
            "movimientos": [_mov(n) for n in paginas if n not in omitir],
            "periodo_inicio": "2026-09-01" if con_p1 else None,
            "periodo_fin": "2026-09-30" if con_p1 else None,
            "saldo_inicial": 10000 if con_p1 else None,
            "saldo_final": 10000 - 100 * total if total in paginas else None,
            "total_cargos": (100 * total if con_p1 else None) if not sin_totales else None,
            "total_abonos": None,
            "advertencias": [f"aviso págs {paginas[0]}-{paginas[-1]}"],
        }
        if con_p1:
            # El bloque 1 termina AL ÚLTIMO: el orden del resultado no puede
            # depender de quién acaba primero.
            time.sleep(0.05)
        return payload, "end_turn"

    return responder


# --- (a) 7 páginas ⇒ 3 bloques en paralelo ---------------------------------


def test_siete_paginas_se_leen_en_tres_bloques_y_se_fusionan_en_orden(
    cliente_por_paginas,
) -> None:
    responder_base = _responder(7)

    def responder(paginas):
        payload, stop = responder_base(paginas)
        if 4 in paginas:
            # Un bloque posterior que «inventa» un total: gana el del bloque 1
            # (primer valor no nulo EN ORDEN de bloque, no de llegada).
            payload["total_cargos"] = 999
        return payload, stop

    cliente = cliente_por_paginas(responder)
    req = _req(_pdf(7))
    res = parsear_estado_cuenta(req)

    assert cliente.paginas() == [[1, 2, 3], [4, 5, 6], [7]]
    assert res.formato == "pdf"
    assert res.total == 7
    assert [m.descripcion for m in res.movimientos] == [f"CARGO PAGINA {n}" for n in range(1, 8)]
    # Totales del bloque 1 (700) cuadran con lo transcrito; cadena de saldos
    # consistente a través de bloques; solo quedan los avisos de la IA, en orden.
    assert res.advertencias == ["aviso págs 1-3", "aviso págs 4-6", "aviso págs 7-7"]
    assert "leído en 3 bloques" in res.notas
    assert "7 páginas" in res.notas

    # uso_ia = SUMA de las tres llamadas.
    assert res.uso_ia is not None
    assert res.uso_ia.modelo == "claude-opus-4-8"
    assert res.uso_ia.input_tokens == 700
    assert res.uso_ia.output_tokens == 70
    assert res.uso_ia.cache_creation_input_tokens == 3
    assert res.uso_ia.cache_read_input_tokens == 6
    assert res.modelo == get_settings().anthropic_model

    # Mismos parámetros de siempre en CADA bloque.
    for _, kwargs in cliente.llamadas:
        assert kwargs["max_tokens"] == 16384
        assert kwargs["system"] == sistema_con_dominio(estado_cuenta._SYSTEM)
        assert kwargs["model"] == get_settings().anthropic_model
    assert cliente.opciones == [{"timeout": 240.0}] * 3

    # Bloque 1: prompt de siempre + nota de páginas, SIN contexto del encabezado.
    b1 = cliente.texto_de([1, 2, 3])
    assert len(cliente.kwargs_de([1, 2, 3])["messages"][0]["content"]) == 2
    assert b1.startswith(estado_cuenta._prompt_pdf(req))
    assert "las páginas 1–3 de 7" in b1
    assert "CONTEXTO DEL ENCABEZADO" not in b1
    assert "PERIODO DEL" not in b1

    # Bloques 2 y 3: sus páginas + el encabezado de la página 1 como texto.
    b2 = cliente.texto_de([4, 5, 6])
    assert "CONTEXTO DEL ENCABEZADO" in b2
    assert "NO transcribas movimientos de este texto" in b2
    assert "PERIODO DEL 01/09/2026 AL 30/09/2026" in b2
    assert "CUENTA 0025830577 MONEDA MXN" in b2
    assert "las páginas 4–6 de 7" in b2
    assert "El documento adjunto trae SOLO esas páginas" in b2
    assert "Transcribe únicamente los renglones del documento PDF adjunto" in b2
    assert "transcribe SOLO los movimientos impresos en estas páginas" in b2
    assert "déjalas en null si no están impresas en estas páginas" in b2
    b3 = cliente.texto_de([7])
    assert "CONTEXTO DEL ENCABEZADO" in b3
    assert "PERIODO DEL 01/09/2026 AL 30/09/2026" in b3
    assert "la página 7 de 7" in b3


def test_paginas_por_bloque_sale_del_setting(cliente_por_paginas, monkeypatch) -> None:
    monkeypatch.setenv("ESTADO_CUENTA_PAGINAS_POR_BLOQUE", "4")
    assert Settings().estado_cuenta_paginas_por_bloque == 4
    monkeypatch.setattr(
        estado_cuenta,
        "get_settings",
        lambda: get_settings().model_copy(update={"estado_cuenta_paginas_por_bloque": 4}),
    )
    cliente = cliente_por_paginas(_responder(7))
    res = parsear_estado_cuenta(_req(_pdf(7)))
    assert cliente.paginas() == [[1, 2, 3, 4], [5, 6, 7]]
    assert res.total == 7


# --- (b) ≤ N páginas ⇒ UNA llamada byte-idéntica ---------------------------


def _llamada_de_siempre(req: ConciliacionParseRequest) -> dict:
    return {
        "model": get_settings().anthropic_model,
        "max_tokens": 16384,
        "system": sistema_con_dominio(estado_cuenta._SYSTEM),
        "messages": [
            {
                "role": "user",
                "content": [
                    {
                        "type": "document",
                        "source": {
                            "type": "base64",
                            "media_type": "application/pdf",
                            "data": req.file_base64,
                        },
                    },
                    {"type": "text", "text": estado_cuenta._prompt_pdf(req)},
                ],
            }
        ],
    }


def test_pdf_de_dos_paginas_va_completo_en_una_llamada_identica(
    cliente_por_paginas, monkeypatch
) -> None:
    cliente = cliente_por_paginas(_responder(2))
    req = _req(_pdf(2))
    res = parsear_estado_cuenta(req)

    assert len(cliente.llamadas) == 1
    paginas, kwargs = cliente.llamadas[0]
    assert paginas == [1, 2]
    # El documento ORIGINAL, el prompt de siempre y nada más.
    assert kwargs == _llamada_de_siempre(req)
    assert cliente.opciones == [{"timeout": 240.0}]
    assert res.total == 2
    assert res.uso_ia is not None and res.uso_ia.input_tokens == 200
    assert "bloques" not in res.notas

    # Y es la MISMA llamada que hace el camino de siempre (sin pypdf).
    monkeypatch.setattr(estado_cuenta, "PdfReader", None)
    cliente_sin_pypdf = cliente_por_paginas(_responder(2))
    parsear_estado_cuenta(req)
    assert cliente_sin_pypdf.llamadas[0][1] == kwargs


def test_pdf_de_tres_paginas_sigue_siendo_una_llamada(cliente_por_paginas) -> None:
    cliente = cliente_por_paginas(_responder(3))
    req = _req(_pdf(3))
    res = parsear_estado_cuenta(req)
    assert len(cliente.llamadas) == 1
    assert cliente.llamadas[0][1] == _llamada_de_siempre(req)
    assert res.total == 3


# --- (c) max_tokens: re-partir; una página sola ⇒ error de siempre ---------


def test_bloque_truncado_se_reparte_a_la_mitad(cliente_por_paginas) -> None:
    # Cualquier bloque de más de una página que traiga la 4 no cabe.
    cliente = cliente_por_paginas(
        _responder(7, truncar=lambda paginas: 4 in paginas and len(paginas) > 1)
    )
    res = parsear_estado_cuenta(_req(_pdf(7)))

    # 4–6 ⇒ 4–5 + 6 ⇒ 4 + 5.
    assert cliente.paginas() == [[1, 2, 3], [4], [4, 5], [4, 5, 6], [5], [6], [7]]
    assert [m.descripcion for m in res.movimientos] == [f"CARGO PAGINA {n}" for n in range(1, 8)]
    assert "leído en 5 bloques" in res.notas
    # Las llamadas truncadas también se cobraron: entran a la suma.
    assert res.uso_ia is not None
    assert res.uso_ia.input_tokens == 100 * (3 + 3 + 2 + 1 + 1 + 1 + 1)
    # El pedazo re-partido lleva su nota y el encabezado.
    assert "la página 4 de 7" in cliente.texto_de([4])
    assert "CONTEXTO DEL ENCABEZADO" in cliente.texto_de([4])


def test_pdf_chico_truncado_tambien_se_reparte(cliente_por_paginas) -> None:
    cliente = cliente_por_paginas(_responder(2, truncar=lambda paginas: len(paginas) > 1))
    req = _req(_pdf(2))
    res = parsear_estado_cuenta(req)
    # La primera llamada fue la de siempre; al truncarse, página por página.
    assert cliente.kwargs_de([1, 2]) == _llamada_de_siempre(req)
    assert cliente.paginas() == [[1], [1, 2], [2]]
    assert res.total == 2
    assert "CONTEXTO DEL ENCABEZADO" not in cliente.texto_de([1])
    assert "CONTEXTO DEL ENCABEZADO" in cliente.texto_de([2])


def test_una_pagina_sola_truncada_es_el_error_de_siempre(cliente_por_paginas) -> None:
    cliente_por_paginas(_responder(7, truncar=lambda paginas: 5 in paginas))
    with pytest.raises(ValueError) as exc:
        parsear_estado_cuenta(_req(_pdf(7)))
    assert str(exc.value) == MSG_TRUNCADO


def test_un_bloque_que_revienta_tumba_toda_la_lectura(cliente_por_paginas) -> None:
    base = _responder(7)

    def responder(paginas):
        if 7 in paginas:
            raise RuntimeError("Claude no disponible")
        return base(paginas)

    cliente_por_paginas(responder)
    with pytest.raises(RuntimeError, match="Claude no disponible"):
        parsear_estado_cuenta(_req(_pdf(7)))


def test_un_bloque_ilegible_tumba_toda_la_lectura(cliente_por_paginas) -> None:
    base = _responder(7)

    def responder(paginas):
        if 4 in paginas:
            return "no es JSON", "end_turn"
        return base(paginas)

    cliente_por_paginas(responder)
    with pytest.raises(ValueError, match="no devolvió una respuesta interpretable"):
        parsear_estado_cuenta(_req(_pdf(7)))


def test_tope_total_de_tiempo(cliente_por_paginas, monkeypatch, caplog) -> None:
    monkeypatch.setattr(estado_cuenta, "_TOPE_TOTAL_PDF_S", 0.3)
    base = _responder(7)

    def responder(paginas):
        if 7 in paginas:
            time.sleep(2.0)
        return base(paginas)

    cliente_por_paginas(responder)
    req = _req(_pdf(7))
    inicio = time.monotonic()
    with (
        caplog.at_level(logging.WARNING, logger="estado_cuenta"),
        pytest.raises(ValueError, match="tardó más de 4 minutos"),
    ):
        parsear_estado_cuenta(req)
    # No espera al bloque que sigue en vuelo (dormiría 2 s).
    assert time.monotonic() - inicio < 1.5
    # Lo ya cobrado de los bloques terminados queda en el log: el API no lo
    # registra en ia_uso sin una respuesta 200.
    assert any("sin registrar en ia_uso" in r.getMessage() for r in caplog.records)


# --- (d) pypdf no puede ⇒ camino de siempre --------------------------------


def test_pdf_corrupto_se_lee_como_siempre(cliente_por_paginas, monkeypatch, caplog) -> None:
    raw = _pdf(7)

    def lector_que_falla(*_args, **_kwargs):
        raise ValueError("xref roto")

    monkeypatch.setattr(estado_cuenta, "PdfReader", lector_que_falla)
    cliente = cliente_por_paginas(_responder(7))
    req = _req(raw)
    with caplog.at_level(logging.WARNING, logger="estado_cuenta"):
        res = parsear_estado_cuenta(req)
    assert len(cliente.llamadas) == 1
    assert cliente.llamadas[0][1] == _llamada_de_siempre(req)
    assert res.total == 7
    assert "bloques" not in res.notas
    assert any("pypdf no pudo abrir" in r.getMessage() for r in caplog.records)


def test_pdf_cifrado_se_lee_como_siempre(claude_fake, caplog) -> None:
    escritor = PdfWriter()
    for pagina in PdfReader(BytesIO(_pdf(7))).pages:
        escritor.add_page(pagina)
    escritor.encrypt(user_password="secreta", owner_password="duenio", algorithm="RC4-128")
    buf = BytesIO()
    escritor.write(buf)
    req = _req(buf.getvalue())

    cliente = claude_fake(estado_cuenta, {"movimientos": [_mov(1)], "advertencias": []})
    with caplog.at_level(logging.WARNING, logger="estado_cuenta"):
        res = parsear_estado_cuenta(req)
    assert len(cliente.llamadas) == 1
    assert cliente.ultima == _llamada_de_siempre(req)
    assert res.total == 1
    assert any("pypdf no pudo abrir" in r.getMessage() for r in caplog.records)


def test_pdf_que_no_se_puede_partir_se_lee_como_siempre(
    cliente_por_paginas, monkeypatch, caplog
) -> None:
    def escritor_que_falla(*_args, **_kwargs):
        raise RuntimeError("fuente rota")

    monkeypatch.setattr(estado_cuenta, "_bloque_b64", escritor_que_falla)
    cliente = cliente_por_paginas(_responder(7))
    req = _req(_pdf(7))
    with caplog.at_level(logging.WARNING, logger="estado_cuenta"):
        res = parsear_estado_cuenta(req)
    assert len(cliente.llamadas) == 1
    assert cliente.llamadas[0][1] == _llamada_de_siempre(req)
    assert res.total == 7
    assert any("pypdf no pudo partir" in r.getMessage() for r in caplog.records)


# --- (e) la cadena de saldos cruza bloques ---------------------------------


def test_cadena_de_saldos_rota_entre_bloques_avisa(cliente_por_paginas) -> None:
    # El bloque 2 se «salta» el movimiento de la página 4: el saldo de la 3
    # (bloque 1) ya no enlaza con el de la 5 (bloque 2).
    cliente_por_paginas(_responder(7, omitir=(4,), sin_totales=True))
    res = parsear_estado_cuenta(_req(_pdf(7)))
    assert res.total == 6
    rotas = [a for a in res.advertencias if "no cuadra" in a]
    assert len(rotas) == 1
    assert "faltar movimientos" in rotas[0]
    assert "2026-09-03 → 2026-09-05" in rotas[0]
    assert "revisa las advertencias" in res.notas


def test_totales_del_bloque_uno_contra_todo_lo_transcrito(cliente_por_paginas) -> None:
    # El total impreso (700, página 1) se compara contra los 7 bloques: si un
    # bloque pierde un movimiento, el aviso sale aunque ese bloque no traiga
    # el resumen.
    cliente_por_paginas(_responder(7, omitir=(7,)))
    res = parsear_estado_cuenta(_req(_pdf(7)))
    assert any("suman 600.00" in a and "700.00" in a for a in res.advertencias)


# --- Revisión adversaria (29-sep-2026) --------------------------------------


def test_pdf_cifrado_solo_con_contrasena_de_dueno_se_parte(cliente_por_paginas) -> None:
    # Los bancos suelen cifrar con contraseña de DUEÑO (sin contraseña para
    # abrir): pypdf lo abre con "" y los bloques salen legibles.
    escritor = PdfWriter()
    for pagina in PdfReader(BytesIO(_pdf(7))).pages:
        escritor.add_page(pagina)
    escritor.encrypt(user_password="", owner_password="duenio", algorithm="AES-256")
    buf = BytesIO()
    escritor.write(buf)

    cliente = cliente_por_paginas(_responder(7))
    res = parsear_estado_cuenta(_req(buf.getvalue()))
    assert cliente.paginas() == [[1, 2, 3], [4, 5, 6], [7]]
    assert [m.descripcion for m in res.movimientos] == [f"CARGO PAGINA {n}" for n in range(1, 8)]
    assert "PERIODO DEL 01/09/2026 AL 30/09/2026" in cliente.texto_de([4, 5, 6])


def test_encabezado_se_extrae_solo_si_hace_falta_y_una_vez(
    cliente_por_paginas, monkeypatch
) -> None:
    extracciones: list[int] = []
    original = estado_cuenta._texto_encabezado

    def contar(lector):
        extracciones.append(1)
        return original(lector)

    monkeypatch.setattr(estado_cuenta, "_texto_encabezado", contar)

    # ≤ N páginas sin truncar: ni se calcula (el camino de siempre no cambia).
    cliente_por_paginas(_responder(3))
    parsear_estado_cuenta(_req(_pdf(3)))
    assert extracciones == []

    # 7 páginas: una sola extracción para los bloques 2 y 3.
    cliente_por_paginas(_responder(7))
    parsear_estado_cuenta(_req(_pdf(7)))
    assert extracciones == [1]


def test_sin_texto_en_la_pagina_uno_los_bloques_van_sin_contexto(
    cliente_por_paginas, monkeypatch
) -> None:
    # PDF escaneado / fuente sin texto: el bloque va sin el contexto (no con
    # un bloque de texto vacío) y la lectura sigue.
    monkeypatch.setattr(estado_cuenta, "_texto_encabezado", lambda _lector: "")
    cliente = cliente_por_paginas(_responder(7))
    res = parsear_estado_cuenta(_req(_pdf(7)))
    assert res.total == 7
    assert len(cliente.kwargs_de([4, 5, 6])["messages"][0]["content"]) == 2
    assert "CONTEXTO DEL ENCABEZADO" not in cliente.texto_de([4, 5, 6])
    assert "las páginas 4–6 de 7" in cliente.texto_de([4, 5, 6])


def test_pdf_de_una_pagina_truncado_es_el_error_de_siempre(cliente_por_paginas) -> None:
    cliente = cliente_por_paginas(_responder(1, truncar=lambda paginas: True))
    req = _req(_pdf(1))
    with pytest.raises(ValueError) as exc:
        parsear_estado_cuenta(req)
    assert str(exc.value) == MSG_TRUNCADO
    # Una sola llamada, la de siempre: no hay nada que re-partir.
    assert len(cliente.llamadas) == 1
    assert cliente.llamadas[0][1] == _llamada_de_siempre(req)


def _lector_con_texto(texto: str) -> SimpleNamespace:
    return SimpleNamespace(pages=[SimpleNamespace(extract_text=lambda: texto)])


def test_encabezado_rescata_el_periodo_despues_del_corte() -> None:
    # PDF impreso del portal: el menú del sitio llena los primeros 1,500
    # caracteres y el periodo (el AÑO de «07 SEP») queda fuera del corte.
    menu = "\n".join(f"Menu del portal opcion {i:03d} Inicio Cuentas Pagos" for i in range(60))
    texto = f"{menu}\nCUENTA 0025830577\nPERIODO DEL 01/09/2026 AL 30/09/2026\n07 SEP ASUR 1,234.56"
    encabezado = estado_cuenta._texto_encabezado(_lector_con_texto(texto))
    assert encabezado.startswith(menu[:1500])
    assert encabezado.endswith("PERIODO DEL 01/09/2026 AL 30/09/2026")
    # Solo los renglones del periodo: ningún movimiento viaja en el contexto.
    assert "ASUR" not in encabezado

    # Con el periodo dentro del corte, el encabezado es el de siempre.
    corto = "ESTADO DE CUENTA\nPERIODO DEL 01/09/2026 AL 30/09/2026\n" + menu
    assert estado_cuenta._texto_encabezado(_lector_con_texto(corto)) == corto[:1500]


def test_encabezado_ilegible_no_rompe_la_lectura() -> None:
    def revienta():
        raise ValueError("fuente rota")

    lector = SimpleNamespace(pages=[SimpleNamespace(extract_text=revienta)])
    assert estado_cuenta._texto_encabezado(lector) == ""


# --- Modelo de la petición (X-IA-Modelo, 2-oct-2026) -----------------------


def test_bloques_en_hilos_usan_el_modelo_de_la_peticion(
    cliente_por_paginas, monkeypatch, request
) -> None:
    """Los bloques corren en hilos de un ThreadPoolExecutor, que NO heredan el
    ContextVar del modelo: `_leer_por_bloques` lo resuelve antes y lo pasa.
    Por HTTP real: dependencia async → endpoint en su hilo → pool de bloques."""
    from fastapi.testclient import TestClient

    from app.main import app

    monkeypatch.setenv("INTERNAL_SHARED_TOKEN", "secreto-de-prueba")
    monkeypatch.setenv("ANTHROPIC_MODEL", "claude-servidor-prueba")
    get_settings.cache_clear()
    request.addfinalizer(get_settings.cache_clear)
    cliente = cliente_por_paginas(_responder(7))

    res = TestClient(app).post(
        "/conciliacion/parse",
        json=_req(_pdf(7)).model_dump(),
        headers={"X-Internal-Token": "secreto-de-prueba", "X-IA-Modelo": "claude-prueba-pdf"},
    )

    assert res.status_code == 200, res.text
    assert cliente.paginas() == [[1, 2, 3], [4, 5, 6], [7]]
    assert [k["model"] for _, k in cliente.llamadas] == ["claude-prueba-pdf"] * 3
    assert res.json()["modelo"] == "claude-prueba-pdf"
