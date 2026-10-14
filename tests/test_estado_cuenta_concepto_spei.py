"""Concepto del SPEI en el PDF del estado de cuenta (13-oct-2026, API 0.0.75).

Caso real (EJE SCOTIA, septiembre 2026): 73 de 98 filas quedaron como
«SEL TRANSF. INTERBANCARIA SPEI» / «TRANSF INTERBANCARIA SPEI» SIN el concepto
que el banco imprime debajo («BANORTE» / «fact vip 7360»): la oficina no
podía saber qué factura pagó cada transferencia. Ahora:

- el prompt exige los RENGLONES DE CONTINUACIÓN (banco, beneficiario u
  ordenante y concepto) pegados al renglón principal, sin CLABE ni clave de
  rastreo, hasta 160 caracteres (y el código recorta a 200);
- la continuación huérfana al inicio de un bloque se ignora SOLO en los
  bloques que no empiezan en la página 1 (`_nota_paginas`);
- la descripción llega en UNA línea (`_descripcion_limpia`);
- una advertencia DETERMINISTA cuenta las transferencias/pagos que se
  leyeron con SOLO la leyenda del banco.

Ninguna prueba llama a Claude (`claude_fake`).
"""

import base64
from io import BytesIO

from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

from app.schemas.conciliacion import ConciliacionParseRequest, MovimientoParseado
from app.services import estado_cuenta
from app.services.estado_cuenta import (
    _LEYENDAS_SIN_CONCEPTO,
    _SYSTEM,
    _advertencia_sin_concepto,
    _advertencias_pdf,
    _descripcion_limpia,
    _movimientos_de,
    _nota_paginas,
    _plano,
    _prompt_pdf,
    parsear_estado_cuenta,
)

AVISO_SIN_CONCEPTO = (
    "transferencia(s) o pago(s) se leyeron SIN su concepto (solo la leyenda del "
    "banco, p. ej. «SEL TRANSF. INTERBANCARIA SPEI»): la IA no juntó los renglones "
    "de abajo (banco, beneficiario y concepto). Vuelve a leer el archivo antes de "
    "importar si necesitas saber qué factura pagó cada una."
)


def _req(raw: bytes | None = None) -> ConciliacionParseRequest:
    return ConciliacionParseRequest(
        filename="scotia eje sep.pdf",
        file_base64=base64.b64encode(raw or _pdf()).decode("ascii"),
        cuenta_moneda="MXN",
    )


def _pdf() -> bytes:
    """PDF de UNA página (≤ N páginas ⇒ una sola llamada, la de siempre)."""
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=letter)
    c.drawString(72, 720, "ESTADO DE CUENTA SCOTIABANK · PERIODO SEP 2026")
    c.drawString(72, 700, "24 SEP SEL TRANSF. INTERBANCARIA SPEI     1,118.12")
    c.drawString(90, 688, "BANORTE")
    c.drawString(90, 676, "fact vip 7360")
    c.showPage()
    c.save()
    return buf.getvalue()


def _mov(descripcion: str | None, monto: float = 100.0) -> MovimientoParseado:
    return MovimientoParseado(
        fecha="2026-09-24", descripcion=descripcion, monto=monto, tipo="CARGO"
    )


# --- Prompt -----------------------------------------------------------------


def test_system_exige_los_renglones_de_continuacion() -> None:
    assert "RENGLONES DE CONTINUACIÓN" in _SYSTEM
    assert "160" in _SYSTEM
    assert "clave de rastreo" in _SYSTEM
    assert "la CLABE" in _SYSTEM
    assert (
        "«SEL TRANSF. INTERBANCARIA SPEI» con «BANORTE» y «fact vip 7360» debajo ⇒ "
        '"SEL TRANSF. INTERBANCARIA SPEI BANORTE fact vip 7360"'
    ) in _SYSTEM
    assert "Hazlo en TODAS las transferencias del documento" in _SYSTEM
    # La literalidad (llave del dedupe) se conserva; el corte de 80 se fue.
    assert "LITERAL" in _SYSTEM
    assert "primeros 80" not in _SYSTEM
    # Scotiabank nombra sus leyendas con continuación.
    assert "POR COBRANZA DOMICILIADA, SCN Pagos a Proveedores) traen debajo" in _SYSTEM
    # La orden de ignorar la continuación huérfana NO va en el system: solo
    # en los bloques que no empiezan en la página 1.
    assert "NO los transcribas ni los pegues" not in _SYSTEM


def test_prompt_del_usuario_recuerda_la_continuacion() -> None:
    texto = _prompt_pdf(_req())
    assert "renglones de continuación" in texto
    assert "banco, beneficiario u ordenante y concepto" in texto
    assert texto.endswith("ninguna fila de más ni de menos.")


def test_continuacion_huerfana_solo_en_bloques_que_no_empiezan_en_la_pagina_1() -> None:
    con = _nota_paginas(3, 6, 7)
    assert "las páginas 4–6 de 7" in con
    assert con.endswith(
        "Si la PRIMERA página de este bloque empieza con renglones sin fecha ni "
        "monto (banco, beneficiario o concepto) que continúan un movimiento de la "
        "página anterior, NO los transcribas ni los pegues a otro movimiento: ese "
        "movimiento lo transcribe otro bloque."
    )
    sin = _nota_paginas(0, 3, 7)
    assert "NO los transcribas ni los pegues" not in sin
    assert sin.endswith("y NO va en advertencias.")
    assert "NO los transcribas ni los pegues" in _nota_paginas(6, 7, 7)


# --- Descripción en una línea -------------------------------------------------


def test_descripcion_limpia_una_linea_tope_y_vacios() -> None:
    assert (
        _descripcion_limpia("SEL TRANSF. INTERBANCARIA SPEI\nBANORTE\n  fact vip 7360")
        == "SEL TRANSF. INTERBANCARIA SPEI BANORTE fact vip 7360"
    )
    assert _descripcion_limpia("x" * 250) == "x" * 200
    assert _descripcion_limpia("   \n\t ") is None
    assert _descripcion_limpia("") is None
    assert _descripcion_limpia(None) is None
    # Corte en un espacio: no deja espacio colgando.
    assert _descripcion_limpia("a" * 199 + " " + "b" * 10) == "a" * 199


def test_movimientos_de_junta_la_continuacion() -> None:
    data = {
        "movimientos": [
            {
                "fecha": "2026-09-24",
                "descripcion": "SEL TRANSF. INTERBANCARIA SPEI\nBANORTE\n  fact vip 7360",
                "monto": 1118.12,
                "tipo": "CARGO",
                "referencia": "00000000000000000001",
            },
            {"fecha": "2026-09-24", "descripcion": "y" * 250, "monto": 5, "tipo": "CARGO"},
            {"fecha": "2026-09-24", "descripcion": "   ", "monto": 7, "tipo": "CARGO"},
        ]
    }
    movs = _movimientos_de(data)
    assert [m.descripcion for m in movs] == [
        "SEL TRANSF. INTERBANCARIA SPEI BANORTE fact vip 7360",
        "y" * 200,
        None,
    ]
    assert movs[0].referencia == "00000000000000000001"


# --- Advertencia determinista ------------------------------------------------


def test_leyendas_son_las_nueve_del_api_en_forma_plana() -> None:
    leyendas = set(_LEYENDAS_SIN_CONCEPTO)
    assert leyendas == {
        "SEL TRANSF INTERBANCARIA SPEI",
        "TRANSF INTERBANCARIA SPEI",
        "TRANSF ENVIO SPEI BCA INTERNET",
        "SEL PAGO DE SERVICIOS",
        "03 POR COBRANZA DOMICILIADA",
        "SCN PAGOS A PROVEEDORES",
        "CGO SPEI",
        "SPEI ENVIADO",
        "SPEI RECIBIDO",
    }
    assert all(_plano(leyenda) == leyenda for leyenda in _LEYENDAS_SIN_CONCEPTO)
    # Mismo criterio que normalizarPlano del API: acentos, puntos y espacios.
    assert _plano("  Sel Tránsf.  Interbancaria   SPEI ") == "SEL TRANSF INTERBANCARIA SPEI"
    assert _plano(None) == ""


def test_advertencia_con_la_leyenda_sola_y_no_con_concepto() -> None:
    movs = [
        _mov("SEL TRANSF. INTERBANCARIA SPEI"),
        _mov("TRANSF INTERBANCARIA SPEI"),
        _mov("SCN Pagos a Proveedores"),
        _mov("SEL TRANSF. INTERBANCARIA SPEI BANORTE fact vip 7360"),
        _mov("ASA CANCUN"),
        _mov(None),
    ]
    assert _advertencia_sin_concepto(movs) == [f"3 {AVISO_SIN_CONCEPTO}"]
    assert _advertencia_sin_concepto(movs[3:]) == []
    assert _advertencia_sin_concepto([]) == []


def test_advertencia_va_primero_y_respeta_el_tope_de_diez() -> None:
    movs = [
        MovimientoParseado(
            fecha="2026-09-24",
            descripcion="SEL TRANSF. INTERBANCARIA SPEI",
            monto=100,
            tipo="CARGO",
            saldo_posterior=1000,
        ),
        # Saldo que no cuadra ⇒ el aviso de la cadena va DESPUÉS.
        MovimientoParseado(
            fecha="2026-09-25", descripcion="ASA CANCUN", monto=50, tipo="CARGO",
            saldo_posterior=10,
        ),
    ]
    data = {"advertencias": [f"aviso de la IA {i}" for i in range(12)]}
    avisos = _advertencias_pdf(data, movs)
    assert avisos[0] == f"1 {AVISO_SIN_CONCEPTO}"
    assert avisos[1].startswith("El saldo corrido del archivo no cuadra")
    assert len(avisos) == 10
    assert avisos[2:] == [f"aviso de la IA {i}" for i in range(8)]


# --- Lectura simulada ---------------------------------------------------------


def test_lectura_pdf_devuelve_la_descripcion_con_el_concepto(claude_fake) -> None:
    cliente = claude_fake(
        estado_cuenta,
        {
            "movimientos": [
                {
                    "fecha": "2026-09-24",
                    "descripcion": "SEL TRANSF. INTERBANCARIA SPEI BANORTE\nfact vip 7360",
                    "monto": 1118.12,
                    "tipo": "CARGO",
                    "referencia": "00000000000000000001",
                    "saldo_posterior": None,
                },
                {
                    "fecha": "2026-09-24",
                    "descripcion": "TRANSF ENVIO SPEI BCA INTERNET",
                    "monto": 4462.75,
                    "tipo": "CARGO",
                    "referencia": None,
                    "saldo_posterior": None,
                },
            ],
            "advertencias": [],
        },
    )
    res = parsear_estado_cuenta(_req())

    assert len(cliente.llamadas) == 1  # 1 página ⇒ la llamada de siempre
    assert [m.descripcion for m in res.movimientos] == [
        "SEL TRANSF. INTERBANCARIA SPEI BANORTE fact vip 7360",
        "TRANSF ENVIO SPEI BCA INTERNET",
    ]
    assert res.advertencias == [f"1 {AVISO_SIN_CONCEPTO}"]
    assert "revisa las advertencias" in res.notas
    assert "RENGLONES DE CONTINUACIÓN" in cliente.texto_system()
    assert "renglones de continuación" in cliente.texto_usuario()
    # Una página sola ⇒ sin nota de bloque ni orden de continuación huérfana.
    assert "NO los transcribas ni los pegues" not in cliente.texto_usuario()
