"""Estado de cuenta BANCARIO (15-sep-2026): CSV/XLSX tolerante y PDF con IA.

Cubre lo que rompía la importación del 15-sep y lo que la dejaba coja:
preámbulo antes de los encabezados, latin-1, separador ';', la columna
REFERENCIA (donde viaja la terminación de la tarjeta) y el saldo corrido
como detector de movimientos faltantes.
"""

import base64
from io import BytesIO

import pytest
from openpyxl import Workbook

from app.schemas.conciliacion import ConciliacionParseRequest
from app.services import estado_cuenta
from app.services.estado_cuenta import parsear_estado_cuenta

# Export real de banco: 3 líneas de preámbulo, separador ';' y columna Saldo.
CSV_BANCO = (
    "Estado de cuenta;;;;;\n"
    "Cuenta;0025830577;Moneda;MXN;;\n"
    "Periodo;01/09/2026 al 30/09/2026;;;;\n"
    ";;;;;\n"
    "Fecha;Concepto;Referencia;Cargo;Abono;Saldo\n"
    "07/09/2026;ASUR CANCUN;0025830577;1,234.56;;8,765.44\n"
    "07/09/2026;AEROPUERTO DE COZUMEL;9155656256;125.82;;8,639.62\n"
    "08/09/2026;SEL TRASPASO ENTRE CUENTAS;174465;;5,000.00;13,639.62\n"
)


def _req(texto: str, filename: str = "scotiabank.csv", encoding: str = "utf-8"):
    return ConciliacionParseRequest(
        filename=filename,
        file_base64=base64.b64encode(texto.encode(encoding)).decode("ascii"),
    )


def test_csv_con_preambulo_separador_y_referencia() -> None:
    res = parsear_estado_cuenta(_req(CSV_BANCO))
    assert res.formato == "csv"
    assert res.total == 3
    assert res.columnas == ["Fecha", "Concepto", "Referencia", "Cargo", "Abono", "Saldo"]

    primero = res.movimientos[0]
    assert primero.fecha == "2026-09-07"
    assert primero.tipo == "CARGO"
    assert primero.monto == 1234.56
    assert primero.descripcion == "ASUR CANCUN"
    # La referencia ya NO se tira: de aquí sale la terminación 0577.
    assert primero.referencia == "0025830577"
    assert primero.saldo_posterior == 8765.44

    ultimo = res.movimientos[2]
    assert ultimo.tipo == "ABONO"
    assert ultimo.monto == 5000.0
    assert ultimo.referencia == "174465"
    # Cadena de saldos consistente ⇒ sin advertencias.
    assert res.advertencias == []


def test_csv_latin1_no_revienta() -> None:
    texto = CSV_BANCO.replace("ASUR CANCUN", "ASUR CANCÚN")
    res = parsear_estado_cuenta(_req(texto, encoding="latin-1"))
    assert res.total == 3
    assert res.movimientos[0].descripcion == "ASUR CANCÚN"


def test_cadena_de_saldos_rota_avisa_de_movimientos_faltantes() -> None:
    roto = CSV_BANCO.replace("8,639.62", "7,000.00")
    res = parsear_estado_cuenta(_req(roto))
    assert res.total == 3
    assert len(res.advertencias) == 1
    assert "no cuadra" in res.advertencias[0]
    assert "faltar movimientos" in res.advertencias[0]


def test_cargo_entre_parentesis_es_cargo() -> None:
    texto = (
        "Fecha,Concepto,Referencia,Importe\n"
        "07/09/2026,ASA MERIDA,0025830585,\"(1,000.00)\"\n"
    )
    res = parsear_estado_cuenta(_req(texto))
    assert res.movimientos[0].tipo == "CARGO"
    assert res.movimientos[0].monto == 1000.0


def test_xlsx_con_preambulo() -> None:
    wb = Workbook()
    ws = wb.active
    ws.append(["Reporte de movimientos"])
    ws.append(["Cuenta", "GASTOS GNRAL"])
    ws.append([])
    ws.append(["Fecha", "Descripcion", "Referencia", "Cargo", "Abono", "Saldo"])
    ws.append(["07/09/2026", "ASA CANCUN", "0025830572", "500.00", "", "1000.00"])
    buf = BytesIO()
    wb.save(buf)
    req = ConciliacionParseRequest(
        filename="banco.xlsx", file_base64=base64.b64encode(buf.getvalue()).decode("ascii")
    )
    res = parsear_estado_cuenta(req)
    assert res.total == 1
    assert res.movimientos[0].referencia == "0025830572"
    assert res.movimientos[0].saldo_posterior == 1000.0


# --- PDF con IA ------------------------------------------------------------

_PDF_OK = {
    "movimientos": [
        {
            "fecha": "2026-09-07",
            "descripcion": "ASUR CANCUN",
            "monto": 1234.56,
            "tipo": "CARGO",
            "referencia": "0025830577",
            "saldo_posterior": 8765.44,
        },
        {
            "fecha": "2026-09-08",
            "descripcion": "SEL TRASPASO ENTRE CUENTAS",
            "monto": 5000,
            "tipo": "ABONO",
            "referencia": None,
            "saldo_posterior": 13765.44,
        },
    ],
    "periodo_inicio": "2026-09-01",
    "periodo_fin": "2026-09-30",
    "total_cargos": 1234.56,
    "total_abonos": 5000,
    "advertencias": [],
}


def _req_pdf(banco: str | None = None):
    return ConciliacionParseRequest(
        filename="estado.pdf", file_base64="UERGRkFLRQ==", banco=banco, cuenta_moneda="MXN"
    )


def test_pdf_lee_referencia_y_saldo_y_manda_el_dominio(claude_fake) -> None:
    cliente = claude_fake(estado_cuenta, _PDF_OK)
    res = parsear_estado_cuenta(_req_pdf(banco="Scotiabank"))

    assert res.formato == "pdf"
    assert res.total == 2
    assert res.movimientos[0].referencia == "0025830577"
    assert res.movimientos[0].saldo_posterior == 8765.44
    assert res.advertencias == []
    assert res.uso_ia is not None and res.uso_ia.input_tokens == 120

    # El bloque de dominio va primero y el prompt exige descripción LITERAL.
    system = cliente.ultima["system"]
    assert len(system) == 2
    assert "CONTEXTO DE DOMINIO" in system[0]["text"]
    assert "LITERAL" in system[1]["text"]
    assert "saldo_posterior" in system[1]["text"]
    assert "SCOTIABANK" in system[1]["text"]
    # El banco de la cuenta viaja en el mensaje del usuario.
    assert "Scotiabank" in cliente.texto_usuario()


def test_pdf_avisa_cuando_los_totales_no_cuadran(claude_fake) -> None:
    payload = {**_PDF_OK, "total_cargos": 9000}
    claude_fake(estado_cuenta, payload)
    res = parsear_estado_cuenta(_req_pdf())
    assert res.total == 2
    assert any("falta" in a for a in res.advertencias)
    assert res.notas  # el panel ve que hay algo que revisar


def test_pdf_avisa_fechas_fuera_del_periodo(claude_fake) -> None:
    payload = {
        **_PDF_OK,
        "movimientos": [{**_PDF_OK["movimientos"][0], "fecha": "2016-09-07"}],
        "total_cargos": None,
        "total_abonos": None,
    }
    claude_fake(estado_cuenta, payload)
    res = parsear_estado_cuenta(_req_pdf())
    assert any("fuera del periodo" in a for a in res.advertencias)


def test_pdf_truncado_falla_con_instrucciones(claude_fake) -> None:
    claude_fake(estado_cuenta, _PDF_OK, stop_reason="max_tokens")
    with pytest.raises(ValueError, match="demasiados movimientos"):
        parsear_estado_cuenta(_req_pdf())
