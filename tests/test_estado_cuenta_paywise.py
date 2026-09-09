"""Parser del estado de cuenta de PAYWISE (9-sep-2026).

Cubre: detección por encabezados con 3 variantes de nombre, CSV y XLSX,
neto derivado (bruto − comisión), reembolsos/estatus omitidos, mapeo manual
de columnas (respaldo) y que el parser genérico bancario siga intacto.
"""

import base64
from io import BytesIO

from openpyxl import Workbook

from app.schemas.conciliacion import ConciliacionParseRequest, MapeoColumnasPaywise
from app.services.estado_cuenta import parsear_estado_cuenta


def _req_csv(texto: str, filename: str = "movimientos.csv", mapeo=None) -> ConciliacionParseRequest:
    return ConciliacionParseRequest(
        filename=filename,
        file_base64=base64.b64encode(texto.encode("utf-8")).decode("ascii"),
        mapeo=mapeo,
    )


def _req_xlsx(filas: list[list], filename: str = "paywise.xlsx") -> ConciliacionParseRequest:
    wb = Workbook()
    ws = wb.active
    for fila in filas:
        ws.append(fila)
    buf = BytesIO()
    wb.save(buf)
    return ConciliacionParseRequest(
        filename=filename, file_base64=base64.b64encode(buf.getvalue()).decode("ascii")
    )


# --- Variante 1: encabezados "oficiales" en español -------------------------
CSV_V1 = (
    "Fecha de operación,ID de operación,Cliente,Estatus,Monto bruto,Comisión,Monto neto,Tarjeta\n"
    "03/09/2026 14:22,PW-12345,Juan Pérez,Aprobado,\"$1,000.00\",88.57,911.43,**** 4242\n"
    "04/09/2026 09:10,PW-12346,Ana López,Aprobado,2500.00,221.43,2278.57,**** 1111\n"
    "04/09/2026 10:00,PW-12347,Luis Ruiz,Rechazado,900.00,79.71,820.29,**** 2222\n"
    "05/09/2026 16:45,PW-12340,Juan Pérez,Reembolso,-1000.00,0.00,-1000.00,**** 4242\n"
)


def test_paywise_v1_detecta_y_normaliza() -> None:
    res = parsear_estado_cuenta(_req_csv(CSV_V1))
    assert res.formato == "paywise"
    assert res.columnas[:2] == ["Fecha de operación", "ID de operación"]
    # Rechazado se omite; el reembolso sale como CARGO.
    assert res.total == 3
    assert "omitida" in res.notas
    m1 = res.movimientos[0]
    assert m1.fecha == "2026-09-03"
    assert m1.tipo == "ABONO"
    assert m1.monto == 911.43  # NETO depositado
    assert m1.monto_bruto == 1000.0
    assert m1.comision == 88.57
    assert m1.referencia == "PW-12345"
    assert m1.estatus == "Aprobado"
    assert "Juan Pérez" in (m1.descripcion or "")
    assert "tarjeta 4242" in (m1.descripcion or "")
    assert "ref PW-12345" in (m1.descripcion or "")
    reembolso = res.movimientos[2]
    assert reembolso.tipo == "CARGO"
    assert reembolso.monto == 1000.0
    assert reembolso.referencia == "PW-12340"


# --- Variante 2: encabezados en inglés, sin neto (se deriva) ----------------
CSV_V2 = (
    "Date,Transaction ID,Status,Amount,Fee,Card last4\n"
    "2026-09-03,ab12cd,approved,1000.00,88.57,4242\n"
    "2026-09-04,ef34gh,approved,500.00,44.29,1111\n"
)


def test_paywise_v2_ingles_deriva_neto() -> None:
    res = parsear_estado_cuenta(_req_csv(CSV_V2, "export.csv"))
    assert res.formato == "paywise"
    assert res.total == 2
    assert "bruto − comisión" in res.notas
    m = res.movimientos[0]
    assert m.fecha == "2026-09-03"
    assert m.monto == 911.43
    assert m.monto_bruto == 1000.0
    assert m.comision == 88.57
    assert m.referencia == "ab12cd"


# --- Variante 3: XLSX, encabezados con "Total"/"Depositado"/"Autorización" ---
def test_paywise_v3_xlsx_variantes() -> None:
    req = _req_xlsx(
        [
            [
                "Fecha pago", "No. Autorización", "Concepto", "Total",
                "Comisión total", "Depositado", "Estado",
            ],
            ["06/09/2026", "778899", "Vuelo CUN-CZM", 3000, 265.71, 2734.29, "Pagado"],
            ["07/09/2026", "778900", "Vuelo CUN-MID", 1200, 106.28, 1093.72, "Cancelado"],
        ]
    )
    res = parsear_estado_cuenta(req)
    assert res.formato == "paywise"
    assert res.total == 1
    m = res.movimientos[0]
    assert m.fecha == "2026-09-06"
    assert m.monto == 2734.29
    assert m.monto_bruto == 3000.0
    assert m.comision == 265.71
    assert m.referencia == "778899"
    assert "Vuelo CUN-CZM" in (m.descripcion or "")


# --- Mapeo manual: encabezados que la detección NO reconoce -----------------
CSV_RARO = (
    "Cuando,Clave,Cobrado,Retenido,Nos llegó\n"
    "03/09/2026,X-1,1000.00,88.57,911.43\n"
    "04/09/2026,X-2,200.00,17.71,182.29\n"
)


def test_generico_no_reconoce_y_mapeo_manual_lo_lee_como_paywise() -> None:
    generico = parsear_estado_cuenta(_req_csv(CSV_RARO, "paywise-sept.csv"))
    # Sin mapeo cae al parser genérico (no hay columnas de monto reconocidas).
    assert generico.formato == "csv"
    assert generico.columnas == ["Cuando", "Clave", "Cobrado", "Retenido", "Nos llegó"]
    assert generico.total == 0

    mapeo = MapeoColumnasPaywise(
        fecha="cuando", bruto="Cobrado", comision="retenido", neto="Nos llegó", referencia="Clave"
    )
    res = parsear_estado_cuenta(_req_csv(CSV_RARO, "paywise-sept.csv", mapeo=mapeo))
    assert res.formato == "paywise"
    assert res.total == 2
    assert res.movimientos[0].monto == 911.43
    assert res.movimientos[0].monto_bruto == 1000.0
    assert res.movimientos[0].comision == 88.57
    assert res.movimientos[0].referencia == "X-1"
    assert res.movimientos[1].fecha == "2026-09-04"


def test_mapeo_manual_columna_inexistente_falla_claro() -> None:
    mapeo = MapeoColumnasPaywise(fecha="Cuando", bruto="NoExiste")
    try:
        parsear_estado_cuenta(_req_csv(CSV_RARO, mapeo=mapeo))
    except ValueError as e:
        assert "NoExiste" in str(e)
    else:  # pragma: no cover
        raise AssertionError("debía fallar")


# --- El parser bancario genérico sigue intacto -------------------------------
CSV_BANCO = (
    "Fecha,Concepto,Cargo,Abono\n"
    "01/09/2026,SPEI RECIBIDO CLIENTE,,15000.00\n"
    "02/09/2026,COMPRA GASOLINA,850.50,\n"
)


def test_banco_generico_sin_cambios() -> None:
    res = parsear_estado_cuenta(_req_csv(CSV_BANCO, "hsbc.csv"))
    assert res.formato == "csv"
    assert res.total == 2
    assert res.movimientos[0].tipo == "ABONO"
    assert res.movimientos[0].monto == 15000.0
    assert res.movimientos[0].monto_bruto is None
    assert res.movimientos[0].comision is None
    assert res.movimientos[1].tipo == "CARGO"
    assert res.movimientos[1].monto == 850.5
    assert res.columnas == ["Fecha", "Concepto", "Cargo", "Abono"]


def test_banco_con_columna_comision_no_se_confunde_con_paywise() -> None:
    """Forma de banco (cargo + abono separados) con una columna "Comisión":
    la detección automática se abstiene y cada renglón conserva su signo."""
    csv = (
        "Fecha,Concepto,Cargo,Abono,Comisión,Saldo\n"
        "01/09/2026,SPEI recibido,,5000.00,0.00,15000.00\n"
        "02/09/2026,Pago proveedor,1200.00,,0.00,13800.00\n"
    )
    res = parsear_estado_cuenta(_req_csv(csv, filename="hsbc.csv"))
    assert res.formato == "csv"
    assert [m.tipo for m in res.movimientos] == ["ABONO", "CARGO"]
    assert all(m.monto_bruto is None for m in res.movimientos)


def test_paywise_referencia_no_toma_cantidad_y_tarjeta_solo_digitos() -> None:
    """Sin columna de referencia, "Cantidad" (contiene "id") NO se usa como
    referencia; "Tipo de tarjeta" (Crédito) no se cuela como "tarjeta dito"."""
    csv = (
        "Fecha,Cantidad,Tipo de tarjeta,Monto bruto,Comisión,Monto neto\n"
        "03/09/2026,1,Crédito,1000.00,88.57,911.43\n"
    )
    res = parsear_estado_cuenta(_req_csv(csv, filename="paywise.csv"))
    assert res.formato == "paywise"
    m = res.movimientos[0]
    assert m.referencia is None
    assert "tarjeta dito" not in (m.descripcion or "")
    assert m.monto == 911.43 and m.monto_bruto == 1000.0
