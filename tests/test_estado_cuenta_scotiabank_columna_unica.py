"""Portal NUEVO de Scotiabank (27-oct-2026): el XLSX trae TODO en una sola
columna separada por «|».

Caso real (EJE SCOTIA, «eje scotia.xlsx» del 27-oct-2026): 82 filas en la
columna A — una fila de resumen de la cuenta, los encabezados «Tipo|Moneda|Uso
futuro|Cuenta|Fecha|Referencia|Importe|Tipo(Cargo/Abono)|Saldo|Transacción|
Leyenda 1|Leyenda 2|Leyenda 3|Información adicional SPEI/SPID|Consecutivo»
(con los acentos rotos) y 80 movimientos del más reciente al más antiguo.
El panel respondía «No se reconocieron columnas de monto». Se congela:

 1. el desempaquetado de la columna única (XLSX y CSV dan lo mismo);
 2. la columna de SENTIDO con texto («Tipo(Cargo/Abono)») y el importe en
    positivo; la columna «Tipo» con «CHQ» no se confunde con ella;
 3. la fecha «yyyy/mm/dd»;
 4. la descripción = transacción + leyendas (sin repetir) + la contraparte
    del SPEI (ordenante o beneficiario), que es la forma del PDF con la que
    el API deduplica y completa el concepto;
 5. la referencia de relleno «00000000000000000001» se omite;
 6. la cadena de saldos se valida al revés (archivo descendente) y sigue
    avisando cuando de verdad falta un movimiento;
 7. un movimiento en otra moneda que la cuenta se avisa.

Los datos son SINTÉTICOS con la estructura exacta del archivo real.
"""

import base64
from io import BytesIO

from openpyxl import Workbook

from app.schemas.conciliacion import ConciliacionParseRequest
from app.services.estado_cuenta import (
    _contraparte_spei,
    _desempaquetar_columna_unica,
    parsear_estado_cuenta,
)

# Encabezado tal cual lo exporta el banco (U+FFFD donde iba el acento).
ENCABEZADO = (
    "Tipo|Moneda|Uso futuro|Cuenta|Fecha|Referencia|Importe|Tipo(Cargo/Abono)|Saldo|"
    "Transacci�n|Leyenda 1|Leyenda 2|Leyenda 3|Informaci�n adicional SPEI/SPID|Consecutivo"
)
RESUMEN = "251708921|17501722409|EMPRESA EJEMPLO SA DE CV|MXN|2026/10/27|3|92450.60|4|6221.60|7"
CUENTA = "00000017501722409"
RELLENO = "00000000000000000001"


def _fila(fecha, ref, importe, tipo, saldo, transaccion, l1="", l2="", l3="", info="", n=1):
    campos = [
        "CHQ", "MXN", "000", CUENTA, fecha, ref, f"{importe:.2f}", tipo, f"{saldo:.2f}",
        transaccion, l1, l2, l3, info, f'"{n}"',
    ]
    return "|".join(campos)


# Del más reciente al más antiguo, como el banco; la cadena de saldos cuadra
# leída al revés: 8,300 + 1,700 = 10,000 − 3,275 = 6,725 − 2,500 = 4,225 −
# 446.60 = 3,778.40 + 5,000 = 8,778.40 + 1,910.60 = 10,689 + 85,840 = 96,529.
LINEAS = [
    _fila(
        "2026/10/19", "00000000000000191026", 85840.00, "Abono", 96529.00,
        "TRANSF INTERBANCARIA SPEI", "EA/BANAMEX", "SERVICIO DE TRANSPORTACION AER", "",
        "SERVICIO DE TRANSPORTACION AEREA/BANAMEX DIA-HORA ABONO 19/OCT 02:17:54 "
        "085907804184329269 OPER Y MANTE A ESTAC DE GAS Y SERV00287770125715 40002",
        n=1,
    ),
    _fila(
        "2026/10/15", RELLENO, 1910.60, "Abono", 10689.00,
        "DEVOLUCION SPEI", "DEVOLUCION SPEI", "SANTANDER", "jimmy 1a oct", "", n=2,
    ),
    _fila(
        "2026/10/06", "00000000001305924084", 5000.00, "Abono", 8778.40,
        "Depositos SBC Camara Presentada No.6790754", n=3,
    ),
    _fila(
        "2026/10/06", "00000000000241579048", 446.60, "Cargo", 3778.40,
        "POR DOMICILIACION", "03 POR COBRANZA DOMICILIADA", "RFC/CURP: AEC810901XX1", "IVA: 0.00",
        n=4,
    ),
    _fila(
        "2026/10/05", "00000000001277025708", 2500.00, "Cargo", 4225.00,
        "SEL TRASPASO ENTRE CUENTAS", "SEL TRASPASO ENTRE CUENTAS", "000227707001 09:00:00", "", "",
        n=5,
    ),
    _fila(
        "2026/10/02", RELLENO, 3275.00, "Cargo", 6725.00,
        "SEL TRANSF. INTERBANCARIA SPEI", "BANORTE", "pre fact corner 4480", "",
        "BANORTE pre fact corner 4480 231233015 12:00:17 2026100240044B36K0000078300000 "
        "Fecha Operaci�n: 02/OCT CORNER CAFE SA DE CV      /0726900",
        n=6,
    ),
    _fila(
        "2026/10/01", "00000000000000011026", 1700.00, "Abono", 10000.00,
        "TRANSF INTERBANCARIA SPEI", "/BANAMEX", "TRASLADO TERRESTRE", "",
        "TRASLADO TERRESTRE/BANAMEX DIA-HORA ABONO 01/OCT 04:15:15 085901679640329266 "
        "MODOZEN SAPI DE CV 002180701406817823 40002",
        n=7,
    ),
]


def _xlsx(lineas: list[str]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "MovimientosCTACHEQ2409_27OCT202"
    for linea in lineas:
        ws.append([linea])
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _req_xlsx(lineas=None, **kw) -> ConciliacionParseRequest:
    return ConciliacionParseRequest(
        filename="eje scotia.xlsx",
        file_base64=base64.b64encode(
            _xlsx(lineas or [RESUMEN, ENCABEZADO, *LINEAS])
        ).decode("ascii"),
        **kw,
    )


def _req_csv(texto: str, **kw) -> ConciliacionParseRequest:
    return ConciliacionParseRequest(
        filename="eje scotia.csv",
        file_base64=base64.b64encode(texto.encode("utf-8")).decode("ascii"),
        **kw,
    )


def test_xlsx_de_una_columna_se_desempaqueta_y_lee_los_movimientos() -> None:
    res = parsear_estado_cuenta(_req_xlsx(cuenta_moneda="MXN"))
    assert res.formato == "excel"
    assert res.total == 7
    assert len(res.columnas) == 15
    assert res.columnas[4] == "Fecha"
    assert res.columnas[7] == "Tipo(Cargo/Abono)"
    assert "una sola columna separada por «|»" in res.notas
    assert "15 columnas" in res.notas
    # Cadena de saldos: el archivo va del más reciente al más antiguo y
    # cuadra leída al revés ⇒ sin advertencias; misma moneda ⇒ nada.
    assert res.advertencias == []


def test_fecha_sentido_importe_referencia_y_saldo() -> None:
    res = parsear_estado_cuenta(_req_xlsx())
    m = res.movimientos
    # Orden del archivo (el reporte del API lo respeta: invariante 55).
    assert [x.fecha for x in m] == [
        "2026-10-19", "2026-10-15", "2026-10-06", "2026-10-06",
        "2026-10-05", "2026-10-02", "2026-10-01",
    ]
    assert [x.tipo for x in m] == ["ABONO", "ABONO", "ABONO", "CARGO", "CARGO", "CARGO", "ABONO"]
    assert [x.monto for x in m] == [85840.0, 1910.6, 5000.0, 446.6, 2500.0, 3275.0, 1700.0]
    assert [x.saldo_posterior for x in m] == [
        96529.0, 10689.0, 8778.4, 3778.4, 4225.0, 6725.0, 10000.0,
    ]
    # Referencia real se conserva; la de relleno «…0001» se omite.
    assert m[0].referencia == "00000000000000191026"
    assert m[2].referencia == "00000000001305924084"
    assert m[1].referencia is None
    assert m[5].referencia is None
    # Un banco genérico: sin bruto/comisión/estatus.
    assert all(x.monto_bruto is None and x.comision is None for x in m)


def test_descripcion_transaccion_leyendas_y_contraparte() -> None:
    res = parsear_estado_cuenta(_req_xlsx())
    d = [x.descripcion for x in res.movimientos]
    # Abono recibido: concepto, luego el banco («/BANAMEX», como lo imprime el
    # PDF) y el ordenante (aunque la CLABE venga pegada y recortada).
    assert d[0] == (
        "TRANSF INTERBANCARIA SPEI SERVICIO DE TRANSPORTACION AER EA/BANAMEX "
        "OPER Y MANTE A ESTAC DE GAS Y SERV"
    )
    # Leyenda 1 idéntica a la transacción: no se repite.
    assert d[1] == "DEVOLUCION SPEI SANTANDER jimmy 1a oct"
    assert d[2] == "Depositos SBC Camara Presentada No.6790754"
    assert d[3] == (
        "POR DOMICILIACION 03 POR COBRANZA DOMICILIADA RFC/CURP: AEC810901XX1 IVA: 0.00"
    )
    assert d[4] == "SEL TRASPASO ENTRE CUENTAS 000227707001 09:00:00"
    # Cargo enviado: la leyenda del banco PRIMERO (el dedupe y el concepto del
    # SPEI del API la buscan como prefijo), luego banco, concepto y beneficiario.
    assert d[5] == (
        "SEL TRANSF. INTERBANCARIA SPEI BANORTE pre fact corner 4480 CORNER CAFE SA DE CV"
    )
    assert d[6] == "TRANSF INTERBANCARIA SPEI TRASLADO TERRESTRE /BANAMEX MODOZEN SAPI DE CV"
    # Nunca la CLABE, la clave de rastreo ni la hora.
    assert not any("002180701406817823" in x or "02:18" in x or "B36K" in x for x in d)


def test_csv_con_pipes_da_lo_mismo_que_el_xlsx() -> None:
    xlsx = parsear_estado_cuenta(_req_xlsx())
    csv_ = parsear_estado_cuenta(_req_csv("\n".join([RESUMEN, ENCABEZADO, *LINEAS]) + "\n"))
    assert csv_.formato == "csv"
    assert csv_.columnas == xlsx.columnas
    assert [m.model_dump() for m in csv_.movimientos] == [
        m.model_dump() for m in xlsx.movimientos
    ]
    assert csv_.advertencias == []


def test_cadena_de_saldos_rota_sigue_avisando_en_orden_descendente() -> None:
    roto = [RESUMEN, ENCABEZADO, *LINEAS[:4], *LINEAS[5:]]  # falta el traspaso del 05-oct
    res = parsear_estado_cuenta(_req_xlsx(roto))
    assert res.total == 6
    assert len(res.advertencias) == 1
    assert "no cuadra" in res.advertencias[0]


def test_otra_moneda_se_avisa() -> None:
    usd = LINEAS[2].replace("CHQ|MXN|", "CHQ|USD|", 1)
    lineas = [RESUMEN, ENCABEZADO, *LINEAS[:2], usd, *LINEAS[3:]]
    res = parsear_estado_cuenta(_req_xlsx(lineas, cuenta_moneda="MXN"))
    assert res.total == 7
    assert any("otra moneda" in a and "1 en USD" in a for a in res.advertencias)


def test_columna_tipo_con_chq_no_es_sentido_y_sin_sentido_manda_el_signo() -> None:
    # Mismo archivo pero SIN la columna de sentido: el importe en positivo se
    # lee como abono (lo de siempre), y la columna «Tipo» (CHQ) no interfiere.
    enc = ENCABEZADO.replace("|Tipo(Cargo/Abono)", "|Uso 2")
    lineas = [ln.replace("|Cargo|", "|x|").replace("|Abono|", "|x|") for ln in LINEAS]
    res = parsear_estado_cuenta(_req_xlsx([RESUMEN, enc, *lineas]))
    assert res.total == 7
    assert all(m.tipo == "ABONO" for m in res.movimientos)


def test_sentido_generico_retiro_deposito_y_cargo_abono_de_texto() -> None:
    texto = (
        "Fecha,Concepto,Tipo de movimiento,Importe\n"
        "07/09/2026,ASUR CANCUN,Retiro,\"1,234.56\"\n"
        "08/09/2026,SEL TRASPASO ENTRE CUENTAS,Depósito,\"5,000.00\"\n"
    )
    res = parsear_estado_cuenta(_req_csv(texto))
    assert [m.tipo for m in res.movimientos] == ["CARGO", "ABONO"]
    assert [m.monto for m in res.movimientos] == [1234.56, 5000.0]
    # Encabezado «Cargo/Abono» con texto: NO es la columna de cargos.
    texto2 = "Fecha,Concepto,Cargo/Abono,Importe\n07/09/2026,ASA MERIDA,Cargo,1000.00\n"
    res2 = parsear_estado_cuenta(_req_csv(texto2))
    assert res2.total == 1
    assert res2.movimientos[0].tipo == "CARGO"
    assert res2.movimientos[0].monto == 1000.0


def test_contraparte_spei() -> None:
    assert _contraparte_spei(
        "TRASLADO AEREO/BANAMEX DIA-HORA ABONO 26/OCT 02:18:49 085906784500329966 "
        "MODOZEN SAPI DE CV 002180701406817823 40002"
    ) == "MODOZEN SAPI DE CV"
    assert _contraparte_spei(
        "AFIRME prestamo aerodinamica 298685903 12:14:29 2026102240044B36K0000078598144 "
        "Fecha Operaci�n: 22/OCT AERODINAMICA DE MONTERREY SA D E CV      /0626911"
    ) == "AERODINAMICA DE MONTERREY SA D E CV"
    assert _contraparte_spei(
        "BBVA MEXICO Beh Kay 25 Jun a 22 Oct 302952758 13:48:55 2026102340044BG3C0000000772345 "
        "Fecha Operaci�n: 23/OCT TRANSPORTES JET SET /01269100119363440"
    ) == "TRANSPORTES JET SET"
    # Sin patrón claro, nada (jamás se inventa un nombre).
    assert _contraparte_spei("") is None
    assert _contraparte_spei("000298692983 12:17:00") is None
    assert _contraparte_spei("RFC/CURP: NFM0307091L9") is None


def test_desempaquetar_solo_cuando_de_verdad_es_una_columna() -> None:
    filas, d = _desempaquetar_columna_unica([[RESUMEN], [ENCABEZADO], [LINEAS[0]]])
    assert d == "|"
    assert len(filas[1]) == 15 and filas[1][7] == "Tipo(Cargo/Abono)"
    # Última celda entrecomillada: csv la limpia.
    assert filas[2][14] == "1"
    # Varias columnas reales: se deja igual.
    tabla = [["Fecha", "Concepto", "Cargo"], ["07/09/2026", "ASUR CANCUN", "100"]]
    assert _desempaquetar_columna_unica(tabla) == (tabla, None)
    # Una columna sin separador consistente: se deja igual.
    suelto = [["Estado de cuenta"], ["Cuenta 123"], ["Saldo 1,000.00"]]
    assert _desempaquetar_columna_unica(suelto) == (suelto, None)
