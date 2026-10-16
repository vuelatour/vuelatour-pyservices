"""Payloads del AFAC POR VUELO (16-oct-2026, API 0.0.89, invariante 90 del
API) para `tests/test_balance_otros_movimientos_afac_vuelo.py`.

Módulo aparte A PROPÓSITO: solo importa lo que ya existía en el HEAD
`575c816` (los constructores de forma real del cierre de mes y las filas
«AFAC · matrícula» del API 0.0.74), para que las huellas «de siempre» se
calculen corriendo ESTOS mismos constructores con el código de ese HEAD
(`_HUELLAS_0088` del test). Nada de aquí calcula dinero: los montos de las
filas AFAC por vuelo son los de la columna PERMISO AFAC (PROVISIÓN) de los
vuelos del libro XA-TST del payload de forma real (#401 $1,082.50, #402
$564.06, #405 $660.00 = su `totales.permiso_afac_mxn` $2,306.56) y los
costos reales son datos de prueba.
"""

from __future__ import annotations

import copy

from tests import _cierre_mes_payloads as pay
from tests import test_balance_otros_movimientos_afac as t_afac

# ---------------------------------------------------------------------------
# Payloads del API 0.0.88 (los de hoy en producción): forma real del cierre
# de mes + las filas «AFAC · matrícula» con `afac_en_filas` (invariante 62).
# ---------------------------------------------------------------------------


def general_0088(variante: str) -> dict:
    p = pay.general_base(variante)
    om = p["consolidado"]["otros_movimientos"]
    om["filas"].extend(copy.deepcopy([t_afac.FILA_AFAC_TST, t_afac.FILA_AFAC_DOS]))
    om["afac_en_filas"] = True
    p["cierre"] = pay.cierre_abierto_vacio()
    return p


def individual_0088() -> dict:
    p = pay.individual_base()
    p["cierre"] = pay.cierre_abierto_vacio()
    return p


def con_llaves_afac_neutras(p: dict, libro: str) -> dict:
    """Las llaves NUEVAS del 0.0.89 en su valor neutro (null / false): lo que
    mandaría un API que las serializara explícitas sin AFAC por vuelo."""
    p = copy.deepcopy(p)
    if libro == "individual":
        p["pagos_afac"] = None
        return p
    om = p["consolidado"]["otros_movimientos"]
    om["afac_por_vuelo"] = False
    om["afac_resumen"] = None
    for f in [*om["filas"], *om["filas_sueltas"]]:
        f["afac"] = None
    for a in p.get("aviones", []):
        a["pagos_afac"] = None
    p["consolidado"]["pagos_afac"] = None
    return p


# ---------------------------------------------------------------------------
# Payload del API 0.0.89: filas `seccion: 'AFAC_VUELO'` con la forma EXACTA
# del contrato (§3.3 + «CAMBIO DEL CLIENTE»: costo real en EGRESO, reservado
# en INGRESO), `afac_por_vuelo` y `afac_resumen` como ÚLTIMAS llaves.
# ---------------------------------------------------------------------------

VUELO_401 = "11111111-1111-4111-8111-111111111401"
VUELO_402 = "11111111-1111-4111-8111-111111111402"
VUELO_405 = "11111111-1111-4111-8111-111111111405"
VUELO_407 = "11111111-1111-4111-8111-111111111407"
GASTO_1 = "22222222-2222-4222-8222-222222222001"
GASTO_2 = "22222222-2222-4222-8222-222222222002"
GASTO_3 = "22222222-2222-4222-8222-222222222003"
COLOR_TST = "#3B82F6"

CONCEPTO_POR_CAPTURAR = "AFAC costo real · por capturar"
CAPTURADO_AT = "2026-10-16T19:05:00.000Z"


def _nota_reservado(tarifa: str, tc: str, horas: str, monto: str) -> str:
    return (
        "Reservado = tarifa AFAC del avión × T.C. de costos del vuelo × horas "
        "cobradas (columna PERMISO AFAC (PROVISIÓN) de la hoja de vuelos): "
        f"{tarifa} × {tc} × {horas} = {monto}"
    )


def _afac(
    vuelo_id: str,
    matricula: str,
    *,
    horas: float | None,
    tarifa: float | None,
    tc: float | None,
    reservado: float | None,
    real: float | None,
    remanente: float | None,
    capturo: str | None = None,
    notas: str | None = None,
    editable: bool = True,
    motivo: str | None = None,
    cierre_marca: str | None = None,
) -> dict:
    return {
        "vuelo_id": vuelo_id,
        "matricula": matricula,
        "horas_cobradas": horas,
        "tarifa_usd_hr": tarifa,
        "tc_costos": tc,
        "reservado_mxn": reservado,
        "costo_real_mxn": real,
        "remanente_mxn": remanente,
        "capturado_por_nombre": capturo,
        "capturado_at": CAPTURADO_AT if capturo else None,
        "notas": notas,
        "mes_contable": "2026-09",
        "editable": editable,
        "motivo_no_editable": motivo,
        "cierre_marca": cierre_marca,
    }


def filas_afac_vuelo() -> list[dict]:
    """Cuatro filas AFAC por vuelo en el orden del API (cronológico):
    #401 con costo real (remanente 82.50), #402 SIN capturar (egreso null,
    remanente = reservado), #407 de XB-DOS (sin tarifa: solo el costo real,
    remanente −300.00) y #405 con costo real mayor que lo reservado
    (remanente −40.00, rojo)."""
    return [
        {
            "clave": "#401 · Cliente Uno",
            "avion_color": COLOR_TST,
            "estado": "COMPLETADO",
            "fecha_vuelo": "2026-09-03",
            "concepto_egreso": "AFAC costo real · capturó Ale el 16/10/2026",
            "egreso_mxn": 1000.0,
            "fecha_egreso": None,
            "concepto_ingreso": "AFAC reservado · XA-TST · 2.50 h × 25.00 USD/hr × T.C. 17.320000",
            "ingreso_mxn": 1082.5,
            "fecha_ingreso": None,
            "remanente_mxn": 82.5,
            "factura": None,
            "nota_ingreso": _nota_reservado("25.00", "17.320000", "2.50", "$1,082.50"),
            "nota_egreso": "Recibo del gestor 0412",
            "seccion": "AFAC_VUELO",
            "afac": _afac(
                VUELO_401,
                "XA-TST",
                horas=2.5,
                tarifa=25.0,
                tc=17.32,
                reservado=1082.5,
                real=1000.0,
                remanente=82.5,
                capturo="Ale",
                notas="Recibo del gestor 0412",
            ),
        },
        {
            "clave": "#402 · Cliente Dos",
            "avion_color": COLOR_TST,
            "estado": "COMPLETADO",
            "fecha_vuelo": "2026-09-08",
            "concepto_egreso": CONCEPTO_POR_CAPTURAR,
            "egreso_mxn": None,
            "fecha_egreso": None,
            "concepto_ingreso": "AFAC reservado · XA-TST · 1.25 h × 25.00 USD/hr × T.C. 18.050000",
            "ingreso_mxn": 564.06,
            "fecha_ingreso": None,
            "remanente_mxn": 564.06,
            "factura": None,
            "nota_ingreso": _nota_reservado("25.00", "18.050000", "1.25", "$564.06"),
            "seccion": "AFAC_VUELO",
            "afac": _afac(
                VUELO_402,
                "XA-TST",
                horas=1.25,
                tarifa=25.0,
                tc=18.05,
                reservado=564.06,
                real=None,
                remanente=564.06,
            ),
        },
        {
            "clave": "#407 · Cliente Cinco",
            "avion_color": None,
            "estado": "COMPLETADO",
            "fecha_vuelo": "2026-09-18",
            "concepto_egreso": "AFAC costo real · capturó Dani el 15/10/2026",
            "egreso_mxn": 300.0,
            "fecha_egreso": None,
            "concepto_ingreso": None,
            "ingreso_mxn": None,
            "fecha_ingreso": None,
            "remanente_mxn": -300.0,
            "factura": None,
            "seccion": "AFAC_VUELO",
            "afac": _afac(
                VUELO_407,
                "XB-DOS",
                horas=3.0,
                tarifa=None,
                tc=17.95,
                reservado=None,
                real=300.0,
                remanente=-300.0,
                capturo="Dani",
            ),
        },
        {
            "clave": "#405 · Cliente Cuatro",
            "avion_color": COLOR_TST,
            "estado": "COMPLETADO",
            "fecha_vuelo": "2026-09-20",
            "concepto_egreso": "AFAC costo real · capturó Ale el 16/10/2026",
            "egreso_mxn": 700.0,
            "fecha_egreso": None,
            "concepto_ingreso": "AFAC reservado · XA-TST · 1.50 h × 25.00 USD/hr × T.C. 17.600000",
            "ingreso_mxn": 660.0,
            "fecha_ingreso": None,
            "remanente_mxn": -40.0,
            "factura": None,
            "nota_ingreso": _nota_reservado("25.00", "17.600000", "1.50", "$660.00"),
            "seccion": "AFAC_VUELO",
            "afac": _afac(
                VUELO_405,
                "XA-TST",
                horas=1.5,
                tarifa=25.0,
                tc=17.6,
                reservado=660.0,
                real=700.0,
                remanente=-40.0,
                capturo="Ale",
            ),
        },
    ]


def pagos_afac() -> list[dict]:
    """Dos pagos marcados «Pago de AFAC» con pesos ($2,250.00) y uno en USD
    sin T.C. (no suma en la comparación)."""
    return [
        {
            "gasto_id": GASTO_1,
            "fecha": "2026-09-01",
            "detalle": "Permisos Vuelos Privados · Yousefh Pineda Bedolla · XA-TST",
            "matricula": "XA-TST",
            "monto_mxn": 1500.0,
            "moneda_original": None,
            "monto_original": None,
        },
        {
            "gasto_id": GASTO_2,
            "fecha": "2026-09-10",
            "detalle": "Derechos AFAC MMCT · XA-TST · vuelo #402",
            "matricula": "XA-TST",
            "monto_mxn": 750.0,
            "moneda_original": None,
            "monto_original": None,
        },
        {
            "gasto_id": GASTO_3,
            "fecha": "2026-09-22",
            "detalle": "Derechos AFAC en USD · XB-DOS",
            "matricula": "XB-DOS",
            "monto_mxn": None,
            "moneda_original": "USD",
            "monto_original": 40.0,
        },
    ]


def resumen_afac(*, pagos: list[dict] | None = None) -> dict:
    pagos = pagos_afac() if pagos is None else pagos
    return {
        "vuelos": 4,
        "vuelos_con_reservado": 3,
        "vuelos_con_real": 3,
        "vuelos_sin_real": 1,
        "reservado_mxn": 2306.56,
        "costo_real_mxn": 2000.0,
        "remanente_mxn": 306.56,
        "pagado_mxn": 2250.0 if pagos else 0.0,
        "pagos_sin_tc": 1 if pagos else 0,
        "diferencia_pagado_mxn": 250.0 if pagos else -2000.0,
        "reales_fuera_de_libro_mxn": 0.0,
        "pagos": pagos,
    }


PENDIENTES_AFAC = [
    "AFAC: 1 vuelo(s) con AFAC reservado ($564.06) sin costo real capturado — "
    "captúralo en 'otros movimientos' (sección AFAC por vuelo).",
    "AFAC: pagado a AFAC en el periodo $2,250.00 (2 gasto(s) marcado(s) «Pago de "
    "AFAC») vs costos reales capturados $2,000.00: diferencia $250.00. Revisa que "
    "cada pago al gestor esté marcado y que cada vuelo tenga su costo real.",
    "AFAC: 1 pago(s) de AFAC en USD sin T.C.: no suman en la comparación.",
]


def general_0089(variante: str = "mensual", *, filas: list[dict] | None = None) -> dict:
    """El general del API 0.0.89: el de forma real del 0.0.88 SIN las filas
    «AFAC · matrícula» (el 0.0.89 ya no las manda ni manda `afac_en_filas`),
    con las filas AFAC por vuelo al FINAL de `filas`, `afac_por_vuelo` y
    `afac_resumen` (últimas llaves) y los avisos «AFAC: …» en pendientes."""
    p = pay.general_base(variante)
    om = p["consolidado"]["otros_movimientos"]
    om["filas"].extend(filas_afac_vuelo() if filas is None else filas)
    if any(f.get("seccion") == "AFAC_VUELO" for f in om["filas"]):
        om["afac_por_vuelo"] = True
    om["afac_resumen"] = resumen_afac()
    p["consolidado"]["pendientes"] = [*p["consolidado"].get("pendientes", []), *PENDIENTES_AFAC]
    p["cierre"] = pay.cierre_abierto_vacio()
    return p


def individual_0089() -> dict:
    """El individual de XA-TST con sus dos pagos marcados en `pagos_afac`
    (ÚLTIMA llave; la forma de `BalanceAvionHojaGastos` con `tc_por_fila` y
    `gasto_id` por fila, como la hoja 'permisos')."""
    p = individual_0088()
    p["pagos_afac"] = {
        "filas": [
            {
                "fecha": "2026-09-01",
                "categoria": "Permiso",
                "detalle": "Permisos Vuelos Privados · Yousefh Pineda Bedolla",
                "monto_mxn": 1500.0,
                "moneda_original": None,
                "monto_original": None,
                "tc_aplicado": None,
                "gasto_id": GASTO_1,
            },
            {
                "fecha": "2026-09-10",
                "categoria": "Permiso",
                "detalle": "Derechos AFAC MMCT · vuelo #402",
                "monto_mxn": 750.0,
                "moneda_original": None,
                "monto_original": None,
                "tc_aplicado": None,
                "gasto_id": GASTO_2,
            },
        ],
        "total_mxn": 2250.0,
        "usd": 128.06,
        "usd_hr": 17.55,
        "tc_por_fila": True,
    }
    return p
