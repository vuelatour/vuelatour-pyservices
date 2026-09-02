"""Nombres de hoja del Balance GENERAL (renombre 1-sep-2026, modelo mental
del equipo): el payload `gastos_empresa` (gastos sin avión ni vuelo) se pinta
en la hoja 'otros gastos' (antes 'gastos VuelaTour') y los parciales del
reparto manual por avión en 'repartidos a aviones' (antes 'otros gastos').
SOLO cambia el renderer del general: el libro INDIVIDUAL y los campos del
contrato del API quedan igual."""

from io import BytesIO

from openpyxl import load_workbook

from app.schemas.reportes import (
    BalanceAvionRequest,
    BalanceGeneralRequest,
)
from app.services.balance_avion_xlsx import (
    render_balance_avion_xlsx,
    render_balance_general_xlsx,
)

_GASTOS_EMPRESA = {
    "filas": [
        {
            "fecha": "2026-08-05",
            "categoria": "INDIRECTO",
            "detalle": "Renta de oficina agosto",
            "monto_mxn": 12000.0,
        },
    ],
    "total_mxn": 12000.0,
    "usd": 600.0,
    "usd_hr": None,
}

_OTROS_GASTOS = {
    "filas": [
        {
            "fecha": "2026-08-15",
            "categoria": "INDIRECTO",
            "detalle": "XB-ABC · Nómina repartida (50 %)",
            "monto_mxn": 5000.0,
            "matricula": "XB-ABC",
            "avion_color": "#FF0000",
        },
    ],
    "total_mxn": 5000.0,
    "usd": 250.0,
    "usd_hr": None,
}


def _general(**extra) -> BalanceGeneralRequest:
    return BalanceGeneralRequest(
        periodo_desde="2026-08-01",
        periodo_hasta="2026-08-31",
        consolidado=BalanceAvionRequest(
            matricula="FLOTA",
            periodo_desde="2026-08-01",
            periodo_hasta="2026-08-31",
            otros_gastos=_OTROS_GASTOS,
        ),
        **extra,
    )


def test_general_renombra_hojas_de_gastos():
    data = render_balance_general_xlsx(_general(gastos_empresa=_GASTOS_EMPRESA))
    wb = load_workbook(BytesIO(data))

    # Nombres nuevos (1-sep-2026); el viejo 'gastos VuelaTour' ya no existe.
    assert "otros gastos" in wb.sheetnames
    assert "repartidos a aviones" in wb.sheetnames
    assert "gastos VuelaTour" not in wb.sheetnames
    # Pestañas hermanas: 'otros gastos' (empresa) va JUNTO y ANTES de
    # 'repartidos a aviones' (la posición que el equipo ya conocía).
    idx = wb.sheetnames.index
    assert idx("otros gastos") + 1 == idx("repartidos a aviones")

    # 'otros gastos' = los gastos de EMPRESA (payload gastos_empresa), con el
    # título exacto (sin sufijo de matrícula) y su fila.
    ws = wb["otros gastos"]
    assert ws.cell(row=1, column=1).value == (
        "Otros gastos — VuelaTour (sin avión ni vuelo)"
    )
    detalles = {c[2].value for c in ws.iter_rows(min_col=1, max_col=4)}
    assert "Renta de oficina agosto" in detalles

    # 'repartidos a aviones' = los parciales del reparto manual por avión
    # (payload otros_gastos del consolidado, que NO cambia de nombre).
    ws_r = wb["repartidos a aviones"]
    assert ws_r.cell(row=1, column=1).value == (
        "Otros gastos repartidos a aviones — FLOTA"
    )
    detalles_r = {c[2].value for c in ws_r.iter_rows(min_col=1, max_col=4)}
    assert "XB-ABC · Nómina repartida (50 %)" in detalles_r


def test_general_sin_gastos_empresa_conserva_repartidos():
    # API viejo (skew): sin `gastos_empresa` no hay hoja de empresa, pero los
    # parciales se pintan igual con el nombre nuevo.
    data = render_balance_general_xlsx(_general())
    wb = load_workbook(BytesIO(data))
    assert "otros gastos" not in wb.sheetnames
    assert "gastos VuelaTour" not in wb.sheetnames
    assert "repartidos a aviones" in wb.sheetnames


def test_individual_conserva_su_hoja_otros_gastos():
    # El libro INDIVIDUAL no cambia: su hoja de parciales sigue siendo
    # 'otros gastos' y no existe 'repartidos a aviones'.
    data = render_balance_avion_xlsx(
        BalanceAvionRequest(matricula="XB-ABC", otros_gastos=_OTROS_GASTOS)
    )
    wb = load_workbook(BytesIO(data))
    assert "otros gastos" in wb.sheetnames
    assert "repartidos a aviones" not in wb.sheetnames
    assert wb["otros gastos"].cell(row=1, column=1).value == (
        "Otros gastos — XB-ABC"
    )
