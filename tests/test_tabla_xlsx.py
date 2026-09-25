"""Tests del export genérico tabla-xlsx: contrato base intacto + resaltes."""

from io import BytesIO

from openpyxl import load_workbook

from app.schemas.tabla import TablaXlsxRequest
from app.services.tabla_xlsx import render_tabla_xlsx


def _req(**extra) -> TablaXlsxRequest:
    return TablaXlsxRequest(
        titulo="Conciliación · Prueba",
        subtitulo="2 movimientos",
        columnas=[
            {"label": "Fecha", "tipo": "texto"},
            {"label": "Cargo", "tipo": "money"},
        ],
        filas=[
            ["2026-08-01", 100.5],
            ["2026-08-02", 200.0],
        ],
        totales=["Totales", 300.5],
        **extra,
    )


def _hoja(xlsx: bytes):
    return load_workbook(BytesIO(xlsx)).active


def test_tabla_sin_resaltes_render_intacto() -> None:
    """Sin `resaltes` el render es el de siempre (contrato de los 4 callers)."""
    ws = _hoja(render_tabla_xlsx(_req()))
    # Título fila 1, subtítulo fila 2, blanco fila 3, encabezado fila 4.
    assert ws.cell(row=4, column=1).value == "Fecha"
    celda = ws.cell(row=5, column=2)
    assert celda.value == 100.5
    assert not celda.font.bold
    # Sin relleno de resalte.
    assert celda.fill.patternType is None


def test_tabla_con_resaltes_pinta_naranja() -> None:
    """La celda marcada queda bold + naranja con fondo suave; el resto igual."""
    ws = _hoja(
        render_tabla_xlsx(
            _req(
                resaltes=[
                    {"fila": 0, "col": 1},
                    # Fuera de rango: se ignora sin tronar el render.
                    {"fila": 99, "col": 1},
                ]
            )
        )
    )
    marcada = ws.cell(row=5, column=2)
    assert marcada.font.bold
    assert str(marcada.font.color.rgb).endswith("ED7D31")
    assert str(marcada.fill.fgColor.rgb).endswith("FFF3E6")
    # Conserva el formato de moneda de su columna.
    assert marcada.number_format == '"$"#,##0.00'
    # La fila NO marcada queda como siempre.
    libre = ws.cell(row=6, column=2)
    assert not libre.font.bold
    assert libre.fill.patternType is None


def test_tabla_varias_hojas() -> None:
    """`hojas` (ADITIVO): un libro con una pestaña por hoja, cada una con su tabla."""
    req = TablaXlsxRequest(
        titulo="Auditoría Paywise",
        columnas=[{"label": "x"}],
        hojas=[
            {
                "titulo": "Cotejo",
                "subtitulo": "3 cruces",
                "columnas": [{"label": "Ref", "tipo": "texto"}, {"label": "Dif", "tipo": "money"}],
                "filas": [["PW-1", 0], ["PW-2", -58.57]],
                "resaltes": [{"fila": 1, "col": 1}],
                "totales": ["Totales", -58.57],
            },
            {
                "titulo": "Paywise sin cobro",
                "columnas": [{"label": "Ref", "tipo": "texto"}],
                "filas": [["PW-9"]],
            },
            {
                # Nombre con caracteres ilegales para Excel → se limpia.
                "titulo": "Cobros / sin: Paywise",
                "columnas": [{"label": "Vuelo", "tipo": "texto"}],
                "filas": [],
            },
        ],
    )
    wb = load_workbook(BytesIO(render_tabla_xlsx(req)))
    assert wb.sheetnames == ["Cotejo", "Paywise sin cobro", "Cobros - sin- Paywise"]
    ws = wb["Cotejo"]
    # Título fila 1, subtítulo fila 2, blanco fila 3, encabezado fila 4.
    assert ws.cell(row=1, column=1).value == "Cotejo"
    assert ws.cell(row=4, column=1).value == "Ref"
    assert ws.cell(row=6, column=2).value == -58.57
    assert ws.cell(row=6, column=2).font.bold
    assert ws.cell(row=7, column=1).value == "Totales"
    ws2 = wb["Paywise sin cobro"]
    # Sin subtítulo: encabezado en fila 3, primera fila en 4.
    assert ws2.cell(row=3, column=1).value == "Ref"
    assert ws2.cell(row=4, column=1).value == "PW-9"


def test_ancho_de_columnas_de_texto_crece_con_su_contenido() -> None:
    """Revisión adversaria 25-sep-2026 (Excel del inventario): una columna de
    TEXTO mide lo de su celda más larga (+2, tope 40) — antes medía lo de su
    encabezado y «Bodega Cancún (anterior)» se cortaba en «(anterior)». Las
    numéricas y las de texto corto quedan EXACTAMENTE como antes; ninguna
    columna se angosta."""
    req = TablaXlsxRequest(
        titulo="Inventario valorizado",
        columnas=[
            {"label": "Ítem"},
            {"label": "Ubicación"},
            {"label": "Código"},
            {"label": "Notas"},
            {"label": "Valor USD (sin T.C.)", "tipo": "money"},
        ],
        filas=[
            ["Aceite multigrado semisintético 15W-50", "Bodega Cancún (anterior)",
             "A1", "x" * 300, 123456789.25],
            ["Filtro", "Bodega del taller de Cozumel", None, "dos\nlíneas", None],
        ],
        totales=["TOTAL", None, None, None, 123456789.25],
    )
    ws = _hoja(render_tabla_xlsx(req))
    anchos = {k: ws.column_dimensions[k].width for k in "ABCDE"}
    assert anchos == {
        "A": len("Aceite multigrado semisintético 15W-50") + 2,  # 40
        "B": len("Bodega del taller de Cozumel") + 2,  # 30 > «(anterior)» 26
        "C": 12,  # texto corto: el mínimo de siempre
        "D": 40,  # tope: 300 caracteres no vuelven la hoja kilométrica
        "E": len("Valor USD (sin T.C.)") + 4,  # numérica: la regla de siempre
    }
