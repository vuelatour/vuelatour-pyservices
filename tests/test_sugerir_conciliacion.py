"""Sugerencia de conciliación (15-sep-2026): payload rico y confianza calibrada.

La IA PROPONE y la persona confirma, así que lo importante aquí es que
nunca pueda afirmar más de lo que los hechos comprobables permiten.
"""

import json

from app.schemas.conciliacion import (
    ConciliacionSugerirRequest,
    GastoCandidato,
    MovimientoSinConciliar,
)
from app.services import estado_cuenta
from app.services.estado_cuenta import sugerir_conciliacion

MOV = MovimientoSinConciliar(
    fecha="2026-09-08",
    monto=125.82,
    descripcion="AEROPUERTO DE COZUMEL",
    tipo="CARGO",
    referencia="9155656256",
    cuenta_alias="Scotiabank GASTOS GNRAL",
    cuenta_moneda="MXN",
)

GASTO_A = GastoCandidato(
    id="g-1",
    fecha="2026-09-07",
    monto=125.82,
    proveedor="ASUR",
    moneda="MXN",
    medio_pago="TARJETA_CORP",
    tarjeta_terminacion="6256",
    categoria="ATERRIZAJE",
    lugar="Aeropuerto de Cozumel",
    nota="Aeropuerto de Cozumel · plataforma",
    matricula="XB-PEV",
    vuelo_folio="1234",
    capturado_por="Luis",
)
GASTO_B = GastoCandidato(
    id="g-2",
    fecha="2026-09-07",
    monto=125.82,
    proveedor="ASUR",
    moneda="MXN",
    medio_pago="TARJETA_CORP",
    tarjeta_terminacion="0577",
    lugar="Aeropuerto de Cozumel",
)


def _pedir(candidatos: list[GastoCandidato], mov: MovimientoSinConciliar = MOV):
    return ConciliacionSugerirRequest(movimiento=mov, candidatos=candidatos)


def test_sin_candidatos_no_llama_a_claude_y_explica() -> None:
    res = sugerir_conciliacion(_pedir([]))
    assert res.gasto_id_sugerido is None
    assert res.confianza == 0.0
    assert res.motivo_sin_match and "Sin candidatos" in res.motivo_sin_match


def test_payload_lleva_el_contexto_que_antes_se_tiraba(claude_fake) -> None:
    cliente = claude_fake(
        estado_cuenta,
        {"gasto_id_sugerido": "g-1", "confianza": 0.9, "razon": "ok", "evidencias": []},
    )
    sugerir_conciliacion(_pedir([GASTO_A]))

    texto = cliente.texto_usuario()
    payload = json.loads(texto[texto.index("{") : texto.rindex("}") + 1])
    mov = payload["movimiento"]
    assert mov["referencia"] == "9155656256"
    assert mov["tipo"] == "CARGO"
    assert mov["cuenta_moneda"] == "MXN"
    # La terminación se deduce de la referencia porque empata con la tarjeta.
    assert mov["terminacion_tarjeta_detectada"] == "6256"

    cand = payload["candidatos"][0]
    for campo in ("lugar", "nota", "categoria", "tarjeta_terminacion", "matricula"):
        assert campo in cand, campo
    # Los campos vacíos no viajan (el JSON va en cada llamada).
    assert "tc_implicito" not in cand

    system = cliente.ultima["system"]
    assert "CONTEXTO DE DOMINIO" in system[0]["text"]
    assert "evidencias" in system[1]["text"]
    assert "terminacion_tarjeta_detectada" in system[1]["text"]


def test_monto_exacto_y_tarjeta_igual_permiten_confianza_alta(claude_fake) -> None:
    claude_fake(
        estado_cuenta,
        {
            "gasto_id_sugerido": "g-1",
            "confianza": 0.95,
            "razon": "mismo aeropuerto",
            "evidencias": ["el proveedor coincide"],
        },
    )
    res = sugerir_conciliacion(_pedir([GASTO_A, GASTO_B]))
    assert res.gasto_id_sugerido == "g-1"
    assert res.confianza == 0.95
    assert any("Monto exacto" in e for e in res.evidencias)
    assert any("6256" in e for e in res.evidencias)
    assert any("1 día" in e for e in res.evidencias)
    # Las evidencias del modelo se conservan después de las duras.
    assert res.evidencias[-1] == "el proveedor coincide"


def test_tarjeta_distinta_tumba_la_confianza_aunque_el_monto_cuadre(claude_fake) -> None:
    claude_fake(
        estado_cuenta,
        {"gasto_id_sugerido": "g-2", "confianza": 0.99, "razon": "monto exacto"},
    )
    res = sugerir_conciliacion(_pedir([GASTO_A, GASTO_B]))
    assert res.gasto_id_sugerido == "g-2"
    assert res.confianza <= 0.3
    assert any("otra tarjeta" in e for e in res.evidencias)


def test_id_inventado_se_descarta_con_motivo(claude_fake) -> None:
    claude_fake(
        estado_cuenta,
        {"gasto_id_sugerido": "g-999", "confianza": 0.9, "razon": "inventado"},
    )
    res = sugerir_conciliacion(_pedir([GASTO_A]))
    assert res.gasto_id_sugerido is None
    assert res.confianza == 0.0
    assert res.motivo_sin_match is not None


def test_alternativas_solo_con_ids_reales_y_sin_repetir_al_elegido(claude_fake) -> None:
    claude_fake(
        estado_cuenta,
        {
            "gasto_id_sugerido": "g-1",
            "confianza": 0.8,
            "razon": "ok",
            "alternativas": [
                {"gasto_id": "g-2", "confianza": 0.4, "razon": "otra tarjeta"},
                {"gasto_id": "g-1", "confianza": 0.9, "razon": "repetido"},
                {"gasto_id": "no-existe", "confianza": 0.7, "razon": "inventado"},
            ],
        },
    )
    res = sugerir_conciliacion(_pedir([GASTO_A, GASTO_B]))
    assert [a.gasto_id for a in res.alternativas] == ["g-2"]
    assert res.alternativas[0].confianza == 0.4


def test_pago_parcial_se_juzga_contra_el_faltante(claude_fake) -> None:
    parcial = GastoCandidato(
        id="g-3",
        fecha="2026-09-08",
        monto=500.0,
        monto_vinculado=374.18,
        faltante=125.82,
        moneda="MXN",
        proveedor="ASUR",
    )
    claude_fake(estado_cuenta, {"gasto_id_sugerido": "g-3", "confianza": 0.95, "razon": "faltante"})
    res = sugerir_conciliacion(_pedir([parcial]))
    assert res.confianza == 0.9
    assert any("faltante" in e for e in res.evidencias)


def test_abono_contra_gasto_queda_marcado(claude_fake) -> None:
    abono = MOV.model_copy(update={"tipo": "ABONO"})
    claude_fake(estado_cuenta, {"gasto_id_sugerido": "g-1", "confianza": 0.95, "razon": "ok"})
    res = sugerir_conciliacion(_pedir([GASTO_A], abono))
    assert res.confianza <= 0.3
    assert any("ABONO" in e for e in res.evidencias)


def test_moneda_cruzada_usa_la_banda_de_tipo_de_cambio(claude_fake) -> None:
    usd = GastoCandidato(
        id="g-usd", fecha="2026-09-07", monto=7.0, moneda="USD", proveedor="FBO", tc_implicito=17.97
    )
    claude_fake(estado_cuenta, {"gasto_id_sugerido": "g-usd", "confianza": 0.95, "razon": "usd"})
    res = sugerir_conciliacion(_pedir([usd]))
    assert res.confianza == 0.8
    assert any("17.97" in e for e in res.evidencias)

    fuera = usd.model_copy(update={"tc_implicito": 90.0})
    claude_fake(estado_cuenta, {"gasto_id_sugerido": "g-usd", "confianza": 0.95, "razon": "usd"})
    res = sugerir_conciliacion(_pedir([fuera]))
    assert res.confianza <= 0.4


def test_payload_real_del_api_no_revienta_por_el_folio_numerico() -> None:
    """El API manda `vuelo_folio` como NÚMERO (`Number(vuelo.folio)`).

    Pydantic v2 no convierte int → str: con `vuelo_folio: str` la petición
    entera respondía 422 y el API lo registraba como «pyservices respondió
    422» — es decir, la sugerencia de IA quedaba muerta en producción en
    cuanto un gasto candidato tenía vuelo. Se acepta int o str y se
    normaliza a texto.
    """
    req = ConciliacionSugerirRequest.model_validate(
        {
            "movimiento": {
                "fecha": "2026-09-08",
                "monto": 125.82,
                "descripcion": "AEROPUERTO DE COZUMEL",
                "referencia": "9155656256",
                "tipo": "CARGO",
                "cuenta_alias": "Scotiabank GASTOS GNRAL",
                "cuenta_moneda": "MXN",
                "terminacion_tarjeta_detectada": "6256",
            },
            "candidatos": [
                {
                    "id": "g-1",
                    "fecha": "2026-09-07",
                    "monto": 125.82,
                    "moneda": "MXN",
                    "tc_implicito": None,
                    "proveedor": "ASUR",
                    "monto_vinculado": 0,
                    "faltante": 125.82,
                    "medio_pago": "TARJETA_CORP",
                    "tarjeta_terminacion": "6256",
                    "categoria": "ATERRIZAJE",
                    "lugar": "Aeropuerto de Cozumel",
                    "nota": "Aeropuerto de Cozumel",
                    "matricula": "XB-PEV",
                    "vuelo_folio": 1234,
                    "capturado_por": "Luis",
                }
            ],
        }
    )
    assert req.candidatos[0].vuelo_folio == "1234"
