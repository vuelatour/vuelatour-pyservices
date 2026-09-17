# CLAUDE.md — vuelatour-pyservices

Reglas de este microservicio (FastAPI, Python 3.12).

## Principios

- **Aquí solo se renderiza e infiere; el negocio vive en `vuelatour-api`.**
  Los payloads llegan ya calculados (p. ej. el reparto de utilidades): no
  recalcular montos aquí — a lo sumo re-sumar totales de columnas para
  mostrar. Si un número se ve mal, el fix casi siempre es en NestJS.
- **Esquemas pydantic ADITIVOS**: los campos nuevos siempre con default
  (`float = 0`, `str | None = None`) para que el deploy del API y el de
  pyservices no tengan que ser simultáneos (skew tolerante en ambos
  sentidos).
- Toda ruta funcional exige `X-Internal-Token` == `INTERNAL_SHARED_TOKEN`
  (mismo valor que en el API). Nada de este servicio se expone a clientes.

## Render

- PDF: WeasyPrint (reportes de vuelo/cotización, HTML+CSS) y ReportLab
  (reparto). En ReportLab con Helvetica **no usar emojis/unicode raro**
  (⚠ se renderiza como caja) — usar texto ("AVISO:").
- Excel: openpyxl. Si agregas una columna a una tabla, revisa los índices de
  `money_cols`, la fila de totales, `widths` y los `merge_cells` de títulos
  (están por índice de columna). Ejemplo vivo: la hoja «combustible» de los
  DOS balances (individual y general la comparten) creció a 7 columnas el
  11-sep-2026 con «PAGO» (`BalanceAvionGastoFila.pago`, medio de pago YA
  legible que arma el API; sin dato = celda VACÍA — aquí no se inventa un
  default) y hubo que mover encabezado, bordes/relleno de los subtotales,
  los cuatro `merge_cells` y los `widths`. Test:
  `tests/test_balance_combustible_pago.py`.
- El reporte por vuelo debe CUADRAR: el desglose (subtotal + TUAS + pernocta
  + extras + ajuste + IVA) suma el total exacto — no omitir líneas del
  desglose canónico v1.3.
- Hoja del CLIENTE sin horas (15-sep-2026, pedido sobre el PDF del folio
  #314 MID→VSA): `cotizacion_pdf.py` ya NO pinta el bloque «Traslados»
  (traslado inicial/final con hora) ni el renglón «Tiempo de vuelo · H:MM h
  por tramo» de la tarjeta «De un vistazo». En su lugar, la columna derecha
  de `.meta` lleva **`<strong>Fecha del vuelo:</strong> dd/mm/aaaa`** —
  `_fecha_vuelo_html` + `_fecha_corta` (hermano de `_fecha_legible`, SIN
  hora), misma fuente que imprimía «Traslado inicial»
  (`fecha_traslado_inicial`, que el API ya resuelve respetando los tramos
  ocultos), en día de PARED de Cancún. Si `fecha_traslado_final` cae en OTRO
  día de pared la etiqueta pasa a «Fechas del vuelo» y el valor al rango
  «15/09/2026 – 17/09/2026»; sin fecha inicial la línea NO se pinta (jamás
  el «Por confirmar» de `_fecha_legible`). Campos ADITIVOS que siguen en el
  schema y ya no se pintan: `avion_tiempo_tramo_hr`. Como el PDF y la vista
  previa salen del MISMO `_build_html`, el cambio es uno solo y la hoja
  editable del panel (`quote-sheet.tsx`) lo replica: ahí salida y regreso
  CON hora siguen capturándose, pero en un bloque `data-cot-ui` que solo
  existe en edición. La cotización INTERNA de oficina NO cambia. Tests:
  `tests/test_cotizacion_pdf.py` (sección «Fecha del vuelo en `.meta`») y
  los fixtures del panel (`npm run gen:hoja-fixture`).
- Cotización de GRUPO (`/reportes/cotizacion-grupo`, 4-sep-2026):
  `cotizacion_grupo_pdf.py` IMPORTA los helpers de `cotizacion_pdf.py`
  (estilos, mapa, itinerario, ficha de aeronave, regla de matrícula) — no
  copiarlos. Solo pinta: el desglose consolidado, subtotal/IVA/total y el
  precio por persona vienen del API; nunca pintar COMISION_VENDEDOR,
  redondeo ni AJUSTE positivo. Toggles `mostrar_*` mandan. El recibo
  (`/pdf/recibo`) etiqueta "Grupo" cuando viene `grupo_folio`.
- Cotización INTERNA v2 (`/reportes/cotizacion-interna`, 8-sep-2026):
  `cotizacion_interna_pdf.py` es UNA hoja carta para la oficina y NUNCA va
  al cliente (banda «Cotización interna · uso exclusivo de oficina», pie
  «Documento interno · generado … · usuario»). Importa branding/formatos de
  `cotizacion_pdf.py` pero NO `_estilos_base` ni `_mostrar_matricula`: aquí
  la matrícula SIEMPRE se ve. Formato de administración: fecha protagonista
  = DÍA DEL VUELO; tabla de tramos RUTA («Cancun-Merida») · FECHA («26-jun»)
  · DISTANCIA MILLAS · TIEMPO VUELO («01:18», incluye calzos) · COSTO POR
  HORA · TOTAL POR TRAMO + fila TOTAL; si `tramos_ajuste_usd` ≠ 0 una fila
  más con su motivo para que Σ tramos + ajuste == «Servicio aéreo» del
  desglose canónico; TUAS solo las COBRADAS (`tuas_cobradas`); comisión del
  vendedor, redondeo/ajuste, IVA, cobros compactos con comisión bancaria y
  notas internas. NO se pintan (aunque un API viejo los mande): tacómetros,
  horas voladas, avión operativo, traslados, partición, gastos, utilidad ni
  CFDI — eso vive en el reporte del vuelo. Todo llega calculado del API;
  aquí solo se formatea («01:18», «26-jun») o se re-suma la columna
  informativa de millas. Aire (11-sep-2026, «todo muy junto»):
  cuerpo y tablas a 9.5 pt (encabezados/aclaraciones 8–8.5 pt, pie 7.5 pt;
  nada baja de ahí), celdas 3 px × 6 px, `@page` 9/12/10 mm. Presupuesto
  MEDIDO: una hoja mientras tramos + cobros + fila de ajuste ≲ 6 (base
  ≈ 730 px + ≈ 40 px por fila sobre ≈ 984 px útiles); con más, segunda hoja
  ORDENADA (`thead` repetido, `page-break-inside: avoid` por fila). Las
  notas se truncan con «…». Cabecera: «Avión cotizado» (siempre) y, si el
  API manda `aeronave_utilizada` (texto u objeto `{matricula, modelo}`), la
  segunda línea «Avión utilizado: …»; con
  `aeronave_cotizada_vs_utilizada_difiere=true` (lo decide el API por ID, no
  por texto: dos aviones comparten modelo) esa línea va en ÁMBAR con la
  marca «Distinto al cotizado». El legado `aeronave_operativa` de la v1 se
  acepta y NO se pinta.
- Vista previa de cotización (`POST /reportes/cotizacion/preview-html`,
  8-sep-2026): MISMO payload `CotizacionPdfRequest` y MISMO
  `cotizacion_pdf._build_html(req, solo_hoja_1=True)` → `text/html` de SOLO
  la hoja 1 (sin hoja «La aeronave» ni fotos) con CSS de pantalla
  (`_estilos_hoja` + `_estilos_pantalla`, ancho fijo `PREVIEW_ANCHO_PX`=794,
  sin `@page`; el pie de `@page :first` va como `.pie-pantalla`),
  `Cache-Control: no-store`, sin WeasyPrint. Regla: el HTML del PDF (default)
  debe seguir byte-idéntico — jamás una réplica aparte de la hoja 1; un
  cambio en la hoja 1 se hace UNA vez en `_build_html` y sale en ambos.
  `_estilos_base` = `_estilos_page` + `_estilos_hoja` (grupo lo importa tal
  cual). Tests: `tests/test_cotizacion_pdf.py` (sección «Vista previa»).
- CSS de la hoja como ARCHIVO ESTÁTICO (form-as-document, 8-sep-2026):
  `app/static/cotizacion-hoja.css` (cuerpo; TODO selector acotado a la
  raíz `.cot-hoja`, que llevan el <body> del PDF, de la preview y del PDF
  de grupo) y `app/static/cotizacion-fuente.css` (`@font-face` Arimo
  regular/bold woff2 base64, OFL en `Arimo-OFL.txt`; archivo GENERADO, la
  receta está en su cabecera). Python solo los lee: `_estilos_hoja()` =
  fuente + cuerpo y es EXACTAMENTE lo que devuelve
  `GET /reportes/cotizacion/hoja.css` (text/css, max-age 3600) al panel;
  `POST /reportes/cotizacion/mapa-svg` {mapa_puntos} devuelve el <svg> del
  mismo `_mapa_svg` (204 sin puntos). Regla: un estilo de la hoja se cambia
  en el .css UNA vez y sale en PDF, preview y panel; jamás copiar CSS al
  panel ni renombrar las clases del marcado.

## Conciliación

- Estado de cuenta de PAYWISE (9-sep-2026): `estado_cuenta._parse_paywise`
  se elige por ENCABEZADOS (`_detectar_columnas_paywise`: fecha + comisión
  + (bruto o neto); `_norm` translitera acentos) o por `mapeo` manual del
  panel (`MapeoColumnasPaywise`, nombres de columna del archivo; la
  respuesta siempre trae `columnas`). Normaliza a `MovimientoParseado` con
  `monto` = NETO depositado, `tipo` ABONO (CARGO en reembolsos/contracargos
  o neto negativo), `referencia` = ID de operación y los ADITIVOS
  `monto_bruto`/`comision`/`estatus`; rechazadas/canceladas/pendientes se
  omiten (conteo en `notas`); neto ausente = bruto − comisión. `formato` =
  `paywise`. El banco genérico (`_parse_tabular`) no cambia; fechas ISO se
  parsean sin `dayfirst` (pandas 3 invertía mes/día). Tests:
  `tests/test_estado_cuenta_paywise.py`.
- `tabla-xlsx` acepta `hojas` (ADITIVO): una pestaña por hoja con su propia
  tabla (auditoría Paywise: Cotejo / Paywise sin cobro / Cobros sin
  Paywise); sin `hojas` el render es idéntico.
- Estado de cuenta BANCARIO tolerante (15-sep-2026): `_leer_tabla` ya no
  supone que el encabezado está en la fila 0 ni que el CSV es UTF-8 con
  comas. `_filas_crudas` lee TODAS las filas (CSV con el módulo `csv` —
  decodifica utf-8-sig → utf-8 → latin-1 y adivina el separador por
  CONSISTENCIA, no por número de columnas, para que una coma dentro de
  «1,234.56» no gane en un archivo separado por `;`), `_fila_encabezado`
  salta el preámbulo del banco (cuenta, periodo, saldo inicial) buscando la
  primera fila con una aguja de fecha Y una de monto, y `_nombres_columnas`
  garantiza nombres únicos. `_parse_tabular` puebla ahora `referencia`
  (columna propia: ahí viaja la terminación de la tarjeta, '0025830577' ⇒
  0577 — el desempate del auto-cruce vive en el API) y `saldo_posterior`;
  la descripción sigue cayendo a la columna de referencia si el archivo no
  trae concepto. `_to_float` entiende «(1,234.56)» como cargo negativo.
  `_advertencias_saldos` valida la cadena `saldo[i] = saldo[i−1] − cargo +
  abono` y avisa dónde faltan movimientos. Tests:
  `tests/test_estado_cuenta_banco.py`.

## IA

- Visión (tacómetro/tickets) usa el modelo de `ANTHROPIC_MODEL`. La lectura
  de tacómetro recibe `ultimo` (último taco del avión) como ancla de
  magnitud; conservar ese contrato.
- Todo punto de IA degrada a captura manual: los errores devuelven
  `legible=false`/`disponible=false`, nunca 500 por fallo del modelo.
- **Contexto de dominio compartido** (`app/services/_dominio.py`,
  15-sep-2026): flota, aeropuertos IATA, alias de proveedores y leyendas del
  banco (ASUR, ASA, AFAC, MERPAGO*, «SEL TRASPASO ENTRE CUENTAS»…), IVA
  16 %, monedas, banda de TC 15–25, DD/MM/AAAA y hora Cancún. TODA llamada a
  Claude usa `sistema_con_dominio(_PROMPT)`: dos bloques `system` con
  `cache_control`, el dominio PRIMERO (prefijo idéntico en todos los
  endpoints ⇒ el caché de Anthropic lo cobra a ×0.10). Un dato de dominio se
  cambia AQUÍ, una vez. Las listas de aeropuertos/proveedores son
  orientativas; la flota solo sirve para ADVERTIR (nunca para borrar una
  matrícula: la flota crece sin que el archivo se entere). Quien quiera
  validar contra catálogos VIVOS los manda en el request
  (`matriculas_flota`, `terminaciones_validas`, `tipos_documento`,
  `rfcs_propios`).
- **Validación determinista** (`app/services/validaciones_ia.py`): lo que se
  puede comprobar con aritmética o catálogo NO se le pregunta al modelo. Se
  comprueba después de la respuesta y se devuelve en `advertencias`
  (ADITIVO en todas las respuestas de IA; el panel/app deben pintarlo).
  Política de confianza, igual en todos lados: si falla el dato PRINCIPAL
  (total del ticket, litros × precio de una carga, RFC de la constancia,
  sumas de una compra, fechas de un vencimiento) → `confianza` se fuerza a
  ≤ 0.3 (`confianza_calibrada`) y el dato se conserva con su aviso; si falla
  un dato SECUNDARIO (desglose de renglones, matrícula, tarjeta) → ese campo
  se limpia o se marca y la confianza del principal NO se castiga.
- **Prompt del PDF de estado de cuenta**: la descripción debe ser LITERAL
  (nunca parafrasear) — el dedupe del API compara descripciones, así que
  parafrasear hace que re-subir el MISMO PDF duplique todo. El prompt trae
  el formato de Scotiabank/BBVA/Banorte/Santander, pide `referencia`
  íntegra, `saldo_posterior`, el periodo y los totales del resumen, y
  `advertencias`; `_advertencias_pdf` valida cadena de saldos, totales
  impresos vs transcritos y fechas dentro del periodo (una extracción
  incompleta que no truncó no se detecta con `stop_reason`).
- **Sugerencias (conciliación y gasto→vuelo)**: la IA PROPONE, la persona
  confirma. El payload ya lleva el contexto que antes se tiraba (movimiento:
  `tipo`, `referencia`, `cuenta_moneda`, `terminacion_tarjeta_detectada`;
  candidato: `lugar`, `nota`, `categoria`, `tarjeta_terminacion`,
  `matricula`, `faltante`, `tc_implicito`…), el modelo debe citar
  `evidencias[]` y puede devolver `alternativas[]` (top 3) y
  `motivo_sin_match`. Los ids se validan contra la lista SIEMPRE. La
  confianza del modelo se TOPA con `_evidencias_deterministas` (monto
  exacto/faltante, terminación igual o distinta, días entre fechas, moneda
  cruzada dentro de la banda de TC, aeropuerto de la ruta): una tarjeta
  distinta o un ABONO contra un gasto bajan el tope a 0.3 aunque el modelo
  diga 0.99. Las fechas del lado gasto→vuelo se comparan con
  `dias_entre_cancun` (día de PARED en Cancún, invariante 4).
- **TIPOS DEL PAYLOAD: sé LIBERAL con lo que llega** (revisión adversaria
  15-sep-2026). `GastoCandidato.vuelo_folio` estaba declarado `str` y el API
  manda el folio como NÚMERO (`Number(vuelo.folio)`): pydantic v2 **no**
  convierte int → str, así que `/conciliacion/sugerir` respondía **422** en
  cuanto un gasto candidato tenía vuelo — el API lo registraba como
  «pyservices respondió 422» y devolvía `disponible:false`, es decir, TODA la
  sugerencia de IA muerta en producción sin un solo error visible. Hoy es
  `str | int | None` con `field_validator(mode="before")` que normaliza a
  texto, y `tests/test_sugerir_conciliacion.py` congela el payload REAL del
  API. Regla: todo campo nuevo que venga de NestJS se declara con el tipo que
  el API realmente serializa (números como `float`/`int`, ids como `str`) o se
  acepta la unión y se normaliza aquí.
- Dos bugs vivos corregidos el 15-sep-2026 al pasar por aquí: (1) `folio`
  se pedía en los prompts de ticket y de combustible y existía en ambos
  esquemas, pero NUNCA se copiaba a la respuesta — el candado
  anti-duplicados del API nunca se prellenaba; (2) `_parse_matricula`
  exigía un dígito y por eso descartaba en silencio XB-PEV, XA-VGV, XB-ANU
  y XB-IJP (las mexicanas son tres letras). Fuente única ahora:
  `normalizar_matricula`.
- CFDI recibido (`recibida_parse.py`) NO usa IA y así debe quedarse. Se
  parsea con `defusedxml` y, además, se rechaza cualquier XML con
  DTD/ENTITY antes de tocar el parser (XXE / billion laughs), candado que
  funciona aunque la dependencia falte. Valida UUID del timbre, versión y
  —si el API manda `rfcs_propios`— que el receptor sea de la empresa:
  `valido=false` + `motivo` en vez de un objeto a medias.
- Tests de IA: NINGUNO llama a Claude. `tests/conftest.py` expone el
  fixture `claude_fake(modulo, payload)` que parchea el `_client` de ese
  módulo (`estado_cuenta`, `anthropic_vision`, `gasto_vuelo`,
  `compras_extract`, `vencimiento_extract` tienen el suyo) y guarda los
  kwargs para verificar el payload y los bloques `system`.

## Entorno

- Python **3.12** obligatorio (sintaxis `X | None`); el python de sistema de
  esta Mac es 3.9 y NO corre el código — validar con `python3 -m ast` /
  tests en CI o Docker.
- `pytest` + `ruff check app tests` antes de commit. Push a `main` = deploy
  automático en Railway (autorizado sin preguntar). OJO con dos ruidos de
  base en esta Mac: `ruff check app` arrastra ~13 E501/E731 VIEJOS en
  `cfdi_fel.py`, `cotizacion_pdf.py`, `dinero_xlsx.py`, `reparto_*.py`,
  `reporte_vuelo_pdf.py` y `schemas/vision.py` (no los introdujo tu cambio:
  compara antes de culparte), y `tests/test_main.py` /
  `tests/test_recibo_pdf.py` fallan localmente porque WeasyPrint necesita
  `libgobject`/GTK, que no está instalado aquí (en Docker/Railway sí).
- `ANTHROPIC_API_KEY` solo en `.env.local` / variables de Railway. Nunca en
  el repo (ya hubo una key expuesta; está pendiente rotarla).
