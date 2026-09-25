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
- **Hoja 'inventario' (tiendita): jamás un USD sumado como MXN**
  (22-sep-2026). El cliente vio «Aceite 15w 50 · 30 · $3,300.00 MXN» sobre
  una entrada de 30 × 110 **USD sin tipo de cambio**: la columna sumaba
  dólares rotulados como pesos (en prod, 67 de 68 ENTRADAS están así). Hoy
  «VALOR A COSTO MXN» lleva SOLO pesos reales y a su derecha aparece
  **«VALOR A COSTO USD (sin T.C.)»** (`MONEY_USD` = `"$"#,##0.00`, el único
  lugar del libro con símbolo) con su propio total y, bajo la tabla, la nota
  «N productos tienen costo en USD sin tipo de cambio: su valor se muestra
  en dólares y no entra al total en pesos». La columna solo existe cuando el
  API manda el desglose nuevo Y hay algo que mostrar (`_hay_usd_sin_tc`:
  `filas_sin_tc`, `total_valor_usd` o alguna fila con `valor_costo_usd` /
  `sin_tc`): con un API viejo —o con un inventario 100 % en pesos— la hoja
  es exactamente la de antes, 9 columnas. Al entrar la columna las del
  periodo corren un lugar (`off`): encabezados, fila TOTALES, `n_cols` (9+off
  para título/notas/merges) y `anchos` se mueven juntos — ejemplo vivo de la
  regla de openpyxl de arriba. La conversión NO se hace aquí ni allá: sin
  T.C. no hay peso que mostrar. Tests:
  `tests/test_balance_inventario_xlsx.py` (caso real, mixto MXN+USD, nota en
  singular/plural, payload viejo ⇒ hoja idéntica).
- **Excel de la REPOSICIÓN de caja chica** (24-sep-2026,
  `POST /reportes/caja-chica-reposicion.xlsx`, `CajaChicaReposicionRequest`
  todo con default y `extra="ignore"`, servicio `caja_chica_xlsx.py`).
  Pedido: «al momento de reembolsar la caja de cada uno … un Excel con lo
  que estoy reembolsando». Una hoja: título, ENCABEZADO (etiqueta en A:C
  —ancho suficiente para «Por reponer antes de esta reposición», que el
  export genérico cortaba a 12 caracteres— valor en D, nota en E:H), TABLA
  de 14 columnas con fecha REAL `dd/mm/yyyy` y montos `"$"#,##0.00`, fila
  TOTAL (Σ Gasto y Σ Reintegro/ajuste que manda el API), bloque TOTALES
  (el renglón `destacado` va en negritas) y AVISOS en naranja; `resaltar`
  pinta fecha y captura de la fila. Impresión horizontal a una hoja de
  ancho. **Aquí no se calcula nada**: filas, saldos por fila, totales y
  diferencia llegan del API (`caja-chica-saldo.util.ts`). Las columnas son
  paridad MANUAL con `COLUMNAS_EXCEL_CAJA` del API (que las usa en su
  respaldo genérico cuando este endpoint todavía no está desplegado).
  **Texto que empieza con «=» se escribe como TEXTO** (`_texto_literal`,
  revisión 24-sep-2026): openpyxl vuelve fórmula cualquier cadena con «=»
  al inicio y una nota de gasto «= 3 taxis» daba un libro que Excel abre con
  error. Ojo: los DEMÁS libros (Libro Dinero, balances) todavía no tienen
  esta guarda — hoy no hay notas así en prod, pero el folio de la factura y
  las notas son texto libre.
  Tests: `tests/test_caja_chica_xlsx.py`.
- El reporte por vuelo debe CUADRAR: el desglose (subtotal + TUAS + pernocta
  + extras + ajuste + IVA) suma el total exacto — no omitir líneas del
  desglose canónico v1.3.
- **Tipo de cambio: hasta 6 DECIMALES, sin ceros de cola** (17-sep-2026,
  pedido sobre la cotización #314). Fuente única `app/services/_formato.py`
  → `_tc_txt` (16.991632 → «16.991632»; 16.9916 → «16.9916»; 18 → «18»;
  `None` → cadena vacía y quien llama decide si pinta «—»). Antes cada
  documento usaba `f"{tc:g}"`, que son seis CIFRAS SIGNIFICATIVAS: el T.C.
  16.991632 con el que 5,885.25 USD dan $100,000.00 MXN se imprimía
  «16.9916», y quien remultiplicaba ese texto obtenía $99,999.81 («cuando
  son muchos decimales como que siempre cambia»). Lo usan los TRES PDFs de
  cotización (`cotizacion_pdf` total MXN, `cotizacion_grupo_pdf` total MXN,
  `cotizacion_interna_pdf` resumen + total MXN + TUA/línea en pesos + T.C.
  del cobro), el recibo (`recibo_pdf`) y el reporte por vuelo (PDF y Excel).
  En Excel el formato de celda `TC` de `dinero_xlsx`/`balance_avion_xlsx`
  pasó a `0.00####` (2 a 6 decimales) por lo mismo. El panel escribe el
  MISMO texto con `fmtTc` y el API persiste `numeric(12,6)`: si los tres no
  coinciden, el operador vuelve a ver dos números para el mismo dato. Aquí
  NO se recalcula el total en pesos — llega del API (`monto_total_mxn`).
  Tests: «T.C. con TODOS sus decimales» en `tests/test_cotizacion_pdf.py`,
  `tests/test_cotizacion_grupo_pdf.py`, `tests/test_cotizacion_interna_pdf.py`
  y `tests/test_recibo_pdf.py`.
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
  = DÍA DEL VUELO; tabla de tramos RUTA («CUN–MID» desde el 22-sep-2026;
  antes el nombre largo «Cancun-Merida») · FECHA («26-jun»)
  · DISTANCIA MILLAS · TIEMPO VUELO (HRS) («1.19» en horas decimales desde
  el 24-sep-2026 —antes «01:18»—, incluye calzos; ver la entrada «Tiempo de
  vuelo en horas decimales» más abajo) · COSTO POR
  HORA · TOTAL POR TRAMO + fila TOTAL; si `tramos_ajuste_usd` ≠ 0 una fila
  más con su motivo para que Σ tramos + ajuste == «Servicio aéreo» del
  desglose canónico; TUAS solo las COBRADAS (`tuas_cobradas`); comisión del
  vendedor, redondeo/ajuste, IVA, cobros compactos con comisión bancaria y
  notas internas. NO se pintan (aunque un API viejo los mande): tacómetros,
  horas voladas, avión operativo, traslados, partición, gastos, utilidad ni
  CFDI — eso vive en el reporte del vuelo. Todo llega calculado del API;
  aquí solo se formatea («1.19», «26-jun») o se re-suma la columna
  informativa de millas. Aire (11-sep-2026, «todo muy junto»):
  cuerpo y tablas a 9.5 pt (encabezados/aclaraciones 8–8.5 pt, pie 7.5 pt;
  nada baja de ahí), celdas 3 px × 6 px, `@page` 9/12/10 mm. Presupuesto
  MEDIDO: una hoja mientras tramos + cobros + fila de ajuste ≲ 6 (base
  ≈ 730 px + ≈ 40 px por fila sobre ≈ 984 px útiles) — **SUPERSEDIDO el
  22-sep-2026 por la poda: hoy ≲ 11 filas, ver la entrada de abajo**; con
  más, segunda hoja ORDENADA (`thead` repetido, `page-break-inside: avoid`
  por fila). Las
  notas se truncan con «…». Cabecera: «Avión cotizado» (siempre) y, si el
  API manda `aeronave_utilizada` (texto u objeto `{matricula, modelo}`), la
  segunda línea «Avión utilizado: …»; con
  `aeronave_cotizada_vs_utilizada_difiere=true` (lo decide el API por ID, no
  por texto: dos aviones comparten modelo) esa línea va en ÁMBAR con la
  marca «Distinto al cotizado». El legado `aeronave_operativa` de la v1 se
  acepta y NO se pinta.
- Cotización interna: PODA de repeticiones y UNA hoja de verdad
  (22-sep-2026, capturas del cliente sobre la #329, que salía en DOS hojas).
  Salieron del documento las filas **«Método de cobro» y «T.C. USD→MXN»** de
  la ficha (el método PREVISTO subió al `<h2>` de «Cobros» —
  `_metodo_previsto_txt`, «Cobros · previsto: Transferencia · comisión
  terminal 8.86 %», también en la fila «Sin cobros registrados»: es el campo
  que decide IVA 16 % vs 0 % y el único sitio con el % de TERMINAL pactado;
  el T.C. sigue en «Total MXN» y en «Equiv. USD»), el **nombre largo del
  tramo** (la celda RUTA abre con la ABREVIATURA «CUN–PCE» y las marcas en
  gris en la MISMA línea, sin `<br>`: 89 → 40 px por fila; RESPALDO
  obligatorio al nombre largo si falta un IATA o es la fila consolidada), la
  **nota gris del IVA** (`iva_nota` ya no se pinta y «16 % de $X» solo si el
  renglón de arriba NO es la base) y el **gris del «Servicio aéreo»** (horas
  × tarifa ya están en «Horas cotizadas»). Refuerzo NO opcional: la fila
  **«Cobrables» carga el motivo y el importe del ajuste** («Horas pactadas
  1.75 h: +$165.00 sobre Σ tramos»); sin él, en las 6 de cada 10
  cotizaciones con ajuste la tabla cerraría en «TOTAL Σ tramos» y el
  desglose en «Servicio aéreo» sin nada que los concilie. Además la hoja
  **incrusta Arimo** (`_estilos_fuente` del PDF del cliente, importada) —
  antes declaraba `'Helvetica Neue', Arial` sin `@font-face` y el contenedor
  de Railway caía a su sans (~8-12 % más ancho): lo que cabía en la Mac se
  desbordaba en producción — y el bloque de Cobros dejó de ser `.bloque`
  (`page-break-inside: avoid`): si algo se desborda bajan dos filas, no el
  bloque entero. Presupuesto nuevo MEDIDO en píxeles con la fuente
  incrustada (Chrome, caja útil 984 px): base ≈ 530 px + 23-40 px por fila
  ⇒ **una hoja mientras filas ≲ 11** (antes ≲ 6). Con los payloads REALES de
  prod: #329 = 718 px y el folio con más tramos (8) = 808 px. Tests:
  `tests/test_cotizacion_interna_pdf.py`.
- **Tiempo de vuelo en horas decimales (24-sep-2026, API 0.0.33)**. Pedido
  del cliente con la captura de una CUN→PTU→CUN: «la parte de tiempo de
  vuelo, lo podemos manejar solo en decimales por favor? … se nos hacen
  raros los tiempos». Cada tramo valía 1.19166… h (125 nm / 120 kt + 0.15
  de calzo) ⇒ «01:12», el total 2.38333 h ⇒ «02:23», y 01:12 + 01:12 ≠
  02:23. Ahora la columna es **«TIEMPO VUELO (HRS)»** (`ENCABEZADO_TIEMPO`
  = «Tiempo vuelo (hrs)»; el CSS pone las mayúsculas) con **2 decimales
  FIJOS** por tramo, en la fila TOTAL y en la fila consolidada, y la nota
  `NOTA_TRAMOS` = «Tiempo de vuelo en horas decimales (1.50 = 1 h 30 min) e
  incluye calzos» (+ calzos, millas, consolidado como siempre). Se PINTAN
  los campos ADITIVOS del API `tiempo_horas` (por tramo; `None` ⇒ «—») y
  `tramos_tiempo_total_horas`, que el API reparte por **RESIDUO MAYOR**
  (`repartirHorasDecimales`, `tramos-costeados.util.ts`) para que **Σ tramos
  mostrados == total mostrado**. Si el payload no trae
  `tramos_tiempo_total_horas` (API previo), `_tiempos_de_tabla` reparte aquí
  con el espejo EXACTO `_repartir_horas_decimales` (`_horas_decimal` =
  medio hacia arriba en ENTEROS con `floor(x + 0.5)` —el `Math.round` del
  API—, jamás `round()`/`:.2f`, que fallan en 1.005). `tiempo_hhmm` /
  `tramos_tiempo_total_hhmm` se siguen aceptando (compatibilidad) y se
  IGNORAN; `_hhmm` se retiró. Solo presentación: ningún importe cambia.
  Misma tabla de casos que el API y el panel (`CASOS_HORAS_DECIMALES` en
  `tests/test_cotizacion_interna_pdf.py`). Tras tocar este armador, el panel
  regenera sus fixtures con `npm run gen:hoja-interna-fixture` (5 escenarios,
  `interna-ptu` = la captura del cliente).
- **Conceptos SIN IVA debajo del IVA** (22-sep-2026, mismo pedido: «que se
  entienda visualmente que no lleva IVA»), en los TRES documentos y en la
  hoja WYSIWYG del panel. Fuente única: `cotizacion_pdf.particionar_por_iva`
  (pura) + `LineaIva` + las etiquetas `ETIQUETA_BASE_GRAVABLE` / `ETIQUETA_
  SIN_IVA` / `ETIQUETA_SUBTOTAL`; `cotizacion_grupo_pdf` y
  `cotizacion_interna_pdf` la IMPORTAN (jamás copiar). Qué hace: separa las
  filas en gravables y exentas (PERNOCTA por clave y `aplica_iva=false` —
  extras exentos y comisión de terminal), pone las exentas DEBAJO del IVA
  bajo el rótulo «No causan IVA» y convierte el renglón de arriba en la BASE
  GRAVABLE («Subtotal gravable» = `iva_base_usd`). Es OBLIGATORIO redefinir
  ese renglón: valía `total − IVA` e incluía los exentos, así que bajarlos y
  dejarlo igual rompería las dos lecturas del Excel de la oficina a la vez
  (ni suma lo de arriba, ni el 16 % de él da el IVA). NADA se recalcula: es
  presentación. Dos candados: **activación CONDICIONAL** (hace falta algún
  exento ≠ 0 *y* IVA > 0 — sin eso el documento sale byte-idéntico al de
  antes: 226 de 231 cotizaciones de prod) y **verificación de las
  identidades** (`Σ gravables == base` y `base + IVA + Σ exentos == total`,
  medio centavo de tolerancia) con **degradación** al layout de siempre si
  no cuadran — el caso real es la línea AJUSTE canónica, que mezcla lo que
  entra a la base con el redondeo que se suma DESPUÉS del IVA (1 de 231).
  Jamás un documento con una columna que no suma. **TERCERA identidad,
  `base × iva_pct == IVA`, obligatoria SOLO cuando la base se DERIVÓ** (el
  PDF del cliente y el de grupo, porque el API todavía no manda
  `iva_base_usd`): una base derivada cumple la segunda identidad POR
  CONSTRUCCIÓN —se despejó de ella— y la primera también cuando el desvío
  vive en una fila de arriba, que es justo lo que hace el REDONDEO
  automático (el armador del cliente lo ABSORBE en «Servicio aéreo»). Sin
  ese candado el PDF imprimiría «Subtotal gravable $3,725.33 / IVA (16 %)
  $594.67» cuando el 16 % de ese renglón son $596.05 — el renglón que la
  oficina lee como «sobre esto se calcula el impuesto». Sin `iva_pct` (API
  viejo) también se degrada: nunca se rotula «Subtotal gravable» un número
  que no se pudo verificar. Verificado contra prod: las 136 cotizaciones con
  IVA cumplen `iva = base × pct` al centavo. El PDF del cliente empezó
  además a LEER `ExtraPdf.aplica_iva` (existía en el esquema y el API ya lo
  mandaba; hasta hoy se ignoraba). Campos ADITIVOS nuevos: `iva_base_usd` en
  `CotizacionPdfRequest` y `CotizacionGrupoPdfRequest` y `aplica_iva` en
  `CotizacionGrupoLineaPdf` — sin ellos la base se deriva `total − IVA −
  Σ exentos` (re-suma de columna, lo único permitido aquí). CSS: `.exentos-row`
  en `app/static/cotizacion-hoja.css` (cliente y grupo; el panel lo copia con
  `npm run sync:hoja-css`) y en `_estilos_interno`. Los 3 sha256 de
  `tests/test_cotizacion_pdf.py` se refrescaron SOLO por la regla nueva del
  CSS compartido: **ningún cuerpo se movió** (el payload «completo» sí trae
  pernocta con IVA, pero su IVA sintético es el 16 % de una base que INCLUYE
  la pernocta —cosa que el motor nunca hace—, así que cae en la degradación).
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
- **La HOJA INTERNA también se comparte con el panel** (Fase 2.1 del
  rediseño del cotizador, 22-sep-2026): la PANTALLA de la cotización pasa a
  ser la hoja interna —la completa, la que la oficina llevaba en Excel— y el
  PDF del cliente queda como salida, así que el documento interno se publica
  igual que el del cliente desde el 8-sep. Tres piezas, mismo patrón:
  (1) `app/static/cotizacion-interna.css` — el CUERPO del CSS que antes
  vivía dentro de `_estilos_interno`, con TODO selector acotado a la raíz
  **`.cot-interna`** (`cotizacion_interna_pdf.CLASE_RAIZ`, hermana de
  `.cot-hoja`) y un bloque que neutraliza el preflight de Tailwind con los
  valores por defecto del agente de usuario (h2 en negrita, img inline,
  table en `display:table`, celdas de 1 px: sin efecto en WeasyPrint).
  Python solo lo lee (`_estilos_cuerpo_interno`, `@lru_cache`);
  `_estilos_hoja_interna()` = fuente + cuerpo y `_estilos_interno(pie)` =
  `_estilos_page_interno(pie)` + eso. En el archivo NO van las reglas
  `@page` ni el `margin: 0` del `<body>`: en el panel no hay página que
  maquetar. (2) El marcado va envuelto en `<div class="cot-interna">` que
  arma **`_cuerpo_interno_html`**, fuente ÚNICA usada por el PDF
  (`_build_html`) y por la vista previa — jamás una réplica: el test afirma
  que la preview es substring del HTML del PDF. (3) Dos rutas espejo de las
  del cliente: `GET /reportes/cotizacion-interna/hoja.css` (text/css,
  max-age 3600, EXACTAMENTE lo que incrusta el PDF) y
  `POST /reportes/cotizacion-interna/preview-html` (mismo payload
  `CotizacionInternaPdfRequest`, devuelve SOLO el cuerpo en `text/html`,
  `no-store`, sin `<style>` ni `@page`, sin importar WeasyPrint; el CSS va
  por la otra ruta). El panel copia el .css con su script de sync y compara
  su hoja contra fixtures generados con `preview-html`. **Paridad
  verificada**: el cuerpo del PDF es byte-idéntico al de antes salvo el
  `<div>` raíz, y el CSS es el mismo texto salvo el prefijo `.cot-interna`
  (medido con 6 payloads reales, incluidos la #329 y el de 8 tramos; alto y
  ancho renderizados idénticos en Chrome headless antes/después). Regla: un
  estilo de la hoja interna se cambia en el .css UNA vez y sale en PDF,
  preview y panel; renombrar una clase del marcado rompe el contrato con el
  panel. Tests: última sección de `tests/test_cotizacion_interna_pdf.py`.

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

## Lectura de facturas emitidas (PDF)

- `POST /facturacion/leer-pdf-emitida` {`pdf_b64`} (24-sep-2026, registro
  «Facturas emitidas» que Mari captura a mano): el API
  (`facturas-emitidas/leer-archivo`) manda el PDF de la factura y prellena
  serie, folio, UUID, fecha, RFC/nombre de emisor y receptor, subtotal, IVA,
  total, moneda, método y forma de pago. Servicio `app/services/
  emitida_pdf.py`, esquemas `LeerPdfEmitida*` en `schemas/facturacion.py`
  (todo con default). **Determinista: pypdf + regex, SIN IA** — y así debe
  quedarse (como `recibida_parse.py`). Nada se guarda aquí; duplicados,
  razón social emisora y cliente sugerido los decide el API.
- **Nunca 500.** 422 SOLO si no es un PDF utilizable (`PdfNoValidoError`:
  base64 roto, sin `%PDF-` en los primeros 1024 bytes, > 11 MB). Protegido
  con contraseña, escaneado (< 30 alfanuméricos), roto o lento ⇒ 200 con
  `texto_extraido=false` + aviso. pypdf corre en un hilo DAEMON con tope
  `TOPE_SEGUNDOS` = 20 s (el API corta a los 30 s; un hilo atorado no frena
  el apagado en un redeploy) y lee máximo 5 páginas; la INTERPRETACIÓN del
  texto corre en otro hilo con lo que sobre del mismo tope (mínimo 1 s): un
  texto patológico —miles de «Total:» apilados— es cuadrático y quedaba
  fuera del tope (la respuesta sale a tiempo; el hilo termina solo). Lo que
  no se encuentra
  sale `None`: el API arma «No encontré: …». Mejor vacío que inventado.
- Extractor TOLERANTE (cada PAC arma su PDF distinto) y probado con el
  layout del ejemplo real (Seguros Inbursa, CFDI 4.0): etiquetas en una
  línea y valores en la siguiente («Emisor: RFC emisor: Régimen…» ⇒
  «NOMBRE RFC 601»), basura binaria del QR (se limpia con NFKC + sin
  caracteres de control), una SEGUNDA línea «Emisor: 26300 Póliza…» sin RFC
  que no pisa al emisor, el RFC del PAC (cadena `||1.1|…|RFC|` y «RFC
  proveedor de certificación») excluido, «Prima total» sí es total y
  «Subtotal»/«Importe total con letra» no, «No. de serie del certificado»
  no es la serie y «Folio fiscal» no es el folio. También: «Serie y Folio:
  A-123», factura del SAT («Efecto de comprobante», montos mezclados con
  etiquetas), bloques «Receptor / Nombre / RFC», totales apilados en columna
  o en fila (se emparejan por posición; si no cuadran, `None`), UUID de
  «CFDI relacionados» nunca se toma como el propio. Avisos ADITIVOS
  (strings; el API los pinta como `LECTURA_PDF`): varios UUID sin etiqueta,
  moneda distinta de MXN/USD, CFDI de Egreso/Pago y «Subtotal + IVA no da
  el total impreso» (salvo descuento/retenciones impresos).
- Revisión adversaria (24-sep-2026), trampas que YA mordían y quedaron
  congeladas en tests: (1) el encabezado de conceptos «CANTIDAD MONEDA IVA
  IMPORTE» cortaba la búsqueda con «Moneda IVA» y nunca llegaba a «MONEDA:
  MXN» ⇒ solo cuenta un código del catálogo c_Moneda (`_MONEDAS_ISO`) y
  «DLS/DLLS/US$» es USD; (2) «Retención IVA: 106.67» se tomaba como EL IVA ⇒
  una etiqueta IVA precedida de «Retención/Ret./retenido» se descarta
  (`_RETENCION_ANTES`) y las retenciones (IVA y/o ISR, `_PREFIJO_RETENCION`)
  entran al cuadre; (3) «Importe total: 1,000.00» de un renglón de concepto
  le ganaba al «Total: 1,160.00» del cuadro ⇒ con subtotal e IVA leídos, el
  total es el que CUADRA (`_extraer_total`), y si ninguno cuadra el de
  siempre + aviso; (4) «Folio interno 555» antes de «Folio: 777» daba el
  folio «INTERNO-555» ⇒ primero las etiquetas con «:»/«#»/«No.» y una serie
  separada por ESPACIO debe ir en mayúsculas; «Serie:» vacía no toma la
  palabra de la línea siguiente («Fecha 2026…»). Fecha también «2026/09/24».
- Dependencia `pypdf>=6.1.3` en `requirements.txt` (piso con los arreglos de
  DoS/bombas de descompresión de pypdf; el Dockerfile instala de
  ahí; `uv.lock` está en `.gitignore`). El import es tolerante: sin pypdf el
  router arranca y la lectura degrada a aviso. Tests:
  `tests/test_leer_pdf_emitida.py` con PDFs SINTÉTICOS de ReportLab (el PDF
  real del cliente NO se commitea).

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
