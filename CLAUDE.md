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
- **Hoja 'inventario': UTILIDAD DE LA TIENDA en dólares** (25-sep-2026, API
  0.0.35). Pedido: «la utilidad… de los productos que compramos, el precio
  que le ponemos en el costo se le saca el 25 %… necesitamos ver las
  ganancias». Desde ese día toda SALIDA a un avión sin precio se cobra a
  costo FIFO × (1 + margen) (config del API `inventario_margen_venta_pct`,
  25 %) y las 10 salidas del 01-sep se re-preciaron: TODAS las ventas de
  prod son USD sobre costo USD sin T.C., cuya utilidad antes no cabía en
  «UTILIDAD MXN» (celda vacía). Campos ADITIVOS: fila `vendido_usd`,
  `utilidad_usd`, `ventas_sin_utilidad` (null ⇒ 0); hoja
  `total_vendido_usd`, `total_utilidad_usd`, `filas_utilidad_incompleta`
  (null ⇒ 0), `margen_venta_pct`. Una salida cuenta en UNA sola moneda (lo
  decide `ventaDeSalida` del API); aquí jamás se convierte ni se suma USD
  con MXN. Con ventas en dólares (`_hay_venta_usd`: alguna fila con
  `vendido_usd`/`utilidad_usd` o un total USD no None — 0.0 sí cuenta)
  entran **«VENDIDO\nUSD» y «UTILIDAD\nUSD»** (`MONEY_USD`) entre «UTILIDAD
  MXN» y «MATRÍCULAS», con sus totales en la fila TOTALES (si el API no
  manda el total, re-suma de ESA columna: `_total_usd_de_columna`). Es un
  SEGUNDO desplazamiento (`off_venta` = 2, solo mueve MATRÍCULAS),
  independiente del `off` de «VALOR A COSTO USD (sin T.C.)»: con los dos la
  hoja tiene 12 columnas y `anchos` = `[34, 42, 16, 18, 12, 14, 12, 14, 14,
  14, 14, 24]`; `n_cols` (título, notas y bloque 2) = `col_matriculas`.
  Notas: `_NOTA_INVENTARIO_VENTA_USD` al pie cuando hay esas columnas;
  `_nota_margen(pct)` al pie cuando llega `margen_venta_pct` (aunque todo
  sea en pesos: «Desde el 25-sep-2026 toda salida sin precio se cobra al
  avión a costo FIFO + 25 % (utilidad de la tienda).»; 0 ⇒ «…a costo FIFO,
  sin utilidad»); y nota ROJA bajo la tabla con `filas_utilidad_incompleta`
  > 0 (o, sin el conteo, filas con `ventas_sin_utilidad` > 0): «N
  producto(s) con ventas en pesos sobre costo en dólares sin tipo de
  cambio: su utilidad no se puede calcular y no aparece en ninguna
  columna.». El BLOQUE 2 (detalle de salidas en pesos con el T.C. promedio
  del libro) NO cambia. **Payload viejo ⇒ hoja idéntica**: verificado
  byte a byte (todos los miembros del .xlsx salvo `docProps/core.xml`) y
  congelado con `_firma_hoja` (huella del layout calculada con el código
  del 24-sep; si cambia, un API 0.0.34 ya no ve su hoja). El **Excel del
  inventario** (panel → `GET /v1/inventory/items/export`) NO tiene código
  aquí: lo arma el export genérico `/pdf/tabla-xlsx` con las columnas que
  manda el API (Ubicación, «Utilidad (MXN)» y «Utilidad (USD)» separadas);
  un test congela que cada moneda cae en su columna. **Anchos del genérico
  (revisión adversaria 25-sep-2026)**: medía cada columna por su ENCABEZADO
  (mínimo 12) y «Ubicación» salía de 13 — «Bodega Cancún (anterior)» se
  cortaba JUSTO en la marca «(anterior)», igual que tres de las cinco
  ubicaciones del catálogo y casi todo nombre de producto (la celda vecina
  siempre trae número: el texto no se desborda, se corta). Hoy
  `tabla_xlsx._ancho_columna`: una columna de TEXTO crece a su celda más
  larga + 2 (tope `ANCHO_TEXTO_MAX` = 40); las numéricas y las de texto
  corto quedan EXACTAMENTE como antes y ninguna se angosta. Aplica a todos
  los exports del genérico (solo anchos: valores, formatos y totales no
  cambian). **Plantilla de alta masiva** (`inventario_masivo.py`): la fila
  de ejemplo decía «Bodega Cancún» —ya no es del catálogo: quien la copiaba
  daba de alta productos «(anterior)»— y hoy dice «Oficina nueva»; las
  instrucciones explican la columna (vacío = sin ubicación; lo que no está
  en el catálogo queda «(anterior)»). `_es_ejemplo_intacto` acepta TAMBIÉN
  el ejemplo viejo: una plantilla descargada antes y subida con el ejemplo
  intacto daba de alta el aceite del ejemplo con 12 piezas. Tests:
  `tests/test_balance_inventario_xlsx.py`, `tests/test_tabla_xlsx.py`,
  `tests/test_inventario_masivo_ubicacion.py`.
- **Inventario: ÚLTIMO PRECIO DE COMPRA + T.C. oficial del día** (25-sep-2026,
  API 0.0.36). Pedido: «que los precios se ajusten en automático al último
  registrado… el remanente de agosto ahora igual su costo de 30 DLS» y «en el
  tipo de cambio, que sea los mismos que usan en las cotizaciones (del día de
  la venta)». El API deja el FIFO: costo = último precio de compra (congelado
  en cada salida al registrarla), valorizado = existencia × último precio ×
  T.C. oficial de HOY, y compras/ventas en dólares LLEGAN YA EN PESOS con el
  T.C. oficial de su día (misma función que la cotización). **Aquí no se
  convierte nada: solo cambian textos.** Campos ADITIVOS:
  `BalanceHojaInventario.regla_costo` (`'ULTIMO_PRECIO'`,
  `REGLA_COSTO_ULTIMO_PRECIO`) y `tc_hoy` (`TcHoy`: `tc`, `fecha_dato`,
  `fuente`, todo opcional — un `tc` null no tumba el libro; revisión
  adversaria: el API usa el MISMO nombre `tc_hoy` como NÚMERO suelto en el
  PATCH de costo, así que `_tc_hoy_liberal` acepta un número —o texto
  numérico— como `{tc}` y cualquier otra cosa como None, y `_tc_liberal`
  vuelve None un `tc` ilegible/NaN: antes cualquiera de esos era un 422 del
  Balance general ENTERO por una nota), fila
  `vendido_usd_original`/`utilidad_usd_original` (NO se pintan: ya cuentan
  en pesos, sumarlos contaría doble) y `CardexLibroRequest.nota` (se pinta
  tal cual en el subtítulo tras «Montos en MXN», sin su punto final). Con
  `regla_costo` (`_es_ultimo_precio`; otro valor o null ⇒ API previo):
  nota al pie `_NOTA_INVENTARIO_ULTIMO_PRECIO` sin «FIFO» con
  `_frase_valor_a_costo` («… al T.C. oficial de hoy (17.6729,
  25/09/2026)», T.C. con `_tc_txt` y fecha del DATO con `_fecha`; sin dato
  no cita número), `_NOTA_INVENTARIO_USD_ULTIMO_PRECIO` /
  `_NOTA_INVENTARIO_VENTA_USD_ULTIMO_PRECIO` (las columnas en dólares ya
  solo son respaldo de filas sin T.C.; tras la migración de datos no llega
  ninguna y se apagan solas con la MISMA lógica de siempre),
  `_nota_margen(pct, ultimo_precio=True)` («al último precio de compra +
  25 %»), el índice del Balance general («detalle de salidas con costo al
  último precio de compra») y **nota bajo el BLOQUE 2** (`_NOTA_BLOQUE2_TC`,
  solo si hay salidas): el detalle de salidas convierte con el T.C. de cada
  GASTO (o el promedio del libro) y la utilidad por ítem con el del día de la
  venta — en las 10 salidas del 01-sep los gastos no traen T.C. y las dos
  utilidades en pesos difieren ≈ $63 MXN en septiembre; la hoja lo explica.
  Esas dos notas van ENVUELTAS con alto estimado (`_pinta_nota_envuelta`,
  `_alto_nota`; los `anchos` se calculan al inicio de `_hoja_inventario`):
  la nota de siempre, combinada sin wrap, se corta al ancho de la hoja.
  `_NOTA_REFACCIONES_GENERAL` conserva «costo FIFO»: solo se pinta sin
  `inventario` (API anterior al 30-ago, que costeaba FIFO). **Payload previo
  ⇒ hoja idéntica** (las dos huellas `_FIRMA_*` siguen iguales; también con
  `regla_costo`/`tc_hoy` en null). El Excel del inventario sigue sin código
  aquí (export genérico): un test congela las columnas del API 0.0.36
  («Último precio de compra» número + «Moneda» + Valor/Vendido/Utilidad en
  MXN). Tests: `tests/test_balance_inventario_xlsx.py` (sección «ÚLTIMO
  PRECIO»), `tests/test_cardex_libro_xlsx.py`.
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
  las notas son texto libre. Excepción (30-sep-2026): las dos columnas de
  la etiqueta de factura de los BALANCES la importan (`from
  app.services.caja_chica_xlsx import _texto_literal`): FACTURA VUELATOUR
  de la hoja maestra y «factura vuelatour» de 'otros movimientos' (si solo
  una la tuviera, el Balance general seguiría abriendo con error). La
  «FACTURA VUELATOUR» del Libro Dinero sigue sin ella.
  Tests: `tests/test_caja_chica_xlsx.py`.
- **Comisión del vendedor como GASTO** (28-sep-2026, API 0.0.39, invariante
  31 del API). Pedido: «¿cómo registro el pago de la comisión a Saab para que
  aparezca en otros movimientos? Si lo capturo como "Otros gastos VuelaTour"
  queda duplicado». El API agrega la categoría `COMISION_VENDEDOR` («Comisión
  del vendedor») y el gasto real ligado al vuelo REEMPLAZA a la PROVISIÓN en
  su fila de «otros movimientos» (Balance general) y de «Otros ingresos»
  (Libro Dinero). **Aquí no se calcula nada**: concepto («pago comisión
  vendedor (X) · gasto real», «· parcial: faltan $… MXN», «· excede $… MXN»,
  filas de solo-egreso «· sin comisión cobrada en la cotización»), monto,
  fecha, remanente y nota de la celda llegan hechos y se pintan tal cual.
  Lo único que cambia son 5 LEYENDAS, y **solo con banderas ADITIVAS** que
  el API manda cuando hay pagos reales: `BalanceHojaOtrosMovimientos.
  hay_pago_vendedor_real: bool | None` (fila 2 de 'otros movimientos', nota
  del RESUMEN y nota de la hoja maestra del GENERAL; `_hay_pago_vendedor_real`
  exige `is True`) y `DineroXlsxRequest.utilidades_comision_vendedor_pagada_mxn:
  float | None` (fila 2 de 'Otros ingresos' con el campo presente; la frase
  «Pagado al vendedor con gasto real en este periodo: $X MXN.» al final de la
  fila 9 de 'utilidades' solo con monto > 0). **Sin bandera el Excel es
  BYTE-IDÉNTICO** al de antes (verificado miembro a miembro del .xlsx salvo
  `docProps/core.xml`, con la bandera ausente, null y false): las leyendas de
  hoy («PROVISIÓN … mientras no exista el gasto real») siguen siendo
  verdaderas mientras no haya gasto real. El libro INDIVIDUAL no trae 'otros
  movimientos' y su nota de la hoja maestra no cambia nunca (`general=True`
  es condición). Visión (IA de tickets) SIN cambio: `valid_cats` descarta la
  categoría y la IA nunca la sugiere; todos los demás esquemas leen
  `categoria` como `str` (un código nuevo no da 422). Tests:
  `tests/test_comision_vendedor_notas.py` (textos viejos LITERALES sin
  bandera, textos nuevos con ella, individual, concepto y nota tal cual,
  rutas con el payload nuevo).
- **FACTURA VUELATOUR en la hoja maestra de los balances** (30-sep-2026, API
  0.0.45). Pedido de Marie para Ale, con la foto del «reporte horas XA-VGV»:
  «se ocupa que diga el num de factura que nosotros emitimos del servicio …
  y si puede salir en el reporte general también». Campo ADITIVO
  `BalanceAvionVuelo.factura_vuelatour: str | None` que el API resuelve con
  su cascada ÚNICA `etiquetasFacturaDeVuelos` (CFDI timbrado vivo →
  facturas EMITIDAS vigentes «A-0424, A-0430» → `vuelo.factura_folio`
  tecleado → etiqueta del estatus «Facturado» / «Factura elaborada y
  enviada» → null): la MISMA etiqueta de «FACTURA VUELATOUR» del Libro
  Dinero y de «factura vuelatour» de 'otros movimientos'. **Aquí no se
  decide nada**: se pinta tal cual. Es del VUELO (multi-avión: todas sus
  filas traen la misma; cancelados también). Validador LIBERAL
  `_factura_liberal`: número ⇒ texto («1234»; 1234.0 ⇒ «1234»), texto en
  blanco ⇒ None, booleano/objeto/NaN ⇒ None — nunca un 422 que tumbe el
  balance por una columna informativa. Layout: una entrada más en `_COLS`
  AL FINAL de STATUS DE COBROS («FACTURA\nVUELATOUR», formato None = texto)
  ⇒ la hoja pasa de 46 a **47 columnas** (AU); el título combinado del grupo
  crece solo (`AH1:AU1`), relleno verde del bloque, ancho 16 (`anchos`), fila
  TOTALES vacía (no está en `_TOTAL_MAP`) y guarda «=» (`_texto_literal`
  importado de `caja_chica_xlsx`; la misma guarda en la columna 10
  «factura vuelatour» de 'otros movimientos'). Por ir al final **no se
  corre nada**: `_COBRO1_COL` = 35 y `_COMISION_COL` = 29 siguen igual, y
  lo demás se busca por atributo. Nota al pie `_NOTA_FACTURA_VUELATOUR` justo después
  de la del ESTATUS DE COBRO (las notas de abajo bajan una fila). Sale en
  el libro INDIVIDUAL y en «reporte horas FLOTA» del GENERAL (los dos usan
  `_hoja_maestra`); 'cobranza' y el RESUMEN no cambian y 'otros
  movimientos' solo cambia con un folio que empiece con «=» (ahora texto). **Payload viejo (sin la clave) ⇒ libro idéntico salvo la columna
  vacía**: verificado celda a celda (valor, formato, relleno, fuente, borde,
  comentario, merges, anchos, alturas) contra el código previo en las dos
  hojas maestras — difieren solo AU2..AU(n), el merge del grupo y la nota
  nueva; las demás pestañas de ambos libros, idénticas. Tests:
  `tests/test_balance_factura_vuelatour.py` (layout e índices, valores tal
  cual, vacío sin clave/null, TOTALES vacía, columnas 1..46 intactas con y
  sin factura, multi-avión, «=A-12» como texto en la maestra y en 'otros
  movimientos', nota al pie, validador, rutas con el JSON del API).
- **EXTENSIÓN DE HORARIO pagada = traslado, como el TUA** (1-oct-2026, API
  0.0.47). Pedido de Ale sobre el #192 (N4142R, CUN-CTM-CUN): «me se está
  poniendo la extensión de servicios como Operación y no va en ese
  apartado» — factura de Chetumal «AE-Extension y/o antelacion de horario»
  $4,549.06 (mismo caso #190 XB-PEV $4,549.04; sin IA, #314 «extensión de
  servicio inspector» $500). **La regla y los montos viven en el API**
  (`desgloseGastoPartes`/`partesDeGasto`): la saca de OPERACIONES/OTROS y
  arma la nota de la celda («Extensión de horario (IVA incluido) $X**» en
  `op_detalle`). Aquí llegan dos llaves ADITIVAS que el API SOLO manda con
  monto ≠ 0: `BalanceAvionVuelo.extension_pagada_mxn` (informativa, no se
  pinta en ninguna columna) y `BalanceAvionTotales.extension_pagada_mxn`
  (Σ; también en el `consolidado` del general). Con ellas, en
  `_hoja_maestra` (libro individual y «reporte horas FLOTA»): (1) renglón
  «Extensión de horario pagada del periodo (solo nota en OPERACIONES, no
  resta en este libro):» justo después del del TUA pagado, mismo formato
  (A:H combinadas, monto en I, «MXN» en J), solo con el total ≠ 0; (2) pie
  ** ampliado `_NOTA_TRASLADOS_EXTENSION` y nota del combustible «(ni el
  TUA pagado ni la extensión de horario restan, ver **)» … «(sin gas, TUA
  ni extensión de horario)» cuando ALGUNA fila o los totales traen la llave
  ≠ 0 (`_hay_extension_pagada`; None o 0 cuentan como ausente). 'otros
  movimientos' y el Libro Dinero NO cambian: el egreso «extensión de
  horario pagada» llega armado en `concepto_egreso`/`nota_egreso`. **Sin la
  llave ⇒ libro byte-idéntico** (todos los miembros del .xlsx salvo
  `docProps/core.xml`, también con la llave en null/0), congelado con
  `_firma_hoja` de las dos hojas maestras calculada con el código previo
  (HEAD ef4d1f4). Las notas de 'Gastos Indirectos'/'repartidos a aviones'
  («El TUA pagado NO va aquí») no se tocaron: la extensión solo existe en
  gastos CON vuelo. Tests: `tests/test_balance_extension_horario.py`
  (montos reales, orden tras el TUA, sin TUA, columnas idénticas con y sin
  llave, pie en individual y general, skew fila/totales, rutas).
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
- **Estado de cuenta PDF por BLOQUES de páginas** (29-sep-2026). El
  Scotiabank mensual (1.18 MB, impreso del portal) daba 422 «demasiados
  movimientos»: el PDF completo iba en UNA llamada y la respuesta se
  truncaba a 16k tokens; aunque cupiera, generarla rebasaba los topes de la
  cadena (240 s por llamada aquí, el API aborta a 270 s, Vercel Hobby mata a
  300 s). `_parse_pdf` ahora abre el PDF con pypdf (`_abrir_pdf`) y lo parte
  en bloques de N páginas consecutivas (`estado_cuenta_paginas_por_bloque`,
  env `ESTADO_CUENTA_PAGINAS_POR_BLOQUE`, default 3) escritos con
  `PdfWriter`; los bloques se piden EN PARALELO
  (`ThreadPoolExecutor(max_workers=4)`, mismo `with_options(timeout=240.0)`,
  mismo modelo, `max_tokens` y `system`) y el resultado se ordena por
  página, no por llegada. Invariantes congelados en
  `tests/test_estado_cuenta_pdf_bloques.py`:
  (1) **≤ N páginas = UNA llamada BYTE-IDÉNTICA a la de siempre** (documento
  ORIGINAL + `_prompt_pdf`; `_PlanPdf.contenido` devuelve
  `_contenido_completo` cuando el bloque es el PDF entero); pypdf ausente,
  PDF corrupto, cifrado con contraseña de usuario o un bloque inicial que
  no se puede escribir ⇒ el camino de siempre (`_leer_completo`) + warning
  en el log `estado_cuenta`; (2) cada bloque lleva el prompt de siempre +
  «Este bloque contiene las páginas X–Y de N; transcribe SOLO…; las claves
  del resumen … déjalas en null…», y el que NO trae la página 1 recibe
  además el CONTEXTO DEL ENCABEZADO (primeros 1,500 caracteres de
  `extract_text()` de la página 1, para el año/periodo/cuenta/moneda) con la
  orden de NO transcribir movimientos de ese texto; (3) `stop_reason ==
  max_tokens` en un bloque ⇒ se re-parte a la mitad (recursivo; el primer
  pedazo lleva la página extra) y el consumo de la llamada truncada SÍ se
  suma; una página sola truncada ⇒ el error de siempre, mismo texto; (4)
  cualquier excepción de un bloque (Claude, JSON ilegible) se propaga tal
  cual y tumba TODA la lectura — jamás se importan parciales — y
  `pool.shutdown(wait=False, cancel_futures=True)` no espera a los bloques
  en vuelo; (5) tope TOTAL `_TOPE_TOTAL_PDF_S` = 260 s (el API aborta a
  270): pasado, 422 «tardó más de 4 minutos» en vez de un corte a ciegas;
  (6) fusión (`_fusionar`, también con 1 bloque): movimientos concatenados
  SIN deduplicar (dos cargos iguales el mismo día son legítimos), resumen =
  primer valor no nulo EN ORDEN de bloque, `advertencias` de la IA
  concatenadas y `_advertencias_pdf` sobre el total (la cadena de saldos y
  los totales impresos cruzan bloques); `uso_ia` = SUMA de todas las
  llamadas (`_sumar_usos`); `notas` dice «PDF de P páginas leído en K
  bloques.» cuando K > 1. Si un estado de cuenta futuro sigue truncando con
  3 páginas, bajar N por env antes de tocar código.
  Revisión adversaria (29-sep-2026), congelada en el mismo archivo: (a) el
  tope de (5) aplica también a un PDF de ≤ N páginas que pypdf sí abre (va
  por el mismo orquestador con un solo bloque); solo `_leer_completo` no lo
  tiene — la llamada y el resultado exitoso no cambian, solo el error a
  los 260 s; (b) el encabezado es `_PlanPdf.encabezado` PEREZOSO
  (`cached_property`): un PDF de ≤ N páginas que no se trunca ni lo extrae
  y un PDF grande lo extrae una vez; (c) si el periodo quedó DESPUÉS de los
  1,500 caracteres (un PDF impreso del portal abre con el menú del sitio),
  `_texto_encabezado` rescata hasta 3 renglones con «periodo»/«fecha de
  corte» — sin él los bloques 2+ no saben el AÑO de «07 SEP» — y nunca un
  renglón de movimiento; (d) la nota de páginas aclara que el documento
  adjunto trae SOLO esas páginas y que eso NO va en `advertencias` (si no,
  cada bloque avisaba «faltan páginas») y el contexto repite «transcribe
  únicamente los renglones del documento PDF adjunto»; (e) una lectura por
  bloques que falla deja en el log `estado_cuenta` el consumo YA cobrado de
  los bloques terminados (el API solo registra `ia_uso` con un 200); (f) un
  PDF cifrado solo con contraseña de DUEÑO (lo usual en bancos) se abre con
  "" y se parte normal (AES-256 probado). Riesgos abiertos: 4 llamadas
  simultáneas de `max_tokens` 16384 pueden toparse con el rate limit de
  salida de la cuenta (el SDK reintenta 429 dos veces; si no alcanza, 502
  «La IA está saturada (límite de peticiones): espera un minuto y
  reintenta», ver «Errores de Claude» en «IA»); con > 12 páginas hacen falta 2 tandas y
  un bloque truncado se re-lee completo antes de partirse, así que el tope
  de 260 s puede cortar; los hilos en vuelo siguen (y cobran) tras un error
  o el tope, y el SDK reintenta también los timeouts.
- **Conciliación de INGRESOS con IA** (24-sep-2026, API 0.0.34, pedido «con
  IA marcar los que sí empatan con los cobros de los vuelos»):
  `POST /conciliacion/sugerir-abonos` (`X-Internal-Token`), servicio
  `app/services/conciliacion_abonos.py` con su PROPIO `_client()`
  (`claude_fake(conciliacion_abonos, …)` lo parchea), esquemas
  `AbonoParaSugerir` / `CandidatoParaAbono` /
  `ConciliacionSugerirAbonos{Request,Response}` / `SugerenciaAbono` en
  `schemas/conciliacion.py` (todo con default; `folio` `str | int` → texto,
  banderas `null` → False, `neto` null ⇒ monto − comisión). El API manda
  lotes de ≤ 10 abonos (`max_length`, 11 ⇒ 422) con sus `candidato_ids`
  PERMITIDOS (prefijo `COBRO_VUELO:` / `SOBRE_GRUPO:` / `INGRESO:`) y un pool
  de ≤ 120 candidatos; aquí UNA llamada (`max_tokens` 6000, timeout 120 s
  con `with_options`, `sistema_con_dominio`). **La IA propone, nada liga**:
  el API revalida ids, calcula `monto_exacto` y registra `ia_uso`
  (`CONCILIACION_ABONOS_SUGERIR`) con el `uso_ia` de la respuesta.
  Errores: `stop_reason == max_tokens` ⇒ **502 `{detail: {error/code:
  IA_RESPUESTA_TRUNCADA, message, uso_ia}}`** (jamás se «repara» un JSON
  cortado; los créditos se gastaron y el API los registra); JSON ilegible ⇒
  422 con `error: IA_RESPUESTA_ILEGIBLE` + `uso_ia`; `APIStatusError` o
  conexión/timeout ⇒ 502 (nunca 500). Sin abonos ⇒ 200 sin llamar a Claude.
  El payload del modelo va SIN campos vacíos, con los candidatos de OTRA
  moneda ya fuera y con `exactos` por abono (ids que cuadran al centavo:
  la aritmética la hace Python, no el modelo). **Post-validación
  determinista** (`_evidencias_abono`, el tope final es el MÍNIMO):
  monto — el abono contra el NETO y el BRUTO del candidato y, en pasarela,
  `monto_bruto` contra el bruto — exacto (≤ 0.01) 0.95 · ≤ $1.00 y ≤ 1 %
  0.85 · ≤ 1 % 0.7 «posible comisión» · ≤ 5 % 0.5 · si no 0.3; monto NO
  exacto y SIN el nombre del cliente en la descripción ⇒ 0.4 (nombre =
  ≥ 2 tokens de ≥ 3 letras en común tras `normalizar_texto_banco`, sin
  «PAGO/VUELO/SPEI…»; es evidencia, no sube el tope); días (día de pared
  Cancún, `dias_entre_cancun`) ≤ 3 libre · 4–15 0.8 · 16–45 0.6 · > 45 0.4;
  `otra_cuenta` 0.5. `confianza = min(modelo, tope)`, también en las
  alternativas. Candidato fuera de SU lista, inventado o de otra moneda ⇒
  descartado (REVISAR, confianza 0); mismo candidato en dos abonos ⇒ gana
  la confianza mayor y el otro pasa a REVISAR («Ese candidato se propuso
  para otro abono con más confianza»); LIGAR sin candidato ⇒ REVISAR;
  candidato + acción que no es LIGAR/REVISAR ⇒ REVISAR; acción desconocida
  ⇒ REVISAR; `categoria_sugerida` fuera de `categorias` ⇒ None. **Anti
  doble conteo** (caso real: abono «MARIA CRISTINA CHAVEZ BADIOLA» 19,380 =
  cobro del vuelo #235, 20,400 − 1,020): REGISTRAR_INGRESO con un
  COBRO_VUELO/SOBRE_GRUPO permitido EXACTO ⇒ REVISAR con «Hay un cobro de
  vuelo con el monto exacto: revísalo antes de registrarlo como otro
  ingreso» y ese cobro PRIMERO en `alternativas` (≤ 0.7); lo mismo con un
  INGRESO ya registrado exacto («… vincúlalo en lugar de registrar otro»);
  REGISTRAR_INGRESO como ANTICIPO_CLIENTE ⇒ REVISAR (quién es el cliente y
  si su vuelo ya existe lo decide la persona). Evidencias ≤ 4 (las duras
  primero), alternativas ≤ 2, textos ≤ 300; una sugerencia por abono en el
  orden del request (sin respuesta del modelo ⇒ REVISAR); `advertencias`
  cuenta las correcciones. `_dominio.py` (v2026-09-24) ganó los alias
  «POCKET DE LATINOAMERICA» (BillPocket agrupado), «REV» (reverso) y «SPEI»
  (trae el nombre del ordenante). Textos de hojas: la nota de «otros
  movimientos» del balance, `_NOTA_GASTOS_EMPRESA` y las de «Otros
  ingresos»/«utilidades» del Libro Dinero ya dicen que las filas `ING-n`
  (otros ingresos sin vuelo, registrados en Ingresos) viven ahí y que los
  anticipos/aportaciones NO — sin cambio de schema: el API las manda como
  filas más (`filas_sueltas` / `otros_ingresos`). Tests:
  `tests/test_sugerir_abonos.py`.
  Revisión adversaria (24-sep-2026), trampas congeladas en tests: (1)
  **`max_retries=0`** en el `with_options` (`MAX_REINTENTOS`): el cliente
  compartido reintenta 2 veces y el SDK reintenta también los TIMEOUTS ⇒ una
  generación lenta duraba 3 × 120 s, el API ya había abortado a los 130 s y
  cada intento se cobraba sin llegar a `ia_uso`; (2) el JSON se lee con
  `raw_decode` desde el primer valor completo — antes una nota después del
  JSON con «}» o un «[» en la prosa de antes tiraban una respuesta buena como
  422; un objeto roto sigue siendo 422 (jamás se rescatan pedazos: una lista
  suelta solo vale si sus objetos traen `movimiento_id`); (3) `num()` de
  `validaciones_ia` rechaza NaN/Infinity (`json.loads` los acepta y
  `"confianza": NaN` tronaba el `le=1` del esquema ⇒ 422 SIN `uso_ia`) y
  TODO fallo al post-validar sale como `IA_RESPUESTA_ILEGIBLE` CON `uso_ia`;
  `stop_reason` `model_context_window_exceeded` también es truncado; el
  resto de `anthropic.APIError` ⇒ 502; (4) ids del modelo con otra
  capitalización o sin el prefijo `TIPO:` se resuelven al id CANÓNICO de la
  lista del abono solo si el uuid es de exactamente un candidato
  (`_resolver_id`; un pedazo de uuid sigue descartado), igual el
  `movimiento_id`; (5) **el candidato que cuadra al centavo SIEMPRE queda a
  la vista**: si el modelo no lo eligió (lo manda a «anticipo», clasifica,
  elige otro no exacto o ni contesta por ese abono) va PRIMERO en
  `alternativas` (≤ 0.7, o la confianza más baja del modelo si ya lo
  listaba); (6) la evidencia «otra cuenta» va antes que la referencia: el
  recorte a 4 nunca tira la que explica el tope.

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

- Visión (tacómetro/tickets) usa el modelo de `modelo_actual()` (ver la
  entrada siguiente; sin elección en el panel = `ANTHROPIC_MODEL`). La
  lectura de tacómetro recibe `ultimo` (último taco del avión) como ancla de
  magnitud; conservar ese contrato.
- **Modelo de Claude elegible desde Configuración** (2-oct-2026, API
  0.0.51). Pedido: «dejar una opción en la configuración para adaptar el
  modelo que quieran utilizar, aunque ahorita dejaremos por default el que
  estamos usando». La elección vive en el API (`configuracion_sistema`,
  clave `ia_modelo`); aquí NO se guarda nada. El API manda el id en el
  header **`X-IA-Modelo`** SOLO cuando hay uno configurado.
  `require_internal_token` (`app/security.py`) lo deja en el ContextVar
  `modelo_ia_peticion` de `app/services/modelo_ia.py` en TODA petición
  (sin header ⇒ None: jamás queda el de la anterior) y **`modelo_actual()`
  es la ÚNICA fuente** del `model=` de cada `messages.create` y del campo
  `modelo` de las respuestas (`uso_ia` sigue saliendo de `resp.model`).
  Header ausente, vacío o que no cumple `^claude-[a-z0-9.-]{3,80}$`
  (`PATRON_ID_MODELO`, la MISMA regex del API y del panel; `fullmatch`,
  espacios a los lados se recortan) ⇒ `ANTHROPIC_MODEL` + warning en el log
  `security` (solo con token válido). API viejo sin header ⇒ todo como
  antes. Aquí no hay catálogo ni tarifas: un id válido que Anthropic no
  conoce da 404 y `TEXTO_MODELO_INEXISTENTE` dice dónde se cambia. Dos
  trampas congeladas en `tests/test_modelo_ia.py`: (1) la dependencia es
  **`async def` a propósito** — FastAPI corre una dependencia `def` en un
  hilo con COPIA del contexto y el ContextVar fijado ahí se pierde (todo
  usaría el del servidor); (2) **un `ThreadPoolExecutor` NO hereda el
  ContextVar**: quien reparta llamadas en hilos propios resuelve
  `modelo_actual()` ANTES y pasa el id (`estado_cuenta._leer_por_bloques` →
  `_llamar_claude(cliente, contenido, modelo)`; test en
  `tests/test_estado_cuenta_pdf_bloques.py`). Las pruebas de AISLAMIENTO
  entre peticiones van con `httpx.ASGITransport` (seguidas en la MISMA
  tarea; simultáneas con `asyncio.gather` en un loop), nunca con
  `TestClient` sin `with`: abre hilo, loop y contexto nuevos por petición y
  una fuga del modelo jamás se vería. Ruta `GET /ia/modelo`
  (internal token) → `{default_servidor, efectivo}`: el API la consulta
  best-effort para pintar «default del servidor» en Configuración.
  Llamada nueva a Claude ⇒ `model=modelo_actual()`, nunca
  `get_settings().anthropic_model`.
- Todo punto de IA degrada a captura manual: los errores devuelven
  `legible=false`/`disponible=false`, nunca 500 por fallo del modelo.
- **Errores de Claude → texto para el operador** (1-oct-2026). Caso real:
  un ADMIN vio «IA no disponible: Claude no disponible (400). Captura a
  mano.» y la oficina creyó que «Claude no está disponible»; era el SALDO de
  créditos agotado, que Anthropic responde como **400
  `invalid_request_error`** («Your credit balance is too low…»). Fuente
  ÚNICA `app/services/ia_errores.py` (pura salvo el log):
  `detalle_error_claude(e)` decide primero por el TEXTO del error
  (`error.message` del cuerpo; sin él, `e.message`; sin distinguir
  mayúsculas, la primera regla de `_REGLAS_POR_TEXTO` gana: saldo, límite
  de gasto «usage limits», pixeles «image dimensions», tipo mal etiquetado
  «image appears to be», peso, formato, foto ilegible, demasiadas fotos,
  documento largo «prompt is too long»/«context limit») y después por el
  STATUS (401/403 llave, 404 + «model», 413 peso, 429 saturada, TODO 5xx
  caída); lo demás cae en «Claude no disponible (<status>): <mensaje>» sin
  el prefijo «Error code: N - » del SDK y jamás un dict/lista/HTML (queda
  «Claude no disponible (400)»). **Contrato de los textos**: SIN punto final
  y SIN «captura a mano» — quien los muestra arma «IA no disponible:
  <texto>. Captura a mano…» (app: `gasto_screen`, `admin_factura_screen`,
  `inventory_item_form_screen`); con punto salía «..». La constancia
  fiscal (degrada a 200, no es 502) usa
  `detalle_error_claude_captura_manual` («…; captura los datos
  manualmente»). `registrar_error_claude(logger, e)` deja status, tipo,
  mensaje (≤ 500) y `request_id`; nunca la llave ni encabezados. **Todo
  `except anthropic.APIStatusError` nuevo en `app/routers` usa los dos**:
  `test_todos_los_routers_usan_la_fuente_unica` (AST) lo exige y cuenta
  EXACTAMENTE 11 manejadores — agregar uno obliga a subir ese número.
  Además el bloque `image` ya no confía en el `media_type` del cliente:
  `app/services/imagen_media_type.media_type_real` lo deduce por bytes
  mágicos (JPEG/PNG/GIF/WEBP; si no reconoce, respeta el declarado) en
  `anthropic_vision._image_block`, `_constancia_blocks` y
  `vencimiento_extract._source_block` — una captura PNG etiquetada JPEG
  daba 400. Pendiente en el API (no aquí): solo `readGastoTicket`,
  `readConstanciaFiscal` e inventario reenvían el `detail`; tacómetro,
  combustible, vencimientos y las sugerencias de conciliación lo tiran, y
  `/conciliacion/parse` y compras lo muestran envuelto en el JSON. Tests:
  `tests/test_ia_errores.py` (clases reales del SDK; los casos clave, armados
  por el cliente REAL sobre `httpx.MockTransport`).
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
  incompleta que no truncó no se detecta con `stop_reason`). Desde el
  29-sep-2026 un PDF de más de N páginas se lee por bloques (ver
  «Conciliación»): el `_SYSTEM` es el mismo; la nota de páginas y el
  contexto del encabezado viajan en el mensaje del usuario.
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
