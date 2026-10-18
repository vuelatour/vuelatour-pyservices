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
  Sin token solo `/`, `/health` y `GET /version` (14-oct-2026, API 0.0.77:
  `{commit, servicio}`, con `commit` = `RAILWAY_GIT_COMMIT_SHA` del
  deployment o null fuera de Railway): es la COMPUERTA verificable con curl
  de los despliegues «pyservices antes del API» (invariante 49 del API).
- **Orden de despliegue del lote 0.0.77 (14-oct-2026; decisiones del cliente
  del 14-oct: anticipo devuelto sin CFDI y conversión ÚNICA del ingreso de un
  avión)**: pyservices (banderas `usd_por_ingreso` y
  `factura_vuelatour_devolucion`; sin ellas, libros byte a byte) ⇒
  **COMPUERTA**: `curl …/version` (o el deployment ACTIVO en Railway) = el
  commit nuevo ⇒ API 0.0.77 (`GET /v1/version`) ⇒ panel. Al revés no hay
  error que avise: el API nuevo con un pyservices viejo publica la hoja
  «Otros ingresos» con los números buenos (la C5 vieja no reproduce y
  `finalizar` la deja como VALOR) pero con B5 = TC PROMEDIO y TRES notas
  falsas (`_NOTA_OTROS_INGRESOS_AVION`, `_NOTA_CASCADA_OTROS_INGRESOS`,
  `_NOTA_OTROS_INGRESOS_AVIONES_GENERAL`) y la nota de FACTURA VUELATOUR
  sin el anticipo devuelto. Rollback: PRIMERO el API a 0.0.76 y solo
  después pyservices (nunca pyservices solo con el 0.0.77 arriba).
- **Orden de despliegue del lote 0.0.75 (CONSOLIDADO en la revisión del
  14-oct-2026; manda sobre lo que diga cada entrada del lote)**: hasta la
  revisión este archivo pedía dos órdenes incompatibles para el MISMO push
  (Comisariato «DESPUÉS del API 0.0.75»; ingreso por avión, ID por cliente y
  TUA «ANTES»). El orden ÚNICO, igual al del CLAUDE.md del API («ORDEN DE
  DESPLIEGUE DEL LOTE 0.0.75»): migraciones `20261013000001` y
  `20261013000009` (dependencia dura del API) → **pyservices COMPLETO** →
  API 0.0.75 → migraciones `20261013000005` y `20261013000002` → panel →
  APK. pyservices ANTES del API es seguro: sin las llaves nuevas todo libro
  sale byte a byte; la sugerencia COMISARIATO de la IA no la aplica nadie
  (la app publicada +80 y el panel de prod solo aplican sugerencias de SU
  catálogo y el API 0.0.74 nunca pisa la categoría) — el 400 IsEnum solo
  lo daría el chip de una app +81 publicada antes de tiempo, con o sin
  pyservices —; y un estado de cuenta leído con los renglones de
  continuación entra igual. Al revés (API primero) el libro de N990GG de
  septiembre saldría sin la pestaña «Otros ingresos» con una utilidad que
  ya los suma, y la CLAVE con ID por cliente cortada. En la ventana entre
  los dos pushes no re-subir estados de cuenta ni regenerar libros.

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
- **Hoja 'inventario': COSTO UNITARIO** (13-oct-2026, API 0.0.74; decisión
  6 del lote del cliente: ver el último precio de compra de cada producto
  en el Balance general). Campos ADITIVOS de fila `costo_unitario_mxn`
  (pesos: compra en MXN tal cual o USD × T.C. oficial de hoy, el mismo
  `tc_hoy` del valorizado) y `costo_unitario_usd` (SOLO si el último
  precio es USD sin T.C., el caso `sin_tc`; en dólares, jamás convertido
  aquí); nunca los dos en una fila. Sin bandera raíz: `_hay_costo_unitario`
  (alguna fila trae cualquiera de los dos) y `_hay_costo_unitario_usd`
  (alguna trae el USD) deciden, como `_hay_usd_sin_tc`. Columnas «COSTO
  UNITARIO MXN» (`MONEY`, ancho 14) y «COSTO UNITARIO USD (sin T.C.)»
  (`MONEY_USD`, 16) entre EXISTENCIA ACTUAL y VALOR A COSTO MXN: tercer
  desplazamiento `oc` (1 o 2) que se SUMA a `off` (así las columnas del
  periodo, `N + off`, no se tocan) e independiente de `off_venta`; con los
  tres a la vez son 14 columnas y `anchos = [34, 42, 14, 16, 16, 18, 12,
  14, 12, 14, 14, 14, 14, 24]`. Es un precio, no un importe: la fila
  TOTALES deja esas celdas EN BLANCO (sin valor ni fórmula; relleno y
  borde del bucle de siempre) y las `ROUND(SUM(...),2)` de los demás
  totales solo cambian de letra (nada nuevo en `ENCABEZADOS_SIN_ROUND`).
  Nota al pie `_NOTA_INVENTARIO_COSTO_UNITARIO` (+ `_USD`) SOLO en la rama
  `regla_costo='ULTIMO_PRECIO'`, tras las notas USD y ANTES de
  `_nota_margen` (los tests hacen `endswith` con la frase del margen):
  dice que existencia × COSTO UNITARIO puede diferir unos centavos del
  VALOR A COSTO (el API lo calcula con el precio y el T.C. completos:
  aceite 84 × 375.55 = 31,546.20 vs 31,546.13) — no se «arregla» aquí ni
  en el API. Payload viejo o campos en null ⇒ hoja byte a byte idéntica
  (huellas `_FIRMA_PAYLOAD_VIEJO` / `_FIRMA_USD_SIN_TC_22SEP` intactas);
  la vista en pantalla (`libro_json`) hereda la columna sola. Tests:
  `tests/test_balance_inventario_xlsx.py` (sección «COSTO UNITARIO»),
  `tests/test_balance_formulas.py` (libros «inventario general · costo
  unitario» y «… USD»).
- **Hoja 'permisos' con T.C. por fila y MONTO MXN como fórmula**
  (13-oct-2026, API 0.0.74; decisión 1 del lote del cliente: ver con qué
  tipo de cambio se convirtió cada gasto a pesos y que MONTO MXN sea una
  fórmula visible, no un número suelto). Campos ADITIVOS:
  `BalanceAvionGastoFila.tc_aplicado: float | None` (el T.C. con el que el
  API convirtió ESA fila —su `tc_gasto` o el promedio CRUDO del periodo—,
  el número exacto que multiplicó; None en MXN, sin T.C. o API previo) y la
  bandera `BalanceAvionHojaGastos.tc_por_fila: bool | None` (true SOLO en
  `permisos`). `_hoja_gastos` gana el kwarg `con_tc` y SOLO la llamada de la
  pestaña 'permisos' de `construir_balance_avion` lo pasa
  (`bool(req.permisos.tc_por_fila)`) junto con la nota `_NOTA_PERMISOS_TC`.
  Con él: columna «T.C.» AL FINAL (`col_tc = n_cols`, letra con
  `get_column_letter`, nunca «G» a mano; ancho 12; VALOR con formato `TC`;
  comentario `_NOTA_ENCABEZADO_TC_PERMISOS` en el encabezado) y, por fila
  (`_fila_monto_con_tc`): gasto en PESOS (moneda y monto original null, el
  contrato de siempre) ⇒ E = «MXN», F = `monto_mxn` (COPIA del API, jamás
  una conversión), T.C. vacío y D = `=ROUND(F{r},2)`; otra moneda con T.C.
  ⇒ E/F como siempre, G = `tc_aplicado` y D = `=ROUND(F{r}*G{r},2)` con el
  `monto_mxn` del API en caché; otra moneda sin T.C. (`monto_mxn` None) ⇒ D
  vacía sin fórmula y G vacía, como hoy; `monto_mxn` sin `tc_aplicado` ⇒
  valor del API. Todas con ROUND exterior (nada nuevo en
  `ENCABEZADOS_SIN_ROUND`) y verificadas por `xlsx_formulas` (una que no
  reproduzca el centavo vuelve a valor: los números no cambian). A5 sigue
  `ROUND(SUM(D8:Dn),2)`; B5..E5, la cascada `'permisos'!$C$5` y el bloque
  AFAC de 'otros movimientos' no cambian. **Sin bandera ⇒ hoja idéntica**
  (API ≤ 0.0.73, o `tc_aplicado` sin la bandera: huella `_firma_hoja` de
  'permisos' y `_FIRMAS_ANTES["tst"]` intactas) y las demás hojas de gastos
  (indirectos, refacciones, otros gastos, repartidos a aviones, fallback
  del general) nunca reciben `con_tc`: byte a byte. La vista en pantalla
  (`libro_json`) hereda la columna sola. Tests:
  `tests/test_balance_formulas.py` (sección 3b: columna y fórmulas sin
  degradar, USD sin T.C., sin bandera idéntica, las demás hojas intactas) y
  `tests/test_libro_json.py`.
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
  del RESUMEN y nota de la hoja maestra del GENERAL —esta última, solo hasta
  el 7-oct-2026: ver «COMISIONES solo en el libro INDIVIDUAL»—;
  `_hay_pago_vendedor_real` exige `is True`) y
  `DineroXlsxRequest.utilidades_comision_vendedor_pagada_mxn:
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
  rutas con el payload nuevo). Desde el 12-oct-2026 (API 0.0.73) la bandera
  `sin_provision_vendedor` MANDA sobre esta: ver «SIN PROVISIÓN del pago al
  vendedor en 'otros movimientos'» más abajo. **SUPERSEDIDO el 13-oct-2026
  (API 0.0.74)**: ya no hay provisión que reemplazar en ningún libro y la
  categoría se partió en dos, «Comisión del vendedor (VuelaTour)» y
  «Comisión del vendedor (avión)» (ver «COMISIÓN DEL VENDEDOR SOLO COMO GASTO
  REAL, EN DOS CATEGORÍAS»).
- **FACTURA VUELATOUR en la hoja maestra de los balances** (30-sep-2026, API
  0.0.45). Pedido de Marie para Ale, con la foto del «reporte horas XA-VGV»:
  «se ocupa que diga el num de factura que nosotros emitimos del servicio …
  y si puede salir en el reporte general también». Campo ADITIVO
  `BalanceAvionVuelo.factura_vuelatour: str | None` que el API resuelve con
  su cascada ÚNICA `etiquetasFacturaDeVuelos` (CFDI timbrado vivo →
  facturas EMITIDAS vigentes «A-0424, A-0430» → `vuelo.factura_folio`
  tecleado → etiqueta del estatus «Facturado» / «Factura elaborada y
  enviada» → desde el 13-oct-2026 (API 0.0.74) «Pendiente de factura» /
  «No lleva factura» según los cobros del vuelo → vacía = sin cobros; ver
  «FACTURA VUELATOUR según los cobros»): la MISMA etiqueta de «FACTURA
  VUELATOUR» del Libro Dinero y de «factura vuelatour» de 'otros
  movimientos'. **Aquí no se decide nada**: se pinta tal cual. Es del VUELO (multi-avión: todas sus
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
  corre nada**: COBRO 1 = 35 y COMISIÓN VENDEDOR = 29 siguen igual (hoy
  `_DISP_MENSUAL.cobro1_col` / `.comision_col`), y lo demás se busca por
  atributo. Nota al pie `_NOTA_FACTURA_VUELATOUR` justo después
  de la del ESTATUS DE COBRO (las notas de abajo bajan una fila); desde el
  13-oct-2026 la nota depende de la bandera raíz
  `factura_vuelatour_segun_cobros` (`_nota_factura_vuelatour`). Sale en
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
- **Balances con FÓRMULAS visibles + caché** (5-oct-2026, pedido: «los
  permisos de AFAC… que tenga las fórmulas visibles»). Toda celda que es
  aritmética de celdas del libro va como fórmula que reproduce EXACTO el
  número del API: `ROUND(…,2)` solo donde el API hace `round2` (fila de
  vuelo L/M/N/X/Y/AI/AJ/AM, por cobrar, Σ de TOTALES y hojas) y sin ROUND
  donde no (AE/AF/AG/AH/AN/AO). Constantes en la fila 1 de la maestra
  (`$B$1` Permiso AFAC USD/hr del request, `$D$1` factor IVA 1.16); el TC
  promedio CRUDO del API es `AVERAGE(V filas)` (la celda TOTALES va
  redondeada: no se divide entre ella). Lo que depende de datos fuera del
  libro (IVA % del cliente, prorrateos multi-avión, TC de cada gasto, tarifa
  AFAC y promedios del GENERAL) sigue como valor. Helper `_formula(ws, …,
  formula, valor_api, fmt)`: `valor_api` None ⇒ celda vacía sin fórmula.
  `app/services/xlsx_formulas.py` (`guardar(wb)` en los dos renders)
  EVALÚA cada fórmula como Excel (ROUND medio-hacia-afuera a 15 cifras) y,
  si no se ve igual que el número del API, la DEGRADA a valor (redondeos
  intermedios del API, empates binarios); luego inyecta el número del API
  en el `<v>` de cada fórmula (la vista previa del teléfono no calcula) con
  `fullCalcOnLoad`. Leído con `data_only=True` el libro es idéntico al de
  antes (huellas congeladas del generador previo): el `<v>` se escribe con
  «%.16g» como `safe_string` de openpyxl, no con `repr` (17 cifras cambiaban
  el último bit del % COBRADO). Si la verificación revienta por algo que no
  es la fórmula, esa celda —o el libro entero, en `guardar`— sale con
  valores: jamás un 500. Tests:
  `tests/test_balance_formulas.py` (espejo de la aritmética del API +
  evaluador independiente `tests/_evaluador_formulas.py`, cuyo
  `verificar_libro` exige ROUND exterior en TODA fórmula salvo referencias
  puras y `ENCABEZADOS_SIN_ROUND` — una fórmula nueva sin ROUND se agrega
  ahí a propósito); los tests que leen números de un balance usan
  `data_only=True`.
- **Bloque «VUELATOUR (empresa)» al final de la hoja 'balance' del GENERAL**
  (6-oct-2026, API 0.0.59; pedido: «en la hoja de balance falta, hasta el
  final, el balance de la empresa VuelaTour»). Payload ADITIVO
  `BalanceGeneralRequest.empresa` (`BalanceEmpresaBloque`, todo opcional;
  `participaciones` null ⇒ []) + `BalanceAvionSocio.es_empresa`; los números
  los calcula el API (`balance-empresa.util.ts`), aquí solo se pintan.
  `_bloque_balance` acumula en `celdas_empresa` la celda MONTO USD real de
  cada socio `es_empresa` (los bloques tienen distinto número de socios:
  jamás calcular la fila) y `_bloque_empresa` la cita: PARTICIPACIÓN =
  `ROUND(SUM(C17,C33),2)` con una línea gris `=C17` por avión (sin socio
  empresa: 0.00 + «sin participación registrada»; `participaciones` sin la
  bandera en los socios ⇒ valores), OTROS GASTOS = `'otros gastos'!$C$5`
  (sin esa hoja, valor), RESULTADO = `ROUND(part+ing−pag−otros+tienda,2)`
  verde/rojo; ingresos propios, egresos propios y tienda van como VALOR
  con nota (salen de MXN con el TC de cada vuelo). Todo por `_formula`
  (verificado y en caché). **Las dos filas de 'otros movimientos' NO son
  caja** (revisión 6-oct): el ingreso es lo COTIZADO de los vuelos del
  periodo no cancelados (COTIZADO/RESERVA incluidos) y `pagos_vendedor_usd`
  es TODO lo que esa hoja resta (pago o provisión al vendedor, TUAs,
  extensión, comisión bancaria, sueltos) — etiquetas «INGRESOS PROPIOS USD
  (… cotizado del periodo)» / «EGRESOS PROPIOS USD (…)» y
  `_NOTA_EMPRESA_BASE` al pie, debajo de la `nota` del API (sin ella, el pie
  fijo `_NOTA_BLOQUE_EMPRESA`; jamás los dos). Base de caja = cambio del API
  y decisión del cliente. `participacion_usd` null ⇒ celda vacía; con
  `empresa` el RESUMEN suma una línea que apunta al bloque. **Sin `empresa`
  ⇒ libro byte-idéntico** (también con `es_empresa` en los socios o
  `empresa: null`; verificado miembro a miembro contra HEAD f509d1a y
  congelado con huellas de layout). Tests:
  `tests/test_balance_empresa_vuelatour.py` (payload con la forma REAL del
  API y espía del registro de `xlsx_formulas`: una fórmula degradada deja
  el mismo número y solo el registro la delata).
- **Nota «cómo se cobró» en cada parcialidad del balance** (6-oct-2026,
  API 0.0.60; pedido: «al lado de la columna STATUS, si ya se pagó, que
  venga cómo se cobró, quién lo cobró y a qué cuenta», y luego «mejor,
  después de cada cobro venga la nota… para no tener tantas columnas
  nuevas»). **SIN columnas nuevas**: en `_hoja_maestra` (libro individual y
  «reporte horas FLOTA») cada celda COBRO n MXN lleva una NOTA (autor
  «VuelaTour», 320 de ancho, alto ≥ 70 según renglones —
  `_comentario_cobro`) «15/09/2026 · $12,522.20 · Transferencia →
  Scotiabank Pesos · Itzi» y la celda STATUS (mismo texto) una nota resumen
  con TODOS los cobros, uno por renglón. Payload ADITIVO en
  `BalanceAvionCobro`: `cobrado_con` (texto YA armado por el API, fuente
  única `etiquetaCobradoCon`; se pinta tal cual), `metodo_etiqueta`,
  `registro` (validador liberal: blanco o no-texto ⇒ None). Helpers puros:
  `_nota_cobro(c)` = fecha · monto · `_cobrado_con(c)`; sin `cobrado_con`
  (API ≤ 0.0.59) se compone `metodo_etiqueta`/`metodo` → `cuenta` (+
  «Registró: X» si llegara `registro`); **sin método, cuenta ni registro ⇒
  sin nota** (los payloads de prueba con solo fecha/monto no cambian).
  `_nota_resumen_cobros(cobros)` (STATUS y la 4.ª agregada: una línea por
  cobro, None si ninguno dice cómo se cobró); `_notas_parcialidades` sigue
  el orden de `_cobros_a_4` (con más de 4 la nota de la 4.ª lleva una línea
  por cobro agrupado). Nada se mueve: ni columnas, ni fórmulas, ni la hoja
  'cobranza'. **Libro sin cobros ⇒ byte-idéntico** (miembro a miembro contra
  HEAD 4125973, huellas congeladas); con cobros solo cambian
  `xl/comments*.xml` y el VML. `tests/test_balance_empresa_vuelatour.py`
  deja fuera de su huella SOLO esas notas (sus payloads traen método y
  cuenta). Tests: `tests/test_balance_cobrado_con.py` (helpers, 2 cobros =
  2 notas + resumen, sin cobros, API viejo, general, > 4 cobros, rutas).
- **Balance MENSUAL y Balance GENERAL con costo por hora** (6-oct-2026, API
  0.0.64; pedido: «el que ya tenemos que se renombre a "Balance mensual" y el
  nuevo botón sea el "Balance general"… se van a agregar estas columnas para
  sacar el costo por hora por cada vuelo», con la fórmula «el total de todos
  los gastos, entre el tiempo volado, entre el tipo de cambio del día, entre
  1.16»). Campo ADITIVO `BalanceGeneralRequest.variante` ('mensual' |
  'general', default 'mensual'; liberal: mayúsculas/espacios se normalizan,
  null o desconocido ⇒ mensual, jamás 422). **Mensual o sin la llave ⇒ libro
  BYTE-IDÉNTICO** al de antes de la variante (hasta el 7-oct-2026: desde
  entonces su hoja de vuelos va sin COMISIONES, `_COLS_FLOTA`, ver abajo)
  (verificado miembro a miembro contra HEAD 93486c1 en los 37
  libros de los tests del balance; congelado con la huella de layout
  `_FIRMAS_HOY` Y con los BYTES: `_MIEMBROS_HOY` = primeros 16 hex del sha256
  de cada miembro del .xlsx salvo `docProps/core.xml`, de 3 libros, calculados
  con el árbol de 93486c1 y openpyxl 3.1.5 — la huella de layout no ve el
  tamaño de las notas (VML), el color ni los otros lados de un borde, ni la
  fuente; si solo cambia la versión de openpyxl, se regeneran con ese árbol).
  La ruta nombra el archivo `balance-{variante}-vuelatour-…` (el API pone el
  suyo al descargar). **General ⇒ cambia «reporte horas FLOTA»** (fuera de
  ella solo las letras con que la citan y la nota de 'repartidos a aviones',
  que en la general dice que el TUA pagado queda en el desglose de TOTAL PARA
  PROVEEDOR (PESOS): `_CPH_NOTA_REPARTIDOS_A_AVIONES`; la de siempre
  nombraba OPERACIONES): `_COLS_GENERAL` (39 columnas desde el 8-oct-2026;
  44 del 6 al 8-oct y 45 con COMISIONES del 6 al 7-oct, ver las entradas
  siguientes) = CLAVE..ESTADO y VENTA iguales; «COSTO POR HORA» = las 12
  columnas de la hoja «utilidades» del cliente en su orden y con sus nombres
  (repite a propósito la venta s/IVA: el bloque se lee solo); STATUS DE
  COBROS y FACTURA iguales. Hasta el 8-oct llevaba también TIEMPO /
  TACÓMETRO y «COSTO TOTAL (MXN)» (COSTO TOTAL + TIPO CAMBIO COSTOS); ver la
  entrada «SIN COLUMNAS DUPLICADAS» de abajo. Fuera OPERACIONES/PILOTO/OTROS/AFAC —su desglose va en la
  NOTA de TOTAL PARA PROVEEDOR (PESOS): «Operación $x · Piloto $y · Otros $z
  · Permiso AFAC $w» + sus líneas de detalle (`_nota_desglose_costo`)— e
  INDICADORES. **Columnas por LLAVE, jamás letras fijas**: `_Disposicion`
  (letra, COBRO 1, total_map, rellenos… de cada juego); los mapas fijos del
  módulo (`_LETRA`, `_COBRO1_COL`, `_COMISION_COL`, `_COBRO_MXN_LETRAS`) se
  RETIRARON en la revisión del 6-oct-2026 —una fórmula nueva escrita con
  ellos apuntaba, en silencio, a las letras del mensual dentro de la
  general—: hoy todo sale de `_DISP_MENSUAL` / `_DISP_GENERAL` (los tests
  congelan las letras del mensual a mano, `_LETRAS_MENSUAL`). Guardia:
  `test_otro_juego_de_columnas_sigue_cuadrando` cambia el juego general
  (columna real antes de VENTA, OPERACIONES dentro del bloque, sin TIPO
  CAMBIO COSTOS) y exige el libro COMPLETO cuadrando: sin fórmulas
  degradadas, evaluador independiente, cada número por su llave, fórmulas
  del bloque y Σ de TOTALES con las letras de ese juego, y las citas de
  'combustible'/'otros gastos'/'repartidos a aviones' en la columna de la
  llave (o valor + nota «calculado por el sistema» si el juego no la tiene).
  Las columnas que repiten un dato son llaves `cph_*` (`_ORIGEN_GENERAL`,
  valor) y HEREDAN las señales de su columna de origen: TOTAL COBRADO S/IVA
  el azul del T.C. oficial y TIEMPO CALZOS/HOBS —divisor de COSTO X HORA— el
  ámbar y la nota del salto de taco interno. IVA X HR y TOTAL PAGADO S/IVA son
  la resta de dos campos del API (`_valor_columna`).
  Fórmulas = aritmética del API citando el propio bloque: COSTO X HORA
  `=PESOS/TC/TIEMPO/$D$1` (orden del API: Y/z/O/1.16), COSTO HR MÁS IVA
  `=DLLS/TIEMPO`, DLLS `=PESOS/TC`, IVA DLLS `=DLLS−DLLS/$D$1`, IVA PESOS
  `=IVA DLLS×TC`, PAGADO S/IVA `=PESOS−IVA PESOS` (sin ROUND, en
  `ENCABEZADOS_SIN_ROUND`), REMANENTE `=ROUND(VENTA AVIÓN MXN−PESOS,2)`; **IVA
  X HR `=ROUND(ROUND(W,2)-ROUND(U,2),2)`**: la resta de las celdas completas
  no daba la diferencia de los dos números a la vista en 1 de cada 4 filas
  (medido) y se degradaba a valor. COSTO TOTAL cita columnas que la variante
  no tiene ⇒ VALOR del API con nota «Calculado por el sistema…» en su
  encabezado (`_ColumnaAusente`); otra hoja que citara una columna ausente ⇒
  valor con nota (`_formula_cita`: jamás una referencia rota). TOTALES = los
  del API por llave (las repetidas llevan el de su dato; promedios del
  consolidado como valor) y, sin huecos (revisión 6-oct-2026: con huecos «la
  fila parece rota»), las del bloque que el API no totaliza
  (`_CPH_TOTALES_PROPIOS`, solo en la variante general): **TOTAL PARA
  PROVEEDOR (DLLS), IVA TOTAL PAGADO (DLLS y PESOS) y TOTAL PAGADO S/IVA = Σ
  de las filas TAL COMO SE VEN**, `=ROUND(SUMPRODUCT(ROUND(rango,2)),2)` con
  caché = round(Σ de los números del API de las filas) — sus celdas de fila
  son fórmulas SIN ROUND y `ROUND(SUM())` sumaría los valores completos (un
  centavo distinto en las cuatro del libro de pruebas; en un mes real casi
  siempre ⇒ se degradaría a valor); **COSTO X HORA `=ROUND(W tot/$D$1,2)`**
  (COSTO HR MÁS IVA de TOTALES = promedio del API, ÷ 1.16 como cada fila; NO
  `AVERAGEIF` de la columna: el consolidado promedia los promedios de cada
  avión, 117.47 vs 120.38 de las filas en el libro de pruebas) e **IVA X HR
  `=ROUND(W tot−U tot,2)`**; sin `costo_hr_prom_usd`, las dos vacías; sin la
  columna que citarían, valor con la nota «calculado por el sistema». Las
  fórmulas citan por llave (`_formula_total_costo_por_hora`) y el valor sale
  de `_valor_total_costo_por_hora`. `xlsx_formulas` y el evaluador de tests
  entienden SUMPRODUCT y ROUND de un rango SOLO dentro de él (fuera, Excel
  haría intersección implícita ⇒ aquí #¡VALOR! ⇒ valor). A1:D1 se conservan; el
  título del bloque «BALANCE GENERAL · COSTO POR HORA» dice en el encabezado
  cuál libro es; notas al pie, renglones TUA/extensión y notas de B1/D1
  propias (`_CPH_*`).
  Las demás hojas tienen los MISMOS números en las dos variantes; las citas a
  TC/horas de la maestra caen en SU columna. **PROVISIÓN en amarillo** (las
  dos variantes; pedido «los que están provisión podrían estar en
  amarillo»): en 'otros movimientos' la fila cuyo egreso trae la marca del
  API «· PROVISIÓN (» EN MAYÚSCULAS (`_MARCA_PROVISION`: «pago comisión
  vendedor (X) · PROVISIÓN (mismo monto…)») —en el concepto o, si la fila
  junta varios egresos del vuelo, en la línea de su nota— pinta concepto y
  monto del egreso con AMBER y al pie (TOTALES + 2, A:C) va «Amarillo =
  provisión (sin gasto real capturado)». El gasto real dice «reemplaza la
  provisión» en minúsculas y NO se pinta; tampoco un PROVEEDOR en mayúsculas
  de una fila suelta («Combustible · PROVISIONES AEREAS DEL SURESTE»: la
  primera versión buscaba la subcadena y lo pintaba). Sin provisiones la
  hoja es la de siempre. Desde el API 0.0.73 (12-oct-2026) la marca YA NO
  VIAJA: sin gasto real la fila no lleva egreso y no hay nada que pintar
  (ver «SIN PROVISIÓN del pago al vendedor en 'otros movimientos'»); la
  lógica del amarillo se conserva tal cual para un API previo. Tests:
  `tests/test_balance_general_costo_hora.py` (+ dos libros de la variante en
  `_LIBROS` de `tests/test_balance_formulas.py`).
- **NEGATIVOS EN ROJO con formato condicional** (8-oct-2026; pedido del
  cliente con la captura del RESUMEN: la GANANCIA −11,299.66 del N58BT en
  rojo). `_rojo_si_negativo(ws, *rangos)` agrega una REGLA de Excel (`cfRule`
  cellIs `lessThan 0`, relleno `FFC7CE` + letra `9C0006` = preset «Relleno
  rojo claro con texto rojo oscuro»), NO un relleno fijo: la evalúa Excel con
  el número que la celda muestre, también si la oficina edita la fórmula.
  Dónde: hoja maestra —`_COLUMNAS_NEGATIVO_ROJO` = REMANENTE VENTA−COSTO,
  GANANCIA MXN/USD y REMANENTE VENTA MENOS COMPRA (variante general), filas +
  TOTALES, las que tenga cada juego de columnas—, RESUMEN flota (GANANCIA
  MXN, aviones + TOTALES), hoja 'balance' (las tres filas UTILIDAD y el
  MONTO USD de cada socio; el RESULTADO del bloque VUELATOUR (empresa)),
  'otros movimientos' (REMANENTE, filas + TOTALES) y 'refacciones' (GANANCIA
  MXN). Los colores fijos que ya había (verde/rojo en UTILIDAD COBRADA y en
  el RESULTADO) se quedan: la regla los pisa solo cuando el valor es negativo.
  El post-proceso de `xlsx_formulas` no toca `<conditionalFormatting>` (su
  regex exige `<c` + espacio). Las huellas de LAYOUT (`_firma_hoja`) no miran
  las reglas; las huellas BYTE A BYTE del libro mensual
  (`tests/test_balance_general_costo_hora.py::test_mensual_es_el_libro_de_hoy`)
  sí cambian con cualquier regla nueva y se recongelan a propósito (8-oct).
  Tests: `tests/test_balance_formato_negativo.py`.
- **BALANCE GENERAL SIN COLUMNAS DUPLICADAS** (8-oct-2026; pedido del
  cliente con captura: «estos valores están duplicados, me ayudan eliminando
  los marcados… y eliminar tacómetro inicial y final (solo en la general)»).
  La variante general (`_COLS_GENERAL`) va SIN el grupo TIEMPO / TACÓMETRO
  (TIEMPO VUELO HR, TACO INICIO, TACO FINAL) ni «COSTO TOTAL (MXN)» (COSTO
  TOTAL MXN, TIPO CAMBIO COSTOS): 39 columnas, el bloque COSTO POR HORA
  empieza en N y STATUS queda en Z. Las tres columnas del bloque que
  repetían esos datos (TIEMPO CALZOS/HOBS (HR), TIPO CAMBIO, TOTAL PARA
  PROVEEDOR (PESOS)) son ahora su ÚNICO lugar: `_disposicion(…,
  costo_por_hora=True)` mete en `letra` un ALIAS por cada par de
  `_ORIGEN_GENERAL` cuyo atributo no tenga columna propia (`tiempo_vuelo` →
  `cph_tiempo_hr`, `tc_costos` → `cph_tc`, `costo_total_mxn` →
  `cph_proveedor_mxn`), así `maestra.total('tc_costos')`, `rango(…)`, las
  citas de 'combustible' / hojas de gastos y las señales por origen (ámbar
  del salto interno en TIEMPO CALZOS/HOBS) siguen funcionando sin tocar a
  nadie; con la columna propia presente, esta manda y no hay alias. La nota
  «calculado por el sistema» pasa del encabezado de COSTO TOTAL MXN al de
  TOTAL PARA PROVEEDOR (PESOS) (`_NOTA_CALCULADO_POR_SISTEMA["cph_proveedor_mxn"]`);
  notas al pie ajustadas (`_CPH_NOTA_AFAC`, `_CPH_NOTA_TOTALES`,
  `_CPH_COMBUSTIBLE_COLA_*`, nueva `_CPH_NOTA_TACOS_AMBAR` en lugar de
  `_NOTA_TACOS_AMBAR`). El Balance mensual y el individual NO cambian (sus
  huellas byte a byte siguen iguales). Fórmulas del bloque intactas: cada
  una cita por llave y la fila es la misma aritmética del API; el caso
  mutado «sin TIPO CAMBIO COSTOS» de los tests se retiró con la columna.
- **COMISIONES a cargo del avión** (6-oct-2026, API 0.0.65; pedido: «cuando
  hay una comisión de un banco, en la parte de total cobrado no refleja el
  monto real que entró a la cuenta» y «la comisión del banco y vendedor se
  puede ir a la columna de (comisiones del vendedor) pero cambiar el nombre a
  "comisiones" y poner notas el tipo de comisión y si hay más de una»;
  respuestas: la del vendedor la absorbe el avión, lo cobrado al cliente por
  ella sigue siendo ingreso de VuelaTour, desde septiembre 2026 por fecha del
  vuelo). **Los números los calcula el API** (fuente única
  `comisionesDelVuelo`, la misma del reparto a socios); aquí se pintan.
  Campos ADITIVOS: cobro `neto_mxn`; vuelo `comisiones_mxn`,
  `comisiones_detalle` (liberal: null/no-lista ⇒ [], solo textos no vacíos),
  `comision_banco_avion_mxn`, `comision_vendedor_prov_mxn`; totales
  `comisiones_mxn`. (1) «COMISIÓN VENDEDOR MXN» → **«COMISIONES MXN»** (llave
  `comisiones_mxn`, misma posición AC en el mensual/individual; en la
  variante general, AL FINAL del bloque COSTO POR HORA —las 12 del cliente
  siguen juntas— y el bloque pasa a 13 columnas; desde el 7-oct-2026 la
  columna vive SOLO en el libro individual, ver la entrada siguiente) con su
  nota de encabezado
  y, por celda, `_nota_comisiones`: `comisiones_detalle` tal cual con «N
  conceptos» arriba si hay más de una línea y el API no lo trae (regex
  `_ENCABEZADO_CONCEPTOS`); sin detalle, una línea por parte del API. Valor
  `_comisiones(x)` = `comisiones_mxn` o, sin él (API ≤ 0.0.64),
  `comision_vendedor_mxn` (null en filas, 0 en TOTALES: su libro no cambia;
  `_valor_total` para TOTALES). (2) **GANANCIA MXN = ROUND(REMANENTE −
  COMISIONES, 2)** y, en la variante general, **REMANENTE VENTA MENOS COMPRA
  = ROUND(VENTA − PROVEEDOR − COMISIONES, 2)** (llave `cph_remanente_mxn`, su
  número es `ganancia_mxn` del API vía `_ORIGEN_GENERAL`, TOTALES = Σ
  ganancia; nota en su encabezado) SOLO en filas con comisiones; sin ellas
  (vuelo anterior o API previo) la fórmula de siempre (`=AA3` /
  `ROUND(K−AA,2)`), byte-idéntica. GANANCIA USD y la cascada de 'balance'
  citan esa ganancia: quedan después de comisiones solas. (3) **COBRO n MXN
  = `_neto_cobro`** (neto o monto); la nota: si `cobrado_con` ya trae el
  desglose del API (`_MARCA_BRUTO`: «Bruto $…»/«neto $…» CON monto —«Registró:
  Neto» es un apodo, no la marca) va «fecha · cobrado_con»; si no,
  `_bruto_neto` («Bruto $X · comisión banco $C · neto $N») ocupa el lugar del
  monto; un cobro sin método/cuenta/registro pero con neto ≠ bruto también
  lleva nota (`_cuenta_algo`). La 4.ª parcialidad agrega netos.
  COBRADO REAL sigue siendo `ROUND(SUM(COBRO n))` (el API manda Σ netos);
  encabezado COBRADO AVIÓN = «COBRADO AVIÓN MXN\n(prorrateado, antes\nde
  comisiones) ****» (saltos explícitos: en una línea envolvía a 4 renglones
  y se recortaba en la fila de 38 pt). (4) 'cobranza': detalle «$19,380.00
  neto (bruto $20,400.00)» (`_monto_detalle_cobranza`) y `_NOTA_COBRANZA_NETO`
  al pie SOLO si algún cobro trae `neto_mxn`. (5) RESUMEN: «COMISIONES MXN» y
  GANANCIA `ROUND(E−F−G−H,2)` cuando el avión trae comisiones ≠ 0 (con 0, la
  de siempre); el API debe mandar en `resumen[].comisiones_mxn` la Σ de
  COMISIONES del libro (si no, la celda queda como valor). Textos: notas al
  pie `_NOTA_COMISIONES` + fragmento de la bandera + `_NOTA_GANANCIA`
  (mensual) o `_CPH_NOTA_REMANENTE` (general), `_RESUMEN_COMISIONES`,
  `_NOTA_COBRADO_REAL` (netos y «antes de comisiones»). 'otros movimientos'
  (R5) y el Libro Dinero no cambian aquí: sus filas llegan armadas. **API
  previo ⇒ el libro de siempre salvo encabezados y sus textos**: verificado
  celda por celda (valor, fórmula, caché, formato, relleno, bordes, notas)
  en los 244 libros de los tests del balance contra el código anterior;
  huellas regeneradas A PROPÓSITO en `test_balance_general_costo_hora`
  (`_FIRMAS_HOY`, `_MIEMBROS_HOY`: cambian sheet1/sheet2/comment1/VML1),
  `test_balance_formulas` (`_FIRMAS_ANTES`), `test_balance_cobrado_con`,
  `test_balance_empresa_vuelatour` y `test_balance_extension_horario`. Tests:
  `tests/test_balance_comisiones.py` (espejo del API con la regla por fila
  vía `ajusta_fila` de `_libro_api`; fórmulas verificadas sin degradar).
  **Revisión (6-oct-2026)**: (a) el **REPARTO a socios** (PDF y Excel) recibe
  en `comisiones_venta_usd` —el campo de siempre, que su cascada ya resta—
  la parte del avión de la comisión bancaria + la provisión del vendedor, y
  seguía imprimiendo «(-) Comisiones de venta» junto a «su pago no es costo
  del avión». Fuente única `app/services/reparto_comisiones.py`: con algún
  avión que absorba comisiones (`comisiones_al_avion`: `comisiones_venta_usd`
  ≠ 0, el mismo criterio con que se imprime la fila y se muestra la columna;
  desde la regla A del 28-ago el API solo lo llena con la regla de
  septiembre, así que no hace falta otra bandera) la fila y la
  columna E se llaman «(-) Comisiones (banco + vendedor)» / «Comisiones (banco
  + vendedor)», el resumen, el pie y las notas del Excel dicen «Con la regla …
  (por fecha del vuelo), el avión absorbe la provisión de la comisión del
  vendedor y su parte de la comisión bancaria de sus cobros» y la línea gris
  del avión que las absorbe (`absorbe_comisiones`) «su provisión resta en
  Comisiones» (cabe en la celda de 120 mm a 7.5 pt: test). **Sin comisiones
  absorbidas, el documento de siempre** (verificado contra b38102d: flujos
  de contenido del PDF y miembros del .xlsx idénticos en 33 payloads, con y
  sin etiqueta). (b) Nombre de la regla ADITIVO `regla_comisiones`
  (`RepartoPdfRequest`, `BalanceAvionRequest`, `BalanceGeneralRequest`;
  validador `regla_comisiones_liberal`). **El API 0.0.65 todavía no lo
  manda**: debe mandar `etiquetaReglaComisiones(vigencia)` («regla sep-2026»)
  — la vigencia es configurable. `_con_regla` cambia «regla de septiembre
  2026» / «desde septiembre 2026» por la etiqueta en las notas de
  COMISIONES, COBRADO REAL, RESUMEN y 'cobranza' (en el general, la del
  request llega al consolidado si este no trae la suya); sin ella
  (`REGLA_COMISIONES_RESPALDO`) el texto de siempre, byte-idéntico. (c) Con
  vuelos de la regla en el periodo (`_regla_vigente`: el consolidado trae
  `totales.comisiones_mxn`, que el API manda SOLO si alguna fila de avión cae
  en la vigencia) la fila 2 de 'otros movimientos' (`_OM_REGLA_COMISIONES`) y
  la base al pie del bloque VUELATOUR (empresa) (`_EMPRESA_REGLA_COMISIONES`,
  también dentro del pie de respaldo) explican la línea de INGRESO «comisión
  del vendedor a cargo del avión …» (la paga el avión, no el cliente) y que la
  comisión bancaria de VuelaTour es solo su parte; sin vuelos de la regla,
  los textos de siempre (huellas intactas). (d) `_NOTA_COBRANZA_NETO` ya no
  dice «las comisiones las absorbe el avión» (un vuelo EXTERNOS va neto y
  ningún avión absorbe nada) y la nota de respaldo de COMISIONES dice
  «provisión» sin «+ IVA». (e) El espejo del API (`_vigente`, `_libro_api`,
  `_flota`) manda la forma EXACTA del JSON: fila vigente sin comisiones con
  `comisiones_mxn` null, las partes en 0 y detalle []; fila EXTERNOS con el
  neto y sin llaves de comisiones; totales del libro si alguna fila trae
  `comisiones_detalle` y de la flota si ALGÚN libro trae `comisiones_mxn`.
  Tests: `tests/test_reparto_comisiones.py` y, en
  `tests/test_balance_comisiones.py`, #403 con la forma real, #504 EXTERNOS
  (mensual y general), la etiqueta del API y las aclaraciones de 'otros
  movimientos' y del bloque empresa. **SUPERSEDIDO el 13-oct-2026 (API
  0.0.74) en lo que toca al vendedor**: lo que absorbe el avión ya no es su
  PROVISIÓN sino el pago capturado como «Comisión del vendedor (avión)»;
  con la bandera `comision_vendedor_solo_gasto_real` los textos de arriba
  que dicen «provisión» tienen su versión `*_GASTO_REAL` /
  `*_DOS_CATEGORIAS` (ver «COMISIÓN DEL VENDEDOR SOLO COMO GASTO REAL, EN
  DOS CATEGORÍAS»). **Rótulo de la columna E (13-oct-2026, lote 2; pendiente
  de las invariantes 50 y 61 del API)**: el encabezado de la columna E del
  Excel del reparto es `ETIQUETA_COLUMNA` («Comisiones (banco + vendedor)»)
  SIEMPRE, también OCULTA sin comisiones absorbidas — antes decía
  «Comisiones venta», el rótulo de la comisión del vendedor de antes del
  28-ago. Es lo ÚNICO que cambia sin comisiones (miembros del .xlsx
  comparados contra el código previo: solo esa celda de `sheet1.xml`, con y
  sin la bandera); con comisiones, byte a byte. La fila del PDF ya decía
  «(-) Comisiones (banco + vendedor)» desde el 6-oct (con 0 no se imprime).
  El nombre se queda «(banco + vendedor)» —el mismo del panel y de la
  cascada del Balance general—: que lo del vendedor es solo el pago «(avión)»
  lo dicen la línea gris y las notas con la bandera. Test:
  `test_xlsx_sin_comisiones_igual_que_siempre`.
- **COMISIONES solo en el libro INDIVIDUAL; el Balance general, antes de
  comisiones** (7-oct-2026, API 0.0.66; pedido: «estas comisiones se están
  duplicando en la general, ya que las tenemos en el apartado de "Otros
  movimientos". Las comisiones deben aparecer en ese apartado pero en el
  individual de los avioncitos» y «en "otros movimientos", en el ingreso
  estás duplicando la comisión»). **El dinero de los socios NO cambia**: la
  cascada, la utilidad cobrada y el reparto son los números del API (iguales
  al 0.0.65, congelados en `_CASCADA_0065`); solo cambia la presentación del
  general. (1) Hoja de vuelos del GENERAL, las DOS variantes, SIN la columna
  COMISIONES: mensual = `_COLS_FLOTA` (`_COLS` sin `comisiones_mxn`, 46
  columnas: de AC en adelante todo corre una — GANANCIA MXN = AC, COBRO 1
  FECHA = AH) con `_DISP_FLOTA`; variante general = `_COLS_GENERAL` de vuelta a
  las 12 columnas del cliente (44 entonces; 39 desde el 8-oct-2026). `_hoja_maestra` elige el juego (el
  individual sigue con `_DISP_MENSUAL`, intacto) y, si el juego no trae
  COMISIONES (`comision_col` None), pinta `_antes_de_comisiones(req)`: fila
  con comisiones ⇒ GANANCIA MXN = `remanente_mxn` del API (fórmula `=AA`) y
  GANANCIA USD = ROUND(remanente ÷ `tc_costos`, 2) calculado como lo evalúa
  Excel (`_numero_de_celda` + `redondear_excel`: la fórmula visible
  ROUND(AC/V,2) lo reproduce y no se degrada) —sin `tc_costos`, VACÍA: el API
  usa el T.C. promedio del libro de SU avión, que el consolidado no tiene—;
  TOTALES: GANANCIA MXN = `totales.remanente_mxn` y GANANCIA USD = Σ de las
  filas tal como se ven (sumadas en orden, como SUM). Sin filas con
  comisiones, el MISMO objeto. **COMISIONES en 0 explícito = null**
  (revisión 7-oct-2026): el API manda null, pero el esquema admite la «regla
  con 0» y `_formula_fila` decide con `_comisiones(v) is None` — con 0.0
  GANANCIA MXN citaba la columna COMISIONES que esta hoja no tiene
  (`_ColumnaAusente`): la celda iba como VALOR y la nota del encabezado se
  cambiaba por «Calculado por el sistema…», falsa. Hoy
  `_fila_antes_de_comisiones` devuelve esa fila SIN la llave
  (`comisiones_mxn`/`comision_vendedor_mxn` None) y, si ninguna fila trae
  comisiones ≠ 0, `_antes_de_comisiones` deja los TOTALES del API: libro
  byte-idéntico al de null (el individual no pasa por aquí). REMANENTE
  VENTA MENOS COMPRA (`cph_remanente_mxn`, `_ORIGEN_GENERAL` → `remanente_mxn`) =
  `ROUND(VENTA − PROVEEDOR, 2)` siempre (definición del cliente). Notas:
  encabezado de GANANCIA MXN y de REMANENTE (`_NOTAS_ENCABEZADO_SIN_
  COMISIONES`), pie `_NOTA_COMISIONES_FLOTA` + `_NOTA_GANANCIA_FLOTA` /
  `_CPH_NOTA_REMANENTE` —sin el fragmento del pago al vendedor: remite a la
  leyenda de 'otros movimientos', así que la bandera `hay_pago_vendedor_real`
  ya no lo cambia (`_MAESTRA_PAGO_VENDEDOR_REAL` se retiró)—,
  `_NOTA_COBRADO_REAL_FLOTA`, `_NOTA_COBRANZA_NETO_FLOTA` (`_hoja_cobranza(...,
  general=True)`) y `_RESUMEN_COMISIONES` (el RESUMEN CONSERVA su columna
  COMISIONES y su GANANCIA después de ellas). (2) Cascada de cada avión en la
  hoja 'balance' del GENERAL: con `balance.comisiones_usd` ≠ 0 y
  `balance.utilidad_antes_comisiones_usd` (ADITIVOS en
  `BalanceAvionBalanceBloque`, `_cascada_con_comisiones`) el bloque abre con
  «UTILIDAD ANTES DE GASTOS USD (antes de comisiones)» y «(−) COMISIONES
  (banco + vendedor) USD» (valores del API, con nota; la etiqueta larga
  envuelve) y UTILIDAD ANTES DE GASTOS = `ROUND(antes − comisiones, 2)` con
  el número de siempre; DESPUÉS, COBRADA y socios, idénticos. Sin los campos,
  con comisiones en 0 o con uno solo: la cascada de siempre. El libro
  INDIVIDUAL los ignora (byte-idéntico: ya resta en su columna). **Las notas
  de esas dos líneas NO prometen cuadres que no siempre se dan** (revisión
  7-oct-2026): `_NOTA_CASCADA_COMISIONES` decía «Son las mismas que descuenta
  el reparto a socios» y no es así — por vuelo es la misma cifra USD
  (`comisionesDelVuelo`), pero el anticipo con comisión bancaria de un vuelo
  que aún no se completa cuenta en el balance y no en el reparto, que solo
  lee COMPLETADOS y CANCELADOS (caso real: #314 de N4142R, CONFIRMADO del
  15-sep-2026, $100,000 con $5,000 de comisión; también #375/#376 de XA-VGV
  en octubre); hoy dice que el total puede no coincidir y por qué.
  `_NOTA_CASCADA_ANTES_COMISIONES` avisa que puede no cuadrar al centavo con
  la Σ de GANANCIA USD de los vuelos del avión en la hoja de vuelos (ahí cada
  remanente va con su T.C. de costos y queda vacía sin él; aquí la comisión
  va con la cifra USD del reparto). Si algún día se quiere el cuadre exacto,
  el API debe mandar por fila `ganancia_antes_comisiones_usd` y sumar eso:
  aquí no se recalcula. (3) Señal
  del API 0.0.66 = algún `aviones[].balance.comisiones_usd`
  (`_avion_cubre_al_vendedor`): con vuelos de la regla, la fila 2 de 'otros
  movimientos', la base del bloque VUELATOUR (empresa) y la nota del RESUMEN
  dicen que en los vuelos COMPLETADOS el pago al vendedor lo cubre el avión
  con su provisión («cubierto por el avión»: solo resta el exceso del pago
  real) y que ya no hay línea de ingreso «a cargo del avión»
  (`_OM_REGLA_COMISIONES_CUBRE_AVION`, `_EMPRESA_REGLA_COMISIONES_CUBRE_AVION`,
  `_RESUMEN_REGLA_CUBRE_AVION`); con el 0.0.65, las de antes (la de 'otros
  movimientos' ya remite a la columna del libro individual). Las filas de
  'otros movimientos' llegan armadas: «cubierto por el avión» NO es la marca
  amarilla. Huellas del GENERAL regeneradas A PROPÓSITO, comparadas celda por
  celda contra HEAD 4189a1e en los 46 libros generales de los tests quitando
  la columna (solo letras de fórmulas corridas, notas y la GANANCIA de filas
  con comisiones); las del individual, intactas (12 libros byte a byte).
  Deploy: pyservices → API 0.0.66. Pendiente que NO es de aquí: la nota al
  pie del libro individual sigue diciendo que el pago al vendedor «sale de
  VuelaTour» (el contrato congeló ese libro) y la del PDF del reparto
  (`reparto_pdf._nota_otros_ingresos`) que VuelaTour «le paga al vendedor»;
  con el API 0.0.66, en los vuelos completados de la regla ese pago lo cubre
  el avión con su provisión. Se ajustan cuando el cliente lo confirme (la
  señal sería `balance.comisiones_usd` en el payload individual, aceptando a
  propósito el cambio de huella del individual solo para esos payloads).
  Tests: `tests/test_balance_comisiones_general.py` (+ los casos del general en
  `test_balance_comisiones`, `test_balance_general_costo_hora`,
  `test_balance_factura_vuelatour`, `test_balance_cobrado_con`,
  `test_comision_vendedor_notas`). **SUPERSEDIDO el 13-oct-2026 (API
  0.0.74)** en lo que toca a la provisión: el pago al vendedor ya no lo
  «cubre el avión con su provisión»; el avión carga el pago REAL y la
  pestaña lo trae «a cargo del avión» con $0.00 (textos `*_GASTO_REAL` con
  la bandera `comision_vendedor_solo_gasto_real`).
- **SIN PROVISIÓN del pago al vendedor en 'otros movimientos'** (12-oct-2026,
  API 0.0.73, invariante 58 del API; pedido del cliente —punto 17 del lote
  del 12-oct— precisado con la captura de la hoja en pantalla: «provisión:
  los gastos que todavía no se hacen no deben aparecer» … «son estas las
  provisiones que están en amarillo las que te digo»: las filas «pago
  comisión vendedor (Pablo Canales) · cubierto por el avión … · PROVISIÓN
  (sin gasto real capturado)» con $0.00 en amarillo y, sin provisión del
  avión, «… · PROVISIÓN (mismo monto que lo cobrado…)» con el monto).
  **El cambio es del API**: mientras no exista el gasto real «Comisión del
  vendedor» ligado al vuelo, la fila ya NO lleva egreso (ni concepto, ni
  monto, ni fecha, ni marca); el ingreso cobrado al cliente sigue (es
  dinero real) y el remanente = ingreso (la fórmula `ROUND(G−D,2)` de
  siempre: el egreso vacío cuenta 0, verificada sin degradar). Con gasto
  real todo sigue EXACTAMENTE como hoy (gasto real pre-vigencia con
  faltan/excede; «cubierto por el avión» con exceso, o con parcial a
  $0.00). **Aquí no se calcula nada: solo cambian 7 leyendas (las 6 del
  contrato y la base del pie del bloque VUELATOUR, abajo), y SOLO con
  la bandera ADITIVA** `sin_provision_vendedor: bool | None = None`
  (`BalanceHojaOtrosMovimientos`, junto a `hay_pago_vendedor_real`, para
  el GENERAL; `BalanceAvionRequest` al nivel RAÍZ, para el libro
  INDIVIDUAL — un API previo no la manda). `_sin_provision_vendedor`
  exige `is True` y MANDA sobre `hay_pago_vendedor_real`
  (`_fragmento_om_pago_vendedor` / `_fragmento_resumen_pago_vendedor`: un
  solo texto con o sin pago real en el periodo). Textos EXACTOS del
  contrato (es-MX): fila 2 de 'otros movimientos'
  `_OM_PAGO_VENDEDOR_SIN_PROVISION` («el pago de la comisión al vendedor va
  apareado en la misma fila SOLO cuando se captura en Gastos el gasto real
  «Comisión del vendedor» ligado al vuelo (con «faltan $…» o «excede $…» si
  no cuadra con lo cobrado); mientras no se capture, la fila no lleva
  egreso ni provisión (pedido del 12-oct-2026: los gastos que todavía no se
  hacen no aparecen). Ese pago NO se captura como «Otros gastos
  VuelaTour»: quedaría duplicado en la hoja 'otros gastos'.»); RESUMEN
  `_RESUMEN_PAGO_VENDEDOR_SIN_PROVISION` («(ingreso cobrado; el pago al
  vendedor se aparea solo cuando se captura el gasto real «Comisión del
  vendedor»: sin él no se provisiona nada en esa hoja).»); con vuelos de
  la regla y el API 0.0.66+ (`cubre_avion`)
  `_OM_REGLA_COMISIONES_CUBRE_AVION_SIN_PROVISION` («… la cubre el avión
  con su PROVISIÓN en su columna COMISIONES; aquí el pago al vendedor
  aparece solo cuando se captura el gasto real y resta únicamente lo que
  ese pago exceda la provisión del avión …»),
  `_EMPRESA_REGLA_COMISIONES_CUBRE_AVION_SIN_PROVISION` («… (los egresos
  propios solo traen lo que un pago real capturado la exceda) …») y
  `_RESUMEN_REGLA_CUBRE_AVION_SIN_PROVISION` («… solo queda lo que un pago
  real capturado la exceda.»); libro INDIVIDUAL con la bandera raíz
  `_MAESTRA_PAGO_VENDEDOR_SIN_PROVISION` («(el pago se aparea ahí cuando se
  captura el gasto real «Comisión del vendedor»).»). Revisión adversaria
  del mismo día: la BASE del pie del bloque VUELATOUR (empresa) decía «… y
  lo pagado o provisionado — no es lo cobrado …» también con la bandera,
  pegada a la regla nueva que dice lo contrario; con la bandera
  `_bloque_empresa` usa `_NOTA_EMPRESA_BASE_SIN_PROVISION` («Ingresos y
  egresos propios = hoja 'otros movimientos': lo cotizado de los vuelos del
  periodo (no cancelados) más los otros ingresos registrados, y lo pagado
  (el pago al vendedor solo cuando se captura el gasto real «Comisión del
  vendedor») — no es lo cobrado: el RESULTADO no es dinero en caja.») con o
  sin regla, también en el pie fijo de respaldo sin `nota`; la `nota` del
  API 0.0.73 que va justo arriba dice a su vez «pago al vendedor (solo el
  gasto real «Comisión del vendedor» capturado; …)». Sin la bandera
  (ausente/null/false) los textos de hoy BYTE A BYTE (las huellas
  congeladas de los demás tests no se mueven); `_OM_REGLA_COMISIONES` /
  `_EMPRESA_REGLA_COMISIONES` (API 0.0.65) no tienen versión con bandera.
  **El amarillo NO se toca** (`_MARCA_PROVISION`, `_LEYENDA_PROVISION`,
  `_es_provision`): no depende de la bandera sino del texto de cada fila;
  con el API nuevo no llega ninguna marca y no se pinta nada (tampoco la
  leyenda); con uno previo sigue pintando. **Lo que NO cambia**: el Libro
  Dinero sigue provisionando (`dinero_xlsx`, se le dirá al cliente), la
  columna COMISIONES del individual y el reparto siguen cargando la
  provisión del vendedor (decisión del 6-oct-2026), y el libro JSON de la
  pantalla es el mismo libro. Consecuencia aceptada (del API): en vuelos
  SIN provisión del avión el remanente y el bloque VUELATOUR suben lo que
  valía la provisión hasta que se capture el gasto real. **Orden de
  deploy: API 0.0.73 antes que pyservices** (con el API nuevo y un
  pyservices previo las filas ya salen sin egreso y solo las leyendas
  quedan un rato desfasadas; al revés, las leyendas nuevas no se activan
  porque falta la bandera). Tests: `tests/test_comision_vendedor_notas.py`
  (sección 12-oct-2026: los 6 textos del contrato y la base del bloque como
  literales, bandera encendida/apagada por cada texto —fila 2 y RESUMEN con
  y sin pago real, las tres «cubre avión», maestra individual, la base del
  bloque VUELATOUR con «cubre avión», con regla sin él, sin regla y en el
  pie fijo sin `nota`—, render del general con el
  payload del 0.0.73 SIN marca en las dos variantes: ninguna celda ámbar,
  sin leyenda, sin «PROVISIÓN (» en la pestaña, fila sin egreso con
  remanente = ingreso y fórmula viva, las tres filas con gasto real
  iguales que hoy, TOTALES = Σ filas; tolerancia: la marca pinta aunque
  viaje la bandera; la bandera de la hoja no toca el individual; rutas
  xlsx y libro JSON). **SUPERSEDIDO el 13-oct-2026 (API 0.0.74)** el «Lo que
  NO cambia» de arriba: el Libro Dinero ya no provisiona y la columna
  COMISIONES del individual y el reparto cargan el pago «Comisión del
  vendedor (avión)», no la provisión (ver «COMISIÓN DEL VENDEDOR SOLO COMO
  GASTO REAL, EN DOS CATEGORÍAS»; la bandera
  `comision_vendedor_solo_gasto_real` manda sobre `sin_provision_vendedor`).
- **FACTURA VUELATOUR según los cobros** (13-oct-2026, API 0.0.74,
  invariante 60 del API; decisión 4 del lote del cliente: un cobro que cae
  en HSBC Dólares / HSBC Pesos / Scotiabank Dólares / Scotiabank Pesos
  exige factura del servicio). **El cambio es del API**: la columna
  «FACTURA VUELATOUR» trae, sin factura registrada, «Pendiente de factura»
  (algún cobro exige factura, o se pidió desde el vuelo) / «No lleva
  factura» (hay cobros y ninguno exige) / vacía = SIN COBROS (antes
  significaba «sin factura»); los textos llegan hechos y se pintan TAL CUAL
  (texto, sin color ni fórmula, con `_texto_literal`) en la hoja maestra
  (individual y «reporte horas FLOTA», `_COLS`/`_COLS_FLOTA`/`_COLS_GENERAL`
  sin cambio), en 'otros movimientos' col J y en el Libro Dinero (hoja 1 y
  'Otros ingresos'; `_factura_liberal` ya acepta cualquier texto). **Aquí
  solo cambia la nota al pie de la hoja maestra**, y SOLO con la bandera
  raíz ADITIVA `BalanceAvionRequest.factura_vuelatour_segun_cobros: bool |
  None = None` (el API 0.0.74 la manda en el libro individual, en cada
  `aviones[]` y en el `consolidado` del general — a diferencia de
  `sin_provision_vendedor`, porque las dos maestras salen de
  `_hoja_maestra` y la nota se lee del `req` que recibe: el consolidado
  en «reporte horas FLOTA», también en la variante general vía
  `_notas_pie_costo_por_hora`): `_nota_factura_vuelatour(req)` devuelve
  `_NOTA_FACTURA_VUELATOUR_SEGUN_COBROS` con `is True` («… sin factura
  registrada dice «Pendiente de factura» cuando algún cobro entró a HSBC o
  Scotiabank (o fue por transferencia, link de pago HSBC o cheque sin cuenta
  capturada) o se pidió factura desde el vuelo, «No lleva factura» cuando
  todos sus cobros fueron por BillPocket, Paywise, efectivo o dólares, y
  vacía = vuelo sin cobros. …») y `_NOTA_FACTURA_VUELATOUR` LITERAL de
  siempre con None/False/ausente (API ≤ 0.0.73): esa constante NO se toca
  ni una letra (la congelan `startswith`/igualdad de
  `test_balance_factura_vuelatour` y las huellas `_firma_hoja` de
  extension_horario, inventario, `_FIRMAS_ANTES` de formulas y
  `_FIRMAS_HOY`/`_MIEMBROS_HOY` de general_costo_hora, cuyos payloads no
  traen la bandera). El Libro Dinero no tiene nota al pie que corregir
  (agregar una cambiaría huellas sin bandera). **Skew en ambos sentidos**:
  pyservices nuevo + API viejo ⇒ libro byte-idéntico; API nuevo +
  pyservices viejo ⇒ `extra="ignore"` descarta la bandera, los textos
  nuevos se pintan igual y solo la nota vieja («vacía = sin factura»)
  queda desfasada; cero 422. Orden de deploy: pyservices → API. Tests:
  `tests/test_balance_factura_vuelatour.py` (esquema aditivo; nota por
  bandera ausente/None/False/True en el individual, en «reporte horas
  FLOTA» y en la variante general, una sola vez y en la MISMA fila tras la
  del ESTATUS DE COBRO; huellas `_firma_hoja` de las DOS maestras SIN
  bandera congeladas con el código de HEAD 12b5e86 —idéntico a 4d206ab
  para esta hoja— y comparadas; con la bandera la ÚNICA celda que cambia
  es la de la nota; columnas 1..46 intactas; «Pendiente de factura» / «No
  lleva factura» como texto en la maestra individual y FLOTA, en col J de
  'otros movimientos' y en el Libro Dinero (hoja 1 y 'Otros ingresos');
  rutas con el JSON del 0.0.74 —raíz, `aviones[]` y `consolidado`— y del
  0.0.73). **ANTICIPO DEVUELTO (14-oct-2026, API 0.0.77, invariante 60 del
  API; decisión del cliente: «¿Un anticipo devuelto lleva CFDI? eso no»)**:
  el dinero devuelto COMPLETO no exige factura salga por donde salga (lo
  decide el API por dinero; lo retenido de una devolución parcial sí exige
  si quedó en una cuenta que la exige) y la columna llega hecha como
  siempre. Aquí SOLO la nota al pie, con la bandera raíz ADITIVA
  `BalanceAvionRequest.factura_vuelatour_devolucion` (`is True`; el API la
  manda justo después de `factura_vuelatour_segun_cobros` en el individual,
  en cada `aviones[]` y en el `consolidado`): `_nota_factura_vuelatour` la
  elige PRIMERO ⇒ `_NOTA_FACTURA_VUELATOUR_DEVOLUCION` («… «Pendiente de
  factura» cuando quedó dinero de algún cobro que entró a HSBC o Scotiabank
  … «No lleva factura» cuando ningún cobro la exige o el dinero se devolvió
  completo (un anticipo devuelto no lleva factura) …»); sin ella
  (None/False/ausente) las DOS notas de siempre, literales y byte a byte
  (`_FIRMA_*_SIN_BANDERA` y las huellas de forma real del 0.0.76 de
  `test_balance_otros_ingresos_avion`). Tests:
  `tests/test_balance_factura_vuelatour.py` (texto literal; manda con o sin
  `segun_cobros`; una sola vez y en la MISMA fila en el individual, FLOTA
  mensual y variante general; contra el libro del 0.0.74 la ÚNICA celda que
  cambia es la de la nota; huellas sin la bandera; ruta con el JSON del
  0.0.77).
- **COMISIÓN DEL VENDEDOR SOLO COMO GASTO REAL, EN DOS CATEGORÍAS**
  (13-oct-2026, API 0.0.74, invariante 61 del API; decisión 3 del lote del
  cliente —«comisión del vendedor SOLO como gasto real en todos lados»— y su
  AJUSTE de las 17:10: «necesitamos una opción que sea Comisión del vendedor
  pero que esa opción sí se cargue al balance del avión, y dejamos el otro
  que se va a balance de vuelatour»). **El cambio es del API**: ya no hay
  PROVISIÓN en ningún libro y el pago al vendedor se captura en una de DOS
  categorías: `COMISION_VENDEDOR` = «Comisión del vendedor (VuelaTour)» (la
  paga VuelaTour en 'otros movimientos'; NUNCA es costo del avión, ni en la
  vigencia) y `COMISION_VENDEDOR_AVION` = «Comisión del vendedor (avión)» (la
  carga el avión del vuelo en su columna COMISIONES, la cascada y el reparto
  a socios, × su parte de la venta, en CUALQUIER fecha, en vuelos
  COMPLETADOS o CANCELADOS con cobros retenidos; la vigencia gobierna solo la
  parte bancaria). Sin pago capturado no se resta nada. En 'otros
  movimientos' la gramática única del API (`egresoPagoVendedorDeVuelo`)
  manda las filas hechas y aquí se pintan TAL CUAL (sin marca, sin amarillo,
  remanente con la fórmula viva `ROUND(G−D,2)`): la «(VuelaTour)» egresa;
  la «(avión)» que carga un avión sale con egreso $0.00 «· a cargo del avión
  XB-… ($… en su columna COMISIONES)» (fila mixta: «· VuelaTour $X · a cargo
  del avión …» con egreso = lo de VuelaTour); la «(avión)» que ningún avión
  carga (vuelo no completado, cancelado sin cobros retenidos, sin avión de la
  flota) la paga VuelaTour con «· incluye «Comisión del vendedor (avión)» que
  ningún avión carga (<motivo>): la paga VuelaTour». El Libro Dinero es la
  CAJA de VuelaTour y resta las DOS (aritmética del API sin cambio).
  **Aquí no se calcula nada: cambian SOLO textos, y SOLO con la bandera
  ADITIVA `comision_vendedor_solo_gasto_real: bool | None = None` (`is
  True`)** — la MISMA bandera trae los textos de las dos categorías (el
  ajuste NO agregó bandera: D3 y ajuste son INSEPARABLES y salen juntos en
  el 0.0.74; con solo 982488d las leyendas dirían que el avión carga «el
  gasto «Comisión del vendedor»», lo que el ajuste corrige). El API la manda
  SIEMPRE desde el 0.0.74 en: (a) la raíz de cada `BalanceAvionRequest`
  —libro individual, cada `aviones[]` del general Y el `consolidado` FLOTA
  (a diferencia de `sin_provision_vendedor`, igual que
  `factura_vuelatour_segun_cobros`)—, (b) `BalanceHojaOtrosMovimientos`, (c)
  `DineroXlsxRequest` y (d) `RepartoPdfRequest`. Con ella MANDA sobre
  `sin_provision_vendedor` y `hay_pago_vendedor_real`. Helpers:
  `_solo_gasto_real(req)` (raíz: notas de la hoja de vuelos con el `req` que
  recibe `_hoja_maestra` —el consolidado en el general—, la cascada de cada
  avión en `_bloque_balance` y la nota multi-avión) y
  `_solo_gasto_real_om(om)` (fila 2, RESUMEN, bloque VUELATOUR); NUNCA
  `_sin_provision_vendedor` como señal del general (el consolidado no la
  trae). **`_fragmento_om_pago_vendedor` / `_fragmento_resumen_pago_vendedor`
  evalúan la bandera PRIMERO** (`*_DOS_CATEGORIAS`); con solo
  `sin_provision_vendedor` (API 0.0.73) siguen los `*_SIN_PROVISION` byte a
  byte. `_bloque_empresa`: base = `_NOTA_EMPRESA_BASE_DOS_CATEGORIAS` con
  `solo_gasto_real`; si no, `_NOTA_EMPRESA_BASE_SIN_PROVISION` con
  `sin_provision`; si no, `_NOTA_EMPRESA_BASE`; luego `+
  _EMPRESA_AFAC_EN_FILAS` con `afac_en_filas` (antes de la regla). Textos
  EXACTOS del contrato §G del ajuste. REESCRITOS EN SU LUGAR (los 10 de D3):
  `_NOTA_ENCABEZADO_COMISIONES_GASTO_REAL` (AC2 del individual; ya no empieza
  con la regla: «COMISIONES = lo que absorbe el avión: …»),
  `_COMISIONES_FUERA_DE_LA_HOJA_GASTO_REAL` (⇒ encabezados GANANCIA /
  REMANENTE de la hoja de vuelos del general),
  `_NOTA_COMISIONES_GASTO_REAL` + `_MAESTRA_PAGO_VENDEDOR_GASTO_REAL` +
  `_NOTA_GANANCIA` (pie del individual; empieza «COMISIONES MXN = lo que
  absorbe el avión»), `_NOTA_COMISIONES_FLOTA_GASTO_REAL` (pie de las dos
  variantes del general: «… por fecha del vuelo, para la parte bancaria»),
  `_RESUMEN_COMISIONES_GASTO_REAL`, `_RESUMEN_REGLA_GASTO_REAL`,
  `_OM_REGLA_COMISIONES_GASTO_REAL`, `_EMPRESA_REGLA_COMISIONES_GASTO_REAL`
  y `_NOTA_CASCADA_COMISIONES_GASTO_REAL` (celda «(−) COMISIONES (banco +
  vendedor) USD»). NUEVOS: `_OM_PAGO_VENDEDOR_DOS_CATEGORIAS` (fila 2),
  `_RESUMEN_PAGO_VENDEDOR_DOS_CATEGORIAS` (RESUMEN) y
  `_NOTA_EMPRESA_BASE_DOS_CATEGORIAS` (base del bloque VUELATOUR), más
  `_NOTA_PARTICIPACION_PAGO_VENDEDOR_AVION`: `_nota_participacion(v,
  solo_gasto_real)` agrega «El pago «Comisión del vendedor (avión)» es la
  excepción: se reparte entre los aviones como la venta (va en COMISIONES).»
  SOLO con la bandera y `comision_vendedor_real_pagos > 0` (la caja de esa
  nota crece de 110 a 150). Reglas de los textos: ninguno nombra la
  categoría vieja («Comisión del vendedor» seguido de la comilla de cierre)
  ni dice «provisión» (única excepción, LITERAL del contrato: la negación
  «sin pago capturado no se provisiona nada en esa hoja» del fragmento del
  RESUMEN); los que pasan por `_con_regla` llevan UNA sola vez «regla de
  septiembre 2026» y ningún otro «septiembre» (todos en `_TEXTOS_CON_MES` y
  en el parametrize de `test_textos_nuevos_con_la_regla_del_api`). El
  respaldo sin detalle de `_nota_comisiones` se rotula «Comisión del vendedor
  (avión) · $…» (antes «(gasto real)»). **Un libro ANTERIOR a la vigencia
  puede traer COMISIONES** (solo del vendedor: un pago «(avión)» absorbido
  en cualquier fecha; la bancaria en 0 y los cobros SIN neto) y entonces
  `_regla_vigente` devuelve la regla y se pintan sus leyendas: están
  redactadas para seguir siendo verdad ahí (`_regla_vigente` y
  `_avion_cubre_al_vendedor` sin cambio). Libro Dinero:
  `_LEYENDA_PAGO_VENDEDOR_GASTO_REAL` (fila 2 de 'Otros ingresos': la caja
  de VuelaTour resta las dos), `_UTILIDADES_FILA6_GASTO_REAL` (fila 6 de
  'utilidades', NUEVA: ya no dice «no restan al avión») y
  `_UTILIDADES_OTROS_INGRESOS_GASTO_REAL` (fila 9) seguida de «Pagado al
  vendedor con gasto real…» (solo > 0) y de `_UTILIDADES_NOTA_COMISION_AVION`
  («De lo pagado al vendedor, $X MXN se capturó como «Comisión del vendedor
  (avión)» en vuelos que carga un avión…») SOLO con la llave ADITIVA NUEVA
  `DineroXlsxRequest.utilidades_comision_vendedor_avion_mxn: float | None =
  None` > 0 (informativa: ya va dentro de lo pagado; es la diferencia
  esperada contra el bloque VUELATOUR del Balance general; el API la manda
  solo > 0). Reparto: `frase_regla(req)` («… el avión absorbe su parte de la
  comisión bancaria de sus cobros y, en cualquier fecha, el pago al vendedor
  capturado como «Comisión del vendedor (avión)» (el de «Comisión del
  vendedor (VuelaTour)» lo paga VuelaTour)»),
  `LINEA_COMISION_VENDEDOR_GASTO_REAL` vía `linea_comision_vendedor(req)` =
  «   incluye comisión del vendedor cotizada (pre-IVA) — pago «(avión)» resta
  en Comisiones» — **medida: 88 caracteres (|Δ| = 2 con la de siempre) y
  301.4 pt en Helvetica-Oblique 7.5 contra un tope de 328.2 pt (120 mm − 12
  de relleno)**: `test_linea_gris_de_las_dos_categorias_medida` —,
  `_nota_otros_ingresos` del PDF («… el pago capturado como «Comisión del
  vendedor (avión)» lo carga el avión ('otros movimientos' … con egreso
  $0.00) y el de «Comisión del vendedor (VuelaTour)» lo paga VuelaTour ahí
  mismo.») y `_nota_comision_por_avion` del Excel (sin nombrar la regla: «su
  pago, si se captura como «Comisión del vendedor (avión)», lo absorbe el
  avión en Comisiones (el de «(VuelaTour)» lo paga VuelaTour)»);
  `absorbe_comisiones` / `comisiones_al_avion` sin cambio (≠ 0 también puede
  venir de un mes anterior a la regla). Llaves ADITIVAS que NO se pintan en
  celda: `BalanceAvionVuelo.comision_vendedor_real_mxn` /
  `comision_vendedor_real_pagos` (los pagos «(avión)» que absorbe la fila;
  respaldo de `_nota_comisiones` y excepción de la nota multi-avión) y
  `BalanceAvionTotales.comision_vendedor_real_mxn` /
  `comision_vendedor_prov_mxn`. Visión (IA de tickets) SIN cambio:
  `valid_cats` y la Literal `CategoriaGasto` no incluyen ninguna de las dos
  (regresión en `tests/test_vision_validaciones.py`). **Sin la bandera
  (ausente/None/False) TODO byte a byte** (también con la llave nueva del
  Dinero: se ignora): huellas `_FIRMAS_HOY`/`_MIEMBROS_HOY`/`_FIRMAS_ANTES`,
  `_FILA_CUBIERTA`/`_FILA_EXCEDE`/`_CASCADA_0065` y los payloads del reparto
  intactos; sin columnas, fórmulas ni merges nuevos (`ENCABEZADOS_SIN_ROUND`
  no cambia). El amarillo (`_MARCA_PROVISION`/`_LEYENDA_PROVISION`/
  `_es_provision`) se queda para APIs previos. **Skew**: API 0.0.74 +
  pyservices previo ⇒ números correctos y leyendas viejas un rato
  (`extra="ignore"`, cero 422); pyservices nuevo + API ≤ 0.0.73 ⇒ sin
  bandera, libros de siempre. Orden de deploy: migraciones del API →
  pyservices (con D3 + D2 + el ajuste) → API. Tests:
  `tests/test_comision_vendedor_notas.py` (sección 13-oct-2026 con los
  literales reescritos y sección «13-oct-2026 · dos categorías»: literales
  §G exactos, sin el nombre viejo ni «provisión», `_fragmento_*` con y sin
  la bandera, base del bloque VUELATOUR con y sin `afac_en_filas` y sin
  `nota`, Libro Dinero filas 2, 6 y 9 con y sin la llave —sin la bandera,
  byte a byte—, esquema y ruta), `tests/test_balance_comisiones.py`
  (`_TEXTOS_CON_MES`, una sola regla por texto, respaldo «(avión)», libro
  ANTERIOR a la vigencia con COMISIONES solo del vendedor que cuadra
  GANANCIA = ROUND(REMANENTE − COMISIONES, 2) sin degradar, nota multi-avión
  con la excepción), `tests/test_balance_comisiones_general.py` (general
  0.0.74 con la fila mixta, la informativa $0.00 y la que ningún avión
  carga: TOTALES = Σ filas, remanente con fórmula viva, `verificar_libro`;
  ninguna leyenda con el nombre viejo; literales),
  `tests/test_balance_otros_movimientos_afac.py` (base + AFAC + regla) y
  `tests/test_reparto_comisiones.py` (PDF y XLSX con las dos categorías,
  línea gris medida, sin bandera byte a byte, regla del API sin
  «septiembre», rutas).
- El reporte por vuelo debe CUADRAR: el desglose (subtotal + TUAS + pernocta
  + extras + ajuste + IVA) suma el total exacto — no omitir líneas del
  desglose canónico v1.3.
- **Reporte por vuelo: pago al vendedor del API 0.0.75** (revisión del lote
  2, 14-oct-2026; invariante 73 del API). Desde el 0.0.75 el API manda en
  `pago_vendedor_usd` el pago REAL al vendedor (los gastos «Comisión del
  vendedor (VuelaTour)» / «(avión)» capturados; 0 si no hay) —lo que resta
  la ganancia— y `neto_vuelatour_usd` = total − ese pago, más los ADITIVOS
  `comision_vendedor_solo_gasto_real: bool | None` y
  `pago_vendedor_cotizado_usd: float | None` (la cotizada con su IVA, dato;
  declarados en `ReporteVueloRequest`: antes pydantic los descartaba). Sin
  esto la sección «Cotización» pintaba «Pago comisión vendedor −$0.00 ·
  Neto VuelaTour = total» (contra el panel, que dice −$580 / $4,060 de la
  MISMA cotización) y la cascada rotulaba «(c/IVA)» un pago real que no es
  IVA (un pago en pesos convertido). Fuente única
  `app/services/reporte_vuelo_vendedor.py` (la usan el PDF y el Excel):
  con la bandera `is True`, la sección «Cotización» pinta
  `pago_cotizacion` (la cotizada) con `etiqueta_pago_cotizacion` («Pago
  comisión vendedor (Juan) (comisión + IVA)» solo si la cotizada supera la
  comisión pre-IVA) y `neto_cotizacion` = round(total − cotizada, 2); la
  cascada resta `pago_balance` (= `pago_vendedor_usd`, el que restó la
  ganancia del API; con 0 no se pinta) rotulado `etiqueta_pago_balance` =
  «(−) Pago real al vendedor (Juan)», sin «(c/IVA)». Sin la bandera (API
  ≤ 0.0.74, o en False) el reporte de siempre BYTE A BYTE (verificado con
  huellas del HTML y de las celdas del Excel en 5 formas de payload).
  Aquí no se calcula ningún pago. Tests: `tests/test_reporte_vuelo_etiquetas.py`
  (esquema; sin pago real, pago real mayor que la comisión sin IVA, #332 sin
  comisión cotizada, y sin bandera / en False igual que siempre).
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

- **Libro de Excel → JSON para la VISTA EN PANTALLA** (12-oct-2026; pedido:
  «ver el balance como el Excel sin descargarlo… mismo formato, colores y
  estructura», y después «editar la celda y que afecte a lo relacionado»).
  `app/services/libro_json.py` (genérico, cualquier libro openpyxl):
  `exportar_libro(wb, titulo, archivo, generado)` verifica las fórmulas
  (`xlsx_formulas.asegurar_valores`: las que no reproducen al API vuelven a
  valor, igual que `guardar`) y vuelca hojas, celdas (`v` crudo, `t` TEXTO
  como lo mostraría Excel —`formatear(valor, number_format)`: `#,##0.00`,
  `"$"#,##0.00`, `0.00####`, `0.0%`, `0.00"%"`, secciones «pos;neg;cero»,
  `General`—, `f` fórmula visible, `s` índice de estilo, `n` nota), estilos
  sin repetir (relleno, color, negrita/cursiva/subrayado, tamaño, alineación,
  ajuste, sangría, bordes `lrtb`, formato), combinadas, anchos/altos,
  paneles congelados, color de pestaña y el formato condicional «cellIs» YA
  RESUELTO (el rojo de los negativos). Contrato en la cabecera del módulo
  (`version` 1; el panel lo tipa 1:1). Fase 2: `vincular(cell, {tipo, id,
  campo})` anota el registro del API de una celda (registro pegado al libro,
  las celdas no admiten atributos) y viaja como `vinculo`. Los balances se
  parten en `construir_balance_avion` / `construir_balance_general`
  (openpyxl sin guardar) + `render_*_xlsx` (= `guardar`) + `libro_balance_*`
  (= JSON); rutas `POST /pdf/balance-avion-libro` y
  `POST /pdf/balance-general-libro` (mismo payload que las .xlsx; `archivo`
  = el nombre con que se descargaría). Regla: lo que se ve en pantalla es lo
  que se descargaría — `tests/test_libro_json.py` compara el JSON contra el
  .xlsx real (hojas, valores de la caché de cada fórmula, combinadas) y
  congela el formato de número y cada rasgo del contrato.
- **Vínculos gasto → celda (13-oct-2026, API 0.0.75, punto 5 del lote).**
  Pedido del cliente: desde Conciliación, abrir el balance del avión o el
  general (pestaña nueva) EN LA CELDA DEL GASTO. El API manda los ids y
  pyservices solo los TRANSPORTA al libro JSON de la vista en pantalla:
  **aquí no se decide a qué libro ni a qué mes va un gasto; solo se
  transporta lo que el API ya acreditó.**
  - **Entrada (ADITIVA, validadores liberales: jamás un 422 por un
    metadato)**: cada renglón de las hojas de gastos (`BalanceAvionGastoFila`)
    trae `gasto_id` (texto recortado; vacío o no-texto ⇒ None; en un parcial
    de reparto manual, el del gasto PADRE) y `gasto_parte` ('reparto' |
    'remanente'; cualquier otra cosa ⇒ None). Cada fila de vuelo
    (`BalanceAvionVuelo`) trae `op_vinculos`, `piloto_vinculos` y
    `otros_vinculos` (`BalanceVinculoGasto {id, rol, monto_mxn, gasto_mxn}`;
    lo que no es lista ⇒ []; solo se quedan los objetos con `id` de texto no
    vacío; `rol` texto en minúsculas, vacío ⇒ 'suma'; montos: número finito
    o texto numérico, lo demás ⇒ None). `vinculos_gasto: true` en
    `BalanceAvionRequest` / `BalanceGeneralRequest` SOLO en la petición del
    libro (cualquier otro valor ⇒ None).
  - **Salida**: `libro_json.vincular_lista(cell, vinculos)` (registro pegado
    al workbook por `id(ws)` + coordenada, como `vincular`; guarda COPIAS,
    una segunda llamada REEMPLAZA, lista vacía no anota) ⇒ la celda gana
    `vinculos: [{tipo:'gasto', id, rol, monto_mxn?, gasto_mxn?, parte?,
    fila?, columna?}]` (una celda SIN valor ni estilo pero con vínculos SÍ se
    exporta) y `exportar_libro(..., vinculos_gasto=True)` agrega
    `vinculos_gasto: true` como ÚLTIMA llave de la raíz (sin la bandera, las
    8 llaves de siempre). Roles: `fila` = la celda es el MONTO de la fila de
    ese gasto; `suma` = la celda lo suma con otros; `nota` = solo va en la
    nota (TUA, extensión de horario); `sin_tc` = USD sin T.C. que aún no
    suma.
  - **Celdas anotadas** (`balance_avion_xlsx`): MONTO MXN de las hojas de
    gastos (`_hoja_gastos`: «Gastos Indirectos» fusionada, «refacciones»,
    «permisos» —también con `tc_por_fila`, donde la celda es FÓRMULA—, y en
    el general «otros gastos», «repartidos a aviones» y el fallback
    «refacciones») y de combustible (`_hoja_combustible`, individual y
    flota) y VENTA AL AVIÓN del bloque 2 de «inventario» (`_hoja_inventario`),
    todas con rol 'fila' (`_anotar_vinculo_fila`, + `parte` si el API la
    manda); OPERACIONES, PILOTO y OTROS de la hoja maestra (individual y
    «reporte horas FLOTA»; `_VINCULOS_ATTR`, `_anotar_vinculos_vuelo`, con
    `fila` = la clave del vuelo y `columna` = el encabezado sin saltos) con
    los roles que manda el API; y TOTAL PARA PROVEEDOR (PESOS) de la variante
    «general» con la UNIÓN por id de las tres listas (`_vinculos_unidos`: un
    vínculo por id en orden de primera aparición, 'suma' si alguna parte
    suma, `monto_mxn` = Σ de las partes que suman —operación 67.14 + FBO
    87.00 de una factura ⇒ 154.14—, `gasto_mxn` el primero no nulo). NO se
    anotan TOTALES, COSTO TOTAL, COMISIONES ni el AFAC.
  - **El .xlsx sale byte a byte**: el registro no toca openpyxl
    (`tests/test_balance_vinculos.py` (f) compara los miembros del zip con y
    sin ids, individual y las dos variantes del general). Sin ids ni
    bandera, el libro JSON es el de siempre (g). Pruebas:
    `tests/test_balance_vinculos.py` (a)–(h) y las dos de
    `tests/test_libro_json.py` (`vinculos` en lista y la bandera raíz).
- **CLAVE con ID de vuelo por cliente (13-oct-2026, API 0.0.75, decisión 6
  del lote).** El ID («january02-211026» = alias del cliente + número de
  vuelo de ESE cliente + DDMMAA del día Cancún) va JUNTO al folio, que sigue
  siendo la llave del dinero. El API manda la CLAVE YA armada y aquí se pinta
  tal cual: «#364 · beh05-011026 · Beh Kay » en la hoja maestra y la
  cobranza, «#364 · beh05-011026» en 'otros movimientos' (este módulo sigue
  agregando « · CANCELADO» en estados no normales, `_clave_otro_movimiento`)
  y «beh05-011026» en el Libro Dinero (+ « (50 % N4142R)» en multi-avión; un
  vuelo de servicio siempre «vtservicio»). `clave_cliente` (ADITIVO, en
  `BalanceAvionVuelo`, `BalanceOtroMovimientoFila`, `DineroVueloFila` y
  `DineroOtroIngresoFila`; viaja SOLO con ID; validador liberal: no-texto o
  vacío ⇒ None) **NO se pinta: solo decide el ancho de la columna A**, POR
  CONTENIDO y con tope — `_formato.ancho_columna_clave(textos, base, tope,
  con_id_cliente=hay_id_cliente(filas))` = `min(max(base, largo pintado + 2),
  tope)`; topes `ANCHO_CLAVE_TOPE_BALANCE` = 48 (maestra base 16, cobranza 22,
  'otros movimientos' 18) y `ANCHO_CLAVE_TOPE_DINERO` = 40 ('dinero-vlos'
  14, 'Otros ingresos' 15, midiendo también las claves de la sección de
  ingresos asignados a un avión). **Sin la llave, los libros son
  byte-idénticos** (huellas congeladas con HEAD 462ce75 en
  `tests/test_clave_cliente_xlsx.py`); con ella solo cambia ese ancho
  (ningún valor, fórmula ni formato). `libro_json` hereda el ancho sin
  tocarlo. Desplegar pyservices ANTES que el API para no ver la CLAVE larga
  cortada en la columna de siempre.
- **PERMISO AFAC con fórmula POR FILA en la hoja de la flota y bloque
  «PERMISOS AFAC» en 'otros movimientos' (12-oct-2026, API 0.0.69).** Pedido
  de Ale: «una de las columnas de Permisos de AFAC necesitamos que venga con
  la fórmula en el excel» (el cliente señaló la columna T, PERMISO AFAC
  (PROVISIÓN)) y «en el balance general en la página de otros movimientos no
  se ven los AFAC, aquí deben estar también». (1) La tarifa AFAC es POR
  AVIÓN y la hoja «reporte horas FLOTA» mezcla aviones, así que `$B$1` no
  sirve ahí: cada fila trae `permiso_afac_usd_hr` (ADITIVO del API) y
  la hoja gana la columna «AFAC USD/HR (tarifa del avión)» AL FINAL
  (`_COL_TARIFA_AFAC`, `_COLS_FLOTA_AFAC`/`_DISP_FLOTA_AFAC`, grupo «TARIFA
  AFAC»; ninguna letra se corre: T sigue siendo PERMISO AFAC) **SOLO cuando
  alguna fila trae la tarifa** (`_trae_tarifa_afac`): con un API previo el
  libro es byte-idéntico al de siempre (y las firmas de los demás tests no
  se mueven); `_formula_fila` escribe `ROUND(<tarifa de la fila> ×
  TIPO CAMBIO COSTOS × HORAS COBRADAS, 2)` en las filas con tarifa y deja el
  valor en las demás (EXTERNOS, avión sin tarifa, API previo). El libro
  INDIVIDUAL conserva `$B$1`; la variante «general» (costo por hora) no
  lleva PERMISO AFAC. Notas: B1 (`_NOTA_B1_TARIFA_POR_AVION`) y el encabezado
  de la columna (`_NOTA_ENCABEZADO_TARIFA_AFAC`). (2) `BalanceHojaOtrosMovimientos.afac`
  (ADITIVO, `BalanceAfacOtrosMovimientos`): `_bloque_afac_otros_movimientos`
  pinta DEBAJO de TOTALES —para no entrar a sus sumas— una fila por avión
  con su provisión del periodo (ingreso, «provisión AFAC del periodo · N
  vuelos · tarifa X USD/hr»), una por permiso pagado (egreso) y SUBTOTAL AFAC
  con fórmulas (Σ y provisión − pagado), con la leyenda de que es INFORMATIVO:
  la provisión ya resta en cada avión y los permisos pagados viven en la hoja
  «permisos» del libro individual. Sin `afac` la pestaña es la de siempre.
  Las firmas de valores NO cambian (los fixtures compartidos no llevan la
  tarifa; `_general_con_tarifa_afac` la pone solo en su test). Tests:
  `test_general_maestra_tarifa_por_avion_y_promedios_como_valor` y
  `test_general_otros_movimientos_bloque_permisos_afac`. **(2) SUPERSEDIDO el
  13-oct-2026 con `afac_en_filas`** (entrada siguiente): el bloque
  informativo queda solo para un API previo (0.0.69–0.0.73).
- **PERMISOS AFAC como FILAS que SUMAN en 'otros movimientos'** (13-oct-2026,
  API 0.0.74, invariante 62 del API; decisión 2 del lote del cliente: «AFAC
  en "otros movimientos" del Balance general: filas en la tabla principal
  que suman en TOTALES (provisión AFAC del periodo = ingreso de VuelaTour
  por avión; cada permiso pagado = egreso), ya no bloque informativo bajo
  totales»). **El cambio es del API**: anexa AL FINAL de `filas` (tras la
  última fila por vuelo) UNA fila «AFAC · matrícula» por avión —INGRESO = la
  provisión AFAC del periodo de su libro (Σ de su columna T); en el EGRESO
  se CITAN todos los gastos «Permiso» de su hoja 'permisos' (sean o no de
  AFAC) con egreso **$0.00** «· a cargo del avión X ($… en su hoja
  «permisos»)»; remanente = provisión— y deja de mandar `afac`. **Revisión
  del 13-oct-2026 (invariante 62 del API), CONFIRMADA por el cliente el
  14-oct-2026 («el AFAC lo cargan al Avión»)**: el permiso pagado resta en
  la cascada de su avión (hoja 'permisos') y en el reparto; egresarlo aquí
  lo contaría DOS veces, así que la fila lo cita sin egresarlo (patrón de
  la «Comisión del vendedor (avión)»); la bandera `afac_en_filas` trae los
  textos que lo dicen (`_OM_SECCION_AFAC`, `_OM_AFAC_EN_FILAS`,
  `_EMPRESA_PAGOS_AFAC`, `_EMPRESA_AFAC_EN_FILAS`, reescritos EN SU LUGAR).
  No hay alternativa pendiente: estos textos NO vuelven a decir «egreso».
  (El candado de un PERMISO sin avión es del API: pre-cierre
  `permisos_sin_avion`, 0.0.77. `_NOTA_PERMISOS_TC` NO se tocó en ese lote:
  es texto impreso que el API ≥ 0.0.74 activa siempre con `tc_por_fila` y
  cambiarlo sin bandera nueva rompería el libro de un API ≤ 0.0.76.)
  **Aquí no se calcula
  nada**: montos, notas y remanente llegan hechos; cada fila pasa por
  `pinta()` de siempre (swatch del avión en A, `ROUND(G{r}-D{r},2)`, notas
  como comentario, rojo si negativo) y SUMA en TOTALES con las MISMAS
  fórmulas (`ROUND(SUM(D4:Dn),2)` / `ROUND(SUM(G4:Gn),2)`). Contrato
  ADITIVO: `BalanceOtroMovimientoFila.seccion: str | None` ('AFAC' en esas
  filas) y `BalanceHojaOtrosMovimientos.afac_en_filas: bool | None`
  (`_afac_en_filas` exige `is True`); `afac` (`BalanceAfacOtrosMovimientos`)
  se CONSERVA en el esquema y se pinta SOLO sin la bandera (un API
  0.0.69–0.0.73). Con la bandera, y SOLO con ella: (1) encabezado de sección
  `_OM_SECCION_AFAC` («PERMISOS AFAC POR AVIÓN · provisión del periodo
  (ingreso de VuelaTour) y permisos pagados (a cargo de su avión: egreso
  $0.00)») justo antes de la
  PRIMERA fila `seccion == 'AFAC'`: en A, negrita NAVY tamaño 10, relleno
  LIGHT en A:J, sin montos (SUM lo ignora), SIN merge y SIN fila en blanco;
  (2) la fila 2 suma `_OM_AFAC_EN_FILAS` AL FINAL (después del texto que
  toque por `sin_provision_vendedor`/`comision_vendedor_solo_gasto_real`/
  `hay_pago_vendedor_real` y de la regla; dice que la fila SÍ suma en
  TOTALES y en el bloque VUELATOUR, qué es su ingreso y su egreso y que los
  permisos de vuelos EXTERNOS no entran); (3) bloque VUELATOUR: etiquetas
  `_EMPRESA_INGRESOS_AFAC` / `_EMPRESA_PAGOS_AFAC` (kwarg `afac_en_filas` de
  `_bloque_empresa`, que pasa `construir_balance_general`) y
  `_EMPRESA_AFAC_EN_FILAS` concatenado a la base del pie
  (`_NOTA_EMPRESA_BASE` o `_NOTA_EMPRESA_BASE_SIN_PROVISION`) ANTES de la
  regla, con `nota` del API y en el pie fijo (`_NOTA_BLOQUE_EMPRESA.replace`
  lo hereda); (4) el bloque informativo `afac` NO se pinta aunque llegara
  (payload mixto: nunca doble). Textos literales del contrato congelados en
  `test_textos_son_los_del_contrato`; ninguno lleva «· PROVISIÓN (» (el
  amarillo no se activa: «PERMISO AFAC (PROVISIÓN)» no tiene el punto medio
  delante). `_TITULO_AFAC_OM` conserva «(informativo)» como discriminador:
  los tests buscan el encabezado nuevo por IGUALDAD, nunca por el prefijo
  «PERMISOS AFAC». **Skew en los dos sentidos**: API 0.0.74 + pyservices
  previo ⇒ las filas se pintan y SUMAN como filas normales (sin encabezado
  ni leyendas, sin bloque) y los números cuadran con el bloque VUELATOUR;
  API 0.0.69–0.0.73 + pyservices nuevo ⇒ el bloque informativo de siempre,
  pestaña idéntica; filas con `seccion` sin la bandera (ningún API las
  manda) ⇒ filas normales sin encabezado. **Sin la bandera, libro byte a
  byte** (todo texto nuevo cuelga de `is True` y el encabezado además de
  `seccion == 'AFAC'`; nada depende de `afac` ni de `len(filas)`): huellas
  `_FIRMAS_ANTES`, `_FIRMAS_HOY`/`_MIEMBROS_HOY` y las de empresa,
  extensión y factura NO se regeneraron (con la llave ausente, null o
  false). La variante «general» (costo por hora) pinta la MISMA pestaña; la
  vista en pantalla (`libro_json`) hereda encabezado y filas solas. Orden de
  deploy recomendado pyservices → API (los dos órdenes son numéricamente
  seguros; con el API primero los INGRESOS PROPIOS suben sin que la etiqueta
  lo diga). Tests: `tests/test_balance_otros_movimientos_afac.py`
  (literales; encabezado, filas, comentarios, rojo y TOTALES con caché y
  `verificar_libro`; un solo encabezado; payload mixto sin bloque; sin
  bandera el bloque informativo y filas sin encabezado; fila 2 con y sin
  bandera sobre cada texto base, también con la regla del API 0.0.74;
  bloque VUELATOUR con y sin bandera, con `sin_provision`, con regla y pie
  fijo; sin amarillo; permisos USD sin T.C. citados con egreso $0.00 y sin
  la marca «sin TC»; esquemas, JSON del 0.0.74 y
  mixto, rutas xlsx y libro JSON; huellas de valores, layout y bytes con la
  llave ausente/null/false; variante general idéntica).
  **SUPERSEDIDO el 16-oct-2026 por AFAC POR VUELO** (entrada siguiente):
  el API 0.0.89 ya no manda `afac_en_filas` ni las filas «AFAC ·
  matrícula»; su camino se conserva por skew con un API 0.0.74–0.0.88.
- **AFAC POR VUELO: reservado vs costo real en 'otros movimientos', con la
  celda del costo real EDITABLE en pantalla** (16-oct-2026, API 0.0.89,
  invariante 90 del API, URGENTE del cliente). Pedido: «debe aparecer en
  balance general en la hoja de Otros movimientos para poder confirmar el
  monto […] 2 columnas juntas, una […] lo reservado de afac […] y otra de lo
  que realmente costó y que en esa celda poder poner el monto y se registre
  sin tenerte que ir a otro lado»; después: «la columna nueva que te pedí
  mejor que sea en la columna de egreso». Decisiones del cliente: UNA fila
  por vuelo; el avión sigue pagando lo RESERVADO (su columna PERMISO AFAC
  (PROVISIÓN)) y el COSTO REAL es egreso de VuelaTour; se escribe solo el
  monto del vuelo; los gastos marcados «Pago de AFAC» dejan de restar al
  avión y se comparan cada mes. **El dinero es del API** (fuente única
  `afac-por-vuelo.util#filasAfacPorVuelo`); aquí se pinta. Contrato ADITIVO:
  `BalanceOtroMovimientoFila.afac` (`BalanceAfacVuelo`, solo en filas
  `seccion == 'AFAC_VUELO'`), `BalanceHojaOtrosMovimientos.afac_por_vuelo`
  (`_afac_por_vuelo` exige `is True`; MANDA sobre `afac_en_filas`) y
  `.afac_resumen` (`BalanceAfacResumen` + `BalanceAfacPago`), y
  `BalanceAvionRequest.pagos_afac` (`BalanceAvionHojaGastos`); todo con
  default y validadores liberales en el metadato (`editable` solo con un
  `true` JSON; jamás un 422). (1) Las filas AFAC_VUELO llegan contiguas al
  final de `filas` y se pintan con `pinta()` como FILAS NORMALES: A clave
  (swatch; CANCELADO en rojo), B fecha, C concepto del costo real («AFAC
  costo real · capturó … el …» / «· por capturar»), **D EGRESO = COSTO
  REAL** (vacía, no 0, si no se ha capturado), F concepto del reservado, **G
  INGRESO = RESERVADO** (valor tal cual; la fórmula con tarifa × T.C. ×
  horas no cabe porque esas celdas no están en la fila: el cálculo va en la
  nota de G), I `ROUND(G-D,2)` y SUMAN en TOTALES con las fórmulas de
  siempre (sin fórmula especial). Con la bandera: fila separadora
  `_OM_SECCION_AFAC_VUELO` antes de la primera (estilo del encabezado por
  avión: negrita NAVY 10, LIGHT A:J, sin montos, sin merge) y, en la celda D
  de cada fila, el formato MONEY (también vacía), el verde suave
  `FILL_AFAC_EDITABLE` si el API la manda `editable` y el vínculo
  `afac_real` del libro JSON (`_vinculo_afac_real`, contrato en la cabecera
  de `libro_json`: `{tipo, campo, vuelo_id, clave, matricula, valor,
  reservado_mxn, editable, motivo, mes, capturado_por, capturado_at}`; solo
  en D, NUNCA en G; sin `vuelo_id` no hay vínculo). Si no fueran contiguas
  (defensa; se loguea): filas normales sin separador. (2) Fila 2:
  `_OM_AFAC_POR_VUELO` en el lugar de `_OM_AFAC_EN_FILAS`. (3) Bloque
  «AFAC DEL PERIODO» (`_bloque_afac_periodo`) DEBAJO de TOTALES (no entra
  a sus sumas) con `afac_resumen` (viaja también solo con pagos marcados,
  sin filas ni bandera): título, leyenda combinada A:J, «AFAC por vuelo» con
  D = `ROUND(SUM(D a:b),2)` (Σ costo real), G = `ROUND(SUM(G a:b),2)` (Σ
  reservado), I = `ROUND(SUM(I a:b),2)` (remanente) sobre el rango de las
  filas AFAC (sin filas o no contiguas ⇒ valores del API); un renglón «pago
  a AFAC» por gasto marcado (detalle en C, monto en D con `vincular_lista`
  `{tipo:'gasto', id, rol:'fila'|'sin_tc', monto_mxn, columna:'pagado a
  AFAC'}`, fecha en E; USD sin T.C. ⇒ D vacía con nota), «PAGADO A AFAC» =
  `ROUND(SUM(D pagos),2)`, «DIFERENCIA» = `ROUND(Dpagado-Dreal,2)` en rojo si
  ≠ 0 y, solo si > 0, «fuera de libro» (red del API, valor). Todas
  verificadas: la que no reproduce el centavo vuelve a valor. (4) Bloque
  VUELATOUR: kwarg `afac_por_vuelo` de `_bloque_empresa` ⇒
  `_EMPRESA_INGRESOS_AFAC_VUELO` / `_EMPRESA_PAGOS_AFAC_VUELO` y el pie suma
  `_EMPRESA_AFAC_POR_VUELO` (antes de la regla). (5) Libro INDIVIDUAL: con
  `pagos_afac` (≥ 1 fila) la fila 3 de 'permisos' lleva en cursiva «N
  pago(s) de AFAC ($total) los paga VuelaTour: no restan aquí (ver Balance
  general · otros movimientos, bloque AFAC DEL PERIODO).» con el detalle en
  su nota (`_linea_pagos_afac`; total = `total_mxn` del API). Ningún texto
  nuevo lleva «sin TC» ni «· PROVISIÓN (». **Skew**: API 0.0.89 + pyservices
  previo ⇒ las filas se pintan y SUMAN como filas normales (TOTALES y bloque
  VUELATOUR correctos), sin separador, sin celda editable ni bloque del
  periodo; pyservices nuevo + API ≤ 0.0.88 ⇒ byte a byte (huellas
  `_HUELLAS_0088` del HEAD 575c816 con los payloads de forma real del 0.0.88
  —cierre de mes + `afac_en_filas`— en `tests/_afac_vuelo_payloads.py`, y
  las de `test_cierre_mes` con las llaves AFAC en neutro). Orden de deploy
  del lote: migración `20261016000005` → pyservices (compuerta `GET
  /version`) → API 0.0.89 → panel. Tests:
  `tests/test_balance_otros_movimientos_afac_vuelo.py` (+ celda vacía con
  vínculo en `test_libro_json.py`). **Revisión adversaria (16-oct-2026)**
  contra los payloads REALES del API 0.0.89 (volcados de `payloadGeneral` /
  `payloadAvion` de la rama afac-real, mensual y general, con un real, un
  vuelo por capturar y dos pagos marcados): nombres 1:1, TOTALES cuentan
  cada fila UNA vez, cero fórmulas degradadas, vínculo `afac_real` solo en D
  de cada fila AFAC_VUELO (75 filas, multi-avión no dueña, solo-real, real 0,
  no contiguas, solo pagos), huellas `_HUELLAS_0088` recalculadas sobre un
  export limpio de 575c816 (idénticas) y skew API 0.0.89 + pyservices 575c816
  (TOTALES correctos, sin vínculo). Dos arreglos: (a) `_es_provision` ignora
  las filas AFAC_VUELO — su `nota_egreso` son las NOTAS LIBRES del operador
  (PUT afac-real), y una nota con «· PROVISIÓN (» pintaba de amarillo la
  celda editable y sacaba la leyenda; (b) la línea de 'permisos' del
  individual dice cuántos pagos van en USD sin T.C. («($1,500.00; 1 en USD
  sin T.C., no suma)»): el `total_mxn` del API (`buildHoja`) los cuenta 0 y
  la línea los presentaba como incluidos (o citaba «$0.00»).
- **Otros ingresos del avión: pestaña 'Otros ingresos', línea «(+) OTROS
  INGRESOS USD», hoja consolidada, Libro Dinero y reparto** (13-oct-2026,
  API 0.0.75; invariante «INGRESO ASIGNADO A UN AVIÓN» del API, migración
  del API 20261013000005). Pedido: un ingreso registrado en Ingresos con
  avión (caso real ING-4 de N990GG, $46,550.80 MXN, «reembolso del pago del
  serv de 50 hrs en toledo») cuenta UNA vez, en el mes de su fecha, para ESE
  avión y deja de ser ingreso de VuelaTour. **El cambio de dinero es del
  API** (fuente única `filaLibroDeIngreso` / `partirIngresosPorAvion`); aquí
  se pinta y NO se convierte nada. (1) **La PRESENCIA de
  `BalanceAvionRequest.otros_ingresos_avion` manda** (`BalanceAvionHojaOtros
  Ingresos`: `filas` de `BalanceAvionOtroIngresoFila`, `total_mxn`, `usd`):
  el libro INDIVIDUAL gana la pestaña «Otros ingresos» entre 'permisos' y
  'balance' (el API 0.0.75 la manda SIEMPRE, vacía = «Sin otros ingresos
  asignados a este avión en el periodo.») y la cascada la línea `_FILA_OTROS
  _INGRESOS` entre «(−) PERMISOS USD» y «UTILIDAD DESPUÉS DE GASTOS USD» —se
  pinta aunque `balance.otros_ingresos_avion_usd` venga None (celda vacía:
  un ingreso USD sin T.C. deja la hoja y la utilidad en null)—, con nota
  `_NOTA_CASCADA_OTROS_INGRESOS` solo con valor. Índices desde UTILIDAD ANTES
  (`k`): 5 otros ingresos (SUMA), 6 DESPUÉS = `ROUND(B{k}-B{k+1}-B{k+2}-
  B{k+3}-B{k+4}+B{k+5},2)`, 7 pendiente, 8 COBRADA = `ROUND(B{k+6}-
  B{k+7},2)`; los socios citan `$B${k+8}` (en el general, `k` ya incluye las
  dos previas de comisiones). Individual: la línea cita `'Otros
  ingresos'!$C$5` (`_refs_balance_individual(..., con_otros_ingresos)` y
  `_CASCADA_REFS`); GENERAL: valor, como el resto del bloque. (2)
  `_hoja_otros_ingresos_avion` (modelada en `_hoja_gastos`, 11 columnas
  `_ENCABEZADOS_OTROS_INGRESOS`, anchos `[12, 22, 50, 15, 9, 11, 15, 14, 15,
  22, 36]`, panel fijo A8, DETALLE y NOTAS envueltos con alto estimado):
  INGRESO MXN G = `ROUND(D*F,2)` en USD con `tc_aplicado` (el T.C. EXACTO
  del ingreso) y `ROUND(D,2)` en MXN (otra cosa: valor); COMISIÓN H = valor;
  NETO I = `ROUND(G-H,2)` (sin comisión, `ROUND(G,2)`); A5 = `ROUND(SUM(I8:
  In),2)`, B5 cita el TC PROMEDIO de la maestra y C5 = `ROUND(A5/AVERAGE(T.C.
  de costos de la maestra),2)` (`_USD_TC_LIBRO`, el `tcHojas` del API; sin
  vuelos, valor: el oficial del cierre). Todas con ROUND (nada nuevo en
  `ENCABEZADOS_SIN_ROUND`) y verificadas: una que no reproduzca el centavo
  vuelve a valor. Nota al pie `_NOTA_OTROS_INGRESOS_AVION` envuelta. (3)
  GENERAL: hoja **'otros ingresos aviones'** después de 'repartidos a
  aviones' y antes de 'inventario', SOLO si `consolidado.otros_ingresos
  _avion` trae filas (el API la manda solo así), título
  `_TITULO_OTROS_INGRESOS_GENERAL`, resumen «TOTAL NETO MXN | TOTAL USD (suma
  de los aviones)» con B5 como VALOR (cada libro convirtió con su T.C.), el
  DETALLE ya llega con «MATRÍCULA · » y se tiñe con `avion_color`; nota
  `_NOTA_OTROS_INGRESOS_AVIONES_GENERAL`. Bandera ADITIVA
  `BalanceHojaOtrosMovimientos.otros_ingresos_por_avion` (`is True`; el API
  0.0.75 la manda SIEMPRE): la fila 2 de 'otros movimientos' suma
  `_OM_OTROS_INGRESOS_POR_AVION` AL FINAL (después de `_OM_AFAC_EN_FILAS`).
  (4) **Libro Dinero**: `DineroXlsxRequest.otros_ingresos_avion` (solo con
  filas; `DineroOtroIngresoFila` gana `matricula`/`avion_color`) se pinta al
  final de 'Otros ingresos' tras TOTALES, una fila en blanco y el encabezado
  `_SECCION_OTROS_INGRESOS_AVION` (A:J, negrita NAVY, relleno LIGHT), con el
  MISMO pintado de las filas de arriba (helper `pinta`) y la CLAVE teñida;
  **NO suman a TOTALES** (el API ya los excluye de `otros_ingresos` y de
  `utilidades_otros_ingresos_mxn`); la fila 2 suma
  `_LEYENDA_OTROS_INGRESOS_AVION`. 'utilidades': con ALGÚN
  `utilidades_aviones[].otros_ingresos_mxn` cada avión aporta 5 columnas
  («OTROS\nINGRESOS\n<mat>» primero; vacía sin dato) y la fila 9 suma
  `_UTILIDADES_NOTA_OTROS_INGRESOS_AVION`. (5) **Reparto**: `RepartoAvion`
  gana `otros_ingresos_avion_usd` (SUMA; `saldo_usd` del API YA lo incluye),
  `otros_ingresos_avion_sin_tc_mxn` y `otros_ingresos_avion`
  (`RepartoOtroIngresoLinea`); null ⇒ 0/[]. PDF (`_filas_otros_ingresos
  _avion`): fila `FILA_OTROS_INGRESOS_AVION` (no gris) antes de las
  deducciones solo con monto ≠ 0, hasta 5 renglones grises «   ING-4 ·
  categoría · concepto» con lo que sumó o «sin T.C.» (recortados con «...»
  al ancho de la celda de 120 mm) y «   y N más», y el gris
  `LINEA_OTROS_INGRESOS_SIN_TC` con los MXN que no entraron; resumen de 5
  columnas (`ENCABEZADOS_RESUMEN_CON_OTROS_INGRESOS`, 5 × 34 mm, el gris de
  «Otros ingresos VuelaTour» pasa a la columna 2) para que venta + otros −
  gastos = saldo. XLSX: columna 15 `ENCABEZADO_OTROS_INGRESOS_AVION` (MONEY,
  borde, total en negrita, ancho 18, FUERA de `money_cols`) y la nota
  `NOTA_OTROS_INGRESOS_AVION` tras la de la tabla, SOLO con dato. **Payload
  viejo (sin las llaves, o en null/vacías/0) ⇒ libros byte a byte**:
  huellas de miembros (balance individual y general, Libro Dinero, Excel del
  reparto) y de los flujos de contenido del PDF calculadas con el código de
  HEAD d58faa1 antes del cambio. Orden de deploy: pyservices ANTES del API
  0.0.75 (con el API primero el libro sale sin pestaña ni línea y la
  utilidad después sube sin explicación; orden consolidado del lote en
  «Principios»). **Nota de la pestaña corregida (revisión del lote 2,
  14-oct-2026)**: `_NOTA_OTROS_INGRESOS_AVION` decía que el reparto «puede
  diferir unos centavos» y era falso — el reparto convierte cada ingreso
  por separado (USD = su neto; MXN ÷ su T.C. propio u oficial de su día) y
  esta hoja divide el total entre el T.C. promedio del libro, así que la
  diferencia crece con la distancia entre esos T.C. (ING-4: 2,710.47 USD
  aquí vs 2,746.70 en el reparto). Hoy dice «… puede diferir del de esta
  hoja según la distancia entre esos T.C. y el promedio del libro (no son
  centavos), como pasa con los gastos.». Solo cambia ese texto (las huellas
  sin las llaves siguen byte a byte). Tests:
  `tests/test_balance_otros_ingresos_avion.py`,
  `tests/test_dinero_otros_ingresos_avion.py`,
  `tests/test_reparto_otros_ingresos_avion.py` y
  `tests/test_libro_json.py` (la pestaña y sus fórmulas en pantalla).
  **SUPERSEDIDO con el API 0.0.77** (entrada siguiente): esa nota y la
  división entre el T.C. promedio quedan solo para un API ≤ 0.0.76.
- **Conversión ÚNICA del ingreso de un avión: «Otros ingresos» por ingreso**
  (14-oct-2026, API 0.0.77; invariante 67 del API; decisión del cliente:
  «unificar la conversión del ingreso de un avión entre balance y reparto
  (ING-4 2,710.47 vs 2,746.70): hazlo, si es lo mejor»). **El cambio de
  dinero es del API**: balance, reparto y tablero convierten CADA ingreso con
  la MISMA función pura (`usdDeIngresoAvion`): USD ⇒ su neto (monto −
  comisión) tal cual; MXN ⇒ NETO ÷ (T.C. propio ?? T.C. del gasto LIGADO ??
  oficial del día de su `fecha`); un REEMBOLSO ligado a un gasto en pesos que
  ESE libro cuenta usa el T.C. con que ese libro convierte su gasto (se
  anulan exactos). ING-4 = 2,746.70 USD en los tres (antes 2,710.47 aquí,
  ÷ el promedio 17.17442). Aquí NO se convierte nada. Contrato ADITIVO:
  `BalanceAvionOtroIngresoFila` gana `tc_usd` (MXN: el T.C. EXACTO, sin
  redondear; USD: null), `neto_usd`, `tc_usd_origen` ('PROPIO' |
  'GASTO_LIGADO' | 'OFICIAL_DEL_DIA' | null) y `tc_usd_nota` (nota armada por
  el API; null con T.C. propio o en dólares) —origen y nota con validador
  liberal: no-texto o vacío ⇒ None—, AL FINAL (antes de `matricula` /
  `avion_color` del consolidado); `BalanceAvionHojaOtrosIngresos` gana
  `usd_por_ingreso` (el API la manda SIEMPRE como ÚLTIMA llave: individual,
  `aviones[]` y consolidado FLOTA; `usd` = round2(Σ `neto_usd`), null si a
  alguna fila le falta, vacía = 0). **Solo con `usd_por_ingreso is True`**
  (`_usd_por_ingreso`): 13 columnas `_ENCABEZADOS_OTROS_INGRESOS_POR_INGRESO`
  (… NETO MXN (I), **T.C. A USD (J)**, **NETO USD (K)**, CUENTA (L), NOTAS
  (M)), anchos `[12, 22, 50, 15, 9, 11, 15, 14, 15, 11, 14, 22, 36]`, nota
  combinada A:M. J = `tc_usd` (formato TC; en dólares, vacía), relleno
  `TC_OFICIAL_FILL` SOLO con `tc_usd_origen == 'OFICIAL_DEL_DIA'` y nota de
  la celda = `tc_usd_nota` (el reembolso ligado: sin azul, con la nota del
  gasto). K (`_celdas_usd_por_ingreso`): pesos ⇒ `ROUND(I/J,2)`; dólares con
  T.C. ⇒ `ROUND(I/F,2)` (reproduce monto − comisión: el error es ≤ 0.01 ÷
  T.C.); otro caso ⇒ el `neto_usd` como valor; sin `neto_usd` ⇒ vacía (jamás
  un 0 falso); todo por `_formula` (si no reproduce el centavo del API,
  `finalizar` la vuelve valor). Resumen individual «TOTAL NETO MXN | T.C. |
  TOTAL USD»: A5 igual, B5 = texto «por ingreso (col. J)» (cursiva MUTED) y
  **C5 = `ROUND(SUM(K8:Kn),2)`** con `hoja.usd` (sin filas, valor 0): la
  cascada SIGUE citando `'Otros ingresos'!$C$5`. General («otros ingresos
  aviones»): mismas 13 columnas y **B5 = `ROUND(SUM(K8:Kn),2)`** (antes
  valor). Notas nuevas, SOLO con la bandera (las de siempre se quedan para
  un API ≤ 0.0.76): `_NOTA_OTROS_INGRESOS_AVION_POR_INGRESO`,
  `_NOTA_OTROS_INGRESOS_AVIONES_GENERAL_POR_INGRESO` y
  `_NOTA_CASCADA_OTROS_INGRESOS_POR_INGRESO` (la de la línea de la cascada
  la elige la hoja del MISMO libro, `_nota_cascada_otros_ingresos`:
  individual y cada bloque del general). Las hojas de GASTOS siguen al T.C.
  promedio del libro (dos T.C. por diseño, réplica del Excel del cliente).
  Reparto: cada línea del detalle gana `tc_origen` (ADITIVO); aquí se
  ignora y el PDF/Excel no cambian. **Sin la bandera, byte a byte**: las
  cuatro llaves de fila sin `usd_por_ingreso` (o en null/false) no mueven
  nada (la bandera manda, no la presencia de llaves) y la huella F12 de la
  revisión se toma con un payload de FORMA REAL del API de producción 0.0.76
  (regla de comisiones con las llaves del 0.0.66, 'permisos' con
  `tc_por_fila`, las tres banderas raíz, socio `es_empresa`, bloque
  VUELATOUR y filas sin las llaves nuevas; NO con `tf._payload_tst()`, que
  no trae `tc_por_fila`): `_HUELLAS_FORMA_REAL_0076` (individual y las dos
  variantes del general, .xlsx y JSON de la vista con `vinculos_gasto`)
  calculadas con el código de HEAD b66df98 antes del cambio. Con la forma
  real del 0.0.77 SOLO cambian «Otros ingresos», los números de la cascada
  que la suman y la nota al pie de FACTURA VUELATOUR (las demás hojas, misma
  huella de layout). Orden de deploy y COMPUERTA: «Principios» (lote
  0.0.77). Tests: `tests/test_balance_otros_ingresos_avion.py` (secciones
  5–7: huellas de forma real, columnas y fórmulas verificadas con
  `verificar_libro` —K8 `=ROUND(I8/J8,2)` = 2,746.70, K9 `=ROUND(I9/F9,2)` =
  95.00, C5 = 2,841.70—, azul y nota por origen, ligado, propio, pesos sin
  T.C., USD sin T.C., neto inconsistente degradado a valor, hoja vacía,
  general, payload mixto, vista en pantalla y rutas con el JSON del 0.0.77),
  `tests/test_libro_json.py` (parametrización «por ingreso») y
  `tests/test_reparto_otros_ingresos_avion.py` (`tc_origen` ignorado).
- **Cierre de mes y vuelos que pasan al mes siguiente** (16-oct-2026, API
  0.0.88, invariante 89 del API; pedido: «los vuelos que aún no se cobren se
  marquen en el balance y se muevan para el siguiente mes hasta que se cobre
  y complete»). El API decide TODO (qué vuelo pasa, a qué mes, con qué
  leyenda, qué suma); aquí solo se PINTAN campos aditivos con default (sin
  ellos ⇒ libros byte a byte). Textos compartidos (etiqueta es-MX del mes,
  sello «Mes cerrado el 16-oct-2026 por Ale» en hora Cancún, títulos de las
  listas): `app/services/cierre_mes.py` — ningún número. Validador ÚNICO de
  textos del cierre: `schemas/reportes.texto_cierre_liberal` (folio numérico
  ⇒ texto; blanco/no-texto ⇒ None; jamás un 422). Lo que consume:
  - **Balance por avión / general** (`BalanceAvionVuelo`): `cierre_marca`
    (texto del API: «Pasa a octubre 2026 (pendiente de cobro: cobrado $X de
    $Y MXN)» / «Viene de septiembre 2026» / «Pendiente de cobro: pasará a
    octubre 2026 al cerrar»), `fila_sombra` (fila del mes de ORIGEN: horas
    reales, TODO el dinero en null), `diferido_a` / `diferido_de`
    ('YYYY-MM'). Raíz `cierre` (`BalanceCierre`: `estado` ABIERTO/CERRADO/
    REABIERTO, `periodo`, `cerrado_at`, `cerrado_por_nombre`,
    `reabierto_at/por_nombre`, listas `pasan` / `vienen` / `pasarian` de
    `{clave, folio, cliente, matricula, fecha, venta_mxn, venta_usd,
    cobrado_mxn, cobrado_usd, por_cobrar_mxn, por_cobrar_usd, destino,
    origen, leyenda}`); general además `resumen_cierre` (`{pasan, vienen}`
    por avión: `matricula, vuelos, venta_mxn, por_cobrar_mxn`). Pinta: en la
    hoja de vuelos y en 'cobranza' la marca EN LUGAR del estado (el estado
    real va en la nota), ámbar si sale/pasará y azul claro (DBEAFE) si viene,
    ESTADO a 30 de ancho y la fila crece a los renglones de la marca; en la
    fila que viene, la celda vacía de TIEMPO VUELO lleva la nota «Horas
    voladas en <mes> (mes real del vuelo)». La sombra NO suma porque sus
    celdas llegan en null (`_formula` con valor None = celda vacía SIN
    fórmula) y la Σ de TOTALES —que la incluye en su rango— reproduce el
    total del API (verificado con `verificar_libro`: cero degradaciones
    nuevas). Bloque «CIERRE DEL MES» (sello + listas como VALORES, sin
    fórmulas) bajo TOTALES del individual y al pie del RESUMEN del general
    (CLAVE en A:B y columna AVIÓN + renglón por avión de `resumen_cierre`;
    si el `cierre` viaja solo en el consolidado, se toma de ahí). Notas al
    pie solo con alguna marca. `libro_json` copia `cierre` a la raíz de la
    vista en pantalla (el panel pinta el estado del mes).
  - **Forma REAL de las dos filas** (revisión 16-oct contra la rama del API
    0.0.88, `aircraft/cierre-balance.util.ts`): `aplicarFilaSombra` pone en
    null TODO el dinero, los T.C. (`tc_venta` / `tc_costos`; con un T.C. la
    AVERAGE de TIPO CAMBIO COSTOS de TOTALES ya no reproduciría al API) y
    las horas COBRADAS, vacía detalles/vínculos/cobros y BORRA las llaves de
    comisiones y extensión; conserva horas y TACOS reales. `aplicarFilaViene`
    pone en null `tiempo_vuelo`, tacos y saltos. `tests/_cierre_mes_payloads.py`
    (`aplicar_fila_sombra` / `aplicar_fila_viene`) copia esas listas campo
    por campo: si el API las cambia, se cambian ahí.
  - **La marca NUNCA reemplaza «CANCELADO»** (`_marca_en_estado`, regla
    única de la hoja de vuelos y de 'cobranza'): TOTAL COTIZACIÓN de
    'cobranza' es `SUMIF(ESTADO, "<>CANCELADO")`; si la marca tapara el
    estado de un cancelado, su cotización entraría a la suma. El API no le
    pone marca a un cancelado; la guarda lo asegura por construcción.
  - **Reparto PDF/XLSX** (`RepartoVueloLinea.cierre_marca`, `diferido_de/a`;
    raíz `cierre` = `RepartoCierre`: `estado`, `cerrado_at`,
    `cerrado_por_nombre`, `version`, `coincide_con_snapshot`, listas
    `diferidos_salen` / `diferidos_entran` / `pendientes_que_pasarian` de
    `{id, folio, cliente, matricula, venta_usd, cobrado_usd, por_cobrar_usd,
    destino, origen, leyenda}`): sello en la portada (PDF) y el subtítulo
    (Excel); la marca bajo el folio del detalle de vuelos (PDF, partida al
    ancho de la columna con `simpleSplit`) o junto al folio (Excel); al
    final de cada avión «Pasan a octubre 2026 (pendientes de cobro, no se
    reparten): #316 Mike Nelson $2,076.40 (cobrado $0.00, por cobrar
    $2,076.40)» / «Vienen de …»; las matrículas sin bloque (externos) en
    «Vuelos del cierre sin avión en este reparto»; en el Excel, sección
    «Cierre del mes» después del detalle de vuelos; `coincide_con_snapshot
    = false` ⇒ AVISO en rojo. El texto del API va ESCAPADO dentro de los
    Paragraph (un «&» o «<» del cliente rompía el marcado).
  - **Libro Dinero** (`DineroVueloFila.cierre_marca`, `diferido_a/de`; raíz
    `cierre` = `BalanceCierre`): la marca junto a la CLAVE (la columna A
    crece por contenido con el tope de siempre) y, al final de 'dinero-vlos',
    «CIERRE DEL MES» con una línea por lista. Un vuelo que PASA no viene en
    `vuelos` del mes de origen (lo dice `cierre.pasan`).
  - **Byte a byte**: `tests/test_cierre_mes.py` congela las huellas
    calculadas con el código de HEAD f0c8ad1 sobre los MISMOS constructores
    (`tests/_cierre_mes_payloads.py`, que solo importa lo que ya existía):
    individual y las dos variantes del general (.xlsx + JSON de la vista),
    reparto PDF (flujos) + XLSX y Libro Dinero — sin las llaves, con ellas
    en neutro (null/false) y con un mes ABIERTO sin vuelos que pasarían (el
    .xlsx idéntico; la vista solo gana la llave `cierre`). Orden de
    despliegue del lote: migración 20261016000004 ⇒ **pyservices** ⇒ API
    0.0.88 ⇒ panel (pyservices antes del API es seguro: el API 0.0.87 no
    manda nada de esto).
  - **Caso #351: «cobrado después», no «pendiente de cobro»** (16-oct-2026,
    API 0.0.91; el cliente: «aparece que aún no se cobran pero sí están
    cobrados»). Los cinco de septiembre (#316, #337, #343, #344, #351)
    pasan a octubre con motivo `COBRO_POSTERIOR`: no estaban cobrados al
    30-sep, pero se cobraron en octubre. Las leyendas siguen siendo del API
    y ya lo dicen; aquí solo cambian los textos PROPIOS, y SOLO cuando el
    item trae `motivo` (sin él, todo byte a byte: las huellas de siempre
    siguen pasando). Campos ADITIVOS con validador liberal
    (`numero_cierre_liberal`: número finito o None, jamás un 422):
    `BalanceCierreVuelo.motivo`, `corte` ('YYYY-MM-DD'),
    `cobrado_despues_usd`, `cobrado_despues_mxn`, `fecha_cobro_despues`;
    `RepartoCierreVuelo.motivo`, `cobrado_posterior_usd`,
    `fecha_cobro_posterior`. Fuente única `app/services/cierre_mes.py`:
    `dia_corto` ('2026-10-04' ⇒ '4-oct'), `corte_del_vuelo` (el `corte` del
    API o, sin él, el último día del mes anterior al destino),
    `es_cobro_posterior`, `detalle_cobro` (las piezas «cobrado al 30-sep $X»
    / «por cobrar al 30-sep $Y» / «cobrado después $Z el 14-oct» o, en un
    COBRO_POSTERIOR sin monto, «cobrado el 14-oct, después del 30-sep»),
    `titulo_lista_reparto` y los títulos con `corte` («Pasan a octubre 2026
    (sin cobro completo al 30-sep: no suman en este mes)», «Pasarán a
    octubre 2026 al cerrar el mes: sin cobro completo al 30-sep (hoy sí
    suman)»; en el reparto «… (sin cobro completo al 30-sep, no se
    reparten)»). Lo usan: las líneas del bloque CIERRE DEL MES del balance
    (encabezados «COBRADO MXN\nAL 30-SEP» / «POR COBRAR MXN\nAL 30-SEP»),
    el Libro Dinero (`linea_vuelo_libro`), el PDF del reparto
    (`_linea_vuelo_cierre`) y el Excel del reparto
    (`ENCABEZADOS_CIERRE_AL_CIERRE`, «Cobrado USD al cierre» / «Por cobrar
    USD al cierre»). Las notas al pie del balance NO cambian. Ojo con el
    skew: el API ≤ 0.0.90 YA manda `motivo` en los items del REPARTO
    (`diferidos_salen` / `diferidos_entran`), así que con este pyservices
    el PDF/Excel del reparto ya rotulan «al 30-sep» y «cobrado después del
    30-sep» antes del API 0.0.91 (es cierto; solo falta la fecha). Orden:
    pyservices ⇒ API 0.0.91 ⇒ panel; todo es aditivo, cualquier orden es
    seguro. Tests: sección 8 de `tests/test_cierre_mes.py` (helpers,
    esquemas liberales, línea del libro por motivo, títulos, bloque del
    balance, Libro Dinero, PDF y Excel del reparto).

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
  3 páginas, bajar N por env antes de tocar código. Desde el 13-oct-2026 los
  bloques con `inicio > 0` llevan además la orden de ignorar la continuación
  huérfana de un SPEI de la página anterior (ver «Prompt del PDF de estado
  de cuenta» en «IA»).
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

- **Comisariato (13-oct-2026, API 0.0.75, migración `20261013000001`).**
  Pedido del cliente: los snacks/bebidas de los PASAJEROS tienen categoría
  propia `COMISARIATO` («Comisariato»: gasto DIRECTO del vuelo, columna
  OTROS del balance del avión; antes se capturaba como OTRO y nunca llegaba
  al avión). Aquí: el `CategoriaGasto` del esquema de visión y `valid_cats`
  de `leer_ticket_gasto` ganan `COMISARIATO` — van JUNTOS: con el valor en
  `valid_cats` y sin él en el Literal, pydantic tumba la lectura entera. El
  prompt `_TICKET_SYSTEM` lo pide para snacks, botanas, bebidas, agua, hielo
  y comisariato de los PASAJEROS (súper/OXXO/comedor cuando el ticket o la
  nota dicen «comisariato», «pax», «pasajeros» o «para el vuelo»); OTRO
  queda para lo que no es ni de la tripulación ni de los pasajeros
  (mensajería, papelería, oficina, varios) y el desempate se conserva: sin
  señas ⇒ TRIPULACIÓN (COMIDA/TAXI/HOTEL); señales de pasajeros ⇒
  COMISARIATO. Los demás esquemas leen `categoria` como `str` (sin cambio) y
  las etiquetas de los libros llegan hechas del API. **NO se tocó** la cola
  «Los servicios FBO sí son costo (columna OTROS)» de los pies del balance
  (`_NOTA_TRASLADOS_*`, `_CPH_TRASLADOS_*`): el diseño la marcaba opcional y
  cambiarla rompe 34 huellas congeladas de libros (cobrado_con, empresa,
  extensión, factura, fórmulas, costo/hora, AFAC); queda incompleta, no
  incorrecta. Orden de deploy: el CONSOLIDADO del lote (ver «Principios»):
  pyservices ANTES del API 0.0.75 (revisión del 14-oct-2026; aquí decía
  «DESPUÉS», y era incompatible con el resto del lote). Es seguro: la app
  publicada (+80) y el panel de prod ignoran una `categoria_sugerida` fuera
  de su catálogo y el API 0.0.74 no pisa la categoría; el APK +81/+82 no
  se publica hasta la migración `20261013000001` y el API 0.0.75. Tests:
  `test_ticket_comisariato_se_entrega` y
  `test_ticket_prompt_pide_comisariato_para_pasajeros` en
  `tests/test_vision_validaciones.py` (el texto vive en el SYSTEM: se lee
  con `cliente.texto_system()`).

- **Pago domiciliado (13-oct-2026, API 0.0.75, migración `20261013000009`;
  el diseño la numeraba 000004).** Medio de pago nuevo `DOMICILIADO`
  («Pago domiciliado»: cargo AUTOMÁTICO a la cuenta —seguros, servicios,
  suscripciones—; bancario, solo oficina, se concilia como una
  transferencia y nunca lleva tarjeta). Aquí NO cambia ningún esquema:
  `medio_pago` es `str` en `GastoCandidato` (conciliación; su `description`
  ya lo nombra) y en `FilaCombustible` (combustible masivo: el catálogo del
  dropdown lo manda el API en `medios_pago`), y la columna PAGO del balance
  llega etiquetada del API. La visión de tickets NO lo sugiere (`medio_pago`
  de ticket y de combustible sigue `Literal[EFECTIVO, TARJETA_CORP,
  TRANSFERENCIA]`: un domiciliado no tiene ticket). Único cambio:
  `_dominio.py` (v2026-10-13) explica las leyendas REALES del estado de
  cuenta «03 POR COBRANZA DOMICILIADA» y «…Carg Recur.» (`COBRANZA
  DOMICILIADA`, `CARG RECUR` en `ALIAS_BANCO`) para la sugerencia de
  conciliación; un candidato DOMICILIADO (sin tarjeta) no recibe el castigo
  «otra tarjeta» aunque la referencia del cargo traiga 20 dígitos. Deploy en
  cualquier orden. Tests: `test_dominio_explica_los_cargos_domiciliados` y
  `test_candidato_domiciliado_viaja_tal_cual_y_sin_castigo_de_tarjeta`
  (`tests/test_sugerir_conciliacion.py`) y la versión del dominio en
  `tests/test_sugerir_abonos.py`.

- **TUA sistémica (13-oct-2026, API 0.0.75; la columna `gasto.tua_incluida`
  es la migración del API `20261013000008`, que el diseño numeraba 000003).**
  Caso #383 (gasto 8c60e428): la IA leyó CORRIDA la tabla resumen de ASUR y
  llamó «Autobuses» a la Tarifa TUA de $559.63; nada se separó y la TUA se
  cargó al avión como operación. Aquí: (1) `GastoTicketResponse.tua_monto`
  ADITIVO en `/vision/gasto` (`float | None`, default None: TUA que cobra el
  documento, IVA incluido, en su moneda); `validaciones_ia.validar_tua_monto`
  (PURA) la limpia: ≤ 0 o ausente ⇒ null sin aviso; > total + 0.02
  (`TOL_SUMA`) ⇒ null + advertencia «La TUA leída (2,000.00) es mayor que el
  total del documento (1,925.27): se descartó.»; si no, a centavos. Dato
  SECUNDARIO (no castiga la confianza) y **NO se vacía cuando
  `validar_suma_conceptos` descarta los renglones** (el panel la enseña como
  sugerencia «La IA leyó una TUA de $X…»). (2) Prompt `_TICKET_SYSTEM`:
  PRIORIDAD 1 de «conceptos» con las secciones canónicas EXACTAS
  (Operaciones, Taxi, Autobuses, Tarifa TUA, Combustible, FOB y Migración),
  etiqueta y Total de la MISMA FILA, la sección de la TUA SIEMPRE «Tarifa
  TUA» y la autocomprobación «Total en Taxi/Autobuses sin servicio de taxi o
  autobús y con TUA ⇒ etiqueta corrida»; clave nueva `"tua_monto"` entre
  «conceptos» y «matricula» (210100/210200/210300 NO son TUA; 130700/230700
  SÍ; «con duda, null»). PRIORIDAD 2, EXCEPCIÓN, claves y descuentos sin
  cambio. **La separación la decide el API** (`common/desglose-gasto.util.ts`:
  oficina `tua_incluida` > renglón TUA reconocido > `tua_monto` SOLO con una
  tabla que cuadra): aquí no se cruzan renglones ni se copia `esTua`, y
  reportes/payloads del balance no cambian. Un API previo ignora la llave
  (viaja dentro de `valor_ia_extraido` y el 0.0.75 la aprovecha al
  reanalizar). Deploy en cualquier orden (ADITIVO). Tests:
  `test_validar_tua_monto` (`tests/test_validaciones_ia.py`) y, en
  `tests/test_vision_validaciones.py`, `test_ticket_tua_monto_se_copia`,
  `…_mayor_que_el_total_se_descarta`, `test_ticket_sin_tua_monto_sigue_igual`,
  `test_ticket_caso_383_tabla_resumen_corrida`,
  `test_ticket_tua_monto_sobrevive_si_se_descartan_conceptos` y
  `test_prompt_tabla_resumen_y_tua_monto`.

- **Folio anti-duplicados de tickets de ESTANCIA del aeropuerto (2-oct-2026).**
  ASUR imprime «COBRO ESTANCIA · Boleto FBOE7»: el «Boleto» es la TARJETA de
  acceso y se repite en cada visita; la IA lo tomaba como `folio` y el candado
  anti-duplicados del API (folio_ticket único) rechazaba el ticket legítimo.
  El prompt de `leer_ticket_gasto` pide además `boleto`, `numero_operacion`
  y `numero_terminal`, y `validaciones_ia.folio_anti_duplicado` (PURA, con
  test) decide el folio: con operación ⇒ «<terminal>-<operación>»
  («1-094-4368») cuando el folio falta, es el boleto o es la operación
  pelona; un folio real distinto se conserva; boleto sin operación ⇒ folio
  null + aviso (nunca bloquear la captura). Un folio corto como «4368»
  NO va solo: chocaría con folios de otros proveedores.

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
  **Concepto del SPEI (13-oct-2026, API 0.0.75)**: caso EJE SCOTIA de
  septiembre — 73 de 98 filas quedaron como «SEL TRANSF. INTERBANCARIA
  SPEI» / «TRANSF INTERBANCARIA SPEI» SIN el concepto que el banco imprime
  debajo («BANORTE» / «fact vip 7360») y la oficina no sabía qué factura
  pagó cada una. El `_SYSTEM` exige los RENGLONES DE CONTINUACIÓN (banco
  destino u origen, beneficiario u ordenante y el concepto tecleado) pegados
  al renglón principal —TAL CUAL y PRIMERO, separados por un espacio, en
  TODAS las transferencias y pagos—, sin CLABE, cuentas/tarjetas, clave de
  rastreo, RFC ni hora, hasta 160 caracteres (se quitó el viejo «córtala a
  80»); el formato de Scotiabank nombra sus leyendas con continuación y
  `_prompt_pdf` lo recuerda. La orden de IGNORAR la continuación huérfana con
  que puede empezar un bloque va SOLO en `_nota_paginas` con `inicio > 0`
  (`_NOTA_CONTINUACION_HUERFANA`; un PDF de ≤ N páginas sigue siendo UNA
  llamada sin ella): ese movimiento lo lee el bloque anterior y se queda sin
  concepto (riesgo aceptado: lo delata el aviso de abajo). Código:
  `_descripcion_limpia` deja la descripción en UNA línea con tope
  `_DESCRIPCION_MAX` = 200 (vacía o solo espacios ⇒ None; CSV, Excel y
  Paywise no cambian) y `_advertencia_sin_concepto` —PRIMERA de
  `_advertencias_pdf`, dentro del tope de 10— cuenta las filas cuya
  descripción en forma `_plano` (= `normalizarPlano` del API: NFD sin marcas
  U+0300–U+036F, mayúsculas, `[^A-Z0-9]+` → espacio) es una de las 9
  `_LEYENDAS_SIN_CONCEPTO` (las MISMAS de `LEYENDAS_SIN_CONCEPTO` del API;
  `test_leyendas_paridad_con_el_api_repo_hermano` lee
  `vuelatour-api/src/modules/conciliacion/enriquecer-descripcion.util.ts` y
  se SALTA sin el repo hermano, como en Railway):
  «N transferencia(s) o pago(s) se leyeron SIN su concepto (solo la leyenda
  del banco, p. ej. «SEL TRANSF. INTERBANCARIA SPEI») …». El esquema NO
  cambia. El API completa la descripción de las filas YA importadas al
  re-subir el archivo (vista previa, «Solo completar conceptos», reglas que
  nunca adivinan y UPDATE con CAS; invariante «concepto del SPEI» del
  CLAUDE.md del API) y su dedupe NO depende de la referencia (los SPEI traen
  la de relleno «00000000000000000001»): una relectura con concepto no
  duplica. Deploy en cualquier orden respecto al API (nada se rompe ni se
  cuenta dos veces); debe estar arriba ANTES de re-subir septiembre. Tests:
  `tests/test_estado_cuenta_concepto_spei.py` y
  `test_continuacion_huerfana_solo_en_los_bloques_sin_la_pagina_1` en
  `tests/test_estado_cuenta_pdf_bloques.py`.
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
  `valido=false` + `motivo` en vez de un objeto a medias. Desde el
  5-oct-2026 devuelve también `serie`/`folio` (atributos del Comprobante,
  recortados, vacío ⇒ None; `serie`/`folio` en minúscula en los 3.2): el
  número de factura de la columna «Notas» del Excel de conciliación.
  **Las llaves `serie`/`folio` SIEMPRE se emiten (null si faltan)**: el cron
  de relectura del API (`recibidas-releer-folio`) usa su PRESENCIA para
  sellar `folio_releido_at` — nunca `exclude_none` en
  `/facturacion/parse-recibida` (sin ellas relee los mismos XML cada 10 min
  sin fin). Tests: `tests/test_recibida_serie_folio.py` (parser y ruta).
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
