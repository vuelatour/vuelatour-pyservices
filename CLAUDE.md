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

## IA

- Visión (tacómetro/tickets) usa el modelo de `ANTHROPIC_MODEL`. La lectura
  de tacómetro recibe `ultimo` (último taco del avión) como ancla de
  magnitud; conservar ese contrato.
- Todo punto de IA degrada a captura manual: los errores devuelven
  `legible=false`/`disponible=false`, nunca 500 por fallo del modelo.

## Entorno

- Python **3.12** obligatorio (sintaxis `X | None`); el python de sistema de
  esta Mac es 3.9 y NO corre el código — validar con `python3 -m ast` /
  tests en CI o Docker.
- `pytest` + `ruff check app tests` antes de commit. Push a `main` = deploy
  automático en Railway (autorizado sin preguntar).
- `ANTHROPIC_API_KEY` solo en `.env.local` / variables de Railway. Nunca en
  el repo (ya hubo una key expuesta; está pendiente rotarla).
