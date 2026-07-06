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
  (están por índice de columna).
- El reporte por vuelo debe CUADRAR: el desglose (subtotal + TUAS + pernocta
  + extras + ajuste + IVA) suma el total exacto — no omitir líneas del
  desglose canónico v1.3.

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
