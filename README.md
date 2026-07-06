# vuelatour-pyservices

Microservicio Python (FastAPI) de **VuelaTour**: visión con IA (Claude),
generación de PDF (WeasyPrint) y Excel (openpyxl), timbrado CFDI (FEL vía
satcfdi/zeep), parseo de estados de cuenta y conciliación asistida.
Desplegado en **Railway** (deploy automático al hacer push a `main`);
lo consume únicamente `vuelatour-api` (nunca los clientes).

> Convenciones para desarrollo: **CLAUDE.md**.

## Requisitos

- Python **3.12** (el `Dockerfile` usa `python:3.12-slim`; el código usa
  sintaxis `X | None`, no corre en 3.9).

## Setup

```bash
python3.12 -m venv .venv && source .venv/bin/activate
pip install -r requirements-dev.txt
cp .env.example .env.local
```

Variables (`.env.local`, nunca commiteadas):

- `ANTHROPIC_API_KEY` · `ANTHROPIC_MODEL` (visión/IA).
- `INTERNAL_SHARED_TOKEN` — debe coincidir con el del API NestJS; todas las
  rutas funcionales exigen el header `X-Internal-Token`.
- `FEL_USUARIO`, `FEL_PASSWORD`, `FEL_WSDL_URL`, `FEL_MODO` (CFDI).

## Run / Test

```bash
uvicorn app.main:app --reload --port 8000   # /docs = Swagger, /health = check
pytest && ruff check app tests
```

## Endpoints

| Ruta | Qué hace |
|---|---|
| `GET /health` | Healthcheck |
| `POST /vision/tacometro` | Lee el horómetro (HOBBS) de una foto con Claude Vision (usa el último taco del avión como ancla de magnitud) |
| `POST /vision/gasto` | Extrae monto/fecha/proveedor/categoría de un ticket |
| `POST /vision/combustible` | Litros / precio-litro / total / aeropuerto |
| `POST /compras/extraer` | Líneas de producto de un PDF de Aircraft Spruce |
| `POST /vencimientos/extraer` | Matrícula/tipo/fechas de pólizas y tarjetas |
| `POST /conciliacion/parse` | Parsea estado de cuenta (CSV/Excel/PDF) |
| `POST /conciliacion/sugerir` | Propone el gasto más probable para un cargo ambiguo |
| `POST /reportes/cotizacion` | PDF de cotización |
| `POST /pdf/reparto` · `/pdf/reparto-xlsx` | Reparto de utilidades (PDF socios / Excel mensual por avión, con horas voladas y avisos de integridad) |
| `POST /pdf/reporte-vuelo` · `/pdf/reporte-vuelo-xlsx` | Reporte consolidado de UN vuelo (desglose con pernocta, manifiesto de pasajeros, horas cotizadas vs voladas, cobros, combustible, gastos) |
| `POST /pdf/tabla-xlsx` | Utilitario genérico tabla→Excel (exports de operación) |
| `POST /pdf/zip` | Ensambla el paquete de cierre (.zip) |
| `POST /facturacion/*` | Timbrado CFDI 4.0, cancelación, nota de crédito, parseo de CFDI recibidos |

Todos los payloads los arma `vuelatour-api` (`PyservicesService`); aquí solo
se validan (pydantic) y se renderiza/infiere.
