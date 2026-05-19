# vuelatour-pyservices

Microservicio Python (FastAPI) de Vuelatour: generación de PDFs y, a futuro,
parseo de Excel, conciliación bancaria con IA y migración de datos.

## Requisitos

- Python 3.12+

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements-dev.txt
copy .env.example .env
```

Edita `.env` y define `SERVICE_TOKEN` con un valor seguro. Ese mismo valor debe
ir en `PYSERVICES_TOKEN` del API NestJS.

## Run

```powershell
uvicorn app.main:app --reload --port 8000
```

- http://localhost:8000/docs — Swagger
- http://localhost:8000/health — healthcheck

## Test

```powershell
pytest
ruff check app tests
```

## Endpoints

| Método | Ruta | Descripción |
|--------|------|-------------|
| GET | `/` | Info del servicio |
| GET | `/health` | Healthcheck |
| POST | `/pdf/cotizacion` | Genera el PDF de una cotización (requiere `X-Service-Token`) |

## Autenticación entre servicios

Las rutas funcionales exigen el header `X-Service-Token`, que debe coincidir con
`SERVICE_TOKEN`. NestJS lo envía automáticamente desde `PyservicesService`.

## Pendiente (roadmap del doc funcional)

- Reportes PDF a socios y mensual por avión.
- Parseo de Excel (openpyxl) y migración de los 9 archivos históricos.
- Conciliación bancaria (pandas + Claude API).
- Archivos de importación para CONTPAQi.
