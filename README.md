# vuelatour-pyservices

Vuelatour Python microservices built with FastAPI.

## Requirements

- Python 3.12+

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements-dev.txt
```

## Run

```powershell
uvicorn app.main:app --reload --port 8000
```

Open http://localhost:8000 and http://localhost:8000/docs

## Configuración

Copia `.env.example` a `.env.local` y completa:

- `ANTHROPIC_API_KEY` — key de Claude (solo en `.env.local`, nunca se commitea).
- `ANTHROPIC_MODEL` — por defecto `claude-sonnet-4-6`.
- `INTERNAL_SHARED_TOKEN` — token compartido con NestJS; debe coincidir en ambos
  servicios. Genera uno con `python -c "import secrets; print(secrets.token_urlsafe(32))"`.

## Endpoints

- `POST /vision/tacometro` — lee el horómetro/tacómetro (HOBBS) de una foto con
  Claude Vision. Requiere header `X-Internal-Token`. Body: `image_base64` +
  `media_type`, o `image_url`. Responde `{lectura, confianza, legible, notas, modelo}`.

## Test

```powershell
pytest
```
