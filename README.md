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

## Test

```powershell
pytest
```
