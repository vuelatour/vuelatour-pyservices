# pyservices (FastAPI) — incluye libs de sistema para WeasyPrint (pango/cairo).
FROM python:3.12-slim

# Dependencias de sistema de WeasyPrint (PDF) + libffi.
RUN apt-get update && apt-get install -y --no-install-recommends \
    libpango-1.0-0 \
    libpangocairo-1.0-0 \
    libcairo2 \
    libgdk-pixbuf-2.0-0 \
    libffi-dev \
    shared-mime-info \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Railway inyecta $PORT. Host :: (IPv6) para la red privada de Railway.
CMD ["sh", "-c", "uvicorn app.main:app --host :: --port ${PORT:-8000}"]
