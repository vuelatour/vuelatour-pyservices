import os

from fastapi import FastAPI

from app.routers import (
    compras,
    conciliacion,
    facturacion,
    gastos,
    ia,
    inventario,
    pdf,
    reportes,
    vencimientos,
    vision,
)

app = FastAPI(
    title="vuelatour-pyservices",
    version="0.1.0",
)

app.include_router(vision.router)
app.include_router(facturacion.router)
app.include_router(reportes.router)
app.include_router(compras.router)
app.include_router(conciliacion.router)
app.include_router(gastos.router)
app.include_router(inventario.router)
app.include_router(vencimientos.router)
app.include_router(pdf.router)
app.include_router(ia.router)


@app.get("/")
def read_root() -> dict[str, str]:
    return {"service": "vuelatour-pyservices", "status": "ok"}


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "healthy"}


@app.get("/version")
def version() -> dict[str, str | None]:
    """Commit DESPLEGADO (14-oct-2026, API 0.0.77; revisión F7): Railway pone
    `RAILWAY_GIT_COMMIT_SHA` en el entorno del deployment y con esta ruta la
    COMPUERTA del despliegue («el deployment ACTIVO de pyservices es el commit
    nuevo» antes de empujar el API) se verifica con curl. Sin token, como
    `/health`: solo dice qué código corre, ningún dato del negocio. Fuera de
    Railway (local, tests) ⇒ `commit` null."""
    return {
        "commit": os.environ.get("RAILWAY_GIT_COMMIT_SHA", "").strip() or None,
        "servicio": "vuelatour-pyservices",
    }
