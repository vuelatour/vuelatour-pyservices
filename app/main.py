from fastapi import FastAPI

from app.routers import compras, facturacion, reportes, vision

app = FastAPI(
    title="vuelatour-pyservices",
    version="0.1.0",
)

app.include_router(vision.router)
app.include_router(facturacion.router)
app.include_router(reportes.router)
app.include_router(compras.router)


@app.get("/")
def read_root() -> dict[str, str]:
    return {"service": "vuelatour-pyservices", "status": "ok"}


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "healthy"}
