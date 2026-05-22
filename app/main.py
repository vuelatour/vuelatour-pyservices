from fastapi import FastAPI

from app.routers import vision

app = FastAPI(
    title="vuelatour-pyservices",
    version="0.1.0",
)

app.include_router(vision.router)


@app.get("/")
def read_root() -> dict[str, str]:
    return {"service": "vuelatour-pyservices", "status": "ok"}


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "healthy"}
