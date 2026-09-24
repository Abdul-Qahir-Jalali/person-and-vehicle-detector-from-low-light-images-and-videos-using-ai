from fastapi import FastAPI
from app.core.config import settings
from app.api.routes import detection

from fastapi.staticfiles import StaticFiles
from fastapi.responses import RedirectResponse
import os

app = FastAPI(
    title=settings.PROJECT_NAME,
    description="A robust modular object detection system for detecting persons and vehicles.",
    version="1.0.0"
)

app.include_router(detection.router, prefix=settings.API_V1_STR, tags=["detection"])

# Mount static files
static_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")
os.makedirs(static_dir, exist_ok=True)
app.mount("/static", StaticFiles(directory=static_dir), name="static")

@app.get("/")
def read_root():
    return RedirectResponse(url="/static/index.html")
