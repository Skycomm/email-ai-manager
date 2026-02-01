"""
API module for web dashboard.
"""

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pathlib import Path

from .routes import router, set_db


def create_app(db=None) -> FastAPI:
    """Create and configure the FastAPI application."""
    app = FastAPI(
        title="Email AI Manager",
        description="AI-powered email management dashboard",
        version="0.1.0",
    )

    # CORS for local development
    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )

    # Include API routes
    app.include_router(router, prefix="/api")

    # Serve static files for dashboard
    dashboard_path = Path(__file__).parent.parent.parent / "dashboard"
    if dashboard_path.exists():
        app.mount("/", StaticFiles(directory=str(dashboard_path), html=True), name="dashboard")

    # Set database if provided
    if db:
        set_db(db)

    return app


__all__ = ["create_app", "router", "set_db"]
