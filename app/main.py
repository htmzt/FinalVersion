from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
import asyncio
import logging
from concurrent.futures import ThreadPoolExecutor
from sqlalchemy.orm import Session

# Database and models
from app.database import engine, SessionLocal
from app import models

# Routers
from app.routers import (
    auth,
    files,
    dashboard,
    po,
    acceptance,
    accounts,
    merged_data,
    summary,
    gap_analysis
)

# Background task utilities
from app.tasks import task_worker, task_queue, task_workers

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Create database tables
models.Base.metadata.create_all(bind=engine)

# Initialize FastAPI app
app = FastAPI(
    title="PO Management API",
    version="1.0.0",
    description="Purchase Order and Acceptance Management System"
)

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Thread pool for background processing
thread_pool = ThreadPoolExecutor(max_workers=4)


# Root endpoint
@app.get("/")
async def root():
    return {
        "message": "SIB API is running",
        "version": "1.0.0",
        "endpoints": {
            "auth": "/login",
            "uploads": "/api/upload",
            "dashboard": "/api/dashboard-analytics",
            "po_data": "/api/po-data",
            "acceptance": "/api/acceptance-data",
            "merged_data": "/api/merged-data",
            "summary": "/api/summary",
            "gap_analysis": "/gap-analysis"
        }
    }


# Include all routers
app.include_router(auth.router)
app.include_router(files.router)
app.include_router(dashboard.router)
app.include_router(po.router)
app.include_router(acceptance.router)
app.include_router(accounts.router)
app.include_router(merged_data.router)
app.include_router(summary.router)
app.include_router(gap_analysis.router)




@app.on_event("startup")
async def startup_event():
    """Initialize startup tasks"""
    db = SessionLocal()
    try:
        models.Base.metadata.create_all(bind=engine)
    except Exception as e:
        logger.error(f"Error during startup: {str(e)}")
    finally:
        db.close()
    
    # Start background workers
    for i in range(2):
        worker = asyncio.create_task(task_worker())
        task_workers.append(worker)
    
    logger.info("PO Management API started successfully with 2 background workers")


@app.on_event("shutdown")
async def shutdown_event():
    """Cleanup on shutdown"""
    # Signal workers to shutdown
    for _ in range(len(task_workers)):
        await task_queue.put(("shutdown", None))
    
    # Wait for all tasks to complete
    await task_queue.join()
    
    # Cancel worker tasks
    for worker in task_workers:
        worker.cancel()
    
    # Wait for workers to finish
    await asyncio.gather(*task_workers, return_exceptions=True)
    
    thread_pool.shutdown(wait=True)
    logger.info("PO Management API shutting down")