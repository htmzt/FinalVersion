from fastapi import APIRouter, Depends, HTTPException, Query
from sqlalchemy.orm import Session
from typing import Optional
import logging

from app.database import get_db
from app.auth import get_current_user
from app.models import User
from app.services.acceptance_service import AcceptanceService

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/api", tags=["acceptance"])


@router.get("/acceptance-data")
async def get_acceptance_data(
    page: int = Query(1, ge=1, description="Page number"),
    per_page: int = Query(50, ge=1, le=1000, description="Items per page (max 1000)"),
    status: Optional[str] = Query(None, description="Filter by status"),
    project_name: Optional[str] = Query(None, description="Filter by project name"),
    search: Optional[str] = Query(None, description="Search in acceptance number or PO number"),
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    """Get paginated Acceptance data with optional filters"""
    try:
        acceptance_service = AcceptanceService(db)
        user_id = str(current_user.id)
        
        return acceptance_service.get_acceptance_data(
            user_id=user_id,
            page=page,
            per_page=per_page,
            status=status,
            project_name=project_name,
            search=search
        )
    except Exception as e:
        logger.error(f"Error in get_acceptance_data: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error fetching acceptance data: {str(e)}")