from fastapi import APIRouter, Depends, HTTPException, Query
from sqlalchemy.orm import Session
from typing import Optional
import logging

from app.database import get_db
from app.auth import get_current_user
from app.models import User
from app.services.po_service import POService

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/api", tags=["purchase-orders"])


@router.get("/po-data")
async def get_po_data(
    page: int = Query(1, ge=1, description="Page number"),
    per_page: int = Query(50, ge=1, le=1000, description="Items per page (max 1000)"),
    project_name: Optional[str] = Query(None, description="Filter by project name"),
    po_status: Optional[str] = Query(None, description="Filter by PO status"),
    search: Optional[str] = Query(None, description="Search in PO number or item description"),
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    """Get paginated PO data with optional filters"""
    try:
        po_service = POService(db)
        user_id = str(current_user.id)
        
        return po_service.get_po_data(
            user_id=user_id,
            page=page,
            per_page=per_page,
            project_name=project_name,
            po_status=po_status,
            search=search
        )
    except Exception as e:
        logger.error(f"Error in get_po_data: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error fetching PO data: {str(e)}")