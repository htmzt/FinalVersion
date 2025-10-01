from fastapi import APIRouter, Depends
from sqlalchemy.orm import Session

from app.database import get_db
from app.auth import get_current_user
from app.models import User
from app.services.dashboard_service import DashboardService

router = APIRouter(prefix="/api", tags=["dashboard"])


@router.get("/data-status")
async def get_data_status(
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    """Check if both PO and Acceptance data exist"""
    dashboard_service = DashboardService(db)
    user_id = str(current_user.id)
    
    return dashboard_service.get_data_status(user_id)


@router.get("/dashboard-analytics")
async def get_dashboard_analytics(
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    """Get comprehensive dashboard analytics"""
    dashboard_service = DashboardService(db)
    user_id = str(current_user.id)
    
    return dashboard_service.get_dashboard_analytics(user_id)


@router.get("/charts-data")
async def get_charts_data(
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    """Get structured data for React charts"""
    dashboard_service = DashboardService(db)
    user_id = str(current_user.id)
    
    return dashboard_service.get_charts_data(user_id)