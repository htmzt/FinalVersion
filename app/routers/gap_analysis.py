from fastapi import APIRouter, Depends, HTTPException, Query
from fastapi.responses import StreamingResponse
from sqlalchemy.orm import Session
from typing import Optional
from app.database import get_db
from app.auth import get_current_user
from app.models import User
from app.services.gap_analysis_service import GapAnalysisService
from io import BytesIO
import logging

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/gap-analysis", tags=["gap-analysis"])

@router.get("/financial-summary")
async def get_gap_financial_summary(
    project_name: Optional[str] = Query(None, description="Filter by project name"),
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    """
    Get gap financial analysis summary by project in tabular format
    
    Returns:
    - Total PO Received (line amounts)
    - GAP PO OK; AC NOK (AC pending amounts) 
    - GAP AC OK; PAC NOK (PAC pending amounts)
    - Total GAP AC & PAC (combined gap amounts)
    - Gap Percentage by Project
    """
    try:
        service = GapAnalysisService(db)
        financial_summary = service.get_gap_financial_summary_by_project(str(current_user.id), project_name)
        
        return {
            "success": True,
            "data": {
                "financial_summary": financial_summary,
                "column_headers": [
                    "Project Name",
                    "Total PO Received", 
                    "GAP PO OK; AC NOK",
                    "GAP AC OK; PAC NOK", 
                    "Total GAP AC & PAC",
                    "Gap Percentage"
                ]
            }
        }
        
    except Exception as e:
        logger.error(f"Error in gap financial summary endpoint: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error retrieving financial summary: {str(e)}")


@router.get("/export/gap-financial-summary")
async def export_gap_financial_summary_excel(
    project_name: Optional[str] = Query(None, description="Filter by project name"),
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    """
    Export gap financial summary directly to Excel (optimized for performance)
    
    Creates an Excel file with gap financial analysis by project:
    - GAP by Project
    - Total PO Received 
    - GAP PO Ok; AC Nok
    - GAP AC OK; PAC Nok
    - Total GAP AC & PAC
    - Pourcentage GAP Par Projet
    
    This endpoint is optimized for speed by calculating summaries directly in SQL.
    """
    try:
        service = GapAnalysisService(db)
        excel_data = service.export_gap_financial_summary_to_excel(str(current_user.id), project_name)
        
        # Create filename
        filename = "gap_financial_summary"
        if project_name:
            safe_project_name = "".join(c for c in project_name if c.isalnum() or c in (' ', '-', '_')).strip()
            filename += f"_{safe_project_name}"
        filename += ".xlsx"
        
        # Return as streaming response
        return StreamingResponse(
            BytesIO(excel_data),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
        
    except Exception as e:
        logger.error(f"Error in gap financial summary Excel export: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error exporting gap financial summary: {str(e)}")