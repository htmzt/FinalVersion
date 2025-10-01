from fastapi import APIRouter, Depends, HTTPException, Query
from fastapi.responses import StreamingResponse
from sqlalchemy.orm import Session
from typing import Optional
from io import BytesIO
import pandas as pd
import logging

from app.database import get_db
from app.auth import get_current_user
from app.models import User
from app.services.summary_service import SummaryService

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/api/summary", tags=["summary"])


@router.get("/monthly")
async def get_monthly_summary(
    year: Optional[int] = Query(None, description="Filter by year (e.g., 2024)"),
    month: Optional[int] = Query(None, ge=1, le=12, description="Filter by month (1-12)"),
    project_name: Optional[str] = Query(None, description="Filter by project name"),
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    """Get monthly summary by project name with optional filters"""
    try:
        summary_service = SummaryService(db)
        user_id = str(current_user.id)
        
        return summary_service.get_monthly_summary(
            user_id=user_id,
            year=year,
            month=month,
            project_name=project_name
        )
    except Exception as e:
        logger.error(f"Error in get_monthly_summary: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error fetching monthly summary: {str(e)}")


@router.get("/yearly")
async def get_yearly_summary(
    project_name: Optional[str] = Query(None, description="Filter by project name"),
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    """Get yearly summary aggregated by project and year"""
    try:
        summary_service = SummaryService(db)
        user_id = str(current_user.id)
        
        return summary_service.get_yearly_summary(
            user_id=user_id,
            project_name=project_name
        )
    except Exception as e:
        logger.error(f"Error in get_yearly_summary: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error fetching yearly summary: {str(e)}")


@router.get("/periods")
async def get_available_periods(
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    """Get available years and months for filtering"""
    try:
        summary_service = SummaryService(db)
        user_id = str(current_user.id)
        
        return summary_service.get_available_periods(user_id)
    except Exception as e:
        logger.error(f"Error in get_available_periods: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error fetching available periods: {str(e)}")


@router.get("/projects")
async def get_project_list(
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    """Get list of available project names for filtering"""
    try:
        summary_service = SummaryService(db)
        user_id = str(current_user.id)
        
        projects = summary_service.get_project_list(user_id)
        return {"projects": projects}
    except Exception as e:
        logger.error(f"Error in get_project_list: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error fetching project list: {str(e)}")


@router.get("/export")
async def export_summary_data(
    year: Optional[int] = Query(None, description="Filter by year"),
    month: Optional[int] = Query(None, ge=1, le=12, description="Filter by month (1-12)"),
    project_name: Optional[str] = Query(None, description="Filter by project name"),
    summary_type: str = Query("monthly", description="Type of summary: 'monthly' or 'yearly'"),
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    """Export summary data to Excel"""
    try:
        summary_service = SummaryService(db)
        user_id = str(current_user.id)
        
        if summary_type == "yearly":
            data = summary_service.get_yearly_summary(user_id, project_name)
            filename = "yearly_summary.xlsx"
        else:
            result = summary_service.get_monthly_summary(user_id, year, month, project_name)
            data = result["summaries"]
            filename = "monthly_summary.xlsx"
        
        if not data:
            raise HTTPException(status_code=404, detail="No data found to export")
        
        # Flatten the nested data structure for Excel export
        flattened_data = []
        for item in data:
            flat_item = {
                'Project Name': item.get('project_name'),
                'Year': item.get('year'),
                'Month': item.get('month'),
                'Month Year': item.get('month_year'),
                'Total Records': item.get('total_records'),
                'Unique POs': item.get('unique_pos'),
                
                # Financial data
                'Total Line Amount': item.get('financial_summary', {}).get('total_line_amount', 0),
                'Total AC Amount': item.get('financial_summary', {}).get('total_ac_amount', 0),
                'Total PAC Amount': item.get('financial_summary', {}).get('total_pac_amount', 0),
                'Total Remaining Amount': item.get('financial_summary', {}).get('total_remaining_amount', 0),
                
                # Status data
                'Closed Count': item.get('status_breakdown', {}).get('closed', 0),
                'Cancelled Count': item.get('status_breakdown', {}).get('cancelled', 0),
                'Pending Count': item.get('status_breakdown', {}).get('pending', 0),
                'Completion Rate %': item.get('status_breakdown', {}).get('completion_rate', 0),
                
                # Payment terms
                'ACPAC 100% Count': item.get('payment_terms_breakdown', {}).get('acpac_100_percent', 0),
                'AC/PAC Split Count': item.get('payment_terms_breakdown', {}).get('ac_pac_split', 0),
                
                # Categories
                'Survey Count': item.get('category_breakdown', {}).get('survey', 0),
                'Transportation Count': item.get('category_breakdown', {}).get('transportation', 0),
                'Site Engineer Count': item.get('category_breakdown', {}).get('site_engineer', 0),
                'Service Count': item.get('category_breakdown', {}).get('service', 0),
            }
            flattened_data.append(flat_item)
        
        df = pd.DataFrame(flattened_data)
        
        # Create Excel file in memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Summary Data', index=False)
            
            # Add a second sheet with overall totals if available
            if summary_type == "monthly" and "overall_totals" in result:
                totals_data = [{
                    'Metric': 'Total Records',
                    'Value': result["overall_totals"]["total_records"]
                }, {
                    'Metric': 'Unique POs',
                    'Value': result["overall_totals"]["unique_pos"]
                }, {
                    'Metric': 'Unique Projects',
                    'Value': result["overall_totals"]["unique_projects"]
                }, {
                    'Metric': 'Total Line Amount',
                    'Value': result["overall_totals"]["financial_totals"]["total_line_amount"]
                }, {
                    'Metric': 'Total AC Amount',
                    'Value': result["overall_totals"]["financial_totals"]["total_ac_amount"]
                }, {
                    'Metric': 'Total PAC Amount',
                    'Value': result["overall_totals"]["financial_totals"]["total_pac_amount"]
                }, {
                    'Metric': 'Overall Completion Rate %',
                    'Value': result["overall_totals"]["overall_completion_rate"]
                }]
                
                totals_df = pd.DataFrame(totals_data)
                totals_df.to_excel(writer, sheet_name='Overall Totals', index=False)
        
        output.seek(0)
        
        return StreamingResponse(
            BytesIO(output.getvalue()),
            media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
        
    except Exception as e:
        logger.error(f"Error exporting summary data: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error exporting summary data: {str(e)}")