from fastapi import APIRouter, Depends, HTTPException, Query
from fastapi.responses import StreamingResponse
from sqlalchemy.orm import Session
from sqlalchemy import text
from typing import Optional
from io import BytesIO
import pandas as pd
import logging

from app.database import get_db
from app.auth import get_current_user
from app.models import User
from app.query import MERGED_DATA_QUERY

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/api/merged-data", tags=["merged-data"])


@router.get("")
async def get_merged_data(
    page: int = Query(1, ge=1, description="Page number"),
    per_page: int = Query(50, ge=1, le=1000, description="Items per page (max 1000)"),
    status: Optional[str] = Query(None, description="Filter by status"),
    category: Optional[str] = Query(None, description="Filter by category"),
    project_name: Optional[str] = Query(None, description="Filter by project name"),
    search: Optional[str] = Query(None, description="Search in PO number or item description"),
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    """Get paginated merged PO and Acceptance data"""
    try:
        user_id = str(current_user.id)
        
        # Build filter conditions
        filter_conditions = ["po.user_id = :user_id"]
        params = {"user_id": user_id}
        
        if project_name:
            filter_conditions.append("po.project_name ILIKE :project_name")
            params["project_name"] = f"%{project_name}%"
            
        if search:
            filter_conditions.append("(po.po_number ILIKE :search OR po.item_description ILIKE :search)")
            params["search"] = f"%{search}%"
        
        where_clause = " AND ".join(filter_conditions)
        base_query = MERGED_DATA_QUERY.format(base_filter=where_clause)
        
        # Apply additional filters
        if status or category:
            filter_subquery = f"SELECT * FROM ({base_query}) as subquery WHERE 1=1"
            if status:
                filter_subquery += " AND subquery.status = :status"
                params["status"] = status
            if category:
                filter_subquery += " AND subquery.category = :category"  
                params["category"] = category
            base_query = filter_subquery
        
        # Count query
        count_query = text(f"SELECT COUNT(*) as total FROM ({base_query}) as count_subquery")
        count_result = db.execute(count_query, params).scalar()
        total_count = count_result or 0
        
        if total_count == 0:
            return {
                "items": [],
                "total_count": 0,
                "page": page,
                "per_page": per_page,
                "total_pages": 0,
                "has_next": False,
                "has_prev": False
            }
        
        # Data query with pagination
        offset = (page - 1) * per_page
        data_query_str = f"{base_query} ORDER BY po_no DESC, po_line ASC LIMIT :limit OFFSET :offset"
        params.update({"limit": per_page, "offset": offset})
        
        data_query = text(data_query_str)
        result = db.execute(data_query, params)
        merged_data = result.mappings().all()
        
        # Calculate pagination info
        total_pages = (total_count + per_page - 1) // per_page
        
        return {
            "items": merged_data,
            "total_count": total_count,
            "page": page,
            "per_page": per_page,
            "total_pages": total_pages,
            "has_next": page < total_pages,
            "has_prev": page > 1
        }
        
    except Exception as e:
        logger.error(f"Error querying merged data: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error querying merged data: {str(e)}")


@router.get("/export")
async def export_merged_data(
    status: Optional[str] = Query(None, description="Filter by status"),
    category: Optional[str] = Query(None, description="Filter by category"),
    project_name: Optional[str] = Query(None, description="Filter by project name"),
    search: Optional[str] = Query(None, description="Search in PO number or item description"),
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    """Export filtered merged PO and Acceptance data to Excel file"""
    try:
        user_id = str(current_user.id)
        
        # Build filter conditions
        filter_conditions = ["po.user_id = :user_id"]
        params = {"user_id": user_id}
        
        if project_name:
            filter_conditions.append("po.project_name ILIKE :project_name")
            params["project_name"] = f"%{project_name}%"
            
        if search:
            filter_conditions.append("(po.po_number ILIKE :search OR po.item_description ILIKE :search)")
            params["search"] = f"%{search}%"
        
        where_clause = " AND ".join(filter_conditions)
        base_query = MERGED_DATA_QUERY.format(base_filter=where_clause)
        
        # Apply additional filters
        if status or category:
            filter_subquery = f"SELECT * FROM ({base_query}) as subquery WHERE 1=1"
            if status:
                filter_subquery += " AND subquery.status = :status"
                params["status"] = status
            if category:
                filter_subquery += " AND subquery.category = :category"  
                params["category"] = category
            base_query = filter_subquery
        
        # Execute query
        query = text(f"{base_query} ORDER BY po_no, po_line")
        result = db.execute(query, params)
        merged_data = result.mappings().all()
        
        if not merged_data:
            raise HTTPException(status_code=404, detail="No data found to export")
        
        # Convert to pandas DataFrame
        data_dict = []
        for item in merged_data:
            data_dict.append({
                'PO ID': item.po_id,
                'PO Number': item.po_no,
                'PO Line': item.po_line,
                'Account': item.account_name,
                'Project': item.project_name,
                'Site Code': item.site_code,
                'Category': item.category,
                'Item Description': item.item_desc,
                'Payment Terms': item.payment_terms,
                'Unit Price': float(item.unit_price) if item.unit_price else 0,
                'Requested Qty': item.req_qty,
                'Line Amount': float(item.line_amount) if item.line_amount else 0,
                'Publish Date': item.publish_date.strftime('%Y-%m-%d') if item.publish_date else '',
                'AC Amount': float(item.ac_amount) if item.ac_amount else 0,
                'AC Date': item.ac_date.strftime('%Y-%m-%d') if item.ac_date else '',
                'PAC Amount': float(item.pac_amount) if item.pac_amount else 0,
                'PAC Date': item.pac_date.strftime('%Y-%m-%d') if item.pac_date else '',
                'Status': item.status,
                'Remaining Amount': float(item.remaining) if item.remaining else 0  
            })
        
        df = pd.DataFrame(data_dict)
        
        # Create Excel file in memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Merged PO Data', index=False)
        
        output.seek(0)
        
        return StreamingResponse(
            BytesIO(output.getvalue()),
            media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            headers={"Content-Disposition": "attachment; filename=filtered_merged_po_data.xlsx"}
        )
        
    except Exception as e:
        logger.error(f"Error exporting data: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error exporting data: {str(e)}")