# app/services/po_service.py
from typing import Dict, Any, Optional
from sqlalchemy.orm import Session
from app.models import PurchaseOrder
from app.services.base_service import BaseService

class POService(BaseService):
    def __init__(self, db: Session):
        super().__init__(db)
    
    def get_po_data(self, user_id: str, page: int, per_page: int, 
                    project_name: Optional[str] = None, 
                    po_status: Optional[str] = None, 
                    search: Optional[str] = None) -> Dict[str, Any]:
        """Get paginated PO data with filters"""
        
        # Base query with user filter
        query = self.db.query(PurchaseOrder).filter(
            PurchaseOrder.user_id == user_id
        )
        
        # Apply filters
        query = self.apply_filters(query, {
            'project_name': project_name,
            'po_status': po_status,
            'search': search
        })
        
        # Order by creation date
        query = query.order_by(PurchaseOrder.created_at.desc())
        
        return self.get_paginated_results(query, page, per_page)
    
    def apply_filters(self, query, filters: Dict[str, Any]):
        """Apply PO-specific filters"""
        if filters.get('project_name'):
            query = query.filter(
                PurchaseOrder.project_name.ilike(f"%{filters['project_name']}%")
            )
        
        if filters.get('po_status'):
            query = query.filter(PurchaseOrder.po_status == filters['po_status'])
        
        if filters.get('search'):
            search_filter = f"%{filters['search']}%"
            query = query.filter(
                (PurchaseOrder.po_number.ilike(search_filter)) |
                (PurchaseOrder.item_description.ilike(search_filter))
            )
        
        return query
    
    def get_po_count(self, user_id: str) -> int:
        """Get total PO count for user"""
        return self.db.query(PurchaseOrder).filter(
            PurchaseOrder.user_id == user_id
        ).count()