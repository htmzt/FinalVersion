# app/services/acceptance_service.py
from typing import Dict, Any, Optional
from sqlalchemy.orm import Session
from app.models import Acceptance
from app.services.base_service import BaseService

class AcceptanceService(BaseService):
    def __init__(self, db: Session):
        super().__init__(db)
    
    def get_acceptance_data(self, user_id: str, page: int, per_page: int,
                          status: Optional[str] = None,
                          project_name: Optional[str] = None,
                          search: Optional[str] = None) -> Dict[str, Any]:
        """Get paginated Acceptance data with filters"""
        
        # Base query
        query = self.db.query(Acceptance).filter(
            Acceptance.user_id == user_id
        )
        
        # Apply filters
        query = self.apply_filters(query, {
            'status': status,
            'project_name': project_name,
            'search': search
        })
        
        # Order by creation date
        query = query.order_by(Acceptance.created_at.desc())
        
        return self.get_paginated_results(query, page, per_page)
    
    def apply_filters(self, query, filters: Dict[str, Any]):
        """Apply Acceptance-specific filters"""
        if filters.get('status'):
            query = query.filter(Acceptance.status == filters['status'])
        
        if filters.get('project_name'):
            query = query.filter(
                Acceptance.project_name.ilike(f"%{filters['project_name']}%")
            )
        
        if filters.get('search'):
            search_filter = f"%{filters['search']}%"
            query = query.filter(
                (Acceptance.acceptance_no.ilike(search_filter)) |
                (Acceptance.po_number.ilike(search_filter))
            )
        
        return query
    
    def get_acceptance_count(self, user_id: str) -> int:
        """Get total acceptance count for user"""
        return self.db.query(Acceptance).filter(
            Acceptance.user_id == user_id
        ).count()