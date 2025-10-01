# app/services/base_service.py
from abc import ABC
from typing import Dict, Any, Optional, List
from sqlalchemy.orm import Session
from sqlalchemy import func
from app.database import get_db

class BaseService(ABC):
    def __init__(self, db: Session):
        self.db = db
    
    def get_paginated_results(self, query, page: int, per_page: int):
        """Generic pagination helper"""
        total_count = query.count()
        
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
        
        offset = (page - 1) * per_page
        items = query.offset(offset).limit(per_page).all()
        total_pages = (total_count + per_page - 1) // per_page
        
        return {
            "items": items,
            "total_count": total_count,
            "page": page,
            "per_page": per_page,
            "total_pages": total_pages,
            "has_next": page < total_pages,
            "has_prev": page > 1
        }
    
    def apply_filters(self, query, filters: Dict[str, Any]):
        """Generic filter application helper"""
        # This will be overridden by specific services
        return query