# app/services/summary_service.py
from typing import Dict, Any, List, Optional
from sqlalchemy.orm import Session
from sqlalchemy import text
from app.services.base_service import BaseService
from app.query import MERGED_DATA_QUERY
import logging

logger = logging.getLogger(__name__)

class SummaryService(BaseService):
    def __init__(self, db: Session):
        super().__init__(db)
    
    def get_monthly_summary(self, 
                           user_id: str,
                           year: Optional[int] = None,
                           month: Optional[int] = None,
                           project_name: Optional[str] = None) -> Dict[str, Any]:
        """Get monthly summary by project name with optional filters"""
        try:
            # Build base filter
            base_filter = "po.user_id = :user_id"
            params = {"user_id": user_id}
            
            # Add project filter if provided
            if project_name:
                base_filter += " AND po.project_name ILIKE :project_name"
                params["project_name"] = f"%{project_name}%"
            
            # Build the summary query
            summary_query = f"""
            SELECT 
                subquery.project_name,
                EXTRACT(YEAR FROM subquery.publish_date) as year,
                EXTRACT(MONTH FROM subquery.publish_date) as month,
                TO_CHAR(subquery.publish_date, 'Month YYYY') as month_year,
                COUNT(*) as total_records,
                COUNT(DISTINCT subquery.po_id) as unique_pos,
                
                -- Financial totals
                COALESCE(SUM(subquery.line_amount), 0) as total_line_amount,
                COALESCE(SUM(subquery.ac_amount), 0) as total_ac_amount,
                COALESCE(SUM(subquery.pac_amount), 0) as total_pac_amount,
                COALESCE(SUM(subquery.remaining), 0) as total_remaining_amount,
                
                -- Status counts
                COUNT(CASE WHEN subquery.status = 'CLOSED' THEN 1 END) as closed_count,
                COUNT(CASE WHEN subquery.status = 'CANCELLED' THEN 1 END) as cancelled_count,
                COUNT(CASE WHEN subquery.status LIKE '%Pending%' THEN 1 END) as pending_count,
                
                -- Payment terms breakdown
                COUNT(CASE WHEN subquery.payment_terms = 'ACPAC 100%' THEN 1 END) as acpac_100_count,
                COUNT(CASE WHEN subquery.payment_terms = 'AC1 80 | PAC 20' THEN 1 END) as ac_pac_split_count,
                
                -- Category breakdown
                COUNT(CASE WHEN subquery.category = 'Survey' THEN 1 END) as survey_count,
                COUNT(CASE WHEN subquery.category = 'Transportation' THEN 1 END) as transportation_count,
                COUNT(CASE WHEN subquery.category = 'Site Engineer' THEN 1 END) as site_engineer_count,
                COUNT(CASE WHEN subquery.category = 'Service' THEN 1 END) as service_count,
                
                -- Date range
                MIN(subquery.publish_date) as earliest_date,
                MAX(subquery.publish_date) as latest_date
                
            FROM (
                {MERGED_DATA_QUERY.format(base_filter=base_filter)}
            ) as subquery
            WHERE subquery.publish_date IS NOT NULL
            """
            
            # Add date filters
            if year:
                summary_query += " AND EXTRACT(YEAR FROM subquery.publish_date) = :year"
                params["year"] = year
            
            if month:
                summary_query += " AND EXTRACT(MONTH FROM subquery.publish_date) = :month"
                params["month"] = month
            
            # Group and order
            summary_query += """
            GROUP BY 
                subquery.project_name,
                EXTRACT(YEAR FROM subquery.publish_date),
                EXTRACT(MONTH FROM subquery.publish_date),
                TO_CHAR(subquery.publish_date, 'Month YYYY')
            ORDER BY 
                year DESC, 
                month DESC, 
                subquery.project_name ASC
            """
            
            result = self.db.execute(text(summary_query), params)
            summaries = []
            
            for row in result:
                completion_rate = 0
                if row.total_records > 0:
                    completion_rate = round((row.closed_count / row.total_records) * 100, 2)
                
                summary = {
                    "project_name": row.project_name,
                    "year": int(row.year) if row.year else None,
                    "month": int(row.month) if row.month else None,
                    "month_year": row.month_year.strip() if row.month_year else None,
                    
                    # Counts
                    "total_records": row.total_records,
                    "unique_pos": row.unique_pos,
                    
                    # Financial summary
                    "financial_summary": {
                        "total_line_amount": float(row.total_line_amount) if row.total_line_amount else 0,
                        "total_ac_amount": float(row.total_ac_amount) if row.total_ac_amount else 0,
                        "total_pac_amount": float(row.total_pac_amount) if row.total_pac_amount else 0,
                        "total_remaining_amount": float(row.total_remaining_amount) if row.total_remaining_amount else 0,
                    },
                    
                    # Status breakdown
                    "status_breakdown": {
                        "closed": row.closed_count,
                        "cancelled": row.cancelled_count,
                        "pending": row.pending_count,
                        "completion_rate": completion_rate
                    },
                    
                    # Payment terms breakdown
                    "payment_terms_breakdown": {
                        "acpac_100_percent": row.acpac_100_count,
                        "ac_pac_split": row.ac_pac_split_count
                    },
                    
                    # Category breakdown
                    "category_breakdown": {
                        "survey": row.survey_count,
                        "transportation": row.transportation_count,
                        "site_engineer": row.site_engineer_count,
                        "service": row.service_count
                    },
                    
                    # Date range
                    "date_range": {
                        "earliest_date": row.earliest_date.isoformat() if row.earliest_date else None,
                        "latest_date": row.latest_date.isoformat() if row.latest_date else None
                    }
                }
                summaries.append(summary)
            
            # Get overall totals
            overall_totals = self._get_overall_totals(user_id, year, month, project_name)
            
            return {
                "summaries": summaries,
                "overall_totals": overall_totals,
                "filters_applied": {
                    "year": year,
                    "month": month,
                    "project_name": project_name
                },
                "total_summary_records": len(summaries)
            }
            
        except Exception as e:
            logger.error(f"Error getting monthly summary: {str(e)}")
            raise
    
    def get_yearly_summary(self, user_id: str, project_name: Optional[str] = None) -> Dict[str, Any]:
        """Get yearly summary aggregated by project and year"""
        try:
            base_filter = "po.user_id = :user_id"
            params = {"user_id": user_id}
            
            if project_name:
                base_filter += " AND po.project_name ILIKE :project_name"
                params["project_name"] = f"%{project_name}%"
            
            yearly_query = f"""
            SELECT 
                subquery.project_name,
                EXTRACT(YEAR FROM subquery.publish_date) as year,
                COUNT(*) as total_records,
                COUNT(DISTINCT subquery.po_id) as unique_pos,
                COALESCE(SUM(subquery.line_amount), 0) as total_line_amount,
                COALESCE(SUM(subquery.ac_amount), 0) as total_ac_amount,
                COALESCE(SUM(subquery.pac_amount), 0) as total_pac_amount,
                COALESCE(SUM(subquery.remaining), 0) as total_remaining_amount,
                COUNT(CASE WHEN subquery.status = 'CLOSED' THEN 1 END) as closed_count,
                COUNT(DISTINCT EXTRACT(MONTH FROM subquery.publish_date)) as months_with_data
            FROM (
                {MERGED_DATA_QUERY.format(base_filter=base_filter)}
            ) as subquery
            WHERE subquery.publish_date IS NOT NULL
            GROUP BY 
                subquery.project_name,
                EXTRACT(YEAR FROM subquery.publish_date)
            ORDER BY 
                year DESC, 
                subquery.project_name ASC
            """
            
            result = self.db.execute(text(yearly_query), params)
            
            return [
                {
                    "project_name": row.project_name,
                    "year": int(row.year) if row.year else None,
                    "total_records": row.total_records,
                    "unique_pos": row.unique_pos,
                    "months_with_data": row.months_with_data,
                    "financial_summary": {
                        "total_line_amount": float(row.total_line_amount) if row.total_line_amount else 0,
                        "total_ac_amount": float(row.total_ac_amount) if row.total_ac_amount else 0,
                        "total_pac_amount": float(row.total_pac_amount) if row.total_pac_amount else 0,
                        "total_remaining_amount": float(row.total_remaining_amount) if row.total_remaining_amount else 0,
                    },
                    "completion_rate": round((row.closed_count / row.total_records) * 100, 2) if row.total_records > 0 else 0
                }
                for row in result
            ]
            
        except Exception as e:
            logger.error(f"Error getting yearly summary: {str(e)}")
            raise
    
    def get_available_periods(self, user_id: str) -> Dict[str, Any]:
        """Get available years and months for filtering"""
        try:
            periods_query = f"""
            SELECT DISTINCT
                EXTRACT(YEAR FROM subquery.publish_date) as year,
                EXTRACT(MONTH FROM subquery.publish_date) as month,
                TO_CHAR(subquery.publish_date, 'FMMonth') as month_name,
                COUNT(*) as record_count
            FROM (
                {MERGED_DATA_QUERY.format(base_filter="po.user_id = :user_id")}
            ) as subquery
            WHERE subquery.publish_date IS NOT NULL
            GROUP BY 
                EXTRACT(YEAR FROM subquery.publish_date),
                EXTRACT(MONTH FROM subquery.publish_date),
                TO_CHAR(subquery.publish_date, 'FMMonth')
            ORDER BY year DESC, month DESC
            """
            
            result = self.db.execute(text(periods_query), {"user_id": user_id})
            
            years = {}
            for row in result:
                year = int(row.year) if row.year else None
                month = int(row.month) if row.month else None
                
                if year not in years:
                    years[year] = {
                        "year": year,
                        "months": [],
                        "total_records": 0
                    }
                
                years[year]["months"].append({
                    "month": month,
                    "month_name": row.month_name.strip() if row.month_name else None,
                    "record_count": row.record_count
                })
                years[year]["total_records"] += row.record_count
            
            return {
                "available_years": sorted(years.keys(), reverse=True) if years else [],
                "year_details": list(years.values()) if years else []
            }
            
        except Exception as e:
            logger.error(f"Error getting available periods: {str(e)}")
            raise
    
    def get_project_list(self, user_id: str) -> List[str]:
        """Get list of available project names for filtering"""
        try:
            projects_query = f"""
            SELECT DISTINCT subquery.project_name
            FROM (
                {MERGED_DATA_QUERY.format(base_filter="po.user_id = :user_id")}
            ) as subquery
            WHERE subquery.project_name IS NOT NULL
            ORDER BY subquery.project_name ASC
            """
            
            result = self.db.execute(text(projects_query), {"user_id": user_id})
            return [row.project_name for row in result if row.project_name]
            
        except Exception as e:
            logger.error(f"Error getting project list: {str(e)}")
            raise
    
    def _get_overall_totals(self, user_id: str, year: Optional[int], month: Optional[int], project_name: Optional[str]) -> Dict[str, Any]:
        """Get overall totals for the filtered dataset"""
        try:
            base_filter = "po.user_id = :user_id"
            params = {"user_id": user_id}
            
            if project_name:
                base_filter += " AND po.project_name ILIKE :project_name"
                params["project_name"] = f"%{project_name}%"
            
            totals_query = f"""
            SELECT 
                COUNT(*) as total_records,
                COUNT(DISTINCT subquery.po_id) as unique_pos,
                COUNT(DISTINCT subquery.project_name) as unique_projects,
                COALESCE(SUM(subquery.line_amount), 0) as total_line_amount,
                COALESCE(SUM(subquery.ac_amount), 0) as total_ac_amount,
                COALESCE(SUM(subquery.pac_amount), 0) as total_pac_amount,
                COALESCE(SUM(subquery.remaining), 0) as total_remaining_amount,
                COUNT(CASE WHEN subquery.status = 'CLOSED' THEN 1 END) as total_closed
            FROM (
                {MERGED_DATA_QUERY.format(base_filter=base_filter)}
            ) as subquery
            WHERE subquery.publish_date IS NOT NULL
            """
            
            # Add date filters
            if year:
                totals_query += " AND EXTRACT(YEAR FROM subquery.publish_date) = :year"
                params["year"] = year
            
            if month:
                totals_query += " AND EXTRACT(MONTH FROM subquery.publish_date) = :month"
                params["month"] = month
            
            result = self.db.execute(text(totals_query), params).first()
            
            return {
                "total_records": result.total_records if result else 0,
                "unique_pos": result.unique_pos if result else 0,
                "unique_projects": result.unique_projects if result else 0,
                "financial_totals": {
                    "total_line_amount": float(result.total_line_amount) if result and result.total_line_amount else 0,
                    "total_ac_amount": float(result.total_ac_amount) if result and result.total_ac_amount else 0,
                    "total_pac_amount": float(result.total_pac_amount) if result and result.total_pac_amount else 0,
                    "total_remaining_amount": float(result.total_remaining_amount) if result and result.total_remaining_amount else 0,
                },
                "overall_completion_rate": round((result.total_closed / result.total_records) * 100, 2) if result and result.total_records > 0 else 0
            }
            
        except Exception as e:
            logger.error(f"Error getting overall totals: {str(e)}")
            return {
                "total_records": 0,
                "unique_pos": 0,
                "unique_projects": 0,
                "financial_totals": {
                    "total_line_amount": 0,
                    "total_ac_amount": 0,
                    "total_pac_amount": 0,
                    "total_remaining_amount": 0,
                },
                "overall_completion_rate": 0
            }