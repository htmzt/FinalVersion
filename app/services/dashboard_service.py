from typing import Dict, Any, List
from sqlalchemy.orm import Session
from sqlalchemy import text, func
from app.models import PurchaseOrder, Acceptance, Account
from app.services.base_service import BaseService
from app.query import MERGED_DATA_QUERY
import logging

logger = logging.getLogger(__name__)

class DashboardService(BaseService):
    def __init__(self, db: Session):
        super().__init__(db)
    
    def get_data_status(self, user_id: str) -> Dict[str, Any]:
        """Check data status with raw counts from purchase_orders and acceptances"""
        # Raw counts from purchase_orders and acceptances tables
        po_count = self.db.query(PurchaseOrder).filter(
            PurchaseOrder.user_id == user_id
        ).count()
        
        acceptance_count = self.db.query(Acceptance).filter(
            Acceptance.user_id == user_id
        ).count()
        
        # Get last upload dates from merged data
        date_query = text(f"""
        SELECT 
            MAX(subquery.publish_date) as last_po_upload,
            MAX(COALESCE(subquery.ac_date, subquery.pac_date)) as last_acceptance_upload
        FROM (
            {MERGED_DATA_QUERY.format(base_filter="po.user_id = :user_id")}
        ) as subquery
        """)
        dates = self.db.execute(date_query, {"user_id": user_id}).first()
        
        return {
            "has_data": po_count > 0 and acceptance_count > 0,
            "po_count": po_count,
            "acceptance_count": acceptance_count,
            "last_po_upload": dates.last_po_upload.isoformat() if dates and dates.last_po_upload else None,
            "last_acceptance_upload": dates.last_acceptance_upload.isoformat() if dates and dates.last_acceptance_upload else None,
            "data_quality": {
                "po_with_acceptances": self._get_matching_po_count(user_id),
                "total_pos": po_count
            }
        }
    
    def get_dashboard_analytics(self, user_id: str) -> Dict[str, Any]:
        """Get comprehensive dashboard analytics with raw counts and merged data for analytics"""
        try:
            # Raw counts from purchase_orders and acceptances tables
            po_count = self.db.query(PurchaseOrder).filter(
                PurchaseOrder.user_id == user_id
            ).count()
            
            acceptance_count = self.db.query(Acceptance).filter(
                Acceptance.user_id == user_id
            ).count()
            
            # Accounts needing review using merged data
            accounts_query = text(f"""
            SELECT COUNT(DISTINCT a.id) as accounts_needing_review
            FROM accounts a
            INNER JOIN (
                {MERGED_DATA_QUERY.format(base_filter="po.user_id = :user_id")}
            ) as subquery ON a.project_name = subquery.project_name
            WHERE a.user_id = :user_id AND a.needs_review = TRUE
            """)
            accounts = self.db.execute(accounts_query, {"user_id": user_id}).scalar()
            
            # Financial totals from merged data
            financial_stats = self._get_financial_stats(user_id)
            
            # Status breakdown, account analysis, and payment terms from merged data
            status_breakdown = self._get_status_breakdown(user_id)
            account_analysis = self._get_account_analysis(user_id)
            payment_terms_dist = self._get_payment_terms_distribution(user_id)
            
            return {
                "basic_stats": {
                    "total_pos": po_count,
                    "total_acceptances": acceptance_count,
                    "accounts_needing_review": accounts or 0,
                    **financial_stats
                },
                "status_breakdown": status_breakdown,
                "project_analysis": account_analysis,  # Renamed to account_analysis internally
                "payment_terms_distribution": payment_terms_dist
            }
        except Exception as e:
            logger.error(f"Error getting dashboard analytics: {str(e)}")
            raise
    
    def get_charts_data(self, user_id: str) -> Dict[str, Any]:
        """Get structured data specifically for React charts using merged data"""
        status_breakdown = self._get_status_breakdown(user_id)
        account_analysis = self._get_account_analysis(user_id)
        
        return {
            "status_pie_chart": {
                "labels": [item["status"] for item in status_breakdown],
                "data": [item["count"] for item in status_breakdown],
                "values": [float(item["total_value"]) for item in status_breakdown]
            },
            "project_bar_chart": {  # Kept key for frontend compatibility
                "labels": [item["account_name"] for item in account_analysis[:10]],  # Top 10
                "data": [float(item["total_value"]) for item in account_analysis[:10]],
                "pending_amounts": [float(item["pending_amount"]) for item in account_analysis[:10]]
            }
        }
    
    def _get_financial_stats(self, user_id: str) -> Dict[str, Any]:
        """Get financial statistics from merged data"""
        stats_query = text(f"""
        SELECT 
            COUNT(*) as total_records,
            COALESCE(SUM(subquery.line_amount), 0) as total_value,
            COALESCE(SUM(subquery.ac_amount), 0) as total_ac_amount,
            COALESCE(SUM(subquery.pac_amount), 0) as total_pac_amount
        FROM (
            {MERGED_DATA_QUERY.format(base_filter="po.user_id = :user_id")}
        ) as subquery
        """)
        
        result = self.db.execute(stats_query, {"user_id": user_id}).first()
        
        return {
            "total_merged_records": result.total_records if result else 0,
            "total_value": float(result.total_value) if result and result.total_value else 0,
            "total_ac_amount": float(result.total_ac_amount) if result and result.total_ac_amount else 0,
            "total_pac_amount": float(result.total_pac_amount) if result and result.total_pac_amount else 0
        }
    
    def _get_status_breakdown(self, user_id: str) -> List[Dict[str, Any]]:
        """Get status breakdown for analytics using merged data"""
        status_query = text(f"""
        SELECT 
            subquery.status,
            COUNT(*) as count,
            COALESCE(SUM(subquery.line_amount), 0) as total_value,
            COALESCE(SUM(subquery.remaining), 0) as pending_amount
        FROM (
            {MERGED_DATA_QUERY.format(base_filter="po.user_id = :user_id")}
        ) as subquery
        GROUP BY subquery.status
        ORDER BY total_value DESC
        """)
        
        result = self.db.execute(status_query, {"user_id": user_id})
        rows = result.fetchall()
        
        total_count = sum([row.count for row in rows]) if rows else 0
        
        return [
            {
                "status": row.status,
                "count": row.count,
                "total_value": float(row.total_value) if row.total_value else 0,
                "pending_amount": float(row.pending_amount) if row.pending_amount else 0,
                "percentage": round((row.count / total_count) * 100, 2) if total_count > 0 else 0
            }
            for row in rows
        ]
    
    def _get_account_analysis(self, user_id: str) -> List[Dict[str, Any]]:
        """Get account-wise analysis using merged data"""
        account_query = text(f"""
        SELECT 
            COALESCE(subquery.account_name, 'Unknown') as account_name,
            COUNT(*) as total_records,
            COALESCE(SUM(subquery.line_amount), 0) as total_value,
            COALESCE(SUM(subquery.remaining), 0) as pending_amount,
            COUNT(CASE WHEN subquery.status = 'CLOSED' THEN 1 END) as closed_count
        FROM (
            {MERGED_DATA_QUERY.format(base_filter="po.user_id = :user_id")}
        ) as subquery
        WHERE subquery.account_name IS NOT NULL OR subquery.account_name IS NULL
        GROUP BY subquery.account_name
        ORDER BY total_value DESC
        LIMIT 20
        """)
        
        result = self.db.execute(account_query, {"user_id": user_id})
        
        return [
            {
                "account_name": row.account_name,
                "total_records": row.total_records,
                "total_value": float(row.total_value) if row.total_value else 0,
                "pending_amount": float(row.pending_amount) if row.pending_amount else 0,
                "completion_rate": round((row.closed_count / row.total_records) * 100, 2) if row.total_records > 0 else 0
            }
            for row in result
        ]
    
    def _get_payment_terms_distribution(self, user_id: str) -> List[Dict[str, Any]]:
        """Get payment terms distribution using merged data"""
        payment_query = text(f"""
        SELECT 
            subquery.payment_terms,
            COUNT(*) as count,
            COALESCE(SUM(subquery.line_amount), 0) as total_value
        FROM (
            {MERGED_DATA_QUERY.format(base_filter="po.user_id = :user_id")}
        ) as subquery
        GROUP BY subquery.payment_terms
        ORDER BY total_value DESC
        """)
        
        result = self.db.execute(payment_query, {"user_id": user_id})
        
        return [
            {
                "payment_terms": row.payment_terms,
                "count": row.count,
                "total_value": float(row.total_value) if row.total_value else 0
            }
            for row in result
        ]
    
    def _get_matching_po_count(self, user_id: str) -> int:
        """Get count of POs that have corresponding acceptances"""
        matching_query = text("""
        SELECT COUNT(DISTINCT CONCAT(po.po_number, '-', po.po_line_no)) as matching_count
        FROM purchase_orders po
        INNER JOIN acceptances a ON po.user_id = a.user_id 
            AND po.po_number = a.po_number 
            AND po.po_line_no = a.po_line_no
        WHERE po.user_id = :user_id
        """)
        
        result = self.db.execute(matching_query, {"user_id": user_id}).scalar()
        return result or 0