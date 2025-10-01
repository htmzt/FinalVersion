# app/processors/acceptance_processor.py
import uuid
from datetime import datetime
from typing import Dict, List, Any
from sqlalchemy.orm import Session
from app.models import Acceptance, AcceptanceStaging
from app.processors.base_etl_processor import BaseETLProcessor
import logging

logger = logging.getLogger(__name__)

class AcceptanceProcessor(BaseETLProcessor):
    def __init__(self, db_session: Session = None):
        super().__init__(db_session)
        
        # Set models for this processor
        self.staging_model = AcceptanceStaging
        self.main_model = Acceptance
        
        # Column mapping specific to Acceptance
        self.column_mapping = {
            'id': 'id',
            'acceptanceno.': 'acceptance_no',
            'status': 'status',
            'rejected_reason': 'rejected_reason',
            'pono.': 'po_number',
            'polineno.': 'po_line_no',
            'shipmentno.': 'shipment_no',
            'item_description': 'item_description',
            'item_description(local)': 'item_description_local',
            'projectcode': 'project_code',
            'projectname': 'project_name',
            'sitecode': 'site_code',
            'sitename': 'site_name',
            'siteid': 'site_id',
            'engineeringcode': 'engineering_code',
            'businesstype': 'business_type',
            'productcategory': 'product_category',
            'requestedqty': 'requested_qty',
            'acceptanceqty': 'acceptance_qty',
            'unitprice': 'unit_price',
            'milestonetype': 'milestone_type',
            'acceptancemilestone': 'acceptance_milestone',
            'cancelremainingqty': 'cancel_remaining_qty',
            'biddingarea': 'bidding_area',
            'customer': 'customer',
            'repoffice': 'rep_office',
            'unit': 'unit',
            'subprojectcode': 'subproject_code',
            'engineeringcategory': 'engineering_category',
            'centerarea': 'center_area',
            'plannedcompletiondate': 'planned_completion_date',
            'actualcompletiondate': 'actual_completion_date',
            'approver': 'approver',
            'currenthandler': 'current_handler',
            'approvalprogress': 'approval_progress',
            'isdpproject': 'isdp_project',
            'applicationsubmitted': 'application_submitted',
            'applicationprocessed': 'application_processed',
            'headerremarks': 'header_remarks',
            'remarks': 'remarks',
            'servicecode': 'service_code',
            'payment_percentage': 'payment_percentage'
        }
    
    def validate_record(self, record: Dict[str, Any], row_num: int) -> List[Dict[str, str]]:
        """Validate Acceptance record"""
        errors = []
        
        if not record.get('acceptance_no'):
            errors.append({
                'row': row_num,
                'field': 'acceptance_no',
                'value': record.get('acceptance_no'),
                'error': 'Acceptance Number is required'
            })
        
        if not record.get('po_number'):
            errors.append({
                'row': row_num,
                'field': 'po_number',
                'value': record.get('po_number'),
                'error': 'PO Number is required'
            })
        
        if not record.get('po_line_no'):
            errors.append({
                'row': row_num,
                'field': 'po_line_no',
                'value': record.get('po_line_no'),
                'error': 'PO Line Number is required'
            })
            
        if not record.get('shipment_no'):
            errors.append({
                'row': row_num,
                'field': 'shipment_no',
                'value': record.get('shipment_no'),
                'error': 'Shipment Number is required'
            })
        
        return errors
    
    def transform_and_load(self, user_id: str) -> bool:
        """Transform staging data and load into main Acceptance table"""
        try:
            user_uuid = uuid.UUID(user_id) if isinstance(user_id, str) else user_id
            
            logger.info(f"Transforming and loading Acceptance data for user {user_id}...")
            
            # Delete existing acceptances for this user (full replace strategy)
            logger.info(f"🧹 Deleting all existing acceptances for user {user_id}")
            deleted_count = self.db.query(Acceptance).filter(
                Acceptance.user_id == user_uuid
            ).delete(synchronize_session=False)
            self.db.commit()
            logger.info(f"🗑️ Deleted {deleted_count} existing acceptances")
            
            # Get valid staging records
            valid_records = self.db.query(AcceptanceStaging).filter(
                AcceptanceStaging.user_id == user_uuid,
                AcceptanceStaging.batch_id == self.batch_id,
                AcceptanceStaging.is_valid == True,
                AcceptanceStaging.is_processed == False
            ).all()
            
            logger.info(f"Processing {len(valid_records)} valid Acceptance records for user {user_id}")
            
            for staging_record in valid_records:
                try:
                    # Create new acceptance
                    new_acceptance = self._create_acceptance_from_staging(user_uuid, staging_record)
                    self.db.add(new_acceptance)
                    logger.info(f"  ➕ Created new Acceptance: {staging_record.acceptance_no}-{staging_record.po_number}-{staging_record.po_line_no}-{staging_record.shipment_no}")
                    self.stats['new_records'] += 1
                    
                except Exception as e:
                    logger.error(f"❌ Error processing acceptance record: {e}")
                    continue
            
            self.db.commit()
            logger.info(f"✅ Successfully processed {self.stats['new_records']} new Acceptance records for user {user_id}")
            return True
            
        except Exception as e:
            logger.error(f"❌ Error in Acceptance transformation: {e}")
            self.db.rollback()
            return False
    
    def _create_acceptance_from_staging(self, user_id: uuid.UUID, staging_record: AcceptanceStaging) -> Acceptance:
        """Create Acceptance from staging record"""
        return Acceptance(
            user_id=user_id,
            acceptance_no=self.safe_string_truncate(staging_record.acceptance_no, 100),
            status=self.safe_string_truncate(staging_record.status, 50),
            rejected_reason=staging_record.rejected_reason,
            po_number=self.safe_string_truncate(staging_record.po_number, 100),
            po_line_no=self.safe_string_truncate(staging_record.po_line_no, 50),
            shipment_no=self.parse_integer(staging_record.shipment_no),
            item_description=staging_record.item_description,
            item_description_local=staging_record.item_description_local,
            project_code=self.safe_string_truncate(staging_record.project_code, 100),
            project_name=self.safe_string_truncate(staging_record.project_name, 255),
            site_code=self.safe_string_truncate(staging_record.site_code, 100),
            site_name=self.safe_string_truncate(staging_record.site_name, 255),
            site_id=self.safe_string_truncate(staging_record.site_id, 255),
            engineering_code=self.safe_string_truncate(staging_record.engineering_code, 100),
            business_type=staging_record.business_type,
            product_category=staging_record.product_category,
            requested_qty=self.parse_integer(staging_record.requested_qty),
            acceptance_qty=self.parse_integer(staging_record.acceptance_qty),
            unit_price=self.parse_decimal(staging_record.unit_price),
            milestone_type=self.safe_string_truncate(staging_record.milestone_type, 100),
            acceptance_milestone=self.safe_string_truncate(staging_record.acceptance_milestone, 100),
            cancel_remaining_qty=staging_record.cancel_remaining_qty,
            bidding_area=self.safe_string_truncate(staging_record.bidding_area, 255),
            customer=staging_record.customer,
            rep_office=self.safe_string_truncate(staging_record.rep_office, 255),
            unit=self.safe_string_truncate(staging_record.unit, 50),
            subproject_code=self.safe_string_truncate(staging_record.subproject_code, 100),
            engineering_category=self.safe_string_truncate(staging_record.engineering_category, 255),
            center_area=self.safe_string_truncate(staging_record.center_area, 255),
            planned_completion_date=self.parse_date(staging_record.planned_completion_date),
            actual_completion_date=self.parse_date(staging_record.actual_completion_date),
            approver=self.safe_string_truncate(staging_record.approver, 255),
            current_handler=staging_record.current_handler,
            approval_progress=self.safe_string_truncate(staging_record.approval_progress, 100),
            isdp_project=self.safe_string_truncate(staging_record.isdp_project, 100),
            application_submitted=self.parse_date(staging_record.application_submitted),
            application_processed=self.parse_date(staging_record.application_processed),
            header_remarks=staging_record.header_remarks,
            remarks=staging_record.remarks,
            service_code=self.parse_decimal(staging_record.service_code),
            payment_percentage=self.safe_string_truncate(staging_record.payment_percentage, 50),
            record_status='active'
        )


def process_user_acceptance_csv(file_path: str, user_id: str) -> Dict:
    """Process user Acceptance CSV file - main function called by FileService"""
    from app.database import SessionLocal
    
    db = SessionLocal()
    try:
        processor = AcceptanceProcessor(db)
        
        logger.info(f"📄 Starting Acceptance CSV processing for user: {user_id}")
        
        # Load CSV into staging
        if not processor.load_csv(file_path, user_id):
            return {
                'success': False,
                'error': 'Failed to load CSV into staging',
                'stats': processor.get_stats()
            }
        
        # Transform and load into main table
        if not processor.transform_and_load(user_id):
            return {
                'success': False,
                'error': 'Failed to transform and load data',
                'stats': processor.get_stats()
            }
        
        processor.print_summary()
        return {
            'success': True,
            'stats': processor.get_stats(),
            'batch_id': str(processor.batch_id)
        }
        
    except Exception as e:
        logger.error(f"Critical error in process_user_acceptance_csv: {e}")
        return {
            'success': False,
            'error': str(e),
            'stats': {}
        }
    finally:
        db.close()