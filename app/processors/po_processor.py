# app/processors/po_processor.py
import uuid
from datetime import datetime
from typing import Dict, List, Any
from sqlalchemy.orm import Session
from app.models import PurchaseOrder, POStaging, POAuditLog, Account
from app.processors.base_etl_processor import BaseETLProcessor
import logging

logger = logging.getLogger(__name__)

class POProcessor(BaseETLProcessor):
    def __init__(self, db_session: Session = None):
        super().__init__(db_session)
        
        # Set models for this processor
        self.staging_model = POStaging
        self.main_model = PurchaseOrder
        
        # Column mapping specific to PO
        self.column_mapping = {
            'id': 'id', 'change_history': 'change_history', 'rep_office': 'rep_office',
            'project_name': 'project_name', 'tax_rate': 'tax_rate', 'site_name': 'site_name',
            'item_description': 'item_description', 'note_to_receiver': 'note_to_receiver',
            'unit_price': 'unit_price', 'due_qty': 'due_qty', 'po_status': 'po_status',
            'po_no.': 'po_number', 'po_line_no.': 'po_line_no', 'item_code': 'item_code',
            'billed_quantity': 'billed_qty', 'requested_qty': 'requested_qty',
            'publish_date': 'publish_date', 'project_code': 'project_code',
            'payment_terms': 'payment_terms', 'customer': 'customer', 'site_code': 'site_code',
            'sub_contract_no.': 'subcontract_no', 'pr_no.': 'pr_no',
            'sales_contract_no.': 'sales_contract_no', 'version_no.': 'version_no',
            'shipment_no.': 'shipment_no', 'item_description(local)': 'item_description_local',
            'quantity_cancel': 'quantity_cancel', 'line_amount': 'line_amount', 'unit': 'unit',
            'currency': 'currency', 'payment_method': 'payment_method', 'bill_to': 'bill_to',
            'ship_to': 'ship_to', 'engineering_code': 'engineering_code',
            'engineering_name': 'engineering_name', 'subproject_code': 'subproject_code',
            'category': 'category', 'center_area': 'center_area',
            'product_category': 'product_category', 'bidding_area': 'bidding_area',
            'start_date': 'start_date', 'end_date': 'end_date', 'expire_date': 'expire_date',
            'acceptance_date': 'acceptance_date', 'ff_buyer': 'ff_buyer',
            'fob_lookup_code': 'fob_lookup_code',
            'pr/po_automation_solution_(only_china)': 'pr_po_automation'
        }
    
    def validate_record(self, record: Dict[str, Any], row_num: int) -> List[Dict[str, str]]:
        """Validate PO record"""
        errors = []
        if not record.get('po_number'):
            errors.append({'row': row_num, 'field': 'po_number', 'error': 'PO Number is required'})
        if not record.get('po_line_no'):
            errors.append({'row': row_num, 'field': 'po_line_no', 'error': 'PO Line Number is required'})
        return errors
    
    def transform_and_load(self, user_id: str) -> bool:
        """Transform staging data and load into main PO table"""
        try:
            user_uuid = uuid.UUID(user_id)
            logger.info(f"Transforming and loading PO data for user {user_id}...")
            
            # Get valid staging records
            valid_records = self.db.query(POStaging).filter(
                POStaging.user_id == user_uuid,
                POStaging.batch_id == self.batch_id,
                POStaging.is_valid == True,
                POStaging.is_processed == False
            ).all()
            
            logger.info(f"Processing {len(valid_records)} valid PO records for user {user_id}")
            
            for staging_record in valid_records:
                try:
                    # Create or get account
                    self._get_or_create_account(user_uuid, staging_record.project_name)
                    
                    po_number = self.safe_string_truncate(staging_record.po_number, 100)
                    po_line_no = self.safe_string_truncate(staging_record.po_line_no, 50)
                    
                    # Check if PO already exists
                    existing_po = self.db.query(PurchaseOrder).filter(
                        PurchaseOrder.user_id == user_uuid,
                        PurchaseOrder.po_number == po_number,
                        PurchaseOrder.po_line_no == po_line_no
                    ).first()
                    
                    if existing_po:
                        # Update existing PO
                        old_values = self._get_record_dict(existing_po)
                        self._update_po_from_staging(existing_po, staging_record)
                        new_values = self._get_record_dict(existing_po)
                        changed_fields = [k for k, v in old_values.items() if v != new_values.get(k)]
                        
                        if changed_fields:
                            # Log audit trail
                            self.db.add(POAuditLog(
                                user_id=user_uuid,
                                batch_id=self.batch_id,
                                po_number=po_number,
                                po_line_no=po_line_no,
                                action='UPDATE',
                                old_values=self.serialize_for_json(old_values),
                                new_values=self.serialize_for_json(new_values),
                                changed_fields=changed_fields
                            ))
                            self.stats['updated_records'] += 1
                    else:
                        # Create new PO
                        new_po = self._create_po_from_staging(user_uuid, staging_record)
                        self.db.add(new_po)
                        self.db.flush()
                        
                        # Log audit trail
                        self.db.add(POAuditLog(
                            user_id=user_uuid,
                            batch_id=self.batch_id,
                            po_number=po_number,
                            po_line_no=po_line_no,
                            action='INSERT',
                            new_values=self.serialize_for_json(self._get_record_dict(new_po))
                        ))
                        self.stats['new_records'] += 1
                    
                    # Mark as processed
                    staging_record.is_processed = True
                    staging_record.processed_at = datetime.utcnow()
                    
                except Exception as e:
                    logger.error(f"Error processing staging record {staging_record.staging_id}: {e}")
                    self.db.rollback()
                    staging_record.is_processed = True
                    staging_record.is_valid = False
                    staging_record.validation_errors = [{'error': str(e)}]
                    self.db.add(staging_record)
                    self.db.commit()
            
            self.db.commit()
            logger.info(f"✅ Processed {self.stats['new_records']} new and {self.stats['updated_records']} updated PO records")
            return True
            
        except Exception as e:
            logger.error(f"❌ Error in PO transformation: {e}")
            self.db.rollback()
            return False
    
    def _map_project_to_account_name(self, project_name: str) -> str:
        """Map project name to account name"""
        if not project_name:
            return 'Other'
        
        project_name_lower = project_name.lower()
        if "iam" in project_name_lower:
            return "IAM Account"
        elif "orange" in project_name_lower:
            return "Orange Account"
        elif "inwi" in project_name_lower:
            return "INWI Account"
        else:
            return "Other"
    
    def _get_or_create_account(self, user_id: uuid.UUID, project_name: str) -> None:
        """Get or create account for project"""
        clean_project_name = (project_name or "Unknown Project").strip()
        
        existing_account = self.db.query(Account).filter(
            Account.user_id == user_id,
            Account.project_name == clean_project_name
        ).first()
        
        if existing_account:
            return
        
        account_name = self._map_project_to_account_name(clean_project_name)
        needs_review = (account_name == "Other")
        
        new_account = Account(
            user_id=user_id,
            project_name=clean_project_name,
            account_name=account_name,
            needs_review=needs_review
        )
        self.db.add(new_account)
        self.db.flush()
        logger.info(f"✨ Created new account '{account_name}' for project '{clean_project_name}'")
    
    def _create_po_from_staging(self, user_id: uuid.UUID, staging: POStaging) -> PurchaseOrder:
        """Create PurchaseOrder from staging record"""
        return PurchaseOrder(
            user_id=user_id,
            po_number=self.safe_string_truncate(staging.po_number, 100),
            po_line_no=self.safe_string_truncate(staging.po_line_no, 50),
            project_name=self.safe_string_truncate(staging.project_name, 255),
            project_code=self.safe_string_truncate(staging.project_code, 100),
            site_name=self.safe_string_truncate(staging.site_name, 255),
            site_code=self.safe_string_truncate(staging.site_code, 100),
            item_code=self.safe_string_truncate(staging.item_code, 100),
            item_description=staging.item_description,
            item_description_local=staging.item_description_local,
            unit_price=self.parse_decimal(staging.unit_price),
            requested_qty=self.parse_integer(staging.requested_qty),
            due_qty=self.parse_integer(staging.due_qty),
            billed_qty=self.parse_integer(staging.billed_qty),
            quantity_cancel=self.parse_integer(staging.quantity_cancel),
            line_amount=self.parse_decimal(staging.line_amount),
            unit=self.safe_string_truncate(staging.unit, 50),
            currency=self.safe_string_truncate(staging.currency, 10),
            tax_rate=self.parse_decimal(staging.tax_rate),
            po_status=self.safe_string_truncate(staging.po_status, 50),
            payment_terms=self.safe_string_truncate(staging.payment_terms, 255),
            payment_method=self.safe_string_truncate(staging.payment_method, 100),
            customer=self.safe_string_truncate(staging.customer, 255),
            rep_office=self.safe_string_truncate(staging.rep_office, 255),
            subcontract_no=self.safe_string_truncate(staging.subcontract_no, 100),
            pr_no=self.safe_string_truncate(staging.pr_no, 100),
            sales_contract_no=self.safe_string_truncate(staging.sales_contract_no, 100),
            version_no=self.safe_string_truncate(staging.version_no, 50),
            shipment_no=self.safe_string_truncate(staging.shipment_no, 100),
            engineering_code=self.safe_string_truncate(staging.engineering_code, 100),
            engineering_name=self.safe_string_truncate(staging.engineering_name, 255),
            subproject_code=self.safe_string_truncate(staging.subproject_code, 100),
            category=self.safe_string_truncate(staging.category, 255),
            center_area=self.safe_string_truncate(staging.center_area, 255),
            product_category=self.safe_string_truncate(staging.product_category, 255),
            bidding_area=self.safe_string_truncate(staging.bidding_area, 255),
            bill_to=staging.bill_to,
            ship_to=staging.ship_to,
            note_to_receiver=staging.note_to_receiver,
            ff_buyer=self.safe_string_truncate(staging.ff_buyer, 255),
            fob_lookup_code=self.safe_string_truncate(staging.fob_lookup_code, 100),
            publish_date=self.parse_date(staging.publish_date),
            start_date=self.parse_date(staging.start_date),
            end_date=self.parse_date(staging.end_date),
            expire_date=self.parse_date(staging.expire_date),
            acceptance_date=self.parse_date(staging.acceptance_date),
            acceptance_date_1=self.parse_date(getattr(staging, 'acceptance_date_1', None)),
            change_history=staging.change_history,
            pr_po_automation=staging.pr_po_automation
        )
    
    def _update_po_from_staging(self, po: PurchaseOrder, staging: POStaging):
        """Update existing PO with staging data"""
        for key, value in self._create_po_from_staging(po.user_id, staging).__dict__.items():
            if not key.startswith('_'):
                setattr(po, key, value)
        po.updated_at = datetime.utcnow()
    
    def _get_record_dict(self, po: PurchaseOrder) -> Dict:
        """Get PO record as dictionary"""
        return {c.name: getattr(po, c.name) for c in po.__table__.columns}


def process_user_csv(file_path: str, user_id: str) -> Dict:
    """Process user CSV file - main function called by FileService"""
    from app.database import SessionLocal
    
    db = SessionLocal()
    try:
        processor = POProcessor(db)
        
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
        logger.error(f"Critical error in process_user_csv: {e}")
        return {
            'success': False,
            'error': str(e),
            'stats': {}
        }
    finally:
        db.close()