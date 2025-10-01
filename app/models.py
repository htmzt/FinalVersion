from sqlalchemy import (
    Column, String, Integer, Boolean, DateTime, Date, Text, 
    DECIMAL, BigInteger, ForeignKey, UniqueConstraint, Index,Numeric
)
from sqlalchemy.dialects.postgresql import UUID
from sqlalchemy.orm import relationship
from sqlalchemy.dialects.postgresql import UUID, JSONB
from sqlalchemy.orm import relationship
from sqlalchemy.sql import func
import uuid
from app.database import Base


class User(Base):
    __tablename__ = "users"
    
    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    email = Column(String(255), unique=True, nullable=False)
    password_hash = Column(String(255), nullable=False)
    prenom = Column(String(100), nullable=False)
    nom = Column(String(100), nullable=False)
    company_name = Column(String(255), nullable=False)
    company_logo = Column(String(500))
    is_active = Column(Boolean, default=True)
    email_verified = Column(Boolean, default=False)
    created_at = Column(DateTime(timezone=True), server_default=func.now())
    updated_at = Column(DateTime(timezone=True), server_default=func.now(), onupdate=func.now())
    last_login = Column(DateTime(timezone=True))
    
    # Relationships
    purchase_orders = relationship("PurchaseOrder", back_populates="user", cascade="all, delete-orphan")
    staging_records = relationship("POStaging", back_populates="user", cascade="all, delete-orphan")
    audit_logs = relationship("POAuditLog", back_populates="user", cascade="all, delete-orphan")

#-----------------------------------------------------------------------------------------------------------------------------
#----------------------------------------PurchaseOrder---------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------


class PurchaseOrder(Base):
    __tablename__ = "purchase_orders"
    
    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    user_id = Column(UUID(as_uuid=True), ForeignKey("users.id", ondelete="CASCADE"), nullable=False)
    
    po_number = Column(String(100), nullable=False)
    po_line_no = Column(String(50), nullable=False)  
    
    project_name = Column(String(255))
    project_code = Column(String(100))
    site_name = Column(String(255))
    site_code = Column(String(100))
    item_code = Column(String(100))
    item_description = Column(Text)
    item_description_local = Column(Text)
    unit_price = Column(DECIMAL(12, 4))
    requested_qty = Column(Integer)
    due_qty = Column(Integer)
    billed_qty = Column(Integer)
    quantity_cancel = Column(Integer)
    line_amount = Column(DECIMAL(15, 2))
    unit = Column(String(50))
    currency = Column(String(10))
    tax_rate = Column(DECIMAL(5, 2))
    po_status = Column(String(50))
    payment_terms = Column(String(255))
    payment_method = Column(String(100))
    customer = Column(String(255))
    rep_office = Column(String(255))
    subcontract_no = Column(String(100))
    pr_no = Column(String(100))
    sales_contract_no = Column(String(100))
    version_no = Column(String(50))
    shipment_no = Column(String(100))
    engineering_code = Column(String(100))
    engineering_name = Column(String(255))
    subproject_code = Column(String(100))
    category = Column(String(255))
    center_area = Column(String(255))
    product_category = Column(String(255))
    bidding_area = Column(String(255))
    bill_to = Column(Text)
    ship_to = Column(Text)
    note_to_receiver = Column(Text)
    ff_buyer = Column(String(255))
    fob_lookup_code = Column(String(100))
    publish_date = Column(Date)
    start_date = Column(Date)
    end_date = Column(Date)
    expire_date = Column(Date)
    acceptance_date = Column(Date)
    acceptance_date_1 = Column(Date)
    change_history = Column(Text)
    pr_po_automation = Column(Text)
    status = Column(String(50), default='active')
    created_at = Column(DateTime(timezone=True), server_default=func.now())
    updated_at = Column(DateTime(timezone=True), server_default=func.now(), onupdate=func.now())
    
    # Relationships
    user = relationship("User", back_populates="purchase_orders")

    __table_args__ = (
        UniqueConstraint('user_id', 'po_number', 'po_line_no', name='uq_user_po_line'),
        Index('idx_user_po_lookup', 'user_id', 'po_number', 'po_line_no'),
        Index('idx_user_po_status', 'user_id', 'po_status'),
        Index('idx_user_project', 'user_id', 'project_code'),
    )

class POStaging(Base):
    __tablename__ = "po_staging"
    
    staging_id = Column(BigInteger, primary_key=True)
    user_id = Column(UUID(as_uuid=True), ForeignKey("users.id", ondelete="CASCADE"), nullable=False)
    batch_id = Column(UUID(as_uuid=True), nullable=False)
    
    row_number = Column(Integer)
    is_processed = Column(Boolean, default=False)
    is_valid = Column(Boolean, default=True)
    validation_errors = Column(JSONB, default=list)
    processed_at = Column(DateTime(timezone=True))
    created_at = Column(DateTime(timezone=True), server_default=func.now())
    # Raw CSV data 
    id = Column(String(100))  
    po_number = Column(String(100), nullable=False)  
    po_line_no = Column(String(50), nullable=False)  
    project_name = Column(String(255))
    project_code = Column(String(100))
    site_name = Column(String(255))
    site_code = Column(String(100))
    item_code = Column(String(100))
    item_description = Column(Text)
    item_description_local = Column(Text)
    unit_price = Column(String(50))
    requested_qty = Column(String(50))
    due_qty = Column(String(50))
    billed_qty = Column(String(50))
    quantity_cancel = Column(String(50))
    line_amount = Column(String(50))
    unit = Column(String(50))
    currency = Column(String(10))
    tax_rate = Column(String(50))
    po_status = Column(String(50))
    payment_terms = Column(String(255))
    payment_method = Column(String(100))
    customer = Column(String(255))
    rep_office = Column(String(255))
    subcontract_no = Column(String(100))
    pr_no = Column(String(100))
    sales_contract_no = Column(String(100))
    version_no = Column(String(50))
    shipment_no = Column(String(100))
    engineering_code = Column(String(100))
    engineering_name = Column(String(255))
    subproject_code = Column(String(100))
    category = Column(String(255))
    center_area = Column(String(255))
    product_category = Column(String(255))
    bidding_area = Column(String(255))
    bill_to = Column(Text)
    ship_to = Column(Text)
    note_to_receiver = Column(Text)
    ff_buyer = Column(String(255))
    fob_lookup_code = Column(String(100))
    publish_date = Column(String(50))
    start_date = Column(String(50))
    end_date = Column(String(50))
    expire_date = Column(String(50))
    acceptance_date = Column(String(50))
    acceptance_date_1 = Column(String(50))
    change_history = Column(Text)
    pr_po_automation = Column(Text)
    
# Relationships
    user = relationship("User", back_populates="staging_records")
    
    __table_args__ = (
        Index('idx_staging_user_batch', 'user_id', 'batch_id'),
        Index('idx_staging_processing', 'user_id', 'is_processed', 'is_valid'),
        Index('idx_staging_po_lookup', 'user_id', 'po_number', 'po_line_no'),
    )


class POAuditLog(Base):
    __tablename__ = "po_audit_log"
    
    id = Column(BigInteger, primary_key=True)
    user_id = Column(UUID(as_uuid=True), ForeignKey("users.id", ondelete="CASCADE"), nullable=False)
    batch_id = Column(UUID(as_uuid=True))
    
    po_number = Column(String(100), nullable=False)
    po_line_no = Column(String(50), nullable=False)
    
    action = Column(String(20), nullable=False)  # INSERT, UPDATE, DELETE
    change_source = Column(String(50), default='csv_upload')
    
    old_values = Column(JSONB)
    new_values = Column(JSONB)
    changed_fields = Column(JSONB)
    
    created_at = Column(DateTime(timezone=True), server_default=func.now())
    # Relationships
    user = relationship("User", back_populates="audit_logs")
    
    __table_args__ = (
        Index('idx_audit_user_po', 'user_id', 'po_number', 'po_line_no'),
        Index('idx_audit_user_batch', 'user_id', 'batch_id'),
        Index('idx_audit_user_action', 'user_id', 'action'),
        Index('idx_audit_created_at', 'created_at'),
    )

#-----------------------------------------------------------------------------------------------------------------------------
#----------------------------------------Subs---------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------

class Account(Base):
    __tablename__ = "accounts"
    
    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    user_id = Column(UUID(as_uuid=True), ForeignKey("users.id", ondelete="CASCADE"), nullable=False)
    account_name = Column(String(255), nullable=False)
    project_name = Column(String(255))
    needs_review = Column(Boolean, default=False)
    created_at = Column(DateTime(timezone=True), server_default=func.now()) 
    updated_at = Column(DateTime(timezone=True), server_default=func.now(), onupdate=func.now()) 
    
    user = relationship("User", backref="accounts")
    
    __table_args__ = (
        UniqueConstraint('user_id', 'project_name', name='uq_user_project'), 
        Index('idx_user_account', 'user_id', 'account_name'),
    )
class Category(Base):
    __tablename__ = "categories"
    
    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    user_id = Column(UUID(as_uuid=True), ForeignKey("users.id", ondelete="CASCADE"), nullable=False)
    category_name = Column(String(255), nullable=False)
    item_description = Column(Text)
    needs_review = Column(Boolean, default=False)
    
    user = relationship("User", backref="categories")
    
    __table_args__ = (
        UniqueConstraint('user_id', 'category_name', name='uq_user_category'),
        Index('idx_user_category', 'user_id', 'category_name'),
    )

class PaymentTerm(Base):
    __tablename__ = "payment_terms"
    
    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    user_id = Column(UUID(as_uuid=True), ForeignKey("users.id", ondelete="CASCADE"), nullable=False)
    payment_term = Column(String(255), nullable=False)
    abbreviation = Column(String(50))
    
    user = relationship("User", backref="payment_terms")
    
    __table_args__ = (
        UniqueConstraint('user_id', 'payment_term', name='uq_user_payment_term'),
        Index('idx_user_payment_term', 'user_id', 'payment_term'),
    )

#-----------------------------------------------------------------------------------------------------------------------------
#----------------------------------------Acceptance---------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------

class Acceptance(Base):
    __tablename__ = "acceptances"

    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    user_id = Column(UUID(as_uuid=True), ForeignKey("users.id", ondelete="CASCADE"), nullable=False)

    acceptance_no = Column(String(100), nullable=False)
    po_number = Column(String(100), nullable=False)
    po_line_no = Column(String(50), nullable=False)
    shipment_no = Column(Integer, nullable=False)

    
    status = Column(String(50)) 
    rejected_reason = Column(Text)
    item_description = Column(Text)
    item_description_local = Column(Text)
    project_code = Column(String(100))
    project_name = Column(String(255))
    site_code = Column(String(100))
    site_name = Column(String(255))
    site_id = Column(String(255))
    engineering_code = Column(String(100))
    business_type = Column(Text)
    product_category = Column(Text)
    requested_qty = Column(Integer)
    acceptance_qty = Column(Integer)
    unit_price = Column(DECIMAL(15, 4))
    milestone_type = Column(String(100))
    acceptance_milestone = Column(String(100))
    cancel_remaining_qty = Column(Text)
    bidding_area = Column(String(255))
    customer = Column(Text)
    rep_office = Column(String(255))
    unit = Column(String(50))
    subproject_code = Column(String(100))
    engineering_category = Column(String(255))
    center_area = Column(String(255))
    planned_completion_date = Column(Date)
    actual_completion_date = Column(Date)
    approver = Column(String(255))
    current_handler = Column(Text)
    approval_progress = Column(String(100))
    isdp_project = Column(String(100))
    application_submitted = Column(Date)
    application_processed = Column(Date)
    header_remarks = Column(Text)
    remarks = Column(Text)
    service_code = Column(DECIMAL(15, 4))
    payment_percentage = Column(String(50))
    record_status = Column(String(50), default='active') 
    created_at = Column(DateTime(timezone=True), server_default=func.now())
    updated_at = Column(DateTime(timezone=True), server_default=func.now(), onupdate=func.now())

    user = relationship("User", backref="acceptances")

    __table_args__ = (
        Index('idx_user_acceptance_lookup', 'user_id', 'acceptance_no', 'po_number', 'po_line_no', 'shipment_no'),
        Index('idx_user_acceptance_status', 'user_id', 'record_status'),
        Index('idx_user_acceptance_project', 'user_id', 'project_code'),
    )



class AcceptanceStaging(Base):
    __tablename__ = "acceptance_staging"

    staging_id = Column(BigInteger, primary_key=True)
    user_id = Column(UUID(as_uuid=True), ForeignKey("users.id", ondelete="CASCADE"), nullable=False)
    batch_id = Column(UUID(as_uuid=True), nullable=False)

    row_number = Column(Integer)
    is_processed = Column(Boolean, default=False)
    is_valid = Column(Boolean, default=True)
    validation_errors = Column(JSONB, default=list)
    processed_at = Column(DateTime(timezone=True))
    created_at = Column(DateTime(timezone=True), server_default=func.now())

    id = Column(String(100))
    acceptance_no = Column(String(100))
    status = Column(String(50))
    rejected_reason = Column(Text)
    po_number = Column(String(100))
    po_line_no = Column(String(50))
    shipment_no = Column(String(50))
    item_description = Column(Text)
    item_description_local = Column(Text)
    project_code = Column(String(100))
    project_name = Column(String(255))
    site_code = Column(String(100))
    site_name = Column(String(255))
    site_id = Column(String(50))
    engineering_code = Column(String(100))
    business_type = Column(Text)
    product_category = Column(Text)
    requested_qty = Column(String(50))
    acceptance_qty = Column(String(50))
    unit_price = Column(String(50))
    milestone_type = Column(String(100))
    acceptance_milestone = Column(String(100))
    cancel_remaining_qty = Column(Text)
    bidding_area = Column(String(255))
    customer = Column(String(50))
    rep_office = Column(String(255))
    unit = Column(String(50))
    subproject_code = Column(String(100))
    engineering_category = Column(String(255))
    center_area = Column(String(255))
    planned_completion_date = Column(String(50))
    actual_completion_date = Column(String(50))
    approver = Column(String(255))
    current_handler = Column(Text)
    approval_progress = Column(String(100))
    isdp_project = Column(String(100))
    application_submitted = Column(String(50))
    application_processed = Column(String(50))
    header_remarks = Column(Text)
    remarks = Column(Text)
    service_code = Column(String(50))
    payment_percentage = Column(String(50))

    user = relationship("User", backref="acceptance_staging_records")

    __table_args__ = (
        Index('idx_acceptance_staging_user_batch', 'user_id', 'batch_id'),
        Index('idx_acceptance_staging_processing', 'user_id', 'is_processed', 'is_valid'),
        Index('idx_acceptance_staging_lookup', 'user_id', 'acceptance_no', 'po_number', 'po_line_no'),
    )
#-----------------------------------------------------------------------------------------------------------------------------
#----------------------------------------MERGED TABLES--------------------------------------------------------------
#----------------------------------------------------------------------------------------------------------------------------- 


