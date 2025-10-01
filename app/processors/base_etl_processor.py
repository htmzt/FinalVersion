# app/processors/base_etl_processor.py
import pandas as pd
import uuid
import json
import os
from datetime import datetime, date
from decimal import Decimal, InvalidOperation
from typing import Dict, List, Any, Optional, Type
import logging
from sqlalchemy.orm import Session
from app.database import SessionLocal
from app.config import CHUNK_SIZE

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class DecimalEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, Decimal):
            return float(obj)
        if isinstance(obj, (date, datetime)):
            return obj.isoformat()
        if isinstance(obj, uuid.UUID):
            return str(obj)
        return super(DecimalEncoder, self).default(obj)

class BaseETLProcessor:
    """Base class for ETL processing with common functionality"""
    
    def __init__(self, db_session: Session = None):
        self.db = db_session or SessionLocal()
        self.batch_id = uuid.uuid4()
        self.stats = {
            'total_rows': 0,
            'processed_rows': 0,
            'failed_rows': 0,
            'new_records': 0,
            'updated_records': 0,
            'validation_errors': []
        }
        
        # These should be overridden by child classes
        self.column_mapping = {}
        self.staging_model = None
        self.main_model = None
    
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.db.close()
    
    def normalize_column_name(self, col_name: str) -> str:
        """Normalize column names to standard format"""
        return col_name.strip().lower().replace(' ', '_').replace('(', '_').replace(')', '_').replace('__', '_').strip('_')
    
    def map_csv_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """Map CSV columns to database fields"""
        df.columns = [self.normalize_column_name(col) for col in df.columns]
        mapped_data = {}
        
        for csv_col, db_field in self.column_mapping.items():
            if csv_col in df.columns:
                mapped_data[db_field] = df[csv_col]
            else:
                mapped_data[db_field] = pd.Series([None] * len(df))
        
        return pd.DataFrame(mapped_data)
    
    def parse_date(self, date_str: str) -> Optional[date]:
        """Parse various date formats"""
        if not date_str or pd.isna(date_str) or str(date_str).strip() == '':
            return None
        
        date_formats = [
            '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', '%d-%m-%Y',
            '%Y/%m/%d', '%d.%m.%Y', '%Y-%m-%d %H:%M:%S', '%m/%d/%Y %H:%M'
        ]
        
        date_str = str(date_str).strip()
        for fmt in date_formats:
            try:
                return datetime.strptime(date_str, fmt).date()
            except ValueError:
                continue
        return None
    
    def parse_decimal(self, value: Any) -> Optional[Decimal]:
        """Parse decimal values with error handling"""
        if pd.isna(value) or value == '' or value is None:
            return None
        try:
            str_value = str(value).strip().replace(',', '').replace(' ', '').replace('%', '')
            return Decimal(str_value)
        except (InvalidOperation, ValueError):
            return None
    
    def parse_integer(self, value: Any) -> Optional[int]:
        """Parse integer values with error handling"""
        if pd.isna(value) or value == '' or value is None:
            return None
        try:
            return int(float(str(value)))
        except (ValueError, OverflowError):
            return None
    
    def safe_string_truncate(self, value: Any, max_length: int = None) -> Optional[str]:
        """Safely truncate string values"""
        if pd.isna(value) or value == '' or value is None:
            return None
        str_value = str(value).strip()
        if max_length and len(str_value) > max_length:
            return str_value[:max_length]
        return str_value if str_value else None
    
    def validate_record(self, record: Dict[str, Any], row_num: int) -> List[Dict[str, str]]:
        """Validate a single record - should be overridden by child classes"""
        return []
    
    def serialize_for_json(self, obj):
        """Serialize objects for JSON storage"""
        if isinstance(obj, dict):
            return {k: self.serialize_for_json(v) for k, v in obj.items()}
        elif isinstance(obj, (list, tuple)):
            return [self.serialize_for_json(item) for item in obj]
        elif isinstance(obj, Decimal):
            return float(obj)
        elif isinstance(obj, (date, datetime)):
            return obj.isoformat()
        elif isinstance(obj, uuid.UUID):
            return str(obj)
        else:
            return obj
    
    def load_csv(self, file_path: str, user_id: str) -> bool:
        """Load CSV file into staging table"""
        try:
            logger.info(f"Loading CSV file for user {user_id}: {file_path}")
            
            # Determine file type and read accordingly
            file_extension = os.path.splitext(file_path)[1].lower()
            if file_extension == '.csv':
                df = pd.read_csv(file_path, dtype=str, keep_default_na=False)
            elif file_extension in ['.xlsx', '.xls']:
                df = pd.read_excel(file_path, dtype=str, keep_default_na=False)
            else:
                raise ValueError("Unsupported file format")
            
            logger.info(f"Found {len(df)} rows in file")
            
            # Map columns to database fields
            df_mapped = self.map_csv_columns(df)
            self.stats['total_rows'] = len(df_mapped)
            
            # Process in chunks
            staging_records = []
            for idx, row in df_mapped.iterrows():
                try:
                    record_data = row.to_dict()
                    
                    # Get valid fields for staging table
                    valid_fields = [column.name for column in self.staging_model.__table__.columns]
                    filtered_data = {k: v for k, v in record_data.items() if k in valid_fields}
                    
                    # Validate record
                    validation_errors = self.validate_record(filtered_data, idx + 1)
                    is_valid = len(validation_errors) == 0
                    
                    if validation_errors:
                        self.stats['validation_errors'].extend(validation_errors)
                    
                    # Create staging record
                    staging_record = self.staging_model(
                        user_id=uuid.UUID(user_id) if isinstance(user_id, str) else user_id,
                        batch_id=self.batch_id,
                        row_number=idx + 1,
                        is_valid=is_valid,
                        validation_errors=validation_errors,
                        **filtered_data
                    )
                    
                    staging_records.append(staging_record)
                    self.stats['processed_rows'] += 1
                    
                except Exception as e:
                    logger.error(f"Error processing row {idx + 1}: {e}")
                    self.stats['failed_rows'] += 1
            
            # Bulk insert staging records
            self.db.bulk_save_objects(staging_records)
            self.db.commit()
            
            logger.info(f"✅ Loaded {self.stats['processed_rows']} rows into staging for user {user_id}")
            return True
            
        except Exception as e:
            logger.error(f"❌ Error loading CSV: {e}")
            self.db.rollback()
            return False
    
    def get_stats(self) -> Dict:
        """Get processing statistics"""
        return self.stats
    
    def print_summary(self):
        """Print processing summary"""
        print("\n" + "=" * 60)
        print("ETL PROCESSING SUMMARY")
        print("=" * 60)
        print(f"Batch ID: {self.batch_id}")
        print(f"Total rows: {self.stats['total_rows']}")
        print(f"Processed rows: {self.stats['processed_rows']}")
        print(f"Failed rows: {self.stats['failed_rows']}")
        print(f"New records: {self.stats['new_records']}")
        print(f"Updated records: {self.stats['updated_records']}")
        print(f"Validation errors: {len(self.stats['validation_errors'])}")
        print("=" * 60)