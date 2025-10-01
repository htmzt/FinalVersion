# app/services/file_service.py
import os
import tempfile
from typing import Dict, Any
from fastapi import UploadFile, HTTPException
from app.services.base_service import BaseService
from app.processors.po_processor import process_user_csv
from app.processors.acceptance_processor import process_user_acceptance_csv
import logging

logger = logging.getLogger(__name__)

class FileService(BaseService):
    
    ALLOWED_EXTENSIONS = {'.csv', '.xlsx', '.xls'}
    MAX_FILE_SIZE = 50 * 1024 * 1024  # 50MB
    
    def validate_file(self, file: UploadFile) -> None:
        """Validate uploaded file"""
        # Check file extension
        file_extension = os.path.splitext(file.filename)[1].lower()
        if file_extension not in self.ALLOWED_EXTENSIONS:
            raise HTTPException(
                status_code=400, 
                detail=f"Unsupported file format. Please upload a CSV or Excel file. Got: {file_extension}"
            )
        
        # Check file size (if available)
        if hasattr(file, 'size') and file.size and file.size > self.MAX_FILE_SIZE:
            raise HTTPException(
                status_code=400,
                detail=f"File size too large. Maximum allowed size is {self.MAX_FILE_SIZE // (1024*1024)}MB"
            )
    
    async def save_temp_file(self, file: UploadFile) -> str:
        """Save uploaded file to temporary location"""
        file_extension = os.path.splitext(file.filename)[1].lower()
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=file_extension) as tmp_file:
            content = await file.read()
            tmp_file.write(content)
            return tmp_file.name
    
    def process_po_file(self, file_path: str, user_id: str) -> Dict[str, Any]:
        """Process PO file"""
        try:
            result = process_user_csv(file_path, user_id)
            return {
                "success": True,
                "message": "PO file processed successfully",
                "stats": result.get('stats', {}),
                "batch_id": result.get('batch_id')
            }
        except Exception as e:
            logger.error(f"Error processing PO file: {str(e)}")
            return {
                "success": False,
                "message": f"Error processing PO file: {str(e)}",
                "stats": {}
            }
        finally:
            self._cleanup_temp_file(file_path)
    
    def process_acceptance_file(self, file_path: str, user_id: str) -> Dict[str, Any]:
        """Process Acceptance file"""
        try:
            result = process_user_acceptance_csv(file_path, user_id)
            return {
                "success": True,
                "message": "Acceptance file processed successfully",
                "stats": result.get('stats', {}),
                "batch_id": result.get('batch_id')
            }
        except Exception as e:
            logger.error(f"Error processing Acceptance file: {str(e)}")
            return {
                "success": False,
                "message": f"Error processing Acceptance file: {str(e)}",
                "stats": {}
            }
        finally:
            self._cleanup_temp_file(file_path)
    
    def _cleanup_temp_file(self, file_path: str) -> None:
        """Clean up temporary file"""
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                logger.info(f"Cleaned up temporary file: {file_path}")
        except Exception as e:
            logger.warning(f"Failed to cleanup file {file_path}: {e}")
    
    def get_file_info(self, file: UploadFile) -> Dict[str, Any]:
        """Get file information"""
        file_extension = os.path.splitext(file.filename)[1].lower()
        
        return {
            "filename": file.filename,
            "extension": file_extension,
            "content_type": file.content_type,
            "size": getattr(file, 'size', 'Unknown')
        }