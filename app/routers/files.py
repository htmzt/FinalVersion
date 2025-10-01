from fastapi import APIRouter, Depends, File, UploadFile
from fastapi.responses import JSONResponse
from sqlalchemy.orm import Session

from app.database import get_db
from app.auth import get_current_user
from app.models import User
from app.services.file_service import FileService
from app.tasks import task_queue

router = APIRouter(prefix="/api", tags=["files"])

@router.post("/upload")
async def upload_file(
    file: UploadFile = File(...),
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    file_service = FileService(db)
    user_id = str(current_user.id)
    
    file_service.validate_file(file)
    file_path = await file_service.save_temp_file(file)
    await task_queue.put(("po_process", (file_path, user_id)))

    return JSONResponse(
        status_code=202,
        content={
            "message": "File upload accepted. Processing has started.",
            "user_id": user_id,
            "file_info": file_service.get_file_info(file)
        }
    )

@router.post("/upload-acceptance")
async def upload_acceptance_file(
    file: UploadFile = File(...),
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    file_service = FileService(db)
    user_id = str(current_user.id)
    
    file_service.validate_file(file)
    file_path = await file_service.save_temp_file(file)
    await task_queue.put(("acceptance_process", (file_path, user_id)))

    return JSONResponse(
        status_code=202,
        content={
            "message": "Acceptance file upload accepted. Processing started.",
            "user_id": user_id,
            "file_info": file_service.get_file_info(file)
        }
    )