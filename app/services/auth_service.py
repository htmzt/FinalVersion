# app/services/auth_service.py
from datetime import datetime, timedelta
from typing import Optional
from sqlalchemy.orm import Session
from app.models import User
from app.auth import verify_password, create_access_token, ACCESS_TOKEN_EXPIRE_MINUTES
from app.services.base_service import BaseService

class AuthService(BaseService):
    def __init__(self, db: Session):
        super().__init__(db)
    
    def authenticate_user(self, email: str, password: str) -> Optional[User]:
        """Authenticate user with email and password"""
        user = self.db.query(User).filter(User.email == email).first()
        if not user or not verify_password(password, user.password_hash):
            return None
        return user
    
    def create_user_token(self, user: User) -> dict:
        """Create access token and return user data"""
        access_token_expires = timedelta(minutes=ACCESS_TOKEN_EXPIRE_MINUTES)
        access_token = create_access_token(
            data={"sub": user.email}, 
            expires_delta=access_token_expires
        )
        
        # Update last login
        user.last_login = datetime.utcnow()
        self.db.commit()
        
        return {
            "access_token": access_token,
            "token_type": "bearer",
            "user_id": str(user.id),
            "email": user.email,
            "name": f"{user.prenom} {user.nom}"
        }