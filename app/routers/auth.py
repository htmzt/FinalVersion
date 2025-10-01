
from fastapi import APIRouter, Depends, HTTPException, status
from fastapi.security import OAuth2PasswordRequestForm
from sqlalchemy.orm import Session
from pydantic import BaseModel, EmailStr, Field, validator
from typing import Optional
from datetime import timedelta, datetime
import logging
import re

from app.database import get_db
from app.models import User
from app.auth import (
    get_password_hash, 
    verify_password, 
    create_access_token,
    get_current_user,
    ACCESS_TOKEN_EXPIRE_MINUTES
)

logger = logging.getLogger(__name__)

# FIXED: Changed prefix to match OAuth2 tokenUrl
router = APIRouter(prefix="/api/auth", tags=["authentication"])

# ==================== Pydantic Schemas ====================

class UserRegistration(BaseModel):
    email: EmailStr
    password: str = Field(..., min_length=8, max_length=100)
    prenom: str = Field(..., min_length=2, max_length=100)
    nom: str = Field(..., min_length=2, max_length=100)
    company_name: str = Field(..., min_length=2, max_length=255)
    company_logo: Optional[str] = None
    
    @validator('password')
    def validate_password(cls, v):
        """
        Password must contain:
        - At least 8 characters
        - At least one uppercase letter
        - At least one lowercase letter
        - At least one digit
        - At least one special character
        """
        if len(v) < 8:
            raise ValueError('Password must be at least 8 characters long')
        if not re.search(r'[A-Z]', v):
            raise ValueError('Password must contain at least one uppercase letter')
        if not re.search(r'[a-z]', v):
            raise ValueError('Password must contain at least one lowercase letter')
        if not re.search(r'\d', v):
            raise ValueError('Password must contain at least one digit')
        if not re.search(r'[!@#$%^&*(),.?":{}|<>]', v):
            raise ValueError('Password must contain at least one special character')
        return v
    
    @validator('email')
    def validate_email(cls, v):
        """Normalize email to lowercase"""
        return v.lower().strip()
    
    class Config:
        json_schema_extra = {
            "example": {
                "email": "admin@example.com",
                "password": "Admin123!",
                "prenom": "Test",
                "nom": "Test",
                "company_name": "SIB",
                "company_logo": "https://example.com/logo.png"
            }
        }


class UserLogin(BaseModel):
    email: EmailStr
    password: str
    
    class Config:
        json_schema_extra = {
            "example": {
                "email": "admin@example.com",
                "password": "Admin123!"
            }
        }


class TokenResponse(BaseModel):
    access_token: str
    token_type: str = "bearer"
    expires_in: int
    user_id: str
    email: str
    name: str
    company_name: str


class UserProfile(BaseModel):
    id: str
    email: str
    prenom: str
    nom: str
    company_name: str
    company_logo: Optional[str]
    is_active: bool
    email_verified: bool
    created_at: datetime
    last_login: Optional[datetime]
    
    class Config:
        from_attributes = True


class PasswordChange(BaseModel):
    old_password: str
    new_password: str = Field(..., min_length=8, max_length=100)
    
    @validator('new_password')
    def validate_password(cls, v):
        """Same validation as registration"""
        if len(v) < 8:
            raise ValueError('Password must be at least 8 characters long')
        if not re.search(r'[A-Z]', v):
            raise ValueError('Password must contain at least one uppercase letter')
        if not re.search(r'[a-z]', v):
            raise ValueError('Password must contain at least one lowercase letter')
        if not re.search(r'\d', v):
            raise ValueError('Password must contain at least one digit')
        if not re.search(r'[!@#$%^&*(),.?":{}|<>]', v):
            raise ValueError('Password must contain at least one special character')
        return v


class PasswordResetRequest(BaseModel):
    email: EmailStr


class MessageResponse(BaseModel):
    message: str


# ==================== Route Handlers ====================

@router.post("/register", response_model=TokenResponse, status_code=status.HTTP_201_CREATED)
async def register(
    user_data: UserRegistration,
    db: Session = Depends(get_db)
):
    """
    Register a new user account
    
    **Requirements:**
    - Unique email address
    - Strong password (8+ chars, uppercase, lowercase, digit, special char)
    - First name (prenom) and last name (nom)
    - Company name
    
    **Returns:**
    - JWT access token
    - User profile information
    
    **Example:**
    ```json
    {
      "email": "admin@example.com",
      "password": "Admin123!",
      "prenom": "Test",
      "nom": "Test",
      "company_name": "SIB"
    }
    ```
    """
    try:
        # Check if email already exists
        existing_user = db.query(User).filter(User.email == user_data.email).first()
        if existing_user:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="Email already registered. Please use a different email or login."
            )
        
        # Create new user
        new_user = User(
            email=user_data.email,
            password_hash=get_password_hash(user_data.password),
            prenom=user_data.prenom,
            nom=user_data.nom,
            company_name=user_data.company_name,
            company_logo=user_data.company_logo,
            is_active=True,
            email_verified=False 
        )
        
        db.add(new_user)
        db.commit()
        db.refresh(new_user)
        
        logger.info(f"✅ New user registered: {new_user.email}")
        
        # Generate JWT token
        access_token_expires = timedelta(minutes=ACCESS_TOKEN_EXPIRE_MINUTES)
        access_token = create_access_token(
            data={"sub": new_user.email},
            expires_delta=access_token_expires
        )
        
        # Update last login
        new_user.last_login = datetime.utcnow()
        db.commit()
        
        return TokenResponse(
            access_token=access_token,
            token_type="bearer",
            expires_in=ACCESS_TOKEN_EXPIRE_MINUTES * 60,  # in seconds
            user_id=str(new_user.id),
            email=new_user.email,
            name=f"{new_user.prenom} {new_user.nom}",
            company_name=new_user.company_name
        )
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"❌ Registration error: {str(e)}")
        db.rollback()
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="An error occurred during registration. Please try again."
        )


@router.post("/login", response_model=TokenResponse)
async def login(
    user_credentials: UserLogin,
    db: Session = Depends(get_db)
):
    """
    Login with email and password
    
    **Returns:**
    - JWT access token (valid for 30 minutes)
    - User profile information
    
    **Example:**
    ```json
    {
      "email": "admin@example.com",
      "password": "Admin123!"
    }
    ```
    """
    try:
        # Find user by email
        user = db.query(User).filter(User.email == user_credentials.email.lower()).first()
        
        if not user:
            raise HTTPException(
                status_code=status.HTTP_401_UNAUTHORIZED,
                detail="Incorrect email or password",
                headers={"WWW-Authenticate": "Bearer"},
            )
        
        # Verify password
        if not verify_password(user_credentials.password, user.password_hash):
            raise HTTPException(
                status_code=status.HTTP_401_UNAUTHORIZED,
                detail="Incorrect email or password",
                headers={"WWW-Authenticate": "Bearer"},
            )
        
        # Check if account is active
        if not user.is_active:
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail="Account is deactivated. Please contact support."
            )
        
        # Generate JWT token
        access_token_expires = timedelta(minutes=ACCESS_TOKEN_EXPIRE_MINUTES)
        access_token = create_access_token(
            data={"sub": user.email},
            expires_delta=access_token_expires
        )
        
        # Update last login
        user.last_login = datetime.utcnow()
        db.commit()
        
        logger.info(f"✅ User logged in: {user.email}")
        
        return TokenResponse(
            access_token=access_token,
            token_type="bearer",
            expires_in=ACCESS_TOKEN_EXPIRE_MINUTES * 60,
            user_id=str(user.id),
            email=user.email,
            name=f"{user.prenom} {user.nom}",
            company_name=user.company_name
        )
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"❌ Login error: {str(e)}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="An error occurred during login. Please try again."
        )


@router.post("/token", response_model=TokenResponse)
async def login_for_swagger(
    form_data: OAuth2PasswordRequestForm = Depends(),
    db: Session = Depends(get_db)
):
    """
    OAuth2 compatible token endpoint for Swagger UI
    
    Use this endpoint with the "Authorize" button in Swagger UI.
    
    **Form Data:**
    - `username`: Your email address
    - `password`: Your password
    """
    try:
        user = db.query(User).filter(User.email == form_data.username.lower()).first()
        
        if not user or not verify_password(form_data.password, user.password_hash):
            raise HTTPException(
                status_code=status.HTTP_401_UNAUTHORIZED,
                detail="Incorrect email or password",
                headers={"WWW-Authenticate": "Bearer"},
            )
        
        if not user.is_active:
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail="Account is deactivated"
            )
        
        access_token_expires = timedelta(minutes=ACCESS_TOKEN_EXPIRE_MINUTES)
        access_token = create_access_token(
            data={"sub": user.email},
            expires_delta=access_token_expires
        )
        
        user.last_login = datetime.utcnow()
        db.commit()
        
        return TokenResponse(
            access_token=access_token,
            token_type="bearer",
            expires_in=ACCESS_TOKEN_EXPIRE_MINUTES * 60,
            user_id=str(user.id),
            email=user.email,
            name=f"{user.prenom} {user.nom}",
            company_name=user.company_name
        )
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"❌ Token error: {str(e)}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="An error occurred. Please try again."
        )


@router.get("/me", response_model=UserProfile)
async def get_current_user_profile(
    current_user: User = Depends(get_current_user)
):
    """
    Get current authenticated user profile
    
    **Requires:** Valid JWT token in Authorization header
    
    **Returns:** Complete user profile information
    """
    return UserProfile(
        id=str(current_user.id),
        email=current_user.email,
        prenom=current_user.prenom,
        nom=current_user.nom,
        company_name=current_user.company_name,
        company_logo=current_user.company_logo,
        is_active=current_user.is_active,
        email_verified=current_user.email_verified,
        created_at=current_user.created_at,
        last_login=current_user.last_login
    )


@router.post("/change-password", response_model=MessageResponse)
async def change_password(
    password_data: PasswordChange,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    """
    Change password for authenticated user
    
    **Requires:** 
    - Valid JWT token
    - Current password
    - New strong password
    
    **Example:**
    ```json
    {
      "old_password": "OldPass123!",
      "new_password": "NewSecurePass456!"
    }
    ```
    """
    try:
        # Verify old password
        if not verify_password(password_data.old_password, current_user.password_hash):
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="Current password is incorrect"
            )
        
        # Check new password is different
        if password_data.old_password == password_data.new_password:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="New password must be different from current password"
            )
        
        # Update password
        current_user.password_hash = get_password_hash(password_data.new_password)
        current_user.updated_at = datetime.utcnow()
        db.commit()
        
        logger.info(f"✅ Password changed for user: {current_user.email}")
        
        return MessageResponse(
            message="Password changed successfully. Please login with your new password."
        )
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"❌ Password change error: {str(e)}")
        db.rollback()
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="An error occurred while changing password"
        )


@router.post("/logout", response_model=MessageResponse)
async def logout(current_user: User = Depends(get_current_user)):
    """
    Logout current user
    
    **Note:** Since JWT tokens are stateless, actual logout happens client-side
    by deleting the token. This endpoint is for logging purposes.
    
    **Client-side actions required:**
    1. Delete token from localStorage/sessionStorage
    2. Clear authentication headers
    3. Redirect to login page
    """
    logger.info(f"🚪 User logged out: {current_user.email}")
    
    return MessageResponse(
        message="Logged out successfully. Please delete your token on the client side."
    )




@router.get("/validate-token")
async def validate_token(current_user: User = Depends(get_current_user)):
    """
    Validate if current JWT token is valid
    
    **Returns:** Success if token is valid
    **Raises:** 401 if token is invalid or expired
    """
    return {
        "valid": True,
        "user_id": str(current_user.id),
        "email": current_user.email
    }