import os
from dotenv import load_dotenv
from pydantic_settings import BaseSettings
from typing import List, Dict
import json

# Load environment variables from .env file
load_dotenv()

class Settings(BaseSettings):
    # Microsoft Graph API Configuration
    OUTLOOK_CLIENT_ID: str = os.getenv("OUTLOOK_CLIENT_ID", "")
    OUTLOOK_CLIENT_SECRET: str = os.getenv("OUTLOOK_CLIENT_SECRET", "")
    OUTLOOK_TENANT_ID: str = os.getenv("OUTLOOK_TENANT_ID", "")
    OUTLOOK_AUTH_CALLBACK_URI: str = os.getenv("OUTLOOK_AUTH_CALLBACK_URI", "http://localhost:5000/auth/callback")
    
    # Single email account (for backward compatibility)
    OUTLOOK_USER_ID: str = os.getenv("OUTLOOK_USER_ID", "")
    
    # Multiple email accounts configuration
    # Can be set as JSON string in environment variable
    # Example: OUTLOOK_USER_IDS='["user1@domain.com", "user2@domain.com", "user3@domain.com"]'
    OUTLOOK_USER_IDS: str = os.getenv("OUTLOOK_USER_IDS", "")
    
    # Google Vertex AI Configuration
    GEMINI_API_KEY: str = os.getenv("GEMINI_API_KEY", "")
    GEMINI_PROJECT_ID: str = os.getenv("GEMINI_PROJECT_ID", "")
    
    # Processing Configuration
    # Processing interval in seconds (default: 1 hour)
    PROCESSING_INTERVAL: int = int(os.getenv("PROCESSING_INTERVAL", "3600"))
    EMAIL_FOLDER: str = os.getenv("EMAIL_FOLDER", "Inbox")
    BATCH_SIZE: int = int(os.getenv("BATCH_SIZE", "10"))
    
    # Batch size per account (when processing multiple accounts)
    BATCH_SIZE_PER_ACCOUNT: int = int(os.getenv("BATCH_SIZE_PER_ACCOUNT", "5"))
    
    # Maximum concurrent accounts to process
    MAX_CONCURRENT_ACCOUNTS: int = int(os.getenv("MAX_CONCURRENT_ACCOUNTS", "3"))
    
    # Logging Configuration
    LOG_LEVEL: str = os.getenv("LOG_LEVEL", "INFO")
    
    # API Configuration
    API_HOST: str = os.getenv("API_HOST", "127.0.0.1")
    API_PORT: int = int(os.getenv("API_PORT", "8000"))
    
    class Config:
        env_file = ".env"
        extra = "allow"  # Allow extra fields

    def get_email_accounts(self) -> List[str]:
        """Get list of email accounts to process."""
        email_accounts = []
        
        # Try to parse multiple email accounts from JSON
        if self.OUTLOOK_USER_IDS:
            try:
                email_accounts = json.loads(self.OUTLOOK_USER_IDS)
                if not isinstance(email_accounts, list):
                    raise ValueError("OUTLOOK_USER_IDS must be a JSON array")
            except json.JSONDecodeError as e:
                raise ValueError(f"Invalid JSON in OUTLOOK_USER_IDS: {e}")
        
        # Fallback to single email account for backward compatibility
        elif self.OUTLOOK_USER_ID:
            email_accounts = [self.OUTLOOK_USER_ID]
        
        if not email_accounts:
            raise ValueError("No email accounts configured. Set either OUTLOOK_USER_ID or OUTLOOK_USER_IDS")
        
        return email_accounts

# Create global settings instance
settings = Settings()

# Validation
def validate_settings():
    required_vars = [
        "OUTLOOK_CLIENT_ID",
        "OUTLOOK_CLIENT_SECRET", 
        "OUTLOOK_TENANT_ID",
        "GEMINI_API_KEY",
        "GEMINI_PROJECT_ID"
    ]
    
    missing_vars = [var for var in required_vars if not getattr(settings, var)]
    
    if missing_vars:
        raise ValueError(f"Missing required environment variables: {', '.join(missing_vars)}")
    
    # Validate that at least one email account is configured
    try:
        email_accounts = settings.get_email_accounts()
        if not email_accounts:
            raise ValueError("No email accounts configured")
    except Exception as e:
        raise ValueError(f"Email account configuration error: {e}")

# Validate settings on import
validate_settings()