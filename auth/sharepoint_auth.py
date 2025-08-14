"""SharePoint authentication handler module - Simplified version using environment variables."""

from dataclasses import dataclass
from datetime import datetime, timedelta
import os
import logging

# No longer need SHAREPOINT_CONFIG as site_url is passed per tool

# Set up logging
logging.basicConfig(level=logging.DEBUG, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s")
logger = logging.getLogger("sharepoint_auth")

@dataclass
class SharePointContext:
    """Context object for SharePoint connection."""
    access_token: str
    token_expiry: datetime
    graph_url: str = "https://graph.microsoft.com/v1.0"

    @property
    def headers(self) -> dict[str, str]:
        """Get authorization headers for API calls."""
        # Log token preview for debugging
        token_preview = f"{self.access_token[:10]}...{self.access_token[-10:]}" if self.access_token and len(self.access_token) > 20 else self.access_token
        logger.debug(f"Using token (preview): {token_preview}")
        
        return {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json",
        }

    def is_token_valid(self) -> bool:
        """Check if the access token is still valid."""
        # Add safety check to handle None expiry
        if not self.token_expiry:
            return False
        is_valid = datetime.now() < self.token_expiry
        logger.debug(f"Token valid: {is_valid}, Expires: {self.token_expiry}")
        return is_valid

    def test_connection(self) -> bool:
        """Test the connection to Microsoft Graph API."""
        try:
            import requests
            # Simple test - just check if we can access the Graph API /me endpoint
            test_url = f"{self.graph_url}/me"
            logger.debug(f"Testing connection to: {test_url}")
            
            response = requests.get(test_url, headers=self.headers)
            
            if response.status_code != 200:
                logger.error(f"Connection test failed: HTTP {response.status_code} - {response.text}")
                return False
                
            logger.info(f"Connection test successful: {response.status_code}")
            return True
        except Exception as e:
            logger.error(f"Error during connection test: {e}")
            return False


async def get_auth_context() -> SharePointContext:
    """Get SharePoint authentication context from environment variables."""
    
    # Get access token from environment variable
    access_token = os.getenv("ACCESS_TOKEN")
    if not access_token:
        logger.error("ACCESS_TOKEN environment variable not set")
        raise ValueError("ACCESS_TOKEN environment variable is required")
    
    logger.info("Using access token from environment variable")
    
    # Get token expiry from environment or default to 1 hour
    expires_in = int(os.getenv("TOKEN_EXPIRES_IN", "3600"))
    expiry = datetime.now() + timedelta(seconds=expires_in)
    logger.info(f"Token expires at {expiry}")
    
    # Return auth context
    context = SharePointContext(
        access_token=access_token,
        token_expiry=expiry
    )
    
    # Test connection immediately
    logger.info("Testing connection with provided token...")
    if not context.test_connection():
        logger.warning("Connection test failed, but continuing anyway...")
    
    return context


async def refresh_token_if_needed(context: SharePointContext) -> None:
    """Refresh token if needed - simplified version just checks validity."""
    if not context.is_token_valid():
        logger.warning("Token appears to be expired based on TOKEN_EXPIRES_IN")
        # In this simplified version, we can't refresh the token
        # The hub should provide a new token before expiry