"""pytest configuration for SharePoint MCP integration tests."""

import sys
from pathlib import Path
from dotenv import load_dotenv

# Add project root to Python path so tests can import modules
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

# Load environment variables from .env file
load_dotenv(project_root / ".env")
