"""Service layer for SharePoint MCP Server.

This module contains the business logic for SharePoint operations,
separated from the MCP protocol layer for better testability.
"""

from .sharepoint_service import SharePointService

__all__ = ['SharePointService']
