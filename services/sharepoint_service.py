"""SharePoint service layer - Business logic for SharePoint operations.

This service layer provides all SharePoint functionality independent of the MCP protocol,
making it easier to test and maintain.
"""

import logging
from typing import Dict, Any, Optional
from auth.sharepoint_auth import SharePointContext
from utils.graph_client import GraphClient
from utils.document_processor import DocumentProcessor
from utils.content_generator import ContentGenerator

logger = logging.getLogger("sharepoint_service")


def parse_site_url(site_url: str) -> tuple[str, str]:
    """Parse SharePoint site URL into domain and site name.

    Args:
        site_url: Full SharePoint site URL

    Returns:
        Tuple of (domain, site_name)
    """
    # Remove protocol
    url_without_protocol = site_url.replace('https://', '').replace('http://', '')

    # Split into parts
    parts = url_without_protocol.split('/')
    domain = parts[0]

    # Get site name (after /sites/)
    if len(parts) > 2 and parts[1] == 'sites':
        site_name = parts[2]
    else:
        site_name = ""

    return domain, site_name


class SharePointService:
    """Service class for SharePoint operations."""

    def __init__(self, context: SharePointContext):
        """Initialize service with SharePoint context.

        Args:
            context: SharePoint authentication context
        """
        self.context = context
        self.graph_client = GraphClient(context)
        self.document_processor = DocumentProcessor()
        self.content_generator = ContentGenerator()

    # Site Operations

    async def get_site_info(self, site_url: str) -> Dict[str, Any]:
        """Get basic information about the SharePoint site."""
        domain, site_name = parse_site_url(site_url)
        site_info = await self.graph_client.get_site_info(domain, site_name)

        return {
            "name": site_info.get("displayName", "Unknown"),
            "description": site_info.get("description", "No description"),
            "created": site_info.get("createdDateTime", "Unknown"),
            "last_modified": site_info.get("lastModifiedDateTime", "Unknown"),
            "web_url": site_info.get("webUrl", site_url),
            "id": site_info.get("id", "Unknown")
        }

    async def list_document_libraries(self, site_url: str) -> Dict[str, Any]:
        """List all document libraries in a site."""
        domain, site_name = parse_site_url(site_url)
        result = await self.graph_client.list_document_libraries(domain, site_name)

        libraries = result.get("value", [])
        formatted_libraries = [{
            "name": lib.get("name", "Unknown"),
            "id": lib.get("id", "Unknown"),
            "webUrl": lib.get("webUrl", ""),
            "description": lib.get("description", ""),
            "created": lib.get("createdDateTime", "Unknown"),
            "modified": lib.get("lastModifiedDateTime", "Unknown")
        } for lib in libraries]

        return {"libraries": formatted_libraries, "count": len(formatted_libraries)}

    async def search_sharepoint(self, site_url: str, query: str) -> Dict[str, Any]:
        """Search content in the SharePoint site."""
        domain, site_name = parse_site_url(site_url)
        return await self.graph_client.search_sharepoint(domain, site_name, query)

    async def create_sharepoint_site(self, display_name: str, alias: str, description: str = "") -> Dict[str, Any]:
        """Create a new SharePoint site."""
        return await self.graph_client.create_site(display_name, alias, description)

    # Folder Operations

    async def list_folders(self, site_id: str, drive_id: str, folder_path: str = "") -> Dict[str, Any]:
        """List all folders in a document library location."""
        result = await self.graph_client.list_drive_items(site_id, drive_id, folder_path, 'folder')

        folders = result.get("value", [])
        formatted_folders = [{
            "name": folder.get("name", "Unknown"),
            "id": folder.get("id", "Unknown"),
            "webUrl": folder.get("webUrl", ""),
            "created": folder.get("createdDateTime", "Unknown"),
            "modified": folder.get("lastModifiedDateTime", "Unknown"),
            "size": folder.get("size", 0),
            "childCount": folder.get("folder", {}).get("childCount", 0)
        } for folder in folders]

        return {"folders": formatted_folders, "count": len(formatted_folders)}

    async def create_folder(self, site_id: str, drive_id: str, folder_path: str) -> Dict[str, Any]:
        """Create a new folder in a document library."""
        return await self.graph_client.create_folder_in_library(site_id, drive_id, folder_path)

    async def delete_folder(self, site_id: str, drive_id: str, folder_id: str) -> Dict[str, Any]:
        """Delete a folder from a document library."""
        return await self.graph_client.delete_drive_item(site_id, drive_id, folder_id)

    async def get_folder_tree(self, site_id: str, drive_id: str, folder_path: str = "", max_depth: int = 10) -> Dict[str, Any]:
        """Get a recursive tree structure of folders."""
        return await self.graph_client.get_folder_tree(site_id, drive_id, folder_path, max_depth)

    # Document Operations

    async def list_documents(self, site_id: str, drive_id: str, folder_path: str = "") -> Dict[str, Any]:
        """List all documents in a document library location."""
        result = await self.graph_client.list_drive_items(site_id, drive_id, folder_path, 'file')

        documents = result.get("value", [])
        formatted_documents = [{
            "name": doc.get("name", "Unknown"),
            "id": doc.get("id", "Unknown"),
            "webUrl": doc.get("webUrl", ""),
            "created": doc.get("createdDateTime", "Unknown"),
            "modified": doc.get("lastModifiedDateTime", "Unknown"),
            "size": doc.get("size", 0),
            "mimeType": doc.get("file", {}).get("mimeType", "Unknown"),
            "createdBy": doc.get("createdBy", {}).get("user", {}).get("displayName", "Unknown"),
            "modifiedBy": doc.get("lastModifiedBy", {}).get("user", {}).get("displayName", "Unknown")
        } for doc in documents]

        return {"documents": formatted_documents, "count": len(formatted_documents)}

    async def get_document_content(self, site_id: str, drive_id: str, item_id: str, filename: str) -> Dict[str, Any]:
        """Get and process content from a SharePoint document."""
        content = await self.graph_client.get_document_content(site_id, drive_id, item_id)
        processed_content = self.document_processor.process_document(content, filename)
        return processed_content

    async def upload_document(self, site_id: str, drive_id: str, folder_path: str,
                            file_name: str, file_content: bytes, content_type: str) -> Dict[str, Any]:
        """Upload a new document to SharePoint."""
        return await self.graph_client.upload_document(
            site_id, drive_id, folder_path, file_name, file_content, content_type
        )

    async def update_document(self, site_id: str, drive_id: str, item_id: str,
                            file_content: bytes, content_type: str = "application/octet-stream") -> Dict[str, Any]:
        """Update the content of an existing document."""
        return await self.graph_client.update_document_content(
            site_id, drive_id, item_id, file_content, content_type
        )

    async def delete_document(self, site_id: str, drive_id: str, item_id: str) -> Dict[str, Any]:
        """Delete a document from a document library."""
        return await self.graph_client.delete_drive_item(site_id, drive_id, item_id)

    # List Operations

    async def create_intelligent_list(self, site_id: str, purpose: str, display_name: str) -> Dict[str, Any]:
        """Create a SharePoint list with AI-optimized schema."""
        return await self.graph_client.create_intelligent_list(site_id, purpose, display_name)

    async def create_list_item(self, site_id: str, list_id: str, fields: Dict[str, Any]) -> Dict[str, Any]:
        """Add new items to SharePoint lists."""
        return await self.graph_client.create_list_item(site_id, list_id, fields)

    async def update_list_item(self, site_id: str, list_id: str, item_id: str, fields: Dict[str, Any]) -> Dict[str, Any]:
        """Update existing list items."""
        return await self.graph_client.update_list_item(site_id, list_id, item_id, fields)

    # Content Creation

    async def create_advanced_document_library(self, site_id: str, display_name: str, doc_type: str = "general") -> Dict[str, Any]:
        """Create a document library with advanced metadata settings."""
        return await self.graph_client.create_advanced_document_library(site_id, display_name, doc_type)

    async def create_modern_page(self, site_id: str, name: str, title: str = "",
                                purpose: str = "general", audience: str = "general") -> Dict[str, Any]:
        """Create a modern SharePoint page with beautiful layout."""
        if not title:
            title = self.content_generator.generate_page_title(purpose, name)

        return await self.graph_client.create_page(site_id, name, title)

    async def create_news_post(self, site_id: str, title: str, description: str = "", content: str = "") -> Dict[str, Any]:
        """Create a news post in a SharePoint site."""
        return await self.graph_client.create_news_post(site_id, title, description, content)
