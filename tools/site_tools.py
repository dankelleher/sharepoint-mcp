"""SharePoint site information tools."""

import json
import logging
from typing import Dict, Any, List, Optional

from mcp.server.fastmcp import FastMCP, Context

from auth.sharepoint_auth import refresh_token_if_needed
from utils.graph_client import GraphClient
from utils.document_processor import DocumentProcessor
from utils.content_generator import ContentGenerator

# Set up logging
logger = logging.getLogger("sharepoint_tools")

def parse_site_url(site_url: str) -> tuple[str, str]:
    """Parse a SharePoint site URL into domain and site name.
    
    Args:
        site_url: Full SharePoint site URL (e.g., https://example.sharepoint.com/sites/test)
    
    Returns:
        Tuple of (domain, site_name)
    """
    # Remove https:// prefix
    url_parts = site_url.replace("https://", "").replace("http://", "")
    
    # Split by forward slash
    parts = url_parts.split("/")
    domain = parts[0]
    
    # Extract site name (if present)
    if len(parts) > 2:
        site_name = parts[2]
    else:
        site_name = "root"
    
    return domain, site_name

def register_site_tools(mcp: FastMCP):
    """Register SharePoint site tools with the MCP server."""
    
    @mcp.tool()
    async def get_site_info(ctx: Context, site_url: str) -> str:
        """Get basic information about the SharePoint site.
        
        Args:
            site_url: The SharePoint site URL (e.g., https://example.sharepoint.com/sites/test)
        """
        logger.info(f"Tool called: get_site_info for {site_url}")
        
        try:
            # Get authentication context from context object
            sp_ctx = ctx.request_context.lifespan_context
            
            # Refresh token if needed
            await refresh_token_if_needed(sp_ctx)
            
            # Create Graph client
            graph_client = GraphClient(sp_ctx)
            
            # Parse site URL
            domain, site_name = parse_site_url(site_url)
            
            logger.info(f"Getting info for site: {site_name} in domain: {domain}")
            
            # Get site info using Graph client
            site_info = await graph_client.get_site_info(domain, site_name)
            
            # Format response
            result = {
                "name": site_info.get("displayName", "Unknown"),
                "description": site_info.get("description", "No description"),
                "created": site_info.get("createdDateTime", "Unknown"),
                "last_modified": site_info.get("lastModifiedDateTime", "Unknown"),
                "web_url": site_info.get("webUrl", site_url),
                "id": site_info.get("id", "Unknown")
            }
            
            logger.info(f"Successfully retrieved site info for: {result['name']}")
            return json.dumps(result, indent=2)
            
        except Exception as e:
            logger.error(f"Error in get_site_info: {str(e)}")
            return f"Error accessing SharePoint: {str(e)}"
            
    @mcp.tool()
    async def list_document_libraries(ctx: Context, site_url: str) -> str:
        """List all document libraries in the SharePoint site.
        
        Args:
            site_url: The SharePoint site URL (e.g., https://example.sharepoint.com/sites/test)
        """
        logger.info(f"Tool called: list_document_libraries for {site_url}")
        
        try:
            # Get authentication context from context object
            sp_ctx = ctx.request_context.lifespan_context
            
            # Refresh token if needed
            await refresh_token_if_needed(sp_ctx)
            
            # Create Graph client
            graph_client = GraphClient(sp_ctx)
            
            # Parse site URL
            domain, site_name = parse_site_url(site_url)
            
            logger.info(f"Listing document libraries for site: {site_name} in domain: {domain}")
            
            # List document libraries using Graph client
            result = await graph_client.list_document_libraries(domain, site_name)
            
            # Extract drive information from response
            drives = result.get("value", [])
            formatted_drives = [{
                    "name": drive.get("name", "Unknown"),
                    "description": drive.get("description", "No description"),
                    "web_url": drive.get("webUrl", "Unknown"),
                    "drive_type": drive.get("driveType", "Unknown"),
                    "id": drive.get("id", "Unknown")
                } for drive in drives]
            
            logger.info(f"Successfully retrieved {len(formatted_drives)} document libraries")
            return json.dumps(formatted_drives, indent=2)
            
        except Exception as e:
            logger.error(f"Error in list_document_libraries: {str(e)}")
            return f"Error listing document libraries: {str(e)}"
    
    @mcp.tool()
    async def search_sharepoint(ctx: Context, site_url: str, query: str) -> str:
        """Search content in the SharePoint site.
        
        Args:
            site_url: The SharePoint site URL (e.g., https://example.sharepoint.com/sites/test)
            query: Search query string
        """
        logger.info(f"Tool called: search_sharepoint for {site_url} with query: {query}")
        
        try:
            # Get authentication context from context object
            sp_ctx = ctx.request_context.lifespan_context
            
            # Refresh token if needed
            await refresh_token_if_needed(sp_ctx)
            
            # Create Graph client
            graph_client = GraphClient(sp_ctx)
            
            # Parse site URL
            domain, site_name = parse_site_url(site_url)
            
            # Search SharePoint
            results = await graph_client.search_sharepoint(domain, site_name, query)
            
            logger.info(f"Search completed with {len(results.get('value', []))} results")
            return json.dumps(results, indent=2)
            
        except Exception as e:
            logger.error(f"Error in search_sharepoint: {str(e)}")
            return f"Error searching SharePoint: {str(e)}"
    
    @mcp.tool()
    async def create_sharepoint_site(
        ctx: Context,
        display_name: str,
        alias: str,
        description: str = ""
    ) -> str:
        """Create a new SharePoint site.
        
        Args:
            display_name: Display name of the site
            alias: Site alias (used in URL)
            description: Site description (optional)
        """
        logger.info(f"Tool called: create_sharepoint_site with name: {display_name}, alias: {alias}")
        
        try:
            # Get authentication context from context object
            sp_ctx = ctx.request_context.lifespan_context
            
            # Refresh token if needed
            await refresh_token_if_needed(sp_ctx)
            
            # Create Graph client
            graph_client = GraphClient(sp_ctx)
            
            # Create the site
            result = await graph_client.create_site(display_name, alias, description)
            
            logger.info(f"Successfully created site: {display_name}")
            return json.dumps(result, indent=2)
            
        except Exception as e:
            logger.error(f"Error in create_sharepoint_site: {str(e)}")
            return f"Error creating SharePoint site: {str(e)}"
    
    @mcp.tool()
    async def create_intelligent_list(
        ctx: Context,
        site_url: str,
        site_id: str,
        purpose: str,
        display_name: str
    ) -> str:
        """Create a SharePoint list with AI-optimized schema based on its purpose.
        
        Args:
            site_url: The SharePoint site URL (e.g., https://example.sharepoint.com/sites/test)
            site_id: ID of the site
            purpose: Purpose of the list (projects, events, tasks, contacts, documents)
            display_name: Display name for the list
        """
        logger.info(f"Tool called: create_intelligent_list for site {site_url}")
        
        try:
            # Get authentication context from context object
            sp_ctx = ctx.request_context.lifespan_context
            
            # Refresh token if needed
            await refresh_token_if_needed(sp_ctx)
            
            # Create Graph client and content generator
            graph_client = GraphClient(sp_ctx)
            generator = ContentGenerator()
            
            # Generate list schema based on purpose
            list_config = generator.generate_list_schema(purpose, display_name)
            
            # Create the list
            result = await graph_client.create_list(site_id, list_config)
            
            logger.info(f"Successfully created intelligent list: {display_name}")
            return json.dumps(result, indent=2)
            
        except Exception as e:
            logger.error(f"Error in create_intelligent_list: {str(e)}")
            return f"Error creating list: {str(e)}"
    
    @mcp.tool()
    async def create_list_item(
        ctx: Context,
        site_url: str,
        site_id: str,
        list_id: str,
        fields: Dict[str, Any]
    ) -> str:
        """Create a new item in a SharePoint list.
        
        Args:
            site_url: The SharePoint site URL (e.g., https://example.sharepoint.com/sites/test)
            site_id: ID of the site
            list_id: ID of the list
            fields: Dictionary of field names and values
        """
        logger.info(f"Tool called: create_list_item for site {site_url}")
        
        try:
            # Get authentication context from context object
            sp_ctx = ctx.request_context.lifespan_context
            
            # Refresh token if needed
            await refresh_token_if_needed(sp_ctx)
            
            # Create Graph client
            graph_client = GraphClient(sp_ctx)
            
            # Create the list item
            result = await graph_client.create_list_item(site_id, list_id, fields)
            
            logger.info(f"Successfully created list item in list {list_id}")
            return json.dumps(result, indent=2)
            
        except Exception as e:
            logger.error(f"Error in create_list_item: {str(e)}")
            return f"Error creating list item: {str(e)}"
    
    @mcp.tool()
    async def update_list_item(
        ctx: Context,
        site_url: str,
        site_id: str,
        list_id: str,
        item_id: str,
        fields: Dict[str, Any]
    ) -> str:
        """Update an existing item in a SharePoint list.
        
        Args:
            site_url: The SharePoint site URL (e.g., https://example.sharepoint.com/sites/test)
            site_id: ID of the site
            list_id: ID of the list
            item_id: ID of the list item
            fields: Dictionary of field names and values to update
        """
        logger.info(f"Tool called: update_list_item for site {site_url}")
        
        try:
            # Get authentication context from context object
            sp_ctx = ctx.request_context.lifespan_context
            
            # Refresh token if needed
            await refresh_token_if_needed(sp_ctx)
            
            # Create Graph client
            graph_client = GraphClient(sp_ctx)
            
            # Update the list item
            result = await graph_client.update_list_item(site_id, list_id, item_id, fields)
            
            logger.info(f"Successfully updated list item {item_id}")
            return json.dumps(result, indent=2)
            
        except Exception as e:
            logger.error(f"Error in update_list_item: {str(e)}")
            return f"Error updating list item: {str(e)}"
    
    @mcp.tool()
    async def create_advanced_document_library(
        ctx: Context,
        site_url: str,
        site_id: str,
        display_name: str,
        doc_type: str = "general"
    ) -> str:
        """Create a document library with advanced metadata settings.
        
        Args:
            site_url: The SharePoint site URL (e.g., https://example.sharepoint.com/sites/test)
            site_id: ID of the site
            display_name: Display name of the library
            doc_type: Type of documents (general, contracts, marketing, reports, projects)
        """
        logger.info(f"Tool called: create_advanced_document_library for site {site_url}")
        
        try:
            # Get authentication context from context object
            sp_ctx = ctx.request_context.lifespan_context
            
            # Refresh token if needed
            await refresh_token_if_needed(sp_ctx)
            
            # Create Graph client and content generator
            graph_client = GraphClient(sp_ctx)
            generator = ContentGenerator()
            
            # Generate library configuration based on document type
            library_config = generator.generate_document_library_schema(doc_type, display_name)
            
            # Create the document library
            result = await graph_client.create_document_library(site_id, library_config)
            
            logger.info(f"Successfully created document library: {display_name}")
            return json.dumps(result, indent=2)
            
        except Exception as e:
            logger.error(f"Error in create_advanced_document_library: {str(e)}")
            return f"Error creating document library: {str(e)}"
    
    @mcp.tool()
    async def upload_document(
        ctx: Context,
        site_url: str,
        site_id: str,
        drive_id: str,
        folder_path: str,
        file_name: str,
        file_content: bytes,
        content_type: str = "application/octet-stream"
    ) -> str:
        """Upload a document to a SharePoint document library.
        
        Args:
            site_url: The SharePoint site URL (e.g., https://example.sharepoint.com/sites/test)
            site_id: ID of the site
            drive_id: ID of the document library
            folder_path: Path to the folder (e.g. 'General' or 'Documents/Folder1')
            file_name: Name of the file to create
            file_content: Content of the file as bytes
            content_type: MIME type of the file
        """
        logger.info(f"Tool called: upload_document to {folder_path}/{file_name} for site {site_url}")
        
        try:
            # Get authentication context from context object
            sp_ctx = ctx.request_context.lifespan_context
            
            # Refresh token if needed
            await refresh_token_if_needed(sp_ctx)
            
            # Create Graph client
            graph_client = GraphClient(sp_ctx)
            
            # Upload the document
            result = await graph_client.upload_document(
                site_id, drive_id, folder_path, file_name, file_content, content_type
            )
            
            logger.info(f"Successfully uploaded document: {file_name}")
            return json.dumps(result, indent=2)
            
        except Exception as e:
            logger.error(f"Error in upload_document: {str(e)}")
            return f"Error uploading document: {str(e)}"
    
    @mcp.tool()
    async def create_modern_page(
        ctx: Context,
        site_url: str,
        site_id: str,
        name: str,
        purpose: str = "general",
        audience: str = "general"
    ) -> str:
        """Create a modern SharePoint page with beautiful layout.
        
        Args:
            site_url: The SharePoint site URL (e.g., https://example.sharepoint.com/sites/test)
            site_id: ID of the site
            name: Name of the page (for URL)
            purpose: Purpose of the page (welcome, dashboard, team, project, announcement)
            audience: Target audience (general, executives, team, customers)
        """
        logger.info(f"Tool called: create_modern_page for site {site_url}")
        
        try:
            # Get authentication context from context object
            sp_ctx = ctx.request_context.lifespan_context
            
            # Refresh token if needed
            await refresh_token_if_needed(sp_ctx)
            
            # Create Graph client and content generator
            graph_client = GraphClient(sp_ctx)
            generator = ContentGenerator()
            
            # Generate page content based on purpose and audience
            page_config = generator.generate_page_content(purpose, audience, name)
            
            # Create the page
            result = await graph_client.create_page(site_id, page_config)
            
            logger.info(f"Successfully created modern page: {name}")
            return json.dumps(result, indent=2)
            
        except Exception as e:
            logger.error(f"Error in create_modern_page: {str(e)}")
            return f"Error creating page: {str(e)}"
    
    @mcp.tool()
    async def create_news_post(
        ctx: Context,
        site_url: str,
        site_id: str,
        title: str,
        description: str = "",
        content: str = ""
    ) -> str:
        """Create a news post in a SharePoint site.
        
        Args:
            site_url: The SharePoint site URL (e.g., https://example.sharepoint.com/sites/test)
            site_id: ID of the site
            title: Title of the news post
            description: Brief description of the news post
            content: HTML or Markdown content of the news post
        """
        logger.info(f"Tool called: create_news_post for site {site_url}")
        
        try:
            # Get authentication context from context object
            sp_ctx = ctx.request_context.lifespan_context
            
            # Refresh token if needed
            await refresh_token_if_needed(sp_ctx)
            
            # Create Graph client and content generator
            graph_client = GraphClient(sp_ctx)
            generator = ContentGenerator()
            
            # Generate news post configuration
            news_config = generator.generate_news_post(title, description, content)
            
            # Create the news post
            result = await graph_client.create_news_post(site_id, news_config)
            
            logger.info(f"Successfully created news post: {title}")
            return json.dumps(result, indent=2)
            
        except Exception as e:
            logger.error(f"Error in create_news_post: {str(e)}")
            return f"Error creating news post: {str(e)}"
    
    @mcp.tool()
    async def get_document_content(
        ctx: Context,
        site_url: str,
        site_id: str,
        drive_id: str,
        item_id: str,
        filename: str
    ) -> str:
        """Get and process content from a SharePoint document.

        Args:
            site_url: The SharePoint site URL (e.g., https://example.sharepoint.com/sites/test)
            site_id: ID of the site
            drive_id: ID of the document library
            item_id: ID of the document
            filename: Name of the file (for content type detection)
        """
        logger.info(f"Tool called: get_document_content for {filename} in site {site_url}")

        try:
            # Get authentication context from context object
            sp_ctx = ctx.request_context.lifespan_context

            # Refresh token if needed
            await refresh_token_if_needed(sp_ctx)

            # Create Graph client and document processor
            graph_client = GraphClient(sp_ctx)
            processor = DocumentProcessor()

            # Download document content
            content = await graph_client.download_document(site_id, drive_id, item_id)

            # Process document based on file type
            processed_content = processor.process_document(content, filename)

            logger.info(f"Successfully retrieved and processed document: {filename}")
            return json.dumps(processed_content, indent=2)

        except Exception as e:
            logger.error(f"Error in get_document_content: {str(e)}")
            return f"Error getting document content: {str(e)}"

    @mcp.tool()
    async def list_folders(
        ctx: Context,
        site_url: str,
        site_id: str,
        drive_id: str,
        folder_path: str = ""
    ) -> str:
        """List all folders in a SharePoint document library location.

        Args:
            site_url: The SharePoint site URL (e.g., https://example.sharepoint.com/sites/test)
            site_id: ID of the site
            drive_id: ID of the document library
            folder_path: Path to the folder (empty string for root)
        """
        logger.info(f"Tool called: list_folders for {site_url} in {folder_path or 'root'}")

        try:
            # Get authentication context from context object
            sp_ctx = ctx.request_context.lifespan_context

            # Refresh token if needed
            await refresh_token_if_needed(sp_ctx)

            # Create Graph client
            graph_client = GraphClient(sp_ctx)

            # List folders
            result = await graph_client.list_drive_items(site_id, drive_id, folder_path, 'folder')

            # Extract and format folder information
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

            logger.info(f"Successfully retrieved {len(formatted_folders)} folders")
            return json.dumps(formatted_folders, indent=2)

        except Exception as e:
            logger.error(f"Error in list_folders: {str(e)}")
            return f"Error listing folders: {str(e)}"

    @mcp.tool()
    async def create_folder(
        ctx: Context,
        site_url: str,
        site_id: str,
        drive_id: str,
        folder_path: str
    ) -> str:
        """Create a new folder in a SharePoint document library.

        Args:
            site_url: The SharePoint site URL (e.g., https://example.sharepoint.com/sites/test)
            site_id: ID of the site
            drive_id: ID of the document library
            folder_path: Path of the folder to create (e.g., "Projects/2024" for nested folders)
        """
        logger.info(f"Tool called: create_folder for {site_url} at {folder_path}")

        try:
            # Get authentication context from context object
            sp_ctx = ctx.request_context.lifespan_context

            # Refresh token if needed
            await refresh_token_if_needed(sp_ctx)

            # Create Graph client
            graph_client = GraphClient(sp_ctx)

            # Create folder
            result = await graph_client.create_folder_in_library(site_id, drive_id, folder_path)

            logger.info(f"Successfully created folder: {folder_path}")
            return json.dumps(result, indent=2)

        except Exception as e:
            logger.error(f"Error in create_folder: {str(e)}")
            return f"Error creating folder: {str(e)}"

    @mcp.tool()
    async def delete_folder(
        ctx: Context,
        site_url: str,
        site_id: str,
        drive_id: str,
        folder_id: str
    ) -> str:
        """Delete a folder from a SharePoint document library.

        Args:
            site_url: The SharePoint site URL (e.g., https://example.sharepoint.com/sites/test)
            site_id: ID of the site
            drive_id: ID of the document library
            folder_id: ID of the folder to delete
        """
        logger.info(f"Tool called: delete_folder for {site_url}, folder_id: {folder_id}")

        try:
            # Get authentication context from context object
            sp_ctx = ctx.request_context.lifespan_context

            # Refresh token if needed
            await refresh_token_if_needed(sp_ctx)

            # Create Graph client
            graph_client = GraphClient(sp_ctx)

            # Delete folder
            result = await graph_client.delete_drive_item(site_id, drive_id, folder_id)

            logger.info(f"Successfully deleted folder: {folder_id}")
            return json.dumps(result, indent=2)

        except Exception as e:
            logger.error(f"Error in delete_folder: {str(e)}")
            return f"Error deleting folder: {str(e)}"

    @mcp.tool()
    async def get_folder_tree(
        ctx: Context,
        site_url: str,
        site_id: str,
        drive_id: str,
        folder_path: str = "",
        max_depth: int = 10
    ) -> str:
        """Get a recursive tree structure of folders in a SharePoint document library.

        Args:
            site_url: The SharePoint site URL (e.g., https://example.sharepoint.com/sites/test)
            site_id: ID of the site
            drive_id: ID of the document library
            folder_path: Starting folder path (empty string for root)
            max_depth: Maximum depth to traverse (default: 10)
        """
        logger.info(f"Tool called: get_folder_tree for {site_url} starting at {folder_path or 'root'}")

        try:
            # Get authentication context from context object
            sp_ctx = ctx.request_context.lifespan_context

            # Refresh token if needed
            await refresh_token_if_needed(sp_ctx)

            # Create Graph client
            graph_client = GraphClient(sp_ctx)

            # Get folder tree
            result = await graph_client.get_folder_tree(site_id, drive_id, folder_path, max_depth)

            logger.info(f"Successfully retrieved folder tree")
            return json.dumps(result, indent=2)

        except Exception as e:
            logger.error(f"Error in get_folder_tree: {str(e)}")
            return f"Error getting folder tree: {str(e)}"

    @mcp.tool()
    async def list_documents(
        ctx: Context,
        site_url: str,
        site_id: str,
        drive_id: str,
        folder_path: str = ""
    ) -> str:
        """List all documents in a SharePoint document library location.

        Args:
            site_url: The SharePoint site URL (e.g., https://example.sharepoint.com/sites/test)
            site_id: ID of the site
            drive_id: ID of the document library
            folder_path: Path to the folder (empty string for root)
        """
        logger.info(f"Tool called: list_documents for {site_url} in {folder_path or 'root'}")

        try:
            # Get authentication context from context object
            sp_ctx = ctx.request_context.lifespan_context

            # Refresh token if needed
            await refresh_token_if_needed(sp_ctx)

            # Create Graph client
            graph_client = GraphClient(sp_ctx)

            # List documents
            result = await graph_client.list_drive_items(site_id, drive_id, folder_path, 'file')

            # Extract and format document information
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

            logger.info(f"Successfully retrieved {len(formatted_documents)} documents")
            return json.dumps(formatted_documents, indent=2)

        except Exception as e:
            logger.error(f"Error in list_documents: {str(e)}")
            return f"Error listing documents: {str(e)}"

    @mcp.tool()
    async def update_document(
        ctx: Context,
        site_url: str,
        site_id: str,
        drive_id: str,
        item_id: str,
        file_content: bytes,
        content_type: str = "application/octet-stream"
    ) -> str:
        """Update the content of an existing document in SharePoint.

        Args:
            site_url: The SharePoint site URL (e.g., https://example.sharepoint.com/sites/test)
            site_id: ID of the site
            drive_id: ID of the document library
            item_id: ID of the document to update
            file_content: New content of the file as bytes
            content_type: MIME type of the file
        """
        logger.info(f"Tool called: update_document for {site_url}, item_id: {item_id}")

        try:
            # Get authentication context from context object
            sp_ctx = ctx.request_context.lifespan_context

            # Refresh token if needed
            await refresh_token_if_needed(sp_ctx)

            # Create Graph client
            graph_client = GraphClient(sp_ctx)

            # Update document
            result = await graph_client.update_document_content(
                site_id, drive_id, item_id, file_content, content_type
            )

            logger.info(f"Successfully updated document: {item_id}")
            return json.dumps(result, indent=2)

        except Exception as e:
            logger.error(f"Error in update_document: {str(e)}")
            return f"Error updating document: {str(e)}"

    @mcp.tool()
    async def delete_document(
        ctx: Context,
        site_url: str,
        site_id: str,
        drive_id: str,
        item_id: str
    ) -> str:
        """Delete a document from a SharePoint document library.

        Args:
            site_url: The SharePoint site URL (e.g., https://example.sharepoint.com/sites/test)
            site_id: ID of the site
            drive_id: ID of the document library
            item_id: ID of the document to delete
        """
        logger.info(f"Tool called: delete_document for {site_url}, item_id: {item_id}")

        try:
            # Get authentication context from context object
            sp_ctx = ctx.request_context.lifespan_context

            # Refresh token if needed
            await refresh_token_if_needed(sp_ctx)

            # Create Graph client
            graph_client = GraphClient(sp_ctx)

            # Delete document
            result = await graph_client.delete_drive_item(site_id, drive_id, item_id)

            logger.info(f"Successfully deleted document: {item_id}")
            return json.dumps(result, indent=2)

        except Exception as e:
            logger.error(f"Error in delete_document: {str(e)}")
            return f"Error deleting document: {str(e)}"