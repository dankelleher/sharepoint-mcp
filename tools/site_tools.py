"""SharePoint MCP tools - Controller layer.

This module provides the MCP tool interfaces that delegate to the service layer.
"""

import json
import logging
from mcp.server.fastmcp import FastMCP, Context
from auth.sharepoint_auth import refresh_token_if_needed
from services.sharepoint_service import SharePointService

logger = logging.getLogger("sharepoint_tools")


def register_site_tools(mcp: FastMCP):
    """Register SharePoint site tools with the MCP server."""

    # Site Operations

    @mcp.tool()
    async def get_site_info(ctx: Context, site_url: str) -> str:
        """Get basic information about the SharePoint site."""
        logger.info(f"Tool called: get_site_info for {site_url}")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            await refresh_token_if_needed(sp_ctx)
            service = SharePointService(sp_ctx)
            result = await service.get_site_info(site_url)
            return json.dumps(result, indent=2)
        except Exception as e:
            logger.error(f"Error in get_site_info: {str(e)}")
            return f"Error accessing SharePoint: {str(e)}"

    @mcp.tool()
    async def list_document_libraries(ctx: Context, site_url: str) -> str:
        """List all document libraries in the SharePoint site."""
        logger.info(f"Tool called: list_document_libraries for {site_url}")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            await refresh_token_if_needed(sp_ctx)
            service = SharePointService(sp_ctx)
            result = await service.list_document_libraries(site_url)
            return json.dumps(result, indent=2)
        except Exception as e:
            logger.error(f"Error in list_document_libraries: {str(e)}")
            return f"Error listing document libraries: {str(e)}"

    @mcp.tool()
    async def search_sharepoint(ctx: Context, site_url: str, query: str) -> str:
        """Search content in the SharePoint site."""
        logger.info(f"Tool called: search_sharepoint for {site_url} with query: {query}")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            await refresh_token_if_needed(sp_ctx)
            service = SharePointService(sp_ctx)
            results = await service.search_sharepoint(site_url, query)
            return json.dumps(results, indent=2)
        except Exception as e:
            logger.error(f"Error in search_sharepoint: {str(e)}")
            return f"Error searching SharePoint: {str(e)}"

    @mcp.tool()
    async def create_sharepoint_site(ctx: Context, display_name: str, alias: str, description: str = "") -> str:
        """Create a new SharePoint site."""
        logger.info(f"Tool called: create_sharepoint_site with name: {display_name}")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            await refresh_token_if_needed(sp_ctx)
            service = SharePointService(sp_ctx)
            result = await service.create_sharepoint_site(display_name, alias, description)
            return json.dumps(result, indent=2)
        except Exception as e:
            logger.error(f"Error in create_sharepoint_site: {str(e)}")
            return f"Error creating SharePoint site: {str(e)}"

    # Folder Operations

    @mcp.tool()
    async def list_folders(ctx: Context, site_url: str, site_id: str, drive_id: str, folder_path: str = "") -> str:
        """List all folders in a SharePoint document library location."""
        logger.info(f"Tool called: list_folders for {site_url} in {folder_path or 'root'}")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            await refresh_token_if_needed(sp_ctx)
            service = SharePointService(sp_ctx)
            result = await service.list_folders(site_id, drive_id, folder_path)
            return json.dumps(result, indent=2)
        except Exception as e:
            logger.error(f"Error in list_folders: {str(e)}")
            return f"Error listing folders: {str(e)}"

    @mcp.tool()
    async def create_folder(ctx: Context, site_url: str, site_id: str, drive_id: str, folder_path: str) -> str:
        """Create a new folder in a SharePoint document library."""
        logger.info(f"Tool called: create_folder for {site_url} at {folder_path}")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            await refresh_token_if_needed(sp_ctx)
            service = SharePointService(sp_ctx)
            result = await service.create_folder(site_id, drive_id, folder_path)
            return json.dumps(result, indent=2)
        except Exception as e:
            logger.error(f"Error in create_folder: {str(e)}")
            return f"Error creating folder: {str(e)}"

    @mcp.tool()
    async def delete_folder(ctx: Context, site_url: str, site_id: str, drive_id: str, folder_id: str) -> str:
        """Delete a folder from a SharePoint document library."""
        logger.info(f"Tool called: delete_folder for {site_url}, folder_id: {folder_id}")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            await refresh_token_if_needed(sp_ctx)
            service = SharePointService(sp_ctx)
            result = await service.delete_folder(site_id, drive_id, folder_id)
            return json.dumps(result, indent=2)
        except Exception as e:
            logger.error(f"Error in delete_folder: {str(e)}")
            return f"Error deleting folder: {str(e)}"

    @mcp.tool()
    async def get_folder_tree(ctx: Context, site_url: str, site_id: str, drive_id: str,
                             folder_path: str = "", max_depth: int = 10) -> str:
        """Get a recursive tree structure of folders in a SharePoint document library."""
        logger.info(f"Tool called: get_folder_tree for {site_url} starting at {folder_path or 'root'}")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            await refresh_token_if_needed(sp_ctx)
            service = SharePointService(sp_ctx)
            result = await service.get_folder_tree(site_id, drive_id, folder_path, max_depth)
            return json.dumps(result, indent=2)
        except Exception as e:
            logger.error(f"Error in get_folder_tree: {str(e)}")
            return f"Error getting folder tree: {str(e)}"

    # Document Operations

    @mcp.tool()
    async def list_documents(ctx: Context, site_url: str, site_id: str, drive_id: str, folder_path: str = "") -> str:
        """List all documents in a SharePoint document library location."""
        logger.info(f"Tool called: list_documents for {site_url} in {folder_path or 'root'}")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            await refresh_token_if_needed(sp_ctx)
            service = SharePointService(sp_ctx)
            result = await service.list_documents(site_id, drive_id, folder_path)
            return json.dumps(result, indent=2)
        except Exception as e:
            logger.error(f"Error in list_documents: {str(e)}")
            return f"Error listing documents: {str(e)}"

    @mcp.tool()
    async def get_document_content(ctx: Context, site_url: str, site_id: str, drive_id: str,
                                  item_id: str, filename: str) -> str:
        """Get and process content from a SharePoint document."""
        logger.info(f"Tool called: get_document_content for {filename} in site {site_url}")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            await refresh_token_if_needed(sp_ctx)
            service = SharePointService(sp_ctx)
            processed_content = await service.get_document_content(site_id, drive_id, item_id, filename)
            return json.dumps(processed_content, indent=2)
        except Exception as e:
            logger.error(f"Error in get_document_content: {str(e)}")
            return f"Error getting document content: {str(e)}"

    @mcp.tool()
    async def upload_document(ctx: Context, site_url: str, site_id: str, drive_id: str,
                            folder_path: str, file_name: str, file_content: bytes,
                            content_type: str = "application/octet-stream") -> str:
        """Upload a new document to SharePoint."""
        logger.info(f"Tool called: upload_document for {file_name} to {site_url}/{folder_path}")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            await refresh_token_if_needed(sp_ctx)
            service = SharePointService(sp_ctx)
            result = await service.upload_document(site_id, drive_id, folder_path, file_name, file_content, content_type)
            return json.dumps(result, indent=2)
        except Exception as e:
            logger.error(f"Error in upload_document: {str(e)}")
            return f"Error uploading document: {str(e)}"

    @mcp.tool()
    async def update_document(ctx: Context, site_url: str, site_id: str, drive_id: str,
                            item_id: str, file_content: bytes, content_type: str = "application/octet-stream") -> str:
        """Update the content of an existing document in SharePoint."""
        logger.info(f"Tool called: update_document for {site_url}, item_id: {item_id}")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            await refresh_token_if_needed(sp_ctx)
            service = SharePointService(sp_ctx)
            result = await service.update_document(site_id, drive_id, item_id, file_content, content_type)
            return json.dumps(result, indent=2)
        except Exception as e:
            logger.error(f"Error in update_document: {str(e)}")
            return f"Error updating document: {str(e)}"

    @mcp.tool()
    async def delete_document(ctx: Context, site_url: str, site_id: str, drive_id: str, item_id: str) -> str:
        """Delete a document from a SharePoint document library."""
        logger.info(f"Tool called: delete_document for {site_url}, item_id: {item_id}")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            await refresh_token_if_needed(sp_ctx)
            service = SharePointService(sp_ctx)
            result = await service.delete_document(site_id, drive_id, item_id)
            return json.dumps(result, indent=2)
        except Exception as e:
            logger.error(f"Error in delete_document: {str(e)}")
            return f"Error deleting document: {str(e)}"

    # List Operations

    @mcp.tool()
    async def create_intelligent_list(ctx: Context, site_url: str, site_id: str, purpose: str, display_name: str) -> str:
        """Create a SharePoint list with AI-optimized schema based on its purpose."""
        logger.info(f"Tool called: create_intelligent_list for site {site_url}")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            await refresh_token_if_needed(sp_ctx)
            service = SharePointService(sp_ctx)
            result = await service.create_intelligent_list(site_id, purpose, display_name)
            return json.dumps(result, indent=2)
        except Exception as e:
            logger.error(f"Error in create_intelligent_list: {str(e)}")
            return f"Error creating list: {str(e)}"

    @mcp.tool()
    async def create_list_item(ctx: Context, site_url: str, site_id: str, list_id: str, fields: dict) -> str:
        """Add new items to SharePoint lists."""
        logger.info(f"Tool called: create_list_item for list {list_id} in site {site_url}")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            await refresh_token_if_needed(sp_ctx)
            service = SharePointService(sp_ctx)
            result = await service.create_list_item(site_id, list_id, fields)
            return json.dumps(result, indent=2)
        except Exception as e:
            logger.error(f"Error in create_list_item: {str(e)}")
            return f"Error creating list item: {str(e)}"

    @mcp.tool()
    async def update_list_item(ctx: Context, site_url: str, site_id: str, list_id: str, item_id: str, fields: dict) -> str:
        """Update existing list items."""
        logger.info(f"Tool called: update_list_item for item {item_id} in list {list_id}")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            await refresh_token_if_needed(sp_ctx)
            service = SharePointService(sp_ctx)
            result = await service.update_list_item(site_id, list_id, item_id, fields)
            return json.dumps(result, indent=2)
        except Exception as e:
            logger.error(f"Error in update_list_item: {str(e)}")
            return f"Error updating list item: {str(e)}"

    # Content Creation

    @mcp.tool()
    async def create_advanced_document_library(ctx: Context, site_url: str, site_id: str,
                                              display_name: str, doc_type: str = "general") -> str:
        """Create a document library with advanced metadata settings."""
        logger.info(f"Tool called: create_advanced_document_library for site {site_url}")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            await refresh_token_if_needed(sp_ctx)
            service = SharePointService(sp_ctx)
            result = await service.create_advanced_document_library(site_id, display_name, doc_type)
            return json.dumps(result, indent=2)
        except Exception as e:
            logger.error(f"Error in create_advanced_document_library: {str(e)}")
            return f"Error creating document library: {str(e)}"

    @mcp.tool()
    async def create_modern_page(ctx: Context, site_url: str, site_id: str, name: str, title: str = "",
                                purpose: str = "general", audience: str = "general") -> str:
        """Create a modern SharePoint page with beautiful layout."""
        logger.info(f"Tool called: create_modern_page for site {site_url}")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            await refresh_token_if_needed(sp_ctx)
            service = SharePointService(sp_ctx)
            result = await service.create_modern_page(site_id, name, title, purpose, audience)
            return json.dumps(result, indent=2)
        except Exception as e:
            logger.error(f"Error in create_modern_page: {str(e)}")
            return f"Error creating page: {str(e)}"

    @mcp.tool()
    async def create_news_post(ctx: Context, site_url: str, site_id: str, title: str,
                              description: str = "", content: str = "") -> str:
        """Create a news post in a SharePoint site."""
        logger.info(f"Tool called: create_news_post for site {site_url}")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            await refresh_token_if_needed(sp_ctx)
            service = SharePointService(sp_ctx)
            result = await service.create_news_post(site_id, title, description, content)
            return json.dumps(result, indent=2)
        except Exception as e:
            logger.error(f"Error in create_news_post: {str(e)}")
            return f"Error creating news post: {str(e)}"
