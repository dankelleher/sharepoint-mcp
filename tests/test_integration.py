"""Comprehensive integration tests for SharePoint MCP Server.

This test suite validates all tools against a live SharePoint environment.
Tests are designed to be run against the civicteamaccount SharePoint instance
on the temp-todo-remove site.

IMPORTANT: These tests will create, modify, and delete content in SharePoint.
Only run against a test site that you can safely modify.
"""

import sys
import os
from pathlib import Path

# Add project root to Python path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import asyncio
import json
import pytest
from datetime import datetime
from auth.sharepoint_auth import get_auth_context
from services.sharepoint_service import SharePointService

# Test configuration
DOMAIN = "civicteamaccount.sharepoint.com"
SITE_NAME = "temp-todo-remove"
SITE_URL = f"https://{DOMAIN}/sites/{SITE_NAME}"

# Test data will be stored here during test run
test_data = {
    "site_id": None,
    "drive_id": None,
    "folder_id": None,
    "document_id": None,
    "list_id": None,
    "list_item_id": None,
    "page_id": None,
    "news_post_id": None,
    "advanced_library_id": None
}


class TestSharePointConnection:
    """Test basic SharePoint connection and authentication."""

    @pytest.mark.asyncio
    async def test_authentication(self):
        """Test that we can authenticate with SharePoint."""
        context = await get_auth_context()
        assert context is not None
        assert context.access_token is not None
        assert len(context.access_token) > 0
        print(f"✓ Authentication successful")

    @pytest.mark.asyncio
    async def test_get_site_info(self):
        """Test retrieving site information."""
        context = await get_auth_context()
        service = SharePointService(context)

        site_info = await service.get_site_info(SITE_URL)

        assert site_info is not None
        assert "id" in site_info
        assert "name" in site_info

        # Store site_id for later tests
        test_data["site_id"] = site_info["id"]
        print(f"✓ Site info retrieved: {site_info['name']}")
        print(f"  Site ID: {test_data['site_id']}")

    @pytest.mark.asyncio
    async def test_create_sharepoint_site(self):
        """Test creating a new SharePoint site.

        NOTE: This test is skipped by default because creating sites is a major
        operation that requires Sites.Manage.All permission and creates permanent
        resources. To enable this test, set environment variable:
        TEST_SITE_CREATION=true
        """
        import os
        if not os.getenv("TEST_SITE_CREATION"):
            pytest.skip("Site creation test disabled. Set TEST_SITE_CREATION=true to enable")

        context = await get_auth_context()
        service = SharePointService(context)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        display_name = f"Test Site {timestamp}"
        alias = f"test-site-{timestamp}"

        result = await service.create_sharepoint_site(
            display_name,
            alias,
            "Test site created by integration tests"
        )

        assert result is not None
        assert "id" in result or "webUrl" in result

        print(f"✓ Created SharePoint site: {display_name}")
        print(f"  Alias: {alias}")
        if "webUrl" in result:
            print(f"  URL: {result['webUrl']}")
        print(f"  NOTE: This site must be manually deleted from SharePoint admin center")


class TestDocumentLibraries:
    """Test document library operations."""

    @pytest.mark.asyncio
    async def test_list_document_libraries(self):
        """Test listing document libraries."""
        context = await get_auth_context()
        service = SharePointService(context)

        result = await service.list_document_libraries(SITE_URL)

        assert result is not None
        assert "libraries" in result
        assert len(result["libraries"]) > 0

        # Store the first drive_id for later tests
        test_data["drive_id"] = result["libraries"][0]["id"]
        print(f"✓ Found {result['count']} document libraries")
        print(f"  Using drive ID: {test_data['drive_id']}")


class TestFolderOperations:
    """Test folder management operations."""

    @pytest.mark.asyncio
    async def test_create_folder(self):
        """Test creating a folder."""
        context = await get_auth_context()
        service = SharePointService(context)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        folder_path = f"test_folder_{timestamp}"

        result = await service.create_folder(
            test_data["site_id"],
            test_data["drive_id"],
            folder_path
        )

        assert result is not None
        assert "id" in result

        test_data["folder_id"] = result["id"]
        test_data["folder_path"] = folder_path
        print(f"✓ Created folder: {folder_path}")
        print(f"  Folder ID: {test_data['folder_id']}")

    @pytest.mark.asyncio
    async def test_list_folders(self):
        """Test listing folders."""
        context = await get_auth_context()
        service = SharePointService(context)

        result = await service.list_folders(
            test_data["site_id"],
            test_data["drive_id"],
            ""
        )

        assert result is not None
        assert "folders" in result
        print(f"✓ Listed {result['count']} folders")

    @pytest.mark.asyncio
    async def test_get_folder_tree(self):
        """Test getting folder tree structure."""
        context = await get_auth_context()
        service = SharePointService(context)

        result = await service.get_folder_tree(
            test_data["site_id"],
            test_data["drive_id"],
            "",
            5  # Max depth
        )

        assert result is not None
        assert "folders" in result
        print(f"✓ Retrieved folder tree with {len(result['folders'])} root folders")


class TestDocumentOperations:
    """Test document management operations."""

    @pytest.mark.asyncio
    async def test_upload_document(self):
        """Test uploading a document."""
        context = await get_auth_context()
        service = SharePointService(context)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_name = f"test_document_{timestamp}.txt"
        file_content = f"Test document created at {datetime.now().isoformat()}".encode('utf-8')

        result = await service.upload_document(
            test_data["site_id"],
            test_data["drive_id"],
            test_data.get("folder_path", ""),
            file_name,
            file_content,
            "text/plain"
        )

        assert result is not None
        assert "id" in result

        test_data["document_id"] = result["id"]
        print(f"✓ Uploaded document: {file_name}")
        print(f"  Document ID: {test_data['document_id']}")

    @pytest.mark.asyncio
    async def test_list_documents(self):
        """Test listing documents."""
        context = await get_auth_context()
        service = SharePointService(context)

        result = await service.list_documents(
            test_data["site_id"],
            test_data["drive_id"],
            test_data.get("folder_path", "")
        )

        assert result is not None
        assert "documents" in result
        print(f"✓ Listed {result['count']} documents")

    @pytest.mark.asyncio
    async def test_get_document_content(self):
        """Test retrieving document content."""
        context = await get_auth_context()
        service = SharePointService(context)

        # get_document_content needs filename parameter
        content = await service.get_document_content(
            test_data["site_id"],
            test_data["drive_id"],
            test_data["document_id"],
            "test_document.txt"
        )

        assert content is not None
        assert "type" in content or "error" not in content
        print(f"✓ Retrieved document content")

    @pytest.mark.asyncio
    async def test_update_document(self):
        """Test updating document content."""
        context = await get_auth_context()
        service = SharePointService(context)

        new_content = f"Updated content at {datetime.now().isoformat()}".encode('utf-8')

        result = await service.update_document(
            test_data["site_id"],
            test_data["drive_id"],
            test_data["document_id"],
            new_content,
            "text/plain"
        )

        assert result is not None
        print(f"✓ Updated document content")


class TestListOperations:
    """Test SharePoint list operations."""

    @pytest.mark.asyncio
    async def test_create_intelligent_list(self):
        """Test creating a list with intelligent schema."""
        context = await get_auth_context()
        service = SharePointService(context)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        list_name = f"Test Projects {timestamp}"

        result = await service.create_intelligent_list(
            test_data["site_id"],
            "projects",
            list_name
        )

        assert result is not None
        assert "id" in result

        test_data["list_id"] = result["id"]
        print(f"✓ Created intelligent list: {list_name}")
        print(f"  List ID: {test_data['list_id']}")

    @pytest.mark.asyncio
    async def test_create_list_item(self):
        """Test creating a list item."""
        context = await get_auth_context()
        service = SharePointService(context)

        fields = {
            "Title": "Test Project",
            "ProjectName": "Integration Test Project",
            "Status": "In Progress",
            "Priority": "High"
        }

        result = await service.create_list_item(
            test_data["site_id"],
            test_data["list_id"],
            fields
        )

        assert result is not None
        assert "id" in result

        test_data["list_item_id"] = result["id"]
        print(f"✓ Created list item")
        print(f"  Item ID: {test_data['list_item_id']}")

    @pytest.mark.asyncio
    async def test_update_list_item(self):
        """Test updating a list item."""
        context = await get_auth_context()
        service = SharePointService(context)

        fields = {
            "Status": "Completed",
            "PercentComplete": 100
        }

        result = await service.update_list_item(
            test_data["site_id"],
            test_data["list_id"],
            test_data["list_item_id"],
            fields
        )

        assert result is not None
        print(f"✓ Updated list item")


class TestSearchOperations:
    """Test SharePoint search functionality."""

    @pytest.mark.asyncio
    async def test_search_sharepoint(self):
        """Test searching SharePoint content."""
        context = await get_auth_context()
        service = SharePointService(context)

        result = await service.search_sharepoint(SITE_URL, "test")

        assert result is not None
        assert "requests" in result or "value" in result
        print(f"✓ Search completed")


class TestContentCreation:
    """Test content creation features."""

    @pytest.mark.asyncio
    async def test_create_advanced_document_library(self):
        """Test creating an advanced document library."""
        context = await get_auth_context()
        service = SharePointService(context)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        library_name = f"Test Advanced Library {timestamp}"

        result = await service.create_advanced_document_library(
            test_data["site_id"],
            library_name,
            "contracts"  # Use contracts type for advanced metadata
        )

        assert result is not None
        assert "id" in result

        test_data["advanced_library_id"] = result["id"]
        print(f"✓ Created advanced document library: {library_name}")
        print(f"  Library ID: {test_data['advanced_library_id']}")

    @pytest.mark.asyncio
    async def test_create_page(self):
        """Test creating a SharePoint page."""
        context = await get_auth_context()
        service = SharePointService(context)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        page_name = f"test-page-{timestamp}"

        result = await service.create_modern_page(
            test_data["site_id"],
            page_name,
            "Test Page"
        )

        assert result is not None
        assert "id" in result

        test_data["page_id"] = result["id"]
        print(f"✓ Created page: {page_name}")

    @pytest.mark.asyncio
    async def test_create_news_post(self):
        """Test creating a news post."""
        context = await get_auth_context()
        service = SharePointService(context)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        news_title = f"Test News Post {timestamp}"

        result = await service.create_news_post(
            test_data["site_id"],
            news_title,
            "This is a test news post created by integration tests",
            "Test content for the news post"
        )

        assert result is not None
        assert "id" in result

        test_data["news_post_id"] = result["id"]
        print(f"✓ Created news post: {news_title}")
        print(f"  News Post ID: {test_data['news_post_id']}")


class TestCleanup:
    """Cleanup test data from SharePoint."""

    @pytest.mark.asyncio
    async def test_delete_document(self):
        """Test deleting a document."""
        if not test_data.get("document_id"):
            pytest.skip("No document to delete")

        context = await get_auth_context()
        service = SharePointService(context)

        result = await service.delete_document(
            test_data["site_id"],
            test_data["drive_id"],
            test_data["document_id"]
        )

        assert result is not None
        print(f"✓ Deleted document")

    @pytest.mark.asyncio
    async def test_delete_folder(self):
        """Test deleting a folder."""
        if not test_data.get("folder_id"):
            pytest.skip("No folder to delete")

        context = await get_auth_context()
        service = SharePointService(context)

        result = await service.delete_folder(
            test_data["site_id"],
            test_data["drive_id"],
            test_data["folder_id"]
        )

        assert result is not None
        print(f"✓ Deleted folder")


# Test execution order
pytest_plugins = ['pytest_asyncio']


if __name__ == "__main__":
    """Run tests directly with proper async support."""
    pytest.main([
        __file__,
        "-v",
        "-s",
        "--tb=short",
        "--asyncio-mode=auto"
    ])
