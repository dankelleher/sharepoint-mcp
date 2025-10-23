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
from utils.graph_client import GraphClient
from tools.site_tools import parse_site_url

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
    "list_item_id": None
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
        client = GraphClient(context)

        site_info = await client.get_site_info(DOMAIN, SITE_NAME)

        assert site_info is not None
        assert "id" in site_info
        assert "displayName" in site_info

        # Store site_id for later tests
        test_data["site_id"] = site_info["id"]
        print(f"✓ Site info retrieved: {site_info['displayName']}")
        print(f"  Site ID: {test_data['site_id']}")


class TestDocumentLibraries:
    """Test document library operations."""

    @pytest.mark.asyncio
    async def test_list_document_libraries(self):
        """Test listing document libraries."""
        context = await get_auth_context()
        client = GraphClient(context)

        result = await client.list_document_libraries(DOMAIN, SITE_NAME)

        assert result is not None
        assert "value" in result
        assert len(result["value"]) > 0

        # Store the first drive_id for later tests
        test_data["drive_id"] = result["value"][0]["id"]
        print(f"✓ Found {len(result['value'])} document libraries")
        print(f"  Using drive ID: {test_data['drive_id']}")


class TestFolderOperations:
    """Test folder management operations."""

    @pytest.mark.asyncio
    async def test_create_folder(self):
        """Test creating a folder."""
        context = await get_auth_context()
        client = GraphClient(context)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        folder_path = f"test_folder_{timestamp}"

        result = await client.create_folder_in_library(
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
        client = GraphClient(context)

        result = await client.list_drive_items(
            test_data["site_id"],
            test_data["drive_id"],
            "",
            "folder"
        )

        assert result is not None
        assert "value" in result
        print(f"✓ Listed {len(result['value'])} folders")

    @pytest.mark.asyncio
    async def test_get_folder_tree(self):
        """Test getting folder tree structure."""
        context = await get_auth_context()
        client = GraphClient(context)

        result = await client.get_folder_tree(
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
        client = GraphClient(context)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_name = f"test_document_{timestamp}.txt"
        file_content = f"Test document created at {datetime.now().isoformat()}".encode('utf-8')

        result = await client.upload_document(
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
        client = GraphClient(context)

        result = await client.list_drive_items(
            test_data["site_id"],
            test_data["drive_id"],
            test_data.get("folder_path", ""),
            "file"
        )

        assert result is not None
        assert "value" in result
        print(f"✓ Listed {len(result['value'])} documents")

    @pytest.mark.asyncio
    async def test_get_document_content(self):
        """Test retrieving document content."""
        context = await get_auth_context()
        client = GraphClient(context)

        content = await client.get_document_content(
            test_data["site_id"],
            test_data["drive_id"],
            test_data["document_id"]
        )

        assert content is not None
        assert len(content) > 0
        print(f"✓ Retrieved document content ({len(content)} bytes)")

    @pytest.mark.asyncio
    async def test_update_document(self):
        """Test updating document content."""
        context = await get_auth_context()
        client = GraphClient(context)

        new_content = f"Updated content at {datetime.now().isoformat()}".encode('utf-8')

        result = await client.update_document_content(
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
        client = GraphClient(context)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        list_name = f"Test Projects {timestamp}"

        result = await client.create_intelligent_list(
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
        client = GraphClient(context)

        fields = {
            "Title": "Test Project",
            "ProjectName": "Integration Test Project",
            "Status": "In Progress",
            "Priority": "High"
        }

        result = await client.create_list_item(
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
        client = GraphClient(context)

        fields = {
            "Status": "Completed",
            "PercentComplete": 100
        }

        result = await client.update_list_item(
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
        client = GraphClient(context)

        result = await client.search_sharepoint(DOMAIN, SITE_NAME, "test")

        assert result is not None
        assert "requests" in result or "value" in result
        print(f"✓ Search completed")


class TestContentCreation:
    """Test content creation features."""

    @pytest.mark.asyncio
    async def test_create_page(self):
        """Test creating a SharePoint page."""
        context = await get_auth_context()
        client = GraphClient(context)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        page_name = f"test-page-{timestamp}"

        result = await client.create_page(
            test_data["site_id"],
            page_name,
            "Test Page"
        )

        assert result is not None
        assert "id" in result

        test_data["page_id"] = result["id"]
        print(f"✓ Created page: {page_name}")


class TestCleanup:
    """Cleanup test data from SharePoint."""

    @pytest.mark.asyncio
    async def test_delete_document(self):
        """Test deleting a document."""
        if not test_data.get("document_id"):
            pytest.skip("No document to delete")

        context = await get_auth_context()
        client = GraphClient(context)

        result = await client.delete_drive_item(
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
        client = GraphClient(context)

        result = await client.delete_drive_item(
            test_data["site_id"],
            test_data["drive_id"],
            test_data["folder_id"]
        )

        assert result is not None
        print(f"✓ Deleted folder")

    @pytest.mark.asyncio
    async def test_delete_list_item(self):
        """Test deleting a list item."""
        if not test_data.get("list_item_id"):
            pytest.skip("No list item to delete")

        context = await get_auth_context()
        client = GraphClient(context)

        result = await client.delete_list_item(
            test_data["site_id"],
            test_data["list_id"],
            test_data["list_item_id"]
        )

        assert result is not None
        print(f"✓ Deleted list item")


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
