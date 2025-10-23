# SharePoint MCP Server - Integration Testing Guide

This document describes how to run the comprehensive integration test suite for the SharePoint MCP Server.

## Overview

The integration test suite (`tests/test_integration.py`) validates all MCP tools against a live SharePoint environment. Tests are configured to run against:

- **SharePoint Account**: `civicteamaccount.sharepoint.com`
- **Site**: `temp-todo-remove`
- **Site URL**: `https://civicteamaccount.sharepoint.com/sites/temp-todo-remove`

⚠️ **IMPORTANT**: These tests will create, modify, and delete content in SharePoint. Only run against a test site that you can safely modify.

## Prerequisites

1. **Python Environment**
   - Python 3.10 or higher
   - Virtual environment activated
   - All dependencies installed

2. **Authentication**
   - Valid `.env` file with credentials:
     ```
     TENANT_ID=your-tenant-id
     CLIENT_ID=your-client-id
     CLIENT_SECRET=your-client-secret
     ```

3. **Permissions**
   - Azure AD app must have the following Application permissions:
     - `Sites.ReadWrite.All`
     - `Sites.Manage.All`
     - `Files.ReadWrite.All`

4. **Test Dependencies**
   ```bash
   pip install -r requirements-dev.txt
   ```

## Running the Tests

### Quick Start

From the project root directory:

```bash
# Activate virtual environment
source venv/bin/activate

# Run all integration tests
pytest tests/test_integration.py -v
```

### Detailed Output

To see detailed output including print statements:

```bash
pytest tests/test_integration.py -v -s
```

### Run Specific Test Classes

```bash
# Test only folder operations
pytest tests/test_integration.py::TestFolderOperations -v

# Test only document operations
pytest tests/test_integration.py::TestDocumentOperations -v

# Test only list operations
pytest tests/test_integration.py::TestListOperations -v
```

### Run Specific Tests

```bash
# Run a single test
pytest tests/test_integration.py::TestFolderOperations::test_create_folder -v
```

### Run Tests with Coverage

```bash
# Install coverage if needed
pip install pytest-cov

# Run with coverage report
pytest tests/test_integration.py --cov=. --cov-report=html
```
