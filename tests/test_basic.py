"""Basic tests for MCP Excel Server."""

import pytest
from pathlib import Path


@pytest.mark.asyncio
async def test_server_initialization():
    """Test that the server initializes correctly."""
    from mcp_excel.server import mcp
    
    assert mcp.name == "mcp-excel-server"
    assert mcp.version == "0.2.0"


@pytest.mark.asyncio
async def test_read_cell(sample_excel_file):
    """Test reading a single cell."""
    from mcp_excel.tools.read import read_cell
    
    result = await read_cell(
        file_path=str(sample_excel_file),
        sheet_name="TestSheet",
        cell="A1",
    )
    
    assert result["success"] is True
    assert result["value"] == "Name"
    assert result["type"] == "string"


@pytest.mark.asyncio
async def test_read_range(sample_excel_file):
    """Test reading a range of cells."""
    from mcp_excel.tools.read import read_range
    
    result = await read_range(
        file_path=str(sample_excel_file),
        sheet_name="TestSheet",
        cell_range="A1:C4",
    )
    
    assert result["success"] is True
    assert len(result["data"]) == 4
    assert result["data"][0] == ["Name", "Age", "City"]


@pytest.mark.asyncio
async def test_list_sheets(sample_excel_file):
    """Test listing sheets."""
    from mcp_excel.tools.read import list_sheets
    
    result = await list_sheets(
        file_path=str(sample_excel_file),
    )
    
    assert result["success"] is True
    assert "TestSheet" in result["sheets"]
    assert result["count"] == 1


@pytest.mark.asyncio
async def test_describe_workbook(sample_excel_file):
    """Test describing a workbook."""
    from mcp_excel.tools.read import describe_workbook
    
    result = await describe_workbook(
        file_path=str(sample_excel_file),
    )
    
    assert result["success"] is True
    assert result["total_sheets"] == 1
    assert result["file"]["name"] == "test.xlsx"


@pytest.mark.asyncio
async def test_get_sheet_info(sample_excel_file):
    """Test getting sheet info."""
    from mcp_excel.tools.read import get_sheet_info
    
    result = await get_sheet_info(
        file_path=str(sample_excel_file),
        sheet_name="TestSheet",
    )
    
    assert result["success"] is True
    assert result["name"] == "TestSheet"
    assert result["row_count"] == 4
    assert len(result["columns"]) == 4


@pytest.mark.asyncio
async def test_search_cells(sample_excel_file):
    """Test searching cells."""
    from mcp_excel.tools.read import search_cells
    
    result = await search_cells(
        file_path=str(sample_excel_file),
        query="Alice",
    )
    
    assert result["success"] is True
    assert result["total_matches"] >= 1
    assert any(m["value"] == "Alice" for m in result["matches"])


@pytest.mark.asyncio
async def test_write_cells(sample_excel_file):
    """Test writing cells."""
    from mcp_excel.tools.write import write_cells
    
    result = await write_cells(
        file_path=str(sample_excel_file),
        sheet_name="TestSheet",
        cell_range="A5:C5",
        values=[["Dave", 40, "Berlin"]],
    )
    
    assert result["success"] is True
    assert result["cells_written"] == 3


@pytest.mark.asyncio
async def test_create_sheet(sample_excel_file):
    """Test creating a sheet."""
    from mcp_excel.tools.write import create_sheet
    
    result = await create_sheet(
        file_path=str(sample_excel_file),
        sheet_name="NewSheet",
    )
    
    assert result["success"] is True
    assert result["sheet_name"] == "NewSheet"


@pytest.mark.asyncio
async def test_get_formula_templates():
    """Test getting formula templates."""
    from mcp_excel.tools.formulas import get_formula_templates
    
    result = await get_formula_templates()
    
    assert result["success"] is True
    assert "sum" in result["templates"]
    assert "average" in result["templates"]


@pytest.mark.asyncio
async def test_backend_open(sample_excel_file):
    """Test backend operations."""
    from mcp_excel.backends.factory import create_backend
    
    backend = create_backend(sample_excel_file)
    backend.open(sample_excel_file)
    
    assert backend.get_sheet_names() == ["TestSheet"]
    assert backend.read_cell("TestSheet", "A1") == "Name"
    
    backend.close()
