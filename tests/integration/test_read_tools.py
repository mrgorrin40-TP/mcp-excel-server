"""Integration tests for read tools."""

import pytest
from pathlib import Path

from mcp_excel.tools.read import read_cell, read_range, get_sheet_info, search_cells


@pytest.mark.asyncio
class TestReadCell:
    """Tests for read_cell tool."""

    async def test_read_string_cell(self, sample_excel_file: Path):
        """Test reading a string cell."""
        result = await read_cell(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            cell="A1",
        )
        assert result["success"] is True
        assert result["value"] == "Name"
        assert result["type"] == "string"

    async def test_read_numeric_cell(self, sample_excel_file: Path):
        """Test reading a numeric cell."""
        result = await read_cell(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            cell="B2",
        )
        assert result["success"] is True
        assert result["value"] == 25
        assert result["type"] == "number"

    async def test_read_empty_cell(self, sample_excel_file: Path):
        """Test reading an empty cell."""
        result = await read_cell(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            cell="Z1",
        )
        assert result["success"] is True
        assert result["value"] is None
        assert result["type"] == "null"

    async def test_read_nonexistent_file(self):
        """Test reading from nonexistent file."""
        result = await read_cell(
            file_path="/nonexistent/file.xlsx",
            sheet_name="Sheet1",
            cell="A1",
        )
        assert result["success"] is False
        assert "error" in result

    async def test_read_nonexistent_sheet(self, sample_excel_file: Path):
        """Test reading from nonexistent sheet."""
        result = await read_cell(
            file_path=str(sample_excel_file),
            sheet_name="NonexistentSheet",
            cell="A1",
        )
        assert result["success"] is False
        assert "error" in result


@pytest.mark.asyncio
class TestReadRange:
    """Tests for read_range tool."""

    async def test_read_full_range(self, sample_excel_file: Path):
        """Test reading a full range."""
        result = await read_range(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            range="A1:D4",
        )
        assert result["success"] is True
        assert len(result["data"]) == 4
        assert result["data"][0] == ["Name", "Age", "City", "Salary"]

    async def test_read_single_cell_range(self, sample_excel_file: Path):
        """Test reading a single cell range."""
        result = await read_range(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            range="A1:A1",
        )
        assert result["success"] is True
        assert result["data"] == [["Name"]]

    async def test_read_with_pagination(self, sample_excel_file: Path):
        """Test reading with pagination."""
        result = await read_range(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            range="A1:D4",
            page_size=2,
        )
        assert result["success"] is True
        assert len(result["data"]) == 2
        assert result["page"] == 1
        assert result["page_size"] == 2

    async def test_read_second_page(self, sample_excel_file: Path):
        """Test reading second page."""
        result = await read_range(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            range="A1:D4",
            page_size=2,
            page=2,
        )
        assert result["success"] is True
        assert len(result["data"]) == 2
        assert result["page"] == 2


@pytest.mark.asyncio
class TestGetSheetInfo:
    """Tests for get_sheet_info tool."""

    async def test_get_sheet_info(self, sample_excel_file: Path):
        """Test getting sheet information."""
        result = await get_sheet_info(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
        )
        assert result["success"] is True
        assert result["name"] == "TestSheet"
        assert result["row_count"] == 4
        assert result["column_count"] == 4
        assert len(result["columns"]) == 4

    async def test_get_sheet_info_with_headers(self, sample_excel_file: Path):
        """Test that headers are detected."""
        result = await get_sheet_info(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
        )
        assert result["success"] is True
        assert len(result["columns"]) > 0


@pytest.mark.asyncio
class TestSearchCells:
    """Tests for search_cells tool."""

    async def test_search_existing_value(self, sample_excel_file: Path):
        """Test searching for existing value."""
        result = await search_cells(
            file_path=str(sample_excel_file),
            query="Alice",
        )
        assert result["success"] is True
        assert result["total_matches"] >= 1
        assert any(m["value"] == "Alice" for m in result["matches"])

    async def test_search_nonexistent_value(self, sample_excel_file: Path):
        """Test searching for nonexistent value."""
        result = await search_cells(
            file_path=str(sample_excel_file),
            query="Nonexistent",
        )
        assert result["success"] is True
        assert result["total_matches"] == 0

    async def test_search_case_insensitive(self, sample_excel_file: Path):
        """Test case-insensitive search."""
        result = await search_cells(
            file_path=str(sample_excel_file),
            query="alice",
        )
        assert result["success"] is True
        assert result["total_matches"] >= 1

    async def test_search_specific_sheet(self, sample_excel_file: Path):
        """Test searching in specific sheet."""
        result = await search_cells(
            file_path=str(sample_excel_file),
            query="Alice",
            sheet_name="TestSheet",
        )
        assert result["success"] is True
        assert result["total_matches"] >= 1
