"""Integration tests for inspect tools."""

import pytest
from pathlib import Path

from mcp_excel.tools.read import list_sheets, describe_workbook
from mcp_excel.tools.inspect import get_column_stats, filter_rows, group_by


@pytest.mark.asyncio
class TestListSheets:
    """Tests for list_sheets tool."""

    async def test_list_sheets(self, sample_excel_file: Path):
        """Test listing sheets."""
        result = await list_sheets(
            file_path=str(sample_excel_file),
        )
        assert result["success"] is True
        assert "TestSheet" in result["sheets"]
        assert result["count"] == 1

    async def test_list_sheets_multiple(self, multi_sheet_excel_file: Path):
        """Test listing multiple sheets."""
        result = await list_sheets(
            file_path=str(multi_sheet_excel_file),
        )
        assert result["success"] is True
        assert result["count"] == 2
        assert "Sales" in result["sheets"]
        assert "Employees" in result["sheets"]

    async def test_list_sheets_nonexistent_file(self):
        """Test listing sheets in nonexistent file."""
        result = await list_sheets(
            file_path="/nonexistent/file.xlsx",
        )
        assert result["success"] is False
        assert "error" in result


@pytest.mark.asyncio
class TestDescribeWorkbook:
    """Tests for describe_workbook tool."""

    async def test_describe_workbook(self, sample_excel_file: Path):
        """Test describing a workbook."""
        result = await describe_workbook(
            file_path=str(sample_excel_file),
        )
        assert result["success"] is True
        assert result["total_sheets"] == 1
        assert result["file"]["name"] == "test.xlsx"
        assert result["file"]["size_kb"] >= 0

    async def test_describe_workbook_multiple_sheets(self, multi_sheet_excel_file: Path):
        """Test describing workbook with multiple sheets."""
        result = await describe_workbook(
            file_path=str(multi_sheet_excel_file),
        )
        assert result["success"] is True
        assert result["total_sheets"] == 2


@pytest.mark.asyncio
class TestGetColumnStats:
    """Tests for get_column_stats tool."""

    async def test_get_column_stats(self, sample_excel_file: Path):
        """Test getting column statistics."""
        result = await get_column_stats(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            column="B",
        )
        assert result["success"] is True
        assert result["count"] == 3
        assert result["mean"] == 30.0
        assert result["min"] == 25
        assert result["max"] == 35

    async def test_get_column_stats_text_column(self, sample_excel_file: Path):
        """Test getting stats for text column."""
        result = await get_column_stats(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            column="A",
        )
        assert result["success"] is True
        assert result["type"] == "no_numeric_data"


@pytest.mark.asyncio
class TestFilterRows:
    """Tests for filter_rows tool."""

    async def test_filter_rows_equals(self, sample_excel_file: Path):
        """Test filtering with equals operator."""
        result = await filter_rows(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            filters=[{"column": "Name", "operator": "==", "value": "Alice"}],
        )
        assert result["success"] is True
        assert result["row_count"] == 1
        assert result["data"][0][0] == "Alice"

    async def test_filter_rows_greater_than(self, sample_excel_file: Path):
        """Test filtering with greater than operator."""
        result = await filter_rows(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            filters=[{"column": "Age", "operator": ">", "value": 25}],
        )
        assert result["success"] is True
        assert result["row_count"] == 2

    async def test_filter_rows_no_matches(self, sample_excel_file: Path):
        """Test filtering with no matches."""
        result = await filter_rows(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            filters=[{"column": "Name", "operator": "==", "value": "Nonexistent"}],
        )
        assert result["success"] is True
        assert result["row_count"] == 0


@pytest.mark.asyncio
class TestGroupBy:
    """Tests for group_by tool."""

    async def test_group_by_sum(self, sample_excel_file: Path):
        """Test grouping with sum aggregation."""
        result = await group_by(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            columns=["City"],
            agg_column="Age",
            agg_func="sum",
        )
        assert result["success"] is True
        assert result["group_count"] > 0

    async def test_group_by_count(self, sample_excel_file: Path):
        """Test grouping with count aggregation."""
        result = await group_by(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            columns=["City"],
            agg_column="Age",
            agg_func="count",
        )
        assert result["success"] is True
        assert result["group_count"] > 0
