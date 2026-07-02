"""Integration tests for table tools."""

from pathlib import Path

import pytest

from mcp_excel.tools.tables import (
    add_table_row,
    create_table,
    delete_table,
    list_tables,
)


@pytest.mark.asyncio
class TestCreateTable:
    """Tests for create_table tool."""

    async def test_create_table(self, sample_excel_file: Path):
        """Test creating a table."""
        result = await create_table(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            table_name="TestTable",
            data_range="A1:D4",
        )
        assert result["success"] is True
        assert result["table_name"] == "TestTable"
        assert result["data_range"] == "A1:D4"

    async def test_create_table_with_style(self, sample_excel_file: Path):
        """Test creating a table with custom style."""
        result = await create_table(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            table_name="StyledTable",
            data_range="A1:D4",
            style="TableStyleLight1",
        )
        assert result["success"] is True
        assert result["style"] == "TableStyleLight1"

    async def test_create_table_nonexistent_file(self):
        """Test creating table in nonexistent file."""
        result = await create_table(
            file_path="/nonexistent/file.xlsx",
            sheet_name="Sheet1",
            table_name="TestTable",
            data_range="A1:D1",
        )
        assert result["success"] is False
        assert "error" in result


@pytest.mark.asyncio
class TestListTables:
    """Tests for list_tables tool."""

    async def test_list_tables_empty(self, sample_excel_file: Path):
        """Test listing tables when none exist."""
        result = await list_tables(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
        )
        assert result["success"] is True
        assert result["count"] == 0

    async def test_list_tables_after_create(self, sample_excel_file: Path):
        """Test listing tables after creating one."""
        await create_table(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            table_name="TestTable",
            data_range="A1:D4",
        )

        result = await list_tables(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
        )
        assert result["success"] is True
        assert result["count"] == 1
        assert result["tables"][0]["name"] == "TestTable"

    async def test_list_tables_nonexistent_file(self):
        """Test listing tables in nonexistent file."""
        result = await list_tables(
            file_path="/nonexistent/file.xlsx",
            sheet_name="Sheet1",
        )
        assert result["success"] is False
        assert "error" in result


@pytest.mark.asyncio
class TestDeleteTable:
    """Tests for delete_table tool."""

    async def test_delete_table(self, sample_excel_file: Path):
        """Test deleting a table."""
        await create_table(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            table_name="ToDelete",
            data_range="A1:D4",
        )

        result = await delete_table(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            table_name="ToDelete",
        )
        assert result["success"] is True
        assert result["deleted_table"] == "ToDelete"

        # Verify table was deleted
        list_result = await list_tables(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
        )
        assert list_result["count"] == 0

    async def test_delete_table_nonexistent(self, sample_excel_file: Path):
        """Test deleting nonexistent table."""
        result = await delete_table(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            table_name="Nonexistent",
        )
        assert result["success"] is False
        assert "error" in result


@pytest.mark.asyncio
class TestAddTableRow:
    """Tests for add_table_row tool."""

    async def test_add_table_row(self, sample_excel_file: Path):
        """Test adding a row to a table."""
        await create_table(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            table_name="AddRowTable",
            data_range="A1:D4",
        )

        result = await add_table_row(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            table_name="AddRowTable",
            values=["New Item", 100, 19.99, "West"],
        )
        assert result["success"] is True
        assert result["values"] == ["New Item", 100, 19.99, "West"]

    async def test_add_table_row_nonexistent_table(self, sample_excel_file: Path):
        """Test adding row to nonexistent table."""
        result = await add_table_row(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            table_name="Nonexistent",
            values=["test"],
        )
        assert result["success"] is False
        assert "error" in result
