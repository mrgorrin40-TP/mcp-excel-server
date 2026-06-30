"""Integration tests for write tools."""

import pytest
from pathlib import Path

from mcp_excel.tools.write import write_cells, write_formula, create_sheet


@pytest.mark.asyncio
class TestWriteCells:
    """Tests for write_cells tool."""

    async def test_write_single_cell(self, sample_excel_file: Path):
        """Test writing a single cell."""
        result = await write_cells(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            range="E1",
            values=[["New Header"]],
        )
        assert result["success"] is True
        assert result["cells_written"] == 1

    async def test_write_multiple_cells(self, sample_excel_file: Path):
        """Test writing multiple cells."""
        result = await write_cells(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            range="E1:F2",
            values=[["X", "Y"], [1, 2]],
        )
        assert result["success"] is True
        assert result["cells_written"] == 4

    async def test_write_to_new_sheet(self, sample_excel_file: Path):
        """Test writing to a new sheet."""
        await create_sheet(
            file_path=str(sample_excel_file),
            sheet_name="NewSheet",
        )
        result = await write_cells(
            file_path=str(sample_excel_file),
            sheet_name="NewSheet",
            range="A1:C1",
            values=[["Col1", "Col2", "Col3"]],
        )
        assert result["success"] is True
        assert result["cells_written"] == 3

    async def test_write_nonexistent_file(self):
        """Test writing to nonexistent file."""
        result = await write_cells(
            file_path="/nonexistent/file.xlsx",
            sheet_name="Sheet1",
            range="A1",
            values=[["test"]],
        )
        assert result["success"] is False
        assert "error" in result

    async def test_write_nonexistent_sheet(self, sample_excel_file: Path):
        """Test writing to nonexistent sheet."""
        result = await write_cells(
            file_path=str(sample_excel_file),
            sheet_name="NonexistentSheet",
            range="A1",
            values=[["test"]],
        )
        assert result["success"] is False
        assert "error" in result


@pytest.mark.asyncio
class TestWriteFormula:
    """Tests for write_formula tool."""

    async def test_write_sum_formula(self, sample_excel_file: Path):
        """Test writing a SUM formula."""
        result = await write_formula(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            cell="E1",
            formula="=SUM(B2:B4)",
        )
        assert result["success"] is True
        assert result["cell"] == "E1"
        assert result["formula"] == "=SUM(B2:B4)"

    async def test_write_average_formula(self, sample_excel_file: Path):
        """Test writing an AVERAGE formula."""
        result = await write_formula(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            cell="E2",
            formula="=AVERAGE(B2:B4)",
        )
        assert result["success"] is True

    async def test_write_formula_without_equals(self, sample_excel_file: Path):
        """Test writing formula without = prefix."""
        result = await write_formula(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            cell="E3",
            formula="SUM(B2:B4)",
        )
        assert result["success"] is True
        assert result["formula"] == "=SUM(B2:B4)"


@pytest.mark.asyncio
class TestCreateSheet:
    """Tests for create_sheet tool."""

    async def test_create_sheet(self, sample_excel_file: Path):
        """Test creating a new sheet."""
        result = await create_sheet(
            file_path=str(sample_excel_file),
            sheet_name="NewSheet",
        )
        assert result["success"] is True
        assert result["sheet_name"] == "NewSheet"

    async def test_create_duplicate_sheet(self, sample_excel_file: Path):
        """Test creating duplicate sheet."""
        result = await create_sheet(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
        )
        assert result["success"] is False
        assert "error" in result

    async def test_create_sheet_nonexistent_file(self):
        """Test creating sheet in nonexistent file."""
        result = await create_sheet(
            file_path="/nonexistent/file.xlsx",
            sheet_name="NewSheet",
        )
        assert result["success"] is False
        assert "error" in result
