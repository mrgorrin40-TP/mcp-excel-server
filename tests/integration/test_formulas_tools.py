"""Integration tests for formula tools."""

from pathlib import Path

import pytest

from mcp_excel.tools.formulas import get_formula_templates, read_formula
from mcp_excel.tools.write import write_formula


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

    async def test_write_complex_formula(self, sample_excel_file: Path):
        """Test writing a complex formula."""
        result = await write_formula(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            cell="E4",
            formula='=IF(B2>30, "Senior", "Junior")',
        )
        assert result["success"] is True


@pytest.mark.asyncio
class TestReadFormula:
    """Tests for read_formula tool."""

    async def test_read_formula(self, sample_excel_file: Path):
        """Test reading a formula."""
        from mcp_excel.tools.write import write_formula as wf

        await wf(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            cell="E1",
            formula="=SUM(B2:B4)",
        )

        result = await read_formula(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            cell="E1",
        )
        assert result["success"] is True
        assert result.get("formula") == "=SUM(B2:B4)"

    async def test_read_formula_without_formula(self, sample_excel_file: Path):
        """Test reading a cell without formula."""
        result = await read_formula(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            cell="A1",
        )
        assert result["success"] is True
        assert result.get("is_formula") is False

    async def test_read_formula_nonexistent_cell(self, sample_excel_file: Path):
        """Test reading formula from nonexistent cell."""
        result = await read_formula(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            cell="Z100",
        )
        assert result["success"] is True
        assert result.get("is_formula") is False


@pytest.mark.asyncio
class TestGetFormulaTemplates:
    """Tests for get_formula_templates tool."""

    async def test_get_formula_templates(self):
        """Test getting formula templates."""
        result = await get_formula_templates()
        assert result["success"] is True
        assert "templates" in result
        assert "sum" in result["templates"]
        assert "average" in result["templates"]
        assert "count" in result["templates"]

    async def test_template_structure(self):
        """Test template structure."""
        result = await get_formula_templates()
        assert result["success"] is True
        template = result["templates"]["sum"]
        assert isinstance(template, str)
        assert "SUM" in template
