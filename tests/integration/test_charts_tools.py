"""Integration tests for chart tools."""

from pathlib import Path

import pytest

from mcp_excel.tools.charts import create_chart, delete_chart, list_charts, modify_chart


@pytest.mark.asyncio
class TestCreateChart:
    """Tests for create_chart tool."""

    async def test_create_bar_chart(self, sample_excel_file: Path):
        """Test creating a bar chart."""
        result = await create_chart(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            chart_type="bar",
            title="Test Bar Chart",
            data_range="A1:B4",
        )
        assert result["success"] is True
        assert result["chart_type"] == "bar"
        assert result["title"] == "Test Bar Chart"

    async def test_create_line_chart(self, sample_excel_file: Path):
        """Test creating a line chart."""
        result = await create_chart(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            chart_type="line",
            title="Test Line Chart",
            data_range="A1:B4",
        )
        assert result["success"] is True
        assert result["chart_type"] == "line"

    async def test_create_pie_chart(self, sample_excel_file: Path):
        """Test creating a pie chart."""
        result = await create_chart(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            chart_type="pie",
            title="Test Pie Chart",
            data_range="A1:B4",
        )
        assert result["success"] is True
        assert result["chart_type"] == "pie"

    async def test_create_scatter_chart(self, sample_excel_file: Path):
        """Test creating a scatter chart."""
        result = await create_chart(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            chart_type="scatter",
            title="Test Scatter Chart",
            data_range="A1:B4",
        )
        assert result["success"] is True
        assert result["chart_type"] == "scatter"

    async def test_create_chart_with_position(self, sample_excel_file: Path):
        """Test creating a chart with custom position."""
        result = await create_chart(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            chart_type="bar",
            title="Positioned Chart",
            data_range="A1:B4",
            position="F2",
        )
        assert result["success"] is True
        assert result["position"] == "F2"

    async def test_create_chart_invalid_type(self, sample_excel_file: Path):
        """Test creating chart with invalid type."""
        result = await create_chart(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            chart_type="invalid",
            title="Invalid Chart",
            data_range="A1:B4",
        )
        assert result["success"] is False
        assert "error" in result

    async def test_create_chart_nonexistent_file(self):
        """Test creating chart in nonexistent file."""
        result = await create_chart(
            file_path="/nonexistent/file.xlsx",
            sheet_name="Sheet1",
            chart_type="bar",
            title="Test",
            data_range="A1:B1",
        )
        assert result["success"] is False
        assert "error" in result


@pytest.mark.asyncio
class TestListCharts:
    """Tests for list_charts tool."""

    async def test_list_charts_empty(self, sample_excel_file: Path):
        """Test listing charts when none exist."""
        result = await list_charts(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
        )
        assert result["success"] is True
        assert result["count"] == 0

    async def test_list_charts_after_create(self, sample_excel_file: Path):
        """Test listing charts after creating one."""
        await create_chart(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            chart_type="bar",
            title="Test Chart",
            data_range="A1:B4",
        )

        result = await list_charts(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
        )
        assert result["success"] is True
        assert result["count"] == 1
        assert result["charts"][0]["title"] == "Test Chart"

    async def test_list_charts_nonexistent_file(self):
        """Test listing charts in nonexistent file."""
        result = await list_charts(
            file_path="/nonexistent/file.xlsx",
            sheet_name="Sheet1",
        )
        assert result["success"] is False
        assert "error" in result


@pytest.mark.asyncio
class TestModifyChart:
    """Tests for modify_chart tool."""

    async def test_modify_chart_title(self, sample_excel_file: Path):
        """Test modifying chart title."""
        await create_chart(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            chart_type="bar",
            title="Original Title",
            data_range="A1:B4",
        )

        result = await modify_chart(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            chart_index=0,
            title="Modified Title",
        )
        assert result["success"] is True
        assert "title" in result["modified_properties"]

    async def test_modify_chart_invalid_index(self, sample_excel_file: Path):
        """Test modifying chart with invalid index."""
        result = await modify_chart(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            chart_index=99,
            title="Test",
        )
        assert result["success"] is False
        assert "error" in result


@pytest.mark.asyncio
class TestDeleteChart:
    """Tests for delete_chart tool."""

    async def test_delete_chart(self, sample_excel_file: Path):
        """Test deleting a chart."""
        await create_chart(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            chart_type="bar",
            title="To Delete",
            data_range="A1:B4",
        )

        result = await delete_chart(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            chart_index=0,
        )
        assert result["success"] is True
        assert result["chart_title"] == "To Delete"

        # Verify chart was deleted
        list_result = await list_charts(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
        )
        assert list_result["count"] == 0

    async def test_delete_chart_invalid_index(self, sample_excel_file: Path):
        """Test deleting chart with invalid index."""
        result = await delete_chart(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            chart_index=99,
        )
        assert result["success"] is False
        assert "error" in result
