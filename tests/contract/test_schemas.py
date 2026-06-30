"""Contract tests for tool schemas."""

import pytest
from typing import Any, Dict, List

from mcp_excel.tools.read import read_cell, read_range, get_sheet_info, search_cells
from mcp_excel.tools.write import write_cells, write_formula, create_sheet
from mcp_excel.tools.inspect import list_sheets, describe_workbook, get_column_stats, filter_rows, group_by
from mcp_excel.tools.formulas import write_formula as formulas_write, read_formula, get_formula_templates


class TestToolSchemas:
    """Contract tests for tool parameter schemas."""

    def test_read_cell_has_correct_parameters(self):
        """Test read_cell has required parameters."""
        import inspect
        sig = inspect.signature(read_cell)
        params = list(sig.parameters.keys())
        assert "file_path" in params
        assert "sheet_name" in params
        assert "cell" in params

    def test_read_range_has_correct_parameters(self):
        """Test read_range has required parameters."""
        import inspect
        sig = inspect.signature(read_range)
        params = list(sig.parameters.keys())
        assert "file_path" in params
        assert "sheet_name" in params
        assert "range" in params
        assert "page_size" in params

    def test_get_sheet_info_has_correct_parameters(self):
        """Test get_sheet_info has required parameters."""
        import inspect
        sig = inspect.signature(get_sheet_info)
        params = list(sig.parameters.keys())
        assert "file_path" in params
        assert "sheet_name" in params

    def test_search_cells_has_correct_parameters(self):
        """Test search_cells has required parameters."""
        import inspect
        sig = inspect.signature(search_cells)
        params = list(sig.parameters.keys())
        assert "file_path" in params
        assert "query" in params
        assert "sheet_name" in params

    def test_write_cells_has_correct_parameters(self):
        """Test write_cells has required parameters."""
        import inspect
        sig = inspect.signature(write_cells)
        params = list(sig.parameters.keys())
        assert "file_path" in params
        assert "sheet_name" in params
        assert "range" in params
        assert "values" in params

    def test_write_formula_has_correct_parameters(self):
        """Test write_formula has required parameters."""
        import inspect
        sig = inspect.signature(write_formula)
        params = list(sig.parameters.keys())
        assert "file_path" in params
        assert "sheet_name" in params
        assert "cell" in params
        assert "formula" in params

    def test_create_sheet_has_correct_parameters(self):
        """Test create_sheet has required parameters."""
        import inspect
        sig = inspect.signature(create_sheet)
        params = list(sig.parameters.keys())
        assert "file_path" in params
        assert "sheet_name" in params

    def test_list_sheets_has_correct_parameters(self):
        """Test list_sheets has required parameters."""
        import inspect
        sig = inspect.signature(list_sheets)
        params = list(sig.parameters.keys())
        assert "file_path" in params

    def test_describe_workbook_has_correct_parameters(self):
        """Test describe_workbook has required parameters."""
        import inspect
        sig = inspect.signature(describe_workbook)
        params = list(sig.parameters.keys())
        assert "file_path" in params

    def test_get_column_stats_has_correct_parameters(self):
        """Test get_column_stats has required parameters."""
        import inspect
        sig = inspect.signature(get_column_stats)
        params = list(sig.parameters.keys())
        assert "file_path" in params
        assert "sheet_name" in params
        assert "column" in params

    def test_filter_rows_has_correct_parameters(self):
        """Test filter_rows has required parameters."""
        import inspect
        sig = inspect.signature(filter_rows)
        params = list(sig.parameters.keys())
        assert "file_path" in params
        assert "sheet_name" in params
        assert "filters" in params
        assert "logic" in params

    def test_group_by_has_correct_parameters(self):
        """Test group_by has required parameters."""
        import inspect
        sig = inspect.signature(group_by)
        params = list(sig.parameters.keys())
        assert "file_path" in params
        assert "sheet_name" in params
        assert "columns" in params
        assert "agg_column" in params
        assert "agg_func" in params


class TestToolReturnTypes:
    """Contract tests for tool return types."""

    @pytest.mark.asyncio
    async def test_read_cell_returns_dict(self, sample_excel_file):
        """Test read_cell returns dictionary."""
        result = await read_cell(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            cell="A1",
        )
        assert isinstance(result, dict)
        assert "success" in result

    @pytest.mark.asyncio
    async def test_read_range_returns_dict(self, sample_excel_file):
        """Test read_range returns dictionary."""
        result = await read_range(
            file_path=str(sample_excel_file),
            sheet_name="TestSheet",
            range="A1:D4",
        )
        assert isinstance(result, dict)
        assert "success" in result
        assert "data" in result

    @pytest.mark.asyncio
    async def test_list_sheets_returns_dict(self, sample_excel_file):
        """Test list_sheets returns dictionary."""
        result = await list_sheets(
            file_path=str(sample_excel_file),
        )
        assert isinstance(result, dict)
        assert "success" in result
        assert "sheets" in result

    @pytest.mark.asyncio
    async def test_describe_workbook_returns_dict(self, sample_excel_file):
        """Test describe_workbook returns dictionary."""
        result = await describe_workbook(
            file_path=str(sample_excel_file),
        )
        assert isinstance(result, dict)
        assert "success" in result
        assert "file" in result

    @pytest.mark.asyncio
    async def test_get_formula_templates_returns_dict(self):
        """Test get_formula_templates returns dictionary."""
        result = await get_formula_templates()
        assert isinstance(result, dict)
        assert "success" in result
        assert "templates" in result


class TestParameterNaming:
    """Contract tests for parameter naming conventions."""

    def test_all_parameters_use_snake_case(self):
        """Test that all tool parameters use snake_case."""
        import inspect
        import re

        snake_case_pattern = re.compile(r'^[a-z][a-z0-9]*(_[a-z0-9]+)*$')

        tools = [
            read_cell, read_range, get_sheet_info, search_cells,
            write_cells, write_formula, create_sheet,
            list_sheets, describe_workbook, get_column_stats, filter_rows, group_by,
            formulas_write, read_formula, get_formula_templates,
        ]

        for tool in tools:
            sig = inspect.signature(tool)
            for param_name in sig.parameters:
                assert snake_case_pattern.match(param_name), \
                    f"Parameter {param_name} in {tool.__name__} is not snake_case"
