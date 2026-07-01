"""Tests for VBA tools."""

import pytest
from unittest.mock import Mock, patch

from mcp_excel.tools.vba import (
    list_vba_modules,
    get_vba_code,
    set_vba_code,
    add_vba_module,
    delete_vba_module,
    rename_vba_module,
    run_macro,
    list_macros,
    get_vba_templates,
    validate_vba_code,
    import_vba_module,
    export_vba_module,
)


def _create_mock_backend():
    """Create a mock backend with VBA support."""
    backend = Mock()
    backend.has_macros.return_value = True
    backend.list_vba_modules.return_value = [
        {"name": "Module1", "type": "standard", "description": ""},
        {"name": "ThisWorkbook", "type": "document", "description": ""},
    ]
    backend.get_vba_code.return_value = 'Sub Test()\n    MsgBox "Hello"\nEnd Sub'
    backend.list_macros.return_value = [
        {"name": "Test", "type": "Sub", "module": "Module1", "is_public": True},
    ]
    return backend


@pytest.mark.asyncio
class TestListVbaModules:
    """Tests for list_vba_modules tool."""

    @patch("mcp_excel.tools.vba.get_backend")
    async def test_list_modules_success(self, mock_get_backend):
        """Test listing VBA modules successfully."""
        mock_get_backend.return_value = _create_mock_backend()

        result = await list_vba_modules(file_path="test.xlsm")

        assert result["success"] is True
        assert result["count"] == 2
        assert result["modules"][0]["name"] == "Module1"

    @patch("mcp_excel.tools.vba.get_backend")
    async def test_list_modules_no_macros(self, mock_get_backend):
        """Test listing modules when file has no macros."""
        backend = Mock()
        backend.has_macros.return_value = False
        mock_get_backend.return_value = backend

        result = await list_vba_modules(file_path="test.xlsx")

        assert result["success"] is False
        assert "no VBA macros" in result["error"]


@pytest.mark.asyncio
class TestGetVbaCode:
    """Tests for get_vba_code tool."""

    @patch("mcp_excel.tools.vba.get_backend")
    async def test_get_code_success(self, mock_get_backend):
        """Test getting VBA code successfully."""
        mock_get_backend.return_value = _create_mock_backend()

        result = await get_vba_code(file_path="test.xlsm", module_name="Module1")

        assert result["success"] is True
        assert result["code"] is not None
        assert "Sub Test()" in result["code"]


@pytest.mark.asyncio
class TestSetVbaCode:
    """Tests for set_vba_code tool."""

    @patch("mcp_excel.tools.vba.get_backend")
    async def test_set_code_success(self, mock_get_backend):
        """Test setting VBA code successfully."""
        backend = _create_mock_backend()
        mock_get_backend.return_value = backend

        new_code = 'Sub NewMacro()\n    MsgBox "New"\nEnd Sub'
        result = await set_vba_code(
            file_path="test.xlsm",
            module_name="Module1",
            code=new_code,
        )

        assert result["success"] is True
        assert result["code_length"] == len(new_code)
        backend.set_vba_code.assert_called_once_with("Module1", new_code)
        backend.save.assert_called_once()


@pytest.mark.asyncio
class TestAddVbaModule:
    """Tests for add_vba_module tool."""

    @patch("mcp_excel.tools.vba.get_backend")
    async def test_add_module_success(self, mock_get_backend):
        """Test adding VBA module successfully."""
        backend = _create_mock_backend()
        mock_get_backend.return_value = backend

        result = await add_vba_module(
            file_path="test.xlsm",
            module_name="NewModule",
            code='Sub New()\nEnd Sub',
            module_type="standard",
        )

        assert result["success"] is True
        assert result["module_name"] == "NewModule"
        backend.add_vba_module.assert_called_once()

    @patch("mcp_excel.tools.vba.get_backend")
    async def test_add_module_invalid_type(self, mock_get_backend):
        """Test adding module with invalid type."""
        mock_get_backend.return_value = _create_mock_backend()

        result = await add_vba_module(
            file_path="test.xlsm",
            module_name="NewModule",
            module_type="invalid",
        )

        assert result["success"] is False
        assert "Invalid module type" in result["error"]


@pytest.mark.asyncio
class TestDeleteVbaModule:
    """Tests for delete_vba_module tool."""

    @patch("mcp_excel.tools.vba.get_backend")
    async def test_delete_module_success(self, mock_get_backend):
        """Test deleting VBA module successfully."""
        backend = _create_mock_backend()
        mock_get_backend.return_value = backend

        result = await delete_vba_module(
            file_path="test.xlsm",
            module_name="Module1",
        )

        assert result["success"] is True
        assert result["deleted_module"] == "Module1"
        backend.delete_vba_module.assert_called_once_with("Module1")


@pytest.mark.asyncio
class TestRenameVbaModule:
    """Tests for rename_vba_module tool."""

    @patch("mcp_excel.tools.vba.get_backend")
    async def test_rename_module_success(self, mock_get_backend):
        """Test renaming VBA module successfully."""
        backend = _create_mock_backend()
        mock_get_backend.return_value = backend

        result = await rename_vba_module(
            file_path="test.xlsm",
            old_name="Module1",
            new_name="RenamedModule",
        )

        assert result["success"] is True
        assert result["old_name"] == "Module1"
        assert result["new_name"] == "RenamedModule"
        backend.rename_vba_module.assert_called_once_with("Module1", "RenamedModule")


@pytest.mark.asyncio
class TestRunMacro:
    """Tests for run_macro tool."""

    @patch("mcp_excel.tools.vba.get_backend")
    async def test_run_macro_success(self, mock_get_backend):
        """Test running macro successfully."""
        backend = _create_mock_backend()
        backend.run_macro.return_value = "Result"
        mock_get_backend.return_value = backend

        result = await run_macro(
            file_path="test.xlsm",
            macro_name="Module1.Test",
        )

        assert result["success"] is True
        assert result["macro_name"] == "Module1.Test"
        backend.run_macro.assert_called_once_with("Module1.Test")


@pytest.mark.asyncio
class TestListMacros:
    """Tests for list_macros tool."""

    @patch("mcp_excel.tools.vba.get_backend")
    async def test_list_macros_success(self, mock_get_backend):
        """Test listing macros successfully."""
        mock_get_backend.return_value = _create_mock_backend()

        result = await list_macros(file_path="test.xlsm")

        assert result["success"] is True
        assert result["count"] == 1
        assert result["macros"][0]["name"] == "Test"


@pytest.mark.asyncio
class TestGetVbaTemplates:
    """Tests for get_vba_templates tool."""

    async def test_get_all_templates(self):
        """Test getting all templates."""
        result = await get_vba_templates()

        assert result["success"] is True
        assert "hello_world" in result["templates"]
        assert "loop_range" in result["templates"]

    async def test_get_specific_template(self):
        """Test getting a specific template."""
        result = await get_vba_templates(template_type="hello_world")

        assert result["success"] is True
        assert "hello_world" in result["templates"]
        assert len(result["templates"]) == 1

    async def test_get_invalid_template(self):
        """Test getting an invalid template."""
        result = await get_vba_templates(template_type="nonexistent")

        assert result["success"] is False
        assert "not found" in result["error"]


@pytest.mark.asyncio
class TestValidateVbaCode:
    """Tests for validate_vba_code tool."""

    async def test_valid_code(self):
        """Test validation of valid VBA code."""
        code = 'Sub Test()\n    MsgBox "Hello"\nEnd Sub'
        result = await validate_vba_code(code=code)

        assert result["success"] is True
        assert result["valid"] is True
        assert result["errors"] == []
        assert "Test" in result["subs"]

    async def test_valid_function(self):
        """Test validation of valid Function."""
        code = 'Function Add(a As Integer, b As Integer) As Integer\n    Add = a + b\nEnd Function'
        result = await validate_vba_code(code=code)

        assert result["success"] is True
        assert result["valid"] is True
        assert "Add" in result["functions"]

    async def test_mismatched_sub(self):
        """Test validation of code with mismatched Sub/End Sub."""
        code = 'Sub Test1()\n    MsgBox "1"\nSub Test2()\n    MsgBox "2"\nEnd Sub'
        result = await validate_vba_code(code=code)

        assert result["success"] is True
        assert result["valid"] is False
        assert len(result["errors"]) > 0

    async def test_mismatched_function(self):
        """Test validation of code with mismatched Function/End Function."""
        code = 'Function Add()\n    Add = 1\nFunction Subtract()\n    Subtract = 2\nEnd Function'
        result = await validate_vba_code(code=code)

        assert result["success"] is True
        assert result["valid"] is False

    async def test_code_with_comments(self):
        """Test validation with comments."""
        code = "' This is a comment\nSub Test()\n    ' Another comment\n    MsgBox \"Hello\"\nEnd Sub"
        result = await validate_vba_code(code=code)

        assert result["success"] is True
        assert result["valid"] is True


@pytest.mark.asyncio
class TestImportVbaModule:
    """Tests for import_vba_module tool."""

    @patch("mcp_excel.tools.vba.get_backend")
    async def test_import_new_module(self, mock_get_backend):
        """Test importing to a new module."""
        backend = _create_mock_backend()
        backend.get_vba_code.side_effect = ValueError("Module not found")
        mock_get_backend.return_value = backend

        code = 'Sub Imported()\n    MsgBox "Imported"\nEnd Sub'
        result = await import_vba_module(
            file_path="test.xlsm",
            module_name="ImportedModule",
            code=code,
        )

        assert result["success"] is True
        assert result["module_name"] == "ImportedModule"
        backend.add_vba_module.assert_called_once()

    @patch("mcp_excel.tools.vba.get_backend")
    async def test_import_existing_module(self, mock_get_backend):
        """Test importing to an existing module."""
        mock_get_backend.return_value = _create_mock_backend()

        code = 'Sub Updated()\n    MsgBox "Updated"\nEnd Sub'
        result = await import_vba_module(
            file_path="test.xlsm",
            module_name="Module1",
            code=code,
        )

        assert result["success"] is True
        backend = mock_get_backend.return_value
        backend.set_vba_code.assert_called_once()


@pytest.mark.asyncio
class TestExportVbaModule:
    """Tests for export_vba_module tool."""

    @patch("mcp_excel.tools.vba.get_backend")
    async def test_export_module_success(self, mock_get_backend):
        """Test exporting VBA module successfully."""
        mock_get_backend.return_value = _create_mock_backend()

        result = await export_vba_module(
            file_path="test.xlsm",
            module_name="Module1",
        )

        assert result["success"] is True
        assert result["module_name"] == "Module1"
        assert "code" in result
