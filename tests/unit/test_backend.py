"""Unit tests for Excel backends."""

from pathlib import Path

import pytest

from mcp_excel.backends.factory import create_backend
from mcp_excel.backends.openpyxl_backend import OpenpyxlBackend


class TestOpenpyxlBackend:
    """Tests for OpenpyxlBackend."""

    def test_open_and_close(self, sample_excel_file: Path):
        """Test opening and closing a workbook."""
        backend = OpenpyxlBackend()
        backend.open(sample_excel_file)
        assert backend._workbook is not None
        backend.close()
        assert backend._workbook is None

    def test_get_sheet_names(self, sample_excel_file: Path):
        """Test getting sheet names."""
        backend = OpenpyxlBackend()
        backend.open(sample_excel_file)
        names = backend.get_sheet_names()
        assert "TestSheet" in names
        assert len(names) == 1
        backend.close()

    def test_read_cell_string(self, sample_excel_file: Path):
        """Test reading a string cell."""
        backend = OpenpyxlBackend()
        backend.open(sample_excel_file)
        value = backend.read_cell("TestSheet", "A1")
        assert value == "Name"
        backend.close()

    def test_read_cell_number(self, sample_excel_file: Path):
        """Test reading a numeric cell."""
        backend = OpenpyxlBackend()
        backend.open(sample_excel_file)
        value = backend.read_cell("TestSheet", "B2")
        assert value == 25
        backend.close()

    def test_read_cell_empty(self, sample_excel_file: Path):
        """Test reading an empty cell."""
        backend = OpenpyxlBackend()
        backend.open(sample_excel_file)
        value = backend.read_cell("TestSheet", "Z1")
        assert value is None
        backend.close()

    def test_read_range(self, sample_excel_file: Path):
        """Test reading a range of cells."""
        backend = OpenpyxlBackend()
        backend.open(sample_excel_file)
        data = backend.read_range("TestSheet", "A1:C2")
        assert len(data) == 2
        assert data[0] == ["Name", "Age", "City"]
        assert data[1][0] == "Alice"
        backend.close()

    def test_write_cell(self, sample_excel_file: Path):
        """Test writing a cell value."""
        backend = OpenpyxlBackend()
        backend.open(sample_excel_file)
        backend.write_cell("TestSheet", "E1", "New Header")
        value = backend.read_cell("TestSheet", "E1")
        assert value == "New Header"
        backend.close()

    def test_write_range(self, sample_excel_file: Path):
        """Test writing a range of values."""
        backend = OpenpyxlBackend()
        backend.open(sample_excel_file)
        values = [["X", "Y", "Z"], [1, 2, 3]]
        backend.write_range("TestSheet", "E1:G2", values)
        data = backend.read_range("TestSheet", "E1:G2")
        assert data == [["X", "Y", "Z"], [1, 2, 3]]
        backend.close()

    def test_create_sheet(self, sample_excel_file: Path):
        """Test creating a new sheet."""
        backend = OpenpyxlBackend()
        backend.open(sample_excel_file)
        backend.create_sheet("NewSheet")
        names = backend.get_sheet_names()
        assert "NewSheet" in names
        backend.close()

    def test_delete_sheet(self, sample_excel_file: Path):
        """Test deleting a sheet."""
        backend = OpenpyxlBackend()
        backend.open(sample_excel_file)
        backend.create_sheet("ToDelete")
        backend.delete_sheet("ToDelete")
        names = backend.get_sheet_names()
        assert "ToDelete" not in names
        backend.close()

    def test_get_used_range(self, sample_excel_file: Path):
        """Test getting used range."""
        backend = OpenpyxlBackend()
        backend.open(sample_excel_file)
        used_range = backend.get_used_range("TestSheet")
        assert used_range is not None
        assert "A1" in used_range
        backend.close()

    def test_get_cell_type(self, sample_excel_file: Path):
        """Test getting cell type."""
        backend = OpenpyxlBackend()
        backend.open(sample_excel_file)
        assert backend.get_cell_type("TestSheet", "A1") == "string"
        assert backend.get_cell_type("TestSheet", "B2") == "number"
        assert backend.get_cell_type("TestSheet", "Z1") == "null"
        backend.close()

    def test_save(self, sample_excel_file: Path):
        """Test saving workbook."""
        backend = OpenpyxlBackend()
        backend.open(sample_excel_file)
        backend.write_cell("TestSheet", "E1", "Saved")
        backend.save()
        backend.close()

        # Reopen and verify
        backend2 = OpenpyxlBackend()
        backend2.open(sample_excel_file)
        value = backend2.read_cell("TestSheet", "E1")
        assert value == "Saved"
        backend2.close()


class TestBackendFactory:
    """Tests for backend factory."""

    def test_create_openpyxl_backend(self, sample_excel_file: Path):
        """Test creating openpyxl backend for .xlsx file."""
        backend = create_backend(sample_excel_file)
        assert isinstance(backend, OpenpyxlBackend)

    def test_create_backend_invalid_extension(self, tmp_path: Path):
        """Test creating backend for invalid file extension."""
        invalid_file = tmp_path / "test.csv"
        invalid_file.write_text("a,b,c")
        with pytest.raises(ValueError, match="Unsupported file format"):
            create_backend(invalid_file)

    def test_backend_interface_compliance(self, sample_excel_file: Path):
        """Test that backend implements required interface."""
        backend = create_backend(sample_excel_file)
        assert hasattr(backend, "open")
        assert hasattr(backend, "close")
        assert hasattr(backend, "read_cell")
        assert hasattr(backend, "read_range")
        assert hasattr(backend, "write_cell")
        assert hasattr(backend, "write_range")
        assert hasattr(backend, "get_sheet_names")
        assert hasattr(backend, "create_sheet")
