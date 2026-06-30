"""Pytest configuration and fixtures."""

import pytest
from pathlib import Path


@pytest.fixture
def sample_excel_file(tmp_path):
    """Create a sample Excel file for testing."""
    from openpyxl import Workbook
    
    wb = Workbook()
    ws = wb.active
    ws.title = "TestSheet"
    
    # Add headers
    ws["A1"] = "Name"
    ws["B1"] = "Age"
    ws["C1"] = "City"
    
    # Add data
    ws["A2"] = "Alice"
    ws["B2"] = 25
    ws["C2"] = "New York"
    
    ws["A3"] = "Bob"
    ws["B3"] = 30
    ws["C3"] = "London"
    
    ws["A4"] = "Charlie"
    ws["B4"] = 35
    ws["C4"] = "Paris"
    
    file_path = tmp_path / "test.xlsx"
    wb.save(file_path)
    return file_path


@pytest.fixture
def empty_excel_file(tmp_path):
    """Create an empty Excel file for testing."""
    from openpyxl import Workbook
    
    wb = Workbook()
    file_path = tmp_path / "empty.xlsx"
    wb.save(file_path)
    return file_path
