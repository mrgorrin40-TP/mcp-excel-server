"""MCP tools for writing Excel data."""

import logging
from typing import Annotated, Any

from fastmcp import FastMCP
from pydantic import Field

from ..utils.backend import get_backend

logger = logging.getLogger(__name__)

# Create tools router
tools = FastMCP("Excel Writing Tools", mask_error_details=True)


@tools.tool(
    name="write_cells",
    description="Write values to cells in an Excel file",
    tags={"excel", "write"},
)
async def write_cells(
    file_path: Annotated[str, Field(description="Absolute path to Excel file")],
    sheet_name: Annotated[str, Field(description="Worksheet name")],
    cell_range: Annotated[str, Field(description="Target range (e.g., 'A1:C3')")],
    values: Annotated[list[list[Any]], Field(description="2D array of values to write")],
) -> dict[str, Any]:
    """Write values to cells in an Excel file."""
    try:
        backend = get_backend(file_path)
        backend.write_range(sheet_name, cell_range, values)
        backend.save()

        # Count cells written
        rows = len(values)
        cols = max(len(row) for row in values) if values else 0

        return {
            "success": True,
            "cells_written": rows * cols,
            "range": cell_range,
            "sheet_name": sheet_name,
        }
    except Exception as e:
        logger.error("Error writing cells: %s", e)
        return {"success": False, "error": str(e)}


@tools.tool(
    name="write_formula",
    description="Add an Excel formula to a cell",
    tags={"excel", "write", "formula"},
)
async def write_formula(
    file_path: Annotated[str, Field(description="Absolute path to Excel file")],
    sheet_name: Annotated[str, Field(description="Worksheet name")],
    cell: Annotated[str, Field(description="Target cell (e.g., 'D1')")],
    formula: Annotated[str, Field(description="Excel formula (with = prefix)")],
) -> dict[str, Any]:
    """Add an Excel formula to a cell."""
    try:
        backend = get_backend(file_path)

        # Ensure formula starts with =
        if not formula.startswith("="):
            formula = "=" + formula

        backend.write_cell(sheet_name, cell, formula)
        backend.save()

        return {
            "success": True,
            "cell": cell,
            "formula": formula,
            "sheet_name": sheet_name,
        }
    except Exception as e:
        logger.error("Error writing formula: %s", e)
        return {"success": False, "error": str(e)}


@tools.tool(
    name="create_sheet",
    description="Create a new worksheet in an Excel file",
    tags={"excel", "write", "sheet"},
)
async def create_sheet(
    file_path: Annotated[str, Field(description="Absolute path to Excel file")],
    sheet_name: Annotated[str, Field(description="Name for new worksheet")],
) -> dict[str, Any]:
    """Create a new worksheet in an Excel file."""
    try:
        backend = get_backend(file_path)
        backend.create_sheet(sheet_name)
        backend.save()

        return {
            "success": True,
            "sheet_name": sheet_name,
            "message": f"Sheet '{sheet_name}' created successfully",
        }
    except Exception as e:
        logger.error("Error creating sheet: %s", e)
        return {"success": False, "error": str(e)}


@tools.tool(
    name="delete_sheet",
    description="Delete a worksheet from an Excel file",
    tags={"excel", "write", "sheet"},
)
async def delete_sheet(
    file_path: Annotated[str, Field(description="Absolute path to Excel file")],
    sheet_name: Annotated[str, Field(description="Worksheet name to delete")],
) -> dict[str, Any]:
    """Delete a worksheet from an Excel file."""
    try:
        backend = get_backend(file_path)
        backend.delete_sheet(sheet_name)
        backend.save()

        return {
            "success": True,
            "sheet_name": sheet_name,
            "message": f"Sheet '{sheet_name}' deleted successfully",
        }
    except Exception as e:
        logger.error("Error deleting sheet: %s", e)
        return {"success": False, "error": str(e)}
