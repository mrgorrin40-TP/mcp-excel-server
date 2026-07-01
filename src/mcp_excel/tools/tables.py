"""MCP tools for Excel structured tables."""

import logging
from typing import Annotated, Any

from fastmcp import FastMCP
from pydantic import Field

from ..backends.factory import create_backend
from ..utils.cache import shared_cache

logger = logging.getLogger(__name__)

# Create tools router
tools = FastMCP("Excel Table Tools", mask_error_details=True)

# Table style options
TABLE_STYLES = [
    "TableStyleMedium2",
    "TableStyleMedium3",
    "TableStyleMedium4",
    "TableStyleMedium5",
    "TableStyleMedium6",
    "TableStyleMedium7",
    "TableStyleMedium8",
    "TableStyleMedium9",
    "TableStyleMedium10",
    "TableStyleLight1",
    "TableStyleLight2",
    "TableStyleLight3",
    "TableStyleLight4",
    "TableStyleLight5",
    "TableStyleLight6",
    "TableStyleLight7",
    "TableStyleLight8",
    "TableStyleLight9",
    "TableStyleLight10",
    "TableStyleDark1",
    "TableStyleDark2",
    "TableStyleDark3",
    "TableStyleDark4",
    "TableStyleDark5",
    "TableStyleDark6",
    "TableStyleDark7",
    "TableStyleDark8",
    "TableStyleDark9",
    "TableStyleDark10",
]


@tools.tool(
    name="create_table",
    description="Create a structured Excel table from a data range",
    tags={"excel", "write", "table"},
)
async def create_table(
    file_path: Annotated[str, Field(description="Absolute path to Excel file")],
    sheet_name: Annotated[str, Field(description="Worksheet name")],
    table_name: Annotated[str, Field(description="Name for the table")],
    data_range: Annotated[str, Field(description="Data range (e.g., 'A1:D100')")],
    has_header: Annotated[bool, Field(description="Whether data has headers")] = True,
    style: Annotated[
        str | None, Field(description="Table style (e.g., 'TableStyleMedium2')")
    ] = None,
) -> dict[str, Any]:
    """Create a structured Excel table from a data range."""
    try:
        backend = shared_cache.get(file_path)
        if backend is None:
            backend = create_backend(file_path)
            backend.open(file_path)

        ws = backend.get_sheet(sheet_name)

        # Import table classes
        from openpyxl.worksheet.table import Table, TableStyleInfo

        # Parse data range
        start, end = data_range.split(":")
        _col_letter_to_index(start)
        int("".join(c for c in start if c.isdigit()))
        _col_letter_to_index(end)
        int("".join(c for c in end if c.isdigit()))

        # Create table
        tab = Table(displayName=table_name, ref=data_range)

        # Set style
        if style and style in TABLE_STYLES:
            tab.tableStyleInfo = TableStyleInfo(
                name=style,
                showFirstColumn=False,
                showLastColumn=False,
                showRowStripes=True,
                showColumnStripes=False,
            )
        else:
            tab.tableStyleInfo = TableStyleInfo(
                name="TableStyleMedium2",
                showFirstColumn=False,
                showLastColumn=False,
                showRowStripes=True,
                showColumnStripes=False,
            )

        # Add table to worksheet
        ws.add_table(tab)

        backend.save()
        shared_cache.put(file_path, backend)

        return {
            "success": True,
            "table_name": table_name,
            "sheet_name": sheet_name,
            "data_range": data_range,
            "has_header": has_header,
            "style": style or "TableStyleMedium2",
        }
    except Exception as e:
        logger.error("Error creating table: %s", e)
        return {"success": False, "error": str(e)}


@tools.tool(
    name="list_tables",
    description="List all structured tables in an Excel worksheet",
    tags={"excel", "read", "table"},
)
async def list_tables(
    file_path: Annotated[str, Field(description="Absolute path to Excel file")],
    sheet_name: Annotated[str, Field(description="Worksheet name")],
) -> dict[str, Any]:
    """List all structured tables in an Excel worksheet."""
    try:
        backend = shared_cache.get(file_path)
        if backend is None:
            backend = create_backend(file_path)
            backend.open(file_path)
            shared_cache.put(file_path, backend)

        ws = backend.get_sheet(sheet_name)

        tables = []
        for table_name, table in ws.tables.items():
            tables.append({
                "name": table_name,
                "display_name": table.displayName,
                "ref": str(table.ref),
                "style": table.tableStyleInfo.name if table.tableStyleInfo else None,
            })

        return {
            "success": True,
            "sheet_name": sheet_name,
            "tables": tables,
            "count": len(tables),
        }
    except Exception as e:
        logger.error("Error listing tables: %s", e)
        return {"success": False, "error": str(e)}


@tools.tool(
    name="delete_table",
    description="Delete a structured table from an Excel worksheet",
    tags={"excel", "write", "table"},
)
async def delete_table(
    file_path: Annotated[str, Field(description="Absolute path to Excel file")],
    sheet_name: Annotated[str, Field(description="Worksheet name")],
    table_name: Annotated[str, Field(description="Name of table to delete")],
) -> dict[str, Any]:
    """Delete a structured table from an Excel worksheet."""
    try:
        backend = shared_cache.get(file_path)
        if backend is None:
            backend = create_backend(file_path)
            backend.open(file_path)

        ws = backend.get_sheet(sheet_name)

        if table_name not in ws.tables:
            return {
                "success": False,
                "error": f"Table '{table_name}' not found",
            }

        # Delete table
        del ws.tables[table_name]

        backend.save()
        shared_cache.put(file_path, backend)

        return {
            "success": True,
            "sheet_name": sheet_name,
            "deleted_table": table_name,
        }
    except Exception as e:
        logger.error("Error deleting table: %s", e)
        return {"success": False, "error": str(e)}


@tools.tool(
    name="add_table_row",
    description="Add a row to an Excel structured table",
    tags={"excel", "write", "table"},
)
async def add_table_row(
    file_path: Annotated[str, Field(description="Absolute path to Excel file")],
    sheet_name: Annotated[str, Field(description="Worksheet name")],
    table_name: Annotated[str, Field(description="Table name")],
    values: Annotated[list[Any], Field(description="Row values")],
) -> dict[str, Any]:
    """Add a row to an Excel structured table."""
    try:
        backend = shared_cache.get(file_path)
        if backend is None:
            backend = create_backend(file_path)
            backend.open(file_path)

        ws = backend.get_sheet(sheet_name)

        if table_name not in ws.tables:
            return {
                "success": False,
                "error": f"Table '{table_name}' not found",
            }

        table = ws.tables[table_name]
        table.add_row(values)

        backend.save()
        shared_cache.put(file_path, backend)

        return {
            "success": True,
            "sheet_name": sheet_name,
            "table_name": table_name,
            "values": values,
        }
    except Exception as e:
        logger.error("Error adding table row: %s", e)
        return {"success": False, "error": str(e)}


def _col_letter_to_index(cell_ref: str) -> int:
    """Convert column letter to 1-based index."""
    from openpyxl.utils import column_index_from_string

    col_letter = "".join(c for c in cell_ref if c.isalpha())
    return int(column_index_from_string(col_letter))
