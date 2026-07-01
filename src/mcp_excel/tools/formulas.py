"""MCP tools for Excel formulas."""

import logging
from typing import Annotated, Any

from fastmcp import FastMCP
from pydantic import Field

from ..backends.factory import create_backend
from ..utils.cache import shared_cache

logger = logging.getLogger(__name__)

# Create tools router
tools = FastMCP("Excel Formula Tools", mask_error_details=True)


@tools.tool(
    name="read_formula",
    description="Read the formula from a cell (returns formula text, not calculated value)",
    tags={"excel", "read", "formula"},
)
async def read_formula(
    file_path: Annotated[str, Field(description="Absolute path to Excel file")],
    sheet_name: Annotated[str, Field(description="Worksheet name")],
    cell: Annotated[str, Field(description="Cell reference (e.g., 'D1')")],
) -> dict[str, Any]:
    """Read the formula from a cell."""
    try:
        backend = shared_cache.get(file_path)
        if backend is None:
            backend = create_backend(file_path)
            backend.open(file_path)
            shared_cache.put(file_path, backend)

        ws = backend.get_sheet(sheet_name)
        cell_obj = ws[cell]

        # Check if cell contains a formula
        if cell_obj.data_type == "f":
            formula = cell_obj.value
            return {
                "success": True,
                "cell": cell,
                "formula": formula,
                "is_formula": True,
                "sheet_name": sheet_name,
            }
        else:
            return {
                "success": True,
                "cell": cell,
                "value": cell_obj.value,
                "is_formula": False,
                "sheet_name": sheet_name,
            }
    except Exception as e:
        logger.error("Error reading formula: %s", e)
        return {"success": False, "error": str(e)}


# Common formula templates
FORMULA_TEMPLATES = {
    "sum": "=SUM({range})",
    "average": "=AVERAGE({range})",
    "count": "=COUNT({range})",
    "counta": "=COUNTA({range})",
    "min": "=MIN({range})",
    "max": "=MAX({range})",
    "median": "=MEDIAN({range})",
    "stdev": "=STDEV({range})",
    "var": "=VAR({range})",
    "if": '=IF({condition},"{true_value}","{false_value}")',
    "sumif": "=SUMIF({range},{criteria},{sum_range})",
    "countif": "=COUNTIF({range},{criteria})",
    "vlookup": "=VLOOKUP({lookup_value},{table_range},{col_index},FALSE)",
    "concatenate": "=CONCATENATE({values})",
    "left": "=LEFT({text},{num_chars})",
    "right": "=RIGHT({text},{num_chars})",
    "len": "=LEN({text})",
    "upper": "=UPPER({text})",
    "lower": "=LOWER({text})",
    "today": "=TODAY()",
    "now": "=NOW()",
}


@tools.tool(
    name="get_formula_templates",
    description="Get common Excel formula templates",
    tags={"excel", "formula", "help"},
)
async def get_formula_templates() -> dict[str, Any]:
    """Get common Excel formula templates."""
    return {
        "success": True,
        "templates": FORMULA_TEMPLATES,
        "usage": (
            "Use these templates as a starting point. "
            "Replace {range}, {criteria}, etc. with actual values."
        ),
    }
