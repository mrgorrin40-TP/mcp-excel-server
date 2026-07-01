"""MCP Excel Server - Main entry point."""

import logging
import sys
from typing import Any

from fastmcp import FastMCP

from .config import settings
from .prompts import templates as prompt_templates
from .resources import excel as excel_resources
from .tools import charts, formulas, inspect, read, tables, write

# Configure logging to stderr (keep stdout clean for MCP)
logging.basicConfig(
    level=getattr(logging, settings.log_level),
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    stream=sys.stderr,
)
logger = logging.getLogger(__name__)

# Create MCP server instance
mcp = FastMCP(
    name="mcp-excel-server",
    instructions=(
        "MCP server for Excel file manipulation - "
        "read, write, formulas, charts, tables, and analysis"
    ),
    version="0.2.0",
    mask_error_details=settings.mask_errors,
)

# Register tools from modules
mcp.mount(read.tools)
mcp.mount(write.tools)
mcp.mount(formulas.tools)
mcp.mount(inspect.tools)
mcp.mount(charts.tools)
mcp.mount(tables.tools)


# Register MCP Resources
@mcp.resource("excel:///{file_path}/metadata")
def get_workbook_metadata(file_path: str) -> dict[str, Any]:
    """Get workbook metadata."""
    return excel_resources.get_workbook_metadata(file_path)


@mcp.resource("excel:///{file_path}/{sheet_name}/metadata")
def get_sheet_metadata(file_path: str, sheet_name: str) -> dict[str, Any]:
    """Get sheet metadata."""
    return excel_resources.get_sheet_metadata(file_path, sheet_name)


# Register MCP Prompts
@mcp.prompt()
def explore_workbook(file_path: str) -> str:
    """Explore and summarize an Excel workbook."""
    result = prompt_templates.get_prompt("explore_workbook", file_path=file_path)
    return str(result.get("text", ""))


@mcp.prompt()
def analyze_data(file_path: str, sheet_name: str) -> str:
    """Perform statistical analysis on worksheet data."""
    result = prompt_templates.get_prompt(
        "analyze_data", file_path=file_path, sheet_name=sheet_name
    )
    return str(result.get("text", ""))


@mcp.prompt()
def compare_sheets(file_path: str, sheet1: str, sheet2: str) -> str:
    """Compare data between two worksheets."""
    result = prompt_templates.get_prompt(
        "compare_sheets", file_path=file_path, sheet1=sheet1, sheet2=sheet2
    )
    return str(result.get("text", ""))


@mcp.prompt()
def create_report(file_path: str, sheet_name: str) -> str:
    """Create a summary report from raw data."""
    result = prompt_templates.get_prompt(
        "create_report", file_path=file_path, sheet_name=sheet_name
    )
    return str(result.get("text", ""))


@mcp.prompt()
def data_quality_check(file_path: str, sheet_name: str) -> str:
    """Check data quality and identify issues."""
    result = prompt_templates.get_prompt(
        "data_quality_check", file_path=file_path, sheet_name=sheet_name
    )
    return str(result.get("text", ""))


logger.info("MCP Excel Server initialized (v0.2.0)")


def main() -> None:
    """Run the MCP server."""
    logger.info("Starting MCP Excel Server (transport=%s)", settings.transport)

    if settings.transport == "http":
        mcp.run(transport="http", host="0.0.0.0", port=8000)
    else:
        mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
