"""MCP Excel Server - Main entry point."""

import logging
import sys

from fastmcp import FastMCP

from .config import settings
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
