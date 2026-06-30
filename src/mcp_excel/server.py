"""MCP Excel Server - Main entry point."""

import sys
import logging
from fastmcp import FastMCP

from .config import settings
from .tools import read, write, formulas, inspect

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
    instructions="MCP server for Excel file manipulation - read, write, formulas, charts, and tables",
    version="0.1.0",
    mask_error_details=settings.mask_errors,
)

# Register tools from modules
mcp.mount(read.tools)
mcp.mount(write.tools)
mcp.mount(formulas.tools)
mcp.mount(inspect.tools)

logger.info("MCP Excel Server initialized")


def main():
    """Run the MCP server."""
    logger.info("Starting MCP Excel Server (transport=%s)", settings.transport)
    
    if settings.transport == "http":
        mcp.run(transport="http", host="0.0.0.0", port=8000)
    else:
        mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
