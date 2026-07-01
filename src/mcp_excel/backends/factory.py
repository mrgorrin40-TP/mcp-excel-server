"""Factory for creating Excel backends."""

import logging
from pathlib import Path

from .base import ExcelBackend
from .openpyxl_backend import OpenpyxlBackend

logger = logging.getLogger(__name__)


def create_backend(file_path: str | Path) -> ExcelBackend:
    """Create an appropriate Excel backend for the given file.

    Args:
        file_path: Path to the Excel file

    Returns:
        ExcelBackend instance

    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If file format is not supported
    """
    path = Path(file_path)

    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")

    # Check file extension
    suffix = path.suffix.lower()

    if suffix == ".xlsx":
        logger.info("Using openpyxl backend for %s", path)
        return OpenpyxlBackend()

    # VBA-enabled workbooks
    if suffix in (".xlsm", ".xlsb", ".xlam"):
        from ..config import settings

        if not settings.vba_enabled:
            raise ValueError(
                "VBA support is disabled. Enable with MCP_EXCEL_VBA_ENABLED=true"
            )

        # Try to use xlwings if available
        try:
            import xlwings as xw  # noqa: F401

            from .xlwings_backend import XlwingsBackend

            logger.info("Using xlwings backend for %s", path)
            return XlwingsBackend()
        except ImportError as err:
            raise ValueError(
                "xlwings is required for VBA macro support. "
                "Install with: pip install 'mcp-excel-server[vba]'"
            ) from err

    raise ValueError(
        f"Unsupported file format: {suffix}. "
        f"Supported formats: .xlsx, .xlsm, .xlsb, .xlam"
    )


def get_available_backends() -> list[str]:
    """Get list of available backend formats."""
    backends = [".xlsx"]

    # Check if xlwings is available
    try:
        import xlwings as xw  # noqa: F401

        backends.extend([".xlsm", ".xlsb", ".xlam"])
    except ImportError:
        logger.info("xlwings not installed, VBA formats not available")

    return backends
