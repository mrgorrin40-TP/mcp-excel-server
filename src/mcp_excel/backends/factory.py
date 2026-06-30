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

    # Add more backends here as needed
    # elif suffix == ".xls":
    #     return XlrdBackend()

    raise ValueError(f"Unsupported file format: {suffix}")


def get_available_backends() -> list[str]:
    """Get list of available backend formats."""
    return [".xlsx"]
