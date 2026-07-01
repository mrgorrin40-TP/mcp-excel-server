"""Backend helper for MCP tools."""

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from ..backends.base import ExcelBackend


def get_backend(file_path: str) -> "ExcelBackend":
    """Get or create a backend for the given file path.

    This helper centralizes the cache pattern used across all tools:
    1. Check if backend exists in cache
    2. If not, create and open a new backend
    3. Store in cache for reuse

    Args:
        file_path: Absolute path to Excel file

    Returns:
        ExcelBackend instance (opened and ready to use)

    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If file format is not supported
    """
    from ..backends.factory import create_backend
    from ..utils.cache import shared_cache

    backend = shared_cache.get(file_path)
    if backend is None:
        backend = create_backend(file_path)
        backend.open(file_path)
        shared_cache.put(file_path, backend)
    return backend
