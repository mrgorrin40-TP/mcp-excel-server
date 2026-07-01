"""Path validation utilities for MCP Excel Server."""

import logging
from pathlib import Path

from ..config import settings

logger = logging.getLogger(__name__)


class PathValidationError(Exception):
    """Raised when a file path fails validation."""

    pass


def validate_file_path(file_path: str) -> str:
    """Validate and normalize a file path.

    Checks:
    1. Path is absolute (no relative paths)
    2. No path traversal attempts (..)
    3. File extension is allowed (.xlsx, .xlsm, .xlsb, .xlam)
    4. Path is within allowed directories (if configured)

    Args:
        file_path: The file path to validate

    Returns:
        Normalized absolute path as string

    Raises:
        PathValidationError: If path fails any validation check
    """
    # Convert to Path object and resolve to absolute
    try:
        path = Path(file_path)
    except Exception as e:
        raise PathValidationError(f"Invalid path: {e}") from e

    # Check for path traversal attempts
    path_str = str(path)
    if ".." in path_str:
        raise PathValidationError(
            f"Path traversal not allowed: {path_str}"
        )

    # Make path absolute if it's relative
    if not path.is_absolute():
        path = Path.cwd() / path
        logger.warning("Relative path converted to absolute: %s", path)

    # Resolve the path (removes . and .. components)
    try:
        resolved = path.resolve()
    except Exception as e:
        raise PathValidationError(f"Cannot resolve path: {e}") from e

    # Check allowed directories if configured
    allowed_dirs = settings.allowed_directories
    if allowed_dirs:
        is_allowed = False
        for allowed_dir in allowed_dirs:
            allowed_path = Path(allowed_dir).resolve()
            try:
                resolved.relative_to(allowed_path)
                is_allowed = True
                break
            except ValueError:
                continue

        if not is_allowed:
            raise PathValidationError(
                f"Access denied: {resolved} is not within allowed directories: {allowed_dirs}"
            )

    return str(resolved)


def get_allowed_extensions() -> list[str]:
    """Get list of allowed file extensions.

    Returns:
        List of allowed extensions (e.g., [".xlsx", ".xlsm"])
    """
    extensions = [".xlsx"]

    # Add VBA formats if enabled
    if settings.vba_enabled:
        extensions.extend([".xlsm", ".xlsb", ".xlam"])

    return extensions


def is_valid_extension(file_path: str) -> bool:
    """Check if file has a valid extension.

    Args:
        file_path: The file path to check

    Returns:
        True if extension is valid, False otherwise
    """
    path = Path(file_path)
    return path.suffix.lower() in get_allowed_extensions()
