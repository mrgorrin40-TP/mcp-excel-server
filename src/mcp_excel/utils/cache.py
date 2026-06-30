"""LRU Cache for Excel workbooks."""

import logging
import time
from collections import OrderedDict
from pathlib import Path
from typing import Any

from ..backends.base import ExcelBackend
from ..backends.factory import create_backend

logger = logging.getLogger(__name__)


class WorkbookCache:
    """LRU cache for Excel workbooks with memory management."""

    def __init__(
        self,
        max_size: int = 5,
        max_memory_mb: int = 1024,
        idle_timeout_seconds: int = 600,
    ):
        """Initialize the cache.
        
        Args:
            max_size: Maximum number of workbooks to cache
            max_memory_mb: Maximum memory usage in MB
            idle_timeout_seconds: Timeout for idle workbooks
        """
        self._cache: OrderedDict[str, tuple[ExcelBackend, float]] = OrderedDict()
        self._max_size = max_size
        self._max_memory_mb = max_memory_mb
        self._idle_timeout = idle_timeout_seconds
        self._last_access = time.time()
        
        logger.info(
            "WorkbookCache initialized: max_size=%d, max_memory=%dMB, timeout=%ds",
            max_size,
            max_memory_mb,
            idle_timeout_seconds,
        )

    def get(self, file_path: str | Path) -> ExcelBackend | None:
        """Get a workbook from cache.
        
        Args:
            file_path: Path to the Excel file
            
        Returns:
            ExcelBackend if cached, None otherwise
        """
        self._cleanup_if_needed()
        
        cache_key = str(Path(file_path).resolve())
        
        if cache_key in self._cache:
            # Move to end (most recently used)
            backend, timestamp = self._cache.pop(cache_key)
            self._cache[cache_key] = (backend, time.time())
            self._last_access = time.time()
            
            logger.debug("Cache hit: %s", cache_key)
            return backend
        
        logger.debug("Cache miss: %s", cache_key)
        return None

    def put(self, file_path: str | Path, backend: ExcelBackend) -> None:
        """Add a workbook to cache.
        
        Args:
            file_path: Path to the Excel file
            backend: ExcelBackend instance
        """
        cache_key = str(Path(file_path).resolve())
        
        # Remove if already exists
        if cache_key in self._cache:
            self._cache.pop(cache_key)
        
        # Add to end
        self._cache[cache_key] = (backend, time.time())
        self._last_access = time.time()
        
        # Evict if over size limit
        while len(self._cache) > self._max_size:
            self._evict_oldest()
        
        logger.debug("Cached workbook: %s (total: %d)", cache_key, len(self._cache))

    def remove(self, file_path: str | Path) -> None:
        """Remove a workbook from cache.
        
        Args:
            file_path: Path to the Excel file
        """
        cache_key = str(Path(file_path).resolve())
        
        if cache_key in self._cache:
            backend, _ = self._cache.pop(cache_key)
            backend.close()
            logger.debug("Removed from cache: %s", cache_key)

    def clear(self) -> None:
        """Clear all cached workbooks."""
        for cache_key, (backend, _) in self._cache.items():
            backend.close()
        
        self._cache.clear()
        logger.info("Cache cleared")

    def _cleanup_if_needed(self) -> None:
        """Cleanup expired entries if needed."""
        current_time = time.time()
        
        # Check idle timeout
        if current_time - self._last_access > self._idle_timeout:
            logger.info("Idle timeout reached, clearing cache")
            self.clear()
            return
        
        # Check for expired entries
        expired_keys = []
        for cache_key, (backend, timestamp) in self._cache.items():
            if current_time - timestamp > self._idle_timeout:
                expired_keys.append(cache_key)
        
        for key in expired_keys:
            backend, _ = self._cache.pop(key)
            backend.close()
            logger.debug("Evicted expired entry: %s", key)

    def _evict_oldest(self) -> None:
        """Evict the oldest entry from cache."""
        if self._cache:
            cache_key, (backend, _) = self._cache.popitem(last=False)
            backend.close()
            logger.debug("Evicted oldest entry: %s", cache_key)

    @property
    def size(self) -> int:
        """Get current cache size."""
        return len(self._cache)

    @property
    def keys(self) -> list[str]:
        """Get list of cached file paths."""
        return list(self._cache.keys())
