"""Rate limiting for MCP Excel Server."""

import logging
import time
from collections import defaultdict
from dataclasses import dataclass
from threading import Lock

logger = logging.getLogger(__name__)


@dataclass
class RateLimitConfig:
    """Rate limit configuration."""

    max_requests: int = 100
    window_seconds: int = 60
    max_macro_calls: int = 10
    macro_window_seconds: int = 60


class RateLimiter:
    """Simple in-memory rate limiter."""

    def __init__(self, config: RateLimitConfig | None = None):
        """Initialize rate limiter.

        Args:
            config: Rate limit configuration
        """
        self._config = config or RateLimitConfig()
        self._requests: dict[str, list[float]] = defaultdict(list)
        self._macro_calls: dict[str, list[float]] = defaultdict(list)
        self._lock = Lock()

    def check_rate_limit(self, client_id: str = "default") -> bool:
        """Check if a request is allowed.

        Args:
            client_id: Client identifier

        Returns:
            True if allowed, False if rate limited
        """
        with self._lock:
            now = time.time()
            cutoff = now - self._config.window_seconds

            # Clean old requests
            self._requests[client_id] = [
                t for t in self._requests[client_id] if t > cutoff
            ]

            # Check limit
            if len(self._requests[client_id]) >= self._config.max_requests:
                logger.warning(
                    "Rate limit exceeded for client %s: %d requests in %ds",
                    client_id,
                    len(self._requests[client_id]),
                    self._config.window_seconds,
                )
                return False

            # Add request
            self._requests[client_id].append(now)
            return True

    def check_macro_limit(self, client_id: str = "default") -> bool:
        """Check if a macro call is allowed.

        Args:
            client_id: Client identifier

        Returns:
            True if allowed, False if rate limited
        """
        with self._lock:
            now = time.time()
            cutoff = now - self._config.macro_window_seconds

            # Clean old calls
            self._macro_calls[client_id] = [
                t for t in self._macro_calls[client_id] if t > cutoff
            ]

            # Check limit
            if len(self._macro_calls[client_id]) >= self._config.max_macro_calls:
                logger.warning(
                    "Macro rate limit exceeded for client %s: %d calls in %ds",
                    client_id,
                    len(self._macro_calls[client_id]),
                    self._config.macro_window_seconds,
                )
                return False

            # Add call
            self._macro_calls[client_id].append(now)
            return True

    def get_usage(self, client_id: str = "default") -> dict[str, int]:
        """Get current usage statistics.

        Args:
            client_id: Client identifier

        Returns:
            Dictionary with current usage counts
        """
        with self._lock:
            now = time.time()
            cutoff = now - self._config.window_seconds
            macro_cutoff = now - self._config.macro_window_seconds

            return {
                "requests": sum(
                    1 for t in self._requests[client_id] if t > cutoff
                ),
                "macro_calls": sum(
                    1 for t in self._macro_calls[client_id] if t > macro_cutoff
                ),
            }


# Global rate limiter instance
rate_limiter = RateLimiter()
