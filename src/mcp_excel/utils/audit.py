"""Audit logging for MCP Excel Server."""

import json
import logging
import time
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)


@dataclass
class AuditEvent:
    """Represents an audit log event."""

    timestamp: float = field(default_factory=time.time)
    event_type: str = ""
    tool_name: str = ""
    file_path: str = ""
    user: str = ""
    success: bool = True
    error: str = ""
    details: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        """Convert to dictionary."""
        return asdict(self)

    def to_json(self) -> str:
        """Convert to JSON string."""
        return json.dumps(self.to_dict())


class AuditLogger:
    """Audit logger for tracking tool usage and security events."""

    def __init__(self, log_file: str | None = None):
        """Initialize audit logger.

        Args:
            log_file: Optional file path for audit log output
        """
        self._log_file = log_file
        self._events: list[AuditEvent] = []

    def log_tool_call(
        self,
        tool_name: str,
        file_path: str,
        success: bool = True,
        error: str = "",
        details: dict[str, Any] | None = None,
    ) -> None:
        """Log a tool call event.

        Args:
            tool_name: Name of the tool called
            file_path: Path to the file being accessed
            success: Whether the call succeeded
            error: Error message if failed
            details: Additional details about the call
        """
        event = AuditEvent(
            event_type="tool_call",
            tool_name=tool_name,
            file_path=file_path,
            success=success,
            error=error,
            details=details or {},
        )

        self._events.append(event)

        # Log to standard logger
        log_level = logging.INFO if success else logging.WARNING
        logger.log(
            log_level,
            "AUDIT: tool=%s file=%s success=%s",
            tool_name,
            file_path,
            success,
        )

        # Write to file if configured
        if self._log_file:
            self._write_to_file(event)

    def log_security_event(
        self,
        event_type: str,
        details: dict[str, Any] | None = None,
    ) -> None:
        """Log a security event.

        Args:
            event_type: Type of security event (e.g., "path_traversal", "rate_limit")
            details: Additional details about the event
        """
        event = AuditEvent(
            event_type=f"security:{event_type}",
            success=False,
            details=details or {},
        )

        self._events.append(event)

        logger.warning("SECURITY: %s %s", event_type, details or {})

        if self._log_file:
            self._write_to_file(event)

    def log_vba_operation(
        self,
        operation: str,
        file_path: str,
        module_name: str = "",
        success: bool = True,
        error: str = "",
    ) -> None:
        """Log a VBA operation.

        Args:
            operation: VBA operation type (e.g., "run_macro", "add_module")
            file_path: Path to the workbook
            module_name: Name of the VBA module
            success: Whether the operation succeeded
            error: Error message if failed
        """
        event = AuditEvent(
            event_type="vba_operation",
            tool_name=operation,
            file_path=file_path,
            success=success,
            error=error,
            details={"module_name": module_name} if module_name else {},
        )

        self._events.append(event)

        log_level = logging.INFO if success else logging.WARNING
        logger.log(
            log_level,
            "AUDIT VBA: operation=%s file=%s module=%s success=%s",
            operation,
            file_path,
            module_name,
            success,
        )

        if self._log_file:
            self._write_to_file(event)

    def get_recent_events(self, count: int = 100) -> list[dict[str, Any]]:
        """Get recent audit events.

        Args:
            count: Number of recent events to return

        Returns:
            List of recent audit events as dictionaries
        """
        return [e.to_dict() for e in self._events[-count:]]

    def _write_to_file(self, event: AuditEvent) -> None:
        """Write event to log file."""
        try:
            if not self._log_file:
                return
            log_path = Path(self._log_file)
            log_path.parent.mkdir(parents=True, exist_ok=True)
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(event.to_json() + "\n")
        except Exception as e:
            logger.error("Failed to write audit log: %s", e)


# Global audit logger instance
audit_logger = AuditLogger()
