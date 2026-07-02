"""Tests for security features: validation, audit, rate limiter."""

from unittest.mock import patch

import pytest

from mcp_excel.utils.audit import AuditLogger
from mcp_excel.utils.rate_limiter import RateLimitConfig, RateLimiter
from mcp_excel.utils.validation import (
    PathValidationError,
    get_allowed_extensions,
    is_valid_extension,
    validate_file_path,
)
from mcp_excel.utils.vba_validator import check_vba_safety, validate_vba_code


class TestPathValidation:
    """Tests for path validation utilities."""

    def test_validate_absolute_path(self, tmp_path):
        """Test validation of absolute path."""
        test_file = tmp_path / "test.xlsx"
        test_file.touch()

        result = validate_file_path(str(test_file))
        assert result == str(test_file.resolve())

    def test_validate_rejects_path_traversal(self):
        """Test rejection of path traversal attempts."""
        with pytest.raises(PathValidationError, match="Path traversal not allowed"):
            validate_file_path("../../../etc/passwd")

    def test_validate_rejects_path_traversal_in_middle(self):
        """Test rejection of path traversal in middle of path."""
        with pytest.raises(PathValidationError, match="Path traversal not allowed"):
            validate_file_path("/data/../../../etc/passwd")

    def test_validate_rejects_unallowed_directory(self):
        """Test rejection when allowed_directories is configured."""
        with patch("mcp_excel.utils.validation.settings") as mock_settings:
            mock_settings.allowed_directories = ["/allowed/path"]

            with pytest.raises(PathValidationError, match="Access denied"):
                validate_file_path("/not/allowed/file.xlsx")

    def test_validate_allows_within_allowed_directory(self, tmp_path):
        """Test path within allowed directory is accepted."""
        with patch("mcp_excel.utils.validation.settings") as mock_settings:
            mock_settings.allowed_directories = [str(tmp_path)]

            test_file = tmp_path / "test.xlsx"
            test_file.touch()

            result = validate_file_path(str(test_file))
            assert result == str(test_file.resolve())

    def test_get_allowed_extensions_vba_disabled(self):
        """Test allowed extensions when VBA is disabled."""
        with patch("mcp_excel.utils.validation.settings") as mock_settings:
            mock_settings.vba_enabled = False
            extensions = get_allowed_extensions()
            assert ".xlsx" in extensions
            assert ".xlsm" not in extensions

    def test_get_allowed_extensions_vba_enabled(self):
        """Test allowed extensions when VBA is enabled."""
        with patch("mcp_excel.utils.validation.settings") as mock_settings:
            mock_settings.vba_enabled = True
            extensions = get_allowed_extensions()
            assert ".xlsx" in extensions
            assert ".xlsm" in extensions
            assert ".xlsb" in extensions
            assert ".xlam" in extensions

    def test_is_valid_extension(self):
        """Test extension validation."""
        assert is_valid_extension("test.xlsx") is True
        assert is_valid_extension("test.txt") is False
        assert is_valid_extension("test.csv") is False


class TestAuditLogger:
    """Tests for audit logging."""

    def test_log_tool_call(self):
        """Test logging a tool call."""
        logger = AuditLogger()
        logger.log_tool_call("read_cell", "/data/test.xlsx", success=True)

        events = logger.get_recent_events()
        assert len(events) == 1
        assert events[0]["event_type"] == "tool_call"
        assert events[0]["tool_name"] == "read_cell"
        assert events[0]["success"] is True

    def test_log_tool_call_with_error(self):
        """Test logging a failed tool call."""
        logger = AuditLogger()
        logger.log_tool_call(
            "write_cells", "/data/test.xlsx", success=False, error="File not found"
        )

        events = logger.get_recent_events()
        assert len(events) == 1
        assert events[0]["success"] is False
        assert events[0]["error"] == "File not found"

    def test_log_security_event(self):
        """Test logging a security event."""
        logger = AuditLogger()
        logger.log_security_event(
            "path_traversal", {"path": "../../../etc/passwd"}
        )

        events = logger.get_recent_events()
        assert len(events) == 1
        assert events[0]["event_type"] == "security:path_traversal"

    def test_log_vba_operation(self):
        """Test logging a VBA operation."""
        logger = AuditLogger()
        logger.log_vba_operation(
            "run_macro", "/data/test.xlsm", module_name="Module1"
        )

        events = logger.get_recent_events()
        assert len(events) == 1
        assert events[0]["event_type"] == "vba_operation"
        assert events[0]["details"]["module_name"] == "Module1"

    def test_get_recent_events_limit(self):
        """Test limiting recent events."""
        logger = AuditLogger()
        for i in range(10):
            logger.log_tool_call("tool", f"/data/{i}.xlsx")

        events = logger.get_recent_events(count=5)
        assert len(events) == 5

    def test_log_to_file(self, tmp_path):
        """Test logging to file."""
        log_file = tmp_path / "audit.log"
        logger = AuditLogger(log_file=str(log_file))
        logger.log_tool_call("read_cell", "/data/test.xlsx")

        assert log_file.exists()
        content = log_file.read_text()
        assert "tool_call" in content
        assert "read_cell" in content


class TestRateLimiter:
    """Tests for rate limiting."""

    def test_allows_within_limit(self):
        """Test requests within limit are allowed."""
        config = RateLimitConfig(max_requests=10, window_seconds=1)
        limiter = RateLimiter(config)

        for _ in range(10):
            assert limiter.check_rate_limit("client1") is True

    def test_rejects_over_limit(self):
        """Test requests over limit are rejected."""
        config = RateLimitConfig(max_requests=5, window_seconds=1)
        limiter = RateLimiter(config)

        for _ in range(5):
            limiter.check_rate_limit("client1")

        assert limiter.check_rate_limit("client1") is False

    def test_separate_clients(self):
        """Test rate limits are per-client."""
        config = RateLimitConfig(max_requests=2, window_seconds=1)
        limiter = RateLimiter(config)

        limiter.check_rate_limit("client1")
        limiter.check_rate_limit("client1")
        assert limiter.check_rate_limit("client1") is False

        assert limiter.check_rate_limit("client2") is True

    def test_macro_rate_limit(self):
        """Test macro-specific rate limiting."""
        config = RateLimitConfig(max_macro_calls=2, macro_window_seconds=1)
        limiter = RateLimiter(config)

        limiter.check_macro_limit("client1")
        limiter.check_macro_limit("client1")
        assert limiter.check_macro_limit("client1") is False

    def test_get_usage(self):
        """Test usage statistics."""
        config = RateLimitConfig(max_requests=10, window_seconds=1)
        limiter = RateLimiter(config)

        limiter.check_rate_limit("client1")
        limiter.check_rate_limit("client1")
        limiter.check_macro_limit("client1")

        usage = limiter.get_usage("client1")
        assert usage["requests"] == 2
        assert usage["macro_calls"] == 1


class TestVBASafetyValidation:
    """Tests for VBA safety validation."""

    def test_safe_code_no_issues(self):
        """Test safe VBA code has no safety issues."""
        code = """
Sub SafeMacro()
    Range("A1").Value = "Hello"
End Sub
"""
        issues = check_vba_safety(code)
        assert len(issues) == 0

    def test_shell_command_detected(self):
        """Test Shell command detection."""
        code = 'Sub RunShell()\n    Shell "cmd.exe /c dir"\nEnd Sub'
        issues = check_vba_safety(code)
        assert len(issues) == 1
        assert "Shell command execution" in issues[0]

    def test_file_system_object_detected(self):
        """Test FileSystemObject detection."""
        code = """
Sub ReadFile()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
End Sub
"""
        issues = check_vba_safety(code)
        assert any("FileSystemObject" in i or "ActiveX" in i for i in issues)

    def test_kill_command_detected(self):
        """Test Kill command detection."""
        code = 'Sub DeleteFile()\n    Kill "C:\\temp\\*.*"\nEnd Sub'
        issues = check_vba_safety(code)
        assert any("File deletion" in i for i in issues)

    def test_sendkeys_detected(self):
        """Test SendKeys detection."""
        code = 'Sub SendKeysExample()\n    SendKeys "{ENTER}"\nEnd Sub'
        issues = check_vba_safety(code)
        assert any("Keyboard input" in i for i in issues)

    def test_validate_vba_code_syntax(self):
        """Test VBA code syntax validation."""
        code = """
Sub Test()
    Dim x As Integer
    x = 10
End Sub
"""
        result = validate_vba_code(code)
        assert result.valid is True
        assert "Test" in result.subs

    def test_validate_vba_code_mismatched_sub(self):
        """Test detection of mismatched Sub/End Sub."""
        code = """
Sub Test()
    Dim x As Integer
End Sub
Sub Test2()
    Dim y As Integer
End Sub
End Sub
"""
        result = validate_vba_code(code)
        assert result.valid is False
        assert any("Mismatched Sub/End Sub" in e for e in result.errors)
