"""VBA code validation utilities."""

import re
from dataclasses import dataclass, field


@dataclass
class VBAValidationResult:
    """Result of VBA code validation."""

    valid: bool = True
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    subs: list[str] = field(default_factory=list)
    functions: list[str] = field(default_factory=list)
    safety_issues: list[str] = field(default_factory=list)


# Potentially dangerous VBA patterns
# Note: VBA syntax uses spaces not parentheses for most commands
DANGEROUS_PATTERNS = [
    (r"(?i)\bShell\s+\"", "Shell command execution"),
    (r"(?i)\bShell\s*\(", "Shell command execution"),
    (r"(?i)\bCreateObject\s*\(", "ActiveX object creation"),
    (r"(?i)\bWScript\.Shell", "Windows Script Host access"),
    (r"(?i)\bFileSystemObject", "File system access"),
    (r"(?i)\bKill\s+\"", "File deletion"),
    (r"(?i)\bKill\s*\(", "File deletion"),
    (r"(?i)\bName\s+.*\s+As\s+", "File rename/move"),
    (r"(?i)\bMkDir\s+\"", "Directory creation"),
    (r"(?i)\bMkDir\s*\(", "Directory creation"),
    (r"(?i)\bRmDir\s+\"", "Directory deletion"),
    (r"(?i)\bRmDir\s*\(", "Directory deletion"),
    (r"(?i)\bSendKeys\s+\"", "Keyboard input simulation"),
    (r"(?i)\bSendKeys\s*\(", "Keyboard input simulation"),
    (r"(?i)\bRegRead\s+\"", "Registry read"),
    (r"(?i)\bRegRead\s*\(", "Registry read"),
    (r"(?i)\bRegWrite\s+\"", "Registry write"),
    (r"(?i)\bRegWrite\s*\(", "Registry write"),
    (r"(?i)\bURLDownloadToFile", "File download from internet"),
    (r"(?i)\bExec\s*\(", "Command execution"),
    (r"(?i)\bApplication\.Run\s*\(", "Dynamic macro execution"),
]


def validate_vba_code(code: str) -> VBAValidationResult:
    """Validate VBA code syntax (basic validation).

    Checks for:
    - Matching Sub/End Sub blocks
    - Matching Function/End Function blocks
    - Variable declarations without type
    - Potentially dangerous code patterns

    Args:
        code: VBA source code to validate

    Returns:
        VBAValidationResult with errors, warnings, and detected procedures
    """
    result = VBAValidationResult()
    lines = code.split("\n")

    for i, line in enumerate(lines, 1):
        stripped = line.strip()

        # Skip comments and empty lines
        if stripped.startswith("'") or not stripped:
            continue

        # Check for Sub/Function declarations
        if stripped.startswith("Sub ") or stripped.startswith("Public Sub "):
            name = stripped.split("Sub ")[1].split("(")[0].strip()
            result.subs.append(name)
        elif stripped.startswith("Function ") or stripped.startswith("Public Function "):
            name = stripped.split("Function ")[1].split("(")[0].strip()
            result.functions.append(name)

        # Check for variable declaration without type
        if "Dim " in stripped and " As " not in stripped and not stripped.startswith("'"):
            result.warnings.append(f"Line {i}: Variable declaration missing type")

    # Check for matching Sub/End Sub
    if result.subs:
        end_subs = sum(1 for line in lines if line.strip() == "End Sub")
        if end_subs != len(result.subs):
            result.errors.append(
                f"Mismatched Sub/End Sub: {len(result.subs)} Sub vs {end_subs} End Sub"
            )

    # Check for matching Function/End Function
    if result.functions:
        end_functions = sum(1 for line in lines if line.strip() == "End Function")
        if end_functions != len(result.functions):
            result.errors.append(
                f"Mismatched Function/End Function: {len(result.functions)} Function vs "
                f"{end_functions} End Function"
            )

    result.valid = len(result.errors) == 0
    return result


def check_vba_safety(code: str) -> list[str]:
    """Check VBA code for potentially dangerous patterns.

    Args:
        code: VBA source code to check

    Returns:
        List of safety issues found
    """
    issues = []

    for pattern, description in DANGEROUS_PATTERNS:
        matches = re.findall(pattern, code)
        if matches:
            issues.append(f"{description}: {len(matches)} occurrence(s) found")

    return issues
