"""Header detection for Excel worksheets."""

import logging
from dataclasses import dataclass
from typing import Any

logger = logging.getLogger(__name__)


@dataclass
class HeaderDetectionResult:
    """Result of header detection."""

    header_row: int
    headers: list[str]
    confidence: float
    sample_data: list[list[Any]]


class HeaderDetector:
    """Detect header rows in Excel data."""

    def __init__(self, max_scan_rows: int = 20):
        """Initialize the header detector.

        Args:
            max_scan_rows: Maximum rows to scan for headers
        """
        self._max_scan_rows = max_scan_rows

    def detect(self, data: list[list[Any]]) -> HeaderDetectionResult:
        """Detect the header row in data.

        Args:
            data: 2D array of cell values

        Returns:
            HeaderDetectionResult with detected headers
        """
        if not data:
            return HeaderDetectionResult(
                header_row=0,
                headers=[],
                confidence=0.0,
                sample_data=[],
            )

        # Scan first N rows
        scan_rows = min(len(data), self._max_scan_rows)

        best_row = 0
        best_score = 0.0
        best_headers = []

        for row_idx in range(scan_rows):
            row = data[row_idx]
            score, headers = self._score_row(row)

            if score > best_score:
                best_score = score
                best_row = row_idx
                best_headers = headers

        # Get sample data after headers
        sample_start = best_row + 1
        sample_end = min(sample_start + 5, len(data))
        sample_data = data[sample_start:sample_end] if sample_start < len(data) else []

        return HeaderDetectionResult(
            header_row=best_row,
            headers=best_headers,
            confidence=best_score,
            sample_data=sample_data,
        )

    def _score_row(self, row: list[Any]) -> tuple[float, list[str]]:
        """Score a row as potential header.

        Returns:
            Tuple of (score, headers)
        """
        if not row:
            return 0.0, []

        # Convert to strings and filter None
        headers = [str(cell) if cell is not None else "" for cell in row]

        # Filter out empty headers
        non_empty = [h for h in headers if h.strip()]

        if not non_empty:
            return 0.0, []

        score = 0.0

        # Score based on string density
        string_count = sum(1 for h in non_empty if not self._is_numeric(h))
        score += (string_count / len(non_empty)) * 0.3

        # Score based on uniqueness
        unique_count = len(set(non_empty))
        score += (unique_count / len(non_empty)) * 0.3

        # Score based on no None/empty values
        completeness = len(non_empty) / len(row)
        score += completeness * 0.2

        # Score based on typical header patterns
        for header in non_empty:
            header_lower = header.lower()

            # Common header keywords
            if any(kw in header_lower for kw in ["name", "date", "id", "type", "status", "value"]):
                score += 0.05

            # All caps (often headers)
            if header.isupper() and len(header) > 1:
                score += 0.05

        # Penalty for numbers (less likely to be headers)
        numeric_count = sum(1 for h in non_empty if self._is_numeric(h))
        score -= (numeric_count / len(non_empty)) * 0.2

        # Ensure score is between 0 and 1
        score = max(0.0, min(1.0, score))

        return score, headers

    def _is_numeric(self, value: str) -> bool:
        """Check if a string represents a number."""
        try:
            float(value)
            return True
        except ValueError:
            return False
