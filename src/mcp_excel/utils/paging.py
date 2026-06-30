"""Pagination service for large Excel datasets."""

import logging
import re
from dataclasses import dataclass

logger = logging.getLogger(__name__)


@dataclass
class PaginationResult:
    """Result of pagination calculation."""

    range_str: str
    start_row: int
    end_row: int
    total_rows: int
    page: int
    total_pages: int
    has_next: bool
    has_previous: bool


class PagingService:
    """Service for paginating Excel data."""

    def __init__(self, cells_limit: int = 4000):
        """Initialize the paging service.

        Args:
            cells_limit: Maximum cells per page
        """
        self._cells_limit = cells_limit
        logger.info("PagingService initialized: cells_limit=%d", cells_limit)

    def calculate_page(
        self,
        used_range: str,
        page: int = 1,
        total_columns: int | None = None,
    ) -> PaginationResult:
        """Calculate pagination for a given range.

        Args:
            used_range: The used range of the worksheet (e.g., "A1:Z1000")
            page: Page number (1-indexed)
            total_columns: Number of columns (calculated from range if not provided)

        Returns:
            PaginationResult with range and pagination info
        """
        # Parse range
        start_match = re.match(r"([A-Z]+)(\d+)", used_range.split(":")[0])
        end_match = re.match(r"([A-Z]+)(\d+)", used_range.split(":")[1])

        if not start_match or not end_match:
            raise ValueError(f"Invalid range format: {used_range}")

        start_col = self._column_to_number(start_match.group(1))
        start_row = int(start_match.group(2))
        end_col = self._column_to_number(end_match.group(1))
        end_row = int(end_match.group(2))

        total_rows = end_row - start_row + 1
        if total_columns is None:
            total_columns = end_col - start_col + 1

        # Calculate rows per page
        rows_per_page = max(1, self._cells_limit // total_columns)
        total_pages = (total_rows + rows_per_page - 1) // rows_per_page

        # Clamp page number
        page = max(1, min(page, total_pages))

        # Calculate range for this page
        page_start_row = start_row + (page - 1) * rows_per_page
        page_end_row = min(page_start_row + rows_per_page - 1, end_row)

        start_col_letter = self._number_to_column(start_col)
        end_col_letter = self._number_to_column(end_col)

        range_str = f"{start_col_letter}{page_start_row}:{end_col_letter}{page_end_row}"

        return PaginationResult(
            range_str=range_str,
            start_row=page_start_row,
            end_row=page_end_row,
            total_rows=total_rows,
            page=page,
            total_pages=total_pages,
            has_next=page < total_pages,
            has_previous=page > 1,
        )

    def get_all_ranges(
        self,
        used_range: str,
        total_columns: int | None = None,
    ) -> list[str]:
        """Get all page ranges for a worksheet.

        Args:
            used_range: The used range of the worksheet
            total_columns: Number of columns

        Returns:
            List of range strings for each page
        """
        result = self.calculate_page(used_range, page=1, total_columns=total_columns)

        ranges = []
        for page in range(1, result.total_pages + 1):
            page_result = self.calculate_page(used_range, page=page, total_columns=total_columns)
            ranges.append(page_result.range_str)

        return ranges

    def _column_to_number(self, column: str) -> int:
        """Convert column letter to number (A=1, B=2, ..., Z=26, AA=27)."""
        result = 0
        for char in column:
            result = result * 26 + (ord(char.upper()) - ord("A") + 1)
        return result

    def _number_to_column(self, number: int) -> str:
        """Convert number to column letter (1=A, 2=B, ..., 26=Z, 27=AA)."""
        result = ""
        while number > 0:
            number, remainder = divmod(number - 1, 26)
            result = chr(65 + remainder) + result
        return result
