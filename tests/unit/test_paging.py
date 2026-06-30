"""Unit tests for pagination utilities."""

import pytest

from mcp_excel.utils.paging import PagingService, PaginationResult


class TestPagingService:
    """Tests for PagingService."""

    def test_calculate_page(self):
        """Test calculating page boundaries."""
        service = PagingService(cells_limit=40)
        result = service.calculate_page("A1:J10", page=1)
        assert result.page == 1
        assert result.total_rows == 10

    def test_calculate_page_second(self):
        """Test calculating second page."""
        service = PagingService(cells_limit=20)
        result = service.calculate_page("A1:J10", page=2)
        assert result.page == 2
        assert result.has_previous is True

    def test_calculate_page_last(self):
        """Test calculating last page."""
        service = PagingService(cells_limit=20)
        result = service.calculate_page("A1:J10", page=5)
        assert result.has_next is False

    def test_calculate_page_beyond_total(self):
        """Test calculating page beyond total pages."""
        service = PagingService(cells_limit=40)
        result = service.calculate_page("A1:J10", page=100)
        assert result.page <= result.total_pages

    def test_custom_cells_limit(self):
        """Test custom cells limit."""
        service = PagingService(cells_limit=100)
        result = service.calculate_page("A1:J10", page=1)
        assert result.total_pages >= 1

    def test_invalid_range_no_colon(self):
        """Test invalid range format without colon."""
        service = PagingService(cells_limit=40)
        with pytest.raises((ValueError, IndexError)):
            service.calculate_page("A1", page=1)

    def test_invalid_range_letters_only(self):
        """Test invalid range with only letters."""
        service = PagingService(cells_limit=40)
        with pytest.raises((ValueError, IndexError)):
            service.calculate_page("ABC:XYZ", page=1)


class TestPaginationResult:
    """Tests for PaginationResult."""

    def test_pagination_result_creation(self):
        """Test creating a PaginationResult."""
        result = PaginationResult(
            range_str="A1:J10",
            start_row=1,
            end_row=10,
            total_rows=100,
            page=1,
            total_pages=10,
            has_next=True,
            has_previous=False,
        )
        assert result.range_str == "A1:J10"
        assert result.page == 1
        assert result.total_rows == 100

    def test_pagination_result_properties(self):
        """Test PaginationResult properties."""
        result = PaginationResult(
            range_str="A1:J10",
            start_row=1,
            end_row=10,
            total_rows=100,
            page=2,
            total_pages=10,
            has_next=True,
            has_previous=True,
        )
        assert result.has_next is True
        assert result.has_previous is True

    def test_pagination_result_first_page(self):
        """Test PaginationResult on first page."""
        result = PaginationResult(
            range_str="A1:J10",
            start_row=1,
            end_row=10,
            total_rows=100,
            page=1,
            total_pages=10,
            has_next=True,
            has_previous=False,
        )
        assert result.has_previous is False
        assert result.has_next is True

    def test_pagination_result_last_page(self):
        """Test PaginationResult on last page."""
        result = PaginationResult(
            range_str="A91:J100",
            start_row=91,
            end_row=100,
            total_rows=100,
            page=10,
            total_pages=10,
            has_next=False,
            has_previous=True,
        )
        assert result.has_previous is True
        assert result.has_next is False
