"""Unit tests for header detection utilities."""

import pytest

from mcp_excel.utils.headers import HeaderDetector, HeaderDetectionResult


class TestHeaderDetector:
    """Tests for HeaderDetector."""

    def test_detect_headers_first_row(self):
        """Test detecting headers in first row."""
        detector = HeaderDetector()
        data = [
            ["Name", "Age", "City"],
            ["Alice", 25, "New York"],
            ["Bob", 30, "London"],
        ]
        result = detector.detect(data)
        assert result.header_row == 0
        assert result.headers == ["Name", "Age", "City"]

    def test_detect_headers_no_headers(self):
        """Test when no clear headers are detected."""
        detector = HeaderDetector()
        data = [
            [1, 2, 3],
            [4, 5, 6],
            [7, 8, 9],
        ]
        result = detector.detect(data)
        assert result.confidence < 0.5

    def test_score_row_text_heavy(self):
        """Test scoring a row with mostly text."""
        detector = HeaderDetector()
        row = ["Name", "Age", "City", "Email"]
        score, headers = detector._score_row(row)
        assert score > 0.3

    def test_score_row_numeric_heavy(self):
        """Test scoring a row with mostly numbers."""
        detector = HeaderDetector()
        row = [1, 2, 3, 4, 5]
        score, headers = detector._score_row(row)
        assert score <= 0.3

    def test_score_row_mixed(self):
        """Test scoring a mixed row."""
        detector = HeaderDetector()
        row = ["Name", 123, "City", 456]
        score, headers = detector._score_row(row)
        assert score > 0

    def test_score_row_empty(self):
        """Test scoring an empty row."""
        detector = HeaderDetector()
        row = []
        score, headers = detector._score_row(row)
        assert score == 0.0
        assert headers == []

    def test_score_row_single_item(self):
        """Test scoring a row with single item."""
        detector = HeaderDetector()
        row = ["Name"]
        score, headers = detector._score_row(row)
        assert score > 0

    def test_detect_with_different_data_types(self):
        """Test detection with various data types."""
        detector = HeaderDetector()
        data = [
            ["ID", "Value", "Active"],
            [1, 100.5, True],
            [2, 200.5, False],
        ]
        result = detector.detect(data)
        assert result.header_row == 0
        assert result.headers == ["ID", "Value", "Active"]

    def test_detect_empty_data(self):
        """Test detection with empty data."""
        detector = HeaderDetector()
        result = detector.detect([])
        assert result.header_row == 0
        assert result.headers == []
        assert result.confidence == 0.0

    def test_header_detection_result_creation(self):
        """Test creating HeaderDetectionResult."""
        result = HeaderDetectionResult(
            header_row=0,
            headers=["Name", "Age"],
            confidence=0.8,
            sample_data=[["Alice", 25]],
        )
        assert result.header_row == 0
        assert result.headers == ["Name", "Age"]
        assert result.confidence == 0.8
