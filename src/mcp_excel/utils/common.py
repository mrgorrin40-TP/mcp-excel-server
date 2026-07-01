"""Common utility functions for MCP Excel Server."""


def col_letter_to_index(cell_ref: str) -> int:
    """Convert column letter(s) to 1-based index.

    Examples:
        'A' -> 1
        'Z' -> 26
        'AA' -> 27

    Args:
        cell_ref: Cell reference containing column letters (e.g., "A", "Z", "AA", "B5")

    Returns:
        1-based column index
    """
    from openpyxl.utils import column_index_from_string

    col_letter = "".join(c for c in cell_ref if c.isalpha())
    return int(column_index_from_string(col_letter))
